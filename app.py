import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="수주 업로드 자동화 대시보드", layout="wide")

st.title("수주 데이터 채널별 서식 자동 분류기 (핵심 컬럼 추출형)")
st.markdown("일반 주문서(Raw Data)를 업로드하면 점포코드 기준으로 분할하여 **발주/배송/상품코드, 수량, 단가, 총액**만 산출하며, **수량이 0인 항목은 자동 제외**됩니다.")

# 파일 업로드
uploaded_file = st.file_uploader("일반 주문서 파일을 업로드하세요 (xlsx, csv)", type=['xlsx', 'xls', 'csv'])

def to_excel(df, sheet_name="Summary(수주업로드용)"):
    """데이터프레임을 엑셀 파일(메모리)로 변환"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

if uploaded_file is not None:
    try:
        # 1. 지능형 시트 탐색 (점포코드 에러 해결 로직)
        if uploaded_file.name.endswith('.csv'):
            raw_df = pd.read_csv(uploaded_file)
        else:
            xls = pd.ExcelFile(uploaded_file)
            target_sheet = xls.sheet_names[0]
            
            for sheet in xls.sheet_names:
                temp_df = pd.read_excel(xls, sheet_name=sheet, nrows=3)
                if '점포코드' in temp_df.columns:
                    target_sheet = sheet
                    break
            
            raw_df = pd.read_excel(xls, sheet_name=target_sheet)
            
        if '점포코드' not in raw_df.columns:
            st.error("❌ 업로드하신 파일의 어떤 시트에서도 '점포코드' 컬럼을 찾을 수 없습니다. 원본 파일을 확인해주세요.")
            st.stop()

        st.success(f"✅ 파일 업로드 성공! (데이터 추출 시트: {target_sheet})")

        # 결측치 제거 및 정수형 변환
        raw_df = raw_df.dropna(subset=['점포코드'])
        raw_df['점포코드'] = pd.to_numeric(raw_df['점포코드'], errors='coerce').fillna(0).astype(int)

        # 2. 채널별 데이터 분할
        emart_mask = ((raw_df['점포코드'] >= 1000) & (raw_df['점포코드'] <= 1999)) | (raw_df['점포코드'] >= 9000)
        traders_mask = (raw_df['점포코드'] >= 2000) & (raw_df['점포코드'] <= 2999)
        nobrand_mask = (raw_df['점포코드'] >= 3000) & (raw_df['점포코드'] <= 3999)

        emart_df = raw_df[emart_mask].copy()
        traders_df = raw_df[traders_mask].copy()
        nobrand_df = raw_df[nobrand_mask].copy()

        # 3. 요청하신 6개 핵심 컬럼만 추출 + 수량 0 제외 함수
        def extract_core_columns(df, channel_type):
            if df.empty: return pd.DataFrame()
            
            # [추가된 로직] 수량을 숫자형으로 변환 후 0인 데이터(혹은 0보다 작은 데이터) 필터링
            df['수량'] = pd.to_numeric(df.get('수량', 0), errors='coerce').fillna(0)
            df = df[df['수량'] > 0].copy()
            
            # 필터링 후 데이터가 없으면 빈 데이터프레임 반환
            if df.empty: return pd.DataFrame()
            
            # 발주코드 로직: 트레이더스는 문서번호(또는 전표번호), 이마트/노브랜드는 지정 양식대로 81010000 고정
            if channel_type == 'traders':
                order_code = df.get('문서번호', df.get('전표번호', '81011010'))
            else:
                order_code = '81010000'

            formatted = pd.DataFrame({
                '발주코드': order_code,
                '배송코드': df.get('센터코드', ''),
                '상품코드': df.get('상품코드', ''),
                '수량': df['수량'],
                '단가': df.get('발주원가', 0),
                'Total Amount': df.get('발주금액', 0)
            })
            return formatted

        # 핵심 서식 변환 실행
        final_emart = extract_core_columns(emart_df, 'emart')
        final_traders = extract_core_columns(traders_df, 'traders')
        final_nobrand = extract_core_columns(nobrand_df, 'nobrand')

        # 4. 화면 출력 및 다운로드 버튼
        st.subheader("데이터 변환 결과 및 다운로드")
        
        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown(f"**이마트** ({len(final_emart)}건)")
            if not final_emart.empty:
                st.dataframe(final_emart.head(5))
                st.download_button(
                    label="📥 이마트 다운로드",
                    data=to_excel(final_emart),
                    file_name="수주업로드_이마트.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("해당 데이터 없음")

        with col2:
            st.markdown(f"**트레이더스** ({len(final_traders)}건)")
            if not final_traders.empty:
                st.dataframe(final_traders.head(5))
                st.download_button(
                    label="📥 트레이더스 다운로드",
                    data=to_excel(final_traders),
                    file_name="수주업로드_트레이더스.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("해당 데이터 없음")

        with col3:
            st.markdown(f"**노브랜드** ({len(final_nobrand)}건)")
            if not final_nobrand.empty:
                st.dataframe(final_nobrand.head(5))
                st.download_button(
                    label="📥 노브랜드 다운로드",
                    data=to_excel(final_nobrand),
                    file_name="수주업로드_노브랜드.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("해당 데이터 없음")

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
