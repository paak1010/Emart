import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="수주 업로드 자동화 대시보드", layout="wide")

st.title("수주 데이터 채널별 서식 자동 분류기")
st.markdown("일반 주문서(Raw Data)를 업로드하면 점포코드 기준으로 분할하고, 내장된 맵핑 규칙에 따라 **센터코드를 배송코드로 자동 변환**하여 산출합니다.")

# 파일 업로드 창 (다시 1개로 단일화)
uploaded_file = st.file_uploader("📦 일반 주문서 파일(Raw Data)을 업로드하세요 (xlsx, csv)", type=['xlsx', 'xls', 'csv'])

def to_excel(df, sheet_name="Summary(수주업로드용)"):
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
            st.error("❌ 일반 주문서 파일의 어떤 시트에서도 '점포코드' 컬럼을 찾을 수 없습니다.")
            st.stop()

        st.success(f"✅ 파일 업로드 성공! (데이터 추출 시트: {target_sheet})")

        # 결측치 제거, 정수형 변환 및 센터코드 문자열 전처리
        raw_df = raw_df.dropna(subset=['점포코드'])
        raw_df['점포코드'] = pd.to_numeric(raw_df['점포코드'], errors='coerce').fillna(0).astype(int)
        raw_df['센터코드'] = raw_df.get('센터코드', '').astype(str).str.replace('.0', '', regex=False).str.strip()

        # 2. 채널별 데이터 분할
        emart_mask = ((raw_df['점포코드'] >= 1000) & (raw_df['점포코드'] <= 1999)) | (raw_df['점포코드'] >= 9000)
        traders_mask = (raw_df['점포코드'] >= 2000) & (raw_df['점포코드'] <= 2999)
        nobrand_mask = (raw_df['점포코드'] >= 3000) & (raw_df['점포코드'] <= 3999)

        emart_df = raw_df[emart_mask].copy()
        traders_df = raw_df[traders_mask].copy()
        nobrand_df = raw_df[nobrand_mask].copy()

        # [핵심 로직] 채널별 센터코드 -> 배송코드 맵핑 딕셔너리
        mapping_dict = {
            'emart': {
                '9110': '81010902',
                '9120': '81010905',
                '9100': '81010903'
            },
            'traders': {
                '9150': '81033036',
                '9102': '89011174',
                '9120': '81011012'
            },
            'nobrand': {
                '9102': '89011175',
                '9130': '81010904',
                '9120': '81010968',
                '9110': '81010969'
            }
        }

        # 3. 추출 및 포맷팅 함수
        def extract_core_columns(df, channel_type):
            if df.empty: return pd.DataFrame()
            
            # 수량 0 필터링
            df['수량'] = pd.to_numeric(df.get('수량', 0), errors='coerce').fillna(0)
            df = df[df['수량'] > 0].copy()
            
            if df.empty: return pd.DataFrame()
            
            # 발주코드 설정
            if channel_type == 'traders':
                order_code = df.get('문서번호', df.get('전표번호', '81011010'))
            else:
                order_code = '81010000'

            # 배송코드 맵핑 (알려주신 딕셔너리 기준)
            # 사전에 정의된 코드가 없으면, 기존 센터코드를 그대로 가져옵니다.
            current_map = mapping_dict.get(channel_type, {})
            mapped_delivery_code = df['센터코드'].map(current_map).fillna(df['센터코드'])

            formatted = pd.DataFrame({
                '발주코드': order_code,
                '배송코드': mapped_delivery_code,
                '상품코드': df.get('상품코드', ''),
                '수량': df['수량'],
                '단가': df.get('발주원가', 0),
                'Total Amount': df.get('발주금액', 0)
            })
            return formatted

        final_emart = extract_core_columns(emart_df, 'emart')
        final_traders = extract_core_columns(traders_df, 'traders')
        final_nobrand = extract_core_columns(nobrand_df, 'nobrand')

        # 4. 화면 출력 및 다운로드
        st.subheader("데이터 변환 결과 및 다운로드")
        
        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown(f"**이마트** ({len(final_emart)}건)")
            if not final_emart.empty:
                st.dataframe(final_emart.head(5))
                st.download_button("📥 이마트 다운로드", data=to_excel(final_emart), file_name="수주업로드_이마트.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("해당 데이터 없음")

        with col2:
            st.markdown(f"**트레이더스** ({len(final_traders)}건)")
            if not final_traders.empty:
                st.dataframe(final_traders.head(5))
                st.download_button("📥 트레이더스 다운로드", data=to_excel(final_traders), file_name="수주업로드_트레이더스.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("해당 데이터 없음")

        with col3:
            st.markdown(f"**노브랜드** ({len(final_nobrand)}건)")
            if not final_nobrand.empty:
                st.dataframe(final_nobrand.head(5))
                st.download_button("📥 노브랜드 다운로드", data=to_excel(final_nobrand), file_name="수주업로드_노브랜드.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("해당 데이터 없음")

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
