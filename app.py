import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="수주 업로드 자동화 대시보드", layout="wide")

st.title("수주 데이터 채널별 서식 자동 분류기")
st.markdown("일반 주문서(Raw Data)를 업로드하면 점포코드 기준으로 분할하여 각 채널의 **Summary(수주업로드용)** 시트 서식에 맞게 산출합니다.")

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
            target_sheet = xls.sheet_names[0] # 기본값은 첫 번째 시트
            
            # 모든 시트를 탐색하며 '점포코드' 컬럼이 있는 진짜 Raw Data 시트 찾기
            for sheet in xls.sheet_names:
                temp_df = pd.read_excel(xls, sheet_name=sheet, nrows=3)
                if '점포코드' in temp_df.columns:
                    target_sheet = sheet
                    break
            
            raw_df = pd.read_excel(xls, sheet_name=target_sheet)
            
        # 안전 장치: 그래도 점포코드가 없으면 에러 메시지 띄우기
        if '점포코드' not in raw_df.columns:
            st.error("❌ 업로드하신 파일의 어떤 시트에서도 '점포코드' 컬럼을 찾을 수 없습니다. 원본 파일을 확인해주세요.")
            st.stop()

        st.success(f"✅ 파일 업로드 성공! (읽어온 시트명: {target_sheet})")

        # 결측치 제거 및 정수형 변환
        raw_df = raw_df.dropna(subset=['점포코드'])
        raw_df['점포코드'] = pd.to_numeric(raw_df['점포코드'], errors='coerce').fillna(0).astype(int)

        # 2. 채널별 데이터 분할 로직
        emart_mask = ((raw_df['점포코드'] >= 1000) & (raw_df['점포코드'] <= 1999)) | (raw_df['점포코드'] >= 9000)
        traders_mask = (raw_df['점포코드'] >= 2000) & (raw_df['점포코드'] <= 2999)
        nobrand_mask = (raw_df['점포코드'] >= 3000) & (raw_df['점포코드'] <= 3999)

        emart_df = raw_df[emart_mask].copy()
        traders_df = raw_df[traders_mask].copy()
        nobrand_df = raw_df[nobrand_mask].copy()

        # 3. 채널별 Summary(수주업로드용) 폼 매핑 함수
        def format_emart_nobrand(df):
            """이마트 및 노브랜드 Summary 시트 서식"""
            if df.empty: return pd.DataFrame()
            
            # Sum Code 생성 (배송코드 + 상품코드)
            center_code = df.get('센터코드', '').astype(str).str.replace('.0', '', regex=False)
            item_code = df.get('상품코드', '').astype(str)
            
            formatted = pd.DataFrame({
                'Sum Code': center_code + item_code,
                '발주코드': '81010000', # 고정 발주코드
                'Unnamed: 2': '',       # 빈 열
                '배송코드': df.get('센터코드', ''),
                '센터명': df.get('센터이름', ''),
                '상품코드': df.get('상품코드', ''),
                'Unnamed: 6': '',       # 빈 열
                'UNIT수량': df.get('수량', 0),
                'UNIT단가': df.get('발주원가', 0),
                'Total Amount': df.get('발주금액', 0),
                'Unnamed: 10': '',      # 빈 열
                '오늘재고': df.get('상품명', '') # 비고/상품명 기록
            })
            # 엑셀 산출 시 컬럼명이 비어보이도록 처리
            formatted.rename(columns={'Unnamed: 2': '', 'Unnamed: 6': '', 'Unnamed: 10': ''}, inplace=True)
            return formatted

        def format_traders(df):
            """이마트 트레이더스 Summary 시트 서식"""
            if df.empty: return pd.DataFrame()
            
            formatted = pd.DataFrame({
                '납품일자': df.get('센터입하일자', ''),
                '발주코드': df.get('문서번호', '81011010'), # 트레이더스는 주로 문서번호 사용
                'Unnamed: 2': '',      # 빈 열
                '배송코드': df.get('센터코드', ''),
                '센터명': df.get('센터이름', ''),
                '상품코드': df.get('상품코드', ''),
                '제품명': df.get('상품명', ''),
                'UNIT수량': df.get('수량', 0),
                'UNIT단가': df.get('발주원가', 0),
                'Total Amount': df.get('발주금액', 0),
                'LOT': df.get('LOT', ''),
                '점포명': df.get('점포명', '')
            })
            formatted.rename(columns={'Unnamed: 2': ''}, inplace=True)
            return formatted

        # 서식 변환 실행
        final_emart = format_emart_nobrand(emart_df)
        final_traders = format_traders(traders_df)
        final_nobrand = format_emart_nobrand(nobrand_df)

        # 4. 화면 출력 및 다운로드 버튼
        st.subheader("데이터 변환 결과 및 다운로드")
        
        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown(f"**이마트** ({len(final_emart)}건)")
            if not final_emart.empty:
                st.dataframe(final_emart.head(5))
                st.download_button(
                    label="📥 이마트 서식 다운로드",
                    data=to_excel(final_emart),
                    file_name="수주업로드_이마트.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("해당되는 데이터가 없습니다.")

        with col2:
            st.markdown(f"**트레이더스** ({len(final_traders)}건)")
            if not final_traders.empty:
                st.dataframe(final_traders.head(5))
                st.download_button(
                    label="📥 트레이더스 서식 다운로드",
                    data=to_excel(final_traders),
                    file_name="수주업로드_트레이더스.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("해당되는 데이터가 없습니다.")

        with col3:
            st.markdown(f"**노브랜드** ({len(final_nobrand)}건)")
            if not final_nobrand.empty:
                st.dataframe(final_nobrand.head(5))
                st.download_button(
                    label="📥 노브랜드 서식 다운로드",
                    data=to_excel(final_nobrand),
                    file_name="수주업로드_노브랜드.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("해당되는 데이터가 없습니다.")

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
