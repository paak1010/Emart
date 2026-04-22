import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="수주 업로드 자동화 대시보드", layout="wide")

st.title("수주 데이터 채널별 서식 자동 분류기")
st.markdown("일반 주문서(Raw Data)를 업로드하면 점포코드 기준(이마트, 트레이더스, 노브랜드)으로 데이터를 분할하여 수주업로드용 지정 서식으로 변환합니다.")

# 파일 업로드
uploaded_file = st.file_uploader("일반 주문서 파일을 업로드하세요 (xlsx, csv)", type=['xlsx', 'xls', 'csv'])

def to_excel(df):
    """데이터프레임을 엑셀 파일(메모리)로 변환하여 첫 번째 시트에 저장"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 무조건 첫 번째 시트(Sheet1)에 저장
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

if uploaded_file is not None:
    try:
        # 확장자에 따른 데이터 로드
        if uploaded_file.name.endswith('.csv'):
            raw_df = pd.read_csv(uploaded_file)
        else:
            raw_df = pd.read_excel(uploaded_file)
            
        st.success("✅ 파일이 성공적으로 업로드되었습니다.")

        # 점포코드가 결측치인 경우 제외 및 정수형 변환 처리
        raw_df = raw_df.dropna(subset=['점포코드'])
        raw_df['점포코드'] = pd.to_numeric(raw_df['점포코드'], errors='coerce').fillna(0).astype(int)

        # 1. 채널별 데이터 분할 로직
        # 이마트: 1000~1999 또는 9000 이상
        emart_mask = ((raw_df['점포코드'] >= 1000) & (raw_df['점포코드'] <= 1999)) | (raw_df['점포코드'] >= 9000)
        # 트레이더스: 2000~2999
        traders_mask = (raw_df['점포코드'] >= 2000) & (raw_df['점포코드'] <= 2999)
        # 노브랜드: 3000~3999
        nobrand_mask = (raw_df['점포코드'] >= 3000) & (raw_df['점포코드'] <= 3999)

        emart_df = raw_df[emart_mask].copy()
        traders_df = raw_df[traders_mask].copy()
        nobrand_df = raw_df[nobrand_mask].copy()

        # 2. 지정된 컬럼 포맷으로 변환하는 함수
        def format_dataframe(df):
            if df.empty:
                return pd.DataFrame()
            
            # 발주코드/문서번호 필드 확인 (파일에 따라 컬럼명이 다를 수 있음)
            order_code_col = '문서번호' if '문서번호' in df.columns else ('전표번호' if '전표번호' in df.columns else '발주코드')
            
            formatted_df = pd.DataFrame({
                '수주일자': df.get('발주일자', ''),
                '납품일자': df.get('센터입하일자', ''),
                '발주코드': df.get(order_code_col, ''),
                '배송 코드': df.get('센터코드', ''),
                'ME코드': df.get('상품코드', ''), # 별도 Product Master 맵핑이 필요하다면 이 부분을 추후 merge로 고도화
                '제품명': df.get('상품명', ''),
                '수량': df.get('수량', 0),
                '가격': df.get('발주원가', 0),
                'total amount': df.get('발주금액', 0)
            })
            return formatted_df

        # 각 채널별 포맷 적용
        final_emart = format_dataframe(emart_df)
        final_traders = format_dataframe(traders_df)
        final_nobrand = format_dataframe(nobrand_df)

        # 3. 화면 출력 및 다운로드 버튼 생성
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
