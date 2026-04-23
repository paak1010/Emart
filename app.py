import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="수주 업로드 자동화 대시보드", layout="wide")

st.title("수주 데이터 채널별 서식 자동 분류기 (VLOOKUP 적용)")
st.markdown("일반 주문서(Raw Data)와 **점포코드(배송코드) 맵핑 파일**을 업로드하면, VLOOKUP을 통해 센터코드를 정확한 배송코드로 변환하여 산출합니다.")

# 1. 파일 업로드 창 2개 분리
col1, col2 = st.columns(2)
with col1:
    uploaded_raw = st.file_uploader("📦 1. 일반 주문서 파일 (Raw Data)", type=['xlsx', 'xls', 'csv'])
with col2:
    uploaded_map = st.file_uploader("🔗 2. 점포코드 맵핑 파일 (기준정보)", type=['xlsx', 'xls', 'csv'])

def to_excel(df, sheet_name="Summary(수주업로드용)"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

if uploaded_raw is not None and uploaded_map is not None:
    try:
        # ==========================================
        # Step 1. 일반 주문서 (Raw Data) 로드 및 시트 탐색
        # ==========================================
        if uploaded_raw.name.endswith('.csv'):
            raw_df = pd.read_csv(uploaded_raw)
        else:
            xls = pd.ExcelFile(uploaded_raw)
            target_sheet = xls.sheet_names[0]
            for sheet in xls.sheet_names:
                temp_df = pd.read_excel(xls, sheet_name=sheet, nrows=3)
                if '점포코드' in temp_df.columns:
                    target_sheet = sheet
                    break
            raw_df = pd.read_excel(xls, sheet_name=target_sheet)
            
        if '점포코드' not in raw_df.columns:
            st.error("❌ 일반 주문서 파일에서 '점포코드' 열을 찾을 수 없습니다.")
            st.stop()

        # ==========================================
        # Step 2. 맵핑 파일 (점포코드 시트) 로드
        # ==========================================
        if uploaded_map.name.endswith('.csv'):
            map_df = pd.read_csv(uploaded_map)
        else:
            # 맵핑 파일의 첫 번째 시트를 읽어옵니다. (점포코드 시트만 따로 저장해서 올리는 것을 권장)
            map_df = pd.read_excel(uploaded_map)
            
        if '센터코드' not in map_df.columns or '배송코드' not in map_df.columns:
            st.error("❌ 맵핑 파일에는 반드시 **'센터코드'**와 **'배송코드'**라는 열 이름이 있어야 VLOOKUP이 가능합니다. 첫 줄(헤더)을 확인해 주세요.")
            st.stop()

        # 병합을 위해 센터코드 데이터 타입을 문자열로 통일하고 공백/소수점 제거
        raw_df['센터코드'] = raw_df['센터코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
        map_df['센터코드'] = map_df['센터코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
        
        # 맵핑 파일에 중복된 센터코드가 있을 경우 첫 번째 값만 남김 (다대일 병합 방지)
        map_df = map_df.drop_duplicates(subset=['센터코드'])

        # ==========================================
        # Step 3. VLOOKUP 실행 (pd.merge)
        # ==========================================
        # raw_df에 map_df의 '배송코드'를 센터코드 기준으로 레프트 조인
        raw_df = pd.merge(raw_df, map_df[['센터코드', '배송코드']], on='센터코드', how='left')
        
        # 맵핑 실패 시(NaN), 빈 칸으로 남기거나 에러 방지용으로 기존 센터코드 유지
        raw_df['배송코드'] = raw_df['배송코드'].fillna('') 

        st.success("✅ 파일 업로드 및 VLOOKUP 데이터 맵핑이 완료되었습니다.")

        # 결측치 제거 및 정수형 변환
        raw_df = raw_df.dropna(subset=['점포코드'])
        raw_df['점포코드'] = pd.to_numeric(raw_df['점포코드'], errors='coerce').fillna(0).astype(int)

        # ==========================================
        # Step 4. 채널별 데이터 분할 및 포맷팅
        # ==========================================
        emart_mask = ((raw_df['점포코드'] >= 1000) & (raw_df['점포코드'] <= 1999)) | (raw_df['점포코드'] >= 9000)
        traders_mask = (raw_df['점포코드'] >= 2000) & (raw_df['점포코드'] <= 2999)
        nobrand_mask = (raw_df['점포코드'] >= 3000) & (raw_df['점포코드'] <= 3999)

        emart_df = raw_df[emart_mask].copy()
        traders_df = raw_df[traders_mask].copy()
        nobrand_df = raw_df[nobrand_mask].copy()

        def extract_core_columns(df, channel_type):
            if df.empty: return pd.DataFrame()
            
            # 수량 0 필터링
            df['수량'] = pd.to_numeric(df.get('수량', 0), errors='coerce').fillna(0)
            df = df[df['수량'] > 0].copy()
            
            if df.empty: return pd.DataFrame()
            
            if channel_type == 'traders':
                order_code = df.get('문서번호', df.get('전표번호', '81011010'))
            else:
                order_code = '81010000'

            formatted = pd.DataFrame({
                '발주코드': order_code,
                '배송코드': df.get('배송코드', ''), # VLOOKUP으로 가져온 새로운 배송코드 매핑
                '상품코드': df.get('상품코드', ''),
                '수량': df['수량'],
                '단가': df.get('발주원가', 0),
                'Total Amount': df.get('발주금액', 0)
            })
            return formatted

        final_emart = extract_core_columns(emart_df, 'emart')
        final_traders = extract_core_columns(traders_df, 'traders')
        final_nobrand = extract_core_columns(nobrand_df, 'nobrand')

        # ==========================================
        # Step 5. 화면 출력 및 다운로드
        # ==========================================
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
