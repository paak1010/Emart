import streamlit as st
import pandas as pd
import os
from io import BytesIO

st.set_page_config(page_title="수주 업로드 자동화 대시보드", layout="wide")

st.title("수주 데이터 통합 자동 분류기 (매핑 자동화)")
st.markdown("""
**일반 주문서(Raw Data)** 파일만 업로드하세요.
깃허브 서버에 내장된 각 서식파일의 `제품명` 시트를 자동으로 취합하여 바코드 변환 및 데이터 합산을 수행합니다.
""")

# ==========================================
# ⚙️ 깃허브에 업로드된 서식 파일 이름 설정
# ==========================================
# (주의: 실제 깃허브에 올려둔 엑셀 파일 이름과 똑같아야 합니다.)
TEMPLATE_FILES = [
    "NEW 이마트 서식파일_20260420납품.xlsx",
    "NEW 이마트 트레이더스(한익스점포확인)_260327납품(평택9여주0대구4).xlsx",
    "NEW 노브랜드_20260409납품.xlsx"
]

@st.cache_data
def load_master_product_data():
    """서버에 있는 여러 서식 파일에서 '제품명' 시트를 모두 읽어와 하나로 합칩니다."""
    appended_data = []
    for file_name in TEMPLATE_FILES:
        # 파일이 실제로 폴더(깃허브)에 존재하는지 확인
        if os.path.exists(file_name):
            xls = pd.ExcelFile(file_name)
            # '제품명'이라는 글자가 포함된 시트 찾기
            prod_sheets = [s for s in xls.sheet_names if '제품명' in s]
            if prod_sheets:
                df = pd.read_excel(xls, sheet_name=prod_sheets[0])
                df.columns = df.columns.str.strip()
                appended_data.append(df)
    
    if not appended_data:
        return None
        
    # 모든 제품명 시트를 하나로 병합
    master_df = pd.concat(appended_data, ignore_index=True)
    
    if '바코드' in master_df.columns:
        # 바코드 전처리 및 중복 데이터(같은 제품) 제거
        master_df['바코드'] = master_df['바코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
        master_df = master_df.drop_duplicates(subset=['바코드'], keep='first')
        
    return master_df

def to_excel(df, sheet_name="통합_수주업로드"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# 1. 맵핑 데이터 자동 로드
prod_df = load_master_product_data()

if prod_df is None or prod_df.empty:
    st.warning("⚠️ 지정된 서식파일을 서버에서 찾을 수 없거나 '제품명' 시트가 없습니다. 깃허브에 서식파일이 정확한 이름으로 올라가 있는지 확인해주세요.")
    st.stop()

# 2. 단일 파일 업로드 창
uploaded_raw = st.file_uploader("📦 일반 주문서 (Raw Data) 업로드", type=['xlsx', 'xls', 'csv'])

if uploaded_raw:
    try:
        # ==========================================
        # Step 1. 일반 주문서 (Raw Data) 로드
        # ==========================================
        if uploaded_raw.name.endswith('.csv'):
            raw_df = pd.read_csv(uploaded_raw)
        else:
            xls_raw = pd.ExcelFile(uploaded_raw)
            target_sheet = xls_raw.sheet_names[0]
            for sheet in xls_raw.sheet_names:
                temp_df = pd.read_excel(xls_raw, sheet_name=sheet, nrows=3)
                if '점포코드' in temp_df.columns:
                    target_sheet = sheet
                    break
            raw_df = pd.read_excel(xls_raw, sheet_name=target_sheet)
            
        if '점포코드' not in raw_df.columns:
            st.error("❌ 일반 주문서 파일에서 '점포코드' 열을 찾을 수 없습니다.")
            st.stop()

        # 필수 컬럼 존재 확인
        if '상품코드(기획)' not in prod_df.columns:
            st.error("❌ 제품명 시트들에 '상품코드(기획)' 열이 존재하지 않습니다.")
            st.stop()
            
        name_col = '상품명(기획)' if '상품명(기획)' in prod_df.columns else ('이마트 상품명' if '이마트 상품명' in prod_df.columns else '상품명')

        raw_df['상품코드'] = raw_df['상품코드'].astype(str).str.replace('.0', '', regex=False).str.strip()

        # ==========================================
        # Step 2. 데이터 전처리
        # ==========================================
        raw_df = raw_df.dropna(subset=['점포코드'])
        raw_df['점포코드'] = pd.to_numeric(raw_df['점포코드'], errors='coerce').fillna(0).astype(int)
        raw_df['센터코드'] = raw_df.get('센터코드', '').astype(str).str.replace('.0', '', regex=False).str.strip()
        raw_df['수량'] = pd.to_numeric(raw_df.get('수량', 0), errors='coerce').fillna(0)
        raw_df['발주금액'] = pd.to_numeric(raw_df.get('발주금액', 0), errors='coerce').fillna(0)
        raw_df['발주원가'] = pd.to_numeric(raw_df.get('발주원가', 0), errors='coerce').fillna(0)
        
        # 수량이 0보다 큰 것만 필터링
        raw_df = raw_df[raw_df['수량'] > 0].copy()

        # ==========================================
        # Step 3. 채널 분류 및 배송코드 매핑
        # ==========================================
        mapping_dict = {
            'E-mart': {'9110': '81010902', '9120': '81010905', '9100': '81010903'},
            'E-mart(TRD)': {'9150': '81033036', '9102': '89011174', '9120': '81011012'},
            'E-mart(노브랜드)': {'9102': '89011175', '9130': '81010904', '9120': '81010968', '9110': '81010969'}
        }

        def determine_customer_and_delivery(row):
            code = row['점포코드']
            center = row['센터코드']
            
            if (1000 <= code <= 1999) or code >= 9000:
                customer = 'E-mart'
                order_code = '81010000'
            elif 2000 <= code <= 2999:
                customer = 'E-mart(TRD)'
                order_code = row.get('문서번호', row.get('전표번호', '81011010'))
            elif 3000 <= code <= 3999:
                customer = 'E-mart(노브랜드)'
                order_code = '81010000'
            else:
                customer = 'Unknown'
                order_code = '81010000'

            delivery_code = mapping_dict.get(customer, {}).get(center, center)
            return pd.Series([customer, order_code, delivery_code])

        raw_df[['Customer', '발주코드', '배송코드']] = raw_df.apply(determine_customer_and_delivery, axis=1)

        # ==========================================
        # Step 4. VLOOKUP 병합 (상품코드 & 상품명)
        # ==========================================
        merged_df = pd.merge(raw_df, prod_df[['바코드', '상품코드(기획)', name_col]], 
                             left_on='상품코드', right_on='바코드', how='left')

        merged_df['최종_상품코드'] = merged_df['상품코드(기획)'].fillna(merged_df['상품코드'])
        merged_df['최종_상품명'] = merged_df[name_col].fillna(merged_df.get('상품명', ''))

        # ==========================================
        # Step 5. 최종 컬럼 구성 및 그룹핑
        # ==========================================
        final_df = merged_df[[
            'Customer', '발주코드', '배송코드', '최종_상품코드', '최종_상품명', '수량', '발주원가', '발주금액'
        ]].copy()
        
        final_df.rename(columns={'최종_상품코드': '상품코드', '최종_상품명': '상품명', '발주원가': '단가', '발주금액': 'Total Amount'}, inplace=True)

        group_cols = ['Customer', '발주코드', '배송코드', '상품코드', '상품명', '단가']
        grouped_df = final_df.groupby(group_cols, dropna=False, as_index=False)[['수량', 'Total Amount']].sum()
        grouped_df = grouped_df.sort_values(by=['Customer', '배송코드'])

        st.success("✅ 자동 맵핑 및 데이터 그룹핑이 성공적으로 완료되었습니다!")

        # ==========================================
        # Step 6. 화면 출력 및 단일 파일 다운로드
        # ==========================================
        st.subheader(f"📊 통합 산출 결과 (총 {len(grouped_df)}건)")
        st.dataframe(grouped_df)

        st.download_button(
            label="📥 통합 수주업로드 다운로드 (클릭)",
            data=to_excel(grouped_df),
            file_name="수주업로드_통합본.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
