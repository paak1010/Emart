import streamlit as st
import pandas as pd
import os
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="수주 업로드 자동화 대시보드", layout="wide")

st.title("수주 데이터 통합 자동 분류기 (최종 완성본)")
st.markdown("""
**일반 주문서(Raw Data)** 파일만 업로드하세요.
- 모든 발주코드는 **81010000**으로 자동 통일됩니다.
- 깃허브 내 서식파일의 `제품명` 시트를 참조하여 바코드를 기획 코드로 변환합니다.
- 배송코드에 맞춰 **'배송처'** 이름이 자동으로 입력됩니다.
- 데이터는 지정된 완벽한 컬럼 순서로 자동 정렬되어 산출됩니다.
""")

# ==========================================
# ⚙️ 1. 현재 날짜 자동 설정
# ==========================================
today_str = datetime.today().strftime("%Y%m%d")

# ==========================================
# ⚙️ 2. 서버(깃허브) 내장 서식 파일 설정
# ==========================================
TEMPLATE_FILES = [
    "NEW 이마트 서식파일_20260420납품.xlsx",
    "NEW 이마트 트레이더스(한익스점포확인)_260327납품(평택9여주0대구4).xlsx",
    "NEW 노브랜드_20260409납품.xlsx"
]

@st.cache_data
def load_master_product_data():
    """서버 내 서식 파일들에서 제품명 시트를 취합하여 마스터 매핑표를 만듭니다."""
    appended_data = []
    for file_name in TEMPLATE_FILES:
        if os.path.exists(file_name):
            xls = pd.ExcelFile(file_name)
            prod_sheets = [s for s in xls.sheet_names if '제품명' in s]
            if prod_sheets:
                df = pd.read_excel(xls, sheet_name=prod_sheets[0])
                df.columns = df.columns.str.strip()
                appended_data.append(df)
    
    if not appended_data:
        return None
        
    master_df = pd.concat(appended_data, ignore_index=True)
    if '바코드' in master_df.columns:
        master_df['바코드'] = master_df['바코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
        master_df = master_df.drop_duplicates(subset=['바코드'], keep='first')
        
    return master_df

def to_excel(df, sheet_name="통합_수주업로드"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# 매핑 데이터 로드
prod_df = load_master_product_data()

if prod_df is None or prod_df.empty:
    st.warning("⚠️ 서버에서 서식파일을 찾을 수 없습니다. 깃허브 파일명을 확인해주세요.")
    st.stop()

# ==========================================
# ⚙️ 3. 파일 업로드 및 데이터 처리
# ==========================================
uploaded_raw = st.file_uploader("📦 일반 주문서 (Raw Data) 업로드", type=['xlsx', 'xls', 'csv'])

if uploaded_raw:
    try:
        # Step 1. Raw Data 로드
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

        # Step 2. 데이터 전처리
        raw_df = raw_df.dropna(subset=['점포코드'])
        raw_df['점포코드'] = pd.to_numeric(raw_df['점포코드'], errors='coerce').fillna(0).astype(int)
        raw_df['센터코드'] = raw_df.get('센터코드', '').astype(str).str.replace('.0', '', regex=False).str.strip()
        raw_df['수량'] = pd.to_numeric(raw_df.get('수량', 0), errors='coerce').fillna(0)
        
        # 점입점일자를 문자열로 변환하여 배송일자로 지정
        raw_df['배송일자'] = raw_df.get('점입점일자', '').astype(str).str.replace('.0', '', regex=False).str.strip()
        
        # 수량이 0보다 큰 건만 남기기
        raw_df = raw_df[raw_df['수량'] > 0].copy() 

        # Step 3. 채널 분류 및 배송코드 매핑
        mapping_dict = {
            'E-mart': {'9110': '81010902', '9120': '81010905', '9100': '81010903'},
            'E-mart(TRD)': {'9150': '81033036', '9102': '89011174', '9120': '81011012'},
            'E-mart(노브랜드)': {'9102': '89011175', '9130': '81010904', '9120': '81010968', '9110': '81010969'}
        }

        def process_row(row):
            code = row['점포코드']
            center = row['센터코드']
            
            if (1000 <= code <= 1999) or code >= 9000:
                customer = 'E-mart'
            elif 2000 <= code <= 2999:
                customer = 'E-mart(TRD)'
            elif 3000 <= code <= 3999:
                customer = 'E-mart(노브랜드)'
            else:
                customer = 'Unknown'

            delivery_code = mapping_dict.get(customer, {}).get(center, center)
            return pd.Series([customer, delivery_code])

        raw_df[['Customer', '배송코드']] = raw_df.apply(process_row, axis=1)

        # Step 4. 상품 정보 매핑 (바코드 -> 기획코드)
        raw_df['상품코드'] = raw_df['상품코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
        name_col = '상품명(기획)' if '상품명(기획)' in prod_df.columns else '상품명'
        
        merged_df = pd.merge(raw_df, prod_df[['바코드', '상품코드(기획)', name_col]], 
                             left_on='상품코드', right_on='바코드', how='left')

        merged_df['최종_상품코드'] = merged_df['상품코드(기획)'].fillna(merged_df['상품코드'])
        merged_df['최종_상품명'] = merged_df[name_col].fillna(merged_df.get('상품명', ''))

        # Step 5. [신규] 배송처 한글명 매핑 데이터 구축
        delivery_name_map = {
            '81010901': '이마트 백암물류센터',
            '81010902': '이마트 시화물류센터',
            '81010903': '이마트 대구물류센터',
            '81010905': '이마트 여주물류센터',
            '81010904': '이마트 노브랜드 여주2물류센터',
            '81010968': '이마트 노브랜드 여주물류센터',
            '81010969': '이마트 노브랜드 시화물류센터',
            '89011175': '이마트 노브랜드 대구물류(신규)',
            '81010906': '이마트 광주물류센터',
            '81033036': '이마트 트레이더스 평택물류'
        }

        # 기본값 및 배송처 매핑 적용
        merged_df['발주코드'] = '81010000'
        merged_df['날짜'] = today_str
        merged_df['배송처'] = merged_df['배송코드'].astype(str).map(delivery_name_map).fillna('')
        
        final_df = merged_df[[
            '날짜', '배송일자', '발주코드', 'Customer', '배송코드', '배송처', '최종_상품코드', '최종_상품명', '수량', '발주원가', '발주금액'
        ]].copy()
        
        final_df.rename(columns={'최종_상품코드': '상품코드', '최종_상품명': '상품명', '발주원가': '단가', '발주금액': 'Total Amount'}, inplace=True)

        # 그룹핑 시 발주코드, Customer 순서 반영 및 배송처 포함
        group_cols = ['날짜', '배송일자', '발주코드', 'Customer', '배송코드', '배송처', '상품코드', '상품명', '단가']
        grouped_df = final_df.groupby(group_cols, dropna=False, as_index=False)[['수량', 'Total Amount']].sum()
        grouped_df = grouped_df.sort_values(by=['Customer', '배송코드'])

        # ⭐ [핵심 변경] 최종 컬럼 순서 (발주코드 -> Customer -> 배송코드 -> 배송처)
        final_column_order = [
            '날짜', '배송일자', '발주코드', 'Customer', '배송코드', '배송처', 
            '상품코드', '상품명', '수량', '단가', 'Total Amount'
        ]
        grouped_df = grouped_df[final_column_order]

        # Step 6. 결과 출력 및 다운로드
        st.success("✅ 배송처 매핑 및 최종 컬럼 정렬, 데이터 합산이 완료되었습니다.")
        st.dataframe(grouped_df)

        st.download_button(
            label="📥 통합 수주업로드 파일 다운로드",
            data=to_excel(grouped_df),
            file_name=f"수주업로드_통합본_{today_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
