import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="수주 업로드 자동화 대시보드", layout="wide")

st.title("수주 데이터 통합 자동 분류기 (매핑 및 그룹핑)")
st.markdown("""
1. 일반 주문서(Raw Data)를 업로드합니다.
2. **제품명 시트**가 포함된 맵핑 파일(서식파일 등)을 업로드합니다.
3. 바코드가 **상품코드(기획)**으로 변환되고, 배송코드와 상품코드가 같은 건은 **자동으로 합산**되어 **단일 파일**로 출력됩니다.
""")

# 파일 업로드 창 (2개로 분리)
col1, col2 = st.columns(2)
with col1:
    uploaded_raw = st.file_uploader("📦 1. 일반 주문서 (Raw Data)", type=['xlsx', 'xls', 'csv'])
with col2:
    uploaded_product = st.file_uploader("📋 2. 제품명 맵핑 파일 ('제품명' 시트 포함)", type=['xlsx', 'xls', 'csv'])

def to_excel(df, sheet_name="통합_수주업로드"):
    """데이터프레임을 엑셀 파일(메모리)로 변환"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

if uploaded_raw and uploaded_product:
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

        # ==========================================
        # Step 2. 제품명 맵핑 파일 로드 ('제품명' 시트)
        # ==========================================
        if uploaded_product.name.endswith('.csv'):
            prod_df = pd.read_csv(uploaded_product)
        else:
            xls_prod = pd.ExcelFile(uploaded_product)
            # '제품명'이라는 단어가 포함된 시트 찾기
            prod_sheet = [s for s in xls_prod.sheet_names if '제품명' in s]
            if not prod_sheet:
                st.error("❌ 맵핑 파일에서 '제품명' 시트를 찾을 수 없습니다.")
                st.stop()
            prod_df = pd.read_excel(xls_prod, sheet_name=prod_sheet[0])
        
        # 맵핑을 위해 컬럼 이름 정리 (공백 제거)
        prod_df.columns = prod_df.columns.str.strip()
        
        # 필수 컬럼 존재 확인 ('바코드', '상품코드(기획)')
        if '바코드' not in prod_df.columns or '상품코드(기획)' not in prod_df.columns:
            st.error("❌ 제품명 시트에 '바코드' 또는 '상품코드(기획)' 열이 없습니다.")
            st.stop()
            
        # 상품명 컬럼 찾기 ('상품명(기획)' 우선, 없으면 '이마트 상품명' 등 사용)
        name_col = '상품명(기획)' if '상품명(기획)' in prod_df.columns else ('이마트 상품명' if '이마트 상품명' in prod_df.columns else '상품명')

        # VLOOKUP을 위한 키값 전처리 (문자열 변환 및 소수점 제거)
        prod_df['바코드'] = prod_df['바코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
        raw_df['상품코드'] = raw_df['상품코드'].astype(str).str.replace('.0', '', regex=False).str.strip()

        # ==========================================
        # Step 3. 데이터 전처리 (결측치 제거, 센터코드 정리 등)
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
        # Step 4. 채널별 분류 및 배송코드 매핑 로직
        # ==========================================
        mapping_dict = {
            'E-mart': {'9110': '81010902', '9120': '81010905', '9100': '81010903'},
            'E-mart(TRD)': {'9150': '81033036', '9102': '89011174', '9120': '81011012'},
            'E-mart(노브랜드)': {'9102': '89011175', '9130': '81010904', '9120': '81010968', '9110': '81010969'}
        }

        def determine_customer_and_delivery(row):
            code = row['점포코드']
            center = row['센터코드']
            
            # 채널(Customer) 분류
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

            # 배송코드 매핑 (사전에 없으면 기존 센터코드 사용)
            delivery_code = mapping_dict.get(customer, {}).get(center, center)
            
            return pd.Series([customer, order_code, delivery_code])

        raw_df[['Customer', '발주코드', '배송코드']] = raw_df.apply(determine_customer_and_delivery, axis=1)

        # ==========================================
        # Step 5. 제품명(바코드) VLOOKUP 매핑
        # ==========================================
        # raw_df의 '상품코드'(바코드) 와 prod_df의 '바코드'를 매핑
        merged_df = pd.merge(raw_df, prod_df[['바코드', '상품코드(기획)', name_col]], 
                             left_on='상품코드', right_on='바코드', how='left')

        # 매핑된 기획 코드가 없으면 원래 바코드를 그대로 사용 (에러 방지)
        merged_df['최종_상품코드'] = merged_df['상품코드(기획)'].fillna(merged_df['상품코드'])
        merged_df['최종_상품명'] = merged_df[name_col].fillna(merged_df.get('상품명', ''))

        # ==========================================
        # Step 6. 최종 컬럼 구성 및 그룹핑 (합치기)
        # ==========================================
        final_df = merged_df[[
            'Customer', '발주코드', '배송코드', '최종_상품코드', '최종_상품명', '수량', '발주원가', '발주금액'
        ]].copy()
        
        final_df.rename(columns={'최종_상품코드': '상품코드', '최종_상품명': '상품명', '발주원가': '단가', '발주금액': 'Total Amount'}, inplace=True)

        # 배송코드와 상품코드가 같은 데이터 합치기 (수량과 Total Amount는 더하고, 나머지는 그룹핑 기준)
        group_cols = ['Customer', '발주코드', '배송코드', '상품코드', '상품명', '단가']
        grouped_df = final_df.groupby(group_cols, dropna=False, as_index=False)[['수량', 'Total Amount']].sum()

        # 정렬 (Customer 별로 예쁘게 보기 위해)
        grouped_df = grouped_df.sort_values(by=['Customer', '배송코드'])

        st.success("✅ 매핑 및 데이터 그룹핑이 성공적으로 완료되었습니다!")

        # ==========================================
        # Step 7. 화면 출력 및 단일 파일 다운로드
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
