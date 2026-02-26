import pandas as pd

def classify_transactions(df, target_months):
    df = df.copy()

    # 1. 헤더 설정 및 데이터 정리
    header_row_idx = df[df.eq("날짜").any(axis=1)].index[0]
    data_df = df.iloc[header_row_idx + 1:].copy()
    data_df.columns = df.iloc[header_row_idx]

    # 2. 날짜 필터링 (2025-11-01 형식 대응)
    data_df['날짜'] = pd.to_datetime(data_df['날짜'])

    # 선택된 월들에 해당하는 데이터만 필터링
    # dt.strftime('%Y-%m')을 쓰면 문자열 비교로 간단히 여러 달을 체크할 수 있습니다.
    filtered_df = data_df[data_df['날짜'].dt.strftime('%Y-%m').isin(target_months)].copy()

    # 3. 정렬: 날짜와 시간 모두 역순 (최신 내역이 위로)
    # 사진 3번을 보면 '시간' 컬럼이 따로 있으므로 이를 활용합니다.
    filtered_df = filtered_df.sort_values(by=['날짜', '시간'], ascending=[False, False])

    # 4. 분류 시작
    classified_data = {}

    # 4-1. 월별 -> 타입별(수입/이체) -> 대분류별(지출) 분류
    for month_str in target_months:
        month_df = filtered_df[filtered_df['날짜'].dt.strftime('%Y-%m') == month_str]
        if month_df.empty:
            continue
            
        m_num = int(month_str.split('-')[1]) # '11' 추출

        # A. 수입 & 이체 (타입별로 분리)
        for t_type in ['수입', '이체']:
            type_df = month_df[month_df['타입'] == t_type]
            if not type_df.empty:
                classified_data[f"{m_num}월 {t_type}"] = type_df

        # B. 지출 (대분류별로 분리)
        expense_df = month_df[month_df['타입'] == '지출'].copy()
        if not expense_df.empty:
            # 금액 컬럼 숫자형 변환 (쉼표 등 제거)
            expense_df['금액_num'] = pd.to_numeric(expense_df['금액'].astype(str).str.replace(',', ''), errors='coerce')

            # 1단계: 대분류별 합계 계산 (index: 대분류, value: 합계)
            category_totals = expense_df.groupby('대분류')['금액_num'].sum().sort_values(ascending=True)
            print(category_totals.index.tolist())

            # 2단계: 합계가 높은 대분류 순서대로 시트 데이터 생성 (category_totals.index: 대분류)
            for cat in category_totals.index:
                if pd.isna(cat): continue
                
                # 대분류별 데이터 추출 (cat: 대분류)
                cat_df = expense_df[expense_df['대분류'] == cat].drop(columns=['금액_num'])
                cat_name = cat.replace('/', ',')
                
                # 시트 이름에 월과 대분류 포함 (합계 순서대로 삽입됨)
                classified_data[f"{m_num}월 지출({cat_name})"] = cat_df

    return classified_data