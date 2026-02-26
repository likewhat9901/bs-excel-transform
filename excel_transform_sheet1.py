import pandas as pd

def get_data_df(df):
    """헤더 행을 찾아 데이터프레임 설정"""
    print("원본 df 크기:", df.shape)
    header_row_idx = df[df.eq("항목").any(axis=1)].index[0]
    print(f"헤더 행 인덱스: {header_row_idx}")
    
    data_df = df.iloc[header_row_idx + 1:].copy()
    data_df.columns = df.iloc[header_row_idx]
    print("\ndata_df 크기:", data_df.shape)
    print("'항목' 컬럼 unique 값:", data_df['항목'].unique())
    
    return data_df

def calculate_totals_and_averages(section_df):
    """월별 데이터로부터 총계와 월평균 계산"""
    month_cols = section_df.iloc[:, 4:]
    section_df['총계'] = pd.to_numeric(month_cols.sum(axis=1), errors='coerce')
    section_df['월평균'] = pd.to_numeric(month_cols.mean(axis=1), errors='coerce')
    return section_df

def extract_and_sort_section(data_df, start_item, end_item, section_name):
    """지정된 항목 범위를 추출하고 총계 기준으로 정렬"""
    print(f"\n=== {section_name} 영역 추출 ===")
    start_idx = data_df[data_df['항목'] == start_item].index[0]
    end_idx = data_df[data_df['항목'] == end_item].index[0] - 1
    print(f"{section_name} 구간: {start_idx} ~ {end_idx}")
    
    section = data_df.loc[start_idx:end_idx].copy()
    section = calculate_totals_and_averages(section)
    
    section['총계_num'] = section['총계']
    sorted_section = section.sort_values(by='총계_num', ascending=False).drop(columns=['총계_num'])
    print(f"정렬 후 {section_name} 섹션:\n{sorted_section}")
    
    return sorted_section, start_idx, end_idx

def transform_summary(df):
    print("원본 df 크기:", df.shape)
    
    data_df = get_data_df(df)
    
    sorted_income, income_start, income_end = extract_and_sort_section(
        data_df, '금융수입', '월수입 총계', '수입'
    )
    
    sorted_expense, expense_start, expense_end = extract_and_sort_section(
        data_df, '경조사', '월지출 총계', '지출'
    )
    
    print("\n=== 최종 결과 ===")
    print("수입 데이터 개수:", len(sorted_income))
    print("지출 데이터 개수:", len(sorted_expense))
    
    return {
        "income": {"data": sorted_income, "start_row": income_start, "end_row": income_end},
        "expense": {"data": sorted_expense, "start_row": expense_start, "end_row": expense_end}
    }

