import sys
from pathlib import Path

from excel_reader import read_banksalad_sheet
from excel_transform_sheet1 import transform_summary
from excel_writer_sheet1 import save_to_excel
from excel_classify_sheet2 import classify_transactions
from excel_writer_category import add_classified_sheets

def transform_banksalad_excel_sheet1(file_path):
    sheet_name = 0
    df = read_banksalad_sheet(file_path, sheet_name)
    result = transform_summary(df)
    print(result)

    # 정렬된 데이터를 원본 파일에 직접 저장
    if save_to_excel(
        file_path,
        result['income']['data'],
        result['expense']['data'],
        result['income']['start_row'],
        result['income']['end_row'],
        result['expense']['start_row'],
        result['expense']['end_row']
    ):
        print("\n✓ 정렬 및 저장 완료!")
    else:
        print("\n❌ 파일 저장 실패")

    pass

def classify_banksalad_excel_sheet2(file_path):
    sheet_name = 1
    df = read_banksalad_sheet(file_path, sheet_name)
    target_months = ['2025-11', '2025-12']
    classified_data = classify_transactions(df, target_months)
    if add_classified_sheets(file_path, classified_data):
        print("✓ 분류 데이터 추가 완료!")
    else:
        print("❌ 분류 데이터 추가 실패")
    pass

def main():
    if len(sys.argv) < 2:
        print("❌ 입력 엑셀 파일을 인자로 전달해주세요")
        print("예: python main.py input.xlsx")
        sys.exit(1)

    file_path = Path(sys.argv[1])

    if not file_path.exists():
        print(f"❌ 파일이 존재하지 않습니다: {file_path}")
        sys.exit(1)
    
    try:
        transform_banksalad_excel_sheet1(file_path)
        classify_banksalad_excel_sheet2(file_path)
    except Exception as e:
        print(f"❌ 메인함수 실행 중 오류 발생: {e}")

if __name__ == "__main__":
    main()