import pandas as pd

def read_banksalad_sheet(file_path, sheet_name):
    """
    엑셀의 특정 시트의 데이터를 읽어옵니다. 
    서식 유지를 위해 skiprows는 사용하지 않습니다.
    """
    try:
        # 특정 시트를 읽어옴
        # header=None으로 읽어야 상단 타이틀까지 포함된 전체 구조 파악이 쉽습니다.
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        return df
    except Exception as e:
        print(f"파일을 읽는 중 오류 발생: '{file_path}'의 '{sheet_name}' 시트를 읽는 중 오류 발생: {e}")
        return None