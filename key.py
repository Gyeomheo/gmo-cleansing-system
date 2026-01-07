import pandas as pd
import logging
import os

# --- 로깅 설정 ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def load_excel_safely(file_name, engine='calamine'):
    """
    (v3) '.xlsx' 파일을 'calamine' 엔진으로 직접 읽어옵니다.
    """
    try:
        logging.info(f"'{engine}' 엔진으로 '{file_name}' 엑셀 파일 로드 시도...")
        return pd.read_excel(file_name, engine=engine)
    except ImportError:
        logging.warning("  -> 'calamine' 엔진을 찾을 수 없습니다.")
        logging.warning("  -> 'openpyxl' 기본 엔진으로 느리게 재시도...")
        return pd.read_excel(file_name, engine='openpyxl')
    except Exception as e:
        logging.error(f"'{file_name}' 엑셀 파일 로드 중 오류 발생: {e}")
        return None

def add_id_and_save_as_csv(): # ⭐️ 함수 이름 변경
    
    input_excel_original = 'SEG_For_TEST.xlsx'
    
    # ⭐️ (변경) 출력 파일 확장자를 .csv로 변경
    output_csv = 'SEG_For_TEST_with_ID.csv' 

    if not os.path.exists(input_excel_original):
        logging.error(f"FATAL: 원본 엑셀 파일 '{input_excel_original}'을(를) 찾을 수 없습니다.")
        return

    # 1. Raw 엑셀 파일 로드 (11초)
    logging.info(f"'{input_excel_original}' 파일 로드 중...")
    df_raw = load_excel_safely(input_excel_original)
    
    if df_raw is None:
        logging.error("파일 로드 실패. 작업을 중단합니다.")
        return

    # 2. 'UniqueID' 열 생성 (1초)
    logging.info("'UniqueID' 열 생성 중...")
    df_raw['UniqueID'] = range(len(df_raw))

    # 3. 'UniqueID' 열을 맨 앞으로 이동 (1초)
    cols = ['UniqueID'] + [col for col in df_raw.columns if col != 'UniqueID']
    df_raw = df_raw[cols]

    # 4. ⭐️ (변경) 새 CSV 파일로 '초고속' 저장
    logging.info(f"'{output_csv}' 파일로 (빠르게) 저장 중...")
    try:
        df_raw.to_csv(output_csv, index=False, encoding='utf-8-sig')
        logging.info(f"✅ 완료! '{output_csv}' 파일이 (빠르게) 생성되었습니다.")
    except Exception as e:
        logging.error(f"CSV 저장 중 오류 발생: {e}")

# --- 메인 실행 ---
if __name__ == "__main__":
    add_id_and_save_as_csv() # ⭐️ 변경된 함수 이름 호출