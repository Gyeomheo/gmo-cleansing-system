import logging
import sys
import glob
import pandas as pd
from pathlib import Path
import warnings

# 경고 메시지 무시 (Excel 로드 시 불필요한 스타일/버전 경고 방지)
warnings.filterwarnings('ignore', category=UserWarning)

try:
    # =========================================================
    # 1. 모듈 임포트 (System Modules)
    # =========================================================
    from src import config
    from src import utils   # 날짜 및 로드 유틸리티
    from src import updater # 마스터 DB 스마트 업데이트
    
    # 파이프라인 로직 로드
    from src.converter import convert_excel_to_json, load_map_from_json
    from src.pipeline import (
        run_cleansing_pipeline, 
        run_ce_product_cleansing, 
        process_mindset_column, 
        process_funding_column, 
        assign_ce_division,
        process_subsidiary_column, # [Step 1] 공백제거
        process_metric_columns,    # [Step 3] 수치표준화 & CPC
        format_mx_data,            # [MX] 최종 포맷터 (Master용)
        format_ce_data,            # [CE] 최종 포맷터 (Master용)
        insert_cleaned_left_of_raw # [NEW] 원본 컬럼 왼쪽에 정제 데이터 삽입
    )
    from src.reporting import create_change_summary, save_to_csv_separated, save_unmapped_reports

except ImportError as e:
    print(f"FATAL: [Import 오류] 필수 모듈을 찾을 수 없습니다. {e}")
    sys.exit(1)

# =========================================================
# 2. 로깅 설정 (Logging Setup)
# =========================================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(config.BASE_DIR / "cleansing_log.txt", mode='w', encoding='utf-8'), 
        logging.StreamHandler(sys.stdout)
    ]
)

def main_workflow():
    logging.info("="*60)
    logging.info(" 🚀 GMO 데이터 자동화 시스템 (Integration Final Build) 가동")
    logging.info("="*60)

    try:
        # =========================================================
        # 1. 매핑 테이블 로드 (Excel -> JSON 캐싱)
        # =========================================================
        logging.info("⚙️ 매핑 테이블 로드 및 초기화 중...")
        
        # Excel 파일을 JSON으로 변환 (속도 최적화)
        convert_excel_to_json(config.MEDIA_MAP_EXCEL, config.MEDIA_MAP_JSON, config.MEDIA_COLS_MAP['key'], config.MEDIA_COLS_MAP['std_cols'])
        convert_excel_to_json(config.PRODUCT_MX_EXCEL, config.PRODUCT_MX_JSON, config.PRODUCT_COLS_MAP_MX['key'], config.PRODUCT_COLS_MAP_MX['std_cols'])
        convert_excel_to_json(config.PRODUCT_CE_EXCEL, config.PRODUCT_CE_JSON, config.PRODUCT_COLS_MAP_CE['key'], config.PRODUCT_COLS_MAP_CE['std_cols'])

        # JSON에서 매핑 데이터프레임 로드
        df_map_media = load_map_from_json(config.MEDIA_MAP_JSON, config.MEDIA_COLS_MAP)
        df_map_mx = load_map_from_json(config.PRODUCT_MX_JSON, config.PRODUCT_COLS_MAP_MX)
        df_map_ce = load_map_from_json(config.PRODUCT_CE_JSON, config.PRODUCT_COLS_MAP_CE)

        # =========================================================
        # 2. 파일 스캔
        # =========================================================
        search_path = str(config.INPUT_DIR / '*.csv') 
        all_raw_files = glob.glob(search_path)
        
        if not all_raw_files:
            logging.warning("📂 1_Input 폴더가 비어있습니다. CSV 파일을 넣어주세요.")
            return
        
        # 마스터 갱신용 데이터 수집 리스트 (최종 포맷 적용된 데이터만 담김)
        collected_mx_data = []
        collected_ce_data = []
        
        # 미매핑 리포트용 집합
        total_unmapped_prod = set() 
        total_unmapped_media = set()    

        # =========================================================
        # 3. 개별 파일 처리 루프
        # =========================================================
        for i, file_path in enumerate(all_raw_files, 1):
            file_name = Path(file_path).name
            logging.info(f"\n📄 [파일 처리 {i}/{len(all_raw_files)}] {file_name}")
            
            try:
                # (A) 로드 (인코딩 자동 감지)
                df_raw = utils.load_csv_safely(file_path)
                if df_raw is None: 
                    logging.error(f"   ❌ 로드 실패: {file_name}")
                    continue

                # (B) 기초 전처리
                # 1. Subsidiary 공백 제거
                df_raw = process_subsidiary_column(df_raw, 'Subsidiary')
                
                # 2. 날짜 필터링 (2025년 이후 데이터만)
                df_raw = utils.process_and_filter_dates(df_raw, 'Date')
                
                if df_raw.empty:
                    logging.warning("   ⚠️ 유효한 날짜(Data)가 없어 스킵합니다.")
                    continue

                # 3. 수치 데이터 표준화 & CPC 계산
                df_raw = process_metric_columns(df_raw)

                # (리포트용) 원본 데이터 보존
                df_raw_original = df_raw.copy()

                # 컬럼명 표준화 (사용자 오타 방지)
                if 'Product Category2' in df_raw.columns:
                    df_raw.rename(columns={'Product Category2': 'Product Category'}, inplace=True)
                if 'Products (optional)' in df_raw.columns:
                    df_raw.rename(columns={'Products (optional)': 'Products'}, inplace=True)

                # (C) 클렌징 파이프라인 시작
                
                # C-1. Media 클렌징 (공통)
                df_cleaned_step1, df_unmapped_media = run_cleansing_pipeline(
                    df_raw, df_map_media, config.MEDIA_COLS_MAP, is_media=True
                )
                total_unmapped_media.update(df_unmapped_media['Unmapped_Key'].tolist())

                # C-2. Product 클렌징 (MX / CE 분기 처리)
                # BU 컬럼 존재 여부로 사업부 판별
                if 'BU' in df_raw.columns:
                    # ---------------------------------------------------
                    # [CASE 1] CE 프로세스
                    # ---------------------------------------------------
                    logging.info("   -> 'BU' 컬럼 감지: CE 로직 적용")
                    df_cleaned_step2, unmapped_prod = run_ce_product_cleansing(
                        df_cleaned_step1, df_map_ce, config.PRODUCT_COLS_MAP_CE
                    )
                    
                    # BU 재할당 로직
                    df_cleaned_final = assign_ce_division(
                        df_cleaned_step2, df_raw_original, config.DIV_RULES, config.AMBIGUOUS_CATS      
                    )
                    
                    # 🛡️ [Safe Filter] CE Smartphones 삭제 (KeyError 방지)
                    cat_col = f"{config.CE_PRODUCT_COLS[0]}_cleaned"
                    if cat_col in df_cleaned_final.columns:
                        mask_ph = (
                            df_cleaned_final[cat_col]
                            .fillna('')
                            .astype(str)
                            .str.strip()
                            .str.lower() == 'smartphones'
                        )
                        if mask_ph.any():
                            dropped = mask_ph.sum()
                            logging.info(f"      ✂️ [Filter] CE 'Smartphones' 데이터 {dropped}행 삭제")
                            df_cleaned_final = df_cleaned_final[~mask_ph].copy()
                    
                    # 🚥 [Data Forking] 마스터용 데이터 생성
                    df_for_master = format_ce_data(df_cleaned_final)
                    
                    if not df_for_master.empty:
                        collected_ce_data.append(df_for_master)

                else:
                    # ---------------------------------------------------
                    # [CASE 2] MX 프로세스
                    # ---------------------------------------------------
                    logging.info("   -> 'BU' 컬럼 없음: MX 로직 적용")
                    df_cleaned_step2, unmapped_prod = run_cleansing_pipeline(
                        df_cleaned_step1, df_map_mx, config.PRODUCT_COLS_MAP_MX
                    )
                    df_cleaned_final = df_cleaned_step2.copy()
                    df_cleaned_final['BU'] = 'MX'
                    
                    # (후처리) MX 전용
                    df_cleaned_final = process_mindset_column(df_cleaned_final, 'Mindset')
                    df_cleaned_final = process_funding_column(df_cleaned_final, 'Funding')

                    # 🚥 [Data Forking] 마스터용 데이터 생성
                    df_for_master = format_mx_data(df_cleaned_final)
                    
                    if not df_for_master.empty:
                        collected_mx_data.append(df_for_master)

                # 공통: 미매핑 키 업데이트
                total_unmapped_prod.update(unmapped_prod['Unmapped_Key'].tolist())
                
                # (D) 개별 파일 저장 (Verification Version)
                # -------------------------------------------------------------
                # ⭐️ [FIX] 원본 순서 유지 + Cleaned 왼쪽 삽입 (Visually Optimized)
                # -------------------------------------------------------------
                df_save = insert_cleaned_left_of_raw(df_cleaned_final.copy())
                
                # 2. 파일명 생성을 위한 날짜 추출
                df_save['Date'] = pd.to_datetime(df_save['Date'], errors='coerce')
                min_date = df_save['Date'].min()
                date_str = min_date.strftime("%m%d") if pd.notnull(min_date) else "Unknown"
                
                file_stem = Path(file_path).stem
                output_filename = f"cleaned_{date_str}~_{file_stem}"
                output_base_path = str(config.OUTPUT_DIR / output_filename)

                # 3. 변경 내역 리포트 생성 및 저장
                summary_prod = create_change_summary(df_raw_original, df_cleaned_final, config.PRODUCT_COLS)
                summary_media = create_change_summary(df_raw_original, df_cleaned_final, config.MEDIA_COLS)
                
                df_to_save = {
                    "Cleaned_Result": df_save,
                    "Summary_Prod": summary_prod,
                    "Summary_Media": summary_media
                }
                save_to_csv_separated(df_to_save, output_base_path)
                logging.info(f"   ✅ 개별 파일 저장 완료 (검증용): {output_filename}")

            except Exception as e_file:
                # 개별 파일 에러가 발생해도 전체 루프는 멈추지 않음 (Resilience)
                logging.error(f"   🚨 [Skip] 파일 처리 중 오류 발생 ({file_name}): {e_file}", exc_info=True)
                continue

        # =========================================================
        # 4. 종료 프로세스 (마스터 갱신)
        # =========================================================
        logging.info("\n💾 [최종 단계] 마스터 DB 동기화 및 리포트 생성")

        # 1. 미매핑 리포트 저장
        save_unmapped_reports(total_unmapped_prod, total_unmapped_media)
        
        # 2. 마스터 DB 스마트 업데이트 (취합된 포맷팅 데이터 사용)
        if collected_mx_data:
            logging.info("🔄 [Updater] MX 마스터 DB 갱신 중...")
            full_mx = pd.concat(collected_mx_data, ignore_index=True)
            updater.update_smart_db(full_mx, "MX")
        else:
            logging.info("ℹ️ 처리된 MX 데이터가 없어 마스터 갱신을 건너뜁니다.")
        
        if collected_ce_data:
            logging.info("🔄 [Updater] CE 마스터 DB 갱신 중...")
            full_ce = pd.concat(collected_ce_data, ignore_index=True)
            updater.update_smart_db(full_ce, "CE")
        else:
            logging.info("ℹ️ 처리된 CE 데이터가 없어 마스터 갱신을 건너뜁니다.")

        logging.info("\n✨ 모든 작업이 성공적으로 완료되었습니다. (2_Output 폴더 확인)")

    except Exception as e:
        # 시스템 전체 레벨의 치명적 오류 (설정 파일 누락 등)
        logging.critical(f"🔥 SYSTEM HALTED: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main_workflow()