import logging
import sys
import pandas as pd
import warnings

# 경고 메시지 무시
warnings.filterwarnings('ignore', category=UserWarning)

try:
    # =========================================================
    # 1. 모듈 임포트
    # =========================================================
    from src import config
    from src import utils   
    from src import updater 
    
    from src.converter import convert_excel_to_json, load_map_from_json
    from src.pipeline import (
        run_cleansing_pipeline, 
        run_ce_product_cleansing, 
        fast_normalize_text,
        process_mindset_column, 
        process_funding_column, 
        assign_ce_division,
        process_subsidiary_column, 
        process_metric_columns,    
        format_mx_data,            
        format_ce_data,            
        insert_cleaned_left_of_raw,
        sanitize_column_headers 
    )
    from src.reporting import create_change_summary, save_to_csv_separated, save_unmapped_reports

except ImportError as e:
    print(f"FATAL: [Import 오류] 필수 모듈을 찾을 수 없습니다. {e}")
    sys.exit(1)

# =========================================================
# 2. 로깅 설정
# =========================================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        # pathlib 경로 사용
        logging.FileHandler(config.BASE_DIR / "cleansing_log.txt", mode='w', encoding='utf-8'), 
        logging.StreamHandler(sys.stdout)
    ]
)


def build_unmapped_product_rows(df_source: pd.DataFrame, unmapped_df: pd.DataFrame, key_source_col: str) -> pd.DataFrame:
    """
    unmapped 키 목록과 원본 제품 컬럼을 조합해,
    C_Key 등록에 바로 사용할 4열 리포트를 생성한다.
    """
    output_cols = [
        'Product_Key_Normalized',
        'Product Category (Raw)',
        'Product Series (Raw)',
        'Products (Raw)'
    ]
    product_cols = ['Product Category', 'Product Series', 'Products']

    if df_source is None or df_source.empty or unmapped_df is None or unmapped_df.empty:
        return pd.DataFrame(columns=output_cols)

    if key_source_col not in df_source.columns or 'Unmapped_Key' not in unmapped_df.columns:
        return pd.DataFrame(columns=output_cols)

    if any(col not in df_source.columns for col in product_cols):
        return pd.DataFrame(columns=output_cols)

    unmapped_keys = set(
        unmapped_df['Unmapped_Key']
        .fillna('')
        .astype(str)
        .str.strip()
        .tolist()
    )
    unmapped_keys.discard('')
    if not unmapped_keys:
        return pd.DataFrame(columns=output_cols)

    source_keys = df_source[key_source_col].fillna('').astype(str).str.strip()
    mask_unmapped = source_keys.isin(unmapped_keys)
    if not mask_unmapped.any():
        return pd.DataFrame(columns=output_cols)

    df_out = df_source.loc[mask_unmapped, product_cols].copy()
    df_out['Product_Key_Normalized'] = fast_normalize_text(df_out['Products'])
    df_out.rename(columns={
        'Product Category': 'Product Category (Raw)',
        'Product Series': 'Product Series (Raw)',
        'Products': 'Products (Raw)'
    }, inplace=True)
    df_out = df_out[output_cols].drop_duplicates().reset_index(drop=True)
    return df_out

def main_workflow():
    logging.info("="*60)
    logging.info(" 🚀 GMO 데이터 자동화 시스템 (Integration Final Build) 가동")
    logging.info("="*60)

    try:
        # =========================================================
        # 1. 매핑 테이블 로드
        # =========================================================
        logging.info("⚙️ 매핑 테이블 로드 및 초기화 중...")
        
        convert_excel_to_json(config.MEDIA_MAP_EXCEL, config.MEDIA_MAP_JSON, config.MEDIA_COLS_MAP['key'], config.MEDIA_COLS_MAP['std_cols'])
        convert_excel_to_json(config.PRODUCT_MX_EXCEL, config.PRODUCT_MX_JSON, config.PRODUCT_COLS_MAP_MX['key'], config.PRODUCT_COLS_MAP_MX['std_cols'])
        convert_excel_to_json(config.PRODUCT_CE_EXCEL, config.PRODUCT_CE_JSON, config.PRODUCT_COLS_MAP_CE['key'], config.PRODUCT_COLS_MAP_CE['std_cols'])

        df_map_media = load_map_from_json(config.MEDIA_MAP_JSON, config.MEDIA_COLS_MAP)
        df_map_mx = load_map_from_json(config.PRODUCT_MX_JSON, config.PRODUCT_COLS_MAP_MX)
        df_map_ce = load_map_from_json(config.PRODUCT_CE_JSON, config.PRODUCT_COLS_MAP_CE)

        # =========================================================
        # 2. 파일 스캔 (Pathlib 활용)
        # =========================================================
        # [Refactor] glob.glob 대신 pathlib.glob 사용
        all_raw_files = list(config.INPUT_DIR.glob('*.csv'))
        
        if not all_raw_files:
            logging.warning(f"📂 {config.INPUT_DIR} 폴더가 비어있습니다. CSV 파일을 넣어주세요.")
            return
        
        collected_mx_data = []
        collected_ce_data = []
        total_unmapped_prod_rows = []
        total_unmapped_media = set()    

        # =========================================================
        # 3. 개별 파일 처리 루프
        # =========================================================
        for i, file_path in enumerate(all_raw_files, 1):
            file_name = file_path.name
            logging.info(f"\n📄 [파일 처리 {i}/{len(all_raw_files)}] {file_name}")
            
            try:
                # (A) 로드
                df_raw = utils.load_csv_safely(file_path)
                if df_raw is None: 
                    logging.error(f"   ❌ 로드 실패: {file_name}")
                    continue

                df_raw = sanitize_column_headers(df_raw)

                # (B) 기초 전처리
                # [Refactor] 상수(Magic String) 제거 적용
                df_raw = process_subsidiary_column(df_raw, config.COL_SUB)
                df_raw = utils.process_and_filter_dates(df_raw, config.COL_DATE)
                
                if df_raw.empty:
                    logging.warning("   ⚠️ 유효한 날짜(Data)가 없어 스킵합니다.")
                    continue

                df_raw = process_metric_columns(df_raw)
                df_raw_original = df_raw.copy()

                # 컬럼명 표준화 (필요시 config에 추가 가능)
                if 'Product Category2' in df_raw.columns:
                    df_raw.rename(columns={'Product Category2': 'Product Category'}, inplace=True)
                if 'Products (optional)' in df_raw.columns:
                    df_raw.rename(columns={'Products (optional)': 'Products'}, inplace=True)

                # (C) 클렌징 파이프라인
                df_cleaned_step1, df_unmapped_media = run_cleansing_pipeline(
                    df_raw, df_map_media, config.MEDIA_COLS_MAP, is_media=True
                )
                total_unmapped_media.update(df_unmapped_media['Unmapped_Key'].tolist())

                # BU 컬럼 존재 여부로 분기 (Config 상수 사용)
                if config.COL_BU in df_raw.columns:
                    # [CASE 1] CE 프로세스
                    logging.info(f"   -> '{config.COL_BU}' 컬럼 감지: CE 로직 적용")
                    df_cleaned_step2, unmapped_prod = run_ce_product_cleansing(
                        df_cleaned_step1, df_map_ce, config.PRODUCT_COLS_MAP_CE
                    )
                    unmapped_key_source_col = 'Product Category'
                    
                    df_cleaned_final = assign_ce_division(
                        df_cleaned_step2, df_raw_original, config.DIV_RULES, config.AMBIGUOUS_CATS      
                    )
                    
                    # 🛡️ [Safe Filter] 키워드 기반 삭제 (Config 활용)
                    cat_col = f"{config.PRODUCT_COLS_MAP_CE['std_cols'][0]}_cleaned" # Product Category_cleaned
                    if cat_col in df_cleaned_final.columns:
                        for ignore_key in config.IGNORE_KEYWORDS:
                            mask_ignore = (
                                df_cleaned_final[cat_col].fillna('').astype(str).str.strip().str.lower() == ignore_key
                            )
                            if mask_ignore.any():
                                logging.info(f"      ✂️ [Filter] '{ignore_key}' 데이터 {mask_ignore.sum()}행 삭제")
                                df_cleaned_final = df_cleaned_final[~mask_ignore].copy()
                    
                    df_for_master = format_ce_data(df_cleaned_final)
                    if not df_for_master.empty: collected_ce_data.append(df_for_master)

                else:
                    # [CASE 2] MX 프로세스
                    logging.info(f"   -> '{config.COL_BU}' 컬럼 없음: MX 로직 적용")
                    df_cleaned_step2, unmapped_prod = run_cleansing_pipeline(
                        df_cleaned_step1, df_map_mx, config.PRODUCT_COLS_MAP_MX
                    )
                    unmapped_key_source_col = 'Products'
                    df_cleaned_final = df_cleaned_step2.copy()
                    df_cleaned_final[config.COL_BU] = 'MX'
                    
                    df_cleaned_final = process_mindset_column(df_cleaned_final, 'Mindset')
                    df_cleaned_final = process_funding_column(df_cleaned_final, 'Funding')

                    df_for_master = format_mx_data(df_cleaned_final)
                    if not df_for_master.empty: collected_mx_data.append(df_for_master)

                # 공통: 미매핑 제품 리포트 업데이트 (정규화 Key + Raw 3열)
                df_unmapped_prod_rows = build_unmapped_product_rows(
                    df_source=df_cleaned_step2,
                    unmapped_df=unmapped_prod,
                    key_source_col=unmapped_key_source_col
                )
                if not df_unmapped_prod_rows.empty:
                    total_unmapped_prod_rows.append(df_unmapped_prod_rows)
                
                # (D) 개별 파일 저장
                df_save = insert_cleaned_left_of_raw(df_cleaned_final.copy())
                
                # 파일명 생성
                df_save[config.COL_DATE] = pd.to_datetime(df_save[config.COL_DATE], errors='coerce')
                min_date = df_save[config.COL_DATE].min()
                date_str = min_date.strftime("%m%d") if pd.notnull(min_date) else "Unknown"
                
                file_stem = file_path.stem
                output_filename = f"cleaned_{date_str}~_{file_stem}"
                output_base_path = str(config.OUTPUT_DIR / output_filename)

                summary_prod = create_change_summary(df_raw_original, df_cleaned_final, config.PRODUCT_COLS)
                summary_media = create_change_summary(df_raw_original, df_cleaned_final, config.MEDIA_COLS)
                
                df_to_save = {
                    "Cleaned_Result": df_save,
                    "Summary_Prod": summary_prod,
                    "Summary_Media": summary_media
                }
                save_to_csv_separated(df_to_save, output_base_path)
                logging.info(f"   ✅ 개별 파일 저장 완료: {output_filename}")

            except Exception as e_file:
                logging.error(f"   🚨 [Skip] 파일 처리 중 오류 발생 ({file_name}): {e_file}", exc_info=True)
                continue

        # =========================================================
        # 4. 종료 프로세스
        # =========================================================
        logging.info("\n💾 [최종 단계] 마스터 DB 동기화 및 리포트 생성")

        if total_unmapped_prod_rows:
            total_unmapped_prod = pd.concat(total_unmapped_prod_rows, ignore_index=True).drop_duplicates()
        else:
            total_unmapped_prod = pd.DataFrame(columns=[
                'Product_Key_Normalized',
                'Product Category (Raw)',
                'Product Series (Raw)',
                'Products (Raw)'
            ])

        save_unmapped_reports(total_unmapped_prod, total_unmapped_media)
        
        if collected_mx_data:
            logging.info("🔄 [Updater] MX 마스터 DB 갱신 중...")
            full_mx = pd.concat(collected_mx_data, ignore_index=True)
            updater.update_smart_db(full_mx, "MX")
        
        if collected_ce_data:
            logging.info("🔄 [Updater] CE 마스터 DB 갱신 중...")
            full_ce = pd.concat(collected_ce_data, ignore_index=True)
            updater.update_smart_db(full_ce, "CE")

        logging.info(f"\n✨ 모든 작업 완료. 결과 폴더: {config.OUTPUT_DIR}")

    except Exception as e:
        logging.critical(f"🔥 SYSTEM HALTED: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main_workflow()
