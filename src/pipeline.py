"""
[Pipeline Logic - Real Final Complete]
- Core: Regex Pre-compilation, Vectorized Normalization
- MX: Priority Scoring System, Multi-Pattern Rescue
- CE: Unique Combination Optimization (with Bespoke Safety Fix)
- Utils: Metric Calculation, Division Assignment (Force Upper Fix)
- Formatters: Final Output Standardization
- Verification: Smart Column Injection (Cleaned Left of Raw)
- Safety: Auto-Sanitize Column Headers & Smart Alias Mapping (New!)
"""

import pandas as pd
import numpy as np
import logging
import re
from typing import Tuple
from . import config

# =========================================================
# 0. [NEW] 컬럼명 위생 처리 (줄바꿈/따옴표 제거)
# =========================================================
def sanitize_column_headers(df: pd.DataFrame) -> pd.DataFrame:
    """
    1. 줄바꿈(\n) -> 공백 변환
    2. 따옴표(") 제거
    3. 다중 공백('  ') -> 단일 공백(' ') 표준화 (핵심!)
    4. 중복 컬럼 자동 제거 (먼저 나온 정상 컬럼 유지)
    """
    
    # 1. 모든 컬럼을 문자열로 변환
    new_cols = df.columns.astype(str)

    new_cols = new_cols.str.replace('\ufeff', '', regex=False)
    
    # 2. 줄바꿈과 캐리지 리턴을 공백으로 변경 (Regex 사용)
    new_cols = new_cols.str.replace(r'[\r\n]+', ' ', regex=True)
    
    # 3. 따옴표 제거
    new_cols = new_cols.str.replace('"', '', regex=False)
    new_cols = new_cols.str.replace("'", '', regex=False)
    
    # 4. [핵심] 공백이 2개 이상인 곳을 1개로 압축 (Normalization)
    new_cols = new_cols.str.replace(r'\s+', ' ', regex=True)
    
    # 5. 양쪽 공백 제거
    new_cols = new_cols.str.strip()
    
    # 6. 컬럼명 적용
    df.columns = new_cols
    
    # 7. [핵심] 이름이 같아진 중복 컬럼 제거 (앞에 있는 정상 데이터 우선)
    # 예: 원본 'Media Spend (USD)'와 청소된 'Media Spend (USD)'가 충돌하면 하나만 남김
    df = df.loc[:, ~df.columns.duplicated()]
    
    return df

# =========================================================
# 1. 정규식 패턴 전역 컴파일
# =========================================================
P_CAMEL_1 = re.compile(r'([a-z])([A-Z])')
P_CAMEL_2 = re.compile(r'([A-Z])([A-Z][a-z])')
P_DIGIT_CHAR = re.compile(r'(\d+)([a-zA-Z]+)')

def fast_normalize_text(series: pd.Series) -> pd.Series:
    """
    [성능 최적화] 데이터 정규화 (CamelCase 분리 및 Title Case 적용)
    """
    s = series.fillna('').astype(str).str.strip()
    s = s.str.replace(P_CAMEL_1, r'\1 \2', regex=True)
    s = s.str.replace(P_CAMEL_2, r'\1 \2', regex=True)
    s = s.str.replace(P_DIGIT_CHAR, r'\1 \2', regex=True)
    
    def _case_adjust(val):
        if not val: return ''
        if val.islower() or val.isupper():
            return val.title()
        return val
    
    return s.map(_case_adjust)

# =========================================================
# (A) MX & Media 파이프라인
# =========================================================
def run_cleansing_pipeline(df_raw: pd.DataFrame, df_map: pd.DataFrame, map_cols: dict, is_media: bool = False) -> Tuple[pd.DataFrame, pd.DataFrame]:
    # ⭐️ [Safety] 로직 시작 전 컬럼명 청소
    df_raw = sanitize_column_headers(df_raw)

    raw_cols = map_cols['raw_cols'] 
    std_cols = map_cols['std_cols'] 
    key_col = map_cols['key']
    normalize_cols = map_cols.get('normalize_cols', raw_cols) 
    
    A_raw, B_raw, C_raw = raw_cols[0], raw_cols[1], raw_cols[2]
    A_std, B_std, C_std = std_cols[0], std_cols[1], std_cols[2]
    A_norm, B_norm, C_norm = f"_norm_{A_raw}", f"_norm_{B_raw}", f"_norm_{C_raw}"

    logging.info(f"  -> (Pipeline) '{C_raw}' 기준 클렌징 (Vectorized)...")
    
    df_proc = df_raw.reset_index().rename(columns={'index': 'original_index'})
    df_proc[A_norm] = df_proc[A_raw]
    df_proc[B_norm] = df_proc[B_raw]
    df_proc[C_norm] = df_proc[C_raw]

    for col, norm_col in zip(raw_cols, [A_norm, B_norm, C_norm]):
        if col in normalize_cols:
            df_proc[norm_col] = fast_normalize_text(df_proc[norm_col])
        else:
            df_proc[norm_col] = df_proc[norm_col].fillna('').astype(str).str.strip()
            
    if is_media:
        df_proc['_v_key'] = df_proc[C_norm] 
        mask_affiliate = (df_proc[A_norm] == 'Affiliate')
        df_proc.loc[mask_affiliate, '_v_key'] = df_proc.loc[mask_affiliate, B_norm]
        merge_left_key = '_v_key'
    else:
        merge_left_key = C_norm 

    df_merged = pd.merge(df_proc, df_map, left_on=merge_left_key, right_on=key_col, how='left')
    
    # Scoring Logic
    norm_a_lower = df_merged[A_norm].str.lower()
    std_a_lower = df_merged[A_std].str.lower()
    norm_b_lower = df_merged[B_norm].str.lower()
    std_b_lower = df_merged[B_std].str.lower()
    norm_c_lower = df_merged[C_norm].str.lower()
    std_c_lower = df_merged[C_std].str.lower()

    a_match = ((df_merged[A_std] == 'NONE') | ((df_merged[A_std] == 'ANY') & (df_merged[A_norm] != '')) | (norm_a_lower == std_a_lower))
    b_match = ((df_merged[B_std] == 'NONE') | ((df_merged[B_std] == 'ANY') & (df_merged[B_norm] != '')) | (norm_b_lower == std_b_lower))
    c_match = ((df_merged[C_std] == 'NONE') | ((df_merged[C_std] == 'ANY') & (df_merged[C_norm] != '')) | (norm_c_lower == std_c_lower))

    if is_media:
        is_aff = (df_merged[A_norm] == 'Affiliate')
        a_match = ((df_merged[A_std] == 'NONE') & is_aff) | a_match
        b_match = ((df_merged[B_std] == 'NONE') & is_aff) | b_match
        c_match = ((df_merged[C_std] == 'NONE') & is_aff) | c_match

    conditions = [
        a_match & b_match & c_match, 
        a_match & b_match,           
        b_match,                    
        df_merged[key_col].notna()   
    ]
    scores = [4, 3, 2, 1] 
    df_merged['priority'] = np.select(conditions, scores, default=0)

    df_merged.sort_values(by=['original_index', 'priority'], ascending=[True, False], inplace=True)
    df_final_rows = df_merged.drop_duplicates(subset=['original_index'], keep='first')
    
    df_final = df_final_rows.copy()
    A_cln, B_cln, C_cln = f"{A_raw}_cleaned", f"{B_raw}_cleaned", f"{C_raw}_cleaned"
    
    df_final[A_cln] = df_final[A_std]
    df_final[B_cln] = df_final[B_std]
    df_final[C_cln] = df_final[C_std]
    
    if is_media:
        mask_aff = (df_final[A_std] == 'Affiliate')
        df_final.loc[mask_aff, C_cln] = df_final.loc[mask_aff, C_raw]
        fail_mask = (df_final['priority'] == 0)
        df_final.loc[fail_mask, [A_cln, B_cln, C_cln]] = df_final.loc[fail_mask, [A_raw, B_raw, C_raw]].values
        mask_aff_fail = (df_final[A_norm] == 'Affiliate') & fail_mask
        if mask_aff_fail.any():
            df_final.loc[mask_aff_fail, B_cln] = 'Others'
            df_final.loc[mask_aff_fail, C_cln] = df_final.loc[mask_aff_fail, C_raw]

    # MX Rescue Logic (Multi)
    if not is_media: 
        mask_multi = (df_final[A_cln] == 'Multi') | (df_final[A_cln].isna())
        if mask_multi.any():
            logging.info(f"     -> (Optimization) 'Multi' {mask_multi.sum()}건 구제 시도...")
            fallback_map = {}
            series_reverse_map = {}
            for row in df_map[[A_std, B_std]].itertuples(index=False):
                a_val, b_val = row[0], row[1]
                if pd.isna(a_val) or pd.isna(b_val): continue
                s_low = str(b_val).lower().strip()
                if s_low not in ['multi', 'others', 'other', 'nan', 'none', '']:
                    series_reverse_map[s_low] = (a_val, b_val)
                cat_key = str(a_val).lower().strip()
                if cat_key not in fallback_map: fallback_map[cat_key] = (a_val, 'Multi')
                current_fb = fallback_map[cat_key][1]
                if 'other' in s_low and 'other' not in str(current_fb).lower():
                    fallback_map[cat_key] = (a_val, b_val)

            target_data = df_final.loc[mask_multi, [A_norm, B_norm]].copy()
            target_data['cat_lower'] = target_data[A_norm].astype(str).str.lower().str.strip()
            target_data['ser_lower'] = target_data[B_norm].astype(str).str.lower().str.strip()
            
            unique_patterns = target_data[['cat_lower', 'ser_lower']].drop_duplicates()
            pattern_results = []
            
            for row in unique_patterns.itertuples(index=False):
                c_low, s_low = row.cat_lower, row.ser_lower
                res_a, res_b = None, None
                if s_low in series_reverse_map: res_a, res_b = series_reverse_map[s_low]
                elif c_low in fallback_map: res_a, res_b = fallback_map[c_low]
                elif c_low in series_reverse_map: res_a, res_b = series_reverse_map[c_low]
                if res_a: pattern_results.append({'cat_lower': c_low, 'ser_lower': s_low, 'A_new': res_a, 'B_new': res_b})
            
            if pattern_results:
                df_pattern_map = pd.DataFrame(pattern_results)
                df_rescued = pd.merge(target_data.reset_index(), df_pattern_map, on=['cat_lower', 'ser_lower'], how='left')
                mask_success = df_rescued['A_new'].notna()
                success_indices = df_rescued.loc[mask_success, 'index'].values
                df_final.loc[success_indices, A_cln] = df_rescued.loc[mask_success, 'A_new'].values
                df_final.loc[success_indices, B_cln] = df_rescued.loc[mask_success, 'B_new'].values
                df_final.loc[success_indices, C_cln] = df_rescued.loc[mask_success, 'B_new'].values

    # Final Cleanup
    fail_mask_final = df_final[A_cln].isna()
    if fail_mask_final.any():
        df_final.loc[fail_mask_final, [A_cln, B_cln, C_cln]] = df_final.loc[fail_mask_final, [A_raw, B_raw, C_raw]].values

    valid_combinations = set((df_map[A_std] + "|" + df_map[B_std] + "|" + df_map[C_std]).unique())
    concat_raw = (df_proc[A_raw].fillna('') + "|" + df_proc[B_raw].fillna('') + "|" + df_proc[C_raw].fillna(''))
    is_already_valid = concat_raw.isin(valid_combinations)
    valid_indices = df_proc.loc[is_already_valid, 'original_index']
    mask_valid_in_final = df_final['original_index'].isin(valid_indices)
    
    if mask_valid_in_final.any():
        df_final.loc[mask_valid_in_final, [A_cln, B_cln, C_cln]] = df_final.loc[mask_valid_in_final, [A_raw, B_raw, C_raw]].values

    unmapped_key_source = merge_left_key if is_media else C_raw
    is_rescued = (df_final[A_cln] != 'Multi') & (df_final[A_cln].notna()) & (df_final['priority'] == 0)
    real_fail_mask = (df_final['priority'] == 0) & (~is_rescued) & (~mask_valid_in_final)
    
    failed_indices = df_final.loc[real_fail_mask, 'original_index'].values
    unmapped_values = df_proc[df_proc['original_index'].isin(failed_indices)][unmapped_key_source].unique()
    df_unmapped_report = pd.DataFrame(unmapped_values, columns=['Unmapped_Key'])

    df_final.set_index('original_index', inplace=True)
    df_final.sort_index(inplace=True)
    
    all_desired_cols = list(df_raw.columns) + [A_cln, B_cln, C_cln]
    cols_to_keep = [col for col in all_desired_cols if col in df_final.columns]
    
    return df_final[cols_to_keep].copy(), df_unmapped_report

# =========================================================
# (B) CE 파이프라인 (⭐️ Bespoke Safety Fix Applied)
# =========================================================
def run_ce_product_cleansing(df_raw: pd.DataFrame, df_map: pd.DataFrame, map_cols: dict) -> Tuple[pd.DataFrame, pd.DataFrame]:
    # ⭐️ [Safety] 로직 시작 전 컬럼명 청소
    df_raw = sanitize_column_headers(df_raw)

    raw_cols = map_cols['raw_cols'] 
    std_cols = map_cols['std_cols'] 
    normalize_cols = map_cols.get('normalize_cols', raw_cols)
    Raw_Cat_Col, Raw_Ser_Col = raw_cols[0], raw_cols[1]
    Std_Cat_Col, Std_Ser_Col = std_cols[0], std_cols[1]
    
    df_proc = df_raw.reset_index().rename(columns={'index': 'original_index'})
    Norm_Cat_Col, Norm_Ser_Col = f"_norm_{Raw_Cat_Col}", f"_norm_{Raw_Ser_Col}"
    
    df_proc[Norm_Cat_Col] = df_proc[Raw_Cat_Col]
    df_proc[Norm_Ser_Col] = df_proc[Raw_Ser_Col]
    
    for col, norm_col in zip([Raw_Cat_Col, Raw_Ser_Col], [Norm_Cat_Col, Norm_Ser_Col]):
        if col in normalize_cols: 
            df_proc[norm_col] = fast_normalize_text(df_proc[norm_col])
        else:
            df_proc[norm_col] = df_proc[norm_col].fillna('').astype(str).str.strip()
    
    key_col = map_cols['key']
    valid_categories_map = {str(k).lower().strip(): v for k, v in zip(df_map[key_col], df_map[Std_Cat_Col])}
    
    series_map_by_cat = {}
    fallback_map_by_cat = {}
    global_series_lookup = {}
    seen_series_counts = {}  # ⭐️ [Safety] 시리즈 등장 횟수 카운트용
    COMMON_KEYWORDS = ['multi', 'others', 'other', 'nan', 'none', '']
    
    # ---------------------------------------------------------
    # 1. Series Frequency Count (중복 방지용 사전 집계)
    # ---------------------------------------------------------
    for cat in df_map[Std_Cat_Col].unique():
        valid_series = df_map.loc[df_map[Std_Cat_Col] == cat, Std_Ser_Col].unique()
        for s in valid_series:
            s_low = str(s).lower().strip()
            if s_low not in COMMON_KEYWORDS:
                seen_series_counts[s_low] = seen_series_counts.get(s_low, 0) + 1

    # ---------------------------------------------------------
    # 2. Build Lookup Tables (중복 시리즈 제외)
    # ---------------------------------------------------------
    for cat in df_map[Std_Cat_Col].unique():
        cat_key = str(cat).lower()
        valid_series = df_map.loc[df_map[Std_Cat_Col] == cat, Std_Ser_Col].unique()
        
        # 카테고리별 시리즈 맵
        series_map_by_cat[cat_key] = {str(s).lower(): s for s in valid_series}
        
        # Fallback 로직
        s_lower_list = [str(s).lower() for s in valid_series]
        if 'others' in s_lower_list: fallback_map_by_cat[cat_key] = next(s for s in valid_series if str(s).lower() == 'others')
        elif 'multi' in s_lower_list: fallback_map_by_cat[cat_key] = next(s for s in valid_series if str(s).lower() == 'multi')
        else: fallback_map_by_cat[cat_key] = 'Multi'
            
        # Global Lookup 등록 (유니크한 경우만)
        for s in valid_series:
            s_low = str(s).lower()
            if s_low not in COMMON_KEYWORDS:
                # ⭐️ [Safety Fix] 오직 한 카테고리에만 존재하는 경우에만 족보 등록
                if seen_series_counts.get(s_low, 0) == 1:
                    global_series_lookup[s_low] = (cat, s)
            
    Res_Cat_Col, Res_Ser_Col, Res_Prd_Col = f"{Raw_Cat_Col}_cleaned", f"{Raw_Ser_Col}_cleaned", f"{raw_cols[2]}_cleaned"
    
    df_proc['cat_lower'] = df_proc[Norm_Cat_Col].astype(str).str.lower().str.strip()
    df_proc['ser_lower'] = df_proc[Norm_Ser_Col].astype(str).str.lower().str.strip()
    df_proc['matched_cat'] = df_proc['cat_lower'].map(valid_categories_map)
    unique_combinations = df_proc[['matched_cat', 'ser_lower']].drop_duplicates()
    unique_results = []
    
    for row in unique_combinations.itertuples(index=False):
        cat_std = row.matched_cat
        raw_ser_lower = row.ser_lower
        res_cat, res_ser = None, None
        
        if pd.isna(cat_std):
            if raw_ser_lower in global_series_lookup: res_cat, res_ser = global_series_lookup[raw_ser_lower]
            else: res_cat, res_ser = 'NON_CE_CATEGORY', 'NON_CE_CATEGORY'
        else:
            cat_key = str(cat_std).lower()
            series_map = series_map_by_cat.get(cat_key, {})
            if raw_ser_lower in series_map: res_cat, res_ser = cat_std, series_map[raw_ser_lower]
            else:
                if raw_ser_lower in global_series_lookup: res_cat, res_ser = global_series_lookup[raw_ser_lower]
                else: res_cat, res_ser = cat_std, fallback_map_by_cat.get(cat_key, 'Multi')
        unique_results.append({'matched_cat': cat_std, 'ser_lower': raw_ser_lower, 'res_cat_val': res_cat, 'res_ser_val': res_ser})
        
    df_lookup = pd.DataFrame(unique_results)
    df_proc = pd.merge(df_proc, df_lookup, on=['matched_cat', 'ser_lower'], how='left')
    df_proc[Res_Cat_Col] = df_proc['res_cat_val']
    df_proc[Res_Ser_Col] = df_proc['res_ser_val']
    df_proc[Res_Prd_Col] = np.where(df_proc[Res_Cat_Col] == 'NON_CE_CATEGORY', 'NON_CE_CATEGORY', df_proc[raw_cols[2]].fillna(''))
    
    mask_fail = (df_proc[Res_Cat_Col] == 'NON_CE_CATEGORY')
    if mask_fail.any():
        df_proc.loc[mask_fail, Res_Cat_Col] = df_proc.loc[mask_fail, Raw_Cat_Col]
        df_proc.loc[mask_fail, Res_Ser_Col] = df_proc.loc[mask_fail, Raw_Ser_Col]
    
    unmapped_rows = df_proc.loc[mask_fail, Raw_Cat_Col].unique()
    df_unmapped_report = pd.DataFrame(unmapped_rows, columns=['Unmapped_Key'])
    
    df_final = df_proc.set_index('original_index')
    df_final.sort_index(inplace=True)
    all_desired_cols = list(df_raw.columns) + [Res_Cat_Col, Res_Ser_Col, Res_Prd_Col]
    cols_to_keep = [col for col in all_desired_cols if col in df_final.columns]
    
    return df_final[cols_to_keep].copy(), df_unmapped_report

# =========================================================
# (C) 유틸리티 (수정된 핵심 로직 포함)
# =========================================================

def process_subsidiary_column(df: pd.DataFrame, col_name: str = 'Subsidiary') -> pd.DataFrame:
    """ [Step 1] Subsidiary 컬럼 앞뒤 공백 제거 """
    if col_name in df.columns:
        df[col_name] = df[col_name].fillna('').astype(str).str.strip()
    return df

def process_mindset_column(df, col_name='Mindset'):
    logging.info(f"  -> (v20) Mindset 클렌징...")
    if col_name not in df.columns: return df
    df[col_name] = df[col_name].fillna('').astype(str)
    mask_cold = (df[col_name].str.strip().str.lower() == 'cold')
    df.loc[mask_cold, col_name] = 'Cold' 
    return df

def process_funding_column(df, col_name='Funding'):
    logging.info(f"  -> (v21) Funding 클렌징...")
    if col_name not in df.columns: return df
    temp_val = df[col_name].fillna('').astype(str).str.strip().str.lower()
    cond_gmo = (temp_val == 'gmo')
    cond_local = (temp_val == 'local')
    df[col_name] = np.select([cond_gmo, cond_local], ['GMO', 'Local'], default=df[col_name])
    return df

# -------------------------------------------------------------------------
# ⭐️ [핵심 수정] BU 강제 대문자화 (Force Uppercase Logic)
# -------------------------------------------------------------------------
def assign_ce_division(df_cleaned: pd.DataFrame, df_raw: pd.DataFrame, div_rules: dict, ambiguous_cats: list) -> pd.DataFrame:
    """
    BU(Division) 할당 로직.
    [중요] 입력 데이터의 BU(da, vd 등)를 무조건 대문자(DA, VD)로 변환한 후 시작합니다.
    """
    logging.info(f"  -> (v41) BU(Division) 업데이트 및 대소문자 강제 보정...")
    df_final = df_cleaned.copy()

    # 1. [FIX] Raw BU 값을 가져와서 무조건 대문자로 초기화
    # 이렇게 하면 'da' -> 'DA', 'vd' -> 'VD'가 기본값으로 확정됩니다.
    if 'BU' in df_raw.columns:
        df_final['BU'] = df_raw['BU'].fillna('').astype(str).str.strip().str.upper()
    else:
        df_final['BU'] = 'MX' # BU 컬럼이 없으면 기본 MX

    # 2. Product Category 기반으로 BU 덮어쓰기 (Rule-based Override)
    # 이미 1번에서 대문자로 깔려있으므로, 룰에 매칭되지 않는 녀석들도 자동으로 대문자가 유지됩니다.
    cat_col = next((c for c in df_final.columns if 'Product Category' in c and c.endswith('_cleaned')), None)
    if not cat_col: return df_final
    
    cat_lower = df_final[cat_col].fillna('').astype(str).str.lower().str.strip()
    
    for div_name, cat_list in div_rules.items():
        mask = cat_lower.isin([c.lower() for c in cat_list])
        df_final.loc[mask, 'BU'] = div_name
            
    # 모호한 카테고리는 Raw BU 유지 (이미 대문자화 되어 있음)

        
    return df_final

def process_metric_columns(df: pd.DataFrame) -> pd.DataFrame:
    logging.info(f"   -> (Metrics) 컬럼명 표준화 및 수치 처리...")
    
    # 1. 일단 컬럼명 위생 처리 (줄바꿈 제거)
    df = sanitize_column_headers(df)

    # 2. ⭐️ [Safety Map] 변종 이름들을 표준 이름으로 '납치'해옵니다.
    # 왼쪽(변종)이 발견되면 오른쪽(표준)으로 즉시 바꿉니다.
    rename_map = {
        # Revenue 관련 변종들
        'Revenue (USD)': 'Revenue',
        'Revenue(USD)': 'Revenue',
        'Total Revenue': 'Revenue',
        'Rev.': 'Revenue',
        
        # Spend 관련 변종들
        'Media Spend (USD)': 'Media Spend (USD)', # 이미 표준이지만 명시
        'Spend': 'Media Spend (USD)',
        'Cost': 'Media Spend (USD)',
        
        # 기타 지표
        'Orders (Count)': 'Orders',
        'App Installs': 'App Install',
        'Installs': 'App Install',
        'App Install (Count)': 'App Install'
    }
    
    # 실제 존재하는 컬럼만 골라서 변경 (에러 방지)
    actual_rename = {k: v for k, v in rename_map.items() if k in df.columns}
    if actual_rename:
        logging.info(f"      🔧 컬럼명 정규화 적용: {actual_rename}")
        df.rename(columns=actual_rename, inplace=True)

    # 3. 수치 변환 로직
    target_metrics = ['Impressions', 'Clicks', 'Media Spend (USD)', 'Media Spend', 'Orders', 'Revenue', 'App Install']
    for col in target_metrics:
        if col in df.columns:
            s = df[col].astype(str).str.strip().str.replace(',', '')
            s = s.replace(['-', '–', '—'], '0') # 대시 기호 처리
            df[col] = pd.to_numeric(s, errors='coerce').fillna(0)
            
    # 4. CPC 재계산
    cost_col = next((c for c in ['Media Spend (USD)', 'Media Spend'] if c in df.columns), None)
    if cost_col and 'Clicks' in df.columns:
        df['CPC'] = np.where(df['Clicks'] > 0, df[cost_col] / df['Clicks'], 0)
        df['CPC'] = df['CPC'].replace([np.inf, -np.inf], 0).fillna(0)
        
    return df

# =========================================================
# (D) 최종 포맷터
# =========================================================
def format_mx_data(df: pd.DataFrame) -> pd.DataFrame:
    df_out = df.copy()
    if 'Media Type 2' in df_out.columns: df_out['Media Type 2 (Raw)'] = df_out['Media Type 2']
    if 'Media Platform' in df_out.columns: df_out['Media Platform (Raw)'] = df_out['Media Platform']

    overwrite_rules = {
        'Media Type 1_cleaned': 'Media Type 1', 'Media Type 2_cleaned': 'Media Type 2',
        'Media Platform_cleaned': 'Media Platform', 'Product Category_cleaned': 'Product Category',
        'Product Series_cleaned': 'Product Series', 'Products_cleaned': 'Products'
    }
    for clean_col, target_col in overwrite_rules.items():
        if clean_col in df_out.columns: df_out[target_col] = df_out[clean_col]

    final_cols = config.MX_OUTPUT_COLS
    for col in final_cols:
        if col not in df_out.columns: df_out[col] = None
    return df_out[final_cols]

def format_ce_data(df: pd.DataFrame) -> pd.DataFrame:
    df_out = df.copy()
    if 'Media Type 2' in df_out.columns: df_out['Media Type 2 (Raw)'] = df_out['Media Type 2']
    if 'Media Platform' in df_out.columns: df_out['Media Platform (Raw)'] = df_out['Media Platform']

    overwrite_rules = {
        'Media Type 1_cleaned': 'Media Type 1', 'Media Type 2_cleaned': 'Media Type 2',
        'Media Platform_cleaned': 'Media Platform', 'Product Category_cleaned': 'Product Category',
        'Product Series_cleaned': 'Product Series', 'Products_cleaned': 'Products (Optional)'
    }
    for clean_col, target_col in overwrite_rules.items():
        if clean_col in df_out.columns: df_out[target_col] = df_out[clean_col]
            
    if 'Products' in df_out.columns and 'Products (Optional)' not in df_out.columns:
        df_out['Products (Optional)'] = df_out['Products']

    final_cols = config.CE_OUTPUT_COLS
    for col in final_cols:
        if col not in df_out.columns: df_out[col] = None
    return df_out[final_cols]

# =========================================================
# (E) 검증용 컬럼 자동 삽입
# =========================================================
def insert_cleaned_left_of_raw(df: pd.DataFrame) -> pd.DataFrame:
    all_cols = list(df.columns)
    cleaned_set = {c for c in all_cols if c.endswith('_cleaned')}
    base_cols = [c for c in all_cols if c not in cleaned_set]
    final_order = []
    
    for col in base_cols:
        target_clean = f"{col}_cleaned"
        if target_clean in cleaned_set:
            final_order.append(target_clean)
            final_order.append(col)
            if target_clean in cleaned_set: cleaned_set.remove(target_clean)
        else:
            final_order.append(col)
            
    if cleaned_set: final_order.extend(list(cleaned_set))
    return df[final_order]