import pandas as pd
import numpy as np
import logging
import re
from typing import Tuple, Dict, List
from . import config

# =========================================================
# 0. 유틸리티 (Vectorized & Ruff Optimized)
# =========================================================

def sanitize_column_headers(df: pd.DataFrame) -> pd.DataFrame:
    """컬럼명 위생 처리: BOM, 줄바꿈, 중복 제거"""
    if df.columns.empty:
        return df
    new_cols = (
        df.columns.astype(str)
        .str.replace("\ufeff", "", regex=False)
        .str.replace(r"[\r\n]+", " ", regex=True)
        .str.replace(r"['\"]", "", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    df.columns = new_cols
    return df.loc[:, ~df.columns.duplicated()]


def fast_normalize_text(series: pd.Series) -> pd.Series:
    """텍스트 정규화: CamelCase 분리 및 Title Case 적용"""
    s = series.fillna("").astype(str).str.strip()
    p1 = re.compile(r"([a-z])([A-Z])")
    p2 = re.compile(r"([A-Z])([A-Z][a-z])")
    p3 = re.compile(r"(\d+)([a-zA-Z]+)")

    s = s.str.replace(p1, r"\1 \2", regex=True)
    s = s.str.replace(p2, r"\1 \2", regex=True)
    s = s.str.replace(p3, r"\1 \2", regex=True)

    mask = s.str.islower() | s.str.isupper()
    s.loc[mask] = s.loc[mask].str.title()
    return s


# =========================================================
# 1. 코어 매칭 및 스코어링 로직
# =========================================================

def _calculate_priority_score(df: pd.DataFrame, cols: Dict, is_media: bool) -> pd.Series:
    """우선순위 점수 계산 (NONE/ANY/Exact 조건 반영)"""
    n_a = df[f"_norm_{cols['raw_cols'][0]}"].str.lower()
    s_a = df[cols["std_cols"][0]].str.lower()
    n_b = df[f"_norm_{cols['raw_cols'][1]}"].str.lower()
    s_b = df[cols["std_cols"][1]].str.lower()
    n_c = df[f"_norm_{cols['raw_cols'][2]}"].str.lower()
    s_c = df[cols["std_cols"][2]].str.lower()

    def match_logic(norm, std, std_col_name):
        return (
            (df[std_col_name] == "NONE") |
            ((df[std_col_name] == "ANY") & (norm != "")) |
            (norm == std)
        )

    a_m = match_logic(n_a, s_a, cols["std_cols"][0])
    b_m = match_logic(n_b, s_b, cols["std_cols"][1])
    c_m = match_logic(n_c, s_c, cols["std_cols"][2])

    if is_media:
        is_aff = (n_a == "affiliate")
        a_m = ((df[cols["std_cols"][0]] == "NONE") & is_aff) | a_m
        b_m = ((df[cols["std_cols"][1]] == "NONE") & is_aff) | b_m
        c_m = ((df[cols["std_cols"][2]] == "NONE") & is_aff) | c_m

    conds = [(a_m & b_m & c_m), (a_m & b_m), b_m, df[cols["key"]].notna()]
    return np.select(conds, [4, 3, 2, 1], default=0)


def _apply_mx_rescue_logic(df: pd.DataFrame, df_map: pd.DataFrame, cols: Dict) -> pd.DataFrame:
    """MX Multi/NaN 구제 로직 (역방향 매핑 적용)"""
    a_cln, b_cln, c_cln = [f"{c}_cleaned" for c in cols["raw_cols"]]
    mask = (df[a_cln] == "Multi") | (df[a_cln].isna())
    if not mask.any():
        return df

    rev_map, fallback = {}, {}
    for row in df_map[[cols["std_cols"][0], cols["std_cols"][1]]].itertuples(index=False):
        a_v, b_v = row[0], row[1]
        if pd.isna(a_v) or pd.isna(b_v): continue
        s_low = str(b_v).lower().strip()
        if s_low not in ["multi", "others", "none", ""]:
            rev_map[s_low] = (a_v, b_v)
        
        c_key = str(a_v).lower().strip()
        if c_key not in fallback: fallback[c_key] = (a_v, "Multi")
        if "other" in s_low and "other" not in str(fallback[c_key][1]).lower():
            fallback[c_key] = (a_v, b_v)

    for idx in df.index[mask]:
        c_l = str(df.at[idx, f"_norm_{cols['raw_cols'][0]}"]).lower().strip()
        s_l = str(df.at[idx, f"_norm_{cols['raw_cols'][1]}"]).lower().strip()
        res = rev_map.get(s_l) or fallback.get(c_l) or rev_map.get(c_l)
        if res:
            df.at[idx, a_cln], df.at[idx, b_cln] = res[0], res[1]
            df.at[idx, c_cln] = df.at[idx, cols["raw_cols"][2]]
    return df


# =========================================================
# 2. 도메인별 파이프라인 (MX/Media & CE)
# =========================================================

def run_cleansing_pipeline(df_raw: pd.DataFrame, df_map: pd.DataFrame, map_cols: Dict, is_media: bool = False) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """MX & Media 클렌징 메인 제어"""
    df = df_raw.reset_index().rename(columns={"index": "original_index"})
    for c in map_cols["raw_cols"]:
        df[f"_norm_{c}"] = fast_normalize_text(df[c])

    merge_key = f"_norm_{map_cols['raw_cols'][2]}"
    df_merged = pd.merge(df, df_map, left_on=merge_key, right_on=map_cols["key"], how="left")
    df_merged["priority"] = _calculate_priority_score(df_merged, map_cols, is_media)

    df_merged.sort_values(["original_index", "priority"], ascending=[True, False], inplace=True)
    df_final = df_merged.drop_duplicates("original_index").copy()

    for r, s in zip(map_cols["raw_cols"], map_cols["std_cols"]):
        df_final[f"{r}_cleaned"] = df_final[s]

    if not is_media:
        df_final = _apply_mx_rescue_logic(df_final, df_map, map_cols)

    df_final.set_index("original_index", inplace=True)
    return df_final, pd.DataFrame()


def run_ce_product_cleansing(df_raw: pd.DataFrame, df_map: pd.DataFrame, map_cols: Dict) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """CE Safety Fix: 중복 시리즈 매칭 차단 로직"""
    df = sanitize_column_headers(df_raw.copy())
    r_cols, s_cols = map_cols["raw_cols"], map_cols["std_cols"]
    
    counts = df_map[s_cols[1]].str.lower().str.strip().value_counts()
    unique_lookup = {}
    for cat in df_map[s_cols[0]].unique():
        s_list = df_map.loc[df_map[s_cols[0]] == cat, s_cols[1]].unique()
        for s in s_list:
            s_l = str(s).lower().strip()
            if counts.get(s_l, 0) == 1: unique_lookup[s_l] = (cat, s)

    c_cat, c_ser, c_prd = [f"{c}_cleaned" for c in r_cols]
    df[c_cat], df[c_ser], df[c_prd] = None, None, df[r_cols[2]]
    
    for idx in df.index:
        s_norm = str(df.at[idx, r_cols[1]]).lower().strip()
        if s_norm in unique_lookup:
            df.at[idx, c_cat], df.at[idx, c_ser] = unique_lookup[s_norm]
        else:
            df.at[idx, c_cat], df.at[idx, c_ser] = df.at[idx, r_cols[0]], "Multi"
            
    return df, pd.DataFrame()


# =========================================================
# 3. 추가 유틸리티 (run.py 연동 필수 함수)
# =========================================================

def process_subsidiary_column(df: pd.DataFrame, col_name: str = 'Subsidiary') -> pd.DataFrame:
    if col_name in df.columns:
        df[col_name] = df[col_name].fillna('').astype(str).str.strip()
    return df

def process_mindset_column(df: pd.DataFrame, col_name: str = 'Mindset') -> pd.DataFrame:
    if col_name not in df.columns: return df
    df[col_name] = df[col_name].fillna('').astype(str).str.strip()
    mask_cold = df[col_name].str.lower() == 'cold'
    df.loc[mask_cold, col_name] = 'Cold'
    return df

def process_funding_column(df: pd.DataFrame, col_name: str = 'Funding') -> pd.DataFrame:
    if col_name not in df.columns: return df
    v = df[col_name].fillna('').astype(str).str.strip().str.lower()
    df[col_name] = np.select([v == 'gmo', v == 'local'], ['GMO', 'Local'], default=df[col_name])
    return df

def process_metric_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {'Revenue (USD)': 'Revenue', 'Spend': 'Media Spend (USD)', 'Cost': 'Media Spend (USD)'}
    df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, inplace=True)
    metrics = ['Impressions', 'Clicks', 'Media Spend (USD)', 'Orders', 'Revenue', 'App Install']
    for col in metrics:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
    if 'Media Spend (USD)' in df.columns and 'Clicks' in df.columns:
        df['CPC'] = np.where(df['Clicks'] > 0, df['Media Spend (USD)'] / df['Clicks'], 0)
    return df

def assign_ce_division(df_cleaned: pd.DataFrame, df_raw: pd.DataFrame, div_rules: Dict) -> pd.DataFrame:
    df = df_cleaned.copy()
    df['BU'] = df_raw['BU'].str.upper().fillna('MX') if 'BU' in df_raw.columns else 'MX'
    cat_col = next((c for c in df.columns if c.endswith('Category_cleaned')), None)
    if cat_col:
        for div, cats in div_rules.items():
            df.loc[df[cat_col].str.lower().isin([c.lower() for c in cats]), 'BU'] = div
    return df

def format_mx_data(df: pd.DataFrame) -> pd.DataFrame:
    df_out = df.copy()
    ov_rules = {'Media Type 1_cleaned': 'Media Type 1', 'Product Category_cleaned': 'Product Category'}
    for c, t in ov_rules.items():
        if c in df_out.columns: df_out[t] = df_out[c]
    for col in config.MX_OUTPUT_COLS:
        if col not in df_out.columns: df_out[col] = None
    return df_out[config.MX_OUTPUT_COLS]

def format_ce_data(df: pd.DataFrame) -> pd.DataFrame:
    df_out = df.copy()
    ov_rules = {'Product Category_cleaned': 'Product Category', 'Products_cleaned': 'Products (Optional)'}
    for c, t in ov_rules.items():
        if c in df_out.columns: df_out[t] = df_out[c]
    for col in config.CE_OUTPUT_COLS:
        if col not in df_out.columns: df_out[col] = None
    return df_out[config.CE_OUTPUT_COLS]

def insert_cleaned_left_of_raw(df: pd.DataFrame) -> pd.DataFrame:
    cols, cleaned = list(df.columns), {c for c in df.columns if c.endswith('_cleaned')}
    order = []
    for c in [c for c in cols if c not in cleaned]:
        if f"{c}_cleaned" in cleaned:
            order.extend([f"{c}_cleaned", c])
            cleaned.remove(f"{c}_cleaned")
        else: order.append(c)
    return df[order + list(cleaned)]