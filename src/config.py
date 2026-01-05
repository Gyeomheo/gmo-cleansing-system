import os
from pathlib import Path

# =========================================================
# 1. 경로 설정 (Path Configuration)
# =========================================================
BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_DIR = BASE_DIR / "1_Input"
OUTPUT_DIR = BASE_DIR / "2_Output"
CONFIG_DIR = BASE_DIR / "3_Config"
BACKUP_DIR = OUTPUT_DIR / "Backup"

INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
BACKUP_DIR.mkdir(parents=True, exist_ok=True)

# =========================================================
# 2. 파일명 설정 (Excel Files)
# =========================================================
MEDIA_MAP_EXCEL = CONFIG_DIR / "media_mapping_MASTER.xlsx"
PRODUCT_MX_EXCEL = CONFIG_DIR / "product_mapping_MX.xlsx"
PRODUCT_CE_EXCEL = CONFIG_DIR / "product_mapping_CE.xlsx"

MEDIA_MAP_JSON = CONFIG_DIR / "media_mapping.json"
PRODUCT_MX_JSON = CONFIG_DIR / "product_mapping_MX.json"
PRODUCT_CE_JSON = CONFIG_DIR / "product_mapping_CE.json"

# =========================================================
# 3. 마스터 DB 경로 설정
# =========================================================
MASTER_FILES = {
    "MX": OUTPUT_DIR / "TOTAL_MASTER_MX.csv",
    "CE": OUTPUT_DIR / "TOTAL_MASTER_CE.csv"
}

# =========================================================
# 4. 컬럼 매핑 설정 (Column Mapping)
# =========================================================
# (1) Media Mapping
MEDIA_COLS_MAP = {
    'raw_cols': ['Media Type 1', 'Media Type 2', 'Media Platform'],
    'std_cols': ['D_Standard', 'E_Standard', 'F_Standard'],
    'key': 'F_Key',
    'normalize_cols': ['Media Type 1', 'Media Type 2', 'Media Platform'] 
}

# (2) Product Mapping (MX)
PRODUCT_COLS_MAP_MX = {
    'raw_cols': ['Product Category', 'Product Series', 'Products'],
    'std_cols': ['A_Standard', 'B_Standard', 'C_Standard'],
    'key': 'C_Key'
}

# (3) Product Mapping (CE)
PRODUCT_COLS_MAP_CE = {
    'raw_cols': ['Product Category', 'Product Series', 'Products'],
    'std_cols': ['A_Standard', 'B_Standard', 'C_Standard'],
    'key': 'A_Key'
}

# 리포트용 헤더
PRODUCT_COLS = ['Product Category', 'Product Series', 'Products']
MEDIA_COLS = ['Media Type 1', 'Media Type 2', 'Media Platform']
CE_PRODUCT_COLS = ['Product Category', 'Product Series', 'Products']

# =========================================================
# 5. 비즈니스 로직 설정
# =========================================================
DIV_RULES = {
    "VD": ["tv", "sound device", "audio", "av", "visual display", "monitor", "smart monitor", "projector"],
    "DA": ["refrigerator", "fridge", "washing machine", "washer", "dryer", "laundry", "air conditioner", "ac", "vacuum", "cleaner", "oven", "cooking", "dishwasher", "air dresser"],
    "MX": ["smartphone", "mobile", "tablet", "wearable", "watch", "buds", "galaxy", "accessory", "pc", "laptop", "notebook"]
}

AMBIGUOUS_CATS = ["memory", "storage", "ssd", "display", "b2b", "others"]

# =========================================================
# 6. [최종 출력 스펙] 마스터 DB용 컬럼 순서
# =========================================================
# [MX]
MX_OUTPUT_COLS = [
    'Subsidiary', 'Sales Channel', 'Partner', 
    'Media Type 1', 'Media Type 2', 'Media Type 2 (Raw)', 
    'Media Platform', 'Media Platform (Raw)', 
    'Funding', 
    'Product Category', 'Product Series', 'Products', 
    'Campaign Name', 'Mindset', 
    'Quarter', 'Month', 'Week', 'Date', 
    'Media Spend (USD)', 'Impressions', 'Clicks', 'CPC', 
    'Orders', 'Revenue', 'App Install'
]

# [CE]
CE_OUTPUT_COLS = [
    'Subsidiary', 'Sales Channel', 'Partner', 
    'Media Type 1', 'Media Type 2', 'Media Type 2 (Raw)', 
    'Media Platform', 'Media Platform (Raw)', 
    'Funding', 'BU', 
    'Product Category', 'Product Series', 'Products (Optional)', 
    'Campaign Name', 'Mindset', 
    'Quarter', 'Month', 'Week', 'Date', 
    'Media Spend (USD)', 'Impressions', 'Clicks', 'CPC', 
    'Orders', 'Revenue', 'App Install'
]

# =========================================================
# 7. 시스템 설정 (System Config)
# =========================================================
# ⭐️ [MISSING FIX] 엑셀 저장 엔진 설정
WRITE_ENGINE = 'openpyxl'