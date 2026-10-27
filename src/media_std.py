"""
media_std.py
~~~~~~~~~~~~
미디어 표준표 검사 — 기존 미디어 매핑 클렌징 뒤에, 결과가 D2C 표준표(3_Config/media_taxonomy.csv)에 있는지 본다.

담당자 확정 (2026-10-20): 미디어는 매핑을 더 늘리지 않는다. 표에 없는 것은 Others 로, 원본 값은
마스터의 'Media Platform (Raw)' 열(포맷터가 원본 Media Platform 을 넣는다)에만 남긴다.

    Platform 이 그 Type 2 에 없음  → Platform = Others  (DSP 는 'Other DSPs')
                                     그 Type 2 에 Others 가 없으면 → Type 2 · Platform = Others
                                     (예: Performance Max 는 Google 만)
    Type 2 가 그 Type 1 에 없음    → Type 2 · Platform = Others  (Paid Search 는 Type 2 가 Text Ads 하나라 Text Ads)
    Affiliate                      → Platform 은 퍼블리셔 원본 그대로(표의 *), Type 2 만 본다
    Type 1 이 표에 없음            → Type 2 가 표에서 한 Type 1 에만 있으면 그 Type 1
                                     (Other / PerformanceMax → Shopping Ads)
                                     못 찾으면 바꾸지 않고 요약에 '확인' 으로만
    Platform 이 비었거나 Other 류      → 표의 '기본' 열에 O 인 Platform (YouTube → Google, 담당자 2026-10-20)
                                     없으면 그 Type 2 의 Platform 이 하나뿐일 때 그 하나
                                     (Performance Max · GDN · Demand Gen → Google)
    대소문자 · 띄어쓰기 · 기호만 다르면 표 표기로 (TikTok → Tiktok, PerformanceMax → Performance Max)
    약어는 3_Config/media_aliases.csv 로 먼저 푼다 (PMAX → Performance Max, FB → Meta …)

원본 값으로 정하는 규칙 (담당자 2026-10-23) — 표 검사보다 먼저, 원본 칸을 보고 _cleaned 를 정한다(원본 칸은 그대로)
    Type 2 원본이 Smart Shopping        → Shopping Ads / Performance Max (Platform 은 매핑 결과 그대로)
    Type 1 · Type 2 원본이 둘 다 App    → 캠페인명에 install 이 있으면 App Install, 없으면 App Engagement
    SEDA: Platform 이 Others · 빈 값인데 'Media Platform_RAW (optional)' 에 값이 있으면 그 값 (담당자 2026-10-27)
                                     → 이어지는 표 검사로 표에 있는 조합(Paid Social · Standard Ads · Meta / Tiktok 등)만 남고
                                       표 밖(App · Tiktok, Jampp …)은 다시 Others
"""
import logging
import re
from pathlib import Path

import pandas as pd

T1, T2, PF = 'Media Type 1_cleaned', 'Media Type 2_cleaned', 'Media Platform_cleaned'
RAW_T1, RAW_T2, CAMPAIGN = 'Media Type 1', 'Media Type 2', 'Campaign Name'
RAW_PF_OPT, SUB = 'Media Platform_RAW (optional)', 'Subsidiary'
PF_FROM_RAW_SUBS = {'SEDA'}     # Platform 이 Others 인데 'Media Platform_RAW (optional)' 에 Meta · Tiktok 등이 있는 법인


ALIASES: dict = {'Type 2': {}, 'Platform': {}}


def load_taxonomy(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    t = pd.read_csv(path, dtype=str, encoding='utf-8-sig').fillna('')
    if t.shape[1] < 4:
        t['기본'] = ''
    t = t.iloc[:, :4]
    t.columns = ['t1', 't2', 'pf', 'default']
    alias_path = path.with_name('media_aliases.csv')
    if alias_path.exists():
        a = pd.read_csv(alias_path, dtype=str, encoding='utf-8-sig').fillna('')
        for col, short, full in a.iloc[:, :3].itertuples(index=False):
            if col.strip() in ALIASES and _key(short):
                ALIASES[col.strip()][_key(short)] = _key(full)
    logging.info(f"'{path.name}' 로드 완료. (표준 조합 {len(t)}개)")
    return t


def _key(v) -> str:
    """비교용 — 소문자 + 영숫자만 ('PerformanceMax' = 'Performance Max')."""
    return re.sub(r'[^a-z0-9*]', '', str(v).lower())


def _canon(values) -> dict:
    return {_key(v): str(v).strip() for v in values}


GENERIC_PF = {'', 'other', 'others', 'none', 'nan', 'etc'}


def _decide(a: str, b: str, p: str, t1c: dict, t2c: dict, pfc: dict, defaults: dict) -> tuple[tuple, str]:
    """(Type1, Type2, Platform) 한 조합 → (표준 조합, 사유)."""
    al, bl, pl = _key(a), _key(b), _key(p)
    bl, pl = ALIASES['Type 2'].get(bl, bl), ALIASES['Platform'].get(pl, pl)
    note = ''
    if al not in t1c:
        owners = [x for x in t1c if bl and bl in t2c[x]]
        if len(owners) != 1:
            return (a, b, p), '확인(Type 1 표에 없음)'
        al, note = owners[0], 'Type 1 ← Type 2 · '
    A = t1c[al]
    if bl not in t2c[al]:
        if 'others' in t2c[al]:
            return (A, 'Others', p if A == 'Affiliate' else 'Others'), 'Type 2 → Others'
        if A != 'Paid Search':
            return (a, b, p), '확인(Type 2 표에 없음)'
        bl = 'textads'                       # Paid Search 의 Type 2 는 Text Ads 하나
    B = t2c[al][bl]
    allowed = pfc[(al, bl)]
    if '*' in allowed or pl in allowed:
        new = (A, B, p if '*' in allowed else allowed[pl])
        return new, (note + '표기 맞춤' if new != (a, b, p) else '')
    if pl in GENERIC_PF and (al, bl) in defaults:
        return (A, B, defaults[(al, bl)]), note + "Platform ← 기본값"
    if pl in GENERIC_PF and len(allowed) == 1:
        return (A, B, next(iter(allowed.values()))), note + 'Platform ← 그 Type 2 의 유일한 Platform'
    if 'others' in allowed:
        return (A, B, 'Others'), note + 'Platform → Others'
    if 'otherdsps' in allowed:
        return (A, B, 'Other DSPs'), note + 'Platform → Other DSPs'
    if 'others' in pfc.get((al, 'others'), {}):
        return (A, 'Others', 'Others'), note + 'Type 2 · Platform → Others'
    return (a, b, p), '확인(Others 자리 없음)'


def _raw_rules(df: pd.DataFrame) -> pd.Series:
    """원본 값으로 정하는 규칙(위 docstring) — df 의 _cleaned 를 바꾸고, 행마다 사유를 돌려준다."""
    why = pd.Series('', index=df.index)
    if RAW_T2 not in df.columns:                       # 마스터 재검사(consolidate · master_media)는 원본 칸이 없다
        return why
    t2 = df[RAW_T2].map(_key)
    smart = t2 == 'smartshopping'
    df.loc[smart, T1], df.loc[smart, T2] = 'Shopping Ads', 'Performance Max'
    why[smart] = '원본 규칙: Smart Shopping = Performance Max'
    if RAW_T1 in df.columns:
        app = (df[RAW_T1].map(_key) == 'app') & (t2 == 'app')
        camp = df[CAMPAIGN].fillna('').astype(str).str.lower() if CAMPAIGN in df.columns else pd.Series('', index=df.index)
        install = camp.str.contains('install', regex=False)
        df.loc[app, T1] = 'App'
        df.loc[app & install, T2] = 'App Install'
        df.loc[app & ~install, T2] = 'App Engagement'
        why[app & install] = '원본 규칙: App · 캠페인명 install → App Install'
        why[app & ~install] = '원본 규칙: App · 캠페인명 install 없음 → App Engagement'
    if RAW_PF_OPT in df.columns and SUB in df.columns:
        # SEDA: Platform 이 Others 면 Platform_RAW 값으로 — 표 검사가 이어서 보므로 표에 있는 조합만 남고 나머지는 다시 Others
        raw_pf = df[RAW_PF_OPT].fillna('').astype(str).str.strip()
        hit = (df[SUB].fillna('').astype(str).str.strip().str.upper().isin(PF_FROM_RAW_SUBS)
               & df[PF].map(_key).isin(GENERIC_PF) & (raw_pf.map(_key) != '') & ~raw_pf.map(_key).isin(GENERIC_PF))
        df.loc[hit, PF] = raw_pf[hit]
        why[hit] = why[hit].where(why[hit] == '', why[hit] + ' · ') + '원본 규칙: Platform ← Platform_RAW(SEDA)'
    return why


def standardize_media(df: pd.DataFrame, tax: pd.DataFrame):
    """(df, 요약 df). _cleaned 미디어 3열만 바꾼다."""
    df = df.copy()
    if tax.empty or any(c not in df.columns for c in (T1, T2, PF)):
        return df, pd.DataFrame()
    orig = df[[T1, T2, PF]].fillna('').astype(str)   # 요약의 '전' = 원본 규칙 · 표 검사 전
    raw_why = _raw_rules(df)
    t1c = _canon(tax['t1'])
    k1, k2 = tax['t1'].map(_key), tax['t2'].map(_key)
    t2c = {a: _canon(tax.loc[k1 == a, 't2']) for a in t1c}
    pfc = {(a, b): _canon(tax.loc[(k1 == a) & (k2 == b), 'pf']) for a in t1c for b in t2c[a]}
    on = tax['default'].str.strip().str.upper() == 'O'
    defaults = {(_key(r.t1), _key(r.t2)): r.pf.strip() for r in tax[on].itertuples(index=False)}

    before = df[[T1, T2, PF]].fillna('').astype(str)
    combos = before.drop_duplicates().itertuples(index=False, name=None)
    decided = {k: _decide(*k, t1c, t2c, pfc, defaults) for k in combos}
    keys = list(zip(before[T1], before[T2], before[PF]))
    after = pd.DataFrame([decided[k][0] for k in keys], index=df.index, columns=[T1, T2, PF])
    why = pd.Series([decided[k][1] for k in keys], index=df.index)
    why = raw_why.where(why == '', raw_why + ' · ' + why).where(raw_why != '', why)
    before = orig
    changed = (after != before).any(axis=1)
    df[[T1, T2, PF]] = after
    df['Media_Rule'] = why
    check = why.str.startswith('확인')
    show = changed | check
    if not show.any():
        return df, pd.DataFrame()
    spend = pd.to_numeric(df.get('Media Spend (USD)', 0), errors='coerce')
    log = pd.DataFrame({'결과': why[show],
                        '전': before.loc[show].agg(' / '.join, axis=1),
                        '후': after.loc[show].agg(' / '.join, axis=1),
                        'Spend': spend.loc[show] if isinstance(spend, pd.Series) else 0})
    summary = (log.groupby(['결과', '전', '후']).agg(Counts=('Spend', 'size'), Spend=('Spend', 'sum'))
               .reset_index().sort_values('Spend', ascending=False))
    summary['Spend'] = summary['Spend'].round(2)
    logging.info(f"   -> (Media) 표준표 검사: 바꿈 {int(changed.sum())}행 · 확인 {int(check.sum())}행")
    return df, summary
