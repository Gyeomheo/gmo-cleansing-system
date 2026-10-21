"""
campaign.py
~~~~~~~~~~~
캠페인명 역추적 — 캠페인명 택소노미의 제품 코드(PR~) · 세부 분류(SB~)로 제품 분류를 보정한다.

예) CN~d2c-ce-tv_PR~nq4k_SB~tv_BS~vd_...  →  PR=nq4k, SB=tv  →  TV / Neo QLED

규칙표: 3_Config/campaign_rules.xlsx (한 줄 = 규칙 하나)
    사용       O 면 사용(비우면 꺼짐)
    법인       SETK 등. 비우거나 * 면 모든 법인
    구분       MX / CE
    PR         제품 코드(대소문자 무시)
    SB         세부 분류 코드. 비우면 아무거나. SB 만 있는 규칙은 가장 마지막(다른 규칙이 안 맞을 때)
    키워드     캠페인명 · Product Series · Products 원본 어디든 들어 있으면(대소문자 · 기호 무시). 비우면 안 봄
               '|' = 그중 하나, '+' = 모두.  예) s26|f1h26 + buds  → (s26 또는 f1h26) 그리고 buds
               '#' 로 시작하면 코드 단어 — 앞뒤가 글자 · 숫자가 아닐 때만(코드끼리 붙은 q8h8b8 은 됨)
                   예) #h8 → 'f2h26-h8' · 'Q8H8B8' 은 맞고 'Watch8' 은 아님
    제외 키워드  이 중 하나라도 있으면 규칙을 안 씀('|' 로 여러 개)   예) Watch Ultra2 규칙에 watch7|watch8
    Category · Series · Products   채울 값. Products 를 비우면 그대로 둔다
    시작일     이 날짜 이후 행만(YYYY-MM-DD). 비우면 전체        예) Watch Ultra 는 2026-07-01 부터 Watch Ultra2
    원본보다 우선  O 면 원본이 구체적이어도 교체(담당자가 정한 규칙만). 키워드에 '+' 인 번들도 교체
    근거       메모

언제 채우나 (2026-10-20 담당자 확정: 원본 우선)
    기본: 매핑 클렌징 결과의 시리즈나 분류가 빈 값 · Multi · Other(s) 일 때만 채운다.
        분류가 이미 구체적이면 규칙과 같을 때만 시리즈를 채운다. 원본이 구체적인데 규칙과 다른 행은
        바꾸지 않고 요약에 '충돌(그대로)'로만 남긴다.
    번들(키워드에 '+'): 한 캠페인에 여러 제품이 같이 있으면 원본보다 우선해서 교체한다.
        예) s26|f1h26 + buds → Multi (S26 · Buds 둘 다 들어 있는 캠페인)
    키워드 규칙이 PR 규칙보다 먼저, 번들이 가장 먼저.

순서: 매핑 엑셀 클렌징 → (여기) 캠페인 보정 → CE 는 BU 할당. 법인 · SB 까지 지정한 규칙이 먼저.
"""
import logging
import re
from pathlib import Path

import pandas as pd

COLS = ['사용', '법인', '구분', 'PR', 'SB', '키워드', '제외 키워드', 'Category', 'Series', 'Products', '시작일',
        '원본보다 우선', '근거']
CAT, SER, PRD = 'Product Category_cleaned', 'Product Series_cleaned', 'Products_cleaned'
ON = {'O', 'Y', '1', '○', 'V'}


def extract_codes(campaign: pd.Series) -> pd.DataFrame:
    """캠페인명 → PR · SB · BS 코드(소문자). 택소노미가 아니면 빈 값."""
    s = campaign.fillna('').astype(str)
    out = pd.DataFrame(index=s.index)
    for tag in ('PR', 'SB', 'BS'):
        out[tag] = s.str.extract(rf'(?:^|_){tag}~([^_]+)', expand=False).fillna('').str.lower().str.strip()
    return out


def is_generic(s: pd.Series) -> pd.Series:
    v = s.fillna('').astype(str).str.strip().str.lower()
    return v.isin(['', 'nan', 'none']) | v.str.startswith(('other', 'multi'))


def load_rules(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=COLS)
    df = pd.read_excel(path, dtype=str).fillna('')
    for c in COLS:
        if c not in df.columns:
            df[c] = ''
    df = df[df['사용'].str.strip().str.upper().isin(ON)].copy()
    for c in ('법인', '구분'):
        df[c] = df[c].str.strip().str.upper().replace('*', '')
    for c in ('PR', 'SB'):
        df[c] = df[c].str.strip().str.lower()
    # 키워드 'a|b + c' → [[a, b], [c]]  (그룹끼리는 모두, 그룹 안은 하나)
    df['kw'] = df['키워드'].map(lambda v: [g for g in ([_key(k) for k in grp.split('|') if _key(k)]
                                                   for grp in str(v).split('+')) if g])
    df['bundle'] = (df['kw'].str.len() > 1) | df['원본보다 우선'].str.strip().str.upper().isin(ON)
    df['since'] = df['시작일'].str.strip().str[:10]
    df['ex'] = df['제외 키워드'].map(lambda v: [_key(k) for k in str(v).split('|') if _key(k)])
    df = df[(df['PR'] != '') | (df['SB'] != '') | (df['kw'].str.len() > 0)]
    sb_only = (df['PR'] == '') & (df['kw'].str.len() == 0)
    # 구체적인 규칙 먼저: 번들 +8, 키워드 +4, 법인 지정 +2, SB 지정 +1, SB 만 있는 규칙 −2(PR 규칙 뒤)
    df['_rank'] = (df['bundle'] * 8 + (df['kw'].str.len() > 0) * 4
                   + (df['법인'] != '') * 2 + (df['SB'] != '') * 1 - sb_only * 2)
    df = df.sort_values('_rank', ascending=False, kind='stable')
    logging.info(f"'{path.name}' 로드 완료. (사용 규칙 {len(df)}개)")
    return df


def _squash(v) -> str:
    """소문자 + 영숫자만 — 'Buds-4' · 'buds 4' · 'Buds4' 를 같게."""
    return re.sub(r'[^a-z0-9]', '', str(v).lower())


def _key(v) -> str:
    """키워드 하나 — '#' 로 시작하면 코드 단어('#' + 영숫자), 아니면 _squash."""
    v = str(v).strip()
    if v.startswith('#'):
        return '#' + _squash(v[1:]) if _squash(v[1:]) else ''
    return _squash(v)


def _has(k: str, squashed: pd.Series, text: pd.Series) -> pd.Series:
    if k.startswith('#'):
        # 앞뒤가 글자 · 숫자가 아니거나, '글자+숫자' 코드끼리만 붙어 있을 때
        pat = rf'(?<![a-z0-9])(?:[a-z]\d+)*{re.escape(k[1:])}(?:[a-z]\d+)*(?![a-z0-9])'
        return text.str.contains(pat, regex=True)
    return squashed.str.contains(k, regex=False)


def _lower(s: pd.Series) -> pd.Series:
    return s.fillna('').astype(str).str.strip().str.lower()


def _rule_rows(r, codes: pd.DataFrame, squashed: pd.Series, sub: pd.Series, dates: pd.Series,
               text: pd.Series) -> pd.Series:
    """규칙 r 의 조건(PR · SB · 법인 · 키워드)에 맞는 행."""
    m = pd.Series(True, index=codes.index)
    if r.PR:
        m &= codes['PR'] == r.PR
    if r.SB:
        m &= codes['SB'] == r.SB
    if r.법인:
        m &= sub == r.법인
    if r.since:
        m &= dates >= r.since
    for group in r.kw:
        m &= pd.concat([_has(k, squashed, text) for k in group], axis=1).any(axis=1)
    for k in r.ex:
        m &= ~_has(k, squashed, text)
    return m


def _tag(r) -> str:
    kw = ' + '.join('|'.join(g) for g in r.kw)
    return '/'.join(x for x in (r.PR and f"PR~{r.PR}", r.SB and f"SB~{r.SB}", kw and f"키워드~{kw}",
                                r.since and f"{r.since}~") if x)


def apply_campaign_rules(df: pd.DataFrame, kind: str, rules: pd.DataFrame):
    """_cleaned 제품 열의 빈 값 · Multi · Other 를 캠페인 규칙으로 채운다(번들은 교체). (df, 요약 df)."""
    df = df.copy()
    df['Campaign_Rule'] = ''
    rules = rules[rules['구분'] == kind] if not rules.empty else rules
    if rules.empty or 'Campaign Name' not in df.columns or CAT not in df.columns:
        return df, pd.DataFrame()

    codes = extract_codes(df['Campaign Name'])
    # 키워드는 캠페인명 + 원본 Product Series · Products 에서 찾는다
    text = df['Campaign Name'].fillna('').astype(str)
    for c in ('Product Series', 'Products'):
        if c in df.columns:
            text = text + ' ' + df[c].fillna('').astype(str)
    text = text.str.lower()
    squashed = text.str.replace(r'[^a-z0-9]', '', regex=True)
    dates = df.get('Date', pd.Series('', index=df.index)).fillna('').astype(str).str[:10]
    sub = df.get('Subsidiary', pd.Series('', index=df.index)).fillna('').astype(str).str.strip().str.upper()
    before = df[[CAT, SER] + ([PRD] if PRD in df.columns else [])].copy()
    taken = pd.Series(False, index=df.index)
    result = pd.Series('', index=df.index)
    proposed = pd.DataFrame('', index=df.index, columns=[CAT, SER])

    for r in rules.itertuples(index=False):
        m = ~taken & _rule_rows(r, codes, squashed, sub, dates, text)
        if not m.any():
            continue
        cat_ok = is_generic(df[CAT]) | (_lower(df[CAT]) == r.Category.lower())
        fill = m if r.bundle else m & (is_generic(df[SER]) | is_generic(df[CAT])) & cat_ok
        differs = (_lower(df[CAT]) != r.Category.lower()) | (_lower(df[SER]) != r.Series.lower())
        conflict = m & ~fill & differs          # 원본이 구체적인데 규칙과 다름 → 그대로 두고 기록만
        for col, val in ((CAT, r.Category), (SER, r.Series), (PRD, r.Products)):
            if val and col in df.columns:
                df.loc[fill, col] = val
        df.loc[fill, 'Campaign_Rule'] = _tag(r)
        df.loc[conflict, 'Campaign_Rule'] = _tag(r) + ' 충돌'
        result[fill] = '번들' if r.bundle else '채움'
        result[conflict] = '충돌(그대로)'
        proposed.loc[conflict, CAT] = r.Category
        proposed.loc[conflict, SER] = r.Series
        taken |= m

    changed = result.isin(['채움', '번들']) & (
        before.fillna('').astype(str) != df[before.columns].fillna('').astype(str)).any(axis=1)
    show = changed | (result == '충돌(그대로)')
    if not show.any():
        return df, pd.DataFrame()
    after_cat = df[CAT].where(changed, proposed[CAT])
    after_ser = df[SER].where(changed, proposed[SER])
    log = pd.DataFrame({
        '결과': result[show], '규칙': df.loc[show, 'Campaign_Rule'].str.replace(' 충돌', ''),
        '전 Category': before.loc[show, CAT], '전 Series': before.loc[show, SER],
        '후(규칙) Category': after_cat[show], '후(규칙) Series': after_ser[show],
    })
    spend = pd.to_numeric(df.get('Media Spend (USD)', 0), errors='coerce')
    log['Spend'] = spend.loc[show] if isinstance(spend, pd.Series) else 0
    keys = [c for c in log.columns if c != 'Spend']
    summary = (log.fillna('').groupby(keys).agg(Counts=('Spend', 'size'), Spend=('Spend', 'sum'))
               .reset_index().sort_values(['결과', 'Spend'], ascending=[True, False]))
    summary['Spend'] = summary['Spend'].round(2)
    n_fill, n_conf = int(changed.sum()), int((result == '충돌(그대로)').sum())
    logging.info(f"   -> (Campaign) 캠페인명 역추적: 채움 {n_fill}행 · 충돌(그대로) {n_conf}행")
    return df, summary
