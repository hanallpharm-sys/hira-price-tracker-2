"""
parser.py
약제급여목록표 엑셀을 정규화된 DataFrame으로 변환한다.

심평원 약제급여목록표는 시점마다 컬럼 순서/명칭이 조금씩 달라질 수 있으므로,
헤더 텍스트를 키워드로 매칭해 유연하게 인식한다.
"""

import re
import pandas as pd


# 정규화 후 사용할 표준 컬럼명 -> 원본 헤더에서 찾을 키워드 우선순위
COLUMN_KEYWORDS = {
    "code":       ["제품코드", "품목기준코드", "보험코드", "edi코드", "기준코드"],
    "name":       ["제품명", "품목명", "약품명"],
    "maker":      ["업체명", "제조업체", "수입업체", "회사명", "업소명", "업체"],
    "spec":       ["규격"],
    "unit":       ["단위"],
    "price":      ["상한금액", "상한가", "급여상한", "금액"],
    "ingr_code":  ["주성분코드"],   # 동일 성분 = 제네릭 경쟁군 그룹핑 키
    "ingredient": ["주성분명", "성분명"],
    "klass":      ["약효분류번호", "약효분류", "식약분류", "분류번호", "분류"],
    "route":      ["투여경로", "투여", "제형"],
}


def _norm(s: str) -> str:
    """헤더 비교용: 공백/괄호/특수문자 제거 후 소문자."""
    return re.sub(r"[\s()\[\]/_\-.]", "", str(s)).lower()


def _find_header_row(raw: pd.DataFrame, scan_rows: int = 15) -> int:
    """
    상단에 제목/안내문이 있는 경우가 많으므로, '제품명' 또는 '상한금액' 류
    키워드가 가장 많이 등장하는 행을 헤더 행으로 추정한다.
    """
    best_row, best_score = 0, -1
    targets = [_norm(k) for ks in COLUMN_KEYWORDS.values() for k in ks]
    for i in range(min(scan_rows, len(raw))):
        cells = [_norm(c) for c in raw.iloc[i].tolist()]
        score = sum(any(t in c for t in targets) for c in cells if c)
        if score > best_score:
            best_row, best_score = i, score
    return best_row


def _map_columns(headers):
    """원본 헤더 리스트 -> {표준명: 원본헤더} 매핑. 우선순위 키워드 순으로 첫 매칭."""
    norm_headers = {h: _norm(h) for h in headers}
    mapping, used = {}, set()
    for std, keywords in COLUMN_KEYWORDS.items():
        for kw in keywords:
            nkw = _norm(kw)
            match = next(
                (h for h, nh in norm_headers.items()
                 if h not in used and nkw in nh),
                None,
            )
            if match:
                mapping[std] = match
                used.add(match)
                break
    return mapping


def detect_label(path: str) -> str | None:
    """
    파일에서 적용월(YYYY-MM)을 추정한다.
    1) 시트명 (예: '2026년6월1일_(21,898)_공개용') 
    2) 파일명 (예: '..._2026_6_1_...')
    실패 시 None.
    """
    import os
    candidates = []
    try:
        xl = pd.ExcelFile(path)
        candidates.extend(xl.sheet_names)
    except Exception:
        pass
    candidates.append(os.path.basename(path))

    for text in candidates:
        # 'YYYY년 M월' 또는 'YYYY년M월'
        m = re.search(r"(20\d{2})\s*년\s*(\d{1,2})\s*월", str(text))
        if not m:
            # 'YYYY_M_' / 'YYYY-M-' / 'YYYY.M.'
            m = re.search(r"(20\d{2})[._\-](\d{1,2})[._\-]", str(text))
        if m:
            y, mo = int(m.group(1)), int(m.group(2))
            if 1 <= mo <= 12:
                return f"{y}-{mo:02d}"
    return None


def parse(path: str, sheet=0) -> pd.DataFrame:
    """
    엑셀 경로 -> 정규화 DataFrame.
    반환 컬럼: code, name, maker, spec, unit, price, klass, route, etc (있는 것만)
    'code'와 'price'는 필수로 간주하고, code를 문자열로 통일한다.
    """
    raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=str)
    hrow = _find_header_row(raw)
    headers = [str(c).strip() for c in raw.iloc[hrow].tolist()]

    df = raw.iloc[hrow + 1:].copy()
    df.columns = headers
    df = df.dropna(how="all")

    mapping = _map_columns(headers)
    if "code" not in mapping or "price" not in mapping:
        raise ValueError(
            f"필수 컬럼(code/price)을 못 찾았습니다. 인식된 헤더: {headers}\n"
            f"매핑 결과: {mapping}"
        )

    out = pd.DataFrame()
    for std, src in mapping.items():
        out[std] = df[src]

    # 코드 정규화: 공백 제거, 소수점 .0 제거, 문자열화
    out["code"] = (
        out["code"].astype(str).str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )
    out = out[out["code"].notna() & (out["code"] != "") & (out["code"].str.lower() != "nan")]

    # 가격 정규화: 콤마/공백 제거 후 숫자
    out["price"] = (
        out["price"].astype(str)
        .str.replace(r"[,\s]", "", regex=True)
    )
    out["price"] = pd.to_numeric(out["price"], errors="coerce")

    # 중복 코드는 마지막 행 유지(드물지만 방어적으로)
    out = out.drop_duplicates(subset="code", keep="last").reset_index(drop=True)
    return out


if __name__ == "__main__":
    import sys
    d = parse(sys.argv[1])
    print(d.head(10).to_string())
    print(f"\n총 {len(d):,}개 품목, 컬럼: {list(d.columns)}")
