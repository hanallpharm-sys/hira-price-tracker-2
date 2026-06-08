"""
streamlit_app.py
약제급여목록표 변동 추적 — Streamlit Community Cloud 배포용

- monthly_data/ 의 월별 엑셀을 읽어 업체(기본: 한올바이오파마) 제품을 추적
- 제품을 검색/선택하면 그 제품의 '변동 이력'을 건별로 누적 표시
  (변동 일자 · 구분(신규/인상/인하/삭제) · 금액 변화 · 사유)
- 변동(인상/인하)·삭제 건에 사유를 직접 입력·저장 (GitHub 토큰 등록 시)
"""
import os
import glob
import json
import base64

import pandas as pd
import requests
import streamlit as st

import parser
import company

DATA_DIR = os.path.join(os.path.dirname(__file__), "monthly_data")
NOTES_PATH = os.path.join(os.path.dirname(__file__), "notes.json")

st.set_page_config(page_title="약제급여목록표 변동 추적", page_icon="💊", layout="wide")


# ───────────── 데이터 로드 ─────────────
@st.cache_data(show_spinner="약제급여목록표를 읽는 중...")
def load_snapshots(file_sig):
    snaps = {}
    for path in sorted(glob.glob(os.path.join(DATA_DIR, "*.xls*"))):
        label = parser.detect_label(path) or os.path.splitext(os.path.basename(path))[0]
        try:
            snaps[label] = parser.parse(path)
        except Exception as e:
            st.warning(f"{os.path.basename(path)} 읽기 실패: {e}")
    return snaps


def file_signature():
    files = sorted(glob.glob(os.path.join(DATA_DIR, "*.xls*")))
    return tuple((os.path.basename(f), os.path.getmtime(f)) for f in files)


# ───────────── 사유(notes) : {제품코드: {월: 사유}} ─────────────
def load_notes():
    notes = {}
    if os.path.exists(NOTES_PATH):
        try:
            with open(NOTES_PATH, encoding="utf-8") as f:
                raw = json.load(f)
            for code, v in raw.items():
                if isinstance(v, dict):
                    notes[code] = dict(v)
        except Exception:
            notes = {}
    for code, mm in st.session_state.get("notes_override", {}).items():
        notes.setdefault(code, {}).update(mm)
    return notes


def github_enabled():
    return "GITHUB_TOKEN" in st.secrets


def save_notes_github(notes: dict) -> bool:
    token = st.secrets["GITHUB_TOKEN"]
    repo = st.secrets.get("GITHUB_REPO", "hanallpharm-sys/hira-price-tracker")
    branch = st.secrets.get("GITHUB_BRANCH", "main")
    url = f"https://api.github.com/repos/{repo}/contents/notes.json"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}
    r = requests.get(url, headers=headers, params={"ref": branch}, timeout=20)
    sha = r.json().get("sha") if r.status_code == 200 else None
    content = base64.b64encode(json.dumps(notes, ensure_ascii=False, indent=2).encode()).decode()
    payload = {"message": "사유 업데이트", "content": content, "branch": branch}
    if sha:
        payload["sha"] = sha
    r = requests.put(url, headers=headers, json=payload, timeout=20)
    return r.status_code in (200, 201)


# ───────────── 건별 변동 이력 ─────────────
def product_events(row, months):
    events, prev, seen = [], None, False
    for i, m in enumerate(months):
        v = row[m]
        if v == "X":
            if seen and i > 0 and row[months[i - 1]] != "X":
                events.append({"월": m, "구분": "삭제", "이전": prev, "이번": None, "변동액": None})
            continue
        v = float(v)
        if not seen:
            if m != months[0]:
                events.append({"월": m, "구분": "신규", "이전": None, "이번": v, "변동액": None})
        elif prev is not None and v != prev:
            d = v - prev
            events.append({"월": m, "구분": "인상" if d > 0 else "인하",
                           "이전": prev, "이번": v, "변동액": d})
        prev, seen = v, True
    return events


def latest_price(row, months):
    for m in reversed(months):
        if row[m] != "X":
            return float(row[m])
    return None


def style_matrix(mat, months):
    disp = mat.rename(columns={"name": "제품명", "status": "상태"})[["제품명"] + months + ["상태"]].copy()
    marks = [company.cell_marks(r, months) for _, r in mat.iterrows()]
    cmap = {"x": "background-color:#fdeceb;color:#b42318;font-weight:700",
            "new": "background-color:#e7f5ed;color:#0f7a4d;font-weight:700",
            "up": "background-color:#fdf0e6;color:#b54708;font-weight:700",
            "down": "background-color:#eaf1fa;color:#1f5fa8;font-weight:700"}

    def fmt(v):
        if v == "X":
            return "X"
        try:
            return f"{float(v):,.0f}"
        except (ValueError, TypeError):
            return "" if v in (None, "") else str(v)
    for m in months:
        disp[m] = disp[m].map(fmt)

    def color(_):
        s = pd.DataFrame("", index=disp.index, columns=disp.columns)
        for i, mk in enumerate(marks):
            for m in months:
                if mk[m] in cmap:
                    s.loc[disp.index[i], m] = cmap[mk[m]]
        return s
    return disp.style.apply(color, axis=None)


# ═════════════════════════ UI ═════════════════════════
st.title("💊 약제급여목록표 변동 추적")
st.caption("HIRA 약제급여목록 및 급여상한금액표 · 업체별 제품 변동 이력")

snapshots = load_snapshots(file_signature())
if not snapshots:
    st.info("아직 데이터가 없습니다. GitHub 레포의 **monthly_data** 폴더에 약제급여목록표 엑셀(.xlsx)을 올리면 자동으로 나타납니다.")
    st.stop()

months = sorted(snapshots.keys())

with st.sidebar:
    st.subheader("설정")
    maker = st.text_input("업체명 필터", value=company.DEFAULT_MAKER)
    st.markdown("---")
    st.markdown(f"**누적 기간**: {months[0]} ~ {months[-1]} ({len(months)}개월)")
    st.markdown("**구분 색상**  \n🟩 신규  🟧 인상  🟦 인하  🟥 삭제")

mat, mon = company.build_matrix(maker, snapshots=snapshots)
if not len(mat):
    st.warning(f"'{maker}' 업체 제품을 찾지 못했습니다. 업체명을 확인해 주세요.")
    st.stop()

notes = load_notes()
events_by_code, name_by_code = {}, {}
for _, r in mat.iterrows():
    events_by_code[r["code"]] = product_events(r, mon)
    name_by_code[r["code"]] = r["name"]

tot = [e for evs in events_by_code.values() for e in evs]
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("전체 품목", len(mat))
c2.metric("신규(건)", sum(1 for e in tot if e["구분"] == "신규"))
c3.metric("인상(건)", sum(1 for e in tot if e["구분"] == "인상"))
c4.metric("인하(건)", sum(1 for e in tot if e["구분"] == "인하"))
c5.metric("삭제(건)", sum(1 for e in tot if e["구분"] == "삭제"))

# 검색 + 목록
st.markdown("### 🔍 제품 검색 · 선택")
query = st.text_input("제품 검색", value="",
                      placeholder="제품명 또는 제품코드 검색  (예: 앱시토, 베노론, 645302132 / 비우면 전체)",
                      label_visibility="collapsed")
rows = []
for _, r in mat.iterrows():
    n_chg = sum(1 for e in events_by_code[r["code"]] if e["구분"] in ("인상", "인하", "삭제"))
    rows.append({"code": r["code"], "제품명": r["name"], "상태": r["status"],
                 "변동건수": n_chg, "최근상한금액": latest_price(r, mon)})
plist = pd.DataFrame(rows)
if query.strip():
    q = query.strip()
    plist = plist[plist["제품명"].str.contains(q, case=False, na=False) | plist["code"].str.contains(q, na=False)]
plist = plist.sort_values(["변동건수", "제품명"], ascending=[False, True]).reset_index(drop=True)

st.caption(f"{len(plist)}개 제품 · 행을 클릭하면 아래에 변동 이력이 표시됩니다.")
event = st.dataframe(
    plist.drop(columns="code"), width="stretch", hide_index=True, height=300,
    on_select="rerun", selection_mode="single-row",
    column_config={"제품명": st.column_config.TextColumn(width="large"),
                   "변동건수": st.column_config.NumberColumn(width="small"),
                   "최근상한금액": st.column_config.NumberColumn(format="%,d")},
)
try:
    sel_rows = event.selection.rows
except Exception:
    sel_rows = []

# 변동 이력
st.markdown("### 📋 제품별 변동 이력")
if not sel_rows:
    st.info("위 목록에서 제품을 클릭하면 그 제품의 변동 이력(일자·구분·사유)이 누적되어 표시됩니다.")
else:
    sel_code = plist.iloc[sel_rows[0]]["code"]
    st.markdown(f"**{name_by_code[sel_code]}**  ·  제품코드 `{sel_code}`")
    evs = events_by_code[sel_code]
    if not evs:
        st.success("이 기간 동안 변동 이력이 없습니다.")
    else:
        ev_df = pd.DataFrame(evs)
        ev_df["사유"] = ev_df["월"].map(lambda m: notes.get(sel_code, {}).get(m, ""))
        ev_df["이전"] = ev_df["이전"].map(lambda x: f"{x:,.0f}" if pd.notna(x) else "-")
        ev_df["이번"] = ev_df["이번"].map(lambda x: f"{x:,.0f}" if pd.notna(x) else "X")
        ev_df["변동액"] = ev_df["변동액"].map(lambda x: f"{x:+,.0f}" if pd.notna(x) else "-")
        ev_df = ev_df.rename(columns={"월": "변동일자(적용월)"})[
            ["변동일자(적용월)", "구분", "이전", "이번", "변동액", "사유"]]
        if github_enabled():
            st.caption("‘사유’ 칸을 클릭해 입력하고 아래 저장을 누르면 모두에게 반영됩니다. (신규는 사유 불필요)")
            edited = st.data_editor(
                ev_df, width="stretch", hide_index=True, key=f"hist_{sel_code}",
                column_config={
                    "변동일자(적용월)": st.column_config.TextColumn(disabled=True),
                    "구분": st.column_config.TextColumn(disabled=True, width="small"),
                    "이전": st.column_config.TextColumn(disabled=True, width="small"),
                    "이번": st.column_config.TextColumn(disabled=True, width="small"),
                    "변동액": st.column_config.TextColumn(disabled=True, width="small"),
                    "사유": st.column_config.TextColumn("사유", width="large")},
            )
            if st.button("💾 이 제품 사유 저장", type="primary"):
                month_notes = dict(notes.get(sel_code, {}))
                for _, er in edited.iterrows():
                    m, txt = er["변동일자(적용월)"], (er["사유"] or "").strip()
                    if txt:
                        month_notes[m] = txt
                    elif m in month_notes:
                        del month_notes[m]
                new_notes = {c: dict(mm) for c, mm in notes.items()}
                if month_notes:
                    new_notes[sel_code] = month_notes
                elif sel_code in new_notes:
                    del new_notes[sel_code]
                with st.spinner("저장 중..."):
                    ok = save_notes_github(new_notes)
                if ok:
                    ov = st.session_state.get("notes_override", {})
                    ov.setdefault(sel_code, {}).update(month_notes)
                    st.session_state["notes_override"] = ov
                    st.success("저장되었습니다. 잠시 후 모든 사용자에게 반영됩니다.")
                else:
                    st.error("저장 실패. GitHub 토큰 권한(Contents: Read and write)을 확인해 주세요.")
        else:
            st.dataframe(ev_df, width="stretch", hide_index=True)
            st.info('사유를 앱에서 직접 입력하려면 GitHub 토큰 등록이 필요합니다(맨 아래 안내). '
                    '토큰 없이 쓰려면 레포의 notes.json을 편집하세요. 형식: {"제품코드": {"2026-06": "사유"}}')

# 전체 매트릭스 + 다운로드
with st.expander("📊 전체 월별 매트릭스 보기"):
    st.dataframe(style_matrix(mat, mon), width="stretch", height=500)

mat_x = mat.copy()
mat_x["사유"] = mat_x["code"].map(
    lambda c: " / ".join(f"{m}:{r}" for m, r in sorted(notes.get(c, {}).items())) if notes.get(c) else "")
tmp = "/tmp/_company_matrix.xlsx"
company.to_excel(maker, mat_x, mon, tmp)
with open(tmp, "rb") as f:
    st.download_button("📥 엑셀로 다운로드 (사유 포함)", data=f.read(),
                       file_name=f"{maker}_월별상한금액_{mon[0]}_{mon[-1]}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with st.expander("ℹ️ 매달 업데이트 / 사유 입력용 토큰 등록 방법"):
    st.markdown(
        "**매달 데이터 추가**  \n"
        "1. GitHub 레포 → monthly_data 폴더 → Add file → Upload files로 새 달 엑셀 올리고 Commit  \n"
        "2. 1~2분 뒤 자동 반영  \n\n"
        "**앱에서 직접 사유 입력 (일회성 토큰 등록)**  \n"
        "1. GitHub → Settings → Developer settings → Fine-grained tokens → Generate new token  \n"
        "2. 이 레포만 선택, 권한 Contents: Read and write 부여 → 토큰 생성·복사  \n"
        "3. Streamlit 앱 우상단 ⋮ → Settings → Secrets에 입력 후 저장:  \n"
        '```\nGITHUB_TOKEN = "복사한_토큰"\nGITHUB_REPO = "hanallpharm-sys/hira-price-tracker"\n```')
