"""
streamlit_app.py
약제급여목록표 변동 추적 — Streamlit Community Cloud 배포용

monthly_data/ 폴더의 약제급여목록표 엑셀들을 읽어,
업체(기본: 한올바이오파마)의 월별 상한금액 변동을 매트릭스로 보여준다.
- 새 달 엑셀을 monthly_data/ 에 올리면(GitHub 웹 업로드) 앱이 자동 갱신.
"""
import os
import glob
import io
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


def load_notes():
    """notes.json(제품코드 -> 사유)을 읽는다. 세션에서 수정한 값이 있으면 우선."""
    notes = {}
    if os.path.exists(NOTES_PATH):
        try:
            with open(NOTES_PATH, encoding="utf-8") as f:
                notes = json.load(f)
        except Exception:
            notes = {}
    notes.update(st.session_state.get("notes_override", {}))
    return notes


def github_enabled():
    return "GITHUB_TOKEN" in st.secrets


def save_notes_github(notes: dict) -> bool:
    """notes.json을 GitHub 레포에 커밋(토큰 필요)."""
    token = st.secrets["GITHUB_TOKEN"]
    repo = st.secrets.get("GITHUB_REPO", "hanallpharm-sys/hira-price-tracker")
    branch = st.secrets.get("GITHUB_BRANCH", "main")
    url = f"https://api.github.com/repos/{repo}/contents/notes.json"
    headers = {"Authorization": f"Bearer {token}",
               "Accept": "application/vnd.github+json"}
    sha = None
    r = requests.get(url, headers=headers, params={"ref": branch}, timeout=20)
    if r.status_code == 200:
        sha = r.json().get("sha")
    content = base64.b64encode(
        json.dumps(notes, ensure_ascii=False, indent=2).encode("utf-8")
    ).decode()
    payload = {"message": "사유 업데이트", "content": content, "branch": branch}
    if sha:
        payload["sha"] = sha
    r = requests.put(url, headers=headers, json=payload, timeout=20)
    return r.status_code in (200, 201)


@st.cache_data(show_spinner="약제급여목록표를 읽는 중...")
def load_snapshots(file_sig):
    """monthly_data 내 모든 엑셀을 파싱해 {label: df} 반환. file_sig로 캐시 무효화."""
    snapshots = {}
    for path in sorted(glob.glob(os.path.join(DATA_DIR, "*.xls*"))):
        label = parser.detect_label(path) or os.path.splitext(os.path.basename(path))[0]
        try:
            snapshots[label] = parser.parse(path)
        except Exception as e:
            st.warning(f"{os.path.basename(path)} 읽기 실패: {e}")
    return snapshots


def file_signature():
    files = sorted(glob.glob(os.path.join(DATA_DIR, "*.xls*")))
    return tuple((os.path.basename(f), os.path.getmtime(f)) for f in files)


def style_matrix(mat, months):
    """매트릭스 셀에 색 입히기: X=빨강, 신규=초록, 인상=주황, 인하=파랑."""
    disp = mat.copy()
    disp = disp.rename(columns={"code": "제품코드", "name": "제품명", "status": "상태"})
    show_cols = ["제품명"] + months + ["상태"]
    disp = disp[show_cols]

    # 셀별 마크 계산
    marks = [company.cell_marks(r, months) for _, r in mat.iterrows()]

    def color_cells(_):
        styles = pd.DataFrame("", index=disp.index, columns=disp.columns)
        for i, mk in enumerate(marks):
            for m in months:
                k = mk[m]
                if k == "x":
                    styles.loc[disp.index[i], m] = "background-color:#fdeceb;color:#b42318;font-weight:700"
                elif k == "new":
                    styles.loc[disp.index[i], m] = "background-color:#e7f5ed;color:#0f7a4d;font-weight:700"
                elif k == "up":
                    styles.loc[disp.index[i], m] = "background-color:#fdf0e6;color:#b54708;font-weight:700"
                elif k == "down":
                    styles.loc[disp.index[i], m] = "background-color:#eaf1fa;color:#1f5fa8;font-weight:700"
        return styles

    def fmt(v):
        if v == "X" or pd.isna(v):
            return "X" if v == "X" else ""
        try:
            return f"{float(v):,.0f}"
        except (ValueError, TypeError):
            return v

    fmt_map = {m: fmt for m in months}
    return disp.style.apply(color_cells, axis=None).format(fmt_map)


# ───────────────────────── UI ─────────────────────────
st.title("💊 약제급여목록표 변동 추적")
st.caption("HIRA 약제급여목록 및 급여상한금액표 · 업체별 월별 상한금액 추적")

snapshots = load_snapshots(file_signature())

if not snapshots:
    st.info(
        "아직 데이터가 없습니다.\n\n"
        "GitHub 레포의 **monthly_data** 폴더에 약제급여목록표 엑셀(.xlsx)을 올리면 "
        "여기에 자동으로 나타납니다. 적용월은 파일명에서 자동 인식됩니다."
    )
    st.stop()

months = sorted(snapshots.keys())

# 사이드바
with st.sidebar:
    st.subheader("설정")
    maker = st.text_input("업체명 필터", value=company.DEFAULT_MAKER)
    st.markdown("---")
    st.markdown(f"**누적 월**: {len(months)}개")
    st.markdown(f"{months[0]} ~ {months[-1]}")
    st.markdown("---")
    st.markdown(
        "**범례**  \n"
        "🟥 X = 그 달에 없음(삭제/미등재)  \n"
        "🟩 신규 진입  \n"
        "🟧 ↑ 전월 대비 인상  \n"
        "🟦 ↓ 전월 대비 인하"
    )

# 매트릭스
mat, mon = company.build_matrix(maker, snapshots=snapshots)

if not len(mat):
    st.warning(f"'{maker}' 업체 제품을 찾지 못했습니다. 업체명을 확인해 주세요.")
    st.stop()

# 요약 지표
changed = mat[mat["status"] != "유지"]
n_add = changed["status"].str.contains("신규").sum()
n_rem = changed["status"].str.contains("삭제").sum()
n_chg = changed["status"].str.contains("변동").sum()
c1, c2, c3, c4 = st.columns(4)
c1.metric("전체 품목", f"{len(mat)}")
c2.metric("신규 등재", f"{int(n_add)}")
c3.metric("삭제", f"{int(n_rem)}")
c4.metric("상한금액 변동", f"{int(n_chg)}")

# 사유(notes) 병합
notes = load_notes()
mat["사유"] = mat["code"].map(lambda c: notes.get(c, "")).fillna("")

st.subheader(f"{maker} · 월별 상한금액")
view = st.radio("보기", ["변동 있는 품목만", "전체"], horizontal=True, label_visibility="collapsed")
table = changed.copy() if view == "변동 있는 품목만" else mat.copy()
table["사유"] = table["code"].map(lambda c: notes.get(c, "")).fillna("")
if not len(table):
    st.success("해당 기간 동안 변동(신규/삭제/금액변경)이 없습니다.")
else:
    st.dataframe(style_matrix(table, mon), use_container_width=True, height=480)

# ── 변동·삭제 사유 기록 (신규 제외) ──
st.markdown("### 변동·삭제 사유 기록")
mask = mat["status"].str.contains("삭제") | mat["status"].str.contains("변동")
editable = mat[mask][["code", "name", "status", "사유"]].copy()
editable = editable.rename(columns={"name": "제품명", "status": "상태"})

if len(editable) == 0:
    st.caption("기록할 변동/삭제 품목이 없습니다.")
elif github_enabled():
    st.caption("사유 칸을 클릭해 입력하고 아래 **저장** 버튼을 누르면 모두에게 반영됩니다.")
    edited = st.data_editor(
        editable, use_container_width=True, hide_index=True, key="notes_editor",
        column_config={
            "code": None,
            "제품명": st.column_config.TextColumn(disabled=True, width="large"),
            "상태": st.column_config.TextColumn(disabled=True, width="small"),
            "사유": st.column_config.TextColumn("사유", width="large"),
        },
    )
    if st.button("💾 사유 저장", type="primary"):
        new_notes = dict(notes)
        for _, r in edited.iterrows():
            txt = (r["사유"] or "").strip()
            if txt:
                new_notes[r["code"]] = txt
            elif r["code"] in new_notes:
                del new_notes[r["code"]]
        with st.spinner("저장 중..."):
            ok = save_notes_github(new_notes)
        if ok:
            st.session_state["notes_override"] = new_notes
            st.success("저장되었습니다. 잠시 후 모든 사용자에게 반영됩니다.")
        else:
            st.error("저장 실패. GitHub 토큰 권한(Contents: Read and write)을 확인해 주세요.")
else:
    st.info(
        "지금은 **읽기 전용**이에요. 앱에서 직접 입력하려면 GitHub 토큰 등록이 필요합니다(아래 안내).\n\n"
        "토큰 없이 쓰려면 레포의 **notes.json**을 GitHub에서 편집하세요. "
        '형식: `{\"제품코드\": \"사유\"}` — 예: `{\"645302132\": \"사용량-약가 연동 인하\"}`'
    )
    st.dataframe(editable.drop(columns="code"), use_container_width=True, hide_index=True)

# 엑셀 다운로드 (사유 포함)
tmp = "/tmp/_company_matrix.xlsx"
company.to_excel(maker, mat, mon, tmp)
with open(tmp, "rb") as f:
    data_bytes = f.read()
st.download_button(
    "📥 엑셀로 다운로드 (사유 포함)", data=data_bytes,
    file_name=f"{maker}_월별상한금액_{mon[0]}_{mon[-1]}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("ℹ️ 매달 업데이트 / 사유 입력용 토큰 등록 방법"):
    st.markdown(
        "**매달 데이터 추가**  \n"
        "1. GitHub 레포 → **monthly_data** 폴더 → Add file → Upload files로 새 달 엑셀 올리고 Commit  \n"
        "2. 1~2분 뒤 자동 반영(열 하나 추가)  \n\n"
        "**앱에서 직접 사유 입력하려면 (일회성 토큰 등록)**  \n"
        "1. GitHub 우상단 프로필 → Settings → Developer settings → "
        "Personal access tokens → **Fine-grained tokens** → Generate new token  \n"
        "2. 이 레포만 선택, 권한 **Contents: Read and write** 부여 → 토큰 생성·복사  \n"
        "3. Streamlit 앱 우상단 **⋮ → Settings → Secrets**에 입력 후 저장:  \n"
        '```\nGITHUB_TOKEN = \"복사한_토큰\"\nGITHUB_REPO = \"hanallpharm-sys/hira-price-tracker\"\n```\n'
        "4. 앱이 재시작되면 사유 칸이 입력 가능해집니다."
    )
