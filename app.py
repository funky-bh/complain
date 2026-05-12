"""Streamlit web UI for the ChannelTalk CS complaint analyzer."""

import io
import os
import tempfile

import bcrypt
import streamlit as st
import yaml
from dotenv import load_dotenv

load_dotenv()

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.yaml")

st.set_page_config(
    page_title="CS 컴플레인 분석기",
    page_icon="📊",
    layout="centered",
)


# ── Auth helpers ─────────────────────────────────────────────────────────────

def load_config() -> dict:
    if not os.path.exists(CONFIG_PATH):
        return {"users": {}}
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return yaml.safe_load(f) or {"users": {}}


def verify_password(username: str, password: str, config: dict) -> bool:
    users = config.get("users", {})
    if username not in users:
        return False
    stored = users[username].get("password", "")
    try:
        return bcrypt.checkpw(password.encode(), stored.encode())
    except Exception:
        return False


# ── Session state init ───────────────────────────────────────────────────────

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "username" not in st.session_state:
    st.session_state.username = ""


# ── Login page ───────────────────────────────────────────────────────────────

def show_login():
    st.title("CS 컴플레인 분석기")
    st.markdown("---")

    config = load_config()
    if not config.get("users"):
        st.error("등록된 사용자가 없습니다. 터미널에서 `python create_user.py`를 실행해 계정을 만들어주세요.")
        return

    with st.form("login_form"):
        st.subheader("로그인")
        username = st.text_input("아이디")
        password = st.text_input("비밀번호", type="password")
        submitted = st.form_submit_button("로그인", use_container_width=True)

    if submitted:
        if verify_password(username, password, config):
            st.session_state.authenticated = True
            st.session_state.username = username
            st.rerun()
        else:
            st.error("아이디 또는 비밀번호가 올바르지 않습니다.")


# ── Main app ─────────────────────────────────────────────────────────────────

def show_app():
    # Sidebar
    with st.sidebar:
        st.markdown(f"**{st.session_state.username}** 님")
        if st.button("로그아웃", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.username = ""
            st.rerun()
        st.markdown("---")
        st.caption("업로드한 파일은 분석 후 즉시 삭제됩니다.")

    st.title("CS 컴플레인 분석기")
    st.markdown("채널톡 내보내기 엑셀 파일을 업로드하면 분석된 결과 파일을 다운로드할 수 있습니다.")
    st.markdown("---")

    uploaded = st.file_uploader("엑셀 파일 선택 (.xlsx)", type=["xlsx"])

    if uploaded is None:
        st.info("파일을 업로드해주세요.")
        return

    st.success(f"파일 업로드 완료: **{uploaded.name}**")

    if st.button("분석 시작", type="primary", use_container_width=True):
        run_analysis(uploaded)


def run_analysis(uploaded):
    from analyze import analyze  # import here to avoid circular at top

    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        st.error("ANTHROPIC_API_KEY가 설정되지 않았습니다. .env 파일을 확인해주세요.")
        return

    progress_bar = st.progress(0, text="분석 준비 중...")
    status_text = st.empty()

    def on_progress(current, total, message):
        progress_bar.progress(current / total, text=f"분석 중... ({current}/{total})")
        status_text.text(message)

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_in:
        tmp_in.write(uploaded.getvalue())
        tmp_in_path = tmp_in.name

    tmp_out_path = tmp_in_path.replace(".xlsx", "_analyzed.xlsx")

    try:
        analyze(tmp_in_path, tmp_out_path, on_progress=on_progress)

        progress_bar.progress(1.0, text="완료!")
        status_text.empty()
        st.success("분석이 완료되었습니다!")

        with open(tmp_out_path, "rb") as f:
            result_bytes = f.read()

        original_stem = uploaded.name.rsplit(".", 1)[0]
        st.download_button(
            label="결과 파일 다운로드",
            data=result_bytes,
            file_name=f"{original_stem}_analyzed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

    except ValueError as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"오류: {e}")
    except Exception as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"예상치 못한 오류가 발생했습니다: {e}")
    finally:
        for path in (tmp_in_path, tmp_out_path):
            try:
                os.unlink(path)
            except FileNotFoundError:
                pass


# ── Entry point ──────────────────────────────────────────────────────────────

if st.session_state.authenticated:
    show_app()
else:
    show_login()
