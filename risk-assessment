import base64
import io
from datetime import datetime, timedelta, timezone

import requests
import streamlit as st
from PIL import Image, ImageOps

st.set_page_config(page_title="위험성평가 사진 전송", page_icon="🦺", layout="centered")

KST = timezone(timedelta(hours=9))
SITES = ["본사", "용인", "문막"]
MAX_PHOTOS = 6


def shrink(raw: bytes, max_side: int = 1280) -> bytes:
    img = ImageOps.exif_transpose(Image.open(io.BytesIO(raw))).convert("RGB")
    img.thumbnail((max_side, max_side))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=82)
    return buf.getvalue()


st.markdown("""<style>
.block-container{max-width:640px;padding-top:1.5rem}
header[data-testid="stHeader"]{display:none}
[data-testid="stFormSubmitButton"] button{height:3.2rem;font-weight:700}
</style>""", unsafe_allow_html=True)

st.title("위험성평가 사진 전송")
st.caption("현장 사진을 올리면 AI가 위험요인을 분석하고, 보고서가 총무팀 SharePoint에 자동 저장됩니다.")

if "sent" not in st.session_state:
    st.session_state.sent = None

if st.session_state.sent:
    d = st.session_state.sent
    st.success(f"{d['site']} 사진 {d['n']}장을 전송했습니다. 분석과 저장에 몇 분 걸릴 수 있습니다.")
    st.info("보고서는 총무팀 SharePoint의 사업장 폴더에 저장됩니다. 이 화면은 닫으셔도 됩니다.")
    if st.button("새로 전송하기", use_container_width=True):
        st.session_state.sent = None
        st.rerun()
    st.stop()

code_required = st.secrets.get("ACCESS_CODE", "")

with st.form("send"):
    site = st.selectbox("사업장", SITES)
    assessor = st.text_input("이름", placeholder="촬영·제출하는 분 이름")
    dept = st.text_input("부서", placeholder="예: 생산1팀")
    task = st.text_area("평가할 작업·장소 설명", placeholder="예: 포장라인 컨베이어 주변, 지게차 통로", height=80)
    files = st.file_uploader(
        f"현장 사진 (최대 {MAX_PHOTOS}장)",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        help="휴대폰에서는 '사진 촬영'을 선택하면 카메라가 열립니다.",
    )
    code = st.text_input("접속 코드", type="password") if code_required else ""
    submit = st.form_submit_button("전송하기", use_container_width=True, type="primary")

if submit:
    errs = []
    if code_required and code != code_required:
        errs.append("접속 코드가 맞지 않습니다.")
    if not assessor.strip():
        errs.append("이름을 입력해 주세요.")
    if not files:
        errs.append("사진을 한 장 이상 올려 주세요.")
    if len(files) > MAX_PHOTOS:
        errs.append(f"사진은 최대 {MAX_PHOTOS}장까지 올릴 수 있습니다.")
    for e in errs:
        st.error(e)
    if not errs:
        try:
            photos = [shrink(f.getvalue()) for f in files]
        except Exception:
            st.error("사진을 읽지 못했습니다. 다른 사진으로 다시 시도해 주세요.")
            st.stop()
        body = {
            "site": site,
            "assessor": assessor.strip(),
            "dept": dept.strip(),
            "task": task.strip(),
            "date": datetime.now(KST).strftime("%Y-%m-%d"),
            "photos": [{"name": f"사진{i + 1}.jpg", "base64": base64.b64encode(p).decode()}
                       for i, p in enumerate(photos)],
        }
        with st.spinner("전송하고 있습니다."):
            try:
                r = requests.post(st.secrets["FLOW_URL"], json=body, timeout=90)
                r.raise_for_status()
            except Exception:
                st.error("전송에 실패했습니다. 잠시 뒤 다시 시도하고, 계속되면 총무팀에 알려 주세요.")
                st.stop()
        st.session_state.sent = {"site": site, "n": len(photos)}
        st.rerun()
