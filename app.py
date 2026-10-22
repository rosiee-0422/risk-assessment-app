import base64
import io
import json
import re
from datetime import datetime, timedelta, timezone

import requests
import streamlit as st
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image, ImageOps

st.set_page_config(page_title="위험성평가 사진 전송", page_icon="🦺", layout="centered")

KST = timezone(timedelta(hours=9))
SITES = ["본사", "용인", "문막"]
KINDS = ["최초", "정기", "수시"]
PARTS = ["순회 점검", "설문조사", "면담"]
MAX_PHOTOS = 6
F_SCALE = ["거의 발생하지 않음", "발생 가능성 낮음", "가끔 발생할 수 있음", "자주 발생할 수 있음", "매우 자주 발생함"]
S_SCALE = ["경미 · 응급처치 수준", "중간 · 휴업이 따르는 부상·질병", "중대 · 중증·영구장애", "치명 · 사망"]
ALLOW, URGENT = 6, 16
FOOT = ("근거: 산업안전보건법 제36조, 같은 법 시행규칙 제37조, 제37조의2(근로자 참여), 제37조의3(근로자 공유), "
        "제37조의4(기록·보존, 3년) (2026. 6. 1. 시행).")
DISCLAIMER = ("이 보고서의 유해·위험요인, 개선대책, 근거법령은 AI가 사진을 바탕으로 제안한 초안입니다. "
              "담당자가 현장에서 확인하고 국가법령정보센터(law.go.kr)에서 조문을 확인한 뒤 확정하세요.")


# ───────── 사진 ─────────
def shrink(raw: bytes, max_side: int = 1280) -> bytes:
    img = ImageOps.exif_transpose(Image.open(io.BytesIO(raw))).convert("RGB")
    img.thumbnail((max_side, max_side))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=82)
    return buf.getvalue()


# ───────── AI 분석 (Power Automate 플로우 A 호출) ─────────
def clamp(v, default, mx):
    try:
        v = round(float(v))
    except (TypeError, ValueError):
        return default
    return v if 1 <= v <= mx else default


def parse_result(text, photo_no):
    m = re.search(r"\{.*\}", text or "", re.S)
    if not m:
        raise ValueError("no json")
    data = json.loads(m.group(0))
    hazards = []
    for h in (data.get("hazards") or [])[:6]:
        if not h or not h.get("hazard"):
            continue
        f, s = clamp(h.get("freq"), 3, 5), clamp(h.get("sev"), 2, 4)
        laws = []
        for l in h.get("laws") or []:
            if not l or not l.get("name"):
                continue
            t = l["name"] + (f" {l['article']}" if l.get("article") else "") + (f" ({l['gist']})" if l.get("gist") else "")
            laws.append(t if (l.get("certain") is not False and l.get("article")) else t + " [조항 확인 필요]")
        hazards.append({
            "photo": photo_no, "process": str(h.get("process") or ""), "hazard": str(h["hazard"]),
            "current": str(h.get("current") or ""), "freq": f, "sev": s,
            "measure": str(h.get("measure") or ""), "laws": laws,
            "afterFreq": clamp(h.get("afterFreq"), max(1, f - 1), 5), "afterSev": clamp(h.get("afterSev"), s, 4),
        })
    return {"summary": str(data.get("summary") or ""), "unsure": str(data.get("unsure") or ""), "hazards": hazards}


def analyze_photo(photo_bytes, photo_no):
    r = requests.post(st.secrets["ANALYZE_URL"], json={"photo": base64.b64encode(photo_bytes).decode()}, timeout=110)
    r.raise_for_status()
    body = r.json()
    text = body.get("text") if isinstance(body, dict) and "text" in body else json.dumps(body, ensure_ascii=False)
    return parse_result(text, photo_no)


# ───────── Word 보고서 (기존 HTML 보고서와 같은 양식) ─────────
def level(score):
    if score <= ALLOW:
        return "허용 가능", "DDEFD9"
    if score >= URGENT:
        return "즉시 개선", "F6D2D0"
    return "개선 필요", "FBEBC4"


def shade(c, hex_fill):
    tcPr = c._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def cell(c, text, bold=False, size=8.5, center=False, fill=None):
    c.text = ""
    p = c.paragraphs[0]
    lines = str(text).split("\n")
    for i, line in enumerate(lines):
        r = p.add_run(line)
        r.bold = bold
        r.font.size = Pt(size)
        if i < len(lines) - 1:
            r.add_break()
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if fill:
        shade(c, fill)


BORDER = "C9CED6"


def borders(tbl, color=BORDER, sz=4):
    tblPr = tbl._tbl.tblPr
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        x = OxmlElement(f"w:{e}")
        x.set(qn("w:val"), "single")
        x.set(qn("w:sz"), str(sz))
        x.set(qn("w:space"), "0")
        x.set(qn("w:color"), color)
        b.append(x)
    tblPr.append(b)


def para_border(p, side="bottom", color="222222", sz=8):
    pPr = p._p.get_or_add_pPr()
    pb = OxmlElement("w:pBdr")
    x = OxmlElement(f"w:{side}")
    x.set(qn("w:val"), "single")
    x.set(qn("w:sz"), str(sz))
    x.set(qn("w:space"), "1")
    x.set(qn("w:color"), color)
    pb.append(x)
    pPr.append(pb)


def widths(tbl, ws):
    tbl.autofit = False
    tblPr = tbl._tbl.tblPr
    for tag in ("w:tblW", "w:tblLayout"):
        e = tblPr.find(qn(tag))
        if e is not None:
            tblPr.remove(e)
    w = OxmlElement("w:tblW")
    w.set(qn("w:w"), str(int(sum(ws) * 567)))
    w.set(qn("w:type"), "dxa")
    tblPr.append(w)
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    for gc, x in zip(tbl._tbl.tblGrid.findall(qn("w:gridCol")), ws):
        gc.set(qn("w:w"), str(int(x * 567)))
    for row in tbl.rows:
        for j, x in enumerate(ws):
            if j < len(row.cells):
                row.cells[j].width = Cm(x)


def heading(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(12)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    para_border(p, "bottom")
    p.paragraph_format.keep_with_next = True


def note(doc, text, size=8.5, bold_label=None):
    p = doc.add_paragraph()
    if bold_label:
        r = p.add_run(bold_label)
        r.bold = True
        r.font.size = Pt(size)
    r = p.add_run(text)
    r.font.size = Pt(size)
    p.paragraph_format.space_before = Pt(3)
    return p


def build_docx(info, result, photos) -> bytes:
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(1.5)
    sec.top_margin = sec.bottom_margin = Cm(1.5)
    st_ = doc.styles["Normal"]
    st_.font.name = "맑은 고딕"
    st_.element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    st_.font.size = Pt(10)

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("위험성평가 결과 보고서")
    r.bold = True
    r.font.size = Pt(18)

    hz = result["hazards"]
    bad = sum(1 for h in hz if h["freq"] * h["sev"] > ALLOW)

    # 결재란
    sg = doc.add_table(rows=2, cols=2)
    sg.style = "Table Grid"
    sg.alignment = WD_TABLE_ALIGNMENT.RIGHT
    cell(sg.rows[0].cells[0], "담당자", True, 9, True, "EEF1F4")
    cell(sg.rows[0].cells[1], "팀장", True, 9, True, "EEF1F4")
    cell(sg.rows[1].cells[0], info["assessor"], False, 9, True)
    cell(sg.rows[1].cells[1], "", False, 9, True)
    sg.rows[1].height = Cm(1.2)
    borders(sg)
    widths(sg, [2.6, 2.6])

    heading(doc, "1. 평가 개요")
    ov = doc.add_table(rows=5, cols=4)
    ov.style = "Table Grid"
    rows = [
        ("사업장", info["site"], "담당 부서", info["dept"] or "-"),
        ("평가 구분", f"{info['kind']}평가", "평가일", info["date"]),
        ("평가 담당자", info["assessor"] or "-", "평가 방법",
         f"빈도·강도 조합 (발생가능성 {len(F_SCALE)}단계 × 심각성 {len(S_SCALE)}단계), 사진 기반 AI 초안 작성 후 담당자 검토"),
        ("대상 작업·장소", info["task"] or "-", "확정 상태", "작성 중 (AI 초안)"),
    ]
    for i, rw in enumerate(rows):
        for j, v in enumerate(rw):
            cell(ov.rows[i].cells[j], v, bold=(j % 2 == 0), size=9, fill="EEF1F4" if j % 2 == 0 else None)
    cell(ov.rows[4].cells[0], "평가 결과", True, 9, fill="EEF1F4")
    mc = ov.rows[4].cells[1].merge(ov.rows[4].cells[3])
    borders(ov)
    widths(ov, [3.0, 9.0, 3.0, 11.7])
    cell(mc, f"유해·위험요인 {len(hz)}건, 허용 불가 수준 {bad}건" + (f"\n{result['summary']}" if result["summary"] else ""), size=9)

    heading(doc, "2. 평가 기준")
    cr = doc.add_table(rows=1, cols=3)
    cr.style = "Table Grid"
    for j, v in enumerate(["구분", "단계", "기준"]):
        cell(cr.rows[0].cells[j], v, True, 9, True, "EEF1F4")
    for name, arr in (("발생가능성", F_SCALE), ("심각성", S_SCALE)):
        for k, txt in enumerate(arr):
            c = cr.add_row().cells
            cell(c[0], name if k == 0 else "", True, 9, fill="EEF1F4")
            cell(c[1], k + 1, False, 9, True)
            cell(c[2], txt, False, 9)
    borders(cr)
    widths(cr, [4.0, 2.5, 20.2])
    note(doc, f"위험성 = 발생가능성 × 심각성 (1~{len(F_SCALE) * len(S_SCALE)}점). {ALLOW}점 이하는 허용 가능, "
              f"{URGENT}점 이상은 즉시 개선, 그 사이는 개선 필요로 보고 개선대책을 수립·이행합니다.")

    heading(doc, "3. 위험성평가표")
    if hz:
        heads = ["번호", "작업·장소", "유해·위험요인", "현재 안전보건조치", "빈도", "강도", "위험성", "개선대책",
                 "근거법령", "개선 후 위험성", "담당자·예정일", "이행"]
        tb = doc.add_table(rows=1, cols=len(heads))
        tb.style = "Table Grid"
        for j, v in enumerate(heads):
            cell(tb.rows[0].cells[j], v, True, 8, True, "EEF1F4")
        for i, h in enumerate(hz):
            b, a = h["freq"] * h["sev"], h["afterFreq"] * h["afterSev"]
            lb, fb = level(b)
            la, fa = level(a)
            c = tb.add_row().cells
            cell(c[0], i + 1, center=True)
            cell(c[1], h["process"] + (f"\n(사진 {h['photo']})" if h["photo"] else ""))
            cell(c[2], h["hazard"])
            cell(c[3], h["current"])
            cell(c[4], h["freq"], center=True)
            cell(c[5], h["sev"], center=True)
            cell(c[6], f"{b}\n{lb}", bold=True, center=True, fill=fb)
            cell(c[7], h["measure"])
            cell(c[8], "\n".join("· " + x for x in h["laws"]) or "-")
            cell(c[9], f"{a}\n{la}", bold=True, center=True, fill=fa)
            cell(c[10], "-")
            cell(c[11], "미완료", center=True)
        borders(tb)
        widths(tb, [1.0, 2.4, 3.8, 3.2, 1.0, 1.0, 1.8, 4.0, 4.0, 1.8, 1.7, 1.3])
    else:
        doc.add_paragraph("등록된 위험요인이 없습니다.")
    if result["unsure"]:
        note(doc, " " + result["unsure"], bold_label="사진만으로 판단하기 어려운 점:")

    heading(doc, "4. 현장 사진")
    if photos:
        cols = 3
        pt = doc.add_table(rows=(len(photos) + cols - 1) // cols, cols=cols)
        pt.style = "Table Grid"
        borders(pt)
        for i, p in enumerate(photos):
            c = pt.rows[i // cols].cells[i % cols]
            c.text = ""
            para = c.paragraphs[0]
            para.add_run().add_picture(io.BytesIO(p), width=Cm(7.6))
            cap = c.add_paragraph(f"사진 {i + 1}")
            cap.runs[0].font.size = Pt(8.5)
    else:
        doc.add_paragraph("등록된 사진이 없습니다.")

    heading(doc, "5. 근로자 참여 및 결과 공유")
    pr = doc.add_table(rows=2, cols=4)
    pr.style = "Table Grid"
    borders(pr)
    widths(pr, [3.5, 9.8, 3.5, 9.9])
    vals = [("참여 근로자", info["workers"] or "-", "근로자대표", info["rep"] or "-"),
            ("참여 방법", ", ".join(info["part"]) or "-", "결과 공유", info["share"] or "-")]
    for i, rw in enumerate(vals):
        for j, v in enumerate(rw):
            cell(pr.rows[i].cells[j], v, bold=(j % 2 == 0), size=9, fill="EEF1F4" if j % 2 == 0 else None)

    fp = note(doc, FOOT, 8)
    para_border(fp, "top", BORDER, 6)
    fp.paragraph_format.space_before = Pt(10)
    n = note(doc, DISCLAIMER, 8)
    n.runs[0].font.color.rgb = RGBColor(0x5A, 0x64, 0x70)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ───────── SharePoint 저장 (Power Automate 플로우 B 호출) ─────────
def push_flow(site, file_name, docx_bytes, photos=None):
    # 사진은 Word 보고서 안에 들어 있으므로 SharePoint에는 .docx만 저장한다.
    body = {"site": site, "fileName": file_name, "docxBase64": base64.b64encode(docx_bytes).decode(), "photos": []}
    r = requests.post(st.secrets["UPLOAD_URL"], json=body, timeout=110)
    r.raise_for_status()


# ───────── 화면 ─────────
st.markdown("""<style>
.block-container{max-width:640px;padding-top:1.5rem}
header[data-testid="stHeader"]{display:none}
[data-testid="stFormSubmitButton"] button{height:3.2rem;font-weight:700}
</style>""", unsafe_allow_html=True)

st.title("위험성평가 사진 전송")
st.caption("현장 사진을 올리면 AI가 위험요인을 분석하고, 보고서가 총무팀 SharePoint에 자동 저장됩니다.")

if "done" not in st.session_state:
    st.session_state.done = None

if st.session_state.done:
    d = st.session_state.done
    st.success(f"{d['site']} · 위험요인 {d['n']}건을 분석했고 보고서를 저장했습니다.")
    if d.get("warn"):
        st.warning(d["warn"])
    if st.button("새로 전송하기", use_container_width=True):
        st.session_state.done = None
        st.rerun()
    st.stop()

code_required = st.secrets.get("ACCESS_CODE", "")

with st.form("send"):
    site = st.selectbox("사업장", SITES)
    assessor = st.text_input("평가 담당자(이름)", placeholder="촬영·제출하는 분 이름")
    dept = st.text_input("담당 부서", placeholder="예: 생산1팀")
    kind = st.selectbox("평가 구분", KINDS, index=1)
    task = st.text_area("평가할 작업·장소 설명", placeholder="예: 포장라인 컨베이어 주변, 지게차 통로", height=80)
    files = st.file_uploader(f"현장 사진 (최대 {MAX_PHOTOS}장)", type=["jpg", "jpeg", "png", "webp"],
                             accept_multiple_files=True,
                             help="휴대폰에서는 '사진 촬영'을 선택하면 카메라가 열립니다.")
    with st.expander("근로자 참여 및 결과 공유 (선택)"):
        workers = st.text_input("참여 근로자", placeholder="이름을 쉼표로 구분")
        rep = st.text_input("근로자대표")
        part = st.multiselect("참여 방법", PARTS, default=["순회 점검"])
        share = st.text_input("결과 공유 방법·일자", placeholder="예: 게시판 게시, 2026-10-23 안전교육")
    code = st.text_input("접속 코드", type="password") if code_required else ""
    submit = st.form_submit_button("분석하고 저장하기", use_container_width=True, type="primary")

if submit:
    errs = []
    if code_required and code != code_required:
        errs.append("접속 코드가 맞지 않습니다.")
    if not assessor.strip():
        errs.append("평가 담당자 이름을 입력해 주세요.")
    if not files:
        errs.append("사진을 한 장 이상 올려 주세요.")
    if len(files) > MAX_PHOTOS:
        errs.append(f"사진은 최대 {MAX_PHOTOS}장까지 올릴 수 있습니다.")
    for e in errs:
        st.error(e)
    if not errs:
        info = {"site": site, "dept": dept.strip(), "assessor": assessor.strip(), "kind": kind,
                "task": task.strip(), "date": datetime.now(KST).strftime("%Y-%m-%d"),
                "workers": workers.strip(), "rep": rep.strip(), "part": part, "share": share.strip()}
        try:
            photos = [shrink(f.getvalue()) for f in files]
        except Exception:
            st.error("사진을 읽지 못했습니다. 다른 사진으로 다시 시도해 주세요.")
            st.stop()

        result = {"summary": "", "unsure": "", "hazards": []}
        bar = st.progress(0.0, text="사진을 분석하고 있습니다.")
        failed = []
        for i, p in enumerate(photos):
            bar.progress(i / len(photos), text=f"사진 {i + 1}/{len(photos)} 분석 중… (장당 10~30초)")
            try:
                r = analyze_photo(p, i + 1)
            except Exception:
                failed.append(i + 1)
                continue
            result["hazards"] += r["hazards"]
            if r["summary"]:
                result["summary"] += (" " if result["summary"] else "") + f"[사진 {i + 1}] {r['summary']}"
            if r["unsure"]:
                result["unsure"] += (" " if result["unsure"] else "") + f"[사진 {i + 1}] {r['unsure']}"
        bar.progress(1.0, text="분석을 마쳤습니다.")
        if len(failed) == len(photos):
            st.error("AI 분석에 실패했습니다. 잠시 뒤 다시 시도하고, 계속되면 총무팀에 알려 주세요.")
            st.stop()

        stamp = datetime.now(KST).strftime("%Y-%m-%d_%H%M%S")
        file_name = f"위험성평가_{site}_{stamp}.docx"
        warn = None
        with st.spinner("보고서를 만들어 SharePoint에 저장하고 있습니다."):
            try:
                push_flow(site, file_name, build_docx(info, result, photos), photos)
            except Exception:
                warn = "분석은 끝났지만 SharePoint 저장에 실패했습니다. 총무팀에 알려 주세요."
        if failed and not warn:
            warn = f"사진 {', '.join(map(str, failed))}번은 분석하지 못해 보고서에서 제외했습니다."
        st.session_state.done = {"site": site, "n": len(result["hazards"]), "warn": warn}
        st.rerun()
