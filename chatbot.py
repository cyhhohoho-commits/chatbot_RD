import streamlit as st
from anthropic import Anthropic
from dotenv import load_dotenv
import base64
import fitz
from pptx import Presentation
from docx import Document
import io
import os
import time

load_dotenv()
client = Anthropic()

with open("prompt.txt", "r", encoding="utf-8") as f:
    SYSTEM_PROMPT = f.read()

def _get_secret(key):
    """secrets.toml(클라우드) 또는 환경변수(.env) 어느 쪽이든 안전하게 읽기"""
    try:
        return st.secrets[key]
    except Exception:
        return os.environ.get(key)

def check_password():
    """APP_PASSWORD가 설정돼 있으면 비밀번호 잠금. 미설정(로컬)이면 통과."""
    app_pw = _get_secret("APP_PASSWORD")
    if not app_pw:
        return True
    if st.session_state.get("authenticated"):
        return True

    def _verify():
        if st.session_state.get("pw_input") == app_pw:
            st.session_state["authenticated"] = True
            st.session_state.pop("pw_input", None)
        else:
            st.session_state["authenticated"] = False

    st.markdown("### 🔒 RD부 업무 안내 챗봇")
    st.text_input("비밀번호를 입력하세요", type="password",
                  key="pw_input", on_change=_verify)
    if st.session_state.get("authenticated") is False:
        st.error("비밀번호가 올바르지 않습니다.")
    return False

def get_image_base64(image_path):
    try:
        with open(image_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except:
        return None

def extract_pptx_text(file_bytes):
    prs = Presentation(io.BytesIO(file_bytes))
    text = ""
    for i, slide in enumerate(prs.slides):
        text += f"\n[슬라이드 {i+1}]\n"
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                text += shape.text + "\n"
    return text

def extract_docx_text(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    text = ""
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    num_counts = {}  # {(numId, ilvl): count} — 자동번호매기기 카운터

    for child in doc.element.body:
        tag = child.tag

        if tag == f'{{{ns}}}p':
            para_text = ''.join(node.text or '' for node in child.iter()
                                if node.tag == f'{{{ns}}}t')

            # 자동번호매기기(numPr) 감지 후 번호 복원
            num_prefix = ''
            pPr = child.find(f'{{{ns}}}pPr')
            if pPr is not None:
                numPr = pPr.find(f'{{{ns}}}numPr')
                if numPr is not None:
                    ilvl_elem = numPr.find(f'{{{ns}}}ilvl')
                    numId_elem = numPr.find(f'{{{ns}}}numId')
                    if ilvl_elem is not None and numId_elem is not None:
                        ilvl = int(ilvl_elem.get(f'{{{ns}}}val', 0))
                        numId = int(numId_elem.get(f'{{{ns}}}val', 0))
                        if numId != 0:
                            key = (numId, ilvl)
                            num_counts[key] = num_counts.get(key, 0) + 1
                            num_prefix = f'{num_counts[key]}) '

            if para_text.strip():
                text += num_prefix + para_text + "\n"

        elif tag == f'{{{ns}}}tbl':
            rows = [r for r in child if r.tag == f'{{{ns}}}tr']
            num_rows = len(rows)
            if num_rows > 1:
                attr_count = num_rows - 1
                first_row_cells = [tc for tc in rows[0] if tc.tag == f'{{{ns}}}tc']
                choice_count = max(0, len(first_row_cells) - 1)
                text += f"[테이블: 속성 {attr_count}개, 보기 {choice_count}개]\n"
                for i, row in enumerate(rows):
                    cells = [tc for tc in row if tc.tag == f'{{{ns}}}tc']
                    cell_texts = []
                    for tc in cells:
                        t_text = ''.join(node.text or '' for node in tc.iter()
                                         if node.tag == f'{{{ns}}}t')
                        cell_texts.append(t_text.strip())
                    label = "[헤더]" if i == 0 else f"[속성{i}]"
                    text += f"  {label} {' | '.join(cell_texts)}\n"

    return text

def extract_pdf_text(file_bytes):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    text = ""
    for i, page in enumerate(doc):
        text += f"\n[페이지 {i+1}]\n"
        text += page.get_text()
    return text

def stream_ai_response(api_messages):
    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=8192,
        system=[{
            "type": "text",
            "text": SYSTEM_PROMPT,
            "cache_control": {"type": "ephemeral"}
        }],
        messages=api_messages
    ) as stream:
        buffer = ""
        for text in stream.text_stream:
            buffer += text
            if "\n\n" in buffer:
                parts = buffer.split("\n\n")
                for part in parts[:-1]:
                    yield part + "\n\n"
                buffer = parts[-1]
        if buffer:
            yield buffer

st.set_page_config(
    page_title="RD부 업무 안내 챗봇",
    page_icon="🤖",
    layout="wide"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&display=swap');
* { font-family: 'Noto Sans KR', sans-serif; }

.stApp { background-color: #f0f4f8; }
#MainMenu, footer { visibility: hidden; }

/* 사이드바 */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #002470 0%, #0057b8 100%) !important;
}
[data-testid="stSidebar"] * { color: white !important; }
[data-testid="stSidebar"] .stButton button {
    background: rgba(255,255,255,0.15) !important;
    border: 1px solid rgba(255,255,255,0.3) !important;
    color: white !important;
    border-radius: 10px !important;
}

/* 사용자 말풍선 - 오른쪽 */
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) {
    display: flex !important;
    flex-direction: row !important;
    justify-content: flex-end !important;
    background: transparent !important;
    box-shadow: none !important;
    padding: 4px 0 !important;
    gap: 0 !important;
}
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) 
[data-testid="stChatMessageContent"] {
    background: #0057b8 !important;
    color: white !important;
    border-radius: 18px 4px 18px 18px !important;
    box-shadow: 0 2px 8px rgba(0,87,184,0.25) !important;
    max-width: 55% !important;
    width: fit-content !important;
    padding: 16px 20px !important;
    margin: 0 !important;
    line-height: 1.6 !important;
    display: inline-block !important;
}
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) 
[data-testid="stChatMessageContent"] p {
    color: white !important;
    margin: 0 !important;
}
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) 
[data-testid="stChatMessageAvatarUser"] {
    display: none !important;
}

/* 챗봇 말풍선 - 왼쪽 */
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarAssistant"]) {
    background: transparent !important;
    box-shadow: none !important;
    padding: 4px 0 !important;
    align-items: flex-start !important;
}
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarAssistant"]) 
[data-testid="stChatMessageContent"] {
    background: white !important;
    color: #222 !important;
    border-radius: 4px 18px 18px 18px !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08) !important;
    max-width: 65% !important;
    padding: 12px 16px !important;
    margin-left: 8px !important;
}
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarAssistant"]) 
[data-testid="stChatMessageAvatarAssistant"] {
    background: linear-gradient(135deg, #003087, #0057b8) !important;
    border-radius: 50% !important;
}

/* 입력창 */
[data-testid="stChatInput"] textarea {
    border-radius: 12px !important;
    border: 1.5px solid #c0d4ec !important;
    background: white !important;
    font-size: 14px !important;
}

/* 파일 업로더 */
[data-testid="stFileUploader"] {
    background: white;
    border: 1.5px dashed #b0c8f0;
    border-radius: 12px;
    padding: 8px 14px;
}

/* 샘플 질문 버튼 */
.stButton button {
    border-radius: 10px !important;
    border: 1.5px solid #c0d4ec !important;
    background: white !important;
    color: #003087 !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    transition: all 0.15s !important;
}
.stButton button:hover {
    background: #0057b8 !important;
    color: white !important;
    border-color: #0057b8 !important;
}

/* 파일 배지 */
.file-badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: #e8f0ff;
    border: 1px solid #b0c8f0;
    border-radius: 8px;
    padding: 5px 12px;
    font-size: 12px;
    color: #003087;
    margin-bottom: 8px;
}

/* 사이드바 박스 */
.info-box {
    background: rgba(255,255,255,0.12);
    border-radius: 12px;
    padding: 14px;
    font-size: 13px;
    line-height: 1.8;
    margin-bottom: 12px;
}
.info-box .box-title {
    font-weight: 700;
    color: #90caf9 !important;
    margin-bottom: 6px;
    font-size: 13px;
}
.warning-box {
    background: rgba(220,50,50,0.2);
    border: 1px solid rgba(255,120,120,0.35);
    border-radius: 12px;
    padding: 14px;
    font-size: 12px;
    line-height: 1.8;
    margin-bottom: 12px;
}
.warning-box .box-title {
    font-weight: 700;
    color: #ffaaaa !important;
    margin-bottom: 6px;
}
</style>
""", unsafe_allow_html=True)

# 비밀번호 잠금 (APP_PASSWORD 설정 시에만 동작)
if not check_password():
    st.stop()

# 상태 초기화
if "messages" not in st.session_state:
    st.session_state.messages = []
if "pending_question" not in st.session_state:
    st.session_state.pending_question = None

SAMPLE_QUESTIONS = [
    "이 설문지는 어느 부서로 의뢰 해야 하나요?",
    "이 조사의 적립금은 대략 얼마정도 인가요?",
    "이지서베이 조사는 어디로 의뢰하나요?",
    "웹업 프로세스가 어떻게 되나요?",
]

# 사이드바
with st.sidebar:
    logo_b64 = get_image_base64("로고.png")
    if logo_b64:
        st.markdown(f"""
            <div style="background:white;border-radius:12px;padding:14px;
                        text-align:center;margin-bottom:16px;">
                <img src="data:image/png;base64,{logo_b64}"
                     style="width:100%;max-width:180px;">
            </div>
        """, unsafe_allow_html=True)

    st.markdown("""
    <div class="info-box">
        <div class="box-title">📌 문의 가능 내용</div>
        ✔ 조사 의뢰처 안내<br>
        &nbsp;&nbsp;&nbsp;<span style="font-size:11px;opacity:0.8;">RD부 · 온라인실사팀 결정</span><br>
        ✔ 설문지 검토<br>
        &nbsp;&nbsp;&nbsp;<span style="font-size:11px;opacity:0.8;">로직·오류 등 이상 여부 확인</span><br>
        ✔ 예상 적립금 안내<br>
        &nbsp;&nbsp;&nbsp;<span style="font-size:11px;opacity:0.8;">응답시간 기반 계산</span><br>
        ✔ 업무 일정 안내<br>
        &nbsp;&nbsp;&nbsp;<span style="font-size:11px;opacity:0.8;">웹업·실사·테이블 일정</span>
    </div>
    <div class="info-box">
        <div class="box-title">📎 첨부 가능 파일</div>
        🖼 이미지 (png, jpg)<br>
        📄 PDF<br>
        📊 PPT (pptx)<br>
        📝 워드 (docx)<br>
        <span style="font-size:11px;opacity:0.7;">* HWP는 PDF 변환 후 첨부</span>
    </div>
    <div class="warning-box">
        <div class="box-title">🔒 보안 주의사항</div>
        아래 내용은 입력/첨부 금지<br>
        ❌ 고객사명/기밀 프로젝트<br>
        ❌ 응답자 개인정보<br>
        ❌ 사내 기밀 문서<br>
        <span style="font-size:11px;opacity:0.75;">
        입력 내용은 외부 AI 서버를 경유합니다</span>
    </div>
    """, unsafe_allow_html=True)

    if st.button("🔄 대화 초기화", use_container_width=True):
        st.session_state.messages = []
        st.session_state.pending_question = None
        st.rerun()

# 메인 헤더
st.markdown("""
<div style="background:linear-gradient(135deg,#002470,#0057b8);
            padding:18px 28px;border-radius:14px;margin-bottom:20px;">
    <div style="display:flex;align-items:center;gap:12px;">
        <div style="width:42px;height:42px;background:rgba(255,255,255,0.2);
                    border-radius:12px;display:flex;align-items:center;
                    justify-content:center;font-size:22px;">🤖</div>
        <div>
            <div style="color:white;font-size:18px;font-weight:700;">
                RD부 업무 안내 챗봇</div>
            <div style="color:rgba(255,255,255,0.8);font-size:13px;">
                설문조사 의뢰 및 RD부 업무 관련 문의를 도와드립니다</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# 샘플 질문
if not st.session_state.messages:
    st.markdown("#### 💬 자주 묻는 질문")
    cols = st.columns(2)
    for i, q in enumerate(SAMPLE_QUESTIONS):
        with cols[i % 2]:
            if st.button(q, key=f"sq_{i}", use_container_width=True):
                st.session_state.pending_question = q
                st.rerun()
    st.markdown("---")

# 이전 대화 표시
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# 파일 업로더
uploaded_file = st.file_uploader(
    "📎 파일을 드래그하거나 클릭해서 첨부 (이미지 / PDF / PPT / 워드)",
    type=["png", "jpg", "jpeg", "pdf", "pptx", "docx"],
)
if uploaded_file:
    st.markdown(
        f'<div class="file-badge">📎 {uploaded_file.name} '
        f'({round(uploaded_file.size/1024,1)}KB)</div>',
        unsafe_allow_html=True
    )

# pending_question 처리
user_input = st.chat_input("문의 내용을 입력하세요...")
if st.session_state.pending_question:
    user_input = st.session_state.pending_question
    st.session_state.pending_question = None

# 메시지 처리
if user_input:
    file_text = ""
    is_image = False
    image_data = None
    image_type = None

    if uploaded_file:
        file_bytes = uploaded_file.read()
        file_name = uploaded_file.name.lower()
        if file_name.endswith((".png", ".jpg", ".jpeg")):
            is_image = True
            image_data = base64.b64encode(file_bytes).decode()
            image_type = uploaded_file.type
        elif file_name.endswith(".pdf"):
            file_text = extract_pdf_text(file_bytes)
        elif file_name.endswith(".pptx"):
            file_text = extract_pptx_text(file_bytes)
        elif file_name.endswith(".docx"):
            file_text = extract_docx_text(file_bytes)

    if is_image and image_data:
        api_message = {
            "role": "user",
            "content": [
                {"type": "image", "source": {
                    "type": "base64",
                    "media_type": image_type,
                    "data": image_data
                }},
                {"type": "text",
                 "text": f"첨부 이미지를 참고해서 답변해줘.\n\n{user_input}"}
            ]
        }
        display_content = f"[이미지 첨부: {uploaded_file.name}]\n{user_input}"
    elif file_text:
        api_message = {
            "role": "user",
            "content": f"첨부 파일 내용:\n{file_text}\n\n"
                       f"위 내용을 참고해서 답변해줘.\n\n{user_input}"
        }
        display_content = f"[파일 첨부: {uploaded_file.name}]\n{user_input}"
    else:
        api_message = {"role": "user", "content": user_input}
        display_content = user_input

    st.session_state.messages.append({
        "role": "user", "content": display_content
    })

    api_messages = []
    for msg in st.session_state.messages[:-1]:
        api_messages.append({
            "role": msg["role"], "content": msg["content"]
        })
    api_messages.append(api_message)

    with st.chat_message("assistant"):
        has_file = bool(file_text or is_image)
        if has_file:
            stages = [
                "파일 읽는 중",
                "설문 구조 분석",
                "답변 작성",
            ]
            progress_placeholder = st.empty()
            for i, stage in enumerate(stages):
                pct = int((i + 1) / len(stages) * 100)
                bar_html = (
                    f'<div style="font-size:0.85em; color:#555; margin-bottom:4px;">'
                    f'분석 중... ({i+1}/{len(stages)}단계) — {stage}</div>'
                    f'<div style="background:#e0e0e0; border-radius:6px; height:10px; width:100%;">'
                    f'<div style="background:linear-gradient(90deg,#6C63FF,#48C9B0); '
                    f'width:{pct}%; height:10px; border-radius:6px; transition:width 0.3s;"></div>'
                    f'</div>'
                )
                progress_placeholder.markdown(bar_html, unsafe_allow_html=True)
                time.sleep(0.25)
            progress_placeholder.empty()
        answer = st.write_stream(stream_ai_response(api_messages))

    st.session_state.messages.append({
        "role": "assistant", "content": answer
    })
    st.rerun()