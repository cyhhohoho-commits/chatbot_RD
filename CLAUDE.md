# CLAUDE.md

RD부 업무 안내 챗봇 프로젝트. 이 파일은 Claude Code가 작업 시 먼저 읽는 지침입니다.

## 프로젝트 개요
- 마크로밀엠브레인 **RD부 업무 안내 챗봇** (Streamlit + Claude API)
- 타부서 직원이 설문조사 의뢰처/ISAS 가능 여부/적립금/업무 프로세스를 문의하면 답변
- 핵심: 설문지(이미지/PDF/PPT/워드)를 첨부받아 **ISAS(RD부 워드 기반 솔루션)로 작업 가능한지** 판단

## 파일 구조
| 파일 | 역할 |
|------|------|
| `chatbot.py` | Streamlit 앱 본체. 파일 파싱 + Claude 스트리밍 호출 + UI(라이트/다크 테마) |
| `prompt.txt` | **챗봇이 쓰는 시스템 프롬프트** (답변 원칙·답변 규칙·ISAS 케이스가 모두 여기 있음) |
| `.streamlit/config.toml` | Streamlit 기본 테마 (입력창 포커스 색 등 위젯 색상) |
| `로고.png` | 사이드바 로고 겸 브라우저 탭 아이콘 |
| `cross_v5.json` | (참고 데이터) |
| `ISAS 가능문항들.docx/pdf` | ISAS 가능 케이스 원본 참고 자료 |
| `ISAS불가능_*.png`, `스크린샷 *.png` | 사용자가 보내주는 ISAS **불가능 케이스 캡처** |
| `백업`, `사용전 파일`, `사용파일` | 작업 폴더 (`백업`에 prompt.txt 날짜별 백업 보관) |
| `챗봇자료` | 프롬프트 반영용 원본 교육자료 (2022~2023 PPT/엑셀). **2026-06-11 요약 반영 완료** — 신규 내용만 추가, 충돌 시 기존 프롬프트 우선 원칙 적용 |

> **중요**: `chatbot.py`와 `prompt.txt`를 혼동하지 말 것.
> - 챗봇의 *답변 내용·판단 규칙·답변 문체*를 바꾸려면 → **`prompt.txt`**
> - 앱의 *동작·UI·파일 파싱·테마*를 바꾸려면 → **`chatbot.py`**

## UI / 디자인 구조 (2026-06 리디자인)

Claude.ai 채팅창 스타일. 차분한 크림 톤 + 다크모드 지원. 이모지 미사용이 디자인 원칙.

- **테마 전환**: 모든 색상은 `chatbot.py` 상단의 `LIGHT_THEME` / `DARK_THEME` 딕셔너리에서
  CSS 변수(`--bg`, `--text` 등)로 주입된다. **색 조정은 이 딕셔너리 값만 바꾸면 됨.**
  다크모드는 우상단 고정 토글 버튼(`key="theme_toggle"`, CSS `.st-key-theme_toggle`)으로
  전환하며, 세션 단위라 새로고침하면 라이트 모드로 초기화된다.
- **파일 첨부**: 별도 업로더 없음. `st.chat_input(accept_file=True, file_type=[...])`로
  채팅창 일체형 (드래그 앤 드롭 지원). 반환값은 `.text`와 `.files`로 분리해서 처리.
  파일만 첨부하고 텍스트가 없으면 "첨부한 파일을 검토해 주세요."로 자동 대체.
- **말풍선**: 사용자는 연한 박스(오른쪽), 챗봇 답변은 말풍선 없이 본문 텍스트. 아바타 숨김.
- **첫 화면**: 대화가 없을 때만 세리프체 인사말(`.welcome`) + 추천 질문 칩 표시.

### CSS 수정 시 주의 (과거에 실제로 깨졌던 부분)
1. **전역 폰트 오버라이드 금지 범위**: Pretendard를 전역 적용하면 Streamlit의
   Material Symbols 아이콘 폰트까지 덮어써서 아이콘이 `keyboard_double_arrow_left` 같은
   글자로 노출된다. `[data-testid="stIconMaterial"]` 등 아이콘 예외 규칙을 지울 것.
   → 지우면 안 되는 규칙: "Material 아이콘 폰트는 덮어쓰지 않기" 블록
2. **stToolbar 숨김의 부작용**: Deploy 버튼을 숨기려고 `stToolbar`를 `visibility: hidden`
   처리하면, 그 안에 있는 **사이드바 다시 열기(>>) 버튼**도 같이 숨겨진다.
   → 지우면 안 되는 규칙: `[data-testid="stExpandSidebarButton"] { visibility: visible }`
3. **토글 버튼 z-index**: Streamlit 헤더가 z-index 999990으로 클릭을 가로채므로
   `.st-key-theme_toggle`은 그보다 높아야 한다 (현재 1000000).

## 실행 방법
가상환경(`.venv`)을 쓰며 PowerShell로 실행한다 (Bash 경로는 인식 안 됨).

```powershell
Start-Process -FilePath "D:\python\작업\플젝3_챗봇\.venv\Scripts\streamlit.exe" `
  -ArgumentList "run", "chatbot.py" `
  -WorkingDirectory "D:\python\작업\플젝3_챗봇" -PassThru
```

- 접속: http://localhost:8501
- API 키는 `.env`(`load_dotenv()`)에서 로드. 모델은 `claude-sonnet-4-6`, 시스템 프롬프트는 프롬프트 캐싱 적용됨.
- Streamlit 1.57 기준. `st.chat_input(accept_file=...)`은 1.43 이상 필요.

## 배포 환경 (클라우드)

이 챗봇은 **Streamlit Community Cloud**에 배포되어 팀원들이 PC 무관하게 사용한다.

- **배포 URL**: https://chatbotrd-rhhzu5wctox6farl7vwhmj.streamlit.app
- **GitHub 저장소**: `cyhhohoho-commits/chatbot_RD` (origin/main), 계정 cyhhohoho-commits
- **클라우드가 읽는 키/비번**: Streamlit 앱 Settings → Secrets 에 설정
  - `ANTHROPIC_API_KEY` = API 키 (로컬 `.env`와 동일)
  - `APP_PASSWORD` = 팀 공용 접속 비밀번호 (`chatbot.py`의 `check_password()`가 사용)
- **로컬 vs 클라우드**: 로컬은 `.env` + 비번 없음으로 작동, 클라우드는 Secrets + 비번 잠금으로 작동. `chatbot.py`의 `_get_secret()`이 양쪽을 자동 처리.
- Python 버전은 Streamlit 앱 설정에서 3.12 권장 (3.14는 일부 라이브러리 설치 실패 가능).

### 배포 파일 (GitHub에 올라가는 것만)
`chatbot.py`, `prompt.txt`, `requirements.txt`, `로고.png`, `CLAUDE.md`, `README.md`, `.gitignore`, `.streamlit/config.toml`
→ `.env`·자료 파일(docx/xlsx/pdf/json)·ISAS 캡처 png 는 `.gitignore`로 **제외** (특히 `.env`는 절대 push 금지).

## 챗봇 내용 수정 → 클라우드 반영 흐름 (중요)

배포된 챗봇은 **GitHub의 `prompt.txt`만 읽는다.** 로컬에서 고치는 것만으로는 반영 안 됨.
수정 후 반드시 GitHub에 push해야 하며, push하면 Streamlit이 자동 감지해 1~2분 뒤 챗봇이 갱신된다.

```
① 로컬에서 prompt.txt(또는 chatbot.py) 수정
        ↓
② GitHub에 push
        ↓
③ Streamlit Cloud가 자동 감지 → 1~2분 뒤 챗봇 자동 업데이트
```

**사용자가 "고치고 배포해줘 / 반영해줘"라고 하면** = ①+② 까지 Claude가 수행한다.
**사용자가 push를 명시적으로 요청하기 전에는 push하지 않는다** (로컬 수정까지만).

```powershell
cd "D:\python\작업\플젝3_챗봇"
git add prompt.txt        # 또는 변경한 파일
git commit -m "프롬프트 수정: <무엇을>"
git push
```

push 후 사용자에게 "1~2분 뒤 클라우드 챗봇에 반영됩니다 (배포 URL)"라고 안내한다.

주의:
- push 전 `git status`로 `.env`가 staged 안 됐는지 확인 (`.gitignore`로 막혀 있지만 재확인).
- 저장소가 Private 권장. Private여도 배포된 앱은 정상 작동.
- 커밋 메시지는 무엇을 바꿨는지 한 줄로 명확히.

## prompt.txt 구조 (수정 시 참고)

prompt.txt는 크게 다음 순서로 구성된다. 새 내용은 맞는 섹션에 넣을 것.

1. **답변 원칙** — 질문 의도 파악, 묻는 것에만 답하기, 답변 길이 비례, 이모지 금지.
   답변 문체/형식 관련 수정은 여기.
2. **용어 정의 / 부서 소개 / 조사 의뢰 프로세스 / 조사 유형별 담당 부서**
3. **진행 방식 판단 기준** (ISAS vs 수작업 vs 혼합)
4. **ISAS 작업 가능 케이스** / **RD부 수작업 가능 케이스** / **ISAS 미지원 케이스**
5. **웹업 비교 / ISAS 자동화 규정 / 설문지 작성 가이드라인**
6. **RD부 전체 업무 프로세스 / 업무별 소요 일정 기준**
7. **답변 방식** — 질문 의도별 답변 범위 표 + 판단 결과별 템플릿
8. **주의사항** — 금지/완화 표현 규칙들
9. **설문지 오류 체크** (로직 오류 / 내용 기반 검토)
10. **예상 응답시간·적립금 계산** — 문항 유형별 기준표와 계산 규칙

## 자주 하는 작업: ISAS 케이스 추가

사용자가 ISAS **불가능/가능 케이스를 캡처(png)와 함께** 하나씩 알려준다. 작업 흐름:

1. **바로 prompt.txt를 수정하지 말 것.** 사용자가 "한번에 반영해줘"라고 할 때까지 케이스를 기억만 한다.
2. 캡처는 `Read` 툴로 이미지를 직접 분석해 구조를 파악한다.
3. 각 케이스를 한 줄로 요약해 사용자에게 확인받는다 (불가 포인트가 여러 개면 모두 분리).
4. "반영해줘" 요청 시, `prompt.txt`의 **알맞은 섹션**에 추가한다:

| 케이스 성격 | 넣을 섹션 |
|------------|----------|
| 테이블형 관련 불가 | `## ISAS 미지원 케이스` > `### 테이블형 문항 관련` |
| 카테고리 관련 불가 | `## ISAS 미지원 케이스` > `### 카테고리 관련` |
| 로테이션 관련 불가 | `## ISAS 미지원 케이스` > `### 로테이션 관련` |
| 그 외 불가 | `## ISAS 미지원 케이스` > `### 기타 미지원 항목` |
| ISAS는 안되지만 RD부 수작업 가능 | `## RD부 수작업 가능 케이스` 표 |
| ISAS로 가능한 케이스 | `## ISAS 작업 가능 케이스` 해당 하위 항목 |

5. 기존 표/항목과 **중복되지 않는지** 먼저 확인하고, 중복이면 기존 행을 보강한다.

### prompt.txt의 핵심 판단 원칙 (요약)
- **한 표 = 단일 문항 유형 / 단일 응답 유형** 이 대원칙. 한 표에 여러 문항·여러 척도·텍스트+수치 혼합은 대부분 불가.
- 테이블형 보기(헤더)는 **반드시 1행**.
- 카테고리는 **1단계까지만** (2중 카테고리 불가). 단 단수/복수/순위형은 1단계 카테고리 가능.
- 드롭박스 문항 유형은 ISAS에 **없음**.
- 일부 불가 문항만 있으면 → RD부 ISAS + 온실 수작업 **혼합 진행** 또는 온실 전체 이관 안내.
- ISAS 불가라도 **RD부 내부 수작업**으로 가능한 건 온실 이관 불필요로 안내.

## 코드 수정 시 주의
- `prompt.txt`는 `encoding="utf-8"`로 읽힌다. 한글 깨짐 주의.
- prompt.txt 수정 전 `백업` 폴더에 날짜 붙여 백업본을 만든다 (예: `prompt_백업_20260610.txt`).
- 표(마크다운 테이블) 형식을 유지할 것 — 프롬프트 내 다른 표와 스타일을 맞춘다.
- 큰 변경 전에는 사용자에게 어느 섹션을 어떻게 바꿀지 먼저 요약해 확인받는다.
- UI 검증은 Playwright MCP로 스크린샷을 찍어 직접 확인한다 (라이트/다크 모두).
