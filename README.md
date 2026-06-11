# StudyAI — 시험공부 보조 프로그램

교수님 강의 녹음 + 강의자료를 AI로 분석해서 시험공부를 도와주는 로컬 웹앱입니다.

---

## 주요 기능

| 기능 | 설명 |
|------|------|
| 🎙️ **STT** | 녹음 파일 → 텍스트 변환 (OpenAI Whisper API) |
| 🧠 **강조 분석** | Claude가 강조 포인트, 핵심 개념, 예상 문제 추출 |
| 📊 **슬라이드 분석** | PPT/PDF 내용 추출 + 강의 내용과 매핑 |
| ✨ **학습 PPT 생성** | 원본 슬라이드 ↔ 이해하기 쉬운 버전 나란히 비교 |
| 🃏 **플래시카드** | 예상 문제 카드 (클릭하면 답 확인) + PPT 마지막 슬라이드에 삽입 |

---

## 빠른 시작

### 1. 환경 설정

```bash
# 프로젝트 폴더로 이동
cd study-assistant

# 패키지 설치
pip install -r requirements.txt

# API 키 설정 (아래 중 하나 선택)

# Windows
set ANTHROPIC_API_KEY=sk-ant-...
set OPENAI_API_KEY=sk-...

# Mac/Linux
export ANTHROPIC_API_KEY=sk-ant-...
export OPENAI_API_KEY=sk-...
```

### 2. 서버 실행

```bash
python main.py
# 또는
uvicorn main:app --reload --port 8000
```

### 3. 브라우저에서 접속

```
http://localhost:8000
```

---

## 사용 흐름

```
1. 🎙️  강의 녹음 파일 업로드  →  Whisper API로 텍스트 변환
2. 📊  강의 PPT/PDF 업로드   →  슬라이드별 텍스트 추출
3. 🧠  [강조 포인트 추출] 클릭 →  Claude가 핵심 개념 + 예상 문제 생성
4. ✨  [학습용 PPT 생성] 클릭  →  원본 vs 이해하기 쉬운 버전 비교 PPT 다운로드
```

---

## 필요한 API 키

| 키 | 용도 | 발급처 |
|----|------|--------|
| `ANTHROPIC_API_KEY` | 강조 분석, PPT 내용 생성 | https://console.anthropic.com |
| `OPENAI_API_KEY` | 녹음 → 텍스트 (Whisper) | https://platform.openai.com |

> 💡 **녹음 파일 없이도** PPT/PDF만 업로드해서 학습 자료 생성 가능합니다.
> 텍스트박스에 직접 필기 내용을 입력해도 됩니다.

---

## 생성되는 PPT 구조

```
슬라이드 1~N:
┌─────────────────────┬──────────────────────┐
│   📄 원본 슬라이드   │  ✨ 이해하기 쉬운 버전 │
│                     │  💡 한 줄 요약         │
│  (원본 텍스트 그대로) │  • 핵심 포인트 1       │
│                     │  • 핵심 포인트 2       │
│                     │  🎯 시험 포인트        │
└─────────────────────┴──────────────────────┘

마지막 슬라이드:
🃏 예상 문제 & 핵심 답변 (플래시카드 형식)
```

---

## 지원 파일 형식

- **녹음**: MP3, WAV, M4A, MP4, OGG, WEBM, FLAC
- **강의자료**: PPTX, PPT, PDF

---

## 폴더 구조

```
study-assistant/
├── main.py            ← 메인 서버 (FastAPI)
├── requirements.txt
├── README.md
├── uploads/           ← 업로드된 파일 (자동 생성)
└── outputs/           ← 생성된 PPT (자동 생성)
```
