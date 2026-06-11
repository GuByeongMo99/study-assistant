"""
시험공부 보조 프로그램 - Study Assistant
FastAPI + OpenAI Whisper API + Claude API + python-pptx

실행: uvicorn main:app --reload --port 8000
"""

import os
import json
import base64
import asyncio
import tempfile
from pathlib import Path
from typing import Optional

import anthropic
import httpx
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import fitz  # PyMuPDF

app = FastAPI(title="Study Assistant")

# ── 디렉터리 ────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# ── Anthropic 클라이언트 ────────────────────────────────
def get_anthropic_client():
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY 환경변수가 설정되지 않았습니다.")
    return anthropic.Anthropic(api_key=api_key)

def get_openai_key():
    key = os.getenv("OPENAI_API_KEY")
    if not key:
        raise HTTPException(status_code=500, detail="OPENAI_API_KEY 환경변수가 설정되지 않았습니다.")
    return key


# ════════════════════════════════════════════════════════
# 1) STT: 녹음 파일 → 텍스트 (OpenAI Whisper API)
# ════════════════════════════════════════════════════════
@app.post("/api/transcribe")
async def transcribe_audio(audio: UploadFile = File(...)):
    """녹음 파일을 업로드하면 텍스트로 변환합니다."""
    allowed = {".mp3", ".mp4", ".wav", ".m4a", ".ogg", ".webm", ".flac"}
    suffix = Path(audio.filename).suffix.lower()
    if suffix not in allowed:
        raise HTTPException(status_code=400, detail=f"지원 형식: {', '.join(allowed)}")

    # 임시 파일 저장
    tmp_path = UPLOAD_DIR / f"audio_{audio.filename}"
    content = await audio.read()
    tmp_path.write_bytes(content)

    openai_key = get_openai_key()
    async with httpx.AsyncClient(timeout=120) as client:
        with open(tmp_path, "rb") as f:
            resp = await client.post(
                "https://api.openai.com/v1/audio/transcriptions",
                headers={"Authorization": f"Bearer {openai_key}"},
                data={"model": "whisper-1", "language": "ko"},
                files={"file": (audio.filename, f, "audio/mpeg")},
            )
    tmp_path.unlink(missing_ok=True)

    if resp.status_code != 200:
        raise HTTPException(status_code=500, detail=f"Whisper 오류: {resp.text}")

    transcript = resp.json().get("text", "")
    return {"transcript": transcript}


# ════════════════════════════════════════════════════════
# 2) 강조 포인트 추출 (Claude)
# ════════════════════════════════════════════════════════
@app.post("/api/extract-emphasis")
async def extract_emphasis(body: dict):
    """
    transcript: STT 결과 텍스트
    slide_texts: [슬라이드별 텍스트 리스트] (선택)
    """
    transcript = body.get("transcript", "")
    slide_texts = body.get("slide_texts", [])

    if not transcript:
        raise HTTPException(status_code=400, detail="transcript가 필요합니다.")

    context = ""
    if slide_texts:
        context = "\n\n[강의 슬라이드 내용]\n" + "\n---\n".join(
            f"슬라이드 {i+1}: {t}" for i, t in enumerate(slide_texts)
        )

    prompt = f"""당신은 대학 강의를 분석하는 학습 도우미입니다.
아래는 교수님의 강의 녹음을 텍스트로 변환한 내용입니다.{context}

[강의 녹음 텍스트]
{transcript}

다음을 분석하여 JSON으로 반환하세요:
{{
  "key_concepts": [
    {{"term": "개념명", "explanation": "간결한 설명", "importance": "high/medium/low"}}
  ],
  "emphasized_points": [
    {{"point": "강조 내용", "context": "관련 문맥", "likely_exam": true/false}}
  ],
  "summary": "전체 강의 3줄 요약",
  "flashcards": [
    {{"question": "예상 문제", "answer": "핵심 답변"}}
  ],
  "slide_notes": {{
    "슬라이드 번호(숫자)": "해당 슬라이드에 대한 교수 설명 보충"
  }}
}}

JSON만 반환하고 다른 텍스트는 포함하지 마세요."""

    client = get_anthropic_client()
    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = message.content[0].text.strip()
    # JSON 코드블록 제거
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        data = {"raw": raw, "error": "JSON 파싱 실패"}

    return data


# ════════════════════════════════════════════════════════
# 3) PPT/PDF 슬라이드 텍스트 추출
# ════════════════════════════════════════════════════════
@app.post("/api/extract-slides")
async def extract_slides(file: UploadFile = File(...)):
    """PPT 또는 PDF에서 슬라이드별 텍스트와 이미지(base64)를 추출합니다."""
    suffix = Path(file.filename).suffix.lower()
    content = await file.read()
    save_path = UPLOAD_DIR / file.filename
    save_path.write_bytes(content)

    slides = []

    if suffix in (".pptx", ".ppt"):
        prs = Presentation(str(save_path))
        for i, slide in enumerate(prs.slides):
            texts = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        line = para.text.strip()
                        if line:
                            texts.append(line)
            slides.append({
                "index": i,
                "text": "\n".join(texts),
                "title": texts[0] if texts else f"슬라이드 {i+1}"
            })

    elif suffix == ".pdf":
        doc = fitz.open(str(save_path))
        for i, page in enumerate(doc):
            text = page.get_text("text").strip()
            lines = [l.strip() for l in text.split("\n") if l.strip()]
            slides.append({
                "index": i,
                "text": "\n".join(lines),
                "title": lines[0] if lines else f"페이지 {i+1}"
            })
        doc.close()
    else:
        raise HTTPException(status_code=400, detail="PPT, PPTX, PDF만 지원합니다.")

    return {"filename": file.filename, "slides": slides, "total": len(slides)}


# ════════════════════════════════════════════════════════
# 4) 이해하기 쉬운 PPT 생성 (Claude + python-pptx)
# ════════════════════════════════════════════════════════
@app.post("/api/generate-study-pptx")
async def generate_study_pptx(body: dict):
    """
    slides: [{index, title, text}]
    analysis: extract-emphasis 결과
    original_filename: 원본 파일명
    """
    slides_data = body.get("slides", [])
    analysis = body.get("analysis", {})
    original_filename = body.get("original_filename", "lecture")

    if not slides_data:
        raise HTTPException(status_code=400, detail="slides 데이터가 필요합니다.")

    # Claude로 각 슬라이드를 쉬운 버전으로 변환
    client = get_anthropic_client()
    slide_notes = analysis.get("slide_notes", {})

    enhanced_slides = []
    for slide in slides_data:
        idx = slide["index"]
        extra_note = slide_notes.get(str(idx), "") or slide_notes.get(idx, "")

        prompt = f"""다음 강의 슬라이드 내용을 대학생이 시험공부할 때 이해하기 쉽도록 재구성하세요.

[원본 슬라이드 제목]: {slide['title']}
[원본 내용]:
{slide['text']}
{f"[교수 강조 보충]: {extra_note}" if extra_note else ""}

JSON으로만 반환:
{{
  "title": "슬라이드 제목 (유지 또는 더 명확하게)",
  "bullet_points": ["핵심 포인트 1 (예시나 비유 포함)", "핵심 포인트 2", ...],
  "key_term": "이 슬라이드의 가장 중요한 용어",
  "one_line_summary": "한 줄 요약",
  "exam_tip": "시험 포인트 (있으면)"
}}"""

        msg = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}]
        )
        raw = msg.content[0].text.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        raw = raw.strip()
        try:
            enhanced = json.loads(raw)
        except:
            enhanced = {
                "title": slide["title"],
                "bullet_points": [slide["text"]],
                "key_term": "",
                "one_line_summary": "",
                "exam_tip": ""
            }
        enhanced["original"] = slide
        enhanced_slides.append(enhanced)

    # PPT 생성
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    # 색상 팔레트
    COLOR_BG      = RGBColor(0x0F, 0x17, 0x2A)   # 딥 네이비
    COLOR_CARD    = RGBColor(0x1A, 0x26, 0x42)   # 카드 배경
    COLOR_ACCENT  = RGBColor(0x4A, 0x9E, 0xFF)   # 파란 하이라이트
    COLOR_GREEN   = RGBColor(0x3D, 0xD6, 0x8C)   # 시험 팁
    COLOR_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
    COLOR_GRAY    = RGBColor(0xA0, 0xAE, 0xC4)
    COLOR_YELLOW  = RGBColor(0xFF, 0xD7, 0x60)

    blank_layout = prs.slide_layouts[6]  # 완전 빈 레이아웃

    def add_rect(slide, left, top, width, height, fill_color, radius=False):
        shape = slide.shapes.add_shape(
            1,  # MSO_SHAPE_TYPE.RECTANGLE
            Inches(left), Inches(top), Inches(width), Inches(height)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
        shape.line.fill.background()
        return shape

    def add_text_box(slide, text, left, top, width, height,
                     font_size=14, color=None, bold=False, align=PP_ALIGN.LEFT, wrap=True):
        txBox = slide.shapes.add_textbox(
            Inches(left), Inches(top), Inches(width), Inches(height)
        )
        tf = txBox.text_frame
        tf.word_wrap = wrap
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = Pt(font_size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = color
        return txBox

    for i, es in enumerate(enhanced_slides):
        slide = prs.slides.add_slide(blank_layout)

        # ── 배경 ──
        add_rect(slide, 0, 0, 13.33, 7.5, COLOR_BG)

        # ── 슬라이드 번호 배지 ──
        add_rect(slide, 0.3, 0.2, 0.7, 0.4, COLOR_ACCENT)
        add_text_box(slide, f"{i+1:02d}", 0.3, 0.2, 0.7, 0.4,
                     font_size=13, color=COLOR_WHITE, bold=True, align=PP_ALIGN.CENTER)

        # ── 제목 ──
        add_text_box(slide, es.get("title", ""), 1.1, 0.18, 9.0, 0.55,
                     font_size=22, color=COLOR_WHITE, bold=True)

        # ── 핵심 용어 배지 ──
        key_term = es.get("key_term", "")
        if key_term:
            add_rect(slide, 10.3, 0.18, 2.7, 0.45, COLOR_ACCENT)
            add_text_box(slide, f"핵심어: {key_term}", 10.3, 0.18, 2.7, 0.45,
                         font_size=11, color=COLOR_WHITE, bold=True, align=PP_ALIGN.CENTER)

        # ── 구분선 ──
        line = slide.shapes.add_shape(1, Inches(0.3), Inches(0.82), Inches(12.73), Inches(0.03))
        line.fill.solid()
        line.fill.fore_color.rgb = COLOR_ACCENT
        line.line.fill.background()

        # ── 원본 슬라이드 패널 (왼쪽) ──
        add_rect(slide, 0.3, 1.0, 5.8, 5.6, COLOR_CARD)
        add_text_box(slide, "📄 원본 슬라이드", 0.4, 1.05, 2.5, 0.35,
                     font_size=10, color=COLOR_GRAY, bold=True)
        original_text = es["original"].get("text", "")
        # 원본 텍스트 (최대 600자)
        display_orig = original_text[:600] + ("..." if len(original_text) > 600 else "")
        add_text_box(slide, display_orig, 0.4, 1.45, 5.5, 4.8,
                     font_size=11, color=COLOR_WHITE)

        # ── 이해하기 쉬운 버전 패널 (오른쪽) ──
        add_rect(slide, 6.4, 1.0, 6.6, 5.6, COLOR_CARD)
        add_text_box(slide, "✨ 이해하기 쉬운 버전", 6.5, 1.05, 3.5, 0.35,
                     font_size=10, color=COLOR_ACCENT, bold=True)

        # 한 줄 요약
        summary = es.get("one_line_summary", "")
        if summary:
            add_rect(slide, 6.5, 1.45, 6.3, 0.45, RGBColor(0x4A, 0x9E, 0xFF))
            add_text_box(slide, f"💡 {summary}", 6.55, 1.47, 6.2, 0.42,
                         font_size=10, color=COLOR_WHITE, bold=True)

        # 불릿 포인트
        bullet_top = 2.05 if summary else 1.55
        bullets = es.get("bullet_points", [])
        for j, bp in enumerate(bullets[:5]):
            y_pos = bullet_top + j * 0.62
            if y_pos > 5.7:
                break
            add_rect(slide, 6.5, y_pos, 0.06, 0.35, COLOR_ACCENT)
            add_text_box(slide, bp, 6.65, y_pos - 0.05, 6.1, 0.65,
                         font_size=11, color=COLOR_WHITE)

        # 시험 팁
        exam_tip = es.get("exam_tip", "")
        if exam_tip:
            add_rect(slide, 6.4, 6.15, 6.6, 0.45, RGBColor(0x1A, 0x3A, 0x2A))
            add_text_box(slide, f"🎯 시험 포인트: {exam_tip}", 6.5, 6.17, 6.4, 0.42,
                         font_size=10, color=COLOR_GREEN, bold=True)

    # 마지막 슬라이드: 플래시카드 요약
    flashcards = analysis.get("flashcards", [])
    if flashcards:
        slide = prs.slides.add_slide(blank_layout)
        add_rect(slide, 0, 0, 13.33, 7.5, COLOR_BG)
        add_text_box(slide, "🃏 예상 문제 & 핵심 답변", 0.5, 0.2, 12.0, 0.6,
                     font_size=24, color=COLOR_YELLOW, bold=True, align=PP_ALIGN.CENTER)
        line = slide.shapes.add_shape(1, Inches(0.3), Inches(0.9), Inches(12.73), Inches(0.03))
        line.fill.solid()
        line.fill.fore_color.rgb = COLOR_YELLOW
        line.line.fill.background()

        cols = 2
        for j, fc in enumerate(flashcards[:6]):
            col = j % cols
            row = j // cols
            x = 0.4 + col * 6.5
            y = 1.1 + row * 2.0
            add_rect(slide, x, y, 6.2, 1.8, COLOR_CARD)
            add_text_box(slide, f"Q. {fc.get('question','')}", x+0.15, y+0.1, 5.9, 0.55,
                         font_size=11, color=COLOR_YELLOW, bold=True)
            add_rect(slide, x+0.1, y+0.68, 6.0, 0.03, COLOR_GRAY)
            add_text_box(slide, f"A. {fc.get('answer','')}", x+0.15, y+0.8, 5.9, 0.85,
                         font_size=11, color=COLOR_WHITE)

    # 저장
    stem = Path(original_filename).stem
    out_name = f"{stem}_study.pptx"
    out_path = OUTPUT_DIR / out_name
    prs.save(str(out_path))

    return {"filename": out_name, "slide_count": len(enhanced_slides)}


# ════════════════════════════════════════════════════════
# 5) 생성된 파일 다운로드
# ════════════════════════════════════════════════════════
@app.get("/api/download/{filename}")
async def download_file(filename: str):
    path = OUTPUT_DIR / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다.")
    return FileResponse(
        str(path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename
    )


# ════════════════════════════════════════════════════════
# 메인 UI
# ════════════════════════════════════════════════════════
@app.get("/", response_class=HTMLResponse)
async def index():
    return HTML_PAGE


HTML_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>StudyAI — 시험공부 보조</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@300;400;500;600;700&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg:       #0F172A;
    --surface:  #1A2642;
    --surface2: #243356;
    --accent:   #4A9EFF;
    --accent2:  #3DD68C;
    --warn:     #FFD760;
    --text:     #E8EEF8;
    --muted:    #A0AEC4;
    --border:   rgba(74,158,255,.2);
    --radius:   14px;
  }

  body {
    font-family: 'Pretendard', 'Noto Sans KR', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
  }

  /* ── 헤더 ── */
  header {
    display: flex; align-items: center; gap: 14px;
    padding: 20px 36px;
    background: rgba(26,38,66,.85);
    backdrop-filter: blur(12px);
    border-bottom: 1px solid var(--border);
    position: sticky; top: 0; z-index: 100;
  }
  .logo { font-size: 22px; font-weight: 700; }
  .logo span { color: var(--accent); }
  .tag { font-size: 11px; background: var(--accent); color: #fff;
         padding: 3px 10px; border-radius: 99px; font-weight: 600; }

  /* ── 레이아웃 ── */
  .container { max-width: 1100px; margin: 0 auto; padding: 36px 24px; }

  /* ── 스텝 카드 ── */
  .step-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
  @media(max-width:700px){ .step-grid { grid-template-columns: 1fr; } }

  .step-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 28px;
    position: relative;
    transition: border-color .2s;
  }
  .step-card:hover { border-color: var(--accent); }
  .step-card.full { grid-column: 1 / -1; }

  .step-num {
    position: absolute; top: -14px; left: 24px;
    background: var(--accent); color: #fff;
    font-size: 11px; font-weight: 700; padding: 3px 12px; border-radius: 99px;
  }
  .step-title { font-size: 17px; font-weight: 700; margin-bottom: 6px; }
  .step-desc  { font-size: 13px; color: var(--muted); margin-bottom: 20px; line-height: 1.6; }

  /* ── 업로드 존 ── */
  .upload-zone {
    border: 2px dashed var(--border);
    border-radius: 10px;
    padding: 28px;
    text-align: center;
    cursor: pointer;
    transition: border-color .2s, background .2s;
    background: rgba(74,158,255,.04);
  }
  .upload-zone:hover, .upload-zone.drag { border-color: var(--accent); background: rgba(74,158,255,.1); }
  .upload-zone input[type=file] { display: none; }
  .upload-icon { font-size: 32px; margin-bottom: 8px; }
  .upload-label { font-size: 14px; color: var(--muted); }
  .upload-label strong { color: var(--accent); }
  .file-name { margin-top: 10px; font-size: 13px; color: var(--accent2); font-weight: 600; }

  /* ── 버튼 ── */
  .btn {
    display: inline-flex; align-items: center; gap: 8px;
    padding: 11px 22px; border-radius: 8px;
    font-size: 14px; font-weight: 600; border: none; cursor: pointer;
    transition: opacity .15s, transform .1s;
  }
  .btn:hover { opacity: .88; }
  .btn:active { transform: scale(.97); }
  .btn-primary { background: var(--accent); color: #fff; }
  .btn-success { background: var(--accent2); color: #0F172A; }
  .btn-warn    { background: var(--warn);    color: #0F172A; }
  .btn:disabled { opacity: .4; cursor: not-allowed; }

  /* ── 진행 로그 ── */
  .log-box {
    background: #0A1020; border-radius: 10px;
    padding: 16px 20px; font-size: 13px; font-family: 'Courier New', monospace;
    color: var(--accent2); line-height: 1.8; max-height: 200px; overflow-y: auto;
    border: 1px solid var(--border);
  }
  .log-box .log-err { color: #FF7070; }
  .log-box .log-info { color: var(--muted); }

  /* ── 결과 패널 ── */
  .result-panel { margin-top: 28px; }
  .result-panel h3 { font-size: 16px; font-weight: 700; margin-bottom: 14px; color: var(--warn); }

  .key-concepts { display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 20px; }
  .concept-badge {
    padding: 6px 14px; border-radius: 99px; font-size: 13px; font-weight: 600;
    border: 1px solid;
  }
  .importance-high   { background: rgba(74,158,255,.15); border-color: var(--accent);  color: var(--accent); }
  .importance-medium { background: rgba(61,214,140,.12); border-color: var(--accent2); color: var(--accent2); }
  .importance-low    { background: rgba(160,174,196,.1); border-color: var(--muted);   color: var(--muted); }

  .flashcard-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
  @media(max-width:600px){ .flashcard-grid { grid-template-columns: 1fr; } }

  .flashcard {
    background: var(--surface2); border-radius: 10px; padding: 16px;
    border: 1px solid var(--border); cursor: pointer;
    transition: border-color .2s;
  }
  .flashcard:hover { border-color: var(--accent); }
  .flashcard .q { font-size: 13px; color: var(--warn); font-weight: 600; margin-bottom: 8px; }
  .flashcard .a { font-size: 13px; color: var(--text); display: none; line-height: 1.6; }
  .flashcard.flipped .a { display: block; }
  .flashcard.flipped .flip-hint { display: none; }
  .flip-hint { font-size: 11px; color: var(--muted); margin-top: 6px; }

  .summary-box {
    background: var(--surface2); border-radius: 10px; padding: 18px;
    border-left: 3px solid var(--accent2); margin-bottom: 20px;
    font-size: 14px; line-height: 1.8; color: var(--text);
  }

  .download-banner {
    background: linear-gradient(135deg, var(--surface2), #1A3A2A);
    border: 1px solid var(--accent2); border-radius: 12px;
    padding: 22px 28px; display: flex; align-items: center; justify-content: space-between;
    gap: 16px; flex-wrap: wrap;
  }
  .download-banner .info { font-size: 15px; }
  .download-banner .info strong { color: var(--accent2); }

  /* ── 스피너 ── */
  .spinner {
    display: inline-block; width: 16px; height: 16px;
    border: 2px solid rgba(255,255,255,.3);
    border-top-color: #fff; border-radius: 50%;
    animation: spin .7s linear infinite;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  textarea {
    width: 100%; background: #0A1020; border: 1px solid var(--border);
    border-radius: 8px; color: var(--text); padding: 14px; font-size: 13px;
    font-family: inherit; resize: vertical; outline: none; line-height: 1.7;
  }
  textarea:focus { border-color: var(--accent); }

  .slide-preview {
    display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 12px;
    max-height: 300px; overflow-y: auto;
  }
  .slide-thumb {
    background: var(--surface2); border-radius: 8px; padding: 12px;
    border: 1px solid var(--border); font-size: 12px;
  }
  .slide-thumb .num { color: var(--accent); font-weight: 700; margin-bottom: 4px; }
  .slide-thumb .text { color: var(--muted); overflow: hidden;
    display: -webkit-box; -webkit-line-clamp: 3; -webkit-box-orient: vertical; }
</style>
</head>
<body>

<header>
  <div class="logo">Study<span>AI</span></div>
  <div class="tag">BETA</div>
</header>

<div class="container">

  <!-- ── STEP 그리드 ── -->
  <div class="step-grid">

    <!-- STEP 1: 녹음 업로드 -->
    <div class="step-card">
      <div class="step-num">STEP 1 — 녹음 파일</div>
      <div class="step-title">강의 녹음 업로드</div>
      <div class="step-desc">교수님 강의 녹음을 업로드하면 자동으로 텍스트로 변환하고 강조 포인트를 추출합니다.</div>
      <label class="upload-zone" id="audioZone">
        <input type="file" id="audioFile" accept=".mp3,.wav,.m4a,.mp4,.ogg,.webm,.flac">
        <div class="upload-icon">🎙️</div>
        <div class="upload-label"><strong>클릭</strong>하거나 파일을 드래그하세요</div>
        <div class="upload-label">MP3, WAV, M4A, MP4 지원</div>
        <div class="file-name" id="audioFileName"></div>
      </label>
      <br>
      <button class="btn btn-primary" id="btnTranscribe" disabled onclick="runTranscribe()">
        🎙️ 텍스트로 변환
      </button>
    </div>

    <!-- STEP 2: 강의자료 업로드 -->
    <div class="step-card">
      <div class="step-num">STEP 2 — 강의자료</div>
      <div class="step-title">PPT / PDF 업로드</div>
      <div class="step-desc">강의 슬라이드를 업로드하세요. 원본과 이해하기 쉬운 버전을 나란히 비교할 수 있는 학습 PPT를 생성합니다.</div>
      <label class="upload-zone" id="slideZone">
        <input type="file" id="slideFile" accept=".pptx,.ppt,.pdf">
        <div class="upload-icon">📊</div>
        <div class="upload-label"><strong>클릭</strong>하거나 파일을 드래그하세요</div>
        <div class="upload-label">PPTX, PPT, PDF 지원</div>
        <div class="file-name" id="slideFileName"></div>
      </label>
      <br>
      <button class="btn btn-primary" id="btnExtractSlides" disabled onclick="runExtractSlides()">
        📊 슬라이드 추출
      </button>
    </div>

    <!-- STEP 3: 텍스트 결과 & 분석 -->
    <div class="step-card full" id="transcriptCard" style="display:none">
      <div class="step-num">STEP 3 — 변환 결과</div>
      <div class="step-title">변환된 텍스트 확인 & 분석</div>
      <div class="step-desc">변환된 텍스트를 확인하고 필요하면 수정한 뒤 강조 포인트를 추출하세요.</div>
      <textarea id="transcriptText" rows="6" placeholder="변환된 텍스트가 여기 표시됩니다..."></textarea>
      <br><br>
      <button class="btn btn-warn" id="btnAnalyze" onclick="runAnalysis()">
        🧠 강조 포인트 추출 & 예상 문제 생성
      </button>
    </div>

    <!-- STEP 4: 슬라이드 미리보기 -->
    <div class="step-card full" id="slidePreviewCard" style="display:none">
      <div class="step-num">STEP 4 — 슬라이드 확인</div>
      <div class="step-title">추출된 슬라이드 미리보기</div>
      <div id="slidePreviewGrid" class="slide-preview"></div>
    </div>

  </div><!-- end step-grid -->

  <!-- ── 진행 로그 ── -->
  <div style="margin-top:24px">
    <div class="log-box" id="logBox">대기 중... 파일을 업로드하고 시작하세요.</div>
  </div>

  <!-- ── 분석 결과 패널 ── -->
  <div class="result-panel" id="resultPanel" style="display:none">

    <h3>📌 강의 요약</h3>
    <div class="summary-box" id="summaryBox"></div>

    <h3>🔑 핵심 개념</h3>
    <div class="key-concepts" id="conceptsBox"></div>

    <h3>🃏 플래시카드 (클릭하면 답 확인)</h3>
    <div class="flashcard-grid" id="flashcardGrid"></div>

    <br>
    <!-- PPT 생성 버튼 -->
    <div id="generatePptSection">
      <button class="btn btn-success" id="btnGenPpt" onclick="runGeneratePPT()">
        ✨ 학습용 PPT 생성하기
      </button>
      <span style="font-size:13px;color:var(--muted);margin-left:12px">원본 + 이해하기 쉬운 버전 나란히 비교</span>
    </div>
  </div>

  <!-- ── 다운로드 배너 ── -->
  <div class="download-banner" id="downloadBanner" style="display:none; margin-top:28px">
    <div class="info">
      <div style="font-size:13px;color:var(--muted);margin-bottom:4px">생성 완료</div>
      <strong id="downloadFilename">study.pptx</strong> 파일이 준비됐습니다.
    </div>
    <button class="btn btn-success" id="btnDownload" onclick="downloadFile()">
      ⬇️ PPT 다운로드
    </button>
  </div>

</div><!-- end container -->

<script>
// ── 상태 ──────────────────────────────────────────────
let state = {
  transcript: "",
  slides: [],
  analysis: null,
  originalFilename: "",
  outputFilename: ""
};

// ── 유틸 ──────────────────────────────────────────────
function log(msg, type="") {
  const box = document.getElementById("logBox");
  const cls = type === "err" ? "log-err" : type === "info" ? "log-info" : "";
  box.innerHTML += `<div class="${cls}">${msg}</div>`;
  box.scrollTop = box.scrollHeight;
}

function setBtn(id, loading, text) {
  const btn = document.getElementById(id);
  btn.disabled = loading;
  btn.innerHTML = loading ? `<span class="spinner"></span> ${text}` : text;
}

async function apiCall(url, options={}) {
  const res = await fetch(url, options);
  if (!res.ok) {
    const err = await res.json().catch(() => ({detail: res.statusText}));
    throw new Error(err.detail || res.statusText);
  }
  return res.json();
}

// ── 파일 선택 ─────────────────────────────────────────
document.getElementById("audioFile").addEventListener("change", function() {
  if (this.files[0]) {
    document.getElementById("audioFileName").textContent = this.files[0].name;
    document.getElementById("btnTranscribe").disabled = false;
  }
});

document.getElementById("slideFile").addEventListener("change", function() {
  if (this.files[0]) {
    document.getElementById("slideFileName").textContent = this.files[0].name;
    document.getElementById("btnExtractSlides").disabled = false;
    state.originalFilename = this.files[0].name;
  }
});

// 드래그앤드롭
["audioZone","slideZone"].forEach(id => {
  const el = document.getElementById(id);
  el.addEventListener("dragover", e => { e.preventDefault(); el.classList.add("drag"); });
  el.addEventListener("dragleave", () => el.classList.remove("drag"));
  el.addEventListener("drop", e => {
    e.preventDefault(); el.classList.remove("drag");
    const inp = el.querySelector("input[type=file]");
    inp.files = e.dataTransfer.files;
    inp.dispatchEvent(new Event("change"));
  });
});

// ── STEP 1: 녹음 변환 ─────────────────────────────────
async function runTranscribe() {
  const file = document.getElementById("audioFile").files[0];
  if (!file) return;
  setBtn("btnTranscribe", true, "변환 중...");
  log(`🎙️ ${file.name} 변환 시작 (Whisper API)...`);
  try {
    const fd = new FormData();
    fd.append("audio", file);
    const data = await apiCall("/api/transcribe", { method:"POST", body: fd });
    state.transcript = data.transcript;
    document.getElementById("transcriptText").value = data.transcript;
    document.getElementById("transcriptCard").style.display = "block";
    log(`✅ 변환 완료 — ${data.transcript.length}자`);
  } catch(e) {
    log(`❌ 변환 실패: ${e.message}`, "err");
  }
  setBtn("btnTranscribe", false, "🎙️ 텍스트로 변환");
}

// ── STEP 2: 슬라이드 추출 ─────────────────────────────
async function runExtractSlides() {
  const file = document.getElementById("slideFile").files[0];
  if (!file) return;
  setBtn("btnExtractSlides", true, "추출 중...");
  log(`📊 ${file.name} 슬라이드 추출 중...`);
  try {
    const fd = new FormData();
    fd.append("file", file);
    const data = await apiCall("/api/extract-slides", { method:"POST", body: fd });
    state.slides = data.slides;
    state.originalFilename = data.filename;
    renderSlidePreview(data.slides);
    document.getElementById("slidePreviewCard").style.display = "block";
    log(`✅ ${data.total}개 슬라이드 추출 완료`);
  } catch(e) {
    log(`❌ 슬라이드 추출 실패: ${e.message}`, "err");
  }
  setBtn("btnExtractSlides", false, "📊 슬라이드 추출");
}

function renderSlidePreview(slides) {
  const grid = document.getElementById("slidePreviewGrid");
  grid.innerHTML = slides.map(s => `
    <div class="slide-thumb">
      <div class="num">슬라이드 ${s.index + 1}</div>
      <div class="text">${s.text || "(내용 없음)"}</div>
    </div>
  `).join("");
}

// ── STEP 3: 강조 포인트 분석 ──────────────────────────
async function runAnalysis() {
  const transcript = document.getElementById("transcriptText").value;
  if (!transcript && state.slides.length === 0) {
    log("⚠️ 녹음 텍스트 또는 슬라이드가 필요합니다.", "err");
    return;
  }
  state.transcript = transcript;
  setBtn("btnAnalyze", true, "분석 중...");
  log("🧠 Claude가 강조 포인트를 분석하고 있습니다...");
  try {
    const body = {
      transcript: state.transcript,
      slide_texts: state.slides.map(s => s.text)
    };
    const data = await apiCall("/api/extract-emphasis", {
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body: JSON.stringify(body)
    });
    state.analysis = data;
    renderAnalysis(data);
    document.getElementById("resultPanel").style.display = "block";
    log(`✅ 분석 완료 — 핵심 개념 ${(data.key_concepts||[]).length}개, 예상문제 ${(data.flashcards||[]).length}개`);
  } catch(e) {
    log(`❌ 분석 실패: ${e.message}`, "err");
  }
  setBtn("btnAnalyze", false, "🧠 강조 포인트 추출 & 예상 문제 생성");
}

function renderAnalysis(data) {
  // 요약
  document.getElementById("summaryBox").textContent = data.summary || "";

  // 핵심 개념
  const concepts = document.getElementById("conceptsBox");
  concepts.innerHTML = (data.key_concepts || []).map(c => `
    <div class="concept-badge importance-${c.importance}" title="${c.explanation}">
      ${c.term}
    </div>
  `).join("");

  // 플래시카드
  const grid = document.getElementById("flashcardGrid");
  grid.innerHTML = (data.flashcards || []).map(fc => `
    <div class="flashcard" onclick="this.classList.toggle('flipped')">
      <div class="q">Q. ${fc.question}</div>
      <div class="a">A. ${fc.answer}</div>
      <div class="flip-hint">👆 클릭하면 답 확인</div>
    </div>
  `).join("");
}

// ── STEP 4: PPT 생성 ──────────────────────────────────
async function runGeneratePPT() {
  if (state.slides.length === 0) {
    log("⚠️ 슬라이드를 먼저 업로드하고 추출하세요.", "err");
    return;
  }
  setBtn("btnGenPpt", true, "PPT 생성 중... (1-3분 소요)");
  log("✨ 이해하기 쉬운 학습 PPT를 생성하고 있습니다...");
  log("   슬라이드 수에 따라 1~3분 걸릴 수 있습니다.", "info");
  try {
    const body = {
      slides: state.slides,
      analysis: state.analysis || {},
      original_filename: state.originalFilename
    };
    const data = await apiCall("/api/generate-study-pptx", {
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body: JSON.stringify(body)
    });
    state.outputFilename = data.filename;
    document.getElementById("downloadFilename").textContent = data.filename;
    document.getElementById("downloadBanner").style.display = "flex";
    log(`🎉 PPT 생성 완료! ${data.slide_count}개 슬라이드 + 플래시카드 슬라이드`);
  } catch(e) {
    log(`❌ PPT 생성 실패: ${e.message}`, "err");
  }
  setBtn("btnGenPpt", false, "✨ 학습용 PPT 생성하기");
}

// ── 다운로드 ──────────────────────────────────────────
function downloadFile() {
  if (state.outputFilename) {
    window.location.href = `/api/download/${state.outputFilename}`;
  }
}
</script>
</body>
</html>
"""

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
