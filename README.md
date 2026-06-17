# CQU Study Guide Automation MVP

DOCX-to-DOCX automation pipeline for turning uploaded study guide documents into a CQU-branded Word document using the provided master template.

## What is implemented

- FastAPI backend with:
  - `POST /generate`
  - `GET /jobs/{job_id}`
  - `GET /download/{job_id}`
- Static upload UI at `/`
- Hyperlink analysis that checks HTTP/HTTPS links and flags broken or manually reviewed URLs
- Local DOCX pipeline:
  - extracts headings, paragraphs, lists, tables, and embedded images
  - reinserts original image binaries into the final DOCX
  - renders content into the CQU template
  - validates output opens, placeholders are gone, headings/tables are present when expected, and images are preserved
- CLI for Stage 1:
  - `python main.py input.docx template.docx output.docx`
- Optional Groq AI polish pass for the next stage.

## Run locally

```bash
python3 -m pip install -r requirements.txt
uvicorn app.main:app --reload
```

Open `http://127.0.0.1:8000`.

## CLI

```bash
python main.py source.docx app/templates/CQU_study_guide_template.docx output.docx \
  --unit-code LAWS13019 \
  --year 2026 \
  --term T1 \
  --week 1 \
  --version 1.0
```

## AI configuration

The MVP runs without AI by default. For Google Gemini testing, set:

```bash
AI_PROVIDER=google
GOOGLE_API_KEY=your_key_here
GOOGLE_MODEL=gemini-2.5-flash
ENABLE_AI_POLISH=true
```

Keep real keys in your shell environment or a local `.env` file. `.env` is ignored by git.

The AI pass is conservative: it can polish text, but the backend rejects responses that change image IDs, image order, or table shapes.
