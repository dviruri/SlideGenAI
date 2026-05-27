import json
import os
import uuid
import tempfile
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

from openai import OpenAI
from fastapi import BackgroundTasks, FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from pptx_builder import generate_presentation

app = FastAPI(title="Slide Generator")

GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
GROQ_ENDPOINT = "https://api.groq.com/openai/v1"
MODEL = "llama-3.3-70b-versatile"

SYSTEM_PROMPT = """You are a professional presentation designer. Given a topic, return ONLY a valid JSON object (no markdown, no code fences) with this structure:

{
  "presentation": {
    "title": "string",
    "subtitle": "string"
  },
  "theme": {
    "primaryColor": "6-char hex, no #",
    "secondaryColor": "6-char hex, no #",
    "accentColor": "6-char hex, no # — vibrant, contrasts with dark bg",
    "darkBg": "6-char hex, no # — dark color for title/conclusion slides",
    "lightBg": "6-char hex, no # — very light, near-white for content slides"
  },
  "slides": [
    {
      "type": "title|content|section_break|stats|two_column|conclusion",
      "title": "string",
      "subtitle": "string (title type only, optional)",
      "bullets": ["string", ...],
      "stats": [{"value": "string", "label": "string"}, ...],
      "columns": [{"title": "string", "bullets": ["string", ...]}, ...]
    }
  ]
}

Rules:
- First slide must be type "title", last must be type "conclusion"
- 6–9 slides total
- "content": up to 5 bullets, ≤10 words each
- "stats": 3–4 impactful numbers with short labels; omit bullets/columns
- "two_column": exactly 2 columns, each with a title and 3–5 bullets; omit stats/bullets
- "section_break": title only, no bullets/stats/columns — use to mark major topic transitions
- Choose colors that fit the subject (tech→blue, nature→green, finance→navy, health→teal)
- primaryColor and darkBg should be dark; lightBg very light (e.g. F5F7FA or FFFFFF)
- Return ONLY the JSON — no text before or after"""


class GenerateRequest(BaseModel):
    subject: str


def _cleanup(path: Path):
    path.unlink(missing_ok=True)


def _get_client():
    if not GROQ_API_KEY:
        raise HTTPException(500, "GROQ_API_KEY is not set. Add it to your .env file.")
    return OpenAI(base_url=GROQ_ENDPOINT, api_key=GROQ_API_KEY)


@app.post("/api/generate")
async def generate(request: GenerateRequest, background_tasks: BackgroundTasks):
    subject = request.subject.strip()
    if not subject:
        raise HTTPException(400, "Subject is required")
    if len(subject) > 500:
        raise HTTPException(400, "Subject is too long (max 500 characters)")

    client = _get_client()

    try:
        response = client.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": f"Create a presentation about: {subject}"},
            ],
            max_tokens=4096,
            temperature=0.7,
        )
    except Exception as e:
        raise HTTPException(502, f"Groq API error: {e}")

    raw = response.choices[0].message.content.strip()

    # Strip markdown fences if the model wrapped the JSON
    if raw.startswith("```"):
        lines = raw.splitlines()
        raw = "\n".join(lines[1:] if lines[0].startswith("```") else lines)
        if raw.endswith("```"):
            raw = raw[: raw.rfind("```")]
    raw = raw.strip()

    try:
        slide_data = json.loads(raw)
    except json.JSONDecodeError as e:
        raise HTTPException(500, f"Model returned invalid JSON: {e}")

    file_id = str(uuid.uuid4())
    output_path = Path(tempfile.gettempdir()) / f"slidegen_{file_id}.pptx"

    try:
        generate_presentation(slide_data, str(output_path))
    except Exception as e:
        raise HTTPException(500, f"PPTX generation failed: {e}")

    title = slide_data.get("presentation", {}).get("title", "presentation")
    safe_title = "".join(c if c.isalnum() or c in " -_" else "" for c in title).strip()[:60]
    filename = f"{safe_title or 'presentation'}.pptx"

    background_tasks.add_task(_cleanup, output_path)

    return FileResponse(
        str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename,
    )


app.mount("/", StaticFiles(directory="static", html=True), name="static")
