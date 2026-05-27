---
title: SlideGen AI
emoji: 📊
colorFrom: blue
colorTo: purple
sdk: docker
app_port: 7860
pinned: false
---

# SlideGen AI

Type any topic — get a polished PowerPoint presentation in seconds.

## How it works

1. Enter a subject in the browser UI
2. Groq (Llama 3.3 70B) designs the slide structure and picks a color theme
3. `python-pptx` renders a styled `.pptx` file
4. Your browser downloads it automatically

## Stack

- **Backend:** FastAPI + python-pptx
- **LLM:** Groq API (Llama 3.3 70B)
- **Frontend:** Vanilla HTML/JS

## Running locally

```bash
cp .env.example .env
# Add your GROQ_API_KEY to .env

pip install -r requirements.txt
python -m uvicorn app:app --host 0.0.0.0 --port 8080
```

Open http://localhost:8080
