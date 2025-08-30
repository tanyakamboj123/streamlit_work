# streamlit_work
# Your Text, Your Style — Auto-Generate a Presentation

Turn bulk text or markdown into a fully-styled PowerPoint deck that matches any uploaded template (.pptx/.potx). Bring your own LLM key (OpenAI, Anthropic, or Gemini). No AI image generation — the app reuses images already in your template.

## Demo
Hosted: https://<your-deployment-url>  <!-- replace with Streamlit Cloud / Render link -->

## Features
- Paste long text/markdown + optional one-line guidance (e.g., “investor pitch”)
- Upload a PowerPoint template/presentation to inherit layouts, fonts, colors
- Reuse template images (logos/motifs) automatically
- Choose provider: OpenAI / Anthropic / Gemini (key never stored)
- Download the generated `.pptx`

## Quick Start (local)
```bash
python -m venv venv
# Windows: venv\Scripts\activate   |  macOS/Linux: source venv/bin/activate
pip install -r requirements.txt  # or see below
streamlit run app.py
