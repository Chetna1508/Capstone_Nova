from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

APP_DIR = Path(__file__).resolve().parent.parent
PROJECT_ROOT = APP_DIR.parent

# Support a repository-level .env as well as a package-local .env.
# Existing process environment variables always take precedence.
load_dotenv(PROJECT_ROOT / ".env")
load_dotenv(APP_DIR / ".env", override=False)

# Google ADK reads GOOGLE_API_KEY directly from the environment when
# using the Gemini Developer API. Never hard-code an API key.
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY", "").strip()
GOOGLE_GENAI_USE_VERTEXAI = os.getenv(
    "GOOGLE_GENAI_USE_VERTEXAI",
    "false",
).strip().lower()

if GOOGLE_API_KEY:
    os.environ["GOOGLE_API_KEY"] = GOOGLE_API_KEY

# Use the Gemini Developer API unless Vertex AI is explicitly selected.
os.environ.setdefault("GOOGLE_GENAI_USE_VERTEXAI", "false")

ASSETS_DIR = Path(
    os.getenv(
        "NOVARETAIL_ASSETS_DIR",
        str(PROJECT_ROOT / "assets"),
    )
).expanduser()

DATASET_PATH = ASSETS_DIR / "novaretail_dataset.xlsx"
TEMPLATE_PATH = ASSETS_DIR / "novaretail_template.pptx"
ICON_DIR = ASSETS_DIR / "icons"

OUTPUT_DIR = APP_DIR / "output"
CHART_DIR = OUTPUT_DIR / "charts"
OUTPUT_PPTX = OUTPUT_DIR / "NovaRetail_QBR.pptx"

MODEL_NAME = os.getenv(
    "NOVARETAIL_MODEL",
    "gemini-2.5-flash",
).strip()

ALLOW_AGENT_FALLBACK = os.getenv(
    "NOVARETAIL_ALLOW_FALLBACK",
    "0",
).strip().lower() in {
    "1",
    "true",
    "yes",
}

ICON_MAPPING = {
    1: ICON_DIR / "growth_chart.png",
    2: ICON_DIR / "globe_region.png",
    3: ICON_DIR / "customer_star.png",
    4: ICON_DIR / "risk_alert.png",
    5: ICON_DIR / "lightbulb_action.png",
}