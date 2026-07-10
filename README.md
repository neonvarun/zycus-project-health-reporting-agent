# Project Health Reporting Agent

An automated PMO reporting tool that reads Excel project plans, calculates Red/Amber/Green health, explains the result in plain English, and creates weekly reports plus a monthly executive presentation.

## What it does

- Reads one or more `.xlsx` project plans.
- Calculates RAG status using deterministic rules for:
  - Schedule slippage
  - Milestone health
  - Blockers and risks
  - SPI-based budget proxy
  - Stakeholder sentiment
  - Data quality
- Uses local Qwen3.5-0.8B Q4_K_M for explanation generation.
- Downloads the model automatically on first use into `.models/`.
- Falls back to rules-based explanations if the local model is unavailable.
- Generates Markdown and JSON weekly reports.
- Generates a 7-slide monthly PowerPoint synthesis.
- Provides a Streamlit dashboard for uploads, analysis, previews, and downloads.

## Setup

Python 3.12 is recommended for the local model runtime.

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

On Windows, if `llama-cpp-python` tries to compile, install the prebuilt CPU wheel:

```powershell
pip install huggingface-hub
pip install --only-binary=:all: llama-cpp-python --extra-index-url https://abetlen.github.io/llama-cpp-python/whl/cpu
```

## Run the CLI

Run the complete pipeline on the sample projects:

```powershell
$env:LLM_PROVIDER="local"
python main.py
```

The first run downloads approximately 553 MB of model weights. Later runs reuse the cached model.

To run without model inference:

```powershell
$env:LLM_PROVIDER="rules"
python main.py
```

## Run the dashboard

```powershell
streamlit run app.py
```

The dashboard supports Excel uploads, project health cards, risk analysis, report previews, PowerPoint generation, and methodology viewing.

## Outputs

```text
outputs/
├── weekly/
│   ├── Project Plan B_weekly_report.md
│   ├── Project Plan B_weekly_report.json
│   ├── S2P Project_weekly_report.md
│   └── S2P Project_weekly_report.json
└── monthly/
    └── project_health_monthly_synthesis.pptx
```

## Design decisions

- RAG scoring is deterministic and auditable; the language model never decides the RAG color.
- The local model only converts calculated metrics into clear executive reasoning.
- Missing sheets, formula errors, unparseable dates, and incomplete fields are recorded as data-quality assumptions.
- Portfolio presentation content is generated from the analyzed project metrics.

## Tests

```powershell
python -m unittest -v tests\test_health_rules.py tests\test_ppt_generator.py tests\test_explainer.py
```

## Submission notes

Include the source code, `data/`, `docs/`, `outputs/`, and this README. Do not include `.env`, `.venv/`, `__pycache__/`, or `.models/`.
