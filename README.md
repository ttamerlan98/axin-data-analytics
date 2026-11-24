# DataLens Linux — Avtomatik Excel Analizator

DataLens Linux is a local-first Azerbaijani-language FastAPI application for
reviewing Excel workbooks, profiling data quality, exploring financial
analytics, reconciling sheets, auditing anomalies, and exporting a styled
multi-sheet report. Uploaded workbook contents remain in application memory
and are discarded when the server stops.

## Features

- Upload `.xlsx` and `.xls` workbooks; every workbook sheet is ingested.
  `.csv` remains supported for tabular data, but has no Excel formatting.
- Profile nulls, exact duplicate rows, trimmed text, text lengths, column types,
  and a transparent 0–100 data-quality score.
- Extract `.xlsx` fills, font colors, bold text, comments, and merged headers.
- Explore DuckDB-backed numeric/category summaries and top-ten, dense-rank,
  difference, running-total, and LAG analytics.
- Reconcile two sheets by key and amount, including exact matches, amount
  differences, one-sided rows, totals, and net differences.
- Detect IQR/negative numeric values and near-duplicate text; generate an
  optional, entirely local Qwen audit.
- Export a styled workbook containing the audit summary, AI audit, color and
  grouped analyses, reconciliation, anomalies, and each sheet's cleaned data.
- Preview safe deterministic cleaning, then explicitly apply or undo it.

## Start the application

Python 3.10+ and pip are required. `llama-cpp-python` may need a native C++
build toolchain if a compatible wheel is unavailable for your Python version.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --host 127.0.0.1 --port 8000
```

Open <http://127.0.0.1:8000>. No external AI API or API key is used.

## Local model

The Qwen model is **not** downloaded automatically. The audit screen's
download button (or `POST /api/ai/download`) fetches
`Qwen2.5-1.5B-Instruct-Q4_K_M.gguf` from
`bartowski/Qwen2.5-1.5B-Instruct-GGUF` into the ignored `models/` directory
(approximately 1.1 GB). `GET /api/ai/status` reports download/load state.
After download, load the model before requesting an AI audit. Inference uses
`n_ctx=4096` and `n_threads=4`; the system prompt restricts it to Excel and
financial-data auditing.

## API highlights

- `POST /api/upload` — ingest one or more files.
- `GET /api/overview` — dataset and sheet overview.
- `POST /api/profile` — data health and quality profile.
- `POST /api/analytics` — statistical and window analytics.
- `POST /api/colors/analyze` — workbook fill-color summaries.
- `POST /api/reconcile` — cross-sheet reconciliation.
- `POST /api/audit/summary`, `POST /api/audit/ai` — two-stage audit.
- `POST /api/export/full` — complete styled `.xlsx` report.
- `GET /api/health`, `GET /api/ai/status` — local service/model status.

## Tests

```bash
python -m unittest discover -s tests -v
```

The tests cover existing font-color extraction plus upload, multi-sheet
ingestion, merged headers, fill/comment extraction, profiling, reconciliation,
and styled report export.
