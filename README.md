# Cel

**Autonomous Excel intelligence that plans before it acts.**

Cel is a multi-agent Excel add-in that interprets natural language, generates executable plans, and validates outcomes — with guardrails that prevent common spreadsheet errors before they happen.

---

## How It Works

```
User: "Create a bar chart of revenue by region"

    ┌──────────────┐
    │  Reasoning   │  Decomposes intent, identifies columns, checks memory
    │    Agent     │  "Revenue=numeric, Region=categorical, use bar chart"
    └──────┬───────┘
           ▼
    ┌──────────────┐
    │   Planner    │  Generates JSON plan with confidence scores
    │    Agent     │  [{action: "create_chart", confidence: 0.85}]
    └──────┬───────┘
           ▼
    ┌──────────────┐
    │   Executor   │  Validates, checkpoints, executes via Office.js
    │              │  Pre-checks: sheet exists, formula syntax, data types
    └──────┬───────┘
           ▼
    ┌──────────────┐
    │  Validator   │  Verifies outcomes, triggers reflection on failure
    └──────────────┘
```

Every request goes through the same pipeline. Charts, formulas, formatting, data generation — all planned, validated, and reversible.

---

## Key Features

| Feature | What It Does |
|---------|-------------|
| **Multi-Agent Pipeline** | Reasoning → Planning → Execution → Validation — each step bounded |
| **Confidence Scoring** | Each plan step scored 0-100%. Low confidence prompts confirmation |
| **Workbook Awareness** | Planner sees full structure: all sheets, headers, tables, named ranges |
| **Cross-Sheet Intelligence** | User says "copy from DataSheet" → planner fetches that sheet's headers |
| **Semantic Resolution** | "Sales" resolves to "Revenue" via synonym mapping — no exact match needed |
| **Pre-Write Validation** | Checks sheet exists, formulas balanced, data types compatible before write |
| **One-Click Undo** | Checkpoint snapshots before every mutation — restore with one click |
| **Persistent Memory** | User rules, formatting conventions, past mistakes carried across sessions |
| **Self-Recovery** | On failure, reflector re-plans using column headers and error context |
| **Any Model** | OpenAI, Anthropic, Google, OpenRouter — any model ID, no hardcoding |

---

## Why Cel

| Capability | Cel | Microsoft Copilot | Pi for Excel |
|-----------|-----|-------------------|--------------|
| Multi-agent planning | Yes | No | No |
| Confidence-based execution | Yes | No | No |
| Cross-sheet awareness | Yes | Limited | No |
| Semantic column resolution | Yes | No | No |
| Pre-write validation | Yes | No | No |
| Checkpoint undo | Yes | No | No |
| Recovery via reflection | Yes | Limited | Basic |
| Dynamic model selection | Any model ID | Fixed | Fixed |
| Backend data analysis | 100k+ rows | No | No |

---

## Quick Start

### Frontend

```bash
npm install
npm run dev
```

Opens at `https://localhost:3000/src/taskpane.html`

### Backend

```bash
cd backend
pip install -r requirements.txt
python3 -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

### Sideload into Excel (macOS)

```bash
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```

Open Excel → Insert → My Add-ins → Browse → select `manifest.xml`

---

## Tech Stack

| Layer | Technology |
|-------|------------|
| Frontend | React, Vite, Office.js, Tailwind |
| Backend | FastAPI, CrewAI, Python |
| AI | OpenAI, Anthropic, Google, OpenRouter |

---

## Project Structure

```
cel/
├── backend/
│   ├── main.py              # FastAPI server, agent pipeline, all endpoints
│   └── requirements.txt
├── src/
│   ├── App.tsx              # Main task pane — execution loop, guardrails, UI
│   ├── lib/
│   │   ├── excel-api.ts     # Office.js wrapper (all Excel operations)
│   │   └── types.ts         # TypeScript interfaces
│   └── components/
│       ├── ChatInput.tsx    # Message input with range preview
│       ├── SettingsPanel.tsx # Model config, memory, conventions
│       └── ConfirmModal.tsx  # Modal confirmation (replaces window.confirm)
├── manifest.xml             # Office.js add-in manifest
└── package.json
```

---

## API Endpoints

| Endpoint | Purpose |
|----------|---------|
| `POST /api/chat` | Main pipeline — reasoning → planning → plan JSON |
| `POST /api/execute` | Execute a plan via Office.js |
| `POST /api/reflect` | Self-correct failed steps and re-plan |
| `POST /api/analyze` | Backend data analysis for large datasets |
| `POST /api/analyze-large` | Chunked analysis for 100k+ row datasets |
| `GET /api/health` | Server status check |

---

## Built With

- [Office.js](https://learn.microsoft.com/en-us/javascript/api/office-apps) — Excel manipulation API
- [CrewAI](https://docs.crewai.com/) — Multi-agent orchestration
- [FastAPI](https://fastapi.tiangolo.com/) — Backend API framework
- [Vite](https://vitejs.dev/) — Frontend build tool
