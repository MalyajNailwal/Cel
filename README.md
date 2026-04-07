# Cel

**Your AI-Powered Excel Command Center**

Cel transforms how you work with spreadsheets. Just tell it what you need in plain English — it understands, plans, and executes.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                                 USER                                         │
│                    "Create pie chart of sales by region"                   │
└──────────────────────────────────┬──────────────────────────────────────────┘
                                   │
                                   ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                            FRONTEND                                          │
│  ┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌──────────────┐  │
│  │  ChatInput │───▶│ ReasoningUI │───▶│  Executor   │───▶│  Office.js   │  │
│  └─────────────┘    └─────────────┘    └─────────────┘    └──────────────┘  │
│         │                                         │                          │
│         │            ┌─────────────┐              │                          │
│         └───────────▶│  TypingInd  │              │                          │
│                      └─────────────┘              │                          │
└──────────────────────────────────────────────────┼──────────────────────────┘
                                                   │
                    ┌──────────────────────────────┼──────────────────────┐
                    │                              ▼                       │
                    │                   ┌──────────────────┐               │
                    │                   │  Backend (FastAPI)│               │
                    │                   └────────┬─────────┘               │
                    │                            │                        │
                    ▼                            ▼                        │
┌─────────────────────────────────────────────────────────────────────────────┐
│                              BACKEND                                         │
│                                                                              │
│  ┌────────────────┐    ┌────────────────┐    ┌────────────────────────┐  │
│  │  Reasoning     │───▶│   Planning      │───▶│      Executor          │  │
│  │    Agent       │    │    Agent        │    │   (Frontend)           │  │
│  │  (CrewAI)      │    │  (CrewAI)       │    │                        │  │
│  └───────┬────────┘    └────────┬────────┘    └────────────────────────┘  │
│          │                      │                                                │
│          │          ┌───────────┴───────────┐                                 │
│          │          │                       │                                 │
│          ▼          ▼                       ▼                                 │
│  ┌────────────────────┐    ┌────────────────────┐    ┌─────────────────┐  │
│  │   Analysis         │    │   Chart             │    │    Data         │  │
│  │   Agent            │    │   Generator         │    │    Generator    │  │
│  │ (stats, trends)    │    │  (Office.js)       │    │ (sample data)   │  │
│  └────────────────────┘    └────────────────────┘    └─────────────────┘  │
│                                                                              │
│  ┌───────────────────────────────────────────────────────────────────────┐  │
│  │                        AI Providers                                   │  │
│  │         OpenAI • Anthropic • Google • OpenRouter (custom)            │  │
│  └───────────────────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## How It Works

1. **You speak** — Type your request in natural language
2. **Reasoning** — Agent breaks down what you need (toggle on/off)
3. **Planning** — Generates step-by-step execution plan
4. **Execution** — Runs operations via Office.js API
5. **Result** — Returns outcome with validation

---

## What It Does

| Capability | Example |
|------------|---------|
| **Charts** | "Pie chart of sales by region" |
| **Data** | "Add 500 employee records" |
| **Formatting** | "Bold header, freeze top row" |
| **Analysis** | "Find trends in this data" |
| **Formulas** | "Average of column C" |
| **Sheets** | "Create new sheet named Sales" |

---

## Tech Stack

| Layer | Technology |
|-------|------------|
| Frontend | React • Vite • Office.js • Tailwind |
| Backend | FastAPI • CrewAI • Python |
| AI | OpenAI • Anthropic • Google • OpenRouter |

---

## Quick Start

### Frontend
```bash
npm run dev
```
Opens: `https://localhost:3000/src/taskpane.html`

### Backend
```bash
cd backend
pip install -r requirements.txt
python3 -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

---

## Common Commands

| Task | Command |
|------|---------|
| Kill frontend | `lsof -ti:3000 \| xargs kill -9` |
| Kill backend | `lsof -ti:8000 \| xargs kill -9` |
| Clear cache | `rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/WebKit/` |
| Build | `npm run build` |

---

## Excel Sideload (macOS)

```bash
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```

Open Excel → Insert → My Add-ins → Browse → select `manifest.xml`

---

## Key Features

- Natural language → Excel operations
- Reasoning agent (toggle on/off)
- Statistical analysis (100K+ rows)
- Smart chart generation
- Multi-sheet management
- Real-time execution tracking

## Agentic Design Patterns

Based on *Agentic Design Patterns* by Antonio Gulli, Cel implements:

| Pattern | How Cel Uses It |
|---------|-----------------|
| **Planning** | Planning Agent breaks tasks into executable steps |
| **Reasoning** | Shows thinking before execution (toggle on/off) |
| **Tool Use** | Office.js API calls (set_values, create_table, charts) |
| **Multi-Agent** | Reasoning → Planning → Validator → Reflection |
| **Reflection** | Auto-recovers from failures (e.g., table overlap) |
| **Guardrails** | Confirms destructive actions, warns on large writes |
| **Goal Monitoring** | Tracks step progress: "3/5 steps completed" |
| **Resource-Aware** | Simple tasks → gpt-4o-mini, Complex → full model |
| **Exception Handling** | Retry logic, sheet verification, overlap detection |
| **Routing** | Analysis → /api/analyze, Charts → Office.js, Data → /api/generate-data |
