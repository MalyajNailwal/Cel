# Cel

**Intelligent Spreadsheet Orchestration**

Cel is an autonomous Excel agent that interprets intent, plans execution, and delivers outcomes — with enterprise-grade guardrails that exceed native capabilities.

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
│         └───────────▶│  Streaming  │              │                          │
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
│  ┌��───────────────┐    ┌────────────────┐    ┌────────────────────────┐  │
│  │  Reasoning     │───▶│   Planning      │───▶│      Validation          │  │
│  │    Agent       │    │    Agent        │    │       Agent              │  │
│  │  (CrewAI)      │    │  (CrewAI)       │    │    (CrewAI)             │  │
│  └───────┬────────┘    └────────┬────────┘    └────────┬───────────┘  │
│          │                      │                      │              │
│          │          ┌───────────┴───────────┐          │              │
│          │          │   Connected Flow   │◀─────────┘              │
│          │          │ (Reasoning→Planner) │                           │
│          ▼          ▼                   ▼                           │
│  ┌─────────────────────────────────────────────────────────────────────┐  │
│  │                    Multi-Layer Guardrails                           │  │
│  │  • Formula Validation    • Error Detection    • Timeout           │  │
│  │  • Pre-execution Review  • Checkpointing     • Recovery           │  │
│  └─────────────────────────────────────────────────────────────────────┘  │
│                                                                              │
│  ┌───────────────────────────────────────────────────────────────────────┐  │
│  │                        AI Providers                                   │  │
│  │         OpenAI • Anthropic • Google • OpenRouter (any model)            │  │
│  └───────────────────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## How It Works

1. **Intent Capture** — Express requirements in natural language
2. **Cognitive Layer** — System decomposes and validates approach (toggleable)
3. **Agent Pipeline** — Reasoning→Planner→Validator (all connected)
4. **Materialization** — Operations applied via native spreadsheet APIs
5. **Outcome Delivery** — Results presented with contextual validation
6. **Recovery** — Self-corrects when execution diverges from plan

---

## Why Cel Exceeds Native Capabilities

| Feature | Cel | Microsoft Agent |
|---------|-----|-----------------|
| **Agent Pipeline** | Reasoning→Planner→Validator (connected) | Single agent |
| **Formula Validation** | ✅ Checks parentheses, arguments | ❌ |
| **Error Detection** | ✅ Warns #N/A, #DIV/0! risk | ❌ Produces errors |
| **Pre-execution Review** | ✅ Shows plan, user confirms | ❌ |
| **Operation Timeout** | ✅ 2-min max, prevents infinite loops | ❌ |
| **Memory System** | ✅ Checkpoints, rules, conventions | ❌ |
| **Recovery** | ✅ Automatic self-correction | Limited |
| **Streaming** | ✅ Real-time response chunks | ❌ |
| **Dynamic Model** | ✅ Any model ID (no hardcoding) | Fixed |

---

## Capabilities

| Domain | Example |
|--------|---------|
| **Visualization** | "Pie chart of sales by region" |
| **Data Generation** | "Add 500 employee records" |
| **Presentation** | "Bold header, freeze top row" |
| **Insight Extraction** | "Find trends in this data" |
| **Computation** | "Average of column C" |
| **Structure** | "Create new sheet named Sales" |

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

## Design Philosophy

Cel implements a multi-agent orchestration pattern where each layer operates with bounded autonomy:

- **Interpretive Layer** — Decomposes ambiguous requests into structured intent
- **Synthesis Layer** — Constructs executable plans with dependency awareness
- **Validation Layer** — Verifies outcomes against expected state
- **Recovery Layer** — Self-corrects when execution diverges from plan

The system maintains contextual awareness across sessions, adapting to user preferences and learning from operational history. Destructive operations are gated by confirmation protocols, and state transitions are checkpointed for reversible workflows.

Provider abstraction enables seamless model switching without architectural changes — the system routes to the appropriate inference endpoint based on configuration, supporting both direct API access and federated gateway patterns.
