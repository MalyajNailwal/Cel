# Cel

**Intelligent Spreadsheet Orchestration**

Cel redefines human-computer interaction with spreadsheets. Rather than navigating menus and formulas, you express intent — the system interprets, plans, and materializes outcomes directly within your workbook.

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

1. **Intent Capture** — Express requirements in natural language
2. **Cognitive Layer** — System decomposes and validates approach (toggleable)
3. **Plan Synthesis** — Generates executable operation sequence
4. **Materialization** — Operations applied via native spreadsheet APIs
5. **Outcome Delivery** — Results presented with contextual validation

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
