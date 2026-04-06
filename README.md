# Cel

**Your AI-Powered Excel Command Center**

Cel transforms Excel into an intelligent assistant that understands plain English and executes complex spreadsheet tasks autonomously.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                                 USER                                         │
│                    "Create pie chart of sales"                              │
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
│  │ (stats, trends)    │    │  (Office.js)       │    │ (sample data)  │  │
│  └────────────────────┘    └────────────────────┘    └─────────────────┘  │
│                                                                              │
│  ┌───────────────────────────────────────────────────────────────────────┐  │
│  │                        AI Providers                                   │  │
│  │         OpenAI • Anthropic • Google • OpenRouter (custom)            │  │
│  └───────────────────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Execution Flow

```
1. USER sends message
        │
        ▼
2. REASONING AGENT analyzes request (if Think enabled)
        │
        ▼
3. PLANNING AGENT generates execution plan
        │
        ▼
4. BACKEND routes to appropriate handler:
   • Analysis → stats/trends on data
   • Charts   → smart chart generation
   • Data     → sample data generation
   • Default  → Excel operations
        │
        ▼
5. FRONTEND EXECUTOR runs Office.js calls
        │
        ▼
6. RESULT returned to user + validation
```

---

## Capabilities

| Feature | Description |
|---------|-------------|
| **Natural Language** | "Create pie chart", "Add 500 rows", "Bold header" |
| **Reasoning Agent** | Shows thinking before execution (toggle on/off) |
| **Analysis** | Stats, trends, outliers on any data size (100K+ rows) |
| **Smart Charts** | Auto-detects column types, recommends best chart |
| **Multi-Provider** | OpenAI, Anthropic, Google, OpenRouter |
| **Dynamic** | No hardcoded headers - analyzes actual data values |

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

## Commands

| Task | Command |
|------|---------|
| Kill frontend | `lsof -ti:3000 \| xargs kill -9` |
| Kill backend | `lsof -ti:8000 \| xargs kill -9` |
| Clear Excel cache | `rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/WebKit/` |
| Build | `npm run build` |

---

## Excel Sideload (macOS)

```bash
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```

Open Excel → Insert → My Add-ins → Browse → select `manifest.xml`

---

## Tech Stack

| Layer | Technology |
|-------|------------|
| Frontend | React • Vite • Office.js • Tailwind |
| Backend | FastAPI • CrewAI • Python |
| AI | OpenAI • Anthropic • Google • OpenRouter |

---

## Features

- Natural language → Excel operations
- Reasoning agent with toggle
- Data analysis (any size)
- Smart chart generation
- Multi-sheet management
- Cell formatting & formulas
- Real-time execution tracking
