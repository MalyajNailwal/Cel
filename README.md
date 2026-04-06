# Cel

**Your AI-Powered Excel Command Center**

Cel transforms Excel into an intelligent assistant that understands plain English and executes complex spreadsheet tasks autonomously.

---

## Capabilities

**Natural Language Commands**
- "Create a pie chart of sales by region"
- "Add a new sheet with 500 employee records"
- "Bold the header row and freeze the top row"
- "Calculate average for column C"

**AI-Powered Analysis**
- Statistical analysis on any dataset size
- Trend detection & outlier identification
- Smart chart recommendations

**Multi-Provider AI**
- OpenAI, Anthropic, Google, OpenRouter
- Custom model support

---

## Quick Start

### Frontend
```bash
npm run dev
```
- Opens: `https://localhost:3000/src/taskpane.html`

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
- Data analysis & statistics
- Smart chart generation
- Multi-sheet management
- Cell formatting & formulas
- Reasoning agent (toggle on/off)
- Real-time execution tracking
