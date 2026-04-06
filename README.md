# Cel - Excel AI Agent

An intelligent Excel add-in powered by AI agents that lets you chat with your spreadsheets.

## Quick Start

### Frontend (Office.js + React + Vite)
```bash
npm run dev
```
- Runs on `https://localhost:3000`
- Access: `https://localhost:3000/src/taskpane.html`

### Backend (FastAPI + CrewAI)
```bash
cd backend
pip install -r requirements.txt
python3 -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```
- Runs on `http://localhost:8000`

## Development Commands

### Kill port if stuck
```bash
lsof -ti:3000 | xargs kill -9  # frontend
lsof -ti:8000 | xargs kill -9  # backend
```

### Clear Excel cache (after manifest changes)
```bash
rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/WebKit/
rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/WebKit/
```

### Excel Add-in Sideload (macOS)
```bash
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```
Then open Excel → Insert → My Add-ins → Browse

### Build for production
```bash
npm run build
```

## Tech Stack
- **Frontend**: React, Vite, Office.js, Tailwind
- **Backend**: FastAPI, CrewAI, Python
- **AI**: OpenAI, Anthropic, Google, OpenRouter

## Features
- Natural language Excel commands
- Data analysis agent
- Multi-provider AI support
- Chart generation
- Formatting & formulas
