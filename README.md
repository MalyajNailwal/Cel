# Cel - Excel AI Agent

An intelligent Excel add-in powered by AI agents that lets you chat with your spreadsheets.

## Quick Start

### Frontend (Office.js + React + Vite)
```bash
npm run dev
```
- Runs on `https://localhost:3000`
- Access: `https://localhost:3000/src/taskpane.html` (for Excel sidebar)

### Backend (FastAPI + CrewAI)
```bash
cd backend
pip install -r requirements.txt
python3 -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```
- Runs on `http://localhost:8000`

## Common Commands

### Kill stuck ports
```bash
lsof -ti:3000 | xargs kill -9  # frontend
lsof -ti:8000 | xargs kill -9  # backend
```

### Clear Excel cache (after manifest changes)
```bash
rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/WebKit/
rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/WebKit/
```

### Build for production
```bash
npm run build
```

### Reinstall dependencies
```bash
npm install
pip install -r backend/requirements.txt
```

## Excel Add-in Sideload (macOS)

1. Create wef folder:
```bash
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```

2. Copy manifest:
```bash
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```

3. Open Excel → Insert → My Add-ins → Browse → select manifest.xml

4. If add-in doesn't appear, clear cache and restart Excel (Cmd+Q)

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Add-in not showing | Clear Excel cache + restart Excel |
| Port 3000 stuck | `lsof -ti:3000 \| xargs kill -9` |
| Port 8000 stuck | `lsof -ti:8000 \| xargs kill -9` |
| HTTPS error | Re-run `mkcert localhost` to regenerate certs |
| Build fails | Delete `dist/` folder, then `npm run build` |

## Tech Stack
- **Frontend**: React, Vite, Office.js, Tailwind
- **Backend**: FastAPI, CrewAI, Python
- **AI**: OpenAI, Anthropic, Google, OpenRouter

## Features
- Natural language Excel commands
- Data analysis agent (stats, trends, outliers)
- Multi-provider AI support
- Chart generation
- Formatting & formulas