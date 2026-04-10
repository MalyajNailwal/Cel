<!-- This section is maintained by the coding agent via lore (https://github.com/BYK/opencode-lore) -->
## Long-term Knowledge

### Architecture

<!-- lore:019d7740-6cd2-7dc8-9279-99fe9cd99683 -->
* **Agent pipeline lacks clarification and adaptive execution**: Current Reasoning→Planner→Validator pipeline has two architectural gaps: (1) No clarification agent — ambiguous requests like 'clean the data' cause planner to guess blindly; a pre-planning clarification step for low-confidence intents would prevent wrong actions. (2) No adaptive executor — all plan steps execute sequentially without real-time adaptation; if step 3 fails, steps 4-7 still run and cascade. Reflection agent runs post-hoc, not during execution. These gaps directly impact outcome quality: wrong assumptions from ambiguity + cascading failures from blind execution.

<!-- lore:019d6743-3189-7a6b-8909-7340f5976d0b -->
* **Cel being positioned as SaaS product**: User views Cel through a SaaS lens, not just a personal tool. This affects architectural decisions: needs user auth, multi-tenant support, cloud storage, and billing eventually. Current localStorage approach is sufficient for now, but migration path to Convex or Supabase should be planned. Project already has Supabase available — prefer it over adding Convex unless Convex offers clear advantages for the specific use case.

<!-- lore:019d6bb6-491c-7cc9-a44a-bd4796e53ce4 -->
* **Model defaults removed - fully dynamic model selection**: Model is now fully dynamic - no hardcoded defaults. All Pydantic request classes in backend/main.py (ChatRequest, ExecuteRequest, AnalyzeRequest, ReflectRequest) use \`model: str = ""\` instead of \`model: str = "gpt-4o"\`. Same change in frontend SettingsPanel.tsx and App.tsx. This forces clients to explicitly provide model ID - any model works (gpt-4o, claude-sonnet, gemini-pro, custom OpenRouter models). Fallback logic was removed since frontend always sends model from settings.

<!-- lore:019d6775-cb40-76de-a459-106c58501ab8 -->
* **Pi for Excel as reference architecture**: Pi for Excel (tmustier/pi-for-excel, 290 stars) serves as the primary reference architecture for Cel's feature roadmap. Key patterns to adopt: (1) Recovery checkpoints — snapshot cell values before mutations for one-click undo; (2) Persistent rules — user/workbook-level guidance stored and followed by AI; (3) Formatting conventions — currency, date, number preferences applied consistently; (4) Session management — multiple conversations per workbook with auto-save/restore; (5) Context compaction — summarize old messages when context gets long. Cel's advantages over Pi: multi-agent planning (Reasoning→Planning→Execution→Validation), backend large-data analysis, native chart creation.

### Gotcha

<!-- lore:019d6795-50a8-7000-aa5d-40a5574c1f44 -->
* **apply\_format execution missing border params**: apply\_format fails at THREE points: (1) AI planner doesn't generate border params; (2) App.tsx execution handler passes borders config but omits \`all: true\` property — ExcelAPI.applyFormat checks \`b.all\` at line 253 before applying borders, so silent failure occurs without it; (3) ExcelAPI.applyFormat must receive valid border object. Fix: Add \`all: true\` to borders config in App.tsx line 1438: \`{ all: true, color: ..., style: ..., weight: ... }\`. Also add 'border' to userWantsSelected regex (line 200) so execution uses selected range.

<!-- lore:019d6d15-a8ea-76fc-98b2-e7933531f83d -->
* **Chart creation uses wrong columns - planner lacks header context**: FIXED: Pass column headers to AI planner for intelligent matching. getSelectedRangeData() now extracts headers from first row and returns them in the range object. Backend extracts headers (both from JSON string or object format) and injects them into planner context via COLUMNS HEADERS section. AI now sees actual column names (e.g., 'Sheet Sheet2 columns: Total Distance, Current Pressure...') and can intelligently match user requests to correct columns instead of guessing wrong axes.

<!-- lore:019d717e-f4b6-7d56-b103-7ce317e629ad -->
* **Chart includes wrong numeric columns**: Chart includes wrong columns - root cause: createChartWithTwoColumns uses Math.min/max (lines 493-497 App.tsx), so 'Player Name' at index 0 and 'Wickets' at index 3 creates range A:D, INCLUDES columns B,C (Matches, Runs) BETWEEN them. Fix: Use exact column letters instead of min/max range: \`const col1Letter = String.fromCharCode(65 + startColIdx + col1Idx); const col2Letter = String.fromCharCode(65 + startColIdx + col2Idx); chartRange = col1Letter + startRow + ':' + col2Letter + endRow\`. Also add 'name' to skipWords (line 432) to prevent duplicate matches to same column.

<!-- lore:019d717e-f4b5-7108-86f4-3d75ed1623aa -->
* **Chart X-axis shows row numbers instead of labels**: Chart X-axis shows row numbers: createChart (excel-api.ts:510) and createChartFromTwoColumns (line 652) still use seriesBy: 'Columns' not 'Auto'. The setCategoryLabels fix WAS applied to createChartFromTwoColumns (lines 655-659) but the seriesBy change was never applied. Fix: Change 'Columns' to Excel.ChartSeriesBy.auto in BOTH functions.

<!-- lore:019d71e9-8663-7e65-b570-7e71a0603566 -->
* **Pie charts fail without data aggregation - need category sums not raw rows**: Pie charts require aggregated data (sum by category), but current code passes raw row data (e.g., C1:C9, E1:E9). For 'Sales by Category', need: Electronics=2700, Fashion=1250, Accessories=950—not individual rows. Fix: Before creating pie charts, aggregate data by summing numeric column for each unique category value, write aggregated summary to helper cells, then create pie chart from aggregated range.

<!-- lore:019d6bab-643c-75b4-ae4d-793ca73e0b4a -->
* **Reasoning agent disconnected from planner - output never used**: Agent pipeline fully connected and verified: Reasoning→Planner receives {reasoning\_output} (main.py:1066). Planner→Validator receives {plan\_info} via ExecuteRequest (App.tsx:1288). Frontend captures plan from streaming (App.tsx:215). Pipeline is COMPLETE - core features (recovery checkpoints, persistent rules, formatting conventions, memory system, dynamic model selection, streaming) are solid. Remaining nice-to-haves: Session management, Context compaction, Audit log, Extensions system.

<!-- lore:019d67cd-f8f9-7f53-a0e6-ea9eab78b324 -->
* **Reasoning agent outputs manual Excel steps instead of agentic actions**: Reasoning agent was generating non-agentic output like "Go to Home tab → click Borders" instead of agentic "I will apply black borders using apply\_format". Fix: Updated REASONING\_PROMPT and reasoning\_task in backend/main.py (lines ~60 and ~1001) to explicitly instruct the AI to describe what actions it WILL take, not manual Excel steps. The prompt now says: "Explain what YOU (the AI assistant) will do, not what the user should do manually in Excel".

<!-- lore:019d7638-5ed7-7790-99a0-0a89fd3cda12 -->
* **Reflection agent hardcodes OpenAI provider**: Reflection agent was hardcoded to OpenAI but is now FIXED. The \`/api/reflect\` endpoint (main.py:2052) now uses CrewAI agents via \`setup\_env()\` like all other endpoints, supporting any configured provider (OpenAI, Anthropic, Google, OpenRouter). Previously it imported \`OpenAI\` directly and ignored the provider setting. The fix replaced the direct OpenAI client call with a CrewAI agent chain: create agent with \`setup\_env()\`, run reflection task, parse JSON from output.

<!-- lore:019d6be1-3a8d-795a-ba59-b1cf10ce9f23 -->
* **window.confirm not supported in Office.js task pane**: Office.js task pane doesn't support window.confirm/alert. Solution: Created ConfirmModal component (src/components/ConfirmModal.tsx) with useConfirm hook that returns async confirm() function. Component renders modal with title, message, confirm/cancel buttons, and variant styles (danger/warning/primary). Import useConfirm in App.tsx, add ConfirmModalComponent to JSX, replace all 5 window.confirm calls with confirm(). Build passes, committed 7b34e7d.

### Pattern

<!-- lore:019d7658-638d-7e16-a544-54fea2b62b0f -->
* **Adding new Excel tools requires changes in 4 files**: Pattern: Adding a new Excel tool (e.g., conditional\_format, find\_replace, merge\_cells) requires coordinated changes across 4 files: (1) excel-api.ts — Office.js wrapper function; (2) agent-tools.ts — tool definition in excelTools array + handler in toolMap; (3) App.tsx — case in executeStep switch + add to VALID\_ACTIONS set; (4) backend/main.py — add to available actions list in TASK\_DESCRIPTION with instructions. Missing any of these 4 causes silent failures or 'unknown action' errors.

<!-- lore:019d7638-5ed9-7ae2-8dd2-e083696335b3 -->
* **Chart column matching logic duplicated frontend/backend**: Dual-path execution: chart requests bypass reasoning/planner pipeline entirely. App.tsx has two independent code paths: (1) General chat: frontend → /api/chat → Reasoning→Planner→Validator→Execution via SSE streaming; (2) Chart/analysis: App.tsx lines 348-843 detect chart/analysis requests FIRST, run frontend column matching, call backend APIs directly (/api/resolve-chart-intent, /api/analyze-large), and return early — never hitting the reasoning/planner pipeline. This means chart requests skip reasoning agent output entirely. Chart matching logic is duplicated between frontend (App.tsx ~450 lines) and backend (/api/resolve-chart-intent). Should consolidate to backend-only to reduce drift. Frontend-first chart path blocks adding reasoning to chart requests without refactoring.

<!-- lore:019d7658-638e-7ab9-988b-209988d9b904 -->
* **Checkpoint-based undo via header button**: Checkpoint system enables one-click undo: checkpoints auto-snapshot cell values before every write mutation via lib/memory.ts (saveCheckpoint function). Undo button in header bar (Rewind icon) calls removeCheckpoint to restore most recent snapshot. Checkpoints stored per-workbook in localStorage. The undo flow: click undo → getCheckpoints() → find most recent → Office.js context.sync to restore cell values → remove checkpoint from storage. This replaces window.confirm-based 'are you sure?' with actual reversibility.

<!-- lore:019d6bc7-6b8f-7cf6-b276-9221e16ce661 -->
* **Excel AI robustness gaps - research-driven priorities**: Excel AI robustness gaps - IMPLEMENTED: All Priority 1 features in App.tsx: (1) Formula Validation - checks unbalanced parentheses, missing VLOOKUP/IF args; (2) Error Detection - warns about VLOOKUP #N/A risk, #DIV/0! in formulas, empty cells, mixed text/numbers; (3) Operation Timeout - 2min max prevents infinite loops. Pre-execution Review was REMOVED - it showed confirmation modal every time, annoying for scale. Now only conditional confirmations: Delete (destructive), Large write (>1000 rows), Address mismatch, Error detection warnings. Plan still visible in chat without blocking execution.

<!-- lore:019d6775-cb41-736a-aac8-2546d972493d -->
* **Memory system for persistent preferences**: Memory and context system: Cel uses localStorage for persistent preferences (cel-user-rules, cel-formatting-conventions) and tracks conversation history for multi-turn context. Frontend sends last 3 exchanges (6 messages) as conversation\_history array with each /api/chat request. Backend ChatRequest accepts conversation\_history: List\[dict], passes it to planner and reasoning agents via {history\_info} placeholder in TASK\_DESCRIPTION. Rules, conventions, and conversation history are separate fields in the API payload, injected into system prompts. Checkpoints auto-snapshot cell values before mutations for one-click undo. UI in SettingsPanel: Memory tab (RulesEditor, CheckpointList, OperationHistory), Conventions tab (ConventionsEditor).

### Preference

<!-- lore:019d6743-3186-7ff3-ad4e-0c573f8a65e1 -->
* **User tends to share GitHub repos seeking solutions**: User tends to share GitHub repos seeking solutions: User frequently shares external GitHub repositories as potential solutions for Cel. Most are overkill or solve different problems. Pattern: user looks for existing libraries before building. Best response: give direct yes/no on usefulness with a brief comparison table, then redirect to what Cel actually needs. Avoid deep dives into repos that don't fit.
<!-- End lore-managed section -->
