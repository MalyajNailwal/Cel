<!-- This section is maintained by the coding agent via lore (https://github.com/BYK/opencode-lore) -->
## Long-term Knowledge

### Architecture

<!-- lore:019d7740-6cd2-7dc8-9279-99fe9cd99683 -->
* **Agent pipeline lacks clarification and adaptive execution**: Agent pipeline gaps now addressed: (1) Clarification agent runs after reasoning, before planning — detects ambiguous requests via JSON response with \`needs\_clarification\` flag, yields \`clarification\` SSE event, frontend stops processing and shows question. (2) Adaptive executor detects critical failures (sheet not found, range invalid, Excel not ready) via isCriticalFailure regex during execution, halts remaining steps, triggers immediate reflection for recovery. Non-critical failures still cascade normally. Both implemented in backend/main.py and src/App.tsx.

<!-- lore:019d6743-3189-7a6b-8909-7340f5976d0b -->
* **Cel being positioned as SaaS product**: User views Cel through a SaaS lens, not just a personal tool. This affects architectural decisions: needs user auth, multi-tenant support, cloud storage, and billing eventually. Current localStorage approach is sufficient for now, but migration path to Convex or Supabase should be planned. Project already has Supabase available — prefer it over adding Convex unless Convex offers clear advantages for the specific use case.

<!-- lore:019eabba-5082-7f4c-b968-c2bdf4bb11df -->
* **Context handoff map - where each LLM operates with incomplete data**: Agent pipeline context gaps — most LLMs operate with incomplete data. Planner is BEST-informed (all context except cross-sheet/formula state). Reasoning agent missing memory\_info and data\_profile — output feeds into planner but may miss user conventions. Validator only summarizes, doesn't verify intent. Reflector has WORST context (only error messages, no headers/schema/memory). Fix for the missing feedback loop: mid-execution re-planning — on step failure, send completed steps + error + original context back to /api/chat so the planner regenerates remaining steps with actual execution state. This is superior to reflection which only sees error messages. No state feedback loop between execution and re-planning except via reflection.

<!-- lore:019d6bb6-491c-7cc9-a44a-bd4796e53ce4 -->
* **Model defaults removed - fully dynamic model selection**: Model is now fully dynamic - no hardcoded defaults. All Pydantic request classes in backend/main.py (ChatRequest, ExecuteRequest, AnalyzeRequest, ReflectRequest) use \`model: str = ""\` instead of \`model: str = "gpt-4o"\`. Same change in frontend SettingsPanel.tsx and App.tsx. This forces clients to explicitly provide model ID - any model works (gpt-4o, claude-sonnet, gemini-pro, custom OpenRouter models). Fallback logic was removed since frontend always sends model from settings.

<!-- lore:019d6775-cb40-76de-a459-106c58501ab8 -->
* **Pi for Excel as reference architecture**: Pi for Excel (tmustier/pi-for-excel, 290 stars) serves as the primary reference architecture for Cel's feature roadmap. Key patterns to adopt: (1) Recovery checkpoints — snapshot cell values before mutations for one-click undo; (2) Persistent rules — user/workbook-level guidance stored and followed by AI; (3) Formatting conventions — currency, date, number preferences applied consistently; (4) Session management — multiple conversations per workbook with auto-save/restore; (5) Context compaction — summarize old messages when context gets long. Cel's advantages over Pi: multi-agent planning (Reasoning→Planning→Execution→Validation), backend large-data analysis, native chart creation.

### Gotcha

<!-- lore:019d6795-50a8-7000-aa5d-40a5574c1f44 -->
* **apply\_format execution missing border params**: apply\_format fails at THREE points: (1) AI planner doesn't generate border params; (2) App.tsx execution handler passes borders config but omits \`all: true\` property — ExcelAPI.applyFormat checks \`b.all\` at line 253 before applying borders, so silent failure occurs without it; (3) ExcelAPI.applyFormat must receive valid border object. Fix: Add \`all: true\` to borders config in App.tsx line 1438: \`{ all: true, color: ..., style: ..., weight: ... }\`. Also add 'border' to userWantsSelected regex (line 200) so execution uses selected range.

<!-- lore:019d717e-f4b6-7d56-b103-7ce317e629ad -->
* **Chart includes wrong numeric columns**: Chart creation gotchas — multiple issues: (1) X-axis shows row numbers instead of labels: createChart and createChartFromTwoColumns use seriesBy: 'Columns' — fix to Excel.ChartSeriesBy.auto in both functions. (2) Wrong columns included: createChartWithTwoColumns uses Math.min/max for range, pulling intermediate columns — use exact column letters instead. Add 'name' to skipWords. (3) Pie charts need aggregation: passing raw row data fails — aggregate by category sum first, write to helper cells, then create chart. (4) Planner lacked header context (FIXED): getSelectedRangeData() now extracts headers, backend injects via COLUMNS HEADERS section.

<!-- lore:019d7b57-6826-789d-bae4-93e1cf9b9bf5 -->
* **Office.js sort hasHeaders flag unreliable - exclude header row manually**: Office.js sort API's \`hasHeaders: true\` parameter doesn't work reliably in all Excel versions, causing the header row to be included in sort and reversing expected order. Fix: parse range address (e.g., 'A1:B4'), create new range starting from row 2 (e.g., 'A2:B4'), sort only data rows. This excludes header entirely and ensures predictable A-Z sorting.

<!-- lore:019eabba-5080-7fcc-a347-805c0688dc01 -->
* **Planner blind to other sheets - cross-sheet operations fail**: The planner only receives the selected range's values (backend/main.py:1737-1753). If user says 'copy data from Sheet2 to Sheet1', the planner has zero visibility into Sheet2's content. Fix via getWorkbookMetadata() — new Office.js function that returns full workbook structure: sheet names, row/col counts, headers, tables, charts, named ranges. Inject this as workbook\_metadata into planner context. For cross-sheet requests, additionally pre-scan the target sheet (headers + first 3 rows) before planning. Schema caching: hash workbook structure, only refetch when sheet count/names change.

<!-- lore:019eabba-5081-7369-a5f2-31f9bfeed132 -->
* **Selected range override discards planner address intelligence**: When userWantsSelected=true (App.tsx:1071-1086), the frontend REPLACES params.address and params.sheet\_name for ALL write operations with the actual selected range. If the planner intelligently chose a different range (e.g., for a formula referencing a specific range), it gets silently overridden. The userWantsSelected flag is computed via regex on user message (line 308) — includes 'selected', 'here', 'this range', 'border'. Formula references and multi-range operations may be corrupted by this override.

<!-- lore:019d6be1-3a8d-795a-ba59-b1cf10ce9f23 -->
* **window.confirm not supported in Office.js task pane**: Office.js task pane doesn't support window.confirm/alert. Solution: Created ConfirmModal component (src/components/ConfirmModal.tsx) with useConfirm hook that returns async confirm() function. Component renders modal with title, message, confirm/cancel buttons, and variant styles (danger/warning/primary). Import useConfirm in App.tsx, add ConfirmModalComponent to JSX, replace all 5 window.confirm calls with confirm(). Build passes, committed 7b34e7d.

### Pattern

<!-- lore:019eabba-5081-7369-a5f2-31fad66f87ba -->
* **agent-tools.ts is dead code - plan-based approach replaced tool calling**: Adding new Excel tools requires changes in 3 files (agent-tools.ts is dead code): (1) excel-api.ts — Office.js wrapper function; (2) App.tsx — case in executeStep switch + add to VALID\_ACTIONS set; (3) backend/main.py — add to available actions list in TASK\_DESCRIPTION with instructions. Missing any causes silent failures. agent-tools.ts has 22 tool definitions but is never imported — plan-based approach (JSON plan → executeStep switch) replaced direct tool-calling. Duplicates info in TASK\_DESCRIPTION.

<!-- lore:019d7638-5ed9-7ae2-8dd2-e083696335b3 -->
* **Chart column matching logic duplicated frontend/backend**: Chart pipeline bypasses entire agent pipeline — HIGH impact. App.tsx:378-896 handles charts in frontend, duplicating ~250 lines of column-matching regex that also exists in backend /api/resolve-chart-intent. No reasoning, no planner, no memory, no reflection on chart failures. Consolidation strategy: add create\_chart action to planner vocabulary (type, xAxis, yAxis, title, style), route chart requests through /api/chat same as everything else, frontend executes resulting plan. This eliminates duplicated regex entirely and gives charts access to multi-agent planning. ~450 lines of duplicate logic between App.tsx and main.py deleted.

<!-- lore:019d7658-638e-7ab9-988b-209988d9b904 -->
* **Checkpoint-based undo via header button**: Checkpoint system enables one-click undo: checkpoints auto-snapshot cell values before every write mutation via lib/memory.ts (saveCheckpoint function). Undo button in header bar (Rewind icon) calls removeCheckpoint to restore most recent snapshot. Checkpoints stored per-workbook in localStorage. The undo flow: click undo → getCheckpoints() → find most recent → Office.js context.sync to restore cell values → remove checkpoint from storage. This replaces window.confirm-based 'are you sure?' with actual reversibility.

<!-- lore:019eabbe-16c5-7d72-9e3a-8f5a21afe59f -->
* **Confidence-based execution tiers**: Confidence-based execution tiers (IMPLEMENTED): Planner outputs confidence field (0.0-1.0) per step. TASK\_DESCRIPTION instructs scoring: 0.90-1.00 high (straightforward writes), 0.70-0.89 medium (charts/formatting), 0.50-0.69 low (ambiguous references), <0.50 very low (guesswork). App.tsx execution policy: <0.50 shows confirmation modal listing uncertain steps, 0.50-0.80 auto-executes but logs info, >0.80 auto-executes silently. Plan preview in chat shows confidence icons: ~ for medium (0.50-0.80), ! for low (<0.50). Planner prompt in main.py TASK\_DESCRIPTION includes scoring rules based on clarity, data availability, action destructiveness.

<!-- lore:019d774a-9881-7117-9944-4ac344d84279 -->
* **Critical failure detection + adaptive halt during execution**: Critical failure detection + adaptive halt: Adaptive executor uses isCriticalFailure regex to detect errors (sheet not found, range invalid, Excel not ready). When detected, remaining steps skipped and /api/reflect called for recovery. Non-critical errors (table overlap, formula errors) still cascade. Gotcha: reflector has worst context of any LLM call — only error messages, no headers/schema/data\_profile/memory — so recovery plans may repeat mistakes.

<!-- lore:019d774a-9880-7e96-af8a-e9e05b9e1d4a -->
* **Data Context Agent - deterministic profiling for planner**: Data profiling injected into planner context via deterministic profile\_data() function (backend/main.py). Expanded scope: null %, duplicate count, mixed type detection, outlier detection per column, PLUS unique value count, sample values, and basic statistics for numeric fields (min, max, mean, median). Output injected via {data\_profile} placeholder in TASK\_DESCRIPTION format string. No LLM needed — pure computation on selected range data. Returns empty string if no data available. Gives planner richer column intelligence for chart generation, formula creation, filtering, and aggregation decisions.

<!-- lore:019d6bc7-6b8f-7cf6-b276-9221e16ce661 -->
* **Excel AI robustness gaps - research-driven priorities**: Excel AI robustness gaps - IMPLEMENTED across multiple phases: Priority 1 (App.tsx): Formula Validation (unbalanced parens, missing VLOOKUP/IF args), Error Detection (VLOOKUP #N/A risk, #DIV/0!, empty cells, mixed text/numbers), Operation Timeout (2min max). Pre-execution Review REMOVED — now only conditional confirmations: Delete (destructive), Large write (>1000 rows), Address mismatch, Error detection warnings. Phase 3.3 added pre-execution validation IN the execution loop: (1) Sheet existence — verifies sheet\_name exists in workbook before writing, skips with error listing available sheets if not found; (2) Formula syntax — checks balanced parentheses and VLOOKUP arg count (needs 4); (3) Data type — validates sort column index is non-negative integer.

<!-- lore:019d6775-cb41-736a-aac8-2546d972493d -->
* **Memory system for persistent preferences**: Memory and context system: Cel uses localStorage for persistent preferences (cel-user-rules, cel-formatting-conventions) and tracks conversation history for multi-turn context. Frontend sends last 3 exchanges (6 messages) as conversation\_history array with each /api/chat request. Backend ChatRequest accepts conversation\_history: List\[dict], passes it to planner and reasoning agents via {history\_info} placeholder in TASK\_DESCRIPTION. Rules, conventions, and conversation history are separate fields in the API payload, injected into system prompts. Checkpoints auto-snapshot cell values before mutations for one-click undo. UI in SettingsPanel: Memory tab (RulesEditor, CheckpointList, OperationHistory), Conventions tab (ConventionsEditor).

<!-- lore:019eabbe-16c4-7407-89c2-222aeb069221 -->
* **Semantic column resolution via synonym maps**: Implement synonym map in backend/main.py for fuzzy matching between user terminology and workbook headers. Example: Sales↔Revenue↔Amount↔Income, Customer↔Client↔Buyer, Product↔Item↔SKU, Date↔Time↔Period. Planner uses this to resolve ambiguous column references — 'use the Sales column' works even if header says 'Revenue Amount'. Match confidence threshold: require >0.6 similarity to use a synonym, otherwise ask clarification. This reduces planner failures from exact-name mismatches and improves natural-language understanding. Maintain map as a dict in main.py, inject into planner context alongside headers.

### Preference

<!-- lore:019d6743-3186-7ff3-ad4e-0c573f8a65e1 -->
* **User tends to share GitHub repos seeking solutions**: User tends to share GitHub repos seeking solutions: User frequently shares external GitHub repositories as potential solutions for Cel. Most are overkill or solve different problems. Pattern: user looks for existing libraries before building. Best response: give direct yes/no on usefulness with a brief comparison table, then redirect to what Cel actually needs. Avoid deep dives into repos that don't fit.
<!-- End lore-managed section -->
