<!-- This section is maintained by the coding agent via lore (https://github.com/BYK/opencode-lore) -->
## Long-term Knowledge

### Architecture

<!-- lore:019d6743-3189-7a6b-8909-7340f5976d0b -->
* **Cel being positioned as SaaS product**: User views Cel through a SaaS lens, not just a personal tool. This affects architectural decisions: needs user auth, multi-tenant support, cloud storage, and billing eventually. Current localStorage approach is sufficient for now, but migration path to Convex or Supabase should be planned. Project already has Supabase available — prefer it over adding Convex unless Convex offers clear advantages for the specific use case.

<!-- lore:019d6bb6-491c-7cc9-a44a-bd4796e53ce4 -->
* **Model defaults removed - fully dynamic model selection**: Model is now fully dynamic - no hardcoded defaults. All Pydantic request classes in backend/main.py (ChatRequest, ExecuteRequest, AnalyzeRequest, ReflectRequest) use \`model: str = ""\` instead of \`model: str = "gpt-4o"\`. Same change in frontend SettingsPanel.tsx and App.tsx. This forces clients to explicitly provide model ID - any model works (gpt-4o, claude-sonnet, gemini-pro, custom OpenRouter models). Fallback logic was removed since frontend always sends model from settings.

<!-- lore:019d6775-cb40-76de-a459-106c58501ab8 -->
* **Pi for Excel as reference architecture**: Pi for Excel (tmustier/pi-for-excel, 290 stars) serves as the primary reference architecture for Cel's feature roadmap. Key patterns to adopt: (1) Recovery checkpoints — snapshot cell values before mutations for one-click undo; (2) Persistent rules — user/workbook-level guidance stored and followed by AI; (3) Formatting conventions — currency, date, number preferences applied consistently; (4) Session management — multiple conversations per workbook with auto-save/restore; (5) Context compaction — summarize old messages when context gets long. Cel's advantages over Pi: multi-agent planning (Reasoning→Planning→Execution→Validation), backend large-data analysis, native chart creation.

### Gotcha

<!-- lore:019d6795-50a8-7000-aa5d-40a5574c1f44 -->
* **apply\_format execution missing border params**: apply\_format fails at THREE points: (1) AI planner doesn't generate border params; (2) App.tsx execution handler passes borders config but omits \`all: true\` property — ExcelAPI.applyFormat checks \`b.all\` at line 253 before applying borders, so silent failure occurs without it; (3) ExcelAPI.applyFormat must receive valid border object. Fix: Add \`all: true\` to borders config in App.tsx line 1438: \`{ all: true, color: ..., style: ..., weight: ... }\`. Also add 'border' to userWantsSelected regex (line 200) so execution uses selected range.

<!-- lore:019d6bab-643c-75b4-ae4d-793ca73e0b4a -->
* **Reasoning agent disconnected from planner - output never used**: Agent pipeline fully connected and verified: Reasoning→Planner receives {reasoning\_output} (main.py:1066). Planner→Validator receives {plan\_info} via ExecuteRequest (App.tsx:1288). Frontend captures plan from streaming (App.tsx:215). Pipeline is COMPLETE - core features (recovery checkpoints, persistent rules, formatting conventions, memory system, dynamic model selection, streaming) are solid. Remaining nice-to-haves: Session management, Context compaction, Audit log, Extensions system.

<!-- lore:019d67cd-f8f9-7f53-a0e6-ea9eab78b324 -->
* **Reasoning agent outputs manual Excel steps instead of agentic actions**: Reasoning agent was generating non-agentic output like "Go to Home tab → click Borders" instead of agentic "I will apply black borders using apply\_format". Fix: Updated REASONING\_PROMPT and reasoning\_task in backend/main.py (lines ~60 and ~1001) to explicitly instruct the AI to describe what actions it WILL take, not manual Excel steps. The prompt now says: "Explain what YOU (the AI assistant) will do, not what the user should do manually in Excel".

<!-- lore:019d6be1-3a8d-795a-ba59-b1cf10ce9f23 -->
* **window.confirm not supported in Office.js task pane**: Office.js task pane runs in a restricted environment that doesn't support browser dialogs like window.conf() and window.alert(). Pre-execution review feature (window.confirm before executing plan) fails with "Function window.confirm is not supported". Fix: Replace all window.confirm calls in App.tsx (lines 607, 625, 645, 662, 723) with a custom modal component using React state, or add try-catch fallback that silently proceeds when confirm is unavailable.

### Pattern

<!-- lore:019d6bc7-6b8f-7cf6-b276-9221e16ce661 -->
* **Excel AI robustness gaps - research-driven priorities**: Excel AI robustness gaps - IMPLEMENTED: All Priority 1 features now complete in App.tsx: (1) Formula Validation - checks unbalanced parentheses, missing VLOOKUP/IF args; (2) Error Detection - warns about VLOOKUP #N/A risk, #DIV/0! in formulas, empty cells, mixed text/numbers; (3) Operation Timeout - 2min max prevents infinite loops; (4) Pre-execution Review - shows full plan with window.confirm before execution. Committed in 0b5dbeb. Nice-to-have: session management, context compaction, audit log.

<!-- lore:019d6775-cb41-736a-aac8-2546d972493d -->
* **Memory system for persistent preferences**: Cel uses a lightweight memory system (lib/memory.ts) for persistent user preferences and formatting conventions. Store in localStorage with keys: 'cel-user-rules' (array of rule strings), 'cel-formatting-conventions' (currency symbol, date format, number format, negative style). Memory context is injected into API calls alongside workbook\_context and selected\_range. Rules and conventions are passed as separate fields to the backend for system prompt injection. UI is built in SettingsPanel with two new tabs: 'Memory' (RulesEditor, CheckpointList, OperationHistory components) and 'Conventions' (ConventionsEditor component). Checkpoints auto-snapshot before write operations with one-click restore.

### Preference

<!-- lore:019d6743-3186-7ff3-ad4e-0c573f8a65e1 -->
* **User tends to share GitHub repos seeking solutions**: User tends to share GitHub repos seeking solutions: User frequently shares external GitHub repositories as potential solutions for Cel. Most are overkill or solve different problems. Pattern: user looks for existing libraries before building. Best response: give direct yes/no on usefulness with a brief comparison table, then redirect to what Cel actually needs. Avoid deep dives into repos that don't fit.
<!-- End lore-managed section -->
