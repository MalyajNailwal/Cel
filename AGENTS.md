<!-- This section is maintained by the coding agent via lore (https://github.com/BYK/opencode-lore) -->
## Long-term Knowledge

### Architecture

<!-- lore:019d6743-3189-7a6b-8909-7340f5976d0b -->
* **Cel being positioned as SaaS product**: User views Cel through a SaaS lens, not just a personal tool. This affects architectural decisions: needs user auth, multi-tenant support, cloud storage, and billing eventually. Current localStorage approach is sufficient for now, but migration path to Convex or Supabase should be planned. Project already has Supabase available — prefer it over adding Convex unless Convex offers clear advantages for the specific use case.

<!-- lore:019d6775-cb40-76de-a459-106c58501ab8 -->
* **Pi for Excel as reference architecture**: Pi for Excel (tmustier/pi-for-excel, 290 stars) serves as the primary reference architecture for Cel's feature roadmap. Key patterns to adopt: (1) Recovery checkpoints — snapshot cell values before mutations for one-click undo; (2) Persistent rules — user/workbook-level guidance stored and followed by AI; (3) Formatting conventions — currency, date, number preferences applied consistently; (4) Session management — multiple conversations per workbook with auto-save/restore; (5) Context compaction — summarize old messages when context gets long. Cel's advantages over Pi: multi-agent planning (Reasoning→Planning→Execution→Validation), backend large-data analysis, native chart creation.

### Pattern

<!-- lore:019d6775-cb41-736a-aac8-2546d972493d -->
* **Memory system for persistent preferences**: Cel uses a lightweight memory system (lib/memory.ts) for persistent user preferences and formatting conventions. Store in localStorage with keys: 'cel-user-rules' (array of rule strings), 'cel-formatting-conventions' (currency symbol, date format, number format, negative style). Memory context is injected into API calls alongside workbook\_context and selected\_range. Rules and conventions are passed as separate fields to the backend for system prompt injection. UI is built in SettingsPanel with two new tabs: 'Memory' (RulesEditor, CheckpointList, OperationHistory components) and 'Conventions' (ConventionsEditor component). Checkpoints auto-snapshot before write operations with one-click restore.

### Preference

<!-- lore:019d6743-3186-7ff3-ad4e-0c573f8a65e1 -->
* **User tends to share GitHub repos seeking solutions**: User tends to share GitHub repos seeking solutions: User frequently shares external GitHub repositories as potential solutions for Cel. Most are overkill or solve different problems. Pattern: user looks for existing libraries before building. Best response: give direct yes/no on usefulness with a brief comparison table, then redirect to what Cel actually needs. Avoid deep dives into repos that don't fit.
<!-- End lore-managed section -->
