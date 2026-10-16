from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
import json
import asyncio
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional, List, Any
import os
import json
import re
import random
from collections import Counter
from contextlib import asynccontextmanager
import logging

from llm import call_llm, aclose as close_llm

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
logger = logging.getLogger("cel")

app = FastAPI()


@asynccontextmanager
async def lifespan(_: FastAPI):
    yield
    await close_llm()


app.router.lifespan_context = lifespan

_origins = os.environ.get(
    "CEL_ALLOWED_ORIGINS",
    "https://localhost:3000,https://127.0.0.1:3000",
).split(",")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[o.strip() for o in _origins if o.strip()],
    allow_methods=["GET", "POST"],
    allow_headers=["Content-Type"],
)


class ChatRequest(BaseModel):
    message: str
    provider: str = "openai"
    model: str = ""
    api_key: str
    workbook_context: Optional[str] = None
    selected_range: Optional[str] = None
    memory_context: Optional[str] = None
    enable_reasoning: Optional[bool] = True
    conversation_history: Optional[List[dict]] = []
    workbook_metadata: Optional[str] = None


class ExecuteRequest(BaseModel):
    message: str
    plan: list = []
    provider: str = "openai"
    model: str = ""
    api_key: str
    results: list


class AnalyzeRequest(BaseModel):
    data: List[List[Any]]
    question: str
    provider: str = "openai"
    model: str = ""
    api_key: str
    headers: Optional[List[str]] = None


class GenerateDataRequest(BaseModel):
    data_type: str = "employee"
    count: int = 100


class ChatResponse(BaseModel):
    content: str
    plan: list
    reasoning: Optional[str] = None


CLARIFICATION_PROMPT = """CLARIFICATION - check BEFORE planning.

If the request is CLEAR and SPECIFIC (e.g. "add bold to A1:A10", "create bar chart",
"sum column B", "replace X with Y", "add borders to selected cells", "sort by Salary
descending", "add a new sheet called Dashboard", "create a table with 100 employees"),
do NOT ask. Return an empty clarification_question.

Only ask ONE question when the request is genuinely ambiguous:
- "clean the data" -> what specifically? duplicates, empty cells, formatting, or other?
- "make it look better" -> borders, colors, bold headers, or column widths?
- "fix the issues" -> which: formula errors, duplicates, empty cells, type mismatches?
- "organize this" -> sort, group, filter, or create a table?
- "format this properly" -> currency, percentages, dates, bold headers, or conditional?

If there ARE column headers available and the user names specific columns, it is CLEAR.
If there is NO selected range and the request does not reference specific data, lean toward asking."""

# Semantic column resolution — maps user terminology to workbook headers
SYNONYM_MAP = {
    "sales": ["revenue", "amount", "income", "earnings", "proceeds", "net sales", "gross sales"],
    "customer": ["client", "buyer", "purchaser", "account", "consumer", "patron"],
    "product": ["item", "sku", "goods", "merchandise", "article", "product name"],
    "date": ["time", "day", "month", "year", "period", "transaction date", "order date"],
    "profit": ["earnings", "net income", "margin", "gain", "net profit", "gross profit"],
    "cost": ["expense", "expenditure", "spending", "outlay", "cogs", "cost of goods"],
    "quantity": ["qty", "count", "volume", "units", "num", "amount"],
    "price": ["rate", "unit price", "cost per unit", "msrp", "list price"],
    "region": ["area", "territory", "zone", "location", "district", "market"],
    "category": ["type", "group", "segment", "class", "subcategory"],
    "employee": ["staff", "worker", "personnel", "team member", "associate"],
    "order": ["purchase", "transaction", "sale", "invoice"],
    "discount": ["reduction", "markdown", "off", "savings"],
    "total": ["sum", "amount", "grand total", "overall"],
    "name": ["title", "label", "description", "full name"],
    "status": ["state", "condition", "stage", "phase"],
    "rating": ["score", "grade", "rank", "evaluation"],
    "feedback": ["review", "comment", "response", "opinion"],
}


def _normalize_text(value: Any) -> str:
    return str(value or "").strip().lower()


def _tokenize(value: str) -> List[str]:
    return [t for t in re.split(r"[^a-z0-9]+", value.lower()) if t]


def _score_header_match(header: str, term: str) -> float:
    h = _normalize_text(header)
    t = _normalize_text(term)
    if not h or not t:
        return 0.0
    if h == t:
        return 1.0
    if h in t or t in h:
        return 0.85
    h_tokens = _tokenize(h)
    t_tokens = _tokenize(t)
    if not h_tokens or not t_tokens:
        return 0.0
    overlap = len([x for x in h_tokens if x in t_tokens])
    if overlap == 0:
        return 0.0
    return overlap / max(len(h_tokens), len(t_tokens))


def _resolve_best_header_index(
    headers: List[str], term: str, threshold: float = 0.55
) -> Optional[int]:
    best_idx = None
    best_score = 0.0
    for i, header in enumerate(headers):
        s = _score_header_match(header, term)
        if s > best_score:
            best_score = s
            best_idx = i
    return best_idx if best_score >= threshold else None


def _is_numeric_column(rows: List[List[Any]], col_idx: int) -> bool:
    num_count = 0
    cat_count = 0
    sample_size = min(20, len(rows))
    for i in range(sample_size):
        val = rows[i][col_idx] if col_idx < len(rows[i]) else None
        if val in (None, ""):
            continue
        cleaned = re.sub(r"[^0-9.\-]", "", str(val))
        if cleaned and re.fullmatch(r"-?\d+(\.\d+)?", cleaned):
            num_count += 1
        else:
            cat_count += 1
    return num_count > cat_count


def _format_formula_context(sr: dict) -> str:
    """List cells that already hold a formula, so the planner does not flatten them."""
    cells = sr.get("formulaCells") or []
    if not cells:
        formulas = sr.get("formulas") or []
        cells = []
        for r, row in enumerate(formulas[:60]):
            for c, value in enumerate(row):
                if isinstance(value, str) and value.startswith("="):
                    letters = ""
                    n = c
                    while n >= 0:
                        letters = chr(65 + (n % 26)) + letters
                        n = n // 26 - 1
                    cells.append({"address": f"{letters}{r + 1}", "formula": value})
    if not cells:
        return ""
    lines = [f"  {c['address']}: {c['formula']}" for c in cells[:25] if isinstance(c, dict)]
    if len(cells) > 25:
        lines.append(f"  ... and {len(cells) - 25} more")
    return "\n".join(lines)


def _format_number_formats(sr: dict) -> str:
    """Surface the number formats already applied, so currency/percent survives."""
    formats = sr.get("numberFormats") or []
    if not formats:
        return ""
    return ", ".join(str(f) for f in formats[:12])


def _numeric_column_list(values: list, headers: list) -> str:
    """Name the columns a chart or aggregate can safely treat as numeric."""
    rows = values[1:] if len(values) > 1 else []
    if not rows:
        return ""
    names = [headers[i] for i in range(len(headers)) if _is_numeric_column(rows, i)]
    return ", ".join(names) if names else "None - all columns look like text"


def resolve_semantic_matches(user_message: str, headers: list) -> str:
    """Map user terminology onto real headers using synonyms + token overlap."""
    if not headers:
        return ""

    msg_lower = user_message.lower()
    matches: list[str] = []

    for term, synonyms in SYNONYM_MAP.items():
        if term not in msg_lower and not any(syn in msg_lower for syn in synonyms):
            continue
        best_idx = _resolve_best_header_index(headers, term, threshold=0.55)
        if best_idx is None:
            for syn in synonyms:
                best_idx = _resolve_best_header_index(headers, syn, threshold=0.55)
                if best_idx is not None:
                    break
        if best_idx is not None:
            matches.append(f'User "{term}" -> header "{headers[best_idx]}"')

    # Token overlap catches literal mentions the synonym map has no entry for.
    for i, header in enumerate(headers):
        if any(f'header "{header}"' in m for m in matches):
            continue
        h = _normalize_text(header)
        if len(h) < 3:
            continue
        if re.search(rf"\b{re.escape(h)}\b", msg_lower):
            matches.append(f'Exact mention -> header "{header}" (index {i})')

    if not matches:
        return ""
    return "SEMANTIC COLUMN MATCHES:\n" + "\n".join(f"  {m}" for m in matches)


SYSTEM_PROMPT = """You are an expert Excel analyst. You create detailed, executable plans for Excel operations.
Each step must have a specific action, target range, and all parameters needed.
Always use A1 notation for ranges. Always read before writing if you need context.

IMPORTANT - For creating tables with data (e.g., "create table with 100 people"):
1. You MUST generate the actual data in the "values" field of set_values
2. The "values" field must be a 2D array: [["Header1", "Header2"], ["row1col1", "row1col2"], ...]
3. Generate at least 10 rows of realistic sample data
4. Then create the table using create_table on that range
5. NEVER leave "values" empty - always populate with actual data"""

TASK_DESCRIPTION = """Create a JSON plan for this Excel request: {message}

{context_info}
{selected_info}
{workbook_metadata_info}

COLUMN HEADERS (use these for charts!):
{headers_context}

SEMANTIC MATCHES (user terms mapped to actual headers):
{semantic_matches}

NUMERIC COLUMNS (safe Y-axis candidates for charts and aggregates):
{numeric_columns}

EXISTING FORMULAS IN THE SELECTION (do NOT overwrite these with literal values;
use set_formulas to change a calculation):
{existing_formulas}

NUMBER FORMATS ALREADY IN USE:
{number_formats}

TABLE SCHEMA (deterministic context agent output):
{schema_context}

DATA QUALITY PROFILE:
{data_profile}

{memory_info}

{history_info}

{clarification_section}

Available actions: get_workbook_structure, get_selected_range, get_range, get_sheet_data,
set_values, set_formulas, apply_format, insert_rows, delete_rows, insert_columns,
delete_columns, add_worksheet, delete_worksheet, create_table, sort_range, auto_fill, create_chart,
conditional_format, find_replace, merge_cells, unmerge_cells, add_dropdown,
copy_range, auto_fit_columns, set_row_height, set_column_width, freeze_panes,
add_comment, add_hyperlink, text_to_columns, remove_duplicates, create_pivot_table

IMPORTANT - For creating tables with sample data:
- First use set_values to write the header row and data rows
- Then use create_table to create an Excel table on that range
- Example: For "100 people" generate realistic sample data in set_values first

IMPORTANT - COLOR FORMATTING:
- When user asks for "color", "grey", "red", "blue", "highlight", "background color", "font color" → use apply_format with fill_color or font_color
- Color names to hex mapping: grey/gray=#808080, red=#FF0000, blue=#0000FF, green=#008000, yellow=#FFFF00, white=#FFFFFF, black=#000000, orange=#FFA500, purple=#800080, pink=#FFC0CB, brown=#A52A2A, navy=#000080, teal=#008080, maroon=#800000, olive=#808000, lime=#00FF00, coral=#FF7F50, salmon=#FA8072, tomato=#FF6347, indigo=#4B0082, violet=#EE82EE, gold=#FFD700, crimson=#DC143C, emerald=#50C878, jade=#00A86B, bronze=#CD7F32, steel=#4682B4, skyblue=#87CEEB, royalblue=#4169E1, dodgerblue=#1E90FF
- User can use ANY color name - if unknown, try common variations or default to #808080 (grey)

IMPORTANT - BORDERS:
- When user asks for "border", "add border", "outer border", "all borders", "grid lines", "pivot style border" → use apply_format with border_all=true, border_color, border_style, border_style options: Continuous, Dash, Dot, DashDot, Double; weight options: Hairline, Thin, Medium, Thick; default: border_style="Continuous", border_weight="Thin", border_color="#000000"
- Example: "add border" → apply_format with border_all=true, border_color="#000000", border_style="Continuous", border_weight="Thin"
- Example: "thick border" → apply_format with border_all=true, border_weight="Thick"

IMPORTANT - COPYING AND MOVING DATA:
- "copy A1:D10 to Sheet2", "move this to another sheet", "duplicate the data" -> use copy_range
- Params: source_address, target_address, source_sheet (optional), target_sheet (optional)
- copy_range preserves values, formulas, and formatting - always prefer it over get_range + set_values

IMPORTANT - PIVOT TABLES / SUMMARIES:
- "pivot by region", "summarize sales by product", "aggregate by category" -> use create_pivot_table
- Params: source_address, target_address, rows (array of field names), values (array of field names), columns (optional)
- Field names must match the header row EXACTLY - use the COLUMN HEADERS section above
- Example: source A1:D50, rows ["Region"], values ["Sales"], target "A1" -> writes a summary table

IMPORTANT - LAYOUT AND ANNOTATION:
- "make columns wider", "auto-fit", "fit the text" -> auto_fit_columns (optional address; omit to fit the whole sheet)
- "make the header taller", "increase row height" -> set_row_height with address and points
- "widen column B" -> set_column_width with address and points
- "freeze the header row", "keep row 1 visible" -> freeze_panes with rows=1 (rows=2, columns=1 also freezes column A)
- "add a note", "comment this cell" -> add_comment with address and text
- "link to X", "add a hyperlink" -> add_hyperlink with address, url, and optional text
- "split this column by comma", "separate the first and last name" -> text_to_columns with address and delimiter
- "remove duplicate rows", "dedupe" -> remove_duplicates (treats row 1 as the header)

IMPORTANT - CONDITIONAL FORMATTING:
- When user says "highlight cells > 100", "color cells where value is negative", "red if less than 50" → use conditional_format
- Params: address, rule_type="cellValue", operator (GreaterThan/LessThan/EqualTo/Between/etc.), formula1 (the threshold value), fill_color (highlight color)
- Example: "highlight cells > 100 in yellow" → conditional_format with address, operator="GreaterThan", formula1="100", fill_color="#FFFF00"
- Example: "red if negative" → conditional_format with operator="LessThan", formula1="0", fill_color="#FF0000"
- Example: "green if between 50 and 100" → conditional_format with operator="Between", formula1="50", formula2="100", fill_color="#00FF00"
- Other rule_type values: "colorScale" (3-colour heat map, optional colors array), "dataBar" (in-cell bars),
  "duplicateValues", "uniqueValues", "blanks", "aboveAverage", "topBottom" (formula1=rank, operator="top"/"bottom"/"percent"/"bottomPercent")

IMPORTANT - FIND AND REPLACE:
- When user says "replace X with Y", "change all 'old' to 'new'", "find and replace" → use find_replace
- Params: address, find_text, replace_text, match_case (optional, default false)
- Example: "replace 'USA' with 'United States'" → find_replace with find_text="USA", replace_text="United States"

IMPORTANT - MERGE/UNMERGE CELLS:
- When user says "merge these cells", "combine A1:D1" → use merge_cells with address
- When user says "unmerge", "split merged cells" → use unmerge_cells with address

IMPORTANT - DROPDOWNS:
- When user says "add dropdown", "dropdown menu", "data validation", "choose from list" → use add_dropdown
- Params: address (cell range, e.g., "B2:B100"), options (array of allowed values), sheet_name
- Example: "dropdown in categories column with theif and not theif" → add_dropdown with address="B2:B100", options=["theif", "not theif"]
- Example: "add a dropdown with yes/no options" → add_dropdown with address="A2:A100", options=["Yes", "No"]

IMPORTANT - SORTING:
- When user says "sort by X ascending/descending", "sort alphabetically", "sort A to Z" → use sort_range
- Params: address (full data range including header row, e.g., "A1:B4"), column_index (0-based index of sort column), ascending (true for A-Z/ascending, false for Z-A/descending)
- Example: "sort by name ascending" on range A1:B4 → sort_range with address="A1:B4", column_index=0, ascending=true
- Example: "sort salary descending" → sort_range with column_index=(index of Salary column), ascending=false
- The API handles header row automatically — first row stays as header during sort

For charts (create_chart), use these params:
- chart_type: "column", "bar", "line", "pie", "pie3D", "doughnut", "area", "scatter", "radar", "surface", "bubble"
- data_range: A1 notation range that contains ALL chart data including headers (e.g., "A1:B10")
- sheet_name: name of the sheet
- title: descriptive chart title (e.g., "Revenue by Region" not "Chart 1")
- position: {{"left": 300, "top": 50, "width": 400, "height": 300}} (optional, pixels from top-left)
- Optional: show_legend (bool), legend_position ("Bottom"/"Right"/"Top"/"Left"), show_data_labels (bool),
  x_axis_title, y_axis_title, gap_width, overlap, series_color
- Stacked variants: "columnStacked", "columnStacked100", "barStacked", "barStacked100", "areaStacked", "lineMarkersStacked"

CHART COLUMN SELECTION — use the COLUMN HEADERS and DATA QUALITY PROFILE above:
- For bar/column charts: categorical column (X-axis) + numeric column (Y-axis)
- For pie charts: one categorical column + one numeric column
- For line charts: one categorical/date column (X-axis) + one numeric column (Y-axis)
- For scatter: two numeric columns
- The DATA QUALITY PROFILE shows type=numeric/text per column — use this to pick correctly
- Include the header row in data_range (e.g., if headers are row 1 and data is rows 2-100, use "A1:B100")
- Title format: "{{Y-axis metric}} by {{X-axis category}}" (e.g., "Revenue by Region", "Sales Trend by Month")

Respond with a JSON object ONLY:
{{
  "reasoning": "1-2 plain sentences on what you will do and why",
  "needs_clarification": false,
  "clarification_question": "",
  "plan": [
    {{
      "action": "tool_name",
      "params": {{"key": "value"}},
      "description": "What this step does",
      "confidence": 0.9
    }}
  ],
  "response": "Natural language summary for the user"
}}

REASONING FIELD:
{reasoning_instruction}

CONFIDENCE SCORING — assign a confidence score (0.0-1.0) to each step:
- 0.90-1.00: High confidence — straightforward write/format with clear data and addresses
- 0.70-0.89: Medium confidence — chart creation, formatting across ranges, multi-step operations
- 0.50-0.69: Low confidence — ambiguous references, inferred column names, cross-sheet operations
- Below 0.50: Very low — guesswork, missing data, unclear user intent
Score based on: clarity of user request, data availability, address precision, action destructiveness.

CRITICAL RULES — FOLLOW THESE EXACTLY:
1. MULTIPLE REQUESTS: If the user asks for multiple things (e.g., "bar graph AND pie chart"), you MUST create SEPARATE steps for EACH item. Do NOT combine them into one step.
2. SELECTED RANGE IS KING: If the user says "selected area", "these cells", "this range", "highlighted", "here" — you MUST use the selected_range address and sheet_name as the target for ALL write operations (set_values, set_formulas, apply_format, create_chart, etc.)
3. When filling data into selected area: Use the selected range's EXACT address as the "address" param for set_values. The sheet_name from selected_range MUST be used as "sheet_name" param.
4. If selected range is A1:D10 on Sheet1 and user says "put fruits data here" → set_values with address="A1:D10", sheet_name="Sheet1"
5. NEVER create data in a different range than what the user selected unless they explicitly say "create a new sheet" or "start from A1"
6. If you need to know current data in selected range, call get_selected_range first
7. For charts, use the selected range's address as data_range. If multiple charts are requested, create multiple create_chart steps.
8. For set_values, params MUST include "address" and "values" (2D array with ACTUAL data - not placeholders)
9. For set_formulas, params must include "address" and "formulas" (2D array)
10. For apply_format, params must include "address" and format properties (bold, fill_color, font_color, number_format, etc.)
11. Be specific with cell ranges — match the selected range dimensions
12. Keep it simple and correct
13. WHEN GENERATING SAMPLE DATA: You MUST include the actual data in the "values" array. Example: "values": [["Name", "Age", "Gender"], ["John", 25, "Male"], ["Jane", 30, "Female"]]
14. NEVER return empty values or placeholder text in the values array - always generate real data
15. NEVER overwrite a cell listed in EXISTING FORMULAS with set_values - use set_formulas to change a calculation
16. Use copy_range for any "copy/move/duplicate data to another sheet" request, never get_range + set_values
17. For "sum by X" / "total per X" / "breakdown by X" use create_pivot_table instead of writing SUMIF formulas"""

VALIDATOR_TASK = """Original request: {message}

Planned approach:
{plan_info}

Execution results:
{results_str}

Provide a clear, friendly summary of what was accomplished. Be specific about cell ranges and what changed."""


REFLECTION_PROMPT = """You are a self-correction agent. Analyze failed Excel steps and create a recovery plan.

Original user request: {message}

Planned steps:
{plan}

Results:
{results}

DATA CONTEXT:
Headers: {headers_context}
Schema: {schema_context}

Rules:
1. Only create recovery steps for actions that FAILED (success: false)
2. If table creation failed with "overlap", suggest a different approach (e.g. just write the data without a table, or use a different range)
3. If set_values failed, fix the data dimensions to match the target range
4. If a formula failed, use the column headers above to build correct references
5. Keep recovery steps minimal - only fix what is broken
6. Use the column headers above to create accurate cell references
7. Return ONLY a valid JSON array of recovery steps

Return format:
[{{"action": "set_values", "params": {{"address": "A1", "values": [], "sheet_name": "Sheet1"}}, "description": "Retry with corrected data"}}]

If no recovery is possible, return: []"""


def compute_statistics(data: List[List[Any]]) -> dict:
    """Compute basic statistics from 2D data array."""
    if not data or not data[0]:
        return {"error": "No data to analyze"}

    flat_data = []
    headers = data[0] if len(data) > 0 else []
    rows = data[1:] if len(data) > 1 else []

    numeric_cols = []
    for col_idx in range(len(headers)):
        col_values = []
        for row in rows:
            if col_idx < len(row):
                val = row[col_idx]
                if val is not None and val != "":
                    try:
                        num = float(
                            str(val)
                            .replace(",", "")
                            .replace("₹", "")
                            .replace("$", "")
                            .replace("€", "")
                        )
                        col_values.append(num)
                    except (ValueError, TypeError):
                        pass
        if col_values:
            numeric_cols.append(
                {
                    "index": col_idx,
                    "name": headers[col_idx]
                    if col_idx < len(headers)
                    else f"Column {col_idx + 1}",
                    "values": col_values,
                }
            )

    stats = {"row_count": len(rows), "col_count": len(headers), "columns": {}}

    for col in numeric_cols:
        vals = col["values"]
        if vals:
            stats["columns"][col["name"]] = {
                "count": len(vals),
                "sum": round(sum(vals), 2),
                "average": round(sum(vals) / len(vals), 2),
                "min": round(min(vals), 2),
                "max": round(max(vals), 2),
                "range": round(max(vals) - min(vals), 2),
            }

    return stats


def find_trends(data: List[List[Any]]) -> dict:
    """Identify trends in numeric columns."""
    if not data or len(data) < 3:
        return {"error": "Need at least 3 rows to detect trends"}

    headers = data[0] if len(data) > 0 else []
    rows = data[1:] if len(data) > 1 else []

    trends = {}
    for col_idx in range(len(headers)):
        values = []
        for row in rows:
            if col_idx < len(row):
                val = row[col_idx]
                if val is not None and val != "":
                    try:
                        num = float(
                            str(val).replace(",", "").replace("₹", "").replace("$", "")
                        )
                        values.append(num)
                    except (ValueError, TypeError):
                        pass

        if len(values) >= 3:
            first_half = values[: len(values) // 2]
            second_half = values[len(values) // 2 :]
            avg_first = sum(first_half) / len(first_half)
            avg_second = sum(second_half) / len(second_half)

            if avg_second > avg_first * 1.1:
                trend = "increasing"
            elif avg_second < avg_first * 0.9:
                trend = "decreasing"
            else:
                trend = "stable"

            trends[
                headers[col_idx] if col_idx < len(headers) else f"Column {col_idx + 1}"
            ] = {
                "trend": trend,
                "change_pct": round(((avg_second - avg_first) / avg_first) * 100, 1)
                if avg_first
                else 0,
            }

    return {"trends": trends} if trends else {"message": "No clear trends detected"}


def detect_outliers(data: List[List[Any]], threshold: float = 2.0) -> dict:
    """Detect outliers using standard deviation method."""
    if not data or len(data) < 4:
        return {"error": "Need at least 4 rows to detect outliers"}

    headers = data[0] if len(data) > 0 else []
    rows = data[1:] if len(data) > 1 else []

    outliers = {}
    for col_idx in range(len(headers)):
        values = []
        for row_idx, row in enumerate(rows):
            if col_idx < len(row):
                val = row[col_idx]
                if val is not None and val != "":
                    try:
                        num = float(
                            str(val).replace(",", "").replace("₹", "").replace("$", "")
                        )
                        values.append((row_idx + 2, num))
                    except (ValueError, TypeError):
                        pass

        if len(values) >= 4:
            nums = [v[1] for v in values]
            mean = sum(nums) / len(nums)
            variance = sum((x - mean) ** 2 for x in nums) / len(nums)
            std_dev = variance**0.5

            col_outliers = []
            for row_num, val in values:
                if abs(val - mean) > threshold * std_dev:
                    col_outliers.append({"row": row_num, "value": val})

            if col_outliers:
                outliers[
                    headers[col_idx]
                    if col_idx < len(headers)
                    else f"Column {col_idx + 1}"
                ] = col_outliers

    return {"outliers": outliers} if outliers else {"message": "No outliers detected"}


def analyze_distribution(data: List[List[Any]]) -> dict:
    """Analyze data distribution and patterns."""
    if not data:
        return {"error": "No data"}

    headers = data[0] if len(data) > 0 else []
    rows = data[1:] if len(data) > 1 else []

    distribution = {}
    for col_idx in range(len(headers)):
        values = [
            row[col_idx]
            for row in rows
            if col_idx < len(row) and row[col_idx] is not None and row[col_idx] != ""
        ]

        if not values:
            continue

        value_counts = Counter([str(v) for v in values])
        top_values = value_counts.most_common(5)

        col_name = (
            headers[col_idx] if col_idx < len(headers) else f"Column {col_idx + 1}"
        )

        is_numeric = all(
            isinstance(v, (int, float))
            or (
                isinstance(v, str)
                and v.replace(",", "").replace(".", "").replace("-", "").isdigit()
            )
            for v in values
            if v is not None
        )

        distribution[col_name] = {
            "unique_count": len(set(str(v) for v in values)),
            "total_count": len(values),
            "top_values": [{"value": v, "count": c} for v, c in top_values],
            "is_numeric": is_numeric,
        }

    return {"distribution": distribution}


def build_table_schema(data: List[List[Any]]) -> dict:
    """Deterministic schema/context agent output for planner and intent resolution."""
    if not data or not data[0]:
        return {"error": "No data", "columns": []}

    headers = [str(h).strip() if h is not None else "" for h in data[0]]
    rows = data[1:] if len(data) > 1 else []
    row_count = len(rows)

    schema_columns = []
    for i, header in enumerate(headers):
        values = [r[i] if i < len(r) else None for r in rows]
        non_empty = [v for v in values if v not in (None, "")]
        null_count = row_count - len(non_empty)
        null_pct = (null_count / row_count * 100.0) if row_count > 0 else 0.0

        numeric_hits = 0
        date_hits = 0
        text_hits = 0
        for v in non_empty[:100]:
            s = str(v).strip()
            cleaned = re.sub(r"[^0-9.\-]", "", s)
            if cleaned and re.fullmatch(r"-?\d+(\.\d+)?", cleaned):
                numeric_hits += 1
            elif re.fullmatch(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", s) or re.fullmatch(
                r"\d{4}[/-]\d{1,2}[/-]\d{1,2}", s
            ):
                date_hits += 1
            else:
                text_hits += 1

        sample_size = max(len(non_empty[:100]), 1)
        numeric_conf = round(numeric_hits / sample_size, 3)
        date_conf = round(date_hits / sample_size, 3)
        if numeric_conf >= 0.6:
            inferred_type = "number"
        elif date_conf >= 0.5:
            inferred_type = "date"
        else:
            inferred_type = "text"

        schema_columns.append(
            {
                "index": i,
                "name": header or f"Column {i + 1}",
                "inferred_type": inferred_type,
                "numeric_confidence": numeric_conf,
                "date_confidence": date_conf,
                "null_pct": round(null_pct, 2),
                "sample_values": [str(v) for v in non_empty[:3]],
            }
        )

    return {
        "total_rows": row_count,
        "total_columns": len(headers),
        "columns": schema_columns,
    }


def profile_data(data: List[List[Any]]) -> str:
    """Deep data profiling for the planner. Returns a human-readable summary with column intelligence."""
    if not data or len(data) < 2:
        return "No data to profile."

    headers = [
        str(h).strip() if h is not None else f"Col{i}" for i, h in enumerate(data[0])
    ]
    rows = data[1:]
    total_rows = len(rows)
    num_cols = len(headers)

    issues = []
    highlights = []
    column_details = []

    # 1. Duplicate detection
    seen = set()
    dup_count = 0
    for row in rows:
        key = tuple(str(v) for v in row)
        if key in seen:
            dup_count += 1
        seen.add(key)
    if dup_count > 0:
        pct = round(dup_count / total_rows * 100, 1)
        issues.append(f"{dup_count} duplicate rows ({pct}%)")

    # 2. Per-column profiling
    for i, header in enumerate(headers):
        values = [r[i] if i < len(r) else None for r in rows]
        non_empty = [v for v in values if v not in (None, "")]
        null_count = total_rows - len(non_empty)
        null_pct = round(null_count / total_rows * 100, 1) if total_rows > 0 else 0

        # Unique value count
        unique_vals = list(set(str(v) for v in non_empty))
        unique_count = len(unique_vals)

        # Sample values (up to 5 unique non-empty)
        sample_vals = unique_vals[:5]

        # Null check
        if null_pct > 20:
            issues.append(f"'{header}': {null_pct}% empty cells")
        elif null_pct > 0:
            highlights.append(f"'{header}': {null_pct}% empty")

        # Type consistency
        num_count = 0
        text_count = 0
        for v in non_empty[:50]:
            cleaned = re.sub(r"[^0-9.\-]", "", str(v))
            if cleaned and re.fullmatch(r"-?\d+(\.\d+)?", cleaned):
                num_count += 1
            else:
                text_count += 1

        detected_type = "text"
        if num_count > len(non_empty) * 0.8 and len(non_empty) > 0:
            detected_type = "numeric"
        elif num_count > 0 and text_count > 0 and len(non_empty) > 2:
            detected_type = "mixed"
            num_pct = round(num_count / len(non_empty) * 100, 1)
            issues.append(
                f"'{header}': mixed types ({num_pct}% numeric, {round(100 - num_pct, 1)}% text)"
            )

        # Numeric statistics (min, max, mean, median)
        numeric_stats = ""
        if detected_type == "numeric" and num_count > 3:
            nums = []
            for v in non_empty[:200]:
                try:
                    nums.append(
                        float(
                            str(v)
                            .replace(",", "")
                            .replace("$", "")
                            .replace("₹", "")
                            .replace("%", "")
                            .strip()
                        )
                    )
                except:
                    pass
            if len(nums) > 3:
                nums_sorted = sorted(nums)
                min_val = nums_sorted[0]
                max_val = nums_sorted[-1]
                mean_val = sum(nums) / len(nums)
                mid = len(nums_sorted) // 2
                median_val = (
                    nums_sorted[mid]
                    if len(nums_sorted) % 2 == 1
                    else (nums_sorted[mid - 1] + nums_sorted[mid]) / 2
                )
                numeric_stats = f"min={min_val}, max={max_val}, mean={round(mean_val, 2)}, median={round(median_val, 2)}"

                # Outlier detection (values > 3 std dev from mean)
                variance = sum((x - mean_val) ** 2 for x in nums) / len(nums)
                std_dev = variance**0.5
                if std_dev > 0:
                    outlier_count = sum(1 for x in nums if abs(x - mean_val) > 3 * std_dev)
                    if outlier_count > 0:
                        highlights.append(
                            f"'{header}': {outlier_count} potential outliers (>3σ)"
                        )

        # Build column detail line
        detail = f"  {header}: type={detected_type}, unique={unique_count}/{total_rows}"
        if sample_vals:
            detail += f", samples={sample_vals}"
        if numeric_stats:
            detail += f" [{numeric_stats}]"
        column_details.append(detail)

    # 3. Empty dataset check
    if total_rows == 0:
        issues.append("Dataset has no data rows (only headers)")

    # Build summary
    summary_parts = [f"DATA PROFILE: {total_rows} rows × {num_cols} columns"]

    if issues:
        summary_parts.append("ISSUES FOUND:")
        for issue in issues[:8]:
            summary_parts.append(f"  ⚠ {issue}")

    if highlights:
        summary_parts.append("NOTES:")
        for h in highlights[:5]:
            summary_parts.append(f"  ℹ {h}")

    if column_details:
        summary_parts.append("COLUMN DETAILS:")
        for detail in column_details:
            summary_parts.append(detail)

    if not issues and not highlights:
        summary_parts.append("Data looks clean — no major issues detected.")

    return "\n".join(summary_parts)


def compute_large_data_stats(data: List[List[Any]]) -> dict:
    """Compute comprehensive statistics for large datasets efficiently."""
    try:
        import pandas as pd
        import numpy as np

        if not data or not data[0]:
            return {"error": "No data to analyze"}

        headers = data[0] if len(data) > 0 else []
        rows = data[1:] if len(data) > 1 else []
        total_rows = len(rows)

        df = pd.DataFrame(rows, columns=headers[: len(rows[0])] if rows else headers)

        results = {
            "overview": {
                "total_rows": total_rows,
                "total_columns": len(headers),
                "headers": list(headers),
            },
            "columns": {},
        }

        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                col_data = df[col].dropna()
                if len(col_data) > 0:
                    results["columns"][col] = {
                        "count": int(col_data.count()),
                        "sum": round(float(col_data.sum()), 2),
                        "average": round(float(col_data.mean()), 2),
                        "min": round(float(col_data.min()), 2),
                        "max": round(float(col_data.max()), 2),
                        "range": round(float(col_data.max() - col_data.min()), 2),
                        "median": round(float(col_data.median()), 2),
                        "std_dev": round(float(col_data.std()), 2)
                        if len(col_data) > 1
                        else 0,
                        "q1": round(float(col_data.quantile(0.25)), 2),
                        "q3": round(float(col_data.quantile(0.75)), 2),
                        "iqr": round(
                            float(col_data.quantile(0.75) - col_data.quantile(0.25)), 2
                        ),
                    }
            else:
                results["columns"][col] = {
                    "count": int(df[col].count()),
                    "unique_count": int(df[col].nunique()),
                    "top_values": df[col].value_counts().head(5).to_dict(),
                    "is_numeric": False,
                }

        return results

    except ImportError:
        # Fallback to basic implementation without pandas
        return _compute_large_data_stats_basic(data)


def _compute_large_data_stats_basic(data: List[List[Any]]) -> dict:
    """Fallback basic stats without pandas."""
    if not data or not data[0]:
        return {"error": "No data to analyze"}

    headers = data[0]
    rows = data[1:]
    total_rows = len(rows)

    results = {
        "overview": {
            "total_rows": total_rows,
            "total_columns": len(headers),
            "headers": list(headers),
        },
        "columns": {},
    }

    for col_idx in range(len(headers)):
        col_name = headers[col_idx]
        col_values = []
        for row in rows:
            if col_idx < len(row) and row[col_idx] not in [None, ""]:
                try:
                    col_values.append(
                        float(
                            str(row[col_idx])
                            .replace(",", "")
                            .replace("₹", "")
                            .replace("$", "")
                            .strip()
                        )
                    )
                except:
                    pass

        if col_values:
            n = len(col_values)
            sorted_vals = sorted(col_values)
            mid = n // 2
            results["columns"][col_name] = {
                "count": n,
                "sum": round(sum(col_values), 2),
                "average": round(sum(col_values) / n, 2),
                "min": round(min(col_values), 2),
                "max": round(max(col_values), 2),
                "range": round(max(col_values) - min(col_values), 2),
                "median": round(
                    sorted_vals[mid]
                    if n % 2 == 1
                    else (sorted_vals[mid - 1] + sorted_vals[mid]) / 2,
                    2,
                ),
                "std_dev": round(
                    (sum((x - sum(col_values) / n) ** 2 for x in col_values) / n)
                    ** 0.5,
                    2,
                )
                if n > 1
                else 0,
            }

    return results


@app.post("/api/analyze-large")
async def analyze_large_data(req: AnalyzeRequest):
    """Analyze large datasets - stats computed locally, only the summary goes to the LLM."""
    stats = compute_large_data_stats(req.data)

    user = f"""Analyze briefly (3-5 points max):

Rows: {stats.get("overview", {}).get("total_rows", "N/A")}
Columns: {stats.get("overview", {}).get("total_columns", "N/A")}
Headers: {stats.get("overview", {}).get("headers", [])}

Stats: {json.dumps(stats.get("columns", {}), indent=None)}

Answer: {req.question}

Keep it short. No asterisks (*)."""

    output = await call_llm(
        req.provider,
        req.model,
        req.api_key,
        "You are an expert at analyzing large datasets. You compute statistics "
        "efficiently without processing every row through the LLM. You focus on "
        "aggregated insights and patterns rather than individual data points.",
        user,
        max_tokens=1500,
    )

    return {"analysis": output, "statistics": stats}


def _generate_sample_data(data_type: str, count: int) -> tuple:
    """Generate realistic sample data. Returns (headers, rows)."""
    import random

    random.seed(42)  # Reproducible data
    count = min(max(count, 1), 10000)

    first_names = [
        "James",
        "Mary",
        "John",
        "Patricia",
        "Robert",
        "Jennifer",
        "Michael",
        "Linda",
        "William",
        "Elizabeth",
        "David",
        "Barbara",
        "Richard",
        "Susan",
        "Joseph",
        "Jessica",
        "Thomas",
        "Sarah",
        "Charles",
        "Karen",
        "Raj",
        "Priya",
        "Amit",
        "Ananya",
        "Vikram",
        "Deepa",
        "Arjun",
        "Meera",
        "Rohan",
        "Kavita",
        "Wei",
        "Xiao",
        "Ming",
        "Fang",
        "Jing",
        "Hiroshi",
        "Yuki",
        "Kenji",
        "Sakura",
        "Akiko",
    ]
    last_names = [
        "Smith",
        "Johnson",
        "Williams",
        "Brown",
        "Jones",
        "Garcia",
        "Miller",
        "Davis",
        "Rodriguez",
        "Martinez",
        "Hernandez",
        "Lopez",
        "Gonzalez",
        "Wilson",
        "Anderson",
        "Thomas",
        "Taylor",
        "Moore",
        "Jackson",
        "Martin",
        "Sharma",
        "Patel",
        "Singh",
        "Kumar",
        "Gupta",
        "Verma",
        "Reddy",
        "Joshi",
        "Rao",
        "Nair",
        "Wang",
        "Li",
        "Zhang",
        "Liu",
        "Chen",
        "Tanaka",
        "Suzuki",
        "Yamamoto",
        "Nakamura",
        "Watanabe",
    ]
    departments = [
        "Engineering",
        "Marketing",
        "Sales",
        "HR",
        "Finance",
        "Operations",
        "IT",
        "Legal",
        "Product",
        "Design",
    ]
    cities = [
        "New York",
        "London",
        "Tokyo",
        "Mumbai",
        "Sydney",
        "Berlin",
        "Paris",
        "Toronto",
        "Singapore",
        "Dubai",
    ]
    products = [
        "Widget A",
        "Widget B",
        "Gadget X",
        "Gadget Y",
        "Tool Pro",
        "Tool Lite",
        "Service Basic",
        "Service Premium",
        "Pack Standard",
        "Pack Enterprise",
    ]
    categories = ["Electronics", "Fashion", "Accessories", "Home & Garden", "Sports"]
    blood_params = [
        "Hemoglobin",
        "WBC",
        "RBC",
        "Platelets",
        "Glucose",
        "Cholesterol",
        "Creatinine",
        "Urea",
        "ALT",
        "AST",
        "Bilirubin",
        "Albumin",
        "Calcium",
        "Iron",
        "Vitamin D",
    ]
    subjects = [
        "Math",
        "Physics",
        "Chemistry",
        "Biology",
        "English",
        "History",
        "Geography",
        "Computer Science",
    ]

    if data_type == "blood_report":
        headers = [
            "Patient Name",
            "Age",
            "Gender",
            "Test Parameter",
            "Result",
            "Unit",
            "Reference Range",
            "Status",
        ]
        rows = []
        for i in range(count):
            name = f"{random.choice(first_names)} {random.choice(last_names)}"
            age = random.randint(18, 85)
            gender = random.choice(["Male", "Female"])
            param = random.choice(blood_params)
            units_map = {
                "Hemoglobin": "g/dL",
                "WBC": "cells/µL",
                "RBC": "million/µL",
                "Platelets": "cells/µL",
                "Glucose": "mg/dL",
                "Cholesterol": "mg/dL",
                "Creatinine": "mg/dL",
                "Urea": "mg/dL",
                "ALT": "U/L",
                "AST": "U/L",
                "Bilirubin": "mg/dL",
                "Albumin": "g/dL",
                "Calcium": "mg/dL",
                "Iron": "µg/dL",
                "Vitamin D": "ng/mL",
            }
            ranges_map = {
                "Hemoglobin": (12.0, 17.5),
                "WBC": (4000, 11000),
                "RBC": (4.2, 5.8),
                "Platelets": (150000, 400000),
                "Glucose": (70, 110),
                "Cholesterol": (0, 200),
                "Creatinine": (0.6, 1.2),
                "Urea": (7, 20),
                "ALT": (7, 56),
                "AST": (10, 40),
                "Bilirubin": (0.1, 1.2),
                "Albumin": (3.5, 5.0),
                "Calcium": (8.5, 10.5),
                "Iron": (60, 170),
                "Vitamin D": (30, 100),
            }
            unit = units_map.get(param, "")
            ref_low, ref_high = ranges_map.get(param, (0, 100))
            result = round(random.uniform(ref_low * 0.7, ref_high * 1.3), 1)
            status = (
                "Normal"
                if ref_low <= result <= ref_high
                else ("Low" if result < ref_low else "High")
            )
            rows.append(
                [
                    name,
                    age,
                    gender,
                    param,
                    result,
                    unit,
                    f"{ref_low}-{ref_high}",
                    status,
                ]
            )

    elif data_type == "employee":
        headers = [
            "Employee ID",
            "Name",
            "Department",
            "Designation",
            "Salary",
            "Join Date",
            "City",
            "Performance Rating",
        ]
        rows = []
        for i in range(count):
            emp_id = f"EMP{1000 + i}"
            name = f"{random.choice(first_names)} {random.choice(last_names)}"
            dept = random.choice(departments)
            designations = {
                "Engineering": [
                    "Software Engineer",
                    "Senior Engineer",
                    "Tech Lead",
                    "Architect",
                ],
                "Marketing": ["Marketing Exec", "Campaign Manager", "Brand Lead"],
                "Sales": ["Sales Rep", "Account Manager", "Sales Director"],
                "HR": ["HR Exec", "HR Manager", "HR Director"],
                "Finance": ["Analyst", "Sr Analyst", "Finance Manager"],
                "Operations": ["Ops Exec", "Ops Manager", "Ops Director"],
                "IT": ["IT Support", "Sys Admin", "IT Manager"],
                "Legal": ["Legal Exec", "Legal Counsel", "Legal Director"],
                "Product": ["Product Analyst", "Product Manager", "Product Director"],
                "Design": ["UI Designer", "UX Designer", "Design Lead"],
            }
            designation = random.choice(designations.get(dept, ["Associate"]))
            salary = random.randint(30000, 150000)
            join_date = f"{random.randint(2018, 2024)}-{random.randint(1, 12):02d}-{random.randint(1, 28):02d}"
            city = random.choice(cities)
            rating = round(random.uniform(2.0, 5.0), 1)
            rows.append(
                [emp_id, name, dept, designation, salary, join_date, city, rating]
            )

    elif data_type == "sales":
        headers = [
            "Order ID",
            "Date",
            "Product",
            "Category",
            "Quantity",
            "Unit Price",
            "Total",
            "Region",
            "Customer",
        ]
        rows = []
        for i in range(count):
            order_id = f"ORD{10000 + i}"
            date = f"2024-{random.randint(1, 12):02d}-{random.randint(1, 28):02d}"
            product = random.choice(products)
            category = random.choice(categories)
            qty = random.randint(1, 50)
            unit_price = round(random.uniform(10, 500), 2)
            total = round(qty * unit_price, 2)
            region = random.choice(["North", "South", "East", "West", "Central"])
            customer = f"{random.choice(first_names)} {random.choice(last_names)}"
            rows.append(
                [
                    order_id,
                    date,
                    product,
                    category,
                    qty,
                    unit_price,
                    total,
                    region,
                    customer,
                ]
            )

    elif data_type == "student":
        headers = [
            "Student ID",
            "Name",
            "Grade",
            "Subject",
            "Marks",
            "Max Marks",
            "Percentage",
            "Result",
        ]
        rows = []
        for i in range(count):
            sid = f"STU{1000 + i}"
            name = f"{random.choice(first_names)} {random.choice(last_names)}"
            grade = random.choice(["9th", "10th", "11th", "12th"])
            subject = random.choice(subjects)
            max_marks = 100
            marks = random.randint(20, 100)
            pct = round((marks / max_marks) * 100, 1)
            result = "Pass" if marks >= 35 else "Fail"
            rows.append([sid, name, grade, subject, marks, max_marks, pct, result])

    else:
        # Generic fallback
        headers = ["ID", "Name", "Value", "Category", "Date"]
        rows = []
        for i in range(count):
            rows.append(
                [
                    i + 1,
                    f"Item {i + 1}",
                    round(random.uniform(10, 1000), 2),
                    random.choice(["A", "B", "C"]),
                    f"2024-{random.randint(1, 12):02d}-{random.randint(1, 28):02d}",
                ]
            )

    return headers, rows


@app.post("/api/generate-data")
async def generate_data(req: GenerateDataRequest):
    """Generate sample data tables (employee, sales, blood report, student, etc.)."""
    try:
        headers, rows = _generate_sample_data(req.data_type, req.count)
        data = [headers] + rows
        return {
            "data": data,
            "headers": headers,
            "row_count": len(rows),
            "col_count": len(headers),
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/chat")
async def chat(req: ChatRequest):
    """Plan an Excel request.

    One LLM call produces reasoning, a clarification verdict, and the plan.
    Previously three sequential CrewAI agents produced the same three outputs.
    """
    context = build_planner_context(req)
    want_reasoning = req.enable_reasoning is not False

    user = TASK_DESCRIPTION.format(
        message=req.message,
        context_info=context["context_info"],
        selected_info=context["selected_info"],
        workbook_metadata_info=context["workbook_metadata_info"],
        headers_context=context["headers_context"] or "No column headers available",
        semantic_matches=context["semantic_matches"] or "No semantic matches found",
        numeric_columns=context["numeric_columns"] or "No numeric columns detected",
        existing_formulas=context["existing_formulas"] or "None - selection has no formulas",
        number_formats=context["number_formats"] or "None - all cells use General format",
        schema_context=context["schema_context"] or "No schema available",
        data_profile=context["data_profile"] or "No data to profile",
        memory_info=context["memory_info"],
        history_info=context["history_info"],
        clarification_section=(
            CLARIFICATION_PROMPT if context["selected_range_address"] else ""
        ),
        reasoning_instruction=(
            "Include a \"reasoning\" field: 1-2 plain sentences on what you will do and why, "
            "using \"I will...\". No markdown, no asterisks, no manual Excel steps."
            if want_reasoning
            else 'Set "reasoning" to an empty string.'
        ),
    )

    try:
        raw = await call_llm(
            req.provider,
            req.model,
            req.api_key,
            SYSTEM_PROMPT,
            user,
            max_tokens=8192,
        )
    except RuntimeError as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    parsed = extract_json_object(raw) or {}

    reasoning = str(parsed.get("reasoning") or "").replace("*", "").strip()

    async def generate():
        if reasoning:
            yield _sse({"type": "reasoning", "content": reasoning})

        if parsed.get("needs_clarification"):
            yield _sse(
                {
                    "type": "clarification",
                    "content": parsed.get(
                        "clarification_question",
                        "Could you be more specific about what you'd like me to do?",
                    ),
                }
            )
            return

        plan = parsed.get("plan") or []
        if not isinstance(plan, list):
            plan = []

        yield _sse(
            {
                "type": "plan",
                "plan": plan,
                "content": parsed.get("response") or "",
            }
        )

    return StreamingResponse(generate(), media_type="text/event-stream")


def _sse(payload: dict) -> str:
    return f"data: {json.dumps(payload)}\n\n"


def extract_json_object(raw: str) -> Optional[dict]:
    """Pull the first balanced JSON object out of an LLM response."""
    start = raw.find("{")
    if start < 0:
        return None
    depth = 0
    in_string = False
    escaped = False
    for i in range(start, len(raw)):
        ch = raw[i]
        if in_string:
            if escaped:
                escaped = False
            elif ch == "\\":
                escaped = True
            elif ch == '"':
                in_string = False
            continue
        if ch == '"':
            in_string = True
        elif ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                try:
                    return json.loads(raw[start : i + 1])
                except json.JSONDecodeError:
                    return None
    return None


def _selected_range_payload(req: ChatRequest) -> Optional[dict]:
    sr = req.selected_range
    if isinstance(sr, str):
        try:
            sr = json.loads(sr)
        except json.JSONDecodeError:
            logger.warning("selected_range is not valid JSON; planner gets no range context")
            return None
    return sr if isinstance(sr, dict) else None


def build_planner_context(req: ChatRequest) -> dict:
    """Assemble every planner input. Returns strings, never raises."""
    ctx = {
        "headers_context": "",
        "schema_context": "",
        "data_profile": "",
        "semantic_matches": "",
        "numeric_columns": "",
        "existing_formulas": "",
        "number_formats": "",
        "selected_range_address": "",
        "context_info": "",
        "selected_info": "",
        "memory_info": "",
        "workbook_metadata_info": "",
        "history_info": "",
    }

    sr = _selected_range_payload(req)
    if sr:
        address = str(sr.get("address") or "")
        sheet_name = str(sr.get("sheetName") or "")
        ctx["selected_range_address"] = address
        ctx["selected_info"] = (
            f"Selected range: {address} on sheet '{sheet_name}' "
            f"({sr.get('rowCount', '?')} rows x {sr.get('columnCount', '?')} cols)"
            if address
            else ""
        )

        values = sr.get("values") or []
        if values and values[0]:
            header_names = [str(h).strip() for h in values[0] if h]
            if header_names:
                ctx["headers_context"] = (
                    f"Sheet '{sheet_name}' columns ({len(header_names)}): "
                    f"{', '.join(header_names)}\nData range: {address}"
                )
            try:
                ctx["schema_context"] = json.dumps(
                    build_table_schema(values), ensure_ascii=False
                )
                ctx["data_profile"] = profile_data(values)
                ctx["semantic_matches"] = resolve_semantic_matches(
                    req.message, header_names
                )
                ctx["numeric_columns"] = _numeric_column_list(values, header_names)
                ctx["existing_formulas"] = _format_formula_context(sr)
                ctx["number_formats"] = _format_number_formats(sr)
            except Exception:
                logger.exception("failed to profile selected range")

    if req.workbook_context:
        ctx["context_info"] = f"Workbook context: {req.workbook_context}"
    if req.memory_context:
        ctx["memory_info"] = f"Memory context:\n{req.memory_context}"

    if req.workbook_metadata:
        try:
            wm = req.workbook_metadata
            if isinstance(wm, str):
                wm = json.loads(wm)
            if isinstance(wm, dict):
                ctx["workbook_metadata_info"] = _format_workbook_metadata(wm)
        except Exception:
            logger.exception("failed to parse workbook_metadata")

    if req.conversation_history:
        lines = []
        for msg in req.conversation_history[-10:]:
            role = msg.get("role", "user")
            content = str(msg.get("content", ""))[:600]
            if role in ("user", "assistant"):
                lines.append(f"{role.capitalize()}: {content}")
        if lines:
            ctx["history_info"] = "RECENT CONVERSATION:\n" + "\n".join(lines)

    return ctx


def _format_workbook_metadata(wm: dict) -> str:
    sheets = wm.get("sheets", []) or []
    lines = [
        f"WORKBOOK STRUCTURE ({len(sheets)} sheets, active: {wm.get('activeSheet', '')}):"
    ]
    for sh in sheets:
        hdrs = ", ".join(sh.get("headers", [])[:8])
        lines.append(
            f"  {sh.get('name')}: {sh.get('rowCount', '?')} rows x "
            f"{sh.get('columnCount', '?')} cols - headers: [{hdrs}]"
        )

    for t in wm.get("tables", []) or []:
        cols = ", ".join(t.get("columns", [])[:6])
        lines.append(
            f"TABLE: {t.get('name')} on {t.get('sheetName')}: "
            f"{t.get('rowCount', '?')} rows, columns: {cols}"
        )

    for n in wm.get("namedRanges", []) or []:
        lines.append(f"NAMED RANGE: {n.get('name')} -> {n.get('address', 'N/A')}")

    prescan = wm.get("prescan", []) or []
    if prescan:
        lines.append("CROSS-SHEET PRESCAN (data from referenced sheets):")
        for ps in prescan:
            hdrs = ", ".join(ps.get("headers", [])[:8])
            lines.append(
                f"  {ps.get('sheetName')} ({ps.get('totalRows', '?')} rows): headers=[{hdrs}]"
            )
            for row in ps.get("sampleRows", []) or []:
                lines.append(f"    sample: {json.dumps(row, ensure_ascii=False)[:200]}")

    return "\n".join(lines)


@app.post("/api/execute")
async def execute_plan(req: ExecuteRequest):
    """Summarise execution results."""
    results_str = json.dumps(req.results, indent=2)
    plan_str = (
        "\n".join(
            f"{i + 1}. {st.get('action', 'unknown')}: {st.get('description', '')}"
            for i, st in enumerate(req.plan)
        )
        or "No plan available"
    )

    try:
        output = await call_llm(
            req.provider,
            req.model,
            req.api_key,
            "You verify that Excel operations completed successfully and summarize "
            "what was done. Be specific about cell ranges and what changed.",
            VALIDATOR_TASK.format(
                message=req.message, plan_info=plan_str, results_str=results_str
            ),
            max_tokens=1200,
        )
    except RuntimeError as exc:
        # Summarisation is cosmetic: never fail the user's actual work over it
        logger.warning("validator LLM failed, using raw summary: %s", exc)
        failed = [r for r in req.results if not r.get("success")]
        output = (
            f"Completed {len(req.results) - len(failed)} of {len(req.results)} steps."
            if failed
            else f"Completed all {len(req.results)} steps."
        )
        if failed:
            output += " Failed: " + "; ".join(
                f"{r.get('action')}: {str(r.get('output', ''))[:120]}" for r in failed
            )

    return {"content": output}


@app.post("/api/analyze")
async def analyze_data(req: AnalyzeRequest):
    """Compute stats/trends/outliers locally, send only the summary to the LLM."""
    stats = compute_statistics(req.data)
    trends = find_trends(req.data)
    outliers = detect_outliers(req.data)
    distribution = analyze_distribution(req.data)

    user = f"""Analyze briefly (3-5 points max):

Headers: {req.headers}
Data: {len(req.data)} rows

Stats: {json.dumps(stats, indent=None)}

Answer: {req.question}

Keep it short, no asterisks (*).

At the end, mention which 2 columns would make a good chart (e.g., "Team vs Total Titles")."""

    try:
        output = await call_llm(
            req.provider,
            req.model,
            req.api_key,
            "You are a concise data analyst. Give brief, point-wise insights.",
            user,
            max_tokens=1500,
        )
    except RuntimeError as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    return {
        "analysis": output,
        "statistics": stats,
        "trends": trends.get("trends", {}),
        "outliers": outliers.get("outliers", {}),
        "distribution": distribution.get("distribution", {}),
    }


@app.post("/api/build-schema")
async def build_schema_endpoint(req: AnalyzeRequest):
    try:
        return {"schema": build_table_schema(req.data)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


class ReflectRequest(BaseModel):
    message: str
    plan: list = []
    results: list = []
    provider: str = "openai"
    model: str = ""
    api_key: str
    selected_range: Optional[str] = None


@app.post("/api/reflect")
async def reflect(req: ReflectRequest):
    """Build a minimal recovery plan for failed steps."""
    failed_results = [r for r in req.results if not r.get("success")]
    if not failed_results:
        return {"recovery": []}

    headers_context = "No headers available"
    schema_context = "No schema available"
    sr = _selected_range_payload(req)
    if sr:
        values = sr.get("values") or []
        if values and values[0]:
            header_names = [str(h).strip() for h in values[0] if h]
            if header_names:
                headers_context = f"Column headers: {', '.join(header_names)}"
            try:
                schema_context = json.dumps(
                    build_table_schema(values), ensure_ascii=False
                )
            except Exception:
                logger.exception("failed to profile selected range for reflection")

    plan_str = "\n".join(
        f"{i + 1}. {st.get('action', 'unknown')}: {st.get('description', '')}"
        for i, st in enumerate(req.plan)
    )
    results_str = "\n".join(
        f"{i + 1}. {r.get('action', 'unknown')}: "
        f"{'ok' if r.get('success') else 'FAILED'} {r.get('output', '')}"
        for i, r in enumerate(req.results)
    )

    try:
        output = await call_llm(
            req.provider,
            req.model,
            req.api_key,
            "You are a self-correction agent. You return ONLY a valid JSON array of "
            "recovery steps. If no recovery is possible, return [].",
            REFLECTION_PROMPT.format(
                message=req.message,
                plan=plan_str,
                results=results_str,
                headers_context=headers_context,
                schema_context=schema_context,
            ),
            max_tokens=2048,
        )
    except RuntimeError as exc:
        logger.warning("reflector LLM failed: %s", exc)
        return {"recovery": []}

    recovery = extract_json_array(output)
    return {"recovery": recovery if isinstance(recovery, list) else []}


def extract_json_array(raw: str) -> Optional[list]:
    start = raw.find("[")
    if start < 0:
        return None
    depth = 0
    in_string = False
    escaped = False
    for i in range(start, len(raw)):
        ch = raw[i]
        if in_string:
            if escaped:
                escaped = False
            elif ch == "\\":
                escaped = True
            elif ch == '"':
                in_string = False
            continue
        if ch == '"':
            in_string = True
        elif ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                try:
                    return json.loads(raw[start : i + 1])
                except json.JSONDecodeError:
                    return None
    return None


@app.get("/api/health")
async def health():
    return {"status": "ok"}
