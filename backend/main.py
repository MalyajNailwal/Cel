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
import io
import base64
import random
import string
from collections import Counter
from datetime import datetime, timedelta

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
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


class ChartIntentRequest(BaseModel):
    data: List[List[Any]]
    question: str
    headers: Optional[List[str]] = None
    start_col_index: Optional[int] = 0


class VerifyChartExecutionRequest(BaseModel):
    request_message: str
    expected_charts: List[dict]
    created_charts: List[dict]


class GenerateDataRequest(BaseModel):
    data_type: str = "employee"
    count: int = 100


class ChatResponse(BaseModel):
    content: str
    plan: list
    reasoning: Optional[str] = None


REASONING_PROMPT = """You are an Excel reasoning agent. Your job is to understand what the user wants and explain your thinking in agentic terms.

IMPORTANT - RESPONSE STYLE:
- ALWAYS describe what YOU (the AI) will do, NOT what user should do manually
- NEVER say "Go to Home tab" or "click" or "select" — that's manual Excel work
- Use phrases like "I will...", "I'll apply...", "I'll create...", "I'll format..."
- Keep it conversational and clear

Example GOOD: "I will add black borders to the selected cells using apply_format"
Example BAD: "Go to Home tab, click Borders, choose Outside Borders"

Analyze and explain:
1. What user wants (1-2 sentences)
2. What data operations needed
3. How you'll approach it (agentic steps, not manual Excel steps)"""


def setup_env(provider: str, model: str, api_key: str) -> str:
    """Configure environment variables for CrewAI and return the model string to use."""
    keys_to_clear = [
        "OPENAI_API_KEY",
        "OPENAI_API_BASE",
        "OPENAI_MODEL_NAME",
        "ANTHROPIC_API_KEY",
        "ANTHROPIC_MODEL_NAME",
        "GOOGLE_API_KEY",
        "GOOGLE_MODEL_NAME",
        "OPENROUTER_API_KEY",
    ]
    for k in keys_to_clear:
        os.environ.pop(k, None)

    if provider == "openrouter":
        os.environ["OPENROUTER_API_KEY"] = api_key
        return "openrouter/" + model
    elif provider == "openai":
        os.environ["OPENAI_API_KEY"] = api_key
    elif provider == "anthropic":
        os.environ["ANTHROPIC_API_KEY"] = api_key
    elif provider == "google":
        os.environ["GOOGLE_API_KEY"] = api_key
    else:
        os.environ["OPENAI_API_KEY"] = api_key

    return model


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

COLUMN HEADERS (use these for charts!):
{headers_context}

TABLE SCHEMA (deterministic context agent output):
{schema_context}

{memory_info}

{history_info}

Reasoning analysis:
{reasoning_output}

Available actions: get_workbook_structure, get_selected_range, get_range, get_sheet_data,
set_values, set_formulas, apply_format, insert_rows, delete_rows, insert_columns,
delete_columns, add_worksheet, delete_worksheet, create_table, sort_range, auto_fill, create_chart,
conditional_format, find_replace, merge_cells, unmerge_cells

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

IMPORTANT - CONDITIONAL FORMATTING:
- When user says "highlight cells > 100", "color cells where value is negative", "red if less than 50" → use conditional_format
- Params: address, rule_type="cellValue", operator (GreaterThan/LessThan/EqualTo/Between/etc.), formula1 (the threshold value), fill_color (highlight color)
- Example: "highlight cells > 100 in yellow" → conditional_format with address, operator="GreaterThan", formula1="100", fill_color="#FFFF00"
- Example: "red if negative" → conditional_format with operator="LessThan", formula1="0", fill_color="#FF0000"
- Example: "green if between 50 and 100" → conditional_format with operator="Between", formula1="50", formula2="100", fill_color="#00FF00"

IMPORTANT - FIND AND REPLACE:
- When user says "replace X with Y", "change all 'old' to 'new'", "find and replace" → use find_replace
- Params: address, find_text, replace_text, match_case (optional, default false)
- Example: "replace 'USA' with 'United States'" → find_replace with find_text="USA", replace_text="United States"

IMPORTANT - MERGE/UNMERGE CELLS:
- When user says "merge these cells", "combine A1:D1" → use merge_cells with address
- When user says "unmerge", "split merged cells" → use unmerge_cells with address

For charts (create_chart), use these params:
- chart_type: "column", "bar", "line", "pie", "pie3D", "doughnut", "area", "scatter", "radar", "surface", "bubble"
- data_range: A1 notation range that contains the data (e.g., "A1:B10")
- sheet_name: name of the sheet
- title: chart title (optional)
- position: {{"left": 300, "top": 50, "width": 400, "height": 300}} (optional, pixels from top-left)

Respond with a JSON object ONLY:
{{
  "plan": [
    {{
      "action": "tool_name",
      "params": {{"key": "value"}},
      "description": "What this step does"
    }}
  ],
  "response": "Natural language summary for the user"
}}

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
14. NEVER return empty values or placeholder text in the values array - always generate real data"""

VALIDATOR_PROMPT = """You verify that Excel operations completed successfully and summarize what was done."""
VALIDATOR_TASK = """Original request: {message}

Planned approach:
{plan_info}

Execution results:
{results_str}

Provide a clear, friendly summary of what was accomplished. Be specific about cell ranges and what changed."""


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
    """Analyze large datasets - handles 100k+ rows efficiently."""
    try:
        from crewai import Agent, Task, Crew, Process

        model = setup_env(req.provider, req.model, req.api_key)

        stats = compute_large_data_stats(req.data)

        ANALYSIS_TASK = f"""Analyze briefly (3-5 points max):

Rows: {stats.get("overview", {}).get("total_rows", "N/A")}
Columns: {stats.get("overview", {}).get("total_columns", "N/A")}
Headers: {stats.get("overview", {}).get("headers", [])}

Stats: {json.dumps(stats.get("columns", {}), indent=None)}

Answer: {req.question}

Keep it short. No asterisks (*)."""

        analyst = Agent(
            role="Large Data Analyst",
            goal="Analyze large datasets efficiently and provide insights",
            backstory="""You are an expert at analyzing large datasets. 
            You compute statistics efficiently without trying to process every row through the LLM.
            You focus on aggregated insights and patterns rather than individual data points.""",
            verbose=False,
            allow_delegation=False,
            llm=model,
        )

        analysis_task = Task(
            description=ANALYSIS_TASK,
            expected_output="Clear insights from large dataset analysis.",
            agent=analyst,
        )

        crew = Crew(
            agents=[analyst],
            tasks=[analysis_task],
            process=Process.sequential,
            verbose=False,
        )

        result = crew.kickoff()
        output = str(result.raw) if hasattr(result, "raw") else str(result)

        return {
            "analysis": output,
            "statistics": stats,
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


class ChartRequest(BaseModel):
    data: List[List[Any]]
    chart_type: str = "line"
    title: str = "Chart"
    x_column: Optional[str] = None
    y_column: Optional[str] = None
    column_name: Optional[str] = None


def generate_single_chart(
    data: List[List[Any]],
    chart_type: str,
    title: str,
    column_name: Optional[str] = None,
) -> Optional[str]:
    """Generate a single chart and return base64 image."""
    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        if not data or len(data) < 2:
            return None

        headers = data[0]
        rows = data[1:]

        sample_size = min(1000, len(rows))
        step = max(1, len(rows) // sample_size)
        sampled_rows = rows[::step]

        fig, ax = plt.subplots(figsize=(10, 6))

        if chart_type == "pie" and column_name:
            try:
                col_idx = headers.index(column_name)
                values = [
                    row[col_idx]
                    for row in sampled_rows
                    if col_idx < len(row) and row[col_idx] is not None
                ]

                from collections import Counter

                counts = Counter([str(v) for v in values])
                labels = list(counts.keys())
                sizes = list(counts.values())

                colors = plt.cm.Set3.colors[: len(labels)]
                ax.pie(
                    sizes,
                    labels=labels,
                    colors=colors,
                    autopct="%1.1f%%",
                    startangle=90,
                    textprops={"fontsize": 9},
                )
                ax.set_title(title, fontsize=14, fontweight="bold", pad=20)
            except Exception as e:
                print(f"Pie chart error: {e}")
                return None

        elif chart_type == "bar":
            if column_name:
                try:
                    col_idx = headers.index(column_name)
                    values = [
                        row[col_idx]
                        for row in sampled_rows
                        if col_idx < len(row) and row[col_idx] is not None
                    ]

                    if all(
                        isinstance(v, (int, float))
                        or (
                            isinstance(v, str)
                            and v.replace(".", "").replace("-", "").isdigit()
                        )
                        for v in values
                        if v is not None
                    ):
                        numeric_values = []
                        for v in values:
                            try:
                                numeric_values.append(
                                    float(
                                        str(v)
                                        .replace(",", "")
                                        .replace("₹", "")
                                        .replace("$", "")
                                        .strip()
                                    )
                                )
                            except:
                                numeric_values.append(0)

                        ax.bar(
                            range(len(numeric_values)),
                            numeric_values,
                            color="#217346",
                            alpha=0.8,
                        )
                        ax.set_title(title, fontsize=14, fontweight="bold")
                        ax.set_xlabel("Index", fontsize=10)
                        ax.set_ylabel(column_name, fontsize=10)
                        ax.grid(True, alpha=0.3, axis="y")
                    else:
                        from collections import Counter

                        counts = Counter([str(v) for v in values])
                        labels = list(counts.keys())[:15]
                        sizes = [counts[l] for l in labels]

                        ax.bar(labels, sizes, color="#217346", alpha=0.8)
                        ax.set_title(title, fontsize=14, fontweight="bold")
                        ax.set_xlabel(column_name, fontsize=10)
                        ax.set_ylabel("Count", fontsize=10)
                        ax.tick_params(axis="x", rotation=45)
                        ax.grid(True, alpha=0.3, axis="y")
                except Exception as e:
                    print(f"Bar chart error: {e}")
                    return None
            else:
                if len(headers) >= 2:
                    y_data = []
                    for row in sampled_rows:
                        if len(row) > 1:
                            try:
                                y_data.append(
                                    float(
                                        str(row[1])
                                        .replace(",", "")
                                        .replace("₹", "")
                                        .replace("$", "")
                                        .strip()
                                    )
                                )
                            except:
                                y_data.append(0)

                    ax.bar(range(len(y_data)), y_data, color="#217346", alpha=0.8)
                    ax.set_title(title, fontsize=14, fontweight="bold")
                    ax.set_xlabel("Index", fontsize=10)
                    ax.set_ylabel(
                        headers[1] if len(headers) > 1 else "Value", fontsize=10
                    )
                    ax.grid(True, alpha=0.3, axis="y")
        else:
            if column_name:
                try:
                    col_idx = headers.index(column_name)
                    values = [
                        row[col_idx]
                        for row in sampled_rows
                        if col_idx < len(row) and row[col_idx] is not None
                    ]
                    numeric_values = []
                    for v in values:
                        try:
                            numeric_values.append(
                                float(
                                    str(v)
                                    .replace(",", "")
                                    .replace("₹", "")
                                    .replace("$", "")
                                    .strip()
                                )
                            )
                        except:
                            numeric_values.append(0)

                    ax.plot(numeric_values, color="#217346", linewidth=0.8, alpha=0.8)
                    ax.set_title(title, fontsize=14, fontweight="bold")
                    ax.set_xlabel("Index", fontsize=10)
                    ax.set_ylabel(column_name, fontsize=10)
                    ax.grid(True, alpha=0.3)
                except:
                    return None
            else:
                if len(headers) >= 2:
                    y_data = []
                    for row in sampled_rows:
                        if len(row) > 1:
                            try:
                                y_data.append(
                                    float(
                                        str(row[1])
                                        .replace(",", "")
                                        .replace("₹", "")
                                        .replace("$", "")
                                        .strip()
                                    )
                                )
                            except:
                                y_data.append(0)

                    ax.plot(y_data, color="#217346", linewidth=0.8, alpha=0.8)
                    ax.set_title(title, fontsize=14, fontweight="bold")
                    ax.grid(True, alpha=0.3)

        plt.tight_layout()

        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=100, bbox_inches="tight")
        plt.close()

        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode()

        return f"data:image/png;base64,{img_base64}"

    except Exception as e:
        print(f"Chart generation error: {e}")
        return None


@app.post("/api/generate-chart")
async def generate_chart(req: ChartRequest):
    """Generate chart from data - samples large data automatically."""
    try:
        chart_image = generate_single_chart(
            req.data, req.chart_type, req.title, req.column_name
        )

        if chart_image:
            return {
                "chart_image": chart_image,
                "sampled_points": min(1000, len(req.data) - 1),
                "total_rows": len(req.data) - 1,
                "chart_type": req.chart_type,
            }
        else:
            return {"error": "Failed to generate chart"}

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/generate-charts")
async def generate_charts(req: ChartRequest):
    """Generate multiple charts from data."""
    try:
        charts = []

        if req.chart_type == "pie" and req.column_name:
            img = generate_single_chart(req.data, "pie", req.title, req.column_name)
            if img:
                charts.append(
                    {
                        "type": "pie",
                        "title": req.title,
                        "image": img,
                        "column": req.column_name,
                    }
                )

        elif req.chart_type == "bar":
            if req.column_name:
                img = generate_single_chart(req.data, "bar", req.title, req.column_name)
                if img:
                    charts.append(
                        {
                            "type": "bar",
                            "title": req.title,
                            "image": img,
                            "column": req.column_name,
                        }
                    )
            else:
                headers = req.data[0] if len(req.data) > 0 else []
                for col_idx in range(1, min(len(headers), 5)):
                    col_name = headers[col_idx]
                    img = generate_single_chart(
                        req.data, "bar", f"{col_name} Distribution", col_name
                    )
                    if img:
                        charts.append(
                            {
                                "type": "bar",
                                "title": f"{col_name} Distribution",
                                "image": img,
                                "column": col_name,
                            }
                        )

        return {"charts": charts, "total_rows": len(req.data) - 1}

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


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
    try:
        from crewai import Agent, Task, Crew, Process

        model = setup_env(req.provider, req.model, req.api_key)

        # Extract headers and schema context from selected_range for intelligent planning
        headers_context = ""
        schema_context = ""
        if req.selected_range:
            try:
                import json

                # Handle both string (old) and object (new) formats
                sr = req.selected_range
                if isinstance(sr, str):
                    try:
                        sr = json.loads(sr)
                    except:
                        sr = None

                if sr and isinstance(sr, dict):
                    values = sr.get("values", [])
                    sheet_name = sr.get("sheetName", "Sheet")
                    address = sr.get("address", "")

                    if values and len(values) > 0:
                        headers = values[0] if values[0] else []
                        header_names = [h for h in headers if h]
                        if header_names:
                            headers_context = f"Sheet '{sheet_name}' columns ({len(header_names)}): {', '.join(header_names)}\nData range: {address}"
                        schema = build_table_schema(values)
                        schema_context = json.dumps(schema, ensure_ascii=False)
            except Exception as e:
                pass  # Keep silent if extraction fails

        context_info = (
            f"Workbook context: {req.workbook_context}" if req.workbook_context else ""
        )
        selected_info = (
            f"Selected range: {req.selected_range}" if req.selected_range else ""
        )
        memory_info = (
            f"Memory context:\n{req.memory_context}" if req.memory_context else ""
        )

        # Build conversation history for multi-turn context
        history_info = ""
        if req.conversation_history and len(req.conversation_history) > 0:
            recent = req.conversation_history[-6:]  # Last 3 exchanges (6 messages)
            history_lines = []
            for msg in recent:
                role = msg.get("role", "user")
                content = str(msg.get("content", ""))[:300]  # Truncate long messages
                if role == "user":
                    history_lines.append(f"User: {content}")
                elif role == "assistant":
                    history_lines.append(f"Assistant: {content}")
            if history_lines:
                history_info = "RECENT CONVERSATION:\n" + "\n".join(history_lines)

        # Only run reasoning agent if enabled
        reasoning = ""
        if req.enable_reasoning or req.enable_reasoning is None:
            reasoner = Agent(
                role="Excel Reasoning Agent",
                goal="Understand what user wants and explain the approach in clear, concise way",
                backstory="""You are a thoughtful Excel assistant. Your job is to understand user requests quickly and explain your thinking in simple, jargon-free language. 

IMPORTANT:
- Keep responses short and clear (2-4 sentences for summary, 3-4 bullet points for approach)
- Use plain language, no technical jargon
- No asterisks or markdown formatting
- Focus on what user wants, not implementation details
- If request is simple (like "sum A1:A10"), just say what you'll do in one sentence
- If complex (like "analyze 1000 rows and make chart"), break into clear steps""",
                verbose=False,
                allow_delegation=False,
                llm=model,
            )

            reasoning_task = Task(
                description=f"""Analyze this user request briefly:

User: {req.message}
{context_info}
{selected_info}
{history_info}

COLUMN HEADERS: {headers_context or "No column headers available"}
TABLE SCHEMA: {schema_context or "No schema available"}

IMPORTANT - Describe what YOU will do, NOT manual steps:
- Say "I will apply...", "I'll format...", "I'll create..." 
- NEVER say "Go to" or "click" or "select" (that's manual work)
- Example: "I will add black borders to B7:C11 using apply_format"

Explain:
1. What user wants (1 sentence)
2. What you will do (agentic steps)
3. Any notes

No asterisks or markdown.""",
                expected_output="Short explanation, no markdown",
                agent=reasoner,
            )

            reasoning_crew = Crew(
                agents=[reasoner],
                tasks=[reasoning_task],
                process=Process.sequential,
                verbose=False,
            )

            reasoning_result = reasoning_crew.kickoff()
            reasoning = (
                str(reasoning_result.raw)
                if hasattr(reasoning_result, "raw")
                else str(reasoning_result)
            )
            reasoning = reasoning.replace("*", "").strip()

        async def generate():
            yield f"data: {json.dumps({'type': 'reasoning', 'content': reasoning})}\n\n"

            planner = Agent(
                role="Excel Task Planner",
                goal="Create a precise, step-by-step JSON plan for Excel operations",
                backstory=SYSTEM_PROMPT,
                verbose=False,
                allow_delegation=False,
                llm=model,
            )

            plan_task = Task(
                description=TASK_DESCRIPTION.format(
                    message=req.message,
                    context_info=context_info,
                    selected_info=selected_info,
                    headers_context=headers_context,
                    schema_context=schema_context or "No schema available",
                    memory_info=memory_info,
                    history_info=history_info,
                    reasoning_output=reasoning
                    if reasoning
                    else "No reasoning available",
                ),
                expected_output="A JSON object with 'plan' array and 'response' string.",
                agent=planner,
            )

            crew = Crew(
                agents=[planner],
                tasks=[plan_task],
                process=Process.sequential,
                verbose=False,
            )

            result = crew.kickoff()
            output = str(result.raw) if hasattr(result, "raw") else str(result)

            plan = []
            response_text = output

            try:
                json_start = output.find("{")
                json_end = output.rfind("}") + 1
                if json_start >= 0 and json_end > json_start:
                    parsed = json.loads(output[json_start:json_end])
                    plan = parsed.get("plan", [])
                    response_text = parsed.get("response", output)
            except:
                response_text = output

            yield f"data: {json.dumps({'type': 'plan', 'plan': plan, 'content': response_text})}\n\n"

        return StreamingResponse(generate(), media_type="text/event-stream")

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/execute")
async def execute_plan(req: ExecuteRequest):
    try:
        from crewai import Agent, Task, Crew, Process

        model = setup_env(req.provider, req.model, req.api_key)

        validator = Agent(
            role="Excel Validator",
            goal="Review execution results and provide a clear summary",
            backstory=VALIDATOR_PROMPT,
            verbose=True,
            allow_delegation=False,
            llm=model,
        )

        results_str = json.dumps(req.results, indent=2)
        plan_str = (
            "\n".join(
                [
                    f"{i + 1}. {s.get('action', 'unknown')}: {s.get('description', '')}"
                    for i, s in enumerate(req.plan)
                ]
            )
            if req.plan
            else "No plan available"
        )

        validate_task = Task(
            description=VALIDATOR_TASK.format(
                message=req.message, plan_info=plan_str, results_str=results_str
            ),
            expected_output="A clear user-friendly summary.",
            agent=validator,
        )

        crew = Crew(
            agents=[validator],
            tasks=[validate_task],
            process=Process.sequential,
            verbose=True,
        )

        result = crew.kickoff()
        output = str(result.raw) if hasattr(result, "raw") else str(result)

        return {"content": output}

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/analyze")
async def analyze_data(req: AnalyzeRequest):
    try:
        from crewai import Agent, Task, Crew, Process

        model = setup_env(req.provider, req.model, req.api_key)

        stats = compute_statistics(req.data)
        trends = find_trends(req.data)
        outliers = detect_outliers(req.data)
        distribution = analyze_distribution(req.data)

        analysis_results = {
            "statistics": stats,
            "trends": trends,
            "outliers": outliers,
            "distribution": distribution,
        }

        ANALYSIS_AGENT_PROMPT = (
            """You are a concise data analyst. Give brief, point-wise insights."""
        )
        ANALYSIS_TASK = f"""Analyze briefly (3-5 points max):

Headers: {req.headers}
Data: {len(req.data)} rows

Stats: {json.dumps(stats, indent=None)}

Answer: {req.question}

Keep it short, no asterisks (*).

At the end, mention which 2 columns would make a good chart (e.g., "Team vs Total Titles")."""

        analyst = Agent(
            role="Data Analyst",
            goal="Provide clear data insights and answer analytical questions",
            backstory=ANALYSIS_AGENT_PROMPT,
            verbose=False,
            allow_delegation=False,
            llm=model,
        )

        analysis_task = Task(
            description=ANALYSIS_TASK,
            expected_output="Clear analytical insights in natural language.",
            agent=analyst,
        )

        crew = Crew(
            agents=[analyst],
            tasks=[analysis_task],
            process=Process.sequential,
            verbose=False,
        )

        result = crew.kickoff()
        output = str(result.raw) if hasattr(result, "raw") else str(result)

        return {
            "analysis": output,
            "statistics": stats,
            "trends": trends.get("trends", {}),
            "outliers": outliers.get("outliers", {}),
            "distribution": distribution.get("distribution", {}),
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/resolve-chart-intent")
async def resolve_chart_intent(req: ChartIntentRequest):
    try:
        if not req.data or len(req.data) < 2:
            return {
                "needs_clarification": True,
                "reason": "Need at least headers and one data row",
            }

        headers = req.headers or req.data[0] or []
        headers = [str(h).strip() if h is not None else "" for h in headers]
        rows = req.data[1:] if len(req.data) > 1 else []
        if len(headers) < 2:
            return {
                "needs_clarification": True,
                "reason": "Need at least two columns for charting",
            }

        raw_message = req.question or ""
        message = _normalize_text(raw_message)
        chart_type = "column"
        if re.search(r"\bpie\s*chart\b|\bpie\b", message):
            chart_type = "pie"
        elif re.search(r"\bline\s*chart\b|\btrend\b|\bline\b", message):
            chart_type = "line"
        elif re.search(r"\bbar\s*chart\b|\bbar\s*graph\b|\bbar\b", message):
            chart_type = "column"

        start_col_index = req.start_col_index or 0
        scores = [0 for _ in headers]
        picked = set()

        # 1) Exact header mention
        for i, h in enumerate(headers):
            if not h:
                continue
            if re.search(rf"\b{re.escape(h.lower())}\b", message):
                scores[i] += 5
                picked.add(i)

        # 2) Explicit column letter references in "column b and d" style
        explicit_tokens = set()
        forward = re.compile(
            r"\bcolumn[s]?\s+([a-z]{1,3}(?:\s*(?:,|and|&)\s*[a-z]{1,3})*)\b",
            re.IGNORECASE,
        )
        reverse = re.compile(
            r"\b([a-z]{1,3}(?:\s*(?:,|and|&)\s*[a-z]{1,3})+)\s+column[s]?\b",
            re.IGNORECASE,
        )

        def collect_tokens(group_text: str):
            parts = re.split(r"\s*(?:,|and|&)\s*", group_text.lower())
            for part in parts:
                token = part.strip()
                if re.fullmatch(r"[a-z]{1,3}", token):
                    explicit_tokens.add(token)

        for m in forward.finditer(message):
            collect_tokens(m.group(1))
        for m in reverse.finditer(message):
            collect_tokens(m.group(1))

        def col_label_to_index(label: str) -> int:
            n = 0
            for ch in label.upper():
                n = n * 26 + (ord(ch) - 64)
            return n - 1

        for token in explicit_tokens:
            absolute = col_label_to_index(token)
            relative = absolute - start_col_index
            if 0 <= relative < len(headers):
                scores[relative] += 4
                picked.add(relative)

        # 3) Phrase-level extraction: "for X and Y", "X vs Y", "X by Y"
        terms = []
        phrase = re.search(
            r"(?:for|plot|graph|chart)\s+(.+?)\s+(?:and|vs|versus|by)\s+(.+?)(?:$|[,.!?])",
            message,
        )
        if phrase:
            terms.extend([phrase.group(1).strip(), phrase.group(2).strip()])
        else:
            pair = re.search(
                r"(.+?)\s+(?:vs|versus|and|by)\s+(.+?)(?:$|[,.!?])", message
            )
            if pair:
                terms.extend([pair.group(1).strip(), pair.group(2).strip()])

        # Multi-pair extraction for requests like:
        # "Pie chart: Category + Sales, Region + Sales"
        # or numbered lines with multiple pairs.
        multi_pairs: List[List[str]] = []
        normalized_lines = []
        for line in raw_message.splitlines():
            cleaned = re.sub(r"^\s*[\d]+[\)\.\-:]*\s*", "", line.strip())
            if cleaned:
                normalized_lines.append(cleaned)
        if not normalized_lines:
            normalized_lines = [raw_message]

        for line in normalized_lines:
            line_l = line.lower()
            # Explicit plus pairs
            for m in re.finditer(
                r"([a-z0-9][a-z0-9 _/-]{0,40})\s*\+\s*([a-z0-9][a-z0-9 _/-]{0,40})",
                line_l,
            ):
                multi_pairs.append([m.group(1).strip(), m.group(2).strip()])
            # "x and y / x vs y / x by y" pairs
            for m in re.finditer(
                r"([a-z0-9][a-z0-9 _/-]{0,40})\s+(and|vs|versus|by)\s+([a-z0-9][a-z0-9 _/-]{0,40})",
                line_l,
            ):
                left = m.group(1).strip()
                right = m.group(3).strip()
                if left not in {
                    "make",
                    "create",
                    "use",
                    "chart",
                    "pie",
                    "bar",
                    "line",
                } and right not in {"chart"}:
                    multi_pairs.append([left, right])

        # Deduplicate and keep high-signal pairs
        pair_seen = set()
        cleaned_pairs: List[List[str]] = []
        noise_terms = {
            "pie",
            "bar",
            "line",
            "chart",
            "graph",
            "plot",
            "use",
            "make",
            "create",
        }
        for left, right in multi_pairs:
            l = left.strip(" -:,.")
            r = right.strip(" -:,.")
            if not l or not r:
                continue
            if l in noise_terms or r in noise_terms:
                continue
            key = (l, r)
            if key not in pair_seen:
                pair_seen.add(key)
                cleaned_pairs.append([l, r])

        for term in terms:
            idx = _resolve_best_header_index(headers, term, threshold=0.55)
            if idx is not None:
                scores[idx] += 3
                picked.add(idx)

        # 4) Token fallback
        skip_words = {
            "bar",
            "graph",
            "chart",
            "pie",
            "line",
            "make",
            "create",
            "show",
            "display",
            "the",
            "a",
            "an",
            "of",
            "to",
            "for",
            "between",
            "with",
            "on",
            "in",
            "and",
            "vs",
            "versus",
            "column",
            "columns",
            "can",
            "you",
        }
        user_tokens = [
            t for t in _tokenize(message) if len(t) > 1 and t not in skip_words
        ]
        for i, h in enumerate(headers):
            hn = _normalize_text(h)
            h_tokens = _tokenize(hn)
            for tok in user_tokens:
                if (
                    tok in hn
                    or hn in tok
                    or any(
                        ht.startswith(tok[:3]) or tok.startswith(ht[:3])
                        for ht in h_tokens
                        if len(ht) >= 3
                    )
                ):
                    scores[i] += 1
                    picked.add(i)
                    break

        ranked = sorted(list(picked), key=lambda i: scores[i], reverse=True)

        # If user used a clear pair phrase but we cannot confidently map 2 columns, ask.
        if len(terms) >= 2 and len(ranked) < 2:
            suggestions = [headers[i] for i in ranked[:3]] if ranked else headers[:3]
            suggestion_text = (
                ", ".join([s for s in suggestions if s])
                or "the column names in your table"
            )
            return {
                "needs_clarification": True,
                "reason": f'Could not confidently map both requested fields from "{terms[0]}" and "{terms[1]}"',
                "clarification_question": f"I found close columns: {suggestion_text}. Which exact two columns should I use?",
                "chart_type": chart_type,
            }

        x_col = None
        y_col = None

        if len(ranked) >= 2:
            cand = ranked[:4]
            numeric = [i for i in cand if _is_numeric_column(rows, i)]
            categorical = [i for i in cand if i not in numeric]
            x_col = categorical[0] if categorical else cand[0]
            y_col = numeric[0] if numeric else (cand[1] if len(cand) > 1 else cand[0])
            if x_col == y_col and len(cand) > 1:
                y_col = cand[1]

        confidence = 0.0
        if x_col is not None and y_col is not None:
            confidence = min(1.0, min(scores[x_col], scores[y_col]) / 5.0)

        # Resolve multiple requested pairs when present.
        chart_requests = []
        if cleaned_pairs:
            for left, right in cleaned_pairs:
                left_idx = _resolve_best_header_index(headers, left, threshold=0.55)
                right_idx = _resolve_best_header_index(headers, right, threshold=0.55)
                if left_idx is None or right_idx is None or left_idx == right_idx:
                    continue
                left_num = _is_numeric_column(rows, left_idx)
                right_num = _is_numeric_column(rows, right_idx)
                if left_num and not right_num:
                    cx, cy = right_idx, left_idx
                elif right_num and not left_num:
                    cx, cy = left_idx, right_idx
                else:
                    cx, cy = left_idx, right_idx
                chart_requests.append(
                    {
                        "chart_type": chart_type,
                        "x_col_index": cx,
                        "y_col_index": cy,
                        "x_header": headers[cx],
                        "y_header": headers[cy],
                        "source_terms": {"left": left, "right": right},
                    }
                )

        if not chart_requests and x_col is not None and y_col is not None:
            chart_requests.append(
                {
                    "chart_type": chart_type,
                    "x_col_index": x_col,
                    "y_col_index": y_col,
                    "x_header": headers[x_col],
                    "y_header": headers[y_col],
                }
            )

        return {
            "needs_clarification": len(chart_requests) == 0,
            "chart_type": chart_type,
            "x_col_index": x_col,
            "y_col_index": y_col,
            "x_header": headers[x_col] if x_col is not None else None,
            "y_header": headers[y_col] if y_col is not None else None,
            "confidence": confidence,
            "chart_requests": chart_requests,
            "clarification_question": (
                "Please confirm the exact two columns to chart."
                if len(chart_requests) == 0
                else ""
            ),
            "ranked_candidates": [
                {"index": i, "header": headers[i], "score": scores[i]}
                for i in ranked[:6]
            ],
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/build-schema")
async def build_schema_endpoint(req: AnalyzeRequest):
    try:
        return {"schema": build_table_schema(req.data)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/verify-chart-execution")
async def verify_chart_execution(req: VerifyChartExecutionRequest):
    try:
        expected = req.expected_charts or []
        created = req.created_charts or []

        if not expected:
            return {
                "ok": False,
                "status": "needs_review",
                "message": "No expected chart mapping was provided for verification.",
            }

        missing = []
        matched = 0
        for exp in expected:
            exp_x = _normalize_text(exp.get("x_header"))
            exp_y = _normalize_text(exp.get("y_header"))
            found = False
            for got in created:
                got_x = _normalize_text(got.get("x_header"))
                got_y = _normalize_text(got.get("y_header"))
                if exp_x and exp_y and got_x == exp_x and got_y == exp_y:
                    found = True
                    break
            if found:
                matched += 1
            else:
                missing.append(
                    {"x_header": exp.get("x_header"), "y_header": exp.get("y_header")}
                )

        ok = matched == len(expected)
        if ok:
            return {
                "ok": True,
                "status": "verified",
                "message": f"Verified {matched}/{len(expected)} chart mappings.",
                "missing": [],
            }

        return {
            "ok": False,
            "status": "mismatch",
            "message": f"Only {matched}/{len(expected)} requested chart mappings were verified.",
            "missing": missing,
            "retry_suggestion": "Retry with explicit pairs (e.g., Category + Sales, Region + Sales).",
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


class ReflectRequest(BaseModel):
    message: str
    plan: list
    results: list
    provider: str = "openai"
    model: str = ""
    api_key: str


REFLECTION_PROMPT = """You are a self-correction agent. Your job is to analyze failed execution steps and create a recovery plan.

Original user request: {message}

Planned steps:
{plan}

Results:
{results}

Rules:
1. Only create recovery steps for actions that FAILED (success: false)
2. If table creation failed with "overlap", suggest a different approach (e.g., just write data without table, or use a different range)
3. If set_values failed, suggest fixing the data dimensions
4. Keep recovery steps minimal - only fix what's broken
5. Return ONLY valid JSON array of recovery steps

Return format:
[{{"action": "set_values", "params": {{"address": "A1", "values": [...], "sheet_name": "Sheet1"}}, "description": "Retry with corrected data"}}]

If no recovery is possible, return: []"""


@app.post("/api/reflect")
async def reflect(req: ReflectRequest):
    try:
        from crewai import Agent, Task, Crew, Process

        failed_results = [r for r in req.results if not r.get("success")]
        if not failed_results:
            return {"recovery": []}

        model = setup_env(req.provider, req.model, req.api_key)

        plan_str = "\n".join(
            [
                f"{i + 1}. {s.get('action', 'unknown')}: {s.get('description', '')}"
                for i, s in enumerate(req.plan)
            ]
        )
        results_str = "\n".join(
            [
                f"{i + 1}. {r.get('action', 'unknown')}: {'✓' if r.get('success') else '✗'} {r.get('output', '')}"
                for i, r in enumerate(req.results)
            ]
        )

        prompt = REFLECTION_PROMPT.format(
            message=req.message,
            plan=plan_str,
            results=results_str,
        )

        reflector = Agent(
            role="Self-Correction Agent",
            goal="Analyze failed Excel operations and create minimal recovery plans",
            backstory="You are a self-correction agent that analyzes failed steps and returns only valid JSON recovery plans.",
            verbose=False,
            allow_delegation=False,
            llm=model,
        )

        reflect_task = Task(
            description=prompt,
            expected_output="A valid JSON array of recovery steps, or empty array [] if no recovery possible.",
            agent=reflector,
        )

        crew = Crew(
            agents=[reflector],
            tasks=[reflect_task],
            process=Process.sequential,
            verbose=False,
        )

        result = crew.kickoff()
        output = str(result.raw) if hasattr(result, "raw") else str(result)

        # Try to extract JSON from response
        json_match = re.search(r"\[.*\]", output, re.DOTALL)
        if json_match:
            recovery = json.loads(json_match.group())
        else:
            recovery = []

        return {"recovery": recovery}

    except Exception as e:
        print(f"Reflection error: {e}")
        return {"recovery": []}


@app.get("/api/health")
async def health():
    return {"status": "ok", "crewai": "connected"}
