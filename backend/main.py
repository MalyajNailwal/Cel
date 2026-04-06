from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional, List, Any
import os
import json
import re
import io
import base64
from collections import Counter

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
    model: str = "gpt-4o"
    api_key: str
    workbook_context: Optional[str] = None
    selected_range: Optional[str] = None


class ExecuteRequest(BaseModel):
    message: str
    provider: str = "openai"
    model: str = "gpt-4o"
    api_key: str
    results: list


class AnalyzeRequest(BaseModel):
    data: List[List[Any]]
    question: str
    provider: str = "openai"
    model: str = "gpt-4o"
    api_key: str
    headers: Optional[List[str]] = None


class ChatResponse(BaseModel):
    content: str
    plan: list


def setup_provider(provider: str, model: str, api_key: str):
    """
    Root fix: Configure environment variables for CrewAI/LiteLLM.
    This avoids passing LangChain objects which cause version conflicts.
    """
    # Clear previous keys to avoid conflicts
    keys_to_clear = [
        "OPENAI_API_KEY",
        "OPENAI_API_BASE",
        "OPENAI_MODEL_NAME",
        "ANTHROPIC_API_KEY",
        "ANTHROPIC_MODEL_NAME",
        "GOOGLE_API_KEY",
        "GOOGLE_MODEL_NAME",
    ]
    for k in keys_to_clear:
        os.environ.pop(k, None)

    if provider == "openrouter":
        os.environ["OPENAI_API_KEY"] = api_key
        os.environ["OPENAI_API_BASE"] = "https://openrouter.ai/api/v1"
        os.environ["OPENAI_MODEL_NAME"] = model
    elif provider == "openai":
        os.environ["OPENAI_API_KEY"] = api_key
        os.environ["OPENAI_MODEL_NAME"] = model
    elif provider == "anthropic":
        os.environ["ANTHROPIC_API_KEY"] = api_key
        os.environ["ANTHROPIC_MODEL_NAME"] = model
    elif provider == "google":
        os.environ["GOOGLE_API_KEY"] = api_key
        os.environ["GOOGLE_MODEL_NAME"] = model


SYSTEM_PROMPT = """You are an expert Excel analyst. You create detailed, executable plans for Excel operations.
Each step must have a specific action, target range, and all parameters needed.
Always use A1 notation for ranges. Always read before writing if you need context."""

TASK_DESCRIPTION = """Create a JSON plan for this Excel request: {message}

{context_info}
{selected_info}

Available actions: get_workbook_structure, get_selected_range, get_range, get_sheet_data,
set_values, set_formulas, apply_format, insert_rows, delete_rows, insert_columns,
delete_columns, add_worksheet, delete_worksheet, create_table, sort_range, auto_fill, create_chart

IMPORTANT - COLOR FORMATTING:
- When user asks for "color", "grey", "red", "blue", "highlight", "background color", "font color" → use apply_format with fill_color or font_color
- Color names to hex mapping: grey/gray=#808080, red=#FF0000, blue=#0000FF, green=#008000, yellow=#FFFF00, white=#FFFFFF, black=#000000, orange=#FFA500, purple=#800080, pink=#FFC0CB, brown=#A52A2A, navy=#000080, teal=#008080, maroon=#800000, olive=#808000, lime=#00FF00, coral=#FF7F50, salmon=#FA8072, tomato=#FF6347, indigo=#4B0082, violet=#EE82EE, gold=#FFD700, crimson=#DC143C, emerald=#50C878, jade=#00A86B, bronze=#CD7F32, steel=#4682B4, skyblue=#87CEEB, royalblue=#4169E1, dodgerblue=#1E90FF
- User can use ANY color name - if unknown, try common variations or default to #808080 (grey)

IMPORTANT - BORDERS:
- When user asks for "border", "add border", "outer border", "all borders", "grid lines", "pivot style border" → use apply_format with border_all=true, border_color, border_style, border_style options: Continuous, Dash, Dot, DashDot, Double; weight options: Hairline, Thin, Medium, Thick; default: border_style="Continuous", border_weight="Thin", border_color="#000000"
- Example: "add border" → apply_format with border_all=true, border_color="#000000", border_style="Continuous", border_weight="Thin"
- Example: "thick border" → apply_format with border_all=true, border_weight="Thick"

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
8. For set_values, params must include "address" and "values" (2D array)
9. For set_formulas, params must include "address" and "formulas" (2D array)
10. For apply_format, params must include "address" and format properties (bold, fill_color, font_color, number_format, etc.)
11. Be specific with cell ranges — match the selected range dimensions
12. Keep it simple and correct"""

VALIDATOR_PROMPT = """You verify that Excel operations completed successfully and summarize what was done."""
VALIDATOR_TASK = """Original request: {message}

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

        setup_provider(req.provider, req.model, req.api_key)

        stats = compute_large_data_stats(req.data)

        ANALYSIS_TASK = f"""You are a Large Data Analyst. Analyze this dataset and answer: {req.question}

Dataset Overview:
- Total Rows: {stats.get("overview", {}).get("total_rows", "N/A")}
- Total Columns: {stats.get("overview", {}).get("total_columns", "N/A")}
- Headers: {stats.get("overview", {}).get("headers", [])}

Column Statistics:
{json.dumps(stats.get("columns", {}), indent=2)}

Provide clear, actionable insights. Focus on:
1. Key findings from the numbers
2. Any patterns or trends
3. Important observations
4. Recommendations if applicable"""

        analyst = Agent(
            role="Large Data Analyst",
            goal="Analyze large datasets efficiently and provide insights",
            backstory="""You are an expert at analyzing large datasets. 
            You compute statistics efficiently without trying to process every row through the LLM.
            You focus on aggregated insights and patterns rather than individual data points.""",
            verbose=False,
            allow_delegation=False,
            llm=req.model,
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


@app.post("/api/generate-chart")
async def generate_chart(req: ChartRequest):
    """Generate chart from data - samples large data automatically."""
    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import numpy as np

        if not req.data or len(req.data) < 2:
            return {"error": "Not enough data"}

        headers = req.data[0]
        rows = req.data[1:]

        sample_size = min(1000, len(rows))
        step = max(1, len(rows) // sample_size)
        sampled_rows = rows[::step]

        x_data = None
        y_data = None

        if req.x_column and req.y_column:
            try:
                x_idx = headers.index(req.x_column) if req.x_column in headers else 0
                y_idx = headers.index(req.y_column) if req.y_column in headers else 1

                x_data = [
                    row[x_idx]
                    for row in sampled_rows
                    if x_idx < len(row) and row[x_idx] is not None
                ]
                y_data = [
                    row[y_idx]
                    for row in sampled_rows
                    if y_idx < len(row) and row[y_idx] is not None
                ]
            except:
                pass

        if not x_data or not y_data:
            if len(headers) >= 2:
                x_data = [
                    row[0]
                    for row in sampled_rows
                    if len(row) > 0 and row[0] is not None
                ]
                y_data = [
                    row[1]
                    for row in sampled_rows
                    if len(row) > 1 and row[1] is not None
                ]
            else:
                x_data = list(range(len(sampled_rows)))
                y_data = [
                    row[0]
                    for row in sampled_rows
                    if len(row) > 0 and row[0] is not None
                ]

        numeric_y = []
        for v in y_data:
            try:
                numeric_y.append(
                    float(str(v).replace(",", "").replace("₹", "").replace("$", ""))
                )
            except:
                numeric_y.append(0)

        plt.figure(figsize=(12, 6))

        if req.chart_type == "bar":
            plt.bar(range(len(numeric_y)), numeric_y, color="#217346", alpha=0.8)
        elif req.chart_type == "pie":
            unique_vals = list(set(numeric_y[:20]))
            counts = [numeric_y[:20].count(v) for v in unique_vals]
            plt.pie(
                counts, labels=[str(v)[:10] for v in unique_vals], autopct="%1.1f%%"
            )
        elif req.chart_type == "scatter":
            plt.scatter(range(len(numeric_y)), numeric_y, alpha=0.5, color="#217346")
        else:
            plt.plot(numeric_y, color="#217346", linewidth=1, alpha=0.8)

        plt.title(req.title, fontsize=14, fontweight="bold")
        plt.xlabel("Index", fontsize=10)
        plt.ylabel("Value", fontsize=10)
        plt.grid(True, alpha=0.3)
        plt.tight_layout()

        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=100, bbox_inches="tight")
        plt.close()

        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode()

        return {
            "chart_image": f"data:image/png;base64,{img_base64}",
            "sampled_points": len(sampled_rows),
            "total_rows": len(rows),
            "chart_type": req.chart_type,
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/chat")
async def chat(req: ChatRequest):
    try:
        from crewai import Agent, Task, Crew, Process

        setup_provider(req.provider, req.model, req.api_key)

        planner = Agent(
            role="Excel Task Planner",
            goal="Create a precise, step-by-step JSON plan for Excel operations",
            backstory=SYSTEM_PROMPT,
            verbose=True,
            allow_delegation=False,
            llm=req.model,
        )

        context_info = (
            f"Workbook context: {req.workbook_context}" if req.workbook_context else ""
        )
        selected_info = (
            f"Selected range: {req.selected_range}" if req.selected_range else ""
        )

        plan_task = Task(
            description=TASK_DESCRIPTION.format(
                message=req.message,
                context_info=context_info,
                selected_info=selected_info,
            ),
            expected_output="A JSON object with 'plan' array and 'response' string.",
            agent=planner,
        )

        crew = Crew(
            agents=[planner],
            tasks=[plan_task],
            process=Process.sequential,
            verbose=True,
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

        return ChatResponse(content=response_text, plan=plan)

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/execute")
async def execute_plan(req: ExecuteRequest):
    try:
        from crewai import Agent, Task, Crew, Process

        setup_provider(req.provider, req.model, req.api_key)

        validator = Agent(
            role="Excel Validator",
            goal="Review execution results and provide a clear summary",
            backstory=VALIDATOR_PROMPT,
            verbose=True,
            allow_delegation=False,
            llm=req.model,
        )

        results_str = json.dumps(req.results, indent=2)

        validate_task = Task(
            description=VALIDATOR_TASK.format(
                message=req.message, results_str=results_str
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

        setup_provider(req.provider, req.model, req.api_key)

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

        ANALYSIS_AGENT_PROMPT = """You are an expert data analyst. You analyze Excel data and provide clear, actionable insights."""
        ANALYSIS_TASK = f"""Analyze this data and answer: {req.question}

Data headers: {req.headers}
Data shape: {len(req.data)} rows, {len(req.data[0]) if req.data else 0} columns

Statistical analysis:
{json.dumps(analysis_results, indent=2)}

Provide a clear, comprehensive analysis in natural language. Highlight key findings, patterns, and any concerns."""

        analyst = Agent(
            role="Data Analyst",
            goal="Provide clear data insights and answer analytical questions",
            backstory=ANALYSIS_AGENT_PROMPT,
            verbose=False,
            allow_delegation=False,
            llm=req.model,
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


@app.get("/api/health")
async def health():
    return {"status": "ok", "crewai": "connected"}
