"""Tests for the Cel backend's LLM pipeline and planner context.

Runs against a local mock OpenAI-compatible server, so no API key or network
access is needed:

    cd backend && python test_pipeline.py
"""

import asyncio
import http.server
import json
import logging
import socketserver
import threading
import unittest

import llm
import main
from fastapi import HTTPException

SHEET = [
    ["Region", "Product", "Sales"],
    ["North", "Widget A", 1200],
    ["South", "Widget B", 800],
]
SELECTED = json.dumps(
    {
        "address": "A1:C3",
        "sheetName": "Sheet1",
        "rowCount": 3,
        "columnCount": 3,
        "values": SHEET,
    }
)

RESPONSES = {}
SEEN = []


class MockLLM(http.server.BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
        SEEN.append(body)
        payload = json.dumps(
            {"choices": [{"message": {"content": RESPONSES.get("next", "")}}]}
        ).encode()
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)


class PipelineTest(unittest.IsolatedAsyncioTestCase):
    @classmethod
    def setUpClass(cls):
        logging.getLogger("httpx").setLevel(logging.WARNING)
        cls.server = socketserver.TCPServer(("127.0.0.1", 0), MockLLM)
        cls.port = cls.server.server_address[1]
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()
        import os

        # unknown provider name routes through the OpenAI-compatible branch
        os.environ["CEL_OPENAI_COMPAT_ENDPOINT"] = f"http://127.0.0.1:{cls.port}/v1"
        llm._client = None

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()

    async def asyncSetUp(self):
        SEEN.clear()
        RESPONSES.clear()
        # force a fresh pooled client so each test re-reads the endpoint env
        import os

        os.environ["CEL_OPENAI_COMPAT_ENDPOINT"] = f"http://127.0.0.1:{self.port}/v1"
        llm._client = None

    async def collect(self, response):
        return "".join([chunk async for chunk in response.body_iterator])

    async def chat(self, message="test", **kw):
        return await main.chat(
            main.ChatRequest(
                message=message,
                provider="mockprov",
                model="test-model",
                api_key="test-key",
                selected_range=SELECTED,
                **kw,
            )
        )

    async def test_plan_request_uses_one_llm_call(self):
        RESPONSES["next"] = json.dumps(
            {
                "reasoning": "I will bold the header.",
                "needs_clarification": False,
                "plan": [
                    {
                        "action": "apply_format",
                        "params": {"address": "A1:C1", "bold": True},
                        "description": "Bold header",
                        "confidence": 0.95,
                    }
                ],
                "response": "Bolded the header.",
            }
        )
        out = await self.collect(await self.chat())
        self.assertEqual(len(SEEN), 1, "planning must be a single LLM call")
        self.assertIn('"type": "reasoning"', out)
        self.assertIn("apply_format", out)

    async def test_planner_receives_numeric_and_semantic_columns(self):
        RESPONSES["next"] = json.dumps({"plan": [], "reasoning": "", "response": ""})
        await self.collect(await self.chat(message="chart revenue by region"))
        prompt = SEEN[0]["messages"][-1]["content"]
        self.assertIn("NUMERIC COLUMNS", prompt)
        self.assertIn("Sales", prompt)
        self.assertIn('User "sales" -> header "Sales"', prompt)

    async def test_clarification_halts_before_plan(self):
        RESPONSES["next"] = json.dumps(
            {
                "reasoning": "Ambiguous.",
                "needs_clarification": True,
                "clarification_question": "Which issues?",
                "plan": [],
                "response": "",
            }
        )
        out = await self.collect(await self.chat(message="fix the issues"))
        self.assertIn('"type": "clarification"', out)
        self.assertNotIn('"type": "plan"', out)

    async def test_non_json_response_does_not_crash(self):
        RESPONSES["next"] = "I am sorry, I cannot do that."
        out = await self.collect(await self.chat())
        self.assertIn('"plan": []', out)

    async def test_provider_error_is_502(self):
        original = main.call_llm

        async def boom(*_a, **_k):
            raise RuntimeError("401 invalid key")

        main.call_llm = boom
        try:
            with self.assertRaises(HTTPException) as ctx:
                await self.chat()
            self.assertEqual(ctx.exception.status_code, 502)
        finally:
            main.call_llm = original

    def test_missing_model_is_rejected(self):
        with self.assertRaises(RuntimeError):
            asyncio.run(llm.call_llm("mockprov", "", "k", "sys", "usr"))

    def test_missing_api_key_is_rejected(self):
        with self.assertRaises(RuntimeError):
            asyncio.run(llm.call_llm("mockprov", "m", "", "sys", "usr"))

    def test_validator_failure_never_fails_user_work(self):
        original = main.call_llm

        async def boom(*_a, **_k):
            raise RuntimeError("validator down")

        main.call_llm = boom
        try:
            out = asyncio.run(
                main.execute_plan(
                    main.ExecuteRequest(
                        message="m",
                        plan=[],
                        provider="mockprov",
                        model="m",
                        api_key="k",
                        results=[
                            {"action": "set_values", "success": True, "output": "ok"},
                            {"action": "apply_format", "success": False, "output": "boom"},
                        ],
                    )
                )
            )
            self.assertIn("Completed 1 of 2 steps", out["content"])
        finally:
            main.call_llm = original

    def test_reflector_failure_returns_empty_recovery(self):
        original = main.call_llm

        async def boom(*_a, **_k):
            raise RuntimeError("reflector down")

        main.call_llm = boom
        try:
            out = asyncio.run(
                main.reflect(
                    main.ReflectRequest(
                        message="m",
                        plan=[],
                        results=[{"action": "x", "success": False, "output": "e"}],
                        provider="mockprov",
                        model="m",
                        api_key="k",
                    )
                )
            )
            self.assertEqual(out["recovery"], [])
        finally:
            main.call_llm = original

    def test_reflector_parses_recovery_array(self):
        RESPONSES["next"] = '[{"action": "set_values", "params": {"address": "A1"}}]'
        out = asyncio.run(
            main.reflect(
                main.ReflectRequest(
                    message="m",
                    plan=[],
                    results=[{"action": "x", "success": False, "output": "e"}],
                    provider="mockprov",
                    model="m",
                    api_key="k",
                )
            )
        )
        self.assertEqual(len(out["recovery"]), 1)

    def test_no_successful_failures_skips_llm(self):
        out = asyncio.run(
            main.reflect(
                main.ReflectRequest(
                    message="m",
                    plan=[],
                    results=[{"action": "x", "success": True, "output": "e"}],
                    provider="mockprov",
                    model="m",
                    api_key="k",
                )
            )
        )
        self.assertEqual(out["recovery"], [])
        self.assertEqual(len(SEEN), 0, "no failures means no LLM call")


class PlannerContextTest(unittest.TestCase):
    def build(self, **kw):
        return main.build_planner_context(
            main.ChatRequest(
                message=kw.get("message", "x"),
                provider="openai",
                model="m",
                api_key="k",
                **kw,
            )
        )

    def test_valid_range_yields_headers_and_types(self):
        ctx = self.build(selected_range=SELECTED)
        self.assertIn("Region", ctx["headers_context"])
        self.assertIn("Sales", ctx["numeric_columns"])

    def test_malformed_range_does_not_raise(self):
        ctx = self.build(selected_range="NOT JSON AT ALL")
        self.assertEqual(ctx["headers_context"], "")
        self.assertEqual(ctx["data_profile"], "")

    def test_malformed_metadata_does_not_raise(self):
        ctx = self.build(selected_range=SELECTED, workbook_metadata="{broken")
        self.assertEqual(ctx["workbook_metadata_info"], "")

    def test_none_payload_does_not_raise(self):
        ctx = self.build(selected_range=None, workbook_metadata=None)
        self.assertEqual(ctx["schema_context"], "")


class FormulaContextTest(unittest.TestCase):
    def test_formula_cells_are_listed(self):
        sr = {
            "formulaCells": [
                {"address": "D2", "formula": "=SUM(B2:C2)"},
                {"address": "D3", "formula": "=SUM(B3:C3)"},
            ]
        }
        out = main._format_formula_context(sr)
        self.assertIn("D2: =SUM(B2:C2)", out)
        self.assertIn("D3: =SUM(B3:C3)", out)

    def test_falls_back_to_formulas_matrix(self):
        sr = {"formulas": [["Name", "Total"], ["a", "=SUM(B2:C2)"]]}
        out = main._format_formula_context(sr)
        self.assertIn("B2: =SUM(B2:C2)", out)

    def test_no_formulas_returns_empty(self):
        self.assertEqual(main._format_formula_context({}), "")
        self.assertEqual(main._format_formula_context({"formulas": [["a", "b"]]}), "")

    def test_number_formats_surfaced(self):
        self.assertEqual(main._format_number_formats({"numberFormats": ["#,##0.00"]}), "#,##0.00")
        self.assertEqual(main._format_number_formats({}), "")


class PlannerContextFormulaTest(unittest.TestCase):
    def build(self, sr):
        return main.build_planner_context(
            main.ChatRequest(
                message="x",
                provider="openai",
                model="m",
                api_key="k",
                selected_range=json.dumps(sr),
            )
        )

    def test_formula_and_format_context_reaches_planner(self):
        ctx = self.build(
            {
                "address": "A1:C3",
                "sheetName": "Sheet1",
                "rowCount": 3,
                "columnCount": 3,
                "values": [["Name", "Qty", "Total"], ["a", 2, 3], ["b", 4, 5]],
                "formulaCells": [{"address": "C2", "formula": "=B2*2"}],
                "numberFormats": ["#,##0.00"],
            }
        )
        self.assertIn("C2: =B2*2", ctx["existing_formulas"])
        self.assertIn("#,##0.00", ctx["number_formats"])


class JsonExtractionTest(unittest.TestCase):
    def test_object_surrounded_by_prose(self):
        got = main.extract_json_object('prose {"plan":[{"a":1}]} more prose')
        self.assertEqual(got, {"plan": [{"a": 1}]})

    def test_braces_inside_strings_do_not_break_balancing(self):
        got = main.extract_json_object('{"p":"}{"}')
        self.assertEqual(got["p"], "}{")

    def test_array_with_nested_braces_in_strings(self):
        got = main.extract_json_array('x [{"params":{"v":"}{"}}] y')
        self.assertEqual(got[0]["params"]["v"], "}{")

    def test_no_json_returns_none(self):
        self.assertIsNone(main.extract_json_object("nothing here"))
        self.assertIsNone(main.extract_json_array("nothing here"))

    def test_truncated_json_returns_none(self):
        self.assertIsNone(main.extract_json_object('{"plan": [{"a": 1}'))


class SemanticResolverTest(unittest.TestCase):
    def test_synonym_maps_to_header(self):
        out = main.resolve_semantic_matches(
            "show revenue by region", ["Region", "Product", "Sales"]
        )
        self.assertIn('User "sales" -> header "Sales"', out)

    def test_exact_mention_is_reported(self):
        out = main.resolve_semantic_matches(
            "sort by Qty", ["Region", "Product", "Qty"]
        )
        self.assertIn('header "Qty"', out)

    def test_no_headers_returns_empty(self):
        self.assertEqual(main.resolve_semantic_matches("anything", []), "")

    def test_numeric_detection(self):
        rows = [["1", "x"], ["2", "y"], ["3", "z"], ["4", "w"]]
        # column 0 holds 1,2,3,4 -> numeric; column 1 holds text
        self.assertEqual(main._numeric_column_list(rows, ["count", "label"]), "count")
        # now every column is text
        text_rows = [["x", "y"], ["a", "b"], ["c", "d"], ["e", "f"]]
        self.assertEqual(
            main._numeric_column_list(text_rows, ["name", "label"]),
            "None - all columns look like text",
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
