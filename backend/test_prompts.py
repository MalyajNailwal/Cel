"""Template wiring and prompt-completeness checks for the Cel backend.

Every prompt constant must (a) render without leftover {placeholders} and
(b) receive every placeholder it declares. The planner silently degrades if a
new context section is added to TASK_DESCRIPTION but not to the call site.

    cd backend && python test_prompts.py
"""

import json
import re
import unittest

import main

TEMPLATES = [
    "TASK_DESCRIPTION",
    "VALIDATOR_TASK",
    "REFLECTION_PROMPT",
    "SYSTEM_PROMPT",
    "CLARIFICATION_PROMPT",
]


def template_source(name: str) -> str:
    match = re.search(name + r' = """(.*?)"""', open("main.py").read(), re.S)
    assert match, f"{name} not found in main.py"
    return match.group(1)


def placeholders(name: str) -> set:
    return set(re.findall(r"\{(\w+)\}", template_source(name)))


def call_kwargs(name: str):
    """Read kwargs from the .format() call. Captures nested parens so a kwarg
    like reasoning_instruction=(...) is not truncated by a naive regex."""
    source = open("main.py").read()
    anchor = source.find(name + ".format(")
    if anchor < 0:
        return None
    i = anchor + len(name) + len(".format(")
    depth = 1
    start = i
    while i < len(source) and depth:
        if source[i] == "(":
            depth += 1
        elif source[i] == ")":
            depth -= 1
        i += 1
    return set(re.findall(r"(\w+)\s*=", source[start : i - 1]))


class TemplateWiringTest(unittest.TestCase):
    def test_every_template_with_placeholders_is_formatted(self):
        for name in TEMPLATES:
            used = placeholders(name)
            if not used:
                continue
            kwargs = call_kwargs(name)
            self.assertIsNotNone(
                kwargs, f"{name} declares {sorted(used)} but is never .format()ed"
            )

    def test_every_placeholder_is_supplied(self):
        for name in TEMPLATES:
            used = placeholders(name)
            kwargs = call_kwargs(name)
            if not used or kwargs is None:
                continue
            missing = used - kwargs
            self.assertFalse(
                missing, f"{name} placeholders never supplied: {sorted(missing)}"
            )

    def test_templates_render_without_leftover_braces(self):
        for name in TEMPLATES:
            used = placeholders(name)
            if not used:
                continue
            rendered = template_source(name).format(**{k: f"<{k}>" for k in used})
            self.assertEqual(
                re.findall(r"\{(\w+)\}", rendered),
                [],
                f"{name} leaves unfilled placeholders after .format()",
            )


class PlannerContextSectionsTest(unittest.TestCase):
    """Every context section must actually reach the planner prompt."""

    def render_prompt(self, selected_range=None, message="sum sales by region") -> str:
        captured = {}

        async def fake_call(*args, **kwargs):
            captured["user"] = args[4]
            return json.dumps({"reasoning": "", "plan": [], "response": ""})

        original = main.call_llm
        main.call_llm = fake_call
        try:
            import asyncio

            asyncio.run(
                main.chat(
                    main.ChatRequest(
                        message=message,
                        provider="p",
                        model="m",
                        api_key="k",
                        selected_range=selected_range,
                    )
                )
            )
        finally:
            main.call_llm = original
        return captured["user"]

    def test_all_context_sections_present(self):
        selected = json.dumps(
            {
                "address": "A1:C3",
                "sheetName": "Sheet1",
                "rowCount": 3,
                "columnCount": 3,
                "values": [
                    ["Region", "Sales", "Qty"],
                    ["North", 1200, 10],
                    ["South", 800, 4],
                ],
                "formulaCells": [{"address": "B2", "formula": "=C2*120"}],
                "numberFormats": ["#,##0"],
            }
        )
        prompt = self.render_prompt(selected)
        for section in (
            "EXISTING FORMULAS IN THE SELECTION",
            "NUMBER FORMATS ALREADY IN USE",
            "NUMERIC COLUMNS",
            "SEMANTIC COLUMN MATCHES",
            "TABLE SCHEMA",
            "DATA PROFILE",
            "CLARIFICATION - check BEFORE planning",
        ):
            self.assertIn(section, prompt, f"planner prompt missing: {section}")

    def test_context_values_reach_prompt(self):
        selected = json.dumps(
            {
                "address": "A1:C3",
                "sheetName": "Sheet1",
                "rowCount": 3,
                "columnCount": 3,
                "values": [
                    ["Region", "Sales", "Qty"],
                    ["North", 1200, 10],
                    ["South", 800, 4],
                ],
                "formulaCells": [{"address": "B2", "formula": "=C2*120"}],
                "numberFormats": ["#,##0"],
            }
        )
        prompt = self.render_prompt(selected)
        self.assertIn("B2: =C2*120", prompt)
        self.assertIn("#,##0", prompt)
        self.assertIn("Sales, Qty", prompt)

    def test_no_selection_omits_clarification_section(self):
        prompt = self.render_prompt(None)
        self.assertNotIn("CLARIFICATION - check BEFORE planning", prompt)
        self.assertEqual(re.findall(r"\{(\w+)\}", prompt), [])

    def test_every_new_action_is_documented(self):
        prompt = self.render_prompt(None)
        for action in (
            "copy_range",
            "create_pivot_table",
            "auto_fit_columns",
            "set_row_height",
            "set_column_width",
            "freeze_panes",
            "add_comment",
            "add_hyperlink",
            "text_to_columns",
            "remove_duplicates",
        ):
            self.assertIn(action, prompt, f"planner prompt never mentions {action}")

    def test_conditional_format_variants_documented(self):
        prompt = self.render_prompt(None)
        for variant in ("colorScale", "dataBar", "duplicateValues", "topBottom"):
            self.assertIn(variant, prompt)

    def test_formula_safety_rules_documented(self):
        prompt = self.render_prompt(None)
        self.assertIn("NEVER overwrite a cell listed in EXISTING FORMULAS", prompt)


if __name__ == "__main__":
    unittest.main(verbosity=2)