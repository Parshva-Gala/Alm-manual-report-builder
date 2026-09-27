"""Focused checks for the chart layout contract, without Excel automation.

Execute the production VBA's pure ChartGrid and ReserveChartRows procedures
through a deliberately small, fail-closed syntax adapter. The geometry is not
reimplemented as a Python preview. A fake worksheet supplies Excel-like row
height rounding and its 409.5pt ceiling. These checks establish rectangle and
reservation invariants, not native PivotChart rendering compatibility.

Run: python -B build/chart_layout_check.py
"""
from pathlib import Path
import re
import unittest


SOURCE = (Path(__file__).resolve().parents[1] / "src/vba/modPD_Charts.bas").read_text(encoding="cp1252")
CONSTANTS = {name: float(value) for name, value in re.findall(
    r"Private Const (\w+) As Double = ([\d.]+)", SOURCE)}


def procedure(name, source=SOURCE):
    match = re.search(r"(?:Public|Private) Function " + name + r"\(.*?\) As \w+\n(.*?)\nEnd Function",
                      source, re.S)
    if not match:
        raise AssertionError(f"Missing production procedure: {name}")
    return match.group(1)


def expression(text):
    text = re.sub(r"\bAnd\b", "and", text)
    text = re.sub(r"\bOr\b", "or", text)
    text = re.sub(r"\bCStr\b", "str", text)
    return text.replace("New Collection", "[]")


def compile_procedure(name, args, namespace):
    """Translate only the numeric/collection subset these production helpers use.

    Unsupported statements fail the check. This intentionally cannot open a
    workbook, call COM, run arbitrary VBA, or automate an application.
    """
    lines = [f"def {name}({args}):"]
    depth = 1
    for original in procedure(name).splitlines():
        raw = original.strip()
        if not raw or raw.startswith("'"):
            continue
        if raw.startswith("Dim "):
            for part in raw[4:].split(","):
                var, typ = part.strip().split(" As ")
                initial = "0.0" if typ == "Double" else "0" if typ == "Long" else "None"
                lines.append("    " * depth + var + " = " + initial)
            continue
        for statement in raw.split(": "):
            if statement in ("End If", "Loop") or statement.startswith("Next "):
                depth -= 1
                continue
            if statement.startswith("If "):
                condition, tail = statement[3:].split(" Then", 1)
                lines.append("    " * depth + "if " + expression(condition) + ":")
                if not tail.strip():
                    depth += 1
                    continue
                statement = tail.strip()
                prefix = "    " * (depth + 1)
            elif statement.startswith("For Each "):
                var, values = statement[9:].split(" In ")
                lines.append("    " * depth + "for " + var + " in " + values + ":")
                depth += 1
                continue
            elif statement.startswith("Do While "):
                lines.append("    " * depth + "while " + expression(statement[9:]) + ":")
                depth += 1
                continue
            else:
                prefix = "    " * depth
            if statement.startswith("Err.Raise "):
                message = re.search(r', "[^"]*", ("[^"]*")$', statement).group(1)
                code = "raise ValueError(" + message + ")"
            elif statement.startswith("boxes.Add "):
                code = "boxes.append(" + expression(statement[10:]) + ")"
            else:
                statement = statement.removeprefix("Set ")
                lhs, separator, rhs = statement.partition(" = ")
                if not separator or not re.fullmatch(r"[\w.()]+", lhs):
                    raise AssertionError(f"Unsupported production statement: {statement}")
                code = ("return " if lhs == name else lhs + " = ") + expression(rhs)
            lines.append(prefix + code)
    if depth != 1:
        raise AssertionError("Unbalanced production control flow")
    exec("\n".join(lines), namespace)
    return namespace[name]


def width_for(size, wide):
    # Select Case expressions come directly from the production procedure.
    for choices, formula in re.findall(r"Case (.*?): WidthFor = (.*)", procedure("WidthFor")):
        if choices == "Else" or size in re.findall(r'"([^"]*)"', choices):
            return eval(expression(formula), {"__builtins__": {}}, dict(CONSTANTS, wide=wide))
    raise AssertionError("WidthFor does not cover the input")


NAMESPACE = dict(CONSTANTS, WidthFor=width_for, Array=lambda *values: tuple(values))
GRID = compile_procedure("ChartGrid", "sizes, wide, high=250", NAMESPACE)
RESERVE = compile_procedure("ReserveChartRows", "ws, firstRow, high", NAMESPACE)


class Row:
    def __init__(self, rows, index):
        self.rows, self.index = rows, index

    @property
    def Height(self):
        return self.rows.heights.get(self.index, 22.5)

    @property
    def RowHeight(self):
        return self.Height

    @RowHeight.setter
    def RowHeight(self, value):
        if not 0 <= value <= 409.5:
            raise ValueError("Excel row-height ceiling")
        self.rows.requested.append(value)
        self.rows.heights[self.index] = round(value * 4) / 4


class Rows:
    count = 1_048_576

    def __init__(self):
        self.heights, self.requested = {}, []

    def __call__(self, index):
        return Row(self, index)


class Sheet:
    def __init__(self):
        self.Rows = Rows()


class ChartLayout(unittest.TestCase):
    def test_known_one_two_three_positions(self):
        self.assertEqual(GRID(["full"], 1032), [(0, 6, 1032, 250)])
        self.assertEqual(GRID(["half"] * 2, 1032), [(0, 6, 510, 250), (522, 6, 510, 250)])
        self.assertEqual(GRID(["third"] * 3, 1032),
                         [(0, 6, 336, 250), (348, 6, 336, 250), (696, 6, 336, 250)])

    def test_mixed_sizes_preserve_order_and_gutters(self):
        sizes = ["third", "two thirds", "half", "third", "full", "half", "half", "third", "third", "third"]
        boxes = GRID(sizes, 1032)
        self.assertEqual([b[1] for b in boxes], [6, 6, 268, 268, 530, 792, 792, 1054, 1054, 1054])
        self.assertEqual(boxes[1], (348, 6, 684, 250))
        self.assertEqual(boxes[4], (0, 530, 1032, 250))

    def test_hundreds_of_mixed_charts_do_not_overlap_or_cross_index(self):
        for wide in (1032, 1064, 779.25):
            for high in (250, 520):
                boxes = GRID(["third", "half", "two thirds", "full", "half"] * 30, wide, high)
                ws = Sheet()
                end = RESERVE(ws, 8, boxes[-1][1] + high + 6)
                reserved = sum(ws.Rows(r).Height for r in range(8, end))
                self.assertGreater(end, 9)
                self.assertTrue(all(0 < h <= 200 for h in ws.Rows.requested))
                for i, (x, y, w, h) in enumerate(boxes):
                    self.assertGreaterEqual(x, 0)
                    self.assertLessEqual(x + w, wide + .001)
                    self.assertGreaterEqual(reserved - (y + h), 6 - .01)
                    for ox, oy, ow, oh in boxes[:i]:
                        overlap = min(x + w, ox + ow) - max(x, ox) > .001 and min(y + h, oy + oh) - max(y, oy) > .001
                        self.assertFalse(overlap)

    def test_own_chart_tall_reservation_and_rounding(self):
        ws = Sheet()
        self.assertEqual(RESERVE(ws, 4, 536), 7)
        self.assertEqual(ws.Rows.requested, [200, 200, 136])
        fractional = Sheet()
        end = RESERVE(fractional, 8, 262.1)
        self.assertGreaterEqual(sum(fractional.Rows(r).Height for r in range(8, end)), 262.1 - .01)

    def test_empty_and_invalid_space(self):
        self.assertEqual(GRID([], 1032), [])
        self.assertEqual(RESERVE(Sheet(), 8, 0), 8)
        with self.assertRaises(ValueError):
            GRID(["third"], 72)
        with self.assertRaises(ValueError):
            GRID(["full"], 1032, 12)
        with self.assertRaises(ValueError):
            RESERVE(Sheet(), 1_048_576, 520)

    def test_reservation_precedes_chart_creation(self):
        body = procedure("PlaceOnGuide")
        self.assertLess(body.index("ClearChartPanels ws"), body.index("DrawChart ws"))
        self.assertLess(body.index("ReserveChartRows(ws, r, bottom)"), body.index("origin = ws.Rows(r).Top"))
        self.assertLess(body.index("origin = ws.Rows(r).Top"), body.index("DrawChart ws"))
        self.assertNotIn(".RowHeight =", body)
        # The former draw-then-enlarge order fails this invariant. Free-floating
        # objects no longer depend on Excel moving anchors as row heights change.
        own = procedure("DrawOnSheet")
        self.assertLess(own.index("ReserveChartRows ws"), own.index("DrawChart(ws"))
        self.assertIn("ws.Delete", own)

    def test_chart_binding_is_followed_by_final_bounds_and_error_cleanup(self):
        body = procedure("DrawChart")
        self.assertLess(body.index("ch.SetSourceData"), body.index("co.Left = l + 8"))
        self.assertLess(body.index("StyleLabels ch"), body.index("co.Left = l + 8"))
        self.assertIn("co.Width = w - 16: co.Height = h - 12", body)
        self.assertIn("If Not co Is Nothing Then co.Delete", body)
        self.assertIn("If Not card Is Nothing Then card.Delete", body)


if __name__ == "__main__":
    unittest.main(verbosity=2)
