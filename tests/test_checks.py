"""Verdict logic and error attribution: the parts where a bug would blame the agent unfairly."""

import openpyxl
import pytest

from failbook.checks import (FAIL, INCONCLUSIVE, PASS, Context, Fail, Inconclusive, RuleResult, RuleSpec,
                             run_rules, verdict_of)


class FakeCase:
    def __init__(self, rules):
        self.rules = rules
        self.outputs = ["book.xlsx"]
        self.inputs = ["book.xlsx"]
        self.module = None


def spec(rid, fn):
    return RuleSpec(id=rid, title=rid, requirement="", title_zh=None, func=fn)


def test_exceptions_map_to_statuses():
    def ok(ctx):
        pass

    def fails(ctx):
        raise Fail("wrong")

    def asserts(ctx):
        assert False, "also wrong"

    def unsure(ctx):
        raise Inconclusive("cannot tell")

    def buggy(ctx):
        raise KeyError("oops")

    case = FakeCase([spec("a", ok), spec("b", fails), spec("c", asserts), spec("d", unsure), spec("e", buggy)])
    results = {r.id: r for r in run_rules(case, ctx=None)}
    assert results["a"].status == PASS
    assert results["b"].status == FAIL and results["b"].detail == "wrong"
    assert results["c"].status == FAIL and "also wrong" in results["c"].detail
    assert results["d"].status == INCONCLUSIVE
    # A crash in the checker must never count against the agent.
    assert results["e"].status == INCONCLUSIVE and "checker error" in results["e"].detail


def test_verdict_precedence():
    r = lambda s: RuleResult("x", "x", s)
    assert verdict_of([r(PASS), r(PASS)]) == PASS
    assert verdict_of([r(PASS), r(INCONCLUSIVE)]) == INCONCLUSIVE
    assert verdict_of([r(INCONCLUSIVE), r(FAIL)]) == FAIL


def _ctx_with_formula(tmp_path, formula, extra_sheet=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    if extra_sheet:
        wb.create_sheet(extra_sheet)
    ws["A1"] = formula
    wb.save(tmp_path / "book.xlsx")
    case = FakeCase([])
    return Context(case, tmp_path, tmp_path, {}, 0, engine=object())


class FakeValues:
    source = "book.xlsx"


def test_name_error_from_missing_sheet_is_a_failure(tmp_path):
    ctx = _ctx_with_formula(tmp_path, "=SUM(Data!C2:C9)")
    with pytest.raises(Fail, match="sheet that does not exist"):
        ctx._raise_for_error(FakeValues(), "S!A1", "#NAME?")


def test_name_error_without_broken_reference_is_inconclusive(tmp_path):
    ctx = _ctx_with_formula(tmp_path, "=XLOOKUP(1,Data!A1:A3,Data!B1:B3)", extra_sheet="Data")
    with pytest.raises(Inconclusive, match="does not support"):
        ctx._raise_for_error(FakeValues(), "S!A1", "#NAME?")


def test_other_error_values_fail(tmp_path):
    ctx = _ctx_with_formula(tmp_path, "=1/0")
    with pytest.raises(Fail, match="#DIV/0!"):
        ctx._raise_for_error(FakeValues(), "S!A1", "#DIV/0!")
