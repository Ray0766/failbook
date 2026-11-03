"""Verdict logic and error attribution: the parts where a bug would blame the agent unfairly."""

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table

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


# -- #NAME? attribution: blame the agent only for a broken reference that explains the error ---------------
# Each workbook has a data sheet D (A1:B4, an Excel table "T" with columns Key and Amount) and a sheet S
# with the formulas. ``calc`` gives the recalculated values (what LibreOffice produced) for S.

def _book(tmp_path, formulas, *, names=None, local_names=None, calc=None, sheets=("D",)):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for title in sheets:
        wb.create_sheet(title)
    if "D" in sheets:
        d = wb["D"]
        d.append(["Key", "Amount"])
        for i in range(1, 4):
            d.append([f"k{i}", i * 10])
        d.add_table(Table(displayName="T", ref="A1:B4"))
    for ref, f in formulas.items():
        ws[ref] = f
    for name, text in (names or {}).items():
        wb.defined_names[name] = DefinedName(name, attr_text=text)
    for name, text in (local_names or {}).items():
        ws.defined_names[name] = DefinedName(name, attr_text=text)
    wb.save(tmp_path / "book.xlsx")

    values = openpyxl.Workbook()
    vs = values.active
    vs.title = "S"
    for title in sheets:
        values.create_sheet(title)
    for ref in formulas:
        vs[ref] = "#NAME?"  # by default every formula cell came back as #NAME?
    for ref, v in (calc or {}).items():
        vs[ref] = v

    class Recalculated:
        source = "book.xlsx"
        label = "base"

    Recalculated.wb = values
    return Context(FakeCase([]), tmp_path, tmp_path, {}, 0, engine=object()), Recalculated()


def _verdict(ctx, values, ref="S!A1"):
    try:
        ctx._raise_for_error(values, ref, "#NAME?")
    except Fail as exc:
        return FAIL, str(exc)
    except Inconclusive as exc:
        return INCONCLUSIVE, str(exc)
    raise AssertionError("#NAME? must never pass")


@pytest.mark.parametrize("formula", [
    "=_xlfn.LET(_xlpm.r,D!A2,SUMIFS(D!B:B,D!A:A,_xlpm.r))",           # LET as Excel saves it
    "=_xlfn.LET(r,D!A2,SUMIFS(D!B:B,D!A:A,r))",                        # LET parameter without _xlpm.
    "=SUM(_xlfn.MAP(D!B2:B4,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2)))",        # LAMBDA parameter
    "=SUM(_xlfn.MAP(D!B2:B4,_xlfn.LAMBDA(v,v*2)))",
    "=_xlfn.XLOOKUP(\"k2\",T[Key],T[Amount])",                         # structured references to a real table
    "=_xlfn.XLOOKUP(\"k2\",T[[#Data],[Key]],T[Amount])",
    "=_xlfn.XLOOKUP(\"k2\",D!A2:A4,D!B2:B4)+ROWS(T)",                   # a table named without brackets
    "=_xlfn.XLOOKUP(\"k2\",D:D!A2:A4,D!B2:B4)",                          # a 3-D reference whose sheets exist
    "=SUM(_xlfn.BYROW(D!B2:B4,_xlfn._xleta.SUM))",                     # an eta-reduced function name
    "=SUM(_xlfn.BYROW(D!B2:B4, SUM))",                                 # ... written without its prefix
    "=_xlfn.XLOOKUP(\"k2\",D!A2:A4,OFFSET(D!B1,1,0):D!B4)",             # range operator after a function
    "=_xlfn.XLOOKUP(\"k2\",D!A2:A4,D!B2:B4)+@D!B2:B4",                  # implicit intersection
    "=_xlfn.XLOOKUP(\"#REF!\",D!A2:A4,D!B2:B4)",                        # "#REF!" as text is not a reference
])
def test_engine_gaps_with_valid_tokens_are_inconclusive(tmp_path, formula):
    ctx, values = _book(tmp_path, {"A1": formula})
    status, detail = _verdict(ctx, values)
    assert status == INCONCLUSIVE, detail
    assert "no broken reference was found" in detail


@pytest.mark.parametrize("formula, names, expected", [
    ("=SUM(Gone!A1:A3)", {}, "refers to a sheet that does not exist ('Gone')"),
    ("=SUM('Sales 2024'!C:C)", {}, "refers to a sheet that does not exist ('Sales 2024')"),
    ("=SUM(D:Gone!B2)", {}, "refers to a sheet that does not exist ('Gone')"),
    ("=SUM(Revenue)", {"Revenue": "Data!$C$2:$C$13"},
     "uses the name Revenue (= Data!$C$2:$C$13), which points to a sheet that does not exist ('Data')"),
    ("=SUM(Revenue)", {"Revenue": "#REF!"}, "uses the name Revenue (= #REF!), which contains a broken reference"),
    ("=SUM(Revenue)", {"Revenue": "Inner*2", "Inner": "Gone!$A$1"},
     "uses the name Revenue (= Inner*2), which points to a sheet that does not exist ('Gone') "
     "(through the name Inner = Gone!$A$1)"),
    ("=SUM(Nope)", {}, "refers to an undefined name 'Nope'"),
    ("=D!B2*Rate", {}, "refers to an undefined name 'Rate'"),  # also a function name, but not used as one
    ("=SUM(Missing[Amount])", {}, "refers to a table that does not exist ('Missing')"),
    ("=SUM(D!#REF!)", {}, "contains a broken reference"),
    ("=_xlfn.LET(_xlpm.r,Gone!A1,_xlpm.r*2)", {}, "refers to a sheet that does not exist ('Gone')"),
])
def test_broken_references_explain_name_errors(tmp_path, formula, names, expected):
    ctx, values = _book(tmp_path, {"A1": formula}, names=names)
    status, detail = _verdict(ctx, values)
    assert status == FAIL, detail
    assert f"because S!A1 {expected}" in detail, detail


def test_sheet_scoped_name_takes_precedence(tmp_path):
    # Rate is broken at workbook level but defined correctly on S, which is what S's formulas use.
    ctx, values = _book(tmp_path, {"A1": "=_xlfn.XLOOKUP(\"k1\",D!A2:A4,D!B2:B4)*Rate"},
                        names={"Rate": "Gone!$A$1"}, local_names={"Rate": "D!$B$2"})
    assert _verdict(ctx, values)[0] == INCONCLUSIVE


def test_only_the_failing_cell_and_its_precedents_are_examined(tmp_path):
    # A2 is broken, but A1 does not depend on it: its #NAME? is an engine gap.
    ctx, values = _book(tmp_path, {"A1": "=_xlfn.XLOOKUP(\"k1\",D!A2:A4,D!B2:B4)", "A2": "=SUM(Gone!A1)"})
    status, detail = _verdict(ctx, values)
    assert status == INCONCLUSIVE, detail
    # ... while a #NAME? in A2 itself is explained by its missing sheet.
    assert _verdict(ctx, values, "S!A2")[0] == FAIL


def test_name_error_is_followed_to_the_precedent_that_causes_it(tmp_path):
    ctx, values = _book(tmp_path, {"A1": "=A2*2", "A2": "=SUM(Revenue)"}, names={"Revenue": "Data!$C$2:$C$13"})
    status, detail = _verdict(ctx, values)
    assert status == FAIL
    assert "because S!A2 uses the name Revenue" in detail


def test_precedents_without_an_error_value_are_not_blamed(tmp_path):
    # A2's formula is broken, but its recalculated value is fine (it is not what A1's #NAME? comes from).
    ctx, values = _book(tmp_path, {"A1": "=_xlfn.XLOOKUP(A2,D!A2:A4,D!B2:B4)", "A2": "=IFERROR(Nope,\"k1\")"},
                        calc={"A2": "k1"})
    assert _verdict(ctx, values)[0] == INCONCLUSIVE


def test_a_stored_error_constant_is_a_failure(tmp_path):
    ctx, values = _book(tmp_path, {"A1": "=A2+1"})
    wb = openpyxl.load_workbook(tmp_path / "book.xlsx")
    wb["S"]["A2"] = "#NAME?"  # pasted values that include the error itself
    wb.save(tmp_path / "book.xlsx")
    values.wb["S"]["A2"] = "#NAME?"
    status, detail = _verdict(ctx, values)
    assert status == FAIL and "S!A2 holds the error value #NAME? as a constant" in detail
