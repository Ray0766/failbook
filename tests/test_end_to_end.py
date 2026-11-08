"""Slower checks that need LibreOffice: the grader and the checker validation on real cases."""

import openpyxl
from openpyxl.worksheet.table import Table

from failbook.case import load_case
from failbook.grade import grade, prepare, verify

from conftest import needs_libreoffice


@needs_libreoffice
def test_verify_one_case_two_seeds():
    outcomes = verify(load_case("xlsx-001"), [0, 1])
    for o in outcomes:
        assert o.ok, o.problems
        assert o.failure.verdict == "fail"
        assert set(o.resaved) == {"reference", *o.alternatives}


@needs_libreoffice
def test_stale_defined_name_is_a_failure_not_an_engine_gap():
    # Formulas rewritten after the rename, but RevenueRange still points at the deleted sheet.
    [o] = verify(load_case("xlsx-006"), [0])
    assert o.ok, o.problems
    fixture = o.fixtures["fail_stale_defined_name"]
    assert fixture.verdict == "fail"
    detail = {r.id: r.detail for r in fixture.rules}["dashboard-values"]
    assert "uses the name RevenueRange" in detail and "sheet that does not exist ('Data')" in detail


def _not_failed(report):
    """Valid formulas the engine may lack must never fail; if undecided, the reason must say so."""
    assert report.verdict != "fail", [(r.id, r.detail) for r in report.rules if r.status == "fail"]
    for r in report.rules:
        if r.status == "inconclusive":
            assert "no broken reference was found" in r.detail, r.detail


@needs_libreoffice
def test_let_formula_is_not_blamed_on_the_agent(tmp_path):
    case = load_case("xlsx-006")
    meta = prepare(case, tmp_path)
    case.solve("reference", tmp_path, meta)
    path = tmp_path / "report.xlsx"
    wb = openpyxl.load_workbook(path)
    wb["Dashboard"]["B2"] = "=_xlfn.LET(_xlpm.r,A2,SUMIFS('Sales 2024'!C:C,'Sales 2024'!A:A,_xlpm.r))"
    wb.save(path)
    _not_failed(grade(case, tmp_path))


@needs_libreoffice
def test_xlookup_through_a_table_is_not_blamed_on_the_agent(tmp_path):
    case = load_case("xlsx-009")
    prepare(case, tmp_path)
    path = tmp_path / "orders.xlsx"
    wb = openpyxl.load_workbook(path)
    products, orders = wb["Products"], wb["Orders"]
    products.add_table(Table(displayName="ProductTable", ref=f"A1:C{products.max_row}"))
    for r in range(2, orders.max_row + 1):
        orders.cell(r, 4, f"=_xlfn.XLOOKUP(B{r},ProductTable[SKU],ProductTable[Category])")
        orders.cell(r, 5, f"=_xlfn.XLOOKUP(B{r},ProductTable[SKU],ProductTable[Unit Price])")
        orders.cell(r, 6, f"=C{r}*E{r}")
    wb.save(path)
    _not_failed(grade(case, tmp_path))
