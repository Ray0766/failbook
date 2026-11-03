"""Reading recalculated values must not depend on whether openpyxl can parse the workbook's charts."""

import re
import zipfile

import openpyxl
import pytest
from openpyxl.chart import LineChart, Reference

from failbook.checks import Values, _open_workbook


def _chart_workbook(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Monthly"
    ws.append(["Month", "Revenue"])
    for month, amount in (("Jan", 100), ("Feb", 120), ("Mar", 90)):
        ws.append([month, amount])
    ws["C2"] = "=B2*2"
    chart = LineChart()
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=4), titles_from_data=True)
    ws.add_chart(chart, "E2")
    wb.save(path)


def _empty_plot_area(src, dst):
    """Copy ``src`` with every chart's plot area emptied of chart groups, as LibreOffice 24.2 saves a line chart
    whose series all point at OFFSET names (only the plot area's shape properties are left)."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info)
            if re.fullmatch(r"xl/charts/chart\d+\.xml", info.filename):
                data, n = re.subn(rb"<((?:c:)?plotArea)>.*</\1>", rb"<\1></\1>", data, flags=re.S)
                assert n == 1
            zout.writestr(info, data)


def test_values_are_read_when_a_chart_has_no_chart_group(tmp_path):
    good, broken = tmp_path / "good.xlsx", tmp_path / "broken.xlsx"
    _chart_workbook(good)
    _empty_plot_area(good, broken)
    try:
        openpyxl.load_workbook(broken, data_only=True)
    except IndexError:
        pass  # openpyxl 3.1 stops here; the loader must not
    values = Values(broken, "base", "broken.xlsx")
    assert values["Monthly!A4"] == "Mar"
    assert values["Monthly!B3"] == 120
    assert values.ws("Monthly").max_row == 4
    assert _open_workbook(broken, data_only=False)["Monthly"]["C2"].value == "=B2*2"  # formulas too


def test_a_workbook_openpyxl_reads_keeps_its_charts(tmp_path):
    good = tmp_path / "good.xlsx"
    _chart_workbook(good)
    values = Values(good, "base", "good.xlsx")
    assert len(values.ws("Monthly")._charts) == 1


def test_other_load_errors_are_raised_unchanged(tmp_path):
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not a zip file")
    with pytest.raises(zipfile.BadZipFile):
        Values(bad, "base", "bad.xlsx")
