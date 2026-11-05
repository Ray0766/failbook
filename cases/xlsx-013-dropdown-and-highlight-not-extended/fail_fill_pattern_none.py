"""Wrong: the highlight's fill is written with patternType="none" next to a red bgColor, as some writers do when the
pattern is left unset. "none" means no background in the file format, and LibreOffice paints no fill. openpyxl cannot
write "none", so this patches styles.xml."""

import re
import shutil
import zipfile

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.worksheet.cell_range import MultiCellRange

from failbook import kit


def solve(workdir, meta):
    path = workdir / "tracker.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb["Tracker"]
    for t in meta["new"]:
        ws.append([t["id"], t["task"], t["owner"], t["status"], kit.from_iso(t["due"])])
        ws.cell(ws.max_row, 5).number_format = kit.DATE_FMT
    last = ws.max_row
    for dv in ws.data_validations.dataValidation:
        dv.sqref = MultiCellRange(f"D2:D{last}")
    ws.conditional_formatting = ConditionalFormattingList()
    ws.conditional_formatting.add(
        f"D2:D{last}", CellIsRule(operator="equal", formula=['"Blocked"'], fill=PatternFill(bgColor="FFFF0000")))
    wb.save(path)
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/styles.xml":
                dxfs = re.search(rb"<dxfs.*?</dxfs>", data, re.S).group(0)
                data = data.replace(dxfs, dxfs.replace(b"<patternFill>", b'<patternFill patternType="none">'))
            out.writestr(item, data)
    shutil.move(tmp, path)
