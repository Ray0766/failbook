import openpyxl

from failbook import kit
from failbook import xlsx as X
from failbook.checks import require, rule

SHEET = "Targets"
HEADERS = ["Region", "2024 Revenue", "Growth Target", "2025 Target"]
PCTS = [3, 4.5, 5, 6, 7.5, 8, 8.5, 10, 12, 15, 18]


def params(seed):
    r = kit.rng(seed, "xlsx-008")
    targets = []
    for region in kit.REGIONS:
        targets.append({
            "region": region,
            "revenue": r.randrange(400_000, 2_500_000, 1000),
            "old": r.choice([0.02, 0.03, 0.04, 0.05, 0.06]),
            "pct": r.choice(PCTS),
        })
    table = kit.md_table(["Region", "Growth target"], [(t["region"], f"{t['pct']:g}%") for t in targets])
    return {"targets": targets, "targets_md": table}


def build(meta, outdir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.append(HEADERS)
    kit.style_header(ws, 1, len(HEADERS))
    for r, t in enumerate(meta["targets"], start=2):
        ws.append([t["region"], t["revenue"], t["old"], f"=B{r}*(1+C{r})"])
        ws.cell(r, 2).number_format = kit.INT_FMT
        ws.cell(r, 3).number_format = kit.PCT_FMT
        ws.cell(r, 4).number_format = kit.INT_FMT
    last = 1 + len(meta["targets"])
    total = last + 1
    ws.cell(total, 1, "Total").font = kit.TOTAL_FONT
    ws.cell(total, 2, f"=SUM(B2:B{last})").number_format = kit.INT_FMT
    ws.cell(total, 3, f"=D{total}/B{total}-1").number_format = kit.PCT_FMT
    ws.cell(total, 4, f"=SUM(D2:D{last})").number_format = kit.INT_FMT
    for c in range(1, 5):
        ws.cell(total, c).font = kit.TOTAL_FONT
    kit.set_widths(ws, {"A": 12, "B": 14, "C": 14, "D": 14})
    wb.save(outdir / "targets.xlsx")


def _amt(x, places=2, sign=False):
    """A number with at least ``places`` and at most 6 decimals, so that two numbers the checker
    tells apart never print the same."""
    whole, _, frac = (f"{x:+,.6f}" if sign else f"{x:,.6f}").rstrip("0").partition(".")
    frac = frac.ljust(places, "0")
    return f"{whole}.{frac}" if frac else whole


def _whole(x):
    return _amt(x, 0)


def _pct(x):
    return _amt(x * 100, 1) + "%"


def _rows(ctx):
    return list(zip(range(2, 2 + len(ctx.meta["targets"])), ctx.meta["targets"]))


def _replanned(meta):
    """[(row, target, revenue, growth)] in the "replanned" scenario: in a copy of the file, the k-th
    region's 2024 Revenue is raised by 10,000 x k and its Growth Target set to its new target plus k
    percentage points (k = 1, 2, ...), so every one of them differs from what the task asks for, by
    a different amount per region, and a 2025 Target, a total or an overall growth figure that is a
    typed number, or that has a region's growth typed into its formula, no longer matches."""
    return [(r, t, t["revenue"] + 10_000 * k, round(t["pct"] / 100 + 0.01 * k, 4))
            for k, (r, t) in enumerate(zip(range(2, 2 + len(meta["targets"])), meta["targets"]), start=1)]


def scenarios(ctx):
    overrides = {}
    for r, _, revenue, growth in _replanned(ctx.meta):
        overrides[X.make_ref(SHEET, (r, 2))] = revenue
        overrides[X.make_ref(SHEET, (r, 3))] = growth
    return {"replanned": overrides}


@rule(
    "growth-values",
    "Each growth target holds the new percentage as a fraction (12% = 0.12)",
    kind="content",
    title_zh="每个增长目标都以小数存储新的百分比（12% = 0.12）",
    requirement="Replace the growth targets with the new ones.",
)
def growth_values(ctx):
    vals = ctx.values().ws(SHEET)
    bad = []
    for r, t in _rows(ctx):
        got = vals.cell(r, 3).value
        exp = t["pct"] / 100
        if not X.is_number(got):
            bad.append(f"{t['region']}: {X.show(got)} is not a number")
        elif X.close(got, t["pct"]):
            bad.append(f"{t['region']}: stored as {got:g}, which is {got:.0%}, not {t['pct']:g}%")
        elif not X.close(got, exp):
            bad.append(f"{t['region']}: {got:g} instead of {exp:g}")
    require(not bad, "; ".join(bad[:3]) + (f" (and {len(bad) - 3} more)" if len(bad) > 3 else ""))


@rule(
    "percent-format-kept",
    "The Growth Target cells are still formatted as percentages",
    kind="structure",
    title_zh="Growth Target 单元格仍为百分比格式",
    requirement="Do not change anything the task does not ask you to change.",
)
def percent_format_kept(ctx):
    ws = ctx.ws(SHEET)
    bad = [t["region"] for r, t in _rows(ctx) if not X.is_percent_format(X.number_format(ws, r, 3))]
    require(not bad, f"Growth Target is no longer formatted as a percentage for {', '.join(bad)}")


@rule(
    "targets-correct",
    "2025 Target, the Total row and the overall growth reflect the new targets",
    kind="value",
    title_zh="2025 Target、Total 行和整体增长率都反映新的目标",
    requirement="The 2025 Target column, the Total row and the overall growth figure must reflect the new targets.",
)
def targets_correct(ctx):
    values = ctx.values()
    total_rev = total_target = 0.0
    for r, t in _rows(ctx):
        exp = t["revenue"] * (1 + t["pct"] / 100)
        total_rev += t["revenue"]
        total_target += exp
        got = ctx.num(values, X.make_ref(SHEET, (r, 4)), f"2025 Target of {t['region']}")
        require(X.close(got, exp), f"2025 Target of {t['region']} is {_whole(got)}, expected {_whole(exp)}")
    total = 2 + len(ctx.meta["targets"])
    got = ctx.num(values, X.make_ref(SHEET, (total, 4)), "Total 2025 Target")
    require(X.close(got, total_target), f"Total 2025 Target is {_whole(got)}, expected {_whole(total_target)}")
    got = ctx.num(values, X.make_ref(SHEET, (total, 3)), "overall growth")
    exp = total_target / total_rev - 1
    require(X.close(got, exp), f"overall growth is {_pct(got)}, expected {_pct(exp)}")


@rule(
    "targets-still-calculated",
    "2025 Target, the Total row and the overall growth still follow the revenue and growth targets",
    kind="preserve",
    title_zh="2025 Target、Total 行和整体增长率仍随收入和增长目标变化",
    requirement="The 2025 Target column, the Total row and the overall growth figure must ... keep being "
                "calculated from the revenue and growth figures, as they are now.",
)
def targets_still_calculated(ctx):
    base, changed = ctx.values(), ctx.values(scenario="replanned")
    rows = _replanned(ctx.meta)
    total = 2 + len(ctx.meta["targets"])
    checks = [(X.make_ref(SHEET, (r, 4)), f"2025 Target of {t['region']}", revenue * (1 + growth), _whole)
              for r, t, revenue, growth in rows]
    total_rev = sum(revenue for _, _, revenue, _ in rows)
    total_target = sum(revenue * (1 + growth) for _, _, revenue, growth in rows)
    checks += [
        (X.make_ref(SHEET, (total, 2)), "Total 2024 Revenue", total_rev, _whole),
        (X.make_ref(SHEET, (total, 4)), "Total 2025 Target", total_target, _whole),
        (X.make_ref(SHEET, (total, 3)), "overall growth", total_target / total_rev - 1, _pct),
    ]
    wrong = []
    for ref, what, exp, fmt in checks:
        got = ctx.num(changed, ref, what)
        if X.close(got, exp):
            continue
        before = base[ref]
        same = ", the same as before the change" if X.is_number(before) and X.close(before, got) else ""
        wrong.append(f"{what} is {fmt(got)}, expected {fmt(exp)}{same}")
    require(not wrong, "in a copy of the file where every 2024 Revenue was raised by a different amount and every "
                       "Growth Target set to its new target plus a different number of percentage points (1 to "
                       f"{len(rows)}), {len(wrong)} of {len(checks)} figures did not follow: " + "; ".join(wrong[:3])
                       + (f" (and {len(wrong) - 3} more)" if len(wrong) > 3 else ""))


@rule(
    "revenue-unchanged",
    "Regions and 2024 revenue are unchanged",
    kind="preserve",
    title_zh="区域与 2024 年收入保持不变",
    requirement="Do not change anything the task does not ask you to change.",
)
def revenue_unchanged(ctx):
    vals = ctx.values().ws(SHEET)
    for r, t in _rows(ctx):
        got = vals.cell(r, 1).value
        require(got == t["region"], f"row {r} holds {X.show(got)}, expected {t['region']!r}")
        got = vals.cell(r, 2).value
        require(X.close(got, t["revenue"]), f"2024 Revenue of {t['region']} is {X.show(got)}, expected {t['revenue']:,}")
