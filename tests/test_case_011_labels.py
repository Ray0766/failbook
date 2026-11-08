"""xlsx-011 reads month labels typed into the chart without trusting the declared point count."""

from xml.etree import ElementTree as ET

import pytest

from failbook.case import load_case

C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


@pytest.fixture(scope="module")
def case011():
    return load_case("xlsx-011").module


def _chart(count, points):
    pts = "".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in points)
    return ET.fromstring(
        f'<c:chartSpace xmlns:c="{C}"><c:chart><c:plotArea><c:lineChart><c:ser>'
        f'<c:cat><c:strLit><c:ptCount val="{count}"/>{pts}</c:strLit></c:cat>'
        f"</c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>")


@pytest.mark.parametrize("text, expected", [
    ("12", 12), ("000000000000000000012", 12), ("0", 0), (" 7 ", 7),
    ("999999999999999999", 999_999_999_999_999_999), ("1" + "0" * 18, 10 ** 18), ("1" + "0" * 30, 10 ** 18),
    ("x", None), ("", None), ("-1", None), ("１２", None),
])
def test_point_numbers_are_bounded_and_ignore_leading_zeros(case011, text, expected):
    assert case011._number(text) == expected


def test_typed_labels_with_a_zero_padded_count_are_read(case011):
    root = _chart("0" * 19 + "12", enumerate(MONTHS))
    assert case011._series_typed_labels(root) == [MONTHS]


def test_a_huge_declared_count_is_reported_without_building_the_list(case011):
    root = _chart("9" * 40, enumerate(MONTHS))
    assert case011._series_typed_labels(root) == [10 ** 18]
