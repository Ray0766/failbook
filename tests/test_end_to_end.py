"""Slower checks that need LibreOffice: the grader and the checker validation on one case."""

from failbook.case import load_case
from failbook.grade import verify

from conftest import needs_libreoffice


@needs_libreoffice
def test_verify_one_case_two_seeds():
    outcomes = verify(load_case("xlsx-001"), [0, 1])
    for o in outcomes:
        assert o.ok, o.problems
        assert o.failure.verdict == "fail"
