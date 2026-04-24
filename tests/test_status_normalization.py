import pytest
from app.domain.models import EmployeeStatus
from app.domain.statuses import normalize_status

STATUS_MAPPING = {
    "office": ["офис", "office"],
    "remote": ["удаленно", "удалённо", "remote", "wfh"],
    "vacation": ["отпуск", "vacation", "ooo"],
    "day_off": ["выходной", "day off"],
}


def test_office_russian():
    assert normalize_status("офис", STATUS_MAPPING) == EmployeeStatus.OFFICE


def test_office_english():
    assert normalize_status("office", STATUS_MAPPING) == EmployeeStatus.OFFICE


def test_remote_russian():
    assert normalize_status("удаленно", STATUS_MAPPING) == EmployeeStatus.REMOTE


def test_remote_alternative_spelling():
    assert normalize_status("удалённо", STATUS_MAPPING) == EmployeeStatus.REMOTE


def test_vacation():
    assert normalize_status("отпуск", STATUS_MAPPING) == EmployeeStatus.VACATION


def test_day_off():
    assert normalize_status("выходной", STATUS_MAPPING) == EmployeeStatus.DAY_OFF


def test_unknown():
    assert normalize_status("что-то_новое", STATUS_MAPPING) == EmployeeStatus.UNKNOWN


def test_whitespace_stripping():
    assert normalize_status("  офис  ", STATUS_MAPPING) == EmployeeStatus.OFFICE


def test_case_insensitive():
    assert normalize_status("ОФИС", STATUS_MAPPING) == EmployeeStatus.OFFICE
