from pathlib import Path

from geodin_ggf_tool.backend import (
    build_final_date_list,
    find_duplicate_dates,
    remove_duplicates_keep_order,
    DEFAULT_DUMMY_DATE,
)


def test_build_final_date_list_fills_dummy_dates():
    dates = ["20240101", "20240202"]
    result = build_final_date_list(dates, max_slots=5, dummy_date=DEFAULT_DUMMY_DATE)

    assert result == [
        "20240101",
        "20240202",
        "19000101",
        "19000101",
        "19000101",
    ]


def test_find_duplicate_dates():
    dates = ["20240101", "20240202", "20240101", "20240303"]
    assert find_duplicate_dates(dates) == ["20240101"]


def test_remove_duplicates_keep_order():
    dates = ["20240101", "20240202", "20240101", "20240303", "20240202"]
    assert remove_duplicates_keep_order(dates) == ["20240101", "20240202", "20240303"]
