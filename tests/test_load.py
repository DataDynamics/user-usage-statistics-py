"""입력 파싱 회귀 테스트: python -m pytest tests"""

from pathlib import Path

import pytest

import usage_stats as u


def load_lines(tmp_path: Path, lines: list[str]):
    path = tmp_path / "log.csv"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return u.load(path)


def test_quoted_rows_with_comma_and_spaces(tmp_path):
    df, errors = load_lines(tmp_path, [
        'View,fharenheit,"887,576",18,2026-10-10 0:11:11,2026-10-10 0:14:16,3:05',
        'View,fharenheit," 887,576 ",18,2026-10-10 0:11:11,2026-10-10 0:14:16,3:05',
        'Batch, user01 , "  1,234,567  " , 5 , 2026-10-10 2:00:00 , 2026-10-10 2:01:00 , 1:00',
    ])
    assert len(errors) == 0
    assert df["rows"].tolist() == [887576, 887576, 1234567]
    assert df["user"].tolist() == ["fharenheit", "fharenheit", "user01"]


@pytest.mark.parametrize("charts", ["18", " 18 ", '"18"', '" 18 "'])
def test_chart_count_trimmed(tmp_path, charts):
    df, errors = load_lines(tmp_path, [
        f'View,u,"1,000",{charts},2026-10-10 0:00:00,2026-10-10 0:00:10,0:10',
    ])
    assert len(errors) == 0
    assert df["charts"].tolist() == [18]


def test_hour_without_zero_padding(tmp_path):
    lines = [f'View,u,"1,000",1,2026-10-10 {h}:05:07,2026-10-10 {h}:05:17,0:10' for h in range(24)]
    df, errors = load_lines(tmp_path, lines)
    assert len(errors) == 0
    assert df["start"].dt.hour.tolist() == list(range(24))
    assert df["end"].dt.hour.tolist() == list(range(24))


@pytest.mark.parametrize("raw, sec", [
    ("0:05", 5), ("9:59", 599), ("10:00", 600), ("59:59", 3599),
    ("60:00", 3600), ("125:30", 7530), (" 3:05 ", 185), ("1:02:03", 3723),
])
def test_elapsed_minutes_grow_digits(raw, sec):
    assert u.parse_elapsed(raw) == sec


def test_bad_row_reported(tmp_path):
    df, errors = load_lines(tmp_path, [
        'View,u,"1,000",1,2026-10-10 0:00:00,2026-10-10 0:00:10,0:10',
        'View,u,"12,x",1,2026-10-10 0:00:00,2026-10-10 0:00:10,0:10',
    ])
    assert len(df) == 1
    assert errors["line_no"].tolist() == [2]
    assert errors["오류컬럼"].tolist() == ["데이터수"]
