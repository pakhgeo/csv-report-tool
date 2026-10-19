"""Тесты очистки и разбора значений."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from clean import (
    CANONICAL_COLUMNS,
    clean_sales,
    detect_separator,
    load_sales,
    normalize_text,
    parse_date,
    parse_number,
    prepare,
    read_text,
)

SAMPLE_DIR = Path(__file__).resolve().parent.parent / "sample_data"


class TestParseNumber:
    @pytest.mark.parametrize(
        "raw, expected",
        [
            ("1250.00", 1250.0),
            ("1250,50", 1250.5),
            ("1 250,50", 1250.5),
            ("1 250.50", 1250.5),
            ("1.250,50", 1250.5),
            ("1,250.50", 1250.5),
            ("1250.00 руб.", 1250.0),
            ("3 400", 3400.0),
            ("1\u00a0250", 1250.0),
        ],
    )
    def test_valid_numbers(self, raw: str, expected: float) -> None:
        assert parse_number(raw) == expected

    @pytest.mark.parametrize("raw", ["", " ", "-", "—", "нет", None, "abc"])
    def test_empty_and_garbage(self, raw: str | None) -> None:
        assert parse_number(raw) is None


class TestParseDate:
    @pytest.mark.parametrize(
        "raw, expected",
        [
            ("2025-01-08", "2025-01-08"),
            ("08.01.2025", "2025-01-08"),
            ("08/01/2025", "2025-01-08"),
            (" 2025-01-08 ", "2025-01-08"),
        ],
    )
    def test_valid_dates(self, raw: str, expected: str) -> None:
        assert parse_date(raw) == pd.Timestamp(expected)

    @pytest.mark.parametrize("raw", ["", "31.02.2025", "не дата", None])
    def test_invalid_dates(self, raw: str | None) -> None:
        assert parse_date(raw) is None


class TestNormalizeText:
    def test_strips_and_collapses(self) -> None:
        assert normalize_text("  Анна \t Петрова  ") == "Анна Петрова"

    def test_removes_nbsp(self) -> None:
        assert normalize_text("Кофе\u00a0в зёрнах") == "Кофе в зёрнах"

    def test_nan_becomes_empty(self) -> None:
        assert normalize_text(float("nan")) == ""


class TestDetectSeparator:
    @pytest.mark.parametrize(
        "header, expected",
        [
            ("date,manager,product", ","),
            ("Дата;Менеджер;Товар", ";"),
            ("date\tmanager\tproduct", "\t"),
        ],
    )
    def test_detects(self, header: str, expected: str) -> None:
        assert detect_separator(header + "\nx,y,z") == expected


class TestReadText:
    def test_detects_cp1251(self) -> None:
        text, encoding = read_text(SAMPLE_DIR / "sales_february.csv")
        assert encoding == "cp1251"
        assert text.splitlines()[0].startswith("Дата;Менеджер")

    def test_detects_utf8(self) -> None:
        _, encoding = read_text(SAMPLE_DIR / "sales_january.csv")
        assert encoding == "utf-8"


class TestCleanSales:
    def test_all_three_files_merge(self) -> None:
        df, stats = prepare(
            [
                SAMPLE_DIR / "sales_january.csv",
                SAMPLE_DIR / "sales_february.csv",
                SAMPLE_DIR / "sales_march.csv",
            ]
        )

        assert stats.files == 3
        assert stats.encodings["sales_february.csv"] == "cp1251"
        assert stats.blank_rows == 2, "в марте пустая строка и строка из пробелов"
        assert stats.duplicates == 1, "в марте полный дубль строки"
        assert stats.bad_quantity == 1, "в марте строка без количества"
        assert stats.bad_price == 1, "в марте строка без цены"
        assert stats.bad_date == 0

        assert stats.input_rows == 30
        assert stats.output_rows == 25
        assert stats.dropped_rows() == stats.input_rows - stats.output_rows

        assert list(df.columns) == list(CANONICAL_COLUMNS)
        assert df["quantity"].dtype == "int64"
        assert (df["quantity"] > 0).all()
        assert (df["price"] > 0).all()
        assert df["date"].is_monotonic_increasing

    def test_manager_whitespace_is_normalized(self) -> None:
        df, _ = prepare([SAMPLE_DIR / "sales_march.csv"])
        assert "Анна" in set(df["manager"])
        assert not any(name != name.strip() for name in df["manager"])

    def test_nbsp_in_product_survives_but_collapses(self) -> None:
        df, _ = prepare([SAMPLE_DIR / "sales_march.csv"])
        assert "Кофе в зёрнах" in set(df["product"])

    def test_empty_input(self) -> None:
        empty = pd.DataFrame(columns=list(CANONICAL_COLUMNS))
        df, stats = clean_sales(empty)
        assert df.empty
        assert stats.input_rows == 0
        assert stats.output_rows == 0

    def test_single_row(self) -> None:
        raw = pd.DataFrame(
            [["2025-01-01", "Анна", "Чай", "Москва", "2", "100,00"]],
            columns=list(CANONICAL_COLUMNS),
        )
        df, stats = clean_sales(raw)
        assert len(df) == 1
        assert df.loc[0, "price"] == 100.0
        assert stats.input_rows == 1
        assert stats.output_rows == 1

    def test_all_rows_broken_gives_empty_result(self) -> None:
        raw = pd.DataFrame(
            [
                ["не дата", "Анна", "Чай", "Москва", "2", "100"],
                ["2025-01-01", "Анна", "Чай", "Москва", "", "100"],
            ],
            columns=list(CANONICAL_COLUMNS),
        )
        df, stats = clean_sales(raw)
        assert df.empty
        assert stats.bad_date == 1
        assert stats.bad_quantity == 1
        assert stats.output_rows == 0

    def test_zero_and_negative_dropped(self) -> None:
        raw = pd.DataFrame(
            [
                ["2025-01-01", "Анна", "Чай", "Москва", "0", "100"],
                ["2025-01-02", "Пётр", "Чай", "Казань", "1", "-5"],
                ["2025-01-03", "Мария", "Чай", "Сочи", "1", "100"],
            ],
            columns=list(CANONICAL_COLUMNS),
        )
        df, stats = clean_sales(raw)
        assert len(df) == 1
        assert stats.bad_quantity == 1
        assert stats.bad_price == 1


class TestMissingColumns:
    def test_reports_missing_columns(self) -> None:
        raw = pd.DataFrame(
            [["2025-01-01", "Анна"]],
            columns=["date", "manager"],
        )
        from clean import standardize_columns

        with pytest.raises(ValueError, match="не хватает колонок"):
            standardize_columns(raw)

    def test_extra_columns_are_reported(self) -> None:
        from clean import standardize_columns

        raw = pd.DataFrame(
            [["2025-01-01", "Анна", "Чай", "Москва", "1", "10", "мусор"]],
            columns=[*CANONICAL_COLUMNS, "Комментарий"],
        )
        df, ignored = standardize_columns(raw)
        assert ignored == ["Комментарий"]
        assert list(df.columns) == list(CANONICAL_COLUMNS)
