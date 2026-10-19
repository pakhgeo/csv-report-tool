"""Тесты агрегации."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from clean import CANONICAL_COLUMNS, prepare
from report import resolve_group_by, summarize

SAMPLE_DIR = Path(__file__).resolve().parent.parent / "sample_data"


def make_df(rows: list[list[object]]) -> pd.DataFrame:
    df = pd.DataFrame(rows, columns=list(CANONICAL_COLUMNS))
    df["date"] = pd.to_datetime(df["date"])
    df["quantity"] = df["quantity"].astype("int64")
    df["price"] = df["price"].astype("float64")
    return df


class TestResolveGroupBy:
    @pytest.mark.parametrize(
        "name, expected",
        [
            ("product", "product"),
            ("Товар", "product"),
            ("товар", "product"),
            ("Менеджер", "manager"),
            ("город", "city"),
            ("дата", "date"),
        ],
    )
    def test_aliases(self, name: str, expected: str) -> None:
        assert resolve_group_by(name)[0] == expected

    def test_unknown_column(self) -> None:
        with pytest.raises(ValueError, match="Нельзя группировать"):
            resolve_group_by("выручка")


class TestSummarize:
    def test_revenue_by_product(self) -> None:
        df = make_df(
            [
                ["2025-01-01", "Анна", "Кофе", "Москва", 2, 100.0],
                ["2025-01-02", "Пётр", "Кофе", "Казань", 3, 100.0],
                ["2025-01-03", "Мария", "Чай", "Сочи", 1, 50.0],
            ]
        )
        result = summarize(df, "product")

        assert result.group_by == "product"
        assert result.totals.orders == 3
        assert result.totals.quantity == 6
        assert result.totals.revenue == 550.0
        assert result.totals.groups == 2
        assert result.totals.avg_order_value == pytest.approx(183.33, abs=0.01)

        coffee = result.table.iloc[0]
        assert coffee["group"] == "Кофе"
        assert coffee["orders"] == 2
        assert coffee["quantity"] == 5
        assert coffee["revenue"] == 500.0
        assert coffee["avg_order_value"] == 250.0
        assert coffee["share_percent"] == pytest.approx(90.91, abs=0.01)

    def test_sorted_by_revenue_desc(self) -> None:
        df = make_df(
            [
                ["2025-01-01", "Анна", "Чай", "Москва", 1, 10.0],
                ["2025-01-02", "Анна", "Кофе", "Москва", 1, 100.0],
                ["2025-01-03", "Анна", "Кофемолка", "Москва", 1, 50.0],
            ]
        )
        table = summarize(df, "product").table
        assert list(table["group"]) == ["Кофе", "Кофемолка", "Чай"]

    def test_share_sums_to_hundred(self) -> None:
        df = make_df(
            [
                ["2025-01-01", "Анна", "Кофе", "Москва", 1, 33.33],
                ["2025-01-02", "Пётр", "Чай", "Казань", 1, 33.33],
                ["2025-01-03", "Мария", "Сок", "Сочи", 1, 33.34],
            ]
        )
        table = summarize(df, "product").table
        assert table["share_percent"].sum() == pytest.approx(100.0, abs=0.01)

    def test_empty_dataframe(self) -> None:
        empty = pd.DataFrame(columns=list(CANONICAL_COLUMNS))
        result = summarize(empty, "product")
        assert result.table.empty
        assert result.totals.orders == 0
        assert result.totals.revenue == 0.0
        assert list(result.table.columns) == [
            "group", "orders", "quantity", "revenue", "avg_order_value", "share_percent",
        ]

    def test_single_row(self) -> None:
        df = make_df([["2025-01-01", "Анна", "Кофе", "Москва", 2, 10.0]])
        result = summarize(df, "manager")
        assert result.totals.groups == 1
        assert result.table.iloc[0]["share_percent"] == 100.0

    def test_empty_group_goes_to_placeholder(self) -> None:
        df = make_df([["2025-01-01", "Анна", "", "Москва", 1, 10.0]])
        result = summarize(df, "product")
        assert result.table.iloc[0]["group"] == "(не указано)"
        assert result.warnings

    def test_nan_group_goes_to_placeholder(self) -> None:
        df = make_df([["2025-01-01", "Анна", float("nan"), "Москва", 1, 10.0]])
        result = summarize(df, "product")
        assert result.table.iloc[0]["group"] == "(не указано)"

    def test_does_not_mutate_input(self) -> None:
        df = make_df([["2025-01-01", "Анна", "Кофе", "Москва", 1, 10.0]])
        before = df.copy()
        summarize(df, "product")
        pd.testing.assert_frame_equal(df, before)

    def test_real_sample_data(self) -> None:
        df, _ = prepare(
            [
                SAMPLE_DIR / "sales_january.csv",
                SAMPLE_DIR / "sales_february.csv",
                SAMPLE_DIR / "sales_march.csv",
            ]
        )
        result = summarize(df, "Товар")
        assert result.totals.orders == len(df) == 25
        assert result.totals.quantity == int(df["quantity"].sum())
        assert result.totals.revenue == pytest.approx(
            float((df["quantity"] * df["price"]).sum()), abs=0.01
        )
        assert result.table["share_percent"].sum() == pytest.approx(100.0, abs=0.05)
        assert set(result.table.columns) == {
            "group", "orders", "quantity", "revenue", "avg_order_value", "share_percent",
        }
