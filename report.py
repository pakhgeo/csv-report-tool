"""Агрегация очищенных данных: сводка по выбранной колонке и итоговая строка."""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from clean import CANONICAL_COLUMNS, normalize_text

# Русские названия колонок, чтобы в CLI можно было писать --group-by Товар.
GROUP_ALIASES: dict[str, str] = {
    "date": "date",
    "дата": "date",
    "manager": "manager",
    "менеджер": "manager",
    "product": "product",
    "товар": "product",
    "city": "city",
    "город": "city",
}

SUM_COLUMNS = ["orders", "quantity", "revenue"]


@dataclass
class Totals:
    """Итоги по всем строкам — для шапки отчёта."""

    orders: int = 0
    quantity: int = 0
    revenue: float = 0.0
    avg_order_value: float = 0.0
    groups: int = 0


@dataclass
class Summary:
    """Результат агрегации: таблица по группам плюс общие итоги."""

    table: pd.DataFrame
    totals: Totals
    group_by: str
    label: str
    warnings: list[str] = field(default_factory=list)


def resolve_group_by(name: str) -> tuple[str, str]:
    """Принимает русское или английское имя колонки. Возвращает (колонка, подпись)."""
    key = normalize_text(name).lower().replace("ё", "е")
    if key not in GROUP_ALIASES:
        allowed = ", ".join(sorted(GROUP_ALIASES))
        raise ValueError(f"Нельзя группировать по «{name}». Доступно: {allowed}")
    column = GROUP_ALIASES[key]
    return column, name


def _empty_table() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "group": pd.Series(dtype="object"),
            "orders": pd.Series(dtype="int64"),
            "quantity": pd.Series(dtype="int64"),
            "revenue": pd.Series(dtype="float64"),
            "avg_order_value": pd.Series(dtype="float64"),
            "share_percent": pd.Series(dtype="float64"),
        }
    )


def summarize(df: pd.DataFrame, group_by: str) -> Summary:
    """Считает количество заказов, штуки и выручку по группам.

    Пустая таблица не ошибка: возвращается пустая сводка с нулевыми итогами,
    чтобы отчёт всё равно построился и клиент увидел «данных нет».
    """
    column, label = resolve_group_by(group_by)
    if column not in CANONICAL_COLUMNS:
        raise ValueError(f"Колонки «{column}» нет в данных")

    warnings: list[str] = []

    if df.empty:
        return Summary(_empty_table(), Totals(), column, label, warnings)

    work = df.copy()
    work["revenue"] = (work["quantity"] * work["price"]).round(2)
    work["group"] = work[column].map(normalize_text)

    if (work["group"] == "").any():
        empty_groups = int((work["group"] == "").sum())
        warnings.append(f"Строк с пустым «{label}»: {empty_groups} — отнесены к «(не указано)»")
        work.loc[work["group"] == "", "group"] = "(не указано)"

    grouped = (
        work.groupby("group", dropna=False)
        .agg(
            orders=("revenue", "size"),
            quantity=("quantity", "sum"),
            revenue=("revenue", "sum"),
        )
        .reset_index()
    )

    grouped["orders"] = grouped["orders"].astype("int64")
    grouped["quantity"] = grouped["quantity"].astype("int64")
    grouped["revenue"] = grouped["revenue"].round(2)
    grouped["avg_order_value"] = (grouped["revenue"] / grouped["orders"]).round(2)

    total_revenue = float(grouped["revenue"].sum())
    if total_revenue > 0:
        grouped["share_percent"] = (grouped["revenue"] / total_revenue * 100).round(2)
    else:
        grouped["share_percent"] = 0.0

    grouped = grouped.sort_values(
        ["revenue", "orders"], ascending=[False, False], kind="stable"
    ).reset_index(drop=True)

    totals = Totals(
        orders=int(grouped["orders"].sum()),
        quantity=int(grouped["quantity"].sum()),
        revenue=round(total_revenue, 2),
        avg_order_value=round(total_revenue / int(grouped["orders"].sum()), 2)
        if grouped["orders"].sum()
        else 0.0,
        groups=len(grouped),
    )

    return Summary(grouped[["group", *SUM_COLUMNS, "avg_order_value", "share_percent"]],
                   totals, column, label, warnings)
