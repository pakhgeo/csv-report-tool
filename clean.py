"""Загрузка и очистка «грязных» выгрузок из Excel и 1С.

Модуль отвечает только за чтение и приведение данных к предсказуемому виду.
Агрегация и вывод отчёта — в report.py.
"""

from __future__ import annotations

import codecs
import io
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Sequence

import pandas as pd

CANONICAL_COLUMNS = ("date", "manager", "product", "city", "quantity", "price")

# Заголовки приходят то по-русски, то по-английски, иногда с «Кол-во».
HEADER_ALIASES: dict[str, set[str]] = {
    "date": {"date", "дата"},
    "manager": {"manager", "менеджер"},
    "product": {"product", "товар"},
    "city": {"city", "город"},
    "quantity": {"quantity", "количество", "кол-во", "колво"},
    "price": {"price", "цена", "стоимость"},
}

DATE_FORMATS = ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y")

_WHITESPACE = re.compile(r"[\s\u00a0\u202f]+")
_ALL_ALIASES = {alias for aliases in HEADER_ALIASES.values() for alias in aliases}
_NUMBER_NOISE = re.compile(r"[^\d,.\-]")
_EMPTY_VALUES = {"", "-", "—", "–", "nan", "none", "n/a", "нет"}


@dataclass
class CleanStats:
    """Что было отброшено и почему — эти цифры идут в отчёт клиенту."""

    files: int = 0
    input_rows: int = 0
    output_rows: int = 0
    blank_rows: int = 0
    bad_date: int = 0
    bad_quantity: int = 0
    bad_price: int = 0
    duplicates: int = 0
    ignored_columns: list[str] = field(default_factory=list)
    encodings: dict[str, str] = field(default_factory=dict)

    def dropped_rows(self) -> int:
        return (
            self.blank_rows
            + self.bad_date
            + self.bad_quantity
            + self.bad_price
            + self.duplicates
        )


def _looks_like_mojibake(text: str) -> bool:
    """Признак неверно определённой кодировки: кириллица, прочитанная как Latin-1.

    UTF-8, прочитанный как cp1251, даёт характерные пары «Р» + строчная буква
    («РљРѕС„Рµ»). Если таких последовательностей много, текст почти наверняка битый.
    """
    patterns = ("Р°", "Р±", "РІ", "Рі", "Рґ", "Рµ", "Р¶", "Р·", "Рє", "Р»", "Рј",
                "РЅ", "Рѕ", "Рї", "СЂ", "СЃ", "С‚", "Сѓ", "С„", "С…", "С†", "С‡",
                "С€", "С‰", "СЋ", "СЏ")
    hits = sum(text.count(pattern) for pattern in patterns)
    return hits >= 5


def read_text(path: str | Path, encoding: str | None = None) -> tuple[str, str]:
    """Читает файл, определяя кодировку. Возвращает текст и название кодировки.

    Если кодировка задана явно, автоопределение не используется — это нужно,
    когда файл битый или когда клиент знает, в чём его сохранил Excel.
    """
    raw = Path(path).read_bytes()

    if encoding:
        try:
            text = raw.decode(encoding)
        except (UnicodeDecodeError, LookupError) as error:
            raise ValueError(
                f"Файл {Path(path).name} не читается как {encoding}: {error}"
            ) from error

        if encoding.lower().replace("-", "") != "utf8" and _looks_like_mojibake(text):
            raise ValueError(
                f"Файл {Path(path).name} при чтении как {encoding} выглядит как "
                "испорченный текст. Похоже, кодировка другая — попробуйте "
                "utf-8 или не указывайте --encoding."
            )
        return text, encoding

    if raw.startswith(codecs.BOM_UTF8):
        return raw.decode("utf-8-sig"), "utf-8-sig"
    for candidate in ("utf-8", "cp1251"):
        try:
            return raw.decode(candidate), candidate
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Не удалось определить кодировку файла: {path}")


def detect_separator(text: str) -> str:
    """Определяет разделитель по первой строке: запятая, точка с запятой или таб."""
    header = text.splitlines()[0] if text.splitlines() else ""
    counts = {",": header.count(","), ";": header.count(";"), "\t": header.count("\t")}
    best = max(counts, key=counts.get)
    return best if counts[best] > 0 else ","


def normalize_text(value: object) -> str:
    """Убирает пробелы по краям и схлопывает внутренние (включая табы и nbsp)."""
    if value is None:
        return ""
    if not isinstance(value, str) and pd.isna(value):
        return ""
    return _WHITESPACE.sub(" ", str(value)).strip()


def parse_number(value: object) -> float | None:
    """Разбирает число в русском или английском формате. None — если числа нет.

    Понимает «1250,50», «1 250,50», «1.250,50», «1,250.50», «1250.00 руб.».
    """
    if value is None:
        return None
    if not isinstance(value, str) and pd.isna(value):
        return None

    text = str(value).strip()
    if text.lower() in _EMPTY_VALUES:
        return None

    text = _WHITESPACE.sub("", text)
    text = _NUMBER_NOISE.sub("", text)
    # «1250.00 руб.» после чистки даёт «1250.00.» — срезаем висящий разделитель.
    text = text.strip(".,")
    if text in _EMPTY_VALUES or text in {"-", ".", ","}:
        return None

    if "," in text and "." in text:
        # Десятичный тот разделитель, который стоит последним.
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        # Одна запятая и 1–2 знака после неё — это дробная часть, иначе разделитель тысяч.
        if text.count(",") == 1 and len(text.split(",")[1]) in (1, 2):
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")

    try:
        return float(text)
    except ValueError:
        return None


def parse_date(value: object) -> pd.Timestamp | None:
    """Разбирает дату в ISO или русском формате. None — если дата не распознана."""
    text = normalize_text(value)
    if text.lower() in _EMPTY_VALUES:
        return None
    for fmt in DATE_FORMATS:
        try:
            return pd.Timestamp(datetime.strptime(text, fmt).date())
        except ValueError:
            continue
    return None


def standardize_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Приводит заголовки к каноническим именам. Возвращает таблицу и лишние колонки."""
    mapping: dict[str, str] = {}
    ignored: list[str] = []

    for column in df.columns:
        key = normalize_text(column).lower()
        canonical = next(
            (name for name, aliases in HEADER_ALIASES.items() if key in aliases),
            None,
        )
        if canonical is None:
            ignored.append(str(column))
        elif canonical in mapping.values():
            raise ValueError(f"Колонка «{column}» дублирует уже найденную «{canonical}»")
        else:
            mapping[column] = canonical

    missing = [name for name in CANONICAL_COLUMNS if name not in mapping.values()]
    if missing:
        found = ", ".join(str(column) for column in df.columns) or "ничего"
        known = sum(
            1 for column in df.columns if normalize_text(column).lower() in _ALL_ALIASES
        )
        hint = ""
        # Ни одной знакомой колонки — это, скорее всего, не выгрузка продаж.
        if known == 0:
            hint = (
                " Похоже, это не выгрузка продаж, а готовый отчёт или другой документ: "
                "нужен файл, где на каждую продажу своя строка."
            )
        elif len(df.columns) <= 2:
            hint = " Колонок слишком мало для выгрузки продаж."
        raise ValueError(
            f"В файле не хватает колонок: {', '.join(missing)}. "
            f"Найдены колонки: {found}.{hint}"
        )

    return df.rename(columns=mapping)[list(CANONICAL_COLUMNS)], ignored


def read_sales_file(
    path: str | Path, encoding: str | None = None
) -> tuple[pd.DataFrame, str, str]:
    """Читает один CSV. Возвращает таблицу, кодировку и разделитель."""
    text, used_encoding = read_text(path, encoding)
    separator = detect_separator(text)

    # rstrip — чтобы завершающий перевод строки не превратился в пустую запись.
    text = text.rstrip("\r\n")
    df = pd.read_csv(
        io.StringIO(text),
        sep=separator,
        dtype=str,
        keep_default_na=False,
        skip_blank_lines=False,
    )
    df.columns = [normalize_text(column) for column in df.columns]

    standardized, ignored = standardize_columns(df)
    if ignored:
        standardized.attrs["ignored_columns"] = ignored
    return standardized, used_encoding, separator


def load_sales(
    paths: Sequence[str | Path], encoding: str | None = None
) -> tuple[pd.DataFrame, CleanStats]:
    """Читает несколько файлов и склеивает их в одну таблицу."""
    frames: list[pd.DataFrame] = []
    stats = CleanStats()
    ignored: list[str] = []

    for path in paths:
        df, used_encoding, _ = read_sales_file(path, encoding)
        frames.append(df)
        stats.files += 1
        stats.encodings[Path(path).name] = used_encoding
        for column in df.attrs.get("ignored_columns", []):
            if column not in ignored:
                ignored.append(column)

    stats.ignored_columns = ignored
    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(
        columns=list(CANONICAL_COLUMNS)
    )
    return combined, stats


def clean_sales(
    df: pd.DataFrame, stats: CleanStats | None = None
) -> tuple[pd.DataFrame, CleanStats]:
    """Очищает данные. Каждый шаг считает отброшенные строки отдельно."""
    stats = stats or CleanStats()
    work = df.copy().reset_index(drop=True)
    stats.input_rows = len(work)

    for column in ("manager", "product", "city"):
        work[column] = work[column].map(normalize_text)

    # 1. Полностью пустые и пробельные строки.
    blank_mask = work[list(CANONICAL_COLUMNS)].apply(
        lambda column: column.map(lambda value: normalize_text(value) == "")
    ).all(axis=1)
    stats.blank_rows = int(blank_mask.sum())
    work = work[~blank_mask]

    # 2. Дата, количество и цена — в типы.
    work["date"] = work["date"].map(parse_date)
    work["quantity"] = work["quantity"].map(parse_number)
    work["price"] = work["price"].map(parse_number)

    bad_date = work["date"].isna()
    stats.bad_date = int(bad_date.sum())
    work = work[~bad_date]

    bad_quantity = work["quantity"].isna() | (work["quantity"] <= 0)
    stats.bad_quantity = int(bad_quantity.sum())
    work = work[~bad_quantity]

    bad_price = work["price"].isna() | (work["price"] <= 0)
    stats.bad_price = int(bad_price.sum())
    work = work[~bad_price]

    # 3. Дубли — уже после нормализации пробелов, иначе « Анна » и «Анна» не схлопнутся.
    duplicate_mask = work.duplicated(keep="first")
    stats.duplicates = int(duplicate_mask.sum())
    work = work[~duplicate_mask]

    work["quantity"] = work["quantity"].astype("int64")
    work = work.sort_values(["date", "manager", "product"], kind="stable")
    work = work.reset_index(drop=True)

    stats.output_rows = len(work)
    return work, stats


def prepare(
    paths: Sequence[str | Path], encoding: str | None = None
) -> tuple[pd.DataFrame, CleanStats]:
    """Читает и очищает набор файлов за один вызов."""
    raw, stats = load_sales(paths, encoding)
    return clean_sales(raw, stats)
