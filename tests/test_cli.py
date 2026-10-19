"""Тесты CLI: успешный путь, коды возврата и сообщения об ошибках."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SAMPLE = ROOT / "sample_data"


def run(*args: str) -> subprocess.CompletedProcess[str]:
    """Запускает main.py так же, как его запустит пользователь.

    PYTHONIOENCODING не задаём специально: CLI должен сам перевести консоль
    в UTF-8, иначе на русской Windows он упадёт на символе ₽.
    """
    return subprocess.run(
        [sys.executable, str(ROOT / "main.py"), *args],
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


class TestSuccess:
    def test_builds_report_from_three_files(self, tmp_path: Path) -> None:
        out = tmp_path / "report.xlsx"
        result = run(
            str(SAMPLE / "sales_january.csv"),
            str(SAMPLE / "sales_february.csv"),
            str(SAMPLE / "sales_march.csv"),
            "-o",
            str(out),
        )
        assert result.returncode == 0, result.stderr
        assert out.exists()
        # cp1251-файл должен читаться без указания кодировки.
        assert "cp1251" in result.stdout
        assert "ИТОГО" in result.stdout
        assert "Готово:" in result.stdout

    def test_group_by_russian_name(self, tmp_path: Path) -> None:
        result = run(
            str(SAMPLE / "sales_january.csv"),
            "-o",
            str(tmp_path / "r.xlsx"),
            "--group-by",
            "Менеджер",
        )
        assert result.returncode == 0, result.stderr
        assert "Анна" in result.stdout

    def test_no_console_is_quiet(self, tmp_path: Path) -> None:
        result = run(
            str(SAMPLE / "sales_january.csv"),
            "-o",
            str(tmp_path / "q.xlsx"),
            "--no-console",
        )
        assert result.returncode == 0
        assert "ИТОГО" not in result.stdout
        assert "Готово:" in result.stdout

    def test_explicit_correct_encoding(self, tmp_path: Path) -> None:
        result = run(
            str(SAMPLE / "sales_february.csv"),
            "-o",
            str(tmp_path / "e.xlsx"),
            "--encoding",
            "cp1251",
        )
        assert result.returncode == 0, result.stderr
        # Группировка по умолчанию — «Товар», поэтому проверяем товар.
        assert "Кофе в зёрнах" in result.stdout
        assert "cp1251" in result.stdout


class TestErrors:
    def test_missing_file(self) -> None:
        result = run("nope.csv")
        assert result.returncode == 2
        assert "Не найдены файлы" in result.stderr

    def test_directory_instead_of_file(self, tmp_path: Path) -> None:
        result = run(str(SAMPLE))
        assert result.returncode == 2
        assert "не файлы" in result.stderr.lower()

    def test_unknown_group_column(self, tmp_path: Path) -> None:
        result = run(str(SAMPLE / "sales_january.csv"), "--group-by", "выручка")
        assert result.returncode == 2
        assert "Нельзя группировать" in result.stderr
        assert "product" in result.stderr

    def test_wrong_encoding_gives_clear_error(self, tmp_path: Path) -> None:
        result = run(
            str(SAMPLE / "sales_january.csv"),
            "-o",
            str(tmp_path / "bad.xlsx"),
            "--encoding",
            "cp1251",
        )
        assert result.returncode == 1
        assert "испорченный текст" in result.stderr
        assert "utf-8" in result.stderr

    def test_file_without_required_columns(self, tmp_path: Path) -> None:
        broken = tmp_path / "broken.csv"
        broken.write_text("date,manager\n2025-01-01,Анна\n", encoding="utf-8")

        result = run(str(broken), "-o", str(tmp_path / "b.xlsx"))
        assert result.returncode == 1
        assert "не хватает колонок" in result.stderr
        # Важно видеть, что именно нашлось, иначе непонятно, что не так с файлом.
        assert "Найдены колонки: date, manager" in result.stderr

    def test_report_saved_as_csv_gives_hint(self, tmp_path: Path) -> None:
        """Клиент может сохранить сам отчёт в CSV и подать его на вход."""
        report_csv = tmp_path / "report.csv"
        report_csv.write_text(
            "Отчёт по продажам;;;;\nФайлов обработано;1;;;\n", encoding="utf-8"
        )

        result = run(str(report_csv), "-o", str(tmp_path / "h.xlsx"))
        assert result.returncode == 1
        assert "не выгрузка продаж" in result.stderr

    def test_unknown_encoding_name(self, tmp_path: Path) -> None:
        result = run(
            str(SAMPLE / "sales_january.csv"),
            "-o",
            str(tmp_path / "x.xlsx"),
            "--encoding",
            "klingon",
        )
        assert result.returncode == 1
        assert "не читается как klingon" in result.stderr


class TestWildcards:
    def test_expands_glob(self, tmp_path: Path) -> None:
        out = tmp_path / "w.xlsx"
        result = run(f"{SAMPLE}/*.csv", "-o", str(out))
        assert result.returncode == 0, result.stderr
        assert "Файлов: 3" in result.stdout

    def test_glob_with_no_matches(self) -> None:
        result = run("sample_data/*.xlsx")
        assert result.returncode == 2
        assert "ничего не найдено" in result.stderr

    def test_same_file_twice_is_not_counted_twice(self, tmp_path: Path) -> None:
        result = run(
            str(SAMPLE / "sales_january.csv"),
            str(SAMPLE / "sales_january.csv"),
            "-o",
            str(tmp_path / "dup.xlsx"),
        )
        assert result.returncode == 0, result.stderr
        assert "Файлов: 1" in result.stdout


class TestHelp:
    def test_help_lists_all_options(self) -> None:
        result = run("--help")
        assert result.returncode == 0
        for option in ("--output", "--group-by", "--encoding", "--no-console"):
            assert option in result.stdout
