"""CLI: собирает отчёт из «грязных» CSV.

Пример:
    python main.py sample_data/*.csv -o output/report.xlsx --group-by Товар
"""

from __future__ import annotations

import argparse
import glob
import sys
from pathlib import Path

from clean import prepare
from output import Report, print_summary, save_excel
from report import GROUP_ALIASES, resolve_group_by, summarize

DEFAULT_OUTPUT = "output/report.xlsx"
DEFAULT_GROUP_BY = "Товар"


def _force_utf8_console() -> None:
    """Переводит консоль в UTF-8.

    На русской Windows кодировка консоли по умолчанию cp1251, и символ ₽
    в ней не существует — печать сводки падала бы с UnicodeEncodeError.
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass

    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
            ctypes.windll.kernel32.SetConsoleCP(65001)
        except (AttributeError, OSError):
            # Консоли может не быть вовсе (вывод в файл) — это не ошибка.
            pass


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Превращает выгрузки из Excel/1С в сводный отчёт.",
        epilog=(
            "Пример: python main.py sample_data/*.csv "
            f"-o {DEFAULT_OUTPUT} --group-by Товар"
        ),
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        metavar="ФАЙЛ",
        help="один или несколько CSV; маски вроде sample_data/*.csv раскрываются сами",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=DEFAULT_OUTPUT,
        help=f"куда сохранить Excel-файл (по умолчанию {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "-g",
        "--group-by",
        default=DEFAULT_GROUP_BY,
        help="по какой колонке группировать: "
        + ", ".join(sorted(GROUP_ALIASES))
        + f" (по умолчанию «{DEFAULT_GROUP_BY}»)",
    )
    parser.add_argument(
        "--encoding",
        default=None,
        help="кодировка входных файлов (по умолчанию определяется автоматически: "
        "utf-8 или cp1251)",
    )
    parser.add_argument(
        "--no-console",
        action="store_true",
        help="не печатать сводку в консоль, только сохранить файл",
    )
    return parser


def expand_inputs(names: list[str]) -> tuple[list[Path], list[str]]:
    """Раскрывает маски вроде sample_data/*.csv самостоятельно.

    PowerShell не раскрывает wildcard'ы перед внешними программами (в отличие
    от bash), поэтому «*.csv» доходил бы до нас буквально и падал с «файл не найден».
    Возвращает найденные пути и шаблоны, по которым ничего не нашлось.
    """
    paths: list[Path] = []
    unmatched: list[str] = []

    for name in names:
        if any(char in name for char in "*?["):
            matches = sorted(glob.glob(name))
            if matches:
                paths.extend(Path(match) for match in matches)
            else:
                unmatched.append(name)
        else:
            paths.append(Path(name))

    return paths, unmatched


def main(argv: list[str] | None = None) -> int:
    _force_utf8_console()
    parser = build_parser()
    args = parser.parse_args(argv)

    paths, unmatched = expand_inputs(args.inputs)

    if unmatched:
        print("По этим маскам ничего не найдено: " + ", ".join(unmatched), file=sys.stderr)
        return 2

    if not paths:
        print("Не указано ни одного файла.", file=sys.stderr)
        return 2

    # Одинаковый файл мог прийти дважды — например, и явно, и по маске.
    seen: set[Path] = set()
    unique_paths: list[Path] = []
    for path in paths:
        resolved = path.resolve()
        if resolved not in seen:
            seen.add(resolved)
            unique_paths.append(path)
    paths = unique_paths

    missing = [str(path) for path in paths if not path.exists()]
    if missing:
        print("Не найдены файлы: " + ", ".join(missing), file=sys.stderr)
        return 2

    not_files = [str(path) for path in paths if not path.is_file()]
    if not_files:
        print("Это не файлы: " + ", ".join(not_files), file=sys.stderr)
        return 2

    try:
        resolve_group_by(args.group_by)
    except ValueError as error:
        print(error, file=sys.stderr)
        return 2

    try:
        data, stats = prepare(paths, args.encoding)
    except ValueError as error:
        print(f"Ошибка в данных: {error}", file=sys.stderr)
        return 1
    except (OSError, PermissionError) as error:
        print(f"Не удалось прочитать файл: {error}", file=sys.stderr)
        return 1

    if data.empty:
        print(
            "После очистки не осталось ни одной корректной строки. "
            "Проверьте формат дат, количества и цены.",
            file=sys.stderr,
        )

    summary = summarize(data, args.group_by)
    report = Report(
        summary=summary,
        data=data,
        stats=stats,
        source_names=[path.name for path in paths],
    )

    try:
        saved = save_excel(report, args.output)
    except PermissionError:
        print(
            f"Не удалось записать «{args.output}»: файл открыт в Excel? "
            "Закройте его и повторите.",
            file=sys.stderr,
        )
        return 1
    except OSError as error:
        print(f"Не удалось записать файл: {error}", file=sys.stderr)
        return 1

    if not args.no_console:
        print_summary(report)

    print(f"Готово: {saved}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
