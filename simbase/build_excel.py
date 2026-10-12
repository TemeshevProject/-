#!/usr/bin/env python3
"""Разбор выгрузок SimBase в отдельные Excel-файлы.

Каждая запись — девять полей:
номер, наименование процесса, тема, компания, филиал,
дата создания, инициатор, бухгалтер, должность.

Записи разделены знаком «;», после которого начинается номер следующей
записи. Точка с запятой внутри темы остаётся частью темы.
"""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import xlsxwriter

COLUMNS = [
    "Номер",
    "Наименование процесса",
    "Тема",
    "Компания",
    "Филиал",
    "Дата создания",
    "Инициатор",
    "Бухгалтер",
    "Должность",
]

SRC = Path("/tmp/simbase")
OUT = Path(__file__).resolve().parent


def parse_records(text: str) -> tuple[list[list], list[str]]:
    text = text.strip()
    if text.startswith("["):
        text = text[1:]
    text = text.rstrip()
    if text.endswith("]"):
        text = text[:-1]

    i = 0
    n = len(text)
    records: list[list] = []
    problems: list[str] = []

    def skip_ws(pos: int) -> int:
        while pos < n and text[pos] in " \t\r\n":
            pos += 1
        return pos

    def read_string(pos: int) -> tuple[str, int]:
        if pos >= n or text[pos] != '"':
            raise ValueError(f"ожидалась кавычка: {text[pos:pos + 60]!r}")
        pos += 1
        chars: list[str] = []
        while pos < n:
            char = text[pos]
            if char == "\\":
                if pos + 1 >= n:
                    chars.append("\\")
                    pos += 1
                    break
                nxt = text[pos + 1]
                chars.append({"n": "\n", "r": "\r", "t": "\t"}.get(nxt, nxt))
                pos += 2
                continue
            if char == '"':
                value = "".join(chars).replace("\r\n", "\n").replace("\r", "\n")
                return value, pos + 1
            chars.append(char)
            pos += 1
        raise ValueError("строка не закрыта кавычкой")

    while True:
        i = skip_ws(i)
        if i >= n:
            break
        match = re.match(r"(\d+)", text[i:])
        if not match:
            problems.append(f"нет номера на позиции {i}: {text[i:i + 80]!r}")
            break
        number = int(match.group(1))
        i += match.end()
        i = skip_ws(i)
        if i >= n or text[i] != ",":
            problems.append(f"после номера {number} нет запятой")
            break
        i += 1

        fields: list[str] = []
        broken = False
        for field_index in range(8):
            i = skip_ws(i)
            try:
                value, i = read_string(i)
            except ValueError as exc:
                problems.append(f"запись {number}, поле {field_index + 2}: {exc}")
                broken = True
                break
            fields.append(value)
            i = skip_ws(i)
            if field_index < 7:
                if i < n and text[i] == ",":
                    i += 1
                else:
                    problems.append(
                        f"запись {number}: нет запятой после поля {field_index + 2}"
                    )
                    broken = True
                    break
        if broken:
            break

        while True:
            i = skip_ws(i)
            if i >= n or text[i] != ";":
                break
            after = skip_ws(i + 1)
            if re.match(r"\d+\s*,", text[after:]):
                i = after
                break
            # Точка с запятой не открывает следующую запись — это часть темы.
            fields[1] += ";"
            i += 1

        records.append([number, *fields])

    return records, problems


def write_xlsx(path: Path, records: list[list]) -> None:
    workbook = xlsxwriter.Workbook(path)
    sheet = workbook.add_worksheet("Данные")
    header = workbook.add_format(
        {
            "bold": True,
            "bg_color": "#1F4E79",
            "font_color": "#FFFFFF",
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
            "border": 1,
        }
    )
    text = workbook.add_format({"valign": "top", "text_wrap": True})
    number = workbook.add_format({"valign": "top", "align": "center"})
    date_fmt = workbook.add_format(
        {"valign": "top", "align": "center", "num_format": "dd.mm.yyyy"}
    )

    for col, title in enumerate(COLUMNS):
        sheet.write(0, col, title, header)

    for row_index, record in enumerate(records, start=1):
        sheet.write_number(row_index, 0, record[0], number)
        for col in range(1, 5):
            sheet.write(row_index, col, record[col], text)
        raw_date = record[5].strip()
        if raw_date:
            try:
                parsed = datetime.strptime(raw_date, "%d.%m.%Y")
            except ValueError:
                sheet.write(row_index, 5, raw_date, text)
            else:
                sheet.write_datetime(row_index, 5, parsed, date_fmt)
        else:
            sheet.write(row_index, 5, "", text)
        for col in range(6, 9):
            sheet.write(row_index, col, record[col], text)

    widths = [10, 28, 70, 36, 42, 16, 28, 28, 36]
    for col, width in enumerate(widths):
        sheet.set_column(col, col, width)
    sheet.set_row(0, 24)
    sheet.freeze_panes(1, 0)
    sheet.autofilter(0, 0, max(len(records), 1), len(COLUMNS) - 1)
    workbook.close()


def main() -> None:
    sources = sorted(SRC.glob("*.txt"))
    if not sources:
        raise SystemExit(f"нет исходных файлов в {SRC}")
    for source in sources:
        text = source.read_text(encoding="utf-8-sig")
        records, problems = parse_records(text)
        numbers = [row[0] for row in records]
        expected = list(range(1, len(records) + 1))
        gaps = []
        if numbers != expected:
            missing = [n for n in expected if n not in set(numbers)]
            extra = [n for n in numbers if numbers.count(n) > 1]
            gaps.append(
                f"номера не идут 1..{len(records)} подряд, пропуски {missing[:8]}, повторы {extra[:8]}"
            )
        bad_dates = []
        for row in records:
            raw = row[5].strip()
            if not raw:
                continue
            try:
                datetime.strptime(raw, "%d.%m.%Y")
            except ValueError:
                bad_dates.append((row[0], raw))
        out = OUT / f"{source.stem}.xlsx"
        write_xlsx(out, records)
        print(
            f"{source.name}: записей {len(records)}, проблем {len(problems)}, "
            f"кривых дат {len(bad_dates)} -> {out.name}"
        )
        for item in problems[:5]:
            print("  ", item)
        for item in gaps:
            print("  ", item)
        for num, raw in bad_dates[:5]:
            print(f"   дата {num}: {raw!r}")


if __name__ == "__main__":
    main()
