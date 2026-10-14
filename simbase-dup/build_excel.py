#!/usr/bin/env python3
"""Разбор выгрузок «ДУП SimBase» в отдельные Excel-файлы.

Запись — девять полей. Записи разделены знаком «;» только если сразу после
него начинается номер следующей записи. Точка с запятой внутри темы остаётся
в теме.

У согласующих восьмое и девятое поля — человек и его должность.
У авторов восьмое поле — должность инициатора, девятое — числовой код,
без кавычек.
"""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import xlsxwriter

SRC = Path("/tmp/simbase2")
OUT = Path(__file__).resolve().parent

APPROVER_COLUMNS = [
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
AUTHOR_COLUMNS = [
    "Номер",
    "Наименование процесса",
    "Тема",
    "Компания",
    "Филиал",
    "Дата создания",
    "Инициатор",
    "Должность",
    "Код",
]


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

    def read_raw(pos: int) -> tuple[str, int]:
        start = pos
        while pos < n and text[pos] not in ",;\r\n":
            pos += 1
        return text[start:pos].strip(), pos

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
                if i < n and text[i] == '"':
                    value, i = read_string(i)
                else:
                    value, i = read_raw(i)
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
                        f"запись {number}: нет запятой после поля {field_index + 2}: {text[i:i + 40]!r}"
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
            fields[1] += ";"
            i += 1

        records.append([number, *fields])

    return records, problems


def write_xlsx(path: Path, records: list[list], columns: list[str], code_column: bool) -> None:
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

    for col, title in enumerate(columns):
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
        sheet.write(row_index, 6, record[6], text)
        sheet.write(row_index, 7, record[7], text)
        if code_column and record[8].isdigit():
            sheet.write_number(row_index, 8, int(record[8]), number)
        else:
            sheet.write(row_index, 8, record[8], text)

    widths = [10, 28, 70, 36, 42, 16, 28, 42, 22]
    for col, width in enumerate(widths):
        sheet.set_column(col, col, width)
    sheet.set_row(0, 24)
    sheet.freeze_panes(1, 0)
    sheet.autofilter(0, 0, max(len(records), 1), len(columns) - 1)
    workbook.close()


def main() -> None:
    sources = sorted(SRC.glob("*.txt"))
    if not sources:
        raise SystemExit(f"нет исходных файлов в {SRC}")
    for source in sources:
        text = source.read_text(encoding="utf-8-sig")
        records, problems = parse_records(text)
        numbers = [row[0] for row in records]
        gaps = []
        if numbers != list(range(1, len(records) + 1)):
            gaps.append(f"номера не идут подряд 1..{len(records)}")
        author = "авторы" in source.stem
        codes = [row[8] for row in records]
        code_ok = all(item.isdigit() for item in codes)
        text_ok = all(not item.isdigit() for item in codes)
        bad_dates = []
        for row in records:
            raw = row[5].strip()
            if not raw:
                continue
            try:
                datetime.strptime(raw, "%d.%m.%Y")
            except ValueError:
                bad_dates.append((row[0], raw))
        semis = sum(row[2].count(";") for row in records)
        columns = AUTHOR_COLUMNS if author else APPROVER_COLUMNS
        out = OUT / f"{source.stem}.xlsx"
        write_xlsx(out, records, columns, code_column=author)
        print(
            f"{source.name}: записей {len(records)}, проблем {len(problems)}, "
            f"дат с ошибкой {len(bad_dates)}, ';' в теме {semis}, "
            f"код-число {code_ok}, должность-текст {text_ok} -> {out.name}"
        )
        for item in problems[:8]:
            print("  ", item)
        for item in gaps:
            print("  ", item)
        for num, raw in bad_dates[:5]:
            print(f"   дата {num}: {raw!r}")


if __name__ == "__main__":
    main()
