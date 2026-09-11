#!/usr/bin/env python3
"""Merge all Excel sheets into one list with company, subdivision, position, FIO."""
import re
from pathlib import Path

import pandas as pd

INPUT = Path("/workspace/drive_download/2026.09.03 Список работников группы.xlsx")
OUTPUT = Path("/workspace/Объединенный_список_работников.xlsx")

SKIP_NAME_RE = re.compile(
    r"^(итого|всего|всеого|всего по|итого по|всего:|итого:)\b",
    re.I,
)
def extract_company_from_text(text: str) -> str:
    text = text.replace("\n", " ")
    parent = None
    for pattern in (
        r"ТОО\s*[«\"]([^»\"]+)[»\"]",
        r"Частной компании\s*[«\"]([^»\"]+)[»\"]",
        r"ЧК\s*[«\"]([^»\"]+)[»\"]",
    ):
        m = re.search(pattern, text, re.I)
        if m:
            parent = m.group(1).strip()
            break

    names = [
        q.strip()
        for q in re.findall(r"[«\"]([^»\"]+)[»\"]", text)
        if not re.fullmatch(r"\d+", q.strip())
    ]
    filial = names[-1] if len(names) >= 2 else None
    if parent and filial and filial != parent:
        return f"{parent} ({filial})"
    return parent or filial or (names[0] if names else "")


def extract_company(df: pd.DataFrame) -> str:
    for i in range(min(5, len(df))):
        for val in df.iloc[i].tolist():
            if pd.isna(val):
                continue
            val = str(val)
            if "штатн" in val.lower() or "тоо" in val.lower() or "частн" in val.lower():
                company = extract_company_from_text(val)
                if company:
                    return company
    return ""


def find_columns(header_row: pd.Series) -> tuple[int | None, int | None]:
    name_col = fio_col = None
    for idx, val in header_row.items():
        if pd.isna(val):
            continue
        s = str(val).lower().replace("\n", " ")
        if "ф.и.о" in s or "ф.и.о." in s:
            fio_col = idx
        if "подраздел" in s and "должност" in s:
            name_col = idx
        elif name_col is None and "должност" in s:
            name_col = idx
    return name_col, fio_col


def find_header_row(df: pd.DataFrame) -> int | None:
    for i in range(min(15, len(df))):
        row = df.iloc[i]
        for val in row:
            if pd.isna(val):
                continue
            s = str(val).lower()
            if "ф.и.о" in s:
                return i
    return None


def section_level(num_val) -> int | None:
    if pd.isna(num_val):
        return None
    if isinstance(num_val, str):
        s = num_val.strip().replace(",", ".")
        if not s or s.lower() == "nan":
            return None
        if s.endswith("."):
            parts = [p for p in s.rstrip(".").split(".") if p]
            return len(parts) if parts else 1
        if re.fullmatch(r"\d+(\.\d+)+", s):
            return len(s.split("."))
        if re.fullmatch(r"\d+", s):
            return 1
        return None
    if isinstance(num_val, (int, float)):
        if isinstance(num_val, float) and not num_val.is_integer():
            return len(str(num_val).split("."))
        return 1
    return None


def is_summary_name(name: str) -> bool:
    n = name.strip().lower()
    if not n or n == "nan":
        return True
    if SKIP_NAME_RE.match(n):
        return True
    if "всего по" in n or "итого по" in n:
        return True
    return False


def parse_sheet(sheet_name: str, df: pd.DataFrame) -> list[dict]:
    company = extract_company(df)
    if not company:
        company = sheet_name.replace(" ФИО", "").strip()

    header_i = find_header_row(df)
    if header_i is None:
        return []

    name_col, fio_col = find_columns(df.iloc[header_i])
    if name_col is None or fio_col is None:
        return []

    num_col = df.iloc[header_i].index[0]
    rows_out: list[dict] = []
    subdiv_stack: list[str] = []

    for i in range(header_i + 1, len(df)):
        row = df.iloc[i]
        name = row[name_col]
        fio = row[fio_col]
        num_val = row[num_col]

        name_s = "" if pd.isna(name) else str(name).strip()
        fio_s = "" if pd.isna(fio) else str(fio).strip()

        if is_summary_name(name_s) and not fio_s:
            continue
        if not name_s and not fio_s:
            continue
        if name_s and "*" in name_s and not fio_s:
            continue

        has_fio = bool(fio_s) and fio_s.lower() != "nan"

        if not has_fio and name_s and not is_summary_name(name_s):
            lvl = section_level(num_val)
            if lvl is not None:
                subdiv_stack = subdiv_stack[: lvl - 1] + [name_s]
            continue

        if has_fio:
            position = name_s
            subdivision = " / ".join(subdiv_stack) if subdiv_stack else ""
            rows_out.append(
                {
                    "Подразделение": subdivision,
                    "Должность": position,
                    "ФИО": fio_s,
                    "Компания": company,
                }
            )

    return rows_out


def main():
    xl = pd.ExcelFile(INPUT)
    all_rows: list[dict] = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(INPUT, sheet_name=sheet, header=None)
        all_rows.extend(parse_sheet(sheet, df))

    result = pd.DataFrame(all_rows)
    result = result.fillna("")
    result = result.sort_values(["Компания", "Подразделение", "Должность", "ФИО"])[
        ["Подразделение", "Должность", "ФИО", "Компания"]
    ]
    result.to_excel(OUTPUT, index=False, sheet_name="Все работники")
    print(f"Sheets: {len(xl.sheet_names)}")
    print(f"Rows: {len(result)}")
    print(f"Saved: {OUTPUT}")
    print("\nBy company:")
    print(result.groupby("Компания").size().to_string())


if __name__ == "__main__":
    main()
