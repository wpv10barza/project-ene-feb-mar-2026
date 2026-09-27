from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import List, Tuple

import pandas as pd
from openpyxl import Workbook


_CODEBLOCK_RE = re.compile(r"```(?:tsv|csv)?\s*\n(.*?)\n```", flags=re.DOTALL | re.IGNORECASE)


def extract_codeblock(text: str) -> str:
    """Si el texto contiene ```...```, devuelve el contenido interno; si no, devuelve text."""
    if not text:
        return ""
    m = _CODEBLOCK_RE.search(text)
    return (m.group(1).strip() if m else text.strip())


def guess_delimiter(sample_line: str) -> str:
    if "\t" in sample_line:
        return "\t"
    if ";" in sample_line and sample_line.count(";") >= sample_line.count(","):
        return ";"
    return ","


def parse_table(table_text: str) -> Tuple[List[List[str]], str]:
    lines = [ln for ln in table_text.splitlines() if ln.strip()]
    if not lines:
        return [], "\t"
    delim = guess_delimiter(lines[0])
    rows = [[cell.strip() for cell in row] for row in csv.reader(lines, delimiter=delim)]
    return rows, delim


def rows_to_dataframe(rows: List[List[str]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    header = rows[0]
    data = rows[1:] if len(rows) > 1 else []
    # Asegura longitudes
    norm = []
    for r in data:
        if len(r) < len(header):
            r = r + [""] * (len(header) - len(r))
        elif len(r) > len(header):
            r = r[: len(header)]
        norm.append(r)
    return pd.DataFrame(norm, columns=header)


def write_xlsx(rows: List[List[str]], out_path: Path, sheet_name: str = "resultado") -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    for r in rows:
        ws.append(r)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
