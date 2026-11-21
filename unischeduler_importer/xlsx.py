"""Read first Excel OOXML worksheet with Python stdlib (ZIP + ElementTree).

No preprocessing loss, no dependency on openpyxl. Supports shared/inline/rich
strings, numeric/boolean/date-formatted cells, cached formula results, and
merged-cell metadata. This is a read-only .xlsx reader, not an Excel engine.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET
from zipfile import ZipFile, BadZipFile
import posixpath
import re

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
P = "{http://schemas.openxmlformats.org/package/2006/relationships}"
MAX_XLSX = 24 * 1024 * 1024
MAX_UNCOMPRESSED = 90 * 1024 * 1024
MAX_CELLS = 100_000


@dataclass
class Worksheet:
    name: str
    rows: dict[int, dict[int, str]]
    merged: list[str]

    def get(self, row: int, col: int) -> str:
        return self.rows.get(row, {}).get(col, "")


def _safe_xml(blob: bytes) -> ET.Element:
    if b"<!DOCTYPE" in blob.upper() or b"<!ENTITY" in blob.upper():
        raise ValueError("XML entities/DOCTYPE are forbidden in workbook files")
    return ET.fromstring(blob)


def _column_number(address: str) -> int:
    letters = re.match(r"[A-Z]+", address)
    if not letters:
        raise ValueError(f"Bad cell address {address}")
    result = 0
    for letter in letters.group(0):
        result = result * 26 + ord(letter) - 64
    return result


def _shared_strings(z: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = _safe_xml(z.read("xl/sharedStrings.xml"))
    return ["".join(n.text or "" for n in si.iter(M + "t")) for si in root.findall(M+"si")]


def _date_styles(z: ZipFile) -> set[int]:
    if "xl/styles.xml" not in z.namelist():
        return set()
    root = _safe_xml(z.read("xl/styles.xml"))
    formats = root.find(M+"numFmts")
    date_fmts = {14, 15, 16, 17, 22, 27, 30, 36, 45, 46, 47, 50, 57}
    if formats is not None:
        for fmt in formats.findall(M+"numFmt"):
            code = fmt.get("formatCode", "").lower()
            code = re.sub(r'"[^\"]*"|\[[^\]]*\]|\\.', "", code)
            if re.search(r"[ymdhsa]", code):
                date_fmts.add(int(fmt.get("numFmtId", "0")))
    xfs = root.find(M+"cellXfs")
    if xfs is None:
        return set()
    return {i for i,xf in enumerate(xfs.findall(M+"xf")) if int(xf.get("numFmtId", "0")) in date_fmts}


def _text_of_cell(c: ET.Element, shared: list[str], date_styles: set[int], is_1904: bool) -> str:
    kind = c.get("t", "n")
    inline = c.find(M+"is")
    v = c.find(M+"v")
    if kind == "inlineStr" and inline is not None:
        return "".join(n.text or "" for n in inline.iter(M+"t"))
    if v is None or v.text is None:
        return ""
    raw = v.text
    if kind == "s":
        return shared[int(raw)]
    if kind == "b":
        return "TRUE" if raw == "1" else "FALSE"
    if kind == "e":
        return raw
    if kind == "d":
        return raw
    if kind in ("str", "inlineStr"):
        return raw
    if kind == "n":
        if int(c.get("s", "0")) in date_styles:
            base = datetime(1904,1,1) if is_1904 else datetime(1899,12,30)
            dt = base + timedelta(days=float(raw))
            return dt.strftime("%d.%m.%Y" if dt.time().isoformat() == "00:00:00" else "%d.%m.%Y %H:%M:%S")
        try:
            number = float(raw)
            if number.is_integer() and abs(number) < 1e15:
                return str(int(number))
        except (OverflowError, ValueError):
            pass
        return raw
    return raw


def read_first_sheet(path: str | Path) -> Worksheet:
    path = Path(path)
    if path.stat().st_size > MAX_XLSX:
        raise ValueError("Workbook exceeds file size limit")
    try:
        with ZipFile(path) as z:
            if sum(item.file_size for item in z.infolist()) > MAX_UNCOMPRESSED:
                raise ValueError("Workbook exceeds decompressed size limit")
            wb = _safe_xml(z.read("xl/workbook.xml"))
            sheets = wb.find(M+"sheets")
            if sheets is None or not list(sheets):
                raise ValueError("Workbook contains no worksheets")
            first = list(sheets)[0]
            name = first.get("name", "Sheet1")
            rel_id = first.get(R+"id")
            relationships = _safe_xml(z.read("xl/_rels/workbook.xml.rels"))
            targets = {rel.get("Id"): rel.get("Target") for rel in relationships.findall(P+"Relationship")}
            target = targets.get(rel_id)
            if not target:
                raise ValueError("Missing relationship for the first worksheet")
            target_path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
            if not target_path.startswith("xl/") or ".." in PurePosixPath(target_path).parts:
                raise ValueError("Invalid worksheet path")
            root = _safe_xml(z.read(target_path))
            shared = _shared_strings(z)
            dates = _date_styles(z)
            props = wb.find(M+"workbookPr")
            is_1904 = props is not None and props.get("date1904") in ("1", "true")
            rows: dict[int, dict[int, str]] = {}
            count = 0
            for r in root.findall(".//"+M+"sheetData/"+M+"row"):
                index = int(r.get("r", "0"))
                if index > MAX_CELLS:
                    raise ValueError("Worksheet exceeds row limit")
                for c in r.findall(M+"c"):
                    count += 1
                    if count > MAX_CELLS:
                        raise ValueError("Too many workbook cells")
                    col = _column_number(c.get("r", ""))
                    val = _text_of_cell(c, shared, dates, is_1904)
                    if val != "":
                        rows.setdefault(index,{})[col] = val
            merges = [el.get("ref", "") for el in root.findall(".//"+M+"mergeCells/"+M+"mergeCell")]
            return Worksheet(name, rows, merges)
    except (KeyError, BadZipFile, ET.ParseError) as ex:
        raise ValueError(f"Invalid or unsupported XLSX document: {ex}") from ex
