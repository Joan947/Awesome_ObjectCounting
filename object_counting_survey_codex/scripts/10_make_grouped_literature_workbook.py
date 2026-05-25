"""Create a modality-grouped review workbook from literature_matrix_clean.xlsx."""

from __future__ import annotations

from copy import copy
from pathlib import Path
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parents[1]
WORKSPACE = ROOT.parent
SOURCE_XLSX = ROOT / "data" / "literature_matrix_clean.xlsx"
STYLE_GUIDE_XLSX = WORKSPACE / "Object_Counting.xlsx"
OUTPUT_XLSX = ROOT / "outputs" / "literature_matrix_grouped_by_modality.xlsx"

REQUESTED_COLUMNS = [
    "paper_title",
    "authors",
    "year",
    "venue",
    "doi",
    "paper_url",
    "abstract",
    "index_keywords",
    "citation_count",
    "dataset_or_benchmark",
    "modality",
    "task_category",
    "input_type",
    "output_type",
    "relevance_score",
    "tier",
    "github_url",
]

SHEET_ORDER = [
    "Survey",
    "Image",
    "Video",
    "3D_PointCloud",
    "Remote_Sensing_Aerial",
    "Medical_Microscopy_Cell",
    "Thermal_Event",
    "Multimodal",
    "Other",
]

TAB_COLORS = {
    "Survey": "8064A2",
    "Image": "4F81BD",
    "Video": "C0504D",
    "3D_PointCloud": "9BBB59",
    "Remote_Sensing_Aerial": "F79646",
    "Medical_Microscopy_Cell": "4BACC6",
    "Thermal_Event": "A64D79",
    "Multimodal": "7F7F7F",
    "Other": "BFBFBF",
}

WIDTHS = {
    "paper_title": 42,
    "authors": 34,
    "year": 10,
    "venue": 28,
    "doi": 22,
    "paper_url": 36,
    "abstract": 70,
    "index_keywords": 46,
    "citation_count": 14,
    "dataset_or_benchmark": 28,
    "modality": 22,
    "task_category": 34,
    "input_type": 24,
    "output_type": 24,
    "relevance_score": 15,
    "tier": 16,
    "github_url": 34,
}


def clean_text(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def load_records() -> list[dict[str, object]]:
    wb = load_workbook(SOURCE_XLSX, read_only=True, data_only=True)
    ws = wb["included"] if "included" in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    headers = [clean_text(value) for value in rows[0]]
    records = []
    for row in rows[1:]:
        if not any(value not in (None, "") for value in row):
            continue
        record = {headers[i]: row[i] if i < len(row) else "" for i in range(len(headers))}
        records.append(record)
    return records


def bucket_for(record: dict[str, object]) -> str:
    modality = clean_text(record.get("modality")).lower()
    task = clean_text(record.get("task_category")).lower()
    title = clean_text(record.get("paper_title")).lower()
    text = " ".join([modality, task, title])
    if "survey" in task or "survey" in title or "review" in title:
        return "Survey"
    if any(term in text for term in ["medical", "microscopy", "cell", "bacteria", "histology"]):
        return "Medical_Microscopy_Cell"
    if any(term in text for term in ["remote", "aerial", "uav", "drone", "satellite"]):
        return "Remote_Sensing_Aerial"
    if any(term in text for term in ["video", "tracking", "temporal"]):
        return "Video"
    if any(term in text for term in ["3d", "point_cloud", "point cloud", "rgb-d", "depth", "lidar"]):
        return "3D_PointCloud"
    if any(term in text for term in ["thermal", "event_camera", "event camera", "infrared"]):
        return "Thermal_Event"
    if "multimodal" in text or "vision-language" in text or "open-vocabulary" in text or "text prompt" in text:
        return "Multimodal"
    if "image" in text or not modality:
        return "Image"
    return "Other"


def score_tuple(record: dict[str, object]) -> tuple[int, int, str]:
    tier_rank = {"A_core": 0, "B_important": 1, "C_background": 2}
    tier = clean_text(record.get("tier"))
    try:
        citations = int(float(clean_text(record.get("citation_count")).replace(",", "")))
    except Exception:
        citations = 0
    return (tier_rank.get(tier, 9), -citations, clean_text(record.get("paper_title")).lower())


def style_header(ws, guide_header_cells) -> None:
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    thin = Side(style="thin", color="D9E2F3")
    for col_idx, column_name in enumerate(REQUESTED_COLUMNS, start=1):
        cell = ws.cell(1, col_idx)
        cell.value = column_name
        if guide_header_cells and col_idx <= len(guide_header_cells):
            guide = guide_header_cells[col_idx - 1]
            cell.font = copy(guide.font)
            cell.alignment = copy(guide.alignment)
            cell.border = copy(guide.border)
        cell.font = Font(name="Calibri", size=11, bold=True, color="1F1F1F")
        cell.fill = header_fill
        cell.border = Border(bottom=thin)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def add_sheet(wb: Workbook, sheet_name: str, records: list[dict[str, object]], guide_header_cells) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.sheet_properties.tabColor = TAB_COLORS.get(sheet_name, "BFBFBF")
    style_header(ws, guide_header_cells)
    records = sorted(records, key=score_tuple)
    for row_idx, record in enumerate(records, start=2):
        for col_idx, column_name in enumerate(REQUESTED_COLUMNS, start=1):
            value = record.get(column_name, "")
            ws.cell(row_idx, col_idx).value = clean_text(value)
        if row_idx % 2 == 0:
            for col_idx in range(1, len(REQUESTED_COLUMNS) + 1):
                ws.cell(row_idx, col_idx).fill = PatternFill("solid", fgColor="F7FBFD")

    max_row = max(ws.max_row, 1)
    max_col = len(REQUESTED_COLUMNS)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"
    ws.row_dimensions[1].height = 28
    ws.sheet_view.showGridLines = True

    for col_idx, column_name in enumerate(REQUESTED_COLUMNS, start=1):
        letter = get_column_letter(col_idx)
        ws.column_dimensions[letter].width = WIDTHS[column_name]
        for cell in ws[letter]:
            cell.alignment = Alignment(vertical="top", wrap_text=column_name in {"abstract", "index_keywords", "task_category", "authors"})
    for row_idx in range(2, max_row + 1):
        ws.row_dimensions[row_idx].height = 42 if row_idx <= 250 else 30

    if records:
        table_ref = f"A1:{get_column_letter(max_col)}{max_row}"
        table_name = re.sub(r"[^A-Za-z0-9_]", "", f"{sheet_name}_Table")[:250]
        table = Table(displayName=table_name, ref=table_ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(table)


def main() -> None:
    records = load_records()
    grouped = {name: [] for name in SHEET_ORDER}
    for record in records:
        grouped.setdefault(bucket_for(record), []).append(record)

    guide_header_cells = []
    if STYLE_GUIDE_XLSX.exists():
        guide_wb = load_workbook(STYLE_GUIDE_XLSX)
        guide_ws = guide_wb["Image"] if "Image" in guide_wb.sheetnames else guide_wb[guide_wb.sheetnames[0]]
        guide_header_cells = [guide_ws.cell(1, col_idx) for col_idx in range(1, guide_ws.max_column + 1)]

    wb = Workbook()
    del wb[wb.sheetnames[0]]
    for sheet_name in SHEET_ORDER:
        add_sheet(wb, sheet_name, grouped.get(sheet_name, []), guide_header_cells)

    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_XLSX)
    print(f"Created {OUTPUT_XLSX}")
    for sheet_name in SHEET_ORDER:
        print(f"{sheet_name}: {len(grouped.get(sheet_name, []))}")


if __name__ == "__main__":
    main()
