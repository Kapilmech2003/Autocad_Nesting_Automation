"""
AutoCAD Nesting Sheet Builder (GUI & Multi-Project CLI)
======================================================
Interactive Dialog & GUI-based application with:
1. Automatic dependency check & self-installation on any computer/laptop.
2. Select any Excel BOM file (*.xlsx).
3. Choose which Worksheet / Vendor Set to convert (or batch-process all sheets).
4. Specify a custom Drawing Sheet Header / Title centered at the top of the CAD file.
5. Select any DXF parts source folder.
6. Select any output save location and filename.
7. Arrange parts in standard AutoCAD grid columns by thickness & material.
8. Stamp metadata (Part ID, Cutlist, Dimensions, Qty, THK, Material).
9. Choose output format: .dxf needed, .dwg needed, or both.

Usage:
  - Interactive UI Window (Default):
      python nesting.py
  - Quick Step-by-Step Dialog Wizard:
      python nesting.py --wizard

  - Command-Line / Headless Execution:
      python nesting.py --cli --excel "path/to/bom.xlsx" --sheet "Set 1" --title "PROJECT ABC - SET 1" --source "path/to/dxfs" --out-dir "path/to/output"
"""
from __future__ import annotations

import argparse
import importlib
import io
import os
import pathlib
import re
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Callable, Any

# ---------------------------------------------------------------------------
# AUTO-DEPENDENCY CHECK & SELF-INSTALLATION
# ---------------------------------------------------------------------------
REQUIRED_PACKAGES = {
    "openpyxl": "openpyxl",
    "ezdxf": "ezdxf",
}


def ensure_dependencies():
    """
    Automatically verify that all required third-party libraries are installed.
    If missing (e.g. running on a new computer or laptop), installs them
    automatically via pip before continuing execution.
    """
    missing = []
    for module_name, pip_name in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(module_name)
        except ImportError:
            missing.append(pip_name)

    if missing:
        print("=" * 60)
        print(" [Auto-Setup] Missing required libraries detected on this system:")
        for pkg in missing:
            print(f"   • {pkg}")
        print(" Automatically installing missing packages... please wait a moment.")
        print("=" * 60)

        for pkg in missing:
            print(f"Installing {pkg}...")
            try:
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", pkg],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.STDOUT
                )
                print(f"✓ Successfully installed {pkg}")
            except Exception:
                # If silent install fails, run with visible output
                print(f"Retrying install of {pkg} with verbose output...")
                subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])
                print(f"✓ Successfully installed {pkg}")

        print("=" * 60)
        print(" [Auto-Setup] All dependencies ready! Launching Nesting Builder...")
        print("=" * 60)


# Run auto-installer
ensure_dependencies()

# Safe imports after ensuring packages exist
import openpyxl
import ezdxf
from ezdxf import bbox as ezbbox

# ---------------------------------------------------------------------------
# DEFAULTS
# ---------------------------------------------------------------------------
DEFAULT_EXCEL = "nesting_list.xlsx"
FALLBACK_EXCEL = "DEFAULT_TEMPLATE.xlsx"
DEFAULT_SOURCE_DIR = "dxf_source"
DEFAULT_OUTPUT_NAME = "combined_nesting"
ALL_SHEETS_OPTION = "(All Sheets - Batch Process)"

# Layout Styling (matches NESTING standard layout)
PADDING_X = 35.0        # mm padding around part inside cell
PADDING_Y = 35.0        # mm padding above/below part inside cell
TEXT_GAP_X = 45.0       # mm gap between part right edge and text block
TEXT_HEIGHT = 14.0      # mm text character height
COLUMN_GAP = 150.0      # mm gap between thickness columns
MIN_TEXT_SPACE = 300.0  # mm reserved width for the MTEXT block

# Title Header Styling at top of CAD drawing
TITLE_TEXT_HEIGHT = 28.0 # mm text height for drawing sheet title
TITLE_GAP_Y = 50.0       # mm vertical gap between top of grid (y=0) and title

# MText height calculation (6 lines with AutoCAD standard line pitch)
NUM_TEXT_LINES = 6
LINE_SPACING_FACTOR = 1.65
TEXT_BLOCK_H = (NUM_TEXT_LINES - 1) * (TEXT_HEIGHT * LINE_SPACING_FACTOR) + TEXT_HEIGHT


# ---------------------------------------------------------------------------
def find_accoreconsole() -> pathlib.Path | None:
    """Locate AutoCAD accoreconsole.exe for DWG conversion."""
    candidates = [
        pathlib.Path(r"C:\Program Files\Autodesk\AutoCAD 2025\accoreconsole.exe"),
        pathlib.Path(r"C:\Program Files\Autodesk\AutoCAD 2024\accoreconsole.exe"),
        pathlib.Path(r"C:\Program Files\Autodesk\AutoCAD 2026\accoreconsole.exe"),
    ]
    for c in candidates:
        if c.exists():
            return c
    autodesk_base = pathlib.Path(r"C:\Program Files\Autodesk")
    if autodesk_base.exists():
        for sub in autodesk_base.glob("AutoCAD */accoreconsole.exe"):
            if sub.exists():
                return sub
    return None


# ---------------------------------------------------------------------------
def get_unique_filepath(target_dir: pathlib.Path, base_name: str, extension: str) -> pathlib.Path:
    """
    Returns an available path. If base_name.ext exists, increments:
    base_name_1.ext, base_name_2.ext, etc.
    """
    ext = extension if extension.startswith(".") else f".{extension}"
    candidate = target_dir / f"{base_name}{ext}"
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = target_dir / f"{base_name}_{counter}{ext}"
        if not candidate.exists():
            return candidate
        counter += 1


# ---------------------------------------------------------------------------
def select_file_dialog(title: str = "Select Excel File", initial_dir: pathlib.Path | None = None) -> pathlib.Path | None:
    """Open a native Windows file dialog to let the user select an Excel file manually."""
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        path = filedialog.askopenfilename(
            title=title,
            initialdir=str(initial_dir) if initial_dir else str(pathlib.Path.cwd()),
            filetypes=[("Excel Files", "*.xlsx;*.xls;*.xlsm"), ("All Files", "*.*")]
        )
        root.destroy()
        if path:
            return pathlib.Path(path)
    except Exception as e:
        print(f"File dialog notification: {e}")

    try:
        user_input = input(f"Please enter path for '{title}' (or press Enter to skip): ").strip().strip('"').strip("'")
        if user_input:
            p = pathlib.Path(user_input)
            if p.exists():
                return p
    except (EOFError, KeyboardInterrupt):
        pass
    return None


# ---------------------------------------------------------------------------
def load_excel_workbook(file_path: pathlib.Path | str) -> openpyxl.Workbook:
    """Safely load an Excel workbook, supporting files currently opened in Microsoft Excel."""
    p = pathlib.Path(file_path)
    try:
        return openpyxl.load_workbook(p, data_only=True)
    except (PermissionError, IOError):
        try:
            import ctypes
            from ctypes import wintypes

            GENERIC_READ = 0x80000000
            FILE_SHARE_READ = 0x00000001
            FILE_SHARE_WRITE = 0x00000002
            FILE_SHARE_DELETE = 0x00000004
            OPEN_EXISTING = 3
            FILE_ATTRIBUTE_NORMAL = 0x80

            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            handle = kernel32.CreateFileW(
                str(p.resolve()),
                GENERIC_READ,
                FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                None,
                OPEN_EXISTING,
                FILE_ATTRIBUTE_NORMAL,
                None
            )
            if handle not in (-1, 0xFFFFFFFF):
                size = kernel32.GetFileSize(handle, None)
                buf = ctypes.create_string_buffer(size)
                bytes_read = wintypes.DWORD()
                kernel32.ReadFile(handle, buf, size, ctypes.byref(bytes_read), None)
                kernel32.CloseHandle(handle)
                data = buf.raw[:bytes_read.value]
                return openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        except Exception:
            pass
        raise


# ---------------------------------------------------------------------------
def get_excel_sheet_names(excel_file: str | pathlib.Path) -> list[str]:
    """Retrieve all sheet names available in the specified Excel workbook."""
    p = pathlib.Path(excel_file)
    if not p.exists() or not p.is_file():
        return []
    try:
        wb = load_excel_workbook(p)
        return list(wb.sheetnames)
    except Exception:
        return []


# ---------------------------------------------------------------------------
def format_dim(val: float) -> str:
    """Format dimension cleanly, e.g. 50.0 -> '50', 50.45 -> '50.5'."""
    r = round(val, 1)
    if r == int(r):
        return str(int(r))
    return f"{r:.1f}"


# ---------------------------------------------------------------------------
def thickness_label(value: Any) -> str:
    """Normalize a thickness number to a clean string (e.g. 1.2, 4, 6)."""
    f = float(value)
    if f == int(f):
        return str(int(f))
    return f"{f:g}"


# ---------------------------------------------------------------------------
def split_text_half(text: str, max_len: int = 18) -> list[str]:
    """
    If text length exceeds max_len, split it into balanced lines ('half half')
    at a natural delimiter (space, underscore, hyphen, dot) near the center.
    If no delimiter exists near the center, splits cleanly at the midpoint.
    """
    text = str(text).strip()
    if len(text) <= max_len:
        return [text]

    mid = len(text) // 2
    best_idx = -1
    best_dist = 9999

    # Look for delimiters in a window around the midpoint (within 35% of string length)
    window = max(2, int(len(text) * 0.35))
    start_pos = max(1, mid - window)
    end_pos = min(len(text) - 1, mid + window + 1)

    for delim in [' ', '_', '-', '.']:
        for i in range(start_pos, end_pos):
            if text[i] == delim:
                dist = abs(i - mid)
                if delim == ' ':
                    dist -= 1  # Give whitespace priority
                if dist < best_dist:
                    best_dist = dist
                    best_idx = i
        if best_idx != -1 and best_dist <= 3:
            break

    if best_idx != -1:
        if text[best_idx] == ' ':
            p1 = text[:best_idx].strip()
            p2 = text[best_idx + 1:].strip()
        elif text[best_idx] in ['_', '-']:
            p1 = text[:best_idx + 1].strip()
            p2 = text[best_idx + 1:].strip()
        else:
            p1 = text[:best_idx].strip()
            p2 = text[best_idx:].strip()

        # If either half is still excessively long (> max_len * 1.7), wrap it recursively
        res = []
        for part in (p1, p2):
            if len(part) > max_len * 1.7:
                res.extend(split_text_half(part, max_len))
            else:
                res.append(part)
        return res
    else:
        return [text[:mid].strip(), text[mid:].strip()]


# ---------------------------------------------------------------------------
def format_part_label(
    part_id: str,
    cutlist: str,
    size_str: str,
    qty_val: Any,
    thk_str: str,
    mat_val: str,
    max_field_len: int = 18
) -> tuple[str, list[str], float, float]:
    """
    Formats the MTEXT label lines for a part cell.
    If PART ID or CUTLIST exceeds max_field_len, it is split into balanced halves
    with aligned indentation.
    Returns: (label_text, lines_list, text_block_h, estimated_text_w)
    """
    lines = []

    # 1. PART ID
    part_id_str = str(part_id).strip() if part_id is not None else ""
    part_splits = split_text_half(part_id_str, max_len=max_field_len)
    if len(part_splits) > 1:
        lines.append(f"PART ID - {part_splits[0]}")
        for continuation in part_splits[1:]:
            lines.append(f"          {continuation}")
    else:
        lines.append(f"PART ID - {part_id_str}")

    # 2. CUTLIST
    cutlist_str = str(cutlist).strip() if cutlist is not None else ""
    if not cutlist_str or cutlist_str.lower() == "none":
        lines.append("CUTLIST - None")
    else:
        cut_splits = split_text_half(cutlist_str, max_len=max_field_len)
        if len(cut_splits) > 1:
            lines.append(f"CUTLIST - {cut_splits[0]}")
            for continuation in cut_splits[1:]:
                lines.append(f"          {continuation}")
        else:
            lines.append(f"CUTLIST - {cutlist_str}")

    # 3. Standard specs
    lines.append(f"SIZE - {size_str}")
    lines.append(f"QTY - {qty_val}")
    lines.append(f"THK - {thk_str}mm")
    lines.append(f"MATERIAL - {mat_val}")

    mtext_str = "\\P".join(lines)
    text_block_h = (len(lines) - 1) * (TEXT_HEIGHT * LINE_SPACING_FACTOR) + TEXT_HEIGHT
    max_chars = max(len(l) for l in lines) if lines else 20
    # Estimated width in mm (average char width ratio ~0.85 * TEXT_HEIGHT in AutoCAD standard font)
    estimated_text_w = max_chars * (TEXT_HEIGHT * 0.85)

    return mtext_str, lines, text_block_h, estimated_text_w


# ---------------------------------------------------------------------------
def read_bom(
    excel_file: str | pathlib.Path,
    sheet_name: str | None = None,
    fallback_excel: str | pathlib.Path | None = None
) -> tuple[list[dict], str]:
    """
    Read the BOM rows from a specific sheet (or active sheet) in the Excel file.
    Returns: (list_of_row_dicts, actual_sheet_name)
    """
    target = pathlib.Path(excel_file)
    fb_path = pathlib.Path(fallback_excel) if fallback_excel else pathlib.Path(FALLBACK_EXCEL)

    if not target.exists():
        if fb_path.exists():
            print(f"'{excel_file}' not found. Using fallback '{fb_path}'...")
            shutil.copy2(fb_path, target)
            target = fb_path
        else:
            print(f"'{excel_file}' and default fallback '{fb_path}' not found.")
            chosen = select_file_dialog("Select Fallback / BOM Excel File")
            if chosen and chosen.exists():
                print(f"Using manually selected fallback: {chosen}")
                shutil.copy2(chosen, target)
                target = chosen
            else:
                raise FileNotFoundError(f"Excel file not found: {excel_file} (fallback: {fb_path})")

    wb = load_excel_workbook(target)

    # Resolve worksheet
    if sheet_name and sheet_name.strip() and sheet_name != ALL_SHEETS_OPTION:
        cleaned_sheet = sheet_name.strip()
        if cleaned_sheet in wb.sheetnames:
            ws = wb[cleaned_sheet]
            actual_sheet_name = cleaned_sheet
        else:
            raise ValueError(
                f"Sheet '{cleaned_sheet}' not found in Excel workbook.\n"
                f"Available sheets: {wb.sheetnames}"
            )
    else:
        ws = wb.active
        actual_sheet_name = ws.title

    header_row = next(ws.iter_rows(min_row=1, max_row=1), None)
    if not header_row:
        raise ValueError(f"Sheet '{actual_sheet_name}' is empty.")

    headers = [str(c.value).strip() if c.value is not None else "" for c in header_row]
    required = {"PartNo", "CutlistName", "DXF_File", "Material", "Thickness_mm", "Qty"}
    missing = required - set(headers)
    if missing:
        raise ValueError(
            f"Sheet '{actual_sheet_name}' is missing required columns: {missing}\n"
            f"Found headers: {headers}"
        )

    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or all(v is None for v in r):
            continue
        row = dict(zip(headers, r))
        if not row.get("DXF_File") and not row.get("PartNo"):
            continue
        rows.append(row)

    return rows, actual_sheet_name


# ---------------------------------------------------------------------------
def _run_accore_script(accore, script_content, work_dir, log_func=print):
    temp_scr = work_dir / f"nesting_autocad_{os.getpid()}_{threading.get_ident()}.scr"
    try:
        temp_scr.write_text(script_content, encoding="utf-8")
        proc = subprocess.run([str(accore), "/s", str(temp_scr)], cwd=str(work_dir), capture_output=True, text=True, timeout=300)
        if proc.stdout.strip(): log_func("[AutoCAD] " + proc.stdout[-4000:])
        if proc.stderr.strip(): log_func("[AutoCAD-ERR] " + proc.stderr[-4000:])
        return proc.returncode, proc.stdout or "", proc.stderr or ""
    except subprocess.TimeoutExpired:
        log_func("[WARNING] AutoCAD Core Console conversion timed out after 300 seconds.")
        return -1, "", "timeout"
    finally:
        try: temp_scr.unlink()
        except Exception: pass


def convert_dxf_with_autocad(dxf_path, cleaned_dxf_path=None, dwg_path=None, log_func=print):
    """DXFIN -> AUDIT -> DXFOUT/SAVEAS using Autodesk AutoCAD Core Console."""
    accore = find_accoreconsole()
    if not accore:
        log_func("[WARNING] AutoCAD accoreconsole.exe not found; AutoCAD round-trip skipped.")
        return False
    src = dxf_path.resolve()
    if not src.exists():
        log_func(f"[WARNING] DXF not found for AutoCAD validation: {src}")
        return False
    if cleaned_dxf_path:
        cleaned_dxf_path = cleaned_dxf_path.resolve(); cleaned_dxf_path.parent.mkdir(parents=True, exist_ok=True)
    if dwg_path:
        dwg_path = dwg_path.resolve(); dwg_path.parent.mkdir(parents=True, exist_ok=True)
    src_str = str(src).replace("\\", "/")
    script = [
    "_.FILEDIA",
    "0",
    "_.CMDECHO",
    "0",
    "_.DXFIN",
    f'"{src_str}"',
    "_.AUDIT",
    "Y",
]

    if dwg_path:
        dwg_str = str(dwg_path).replace("\\", "/")
    script += [
        "_.SAVEAS",
        "2018",
        f'"{dwg_str}"',
    ]

    script += [
    "_.QUIT",
    "Y",
    "",
]

    rc, stdout, stderr = _run_accore_script(
    accore,
    "\n".join(script),
    src.parent,
    log_func,
)

    if rc != 0:
        log_func(f"[WARNING] AutoCAD Core Console returned exit code {rc}.")

    if cleaned_dxf_path and cleaned_dxf_path.exists():
        log_func(f"AutoCAD-validated DXF: {cleaned_dxf_path}")

    if dwg_path and dwg_path.exists():
        log_func(f"AutoCAD-generated DWG: {dwg_path}")

    return bool(
        (cleaned_dxf_path and cleaned_dxf_path.exists())
        or (dwg_path and dwg_path.exists())
)


def convert_to_dwg(dxf_path, dwg_path, log_func=print):
    return convert_dxf_with_autocad(dxf_path, dwg_path=dwg_path, log_func=log_func)


# ---------------------------------------------------------------------------
def sanitize_filename(name: str) -> str:
    """Sanitize string for safe filenames."""
    return re.sub(r'[\\/*?:"<>| ]', '_', name)


def _clean_entity_for_autocad(entity):
    """Remove source-specific metadata before moving geometry into a clean DXF."""
    try: entity.discard_xdata()
    except Exception: pass
    try: entity.discard_app_data()
    except Exception: pass
    try: entity.discard_extension_dict()
    except Exception: pass
    try: entity.set_reactors([])
    except Exception: pass
    try:
        if entity.dxf.hasattr("linetype"): entity.dxf.linetype = "BYLAYER"
    except Exception: pass
    try: entity.dxf.layer = "0"
    except Exception: pass
    try: entity.dxf.color = 7
    except Exception: pass
    return entity


def _copy_geometry_autocad_safe(source_entity, target_msp, dx=0.0, dy=0.0):
    """Copy manufacturing geometry without carrying source BLOCK/table/proxy dependencies."""
    etype = source_entity.dxftype()
    if etype in {"INSERT", "POLYLINE"}:
        try:
            return sum(_copy_geometry_autocad_safe(v, target_msp, dx, dy) for v in source_entity.virtual_entities())
        except Exception:
            return 0
    supported = {"LINE", "CIRCLE", "ARC", "LWPOLYLINE", "ELLIPSE", "SPLINE"}
    if etype not in supported:
        return 0
    try:
        new_entity = source_entity.copy()
        _clean_entity_for_autocad(new_entity)
        new_entity.translate(dx, dy, 0.0)
        target_msp.add_entity(new_entity)
        return 1
    except Exception:
        return 0


def _audit_ezdxf_document(doc, log_func=print, label="drawing"):
    try:
        auditor = doc.audit()
        issues = auditor.run()
        if auditor.has_issues:
            log_func(f"[WARNING] ezdxf audit found {len(issues)} issue(s) in {label}.")
        if auditor.has_fixes:
            log_func(f"[INFO] ezdxf audit repaired {len(auditor.fixes)} issue(s) in {label}.")
        return auditor
    except Exception as exc:
        log_func(f"[WARNING] Could not complete ezdxf audit for {label}: {exc}")
        return None


# ---------------------------------------------------------------------------
def build_combined_nesting(
    rows: list[dict],
    source_dir: pathlib.Path,
    output_dxf: pathlib.Path | None = None,
    output_dwg: pathlib.Path | None = None,
    need_dxf: bool = True,
    need_dwg: bool = False,
    drawing_title: str | None = None,
    log_func: Callable[[str], None] = print
) -> tuple[int, int]:
    """
    Builds the combined DXF with centered drawing header title and converts it to DWG.
    Returns: (placed_count, failed_count)
    """
    doc = ezdxf.new("R2018")

    # Setup layers matching AutoCAD reference format
    if "FRAME" not in doc.layers:
        doc.layers.new("FRAME", dxfattribs={"color": 3})  # Green border/grid lines
    if "PART_TEXT" not in doc.layers:
        doc.layers.new("PART_TEXT", dxfattribs={"color": 7})  # White text annotations
    if "TITLE" not in doc.layers:
        doc.layers.new("TITLE", dxfattribs={"color": 4})  # Cyan header title layer
    msp = doc.modelspace()

    # Group parts by thickness
    columns_data: dict[str, list[dict]] = {}
    for row in rows:
        thk_key = thickness_label(row["Thickness_mm"])
        columns_data.setdefault(thk_key, []).append(row)

    # Sort columns left-to-right by numeric thickness
    sorted_thicknesses = sorted(columns_data.keys(), key=lambda k: float(k))

    placed_count = 0
    failed_count = 0
    x_cursor = 0.0

    for thk_key in sorted_thicknesses:
        parts_list = columns_data[thk_key]

        measured_parts = []
        max_part_width = 0.0

        for row in parts_list:
            dxf_filename = str(row.get("DXF_File", "")).strip()
            dxf_path = source_dir / dxf_filename
            if not dxf_path.exists():
                log_func(f"[WARNING] Missing DXF file: '{dxf_filename}' in '{source_dir}'")
                failed_count += 1
                continue
            try:
                part_doc = ezdxf.readfile(dxf_path)
            except Exception as exc:
                log_func(f"[WARNING] Error reading '{dxf_filename}': {exc}")
                failed_count += 1
                continue

            _audit_ezdxf_document(part_doc, log_func, label=dxf_filename)
            part_msp = part_doc.modelspace()
            cache_box = ezbbox.extents(part_msp)
            if cache_box is None or not cache_box.has_data:
                log_func(f"[WARNING] Empty geometry in '{dxf_filename}' - skipped")
                failed_count += 1
                continue

            p_w = cache_box.extmax.x - cache_box.extmin.x
            p_h = cache_box.extmax.y - cache_box.extmin.y
            max_part_width = max(max_part_width, p_w)

            qty_val = int(row["Qty"]) if float(row["Qty"]).is_integer() else row["Qty"]
            mat_val = str(row["Material"]).strip().upper()
            part_id = str(row.get("PartNo", "")).strip()
            cutlist = str(row.get("CutlistName", "")).strip()
            size_str = f"{format_dim(p_w)} x {format_dim(p_h)} mm"

            label_text, label_lines, part_text_h, part_text_w = format_part_label(
                part_id=part_id,
                cutlist=cutlist,
                size_str=size_str,
                qty_val=qty_val,
                thk_str=thk_key,
                mat_val=mat_val
            )

            # Cell height fits the maximum of part height or full text block height + margins
            cell_h = max(p_h, part_text_h) + PADDING_Y * 2

            measured_parts.append({
                "row": row,
                "part_doc": part_doc,
                "part_msp": part_msp,
                "box": cache_box,
                "width": p_w,
                "height": p_h,
                "cell_height": cell_h,
                "label_text": label_text,
                "text_block_h": part_text_h,
                "text_width": part_text_w
            })

        if not measured_parts:
            continue

        # Column width & total height (guaranteeing green boundary line NEVER overlays text)
        max_text_width = max(p["text_width"] for p in measured_parts)
        col_text_space = max(MIN_TEXT_SPACE, max_text_width + 40.0)
        col_width = PADDING_X + max_part_width + TEXT_GAP_X + col_text_space + PADDING_X
        total_col_height = sum(p["cell_height"] for p in measured_parts)

        col_x0 = x_cursor
        col_x1 = x_cursor + col_width
        col_y0 = 0.0
        col_y1 = -total_col_height

        # Outer bounding rectangle frame (Green)
        rect_pts = [(col_x0, col_y0), (col_x1, col_y0), (col_x1, col_y1), (col_x0, col_y1)]
        msp.add_lwpolyline(rect_pts, close=True, dxfattribs={"layer": "FRAME", "color": 3})

        # Place cells and separator lines
        current_y_top = col_y0

        for idx, item in enumerate(measured_parts):
            row = item["row"]
            cell_h = item["cell_height"]
            cell_y_bot = current_y_top - cell_h

            # Horizontal divider line (Green)
            if idx < len(measured_parts) - 1:
                msp.add_line(
                    (col_x0, cell_y_bot),
                    (col_x1, cell_y_bot),
                    dxfattribs={"layer": "FRAME", "color": 3}
                )

            # Place Part geometry inside cell (left aligned + centered vertically)
            p_box = item["box"]
            p_w = item["width"]
            p_h = item["height"]

            part_target_x = col_x0 + PADDING_X
            part_target_y = cell_y_bot + (cell_h - p_h) / 2.0

            dx = part_target_x - p_box.extmin.x
            dy = part_target_y - p_box.extmin.y

            copied_geometry = 0
            skipped_geometry = 0
            for e in item["part_msp"]:
                copied = _copy_geometry_autocad_safe(e, msp, dx=dx, dy=dy)
                if copied:
                    copied_geometry += copied
                else:
                    skipped_geometry += 1
            if skipped_geometry:
                log_func(f"[INFO] '{row.get('DXF_File', '')}': {copied_geometry} geometry entities copied; {skipped_geometry} non-cut/unsupported entities skipped.")

            # Place MTEXT annotation to the right of the part
            text_x = col_x0 + PADDING_X + max_part_width + TEXT_GAP_X
            text_top_y = cell_y_bot + (cell_h + item["text_block_h"]) / 2.0

            msp.add_mtext(
                item["label_text"],
                dxfattribs={
                    "layer": "PART_TEXT",
                    "color": 7,
                    "char_height": TEXT_HEIGHT,
                    "insert": (text_x, text_top_y),
                    "attachment_point": 1
                }
            )

            placed_count += 1
            current_y_top = cell_y_bot

        x_cursor += col_width + COLUMN_GAP

    # Stamp Drawing Header Title centered across all columns at the top of the CAD sheet (y > 0)
    if drawing_title and drawing_title.strip():
        title_clean = drawing_title.strip().upper()
        total_layout_width = max(x_cursor - COLUMN_GAP, 100.0) if x_cursor > 0 else 500.0
        center_x = total_layout_width / 2.0
        title_insert_y = TITLE_GAP_Y + TITLE_TEXT_HEIGHT

        # Clean top-center MTEXT without escaping bugs
        msp.add_mtext(
            title_clean,
            dxfattribs={
                "layer": "TITLE",
                "color": 4,  # Cyan
                "char_height": TITLE_TEXT_HEIGHT,
                "insert": (center_x, title_insert_y),
                "attachment_point": 2  # Top-Center (horizontally centered across all columns)
            }
        )
        # Decorative underline spanning the entire width of all columns
        underline_y = TITLE_GAP_Y - 10.0
        msp.add_line((0.0, underline_y), (total_layout_width, underline_y), dxfattribs={"layer": "TITLE", "color": 4})
        log_func(f"Stamped Centered Drawing Header Title: '{title_clean}' (width: {total_layout_width:.1f}mm).")

    # Ensure at least one format is requested
    if not need_dxf and not need_dwg:
        need_dxf = True

    _audit_ezdxf_document(doc, log_func, label="combined nesting drawing")

    # Save output DXF if requested
    if need_dxf and output_dxf:
        output_dxf.parent.mkdir(parents=True, exist_ok=True)
        try:
            doc.saveas(str(output_dxf))
            log_func(f"Generated DXF: {output_dxf}")
        except PermissionError:
            alt_dxf = get_unique_filepath(output_dxf.parent, output_dxf.stem, ".dxf")
            log_func(f"Warning: '{output_dxf.name}' is currently locked. Saved as '{alt_dxf.name}'.")
            doc.saveas(str(alt_dxf))
            output_dxf = alt_dxf

    # AutoCAD round-trip makes the final DXF/DWG Autodesk-authored and audited.
    if (need_dxf and output_dxf and output_dxf.exists()) or (need_dwg and output_dwg):
        source_for_autocad = output_dxf if (need_dxf and output_dxf and output_dxf.exists()) else None
        temp_dxf = None
        if source_for_autocad is None:
            temp_dxf = get_unique_filepath(output_dwg.parent, f"__temp_{output_dwg.stem}", ".dxf")
            doc.saveas(str(temp_dxf)); source_for_autocad = temp_dxf
        roundtrip_dxf = None
        if need_dxf and output_dxf:
            roundtrip_dxf = get_unique_filepath(output_dxf.parent, f"__autocad_clean_{output_dxf.stem}", ".dxf")
        autocad_ok = convert_dxf_with_autocad(source_for_autocad, cleaned_dxf_path=roundtrip_dxf, dwg_path=output_dwg if need_dwg else None, log_func=log_func)
        if roundtrip_dxf and roundtrip_dxf.exists():
            try:
                if output_dxf.exists(): output_dxf.unlink()
                roundtrip_dxf.replace(output_dxf)
                log_func(f"✓ Final DXF replaced with AutoCAD-audited DXF: {output_dxf}")
            except Exception as exc:
                log_func(f"[WARNING] Could not replace DXF with audited copy: {exc}")
        elif need_dxf and not autocad_ok:
            log_func("[WARNING] AutoCAD validation unavailable/failed; keeping the ezdxf-generated DXF.")
        if temp_dxf and temp_dxf.exists():
            try: temp_dxf.unlink()
            except Exception: pass

    return placed_count, failed_count


# ---------------------------------------------------------------------------
def run_single_sheet_nesting(
    excel_path: pathlib.Path,
    sheet_name: str | None,
    source_dir: pathlib.Path,
    output_dir: pathlib.Path,
    output_base_name: str,
    drawing_title: str | None = None,
    need_dxf: bool = True,
    need_dwg: bool = False,
    fallback_excel: pathlib.Path | None = None,
    log_func: Callable[[str], None] = print
) -> dict:
    """Execute nesting generation for a single Excel worksheet."""
    rows, actual_sheet = read_bom(excel_path, sheet_name=sheet_name, fallback_excel=fallback_excel)
    if not rows:
        raise ValueError(f"No valid part rows found in sheet '{actual_sheet}'.")

    log_func(f"Loaded {len(rows)} part entries from Sheet: '{actual_sheet}'.")

    # Output file paths
    out_dxf = None
    out_dwg = None

    if need_dxf:
        out_dxf = get_unique_filepath(output_dir, output_base_name, ".dxf")
        log_func(f"Target Output DXF: {out_dxf.name}")
    else:
        out_dxf = output_dir / f"{output_base_name}.dxf"

    if need_dwg:
        out_dwg = get_unique_filepath(output_dir, output_base_name, ".dwg")
        log_func(f"Target Output DWG: {out_dwg.name}")

    # Build nesting
    placed_count, failed_count = build_combined_nesting(
        rows,
        source_dir,
        output_dxf=out_dxf,
        output_dwg=out_dwg,
        need_dxf=need_dxf,
        need_dwg=need_dwg,
        drawing_title=drawing_title,
        log_func=log_func
    )

    return {
        "sheet_name": actual_sheet,
        "placed": placed_count,
        "failed": failed_count,
        "out_dxf": out_dxf if (need_dxf and out_dxf and out_dxf.exists()) else None,
        "out_dwg": out_dwg if (need_dwg and out_dwg and out_dwg.exists()) else None,
        "need_dxf": need_dxf,
        "need_dwg": need_dwg,
    }


# ---------------------------------------------------------------------------
def run_nesting_process(
    excel_path: pathlib.Path | str,
    source_dir: pathlib.Path | str,
    output_dir: pathlib.Path | str,
    sheet_name: str | None = None,
    drawing_title: str | None = None,
    output_name: str = DEFAULT_OUTPUT_NAME,
    need_dxf: bool = True,
    need_dwg: bool = False,
    fallback_excel: pathlib.Path | str | None = None,
    log_func: Callable[[str], None] = print
) -> dict:
    """
    Main pipeline function with multi-sheet and centered in-drawing title heading support.
    """
    excel_p = pathlib.Path(excel_path).resolve()
    source_p = pathlib.Path(source_dir).resolve()
    out_dir_p = pathlib.Path(output_dir).resolve()
    out_dir_p.mkdir(parents=True, exist_ok=True)

    fb_p = pathlib.Path(fallback_excel).resolve() if fallback_excel else None

    # Resolve formats
    if not need_dxf and not need_dwg:
        need_dxf = True

    fmt_labels = []
    if need_dxf:
        fmt_labels.append(".DXF")
    if need_dwg:
        fmt_labels.append(".DWG")

    log_func("=" * 60)
    log_func("Nesting Build Process Started")
    log_func("=" * 60)
    log_func(f"BOM Excel File:    {excel_p}")
    if fb_p:
        log_func(f"Fallback Template: {fb_p}")
    log_func(f"DXF Source Folder: {source_p}")
    log_func(f"Output Save Folder:{out_dir_p}")
    log_func(f"Output Formats:    {', '.join(fmt_labels)}")
    if drawing_title:
        log_func(f"Drawing Title:     {drawing_title}")

    # Validate inputs
    if not excel_p.exists() and (not fb_p or not fb_p.exists()):
        raise FileNotFoundError(f"Excel BOM file does not exist: {excel_p}")

    if not source_p.exists() or not source_p.is_dir():
        raise FileNotFoundError(f"DXF source folder does not exist: {source_p}")

    base_name = output_name.strip() if output_name else DEFAULT_OUTPUT_NAME

    # Check if user requested batch processing of all sheets
    if sheet_name == ALL_SHEETS_OPTION or sheet_name == "__ALL__":
        all_sheets = get_excel_sheet_names(excel_p)
        if not all_sheets:
            raise ValueError(f"No worksheets found in {excel_p.name}")

        log_func(f"Batch processing {len(all_sheets)} sheets: {', '.join(all_sheets)}")
        results_list = []
        total_placed = 0
        total_failed = 0

        for s_name in all_sheets:
            log_func("-" * 50)
            log_func(f"Processing Sheet / Set: '{s_name}'...")
            sheet_suffix = sanitize_filename(s_name)
            s_base_name = f"{base_name}_{sheet_suffix}"

            # Customize title for each sheet in batch
            if drawing_title and drawing_title.strip():
                sheet_title = f"{drawing_title.strip()} - {s_name.upper()}"
            else:
                sheet_title = s_name.upper()

            try:
                res = run_single_sheet_nesting(
                    excel_path=excel_p,
                    sheet_name=s_name,
                    source_dir=source_p,
                    output_dir=out_dir_p,
                    output_base_name=s_base_name,
                    drawing_title=sheet_title,
                    need_dxf=need_dxf,
                    need_dwg=need_dwg,
                    fallback_excel=fb_p,
                    log_func=log_func
                )
                results_list.append(res)
                total_placed += res["placed"]
                total_failed += res["failed"]
                log_func(f"Sheet '{s_name}': {res['placed']} parts placed, {res['failed']} issues.")
            except Exception as e:
                log_func(f"[WARNING] Skipping sheet '{s_name}': {e}")

        if not results_list:
            raise ValueError("None of the sheets could be processed. Please check table headers.")

        log_func("=" * 60)
        log_func(f"BATCH FINISHED: {len(results_list)} sheets processed.")
        log_func(f"Total Placed: {total_placed} parts, Total Issues: {total_failed}.")
        log_func(f"Output Directory: {out_dir_p}")
        log_func("=" * 60)

        first_res = results_list[0]
        return {
            "success": True,
            "is_batch": True,
            "placed": total_placed,
            "failed": total_failed,
            "sheets_processed": len(results_list),
            "out_dxf": first_res.get("out_dxf"),
            "out_dwg": first_res.get("out_dwg"),
            "need_dxf": need_dxf,
            "need_dwg": need_dwg,
            "results": results_list
        }
    else:
        # Single sheet processing
        final_title = drawing_title.strip() if drawing_title and drawing_title.strip() else None
        res = run_single_sheet_nesting(
            excel_path=excel_p,
            sheet_name=sheet_name,
            source_dir=source_p,
            output_dir=out_dir_p,
            output_base_name=base_name,
            drawing_title=final_title,
            need_dxf=need_dxf,
            need_dwg=need_dwg,
            fallback_excel=fb_p,
            log_func=log_func
        )

        log_func("=" * 60)
        log_func(f"Summary for Sheet '{res['sheet_name']}': {res['placed']} parts placed, {res['failed']} issues.")
        if res.get("out_dxf") and res["out_dxf"].exists():
            log_func(f"Generated DXF: {res['out_dxf']}")
        if res.get("out_dwg") and res["out_dwg"].exists():
            log_func(f"Generated DWG: {res['out_dwg']}")
        log_func("=" * 60)

        return {
            "success": True,
            "is_batch": False,
            **res
        }


# ---------------------------------------------------------------------------
def process_project(
    project_dir: pathlib.Path,
    excel_file: str | None = None,
    sheet_name: str | None = None,
    drawing_title: str | None = None,
    fallback_excel: str | None = None,
    source_dir_name: str | None = None,
    output_name: str | None = None,
    need_dxf: bool = True,
    need_dwg: bool = False
):
    """Process a project directory (CLI / Batch support)."""
    project_dir = project_dir.resolve()

    # 1. Resolve Fallback Excel
    fallback_path = (project_dir / fallback_excel) if fallback_excel else None

    # 2. Resolve Primary Excel BOM path
    if excel_file:
        excel_path = (project_dir / excel_file) if not pathlib.Path(excel_file).is_absolute() else pathlib.Path(excel_file)
    else:
        if (project_dir / DEFAULT_EXCEL).exists():
            excel_path = project_dir / DEFAULT_EXCEL
        elif fallback_path and fallback_path.exists():
            excel_path = fallback_path
        elif (project_dir / FALLBACK_EXCEL).exists():
            excel_path = project_dir / FALLBACK_EXCEL
        else:
            candidates = [f for f in project_dir.glob("*.xlsx") if not f.name.startswith("~$")]
            excel_path = candidates[0] if candidates else project_dir / DEFAULT_EXCEL

    # 3. Resolve DXF source folder
    if source_dir_name:
        source_path = (project_dir / source_dir_name) if not pathlib.Path(source_dir_name).is_absolute() else pathlib.Path(source_dir_name)
    else:
        if (project_dir / DEFAULT_SOURCE_DIR).exists():
            source_path = project_dir / DEFAULT_SOURCE_DIR
        else:
            source_path = project_dir

    try:
        run_nesting_process(
            excel_path=excel_path,
            sheet_name=sheet_name,
            drawing_title=drawing_title,
            source_dir=source_path,
            output_dir=project_dir,
            output_name=output_name or DEFAULT_OUTPUT_NAME,
            need_dxf=need_dxf,
            need_dwg=need_dwg,
            fallback_excel=fallback_path,
            log_func=print
        )
    except Exception as exc:
        print(f"Error processing project {project_dir.name}: {exc}")


# ---------------------------------------------------------------------------
# TKINTER UI APPLICATION
# ---------------------------------------------------------------------------
class NestingAppGUI:
    """Modern Tkinter UI for Interactive Dialog-based Nesting with Drawing Title support."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("AutoCAD Nesting Sheet Builder")
        self.root.geometry("880x800")
        self.root.minsize(740, 620)

        # Style configuration
        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.style.configure("TFrame", background="#F5F6F8")
        self.style.configure("TLabelframe", background="#F5F6F8", font=("Segoe UI", 10, "bold"))
        self.style.configure("TLabelframe.Label", background="#F5F6F8", foreground="#2C3E50", font=("Segoe UI", 10, "bold"))
        self.style.configure("TLabel", background="#F5F6F8", font=("Segoe UI", 9))
        self.style.configure("Header.TLabel", background="#2C3E50", foreground="#FFFFFF", font=("Segoe UI", 14, "bold"))
        self.style.configure("SubHeader.TLabel", background="#2C3E50", foreground="#BDC3C7", font=("Segoe UI", 9))
        self.style.configure("TButton", font=("Segoe UI", 9))
        self.style.configure("TCombobox", font=("Segoe UI", 9))

        # StringVars
        self.var_excel = tk.StringVar()
        self.var_sheet = tk.StringVar()
        self.var_title = tk.StringVar(value="")
        self.var_source = tk.StringVar()
        self.var_fallback = tk.StringVar()
        self.var_out_dir = tk.StringVar()
        self.var_out_name = tk.StringVar(value=DEFAULT_OUTPUT_NAME)
        self.var_need_dxf = tk.BooleanVar(value=True)
        self.var_need_dwg = tk.BooleanVar(value=False)
        self.var_status = tk.StringVar(value="Ready. Select Excel BOM, choose worksheet/set, and enter drawing title.")

        self.available_sheets: list[str] = []
        self.last_out_dir: pathlib.Path | None = None
        self.is_running = False

        self._create_widgets()
        self._auto_detect_defaults()

    def _create_widgets(self):
        # 1. Header Banner
        header_frame = tk.Frame(self.root, bg="#2C3E50", padx=16, pady=12)
        header_frame.pack(fill=tk.X)

        title_lbl = ttk.Label(header_frame, text="📐 AutoCAD Nesting Sheet Builder", style="Header.TLabel")
        title_lbl.pack(anchor=tk.W)

        sub_lbl = ttk.Label(
            header_frame,
            text="Interactive Multi-Sheet Nesting: Select Excel BOM, pick Vendor Set, set in-drawing Title, and save.",
            style="SubHeader.TLabel"
        )
        sub_lbl.pack(anchor=tk.W, pady=(2, 0))

        # Main Container
        main_frame = ttk.Frame(self.root, padding=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 2. Input Selection Group
        in_group = ttk.LabelFrame(main_frame, text=" 1. Input Files & Project Details ", padding=10)
        in_group.pack(fill=tk.X, pady=(0, 8))

        # BOM Excel File
        r = 0
        ttk.Label(in_group, text="BOM Excel File (*.xlsx):").grid(row=r, column=0, sticky=tk.W, pady=4)
        ent_excel = ttk.Entry(in_group, textvariable=self.var_excel, width=52)
        ent_excel.grid(row=r, column=1, sticky=tk.EW, padx=6, pady=4)
        btn_excel = ttk.Button(in_group, text="Browse File...", command=self.browse_excel)
        btn_excel.grid(row=r, column=2, padx=2, pady=4)

        # Worksheet / Vendor Set Dropdown
        r += 1
        ttk.Label(in_group, text="Select Sheet / Vendor Set:").grid(row=r, column=0, sticky=tk.W, pady=4)

        sheet_subframe = ttk.Frame(in_group)
        sheet_subframe.grid(row=r, column=1, sticky=tk.EW, padx=6, pady=4)

        self.cbo_sheet = ttk.Combobox(
            sheet_subframe,
            textvariable=self.var_sheet,
            state="readonly",
            font=("Segoe UI", 9)
        )
        self.cbo_sheet.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cbo_sheet.bind("<<ComboboxSelected>>", self._on_sheet_selected)

        btn_refresh = ttk.Button(sheet_subframe, text="🔄 Reload Sheets", command=self.reload_sheets, width=15)
        btn_refresh.pack(side=tk.RIGHT, padx=(6, 0))

        # Drawing Sheet Title / Heading (Inside Drawing)
        r += 1
        ttk.Label(in_group, text="Drawing Title / Heading:").grid(row=r, column=0, sticky=tk.W, pady=4)
        title_subframe = ttk.Frame(in_group)
        title_subframe.grid(row=r, column=1, sticky=tk.EW, padx=6, pady=4)

        ent_title = ttk.Entry(title_subframe, textvariable=self.var_title, font=("Segoe UI", 9))
        ent_title.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Label(
            in_group,
            text="(Centered at top of drawing)",
            foreground="#7F8C8D"
        ).grid(row=r, column=2, sticky=tk.W, padx=2, pady=4)

        # DXF Source Folder
        r += 1
        ttk.Label(in_group, text="DXF Parts Folder:").grid(row=r, column=0, sticky=tk.W, pady=4)
        ent_source = ttk.Entry(in_group, textvariable=self.var_source, width=52)
        ent_source.grid(row=r, column=1, sticky=tk.EW, padx=6, pady=4)
        btn_source = ttk.Button(in_group, text="Browse Folder...", command=self.browse_source)
        btn_source.grid(row=r, column=2, padx=2, pady=4)

        # Fallback Excel Template (Optional)
        r += 1
        ttk.Label(in_group, text="Template / Fallback Excel:").grid(row=r, column=0, sticky=tk.W, pady=4)
        ent_fallback = ttk.Entry(in_group, textvariable=self.var_fallback, width=52)
        ent_fallback.grid(row=r, column=1, sticky=tk.EW, padx=6, pady=4)
        btn_fallback = ttk.Button(in_group, text="Browse File...", command=self.browse_fallback)
        btn_fallback.grid(row=r, column=2, padx=2, pady=4)

        in_group.columnconfigure(1, weight=1)

        # 3. Output Configuration Group
        out_group = ttk.LabelFrame(main_frame, text=" 2. Output & Save Location ", padding=10)
        out_group.pack(fill=tk.X, pady=(0, 8))

        # Save Directory
        r = 0
        ttk.Label(out_group, text="Save Output Folder:").grid(row=r, column=0, sticky=tk.W, pady=4)
        ent_out_dir = ttk.Entry(out_group, textvariable=self.var_out_dir, width=52)
        ent_out_dir.grid(row=r, column=1, sticky=tk.EW, padx=6, pady=4)
        btn_out_dir = ttk.Button(out_group, text="Browse Folder...", command=self.browse_out_dir)
        btn_out_dir.grid(row=r, column=2, padx=2, pady=4)

        # Output Base Name
        r += 1
        name_subframe = ttk.Frame(out_group)
        name_subframe.grid(row=r, column=0, columnspan=3, sticky=tk.EW, pady=4)

        ttk.Label(name_subframe, text="Output Base Name:").pack(side=tk.LEFT)
        ent_name = ttk.Entry(name_subframe, textvariable=self.var_out_name, width=26)
        ent_name.pack(side=tk.LEFT, padx=6)
        ttk.Label(name_subframe, text="(Auto-increments output filename if exists)", foreground="#7F8C8D").pack(side=tk.LEFT)

        btn_save_as = ttk.Button(name_subframe, text="Save As Dialog...", command=self.browse_save_as)
        btn_save_as.pack(side=tk.RIGHT, padx=2)

        # Output Format Options (2 Options: .dxf needed, .dwg needed)
        r += 1
        fmt_subframe = ttk.Frame(out_group)
        fmt_subframe.grid(row=r, column=0, columnspan=3, sticky=tk.W, pady=(6, 2))

        ttk.Label(fmt_subframe, text="Select Output Format:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 10))

        chk_dxf = ttk.Checkbutton(
            fmt_subframe,
            text=".dxf needed",
            variable=self.var_need_dxf,
            command=self._on_format_changed
        )
        chk_dxf.pack(side=tk.LEFT, padx=(0, 16))

        chk_dwg = ttk.Checkbutton(
            fmt_subframe,
            text=".dwg needed",
            variable=self.var_need_dwg,
            command=self._on_format_changed
        )
        chk_dwg.pack(side=tk.LEFT, padx=(0, 16))

        self.lbl_format_hint = ttk.Label(fmt_subframe, text="(Only .DXF will be saved)", foreground="#2980B9")
        self.lbl_format_hint.pack(side=tk.LEFT)

        out_group.columnconfigure(1, weight=1)

        # 4. Action Buttons Bar
        act_frame = ttk.Frame(main_frame, padding=(0, 4))
        act_frame.pack(fill=tk.X, pady=(0, 6))

        self.btn_run = tk.Button(
            act_frame,
            text="▶ Generate Nesting (.DXF)",
            command=self.start_build_thread,
            bg="#27AE60",
            fg="#FFFFFF",
            activebackground="#2ECC71",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 10, "bold"),
            padx=16,
            pady=6,
            relief=tk.RAISED,
            cursor="hand2"
        )
        self.btn_run.pack(side=tk.LEFT, padx=(0, 6))


        btn_wizard = tk.Button(
            act_frame,
            text="🧙 Step-by-Step Wizard",
            command=self.run_wizard_dialogs,
            bg="#2980B9",
            fg="#FFFFFF",
            activebackground="#3498DB",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=6,
            relief=tk.RAISED,
            cursor="hand2"
        )
        btn_wizard.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_open_folder = ttk.Button(act_frame, text="📂 Open Output Folder", command=self.open_output_folder)
        self.btn_open_folder.pack(side=tk.LEFT, padx=(0, 6))

        btn_clear = ttk.Button(act_frame, text="🧹 Reset Fields", command=self.reset_fields)
        btn_clear.pack(side=tk.RIGHT)

        # 5. Live Activity Log Area
        log_group = ttk.LabelFrame(main_frame, text=" 3. Live Execution Activity ", padding=6)
        log_group.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        self.txt_log = tk.Text(
            log_group,
            bg="#1E1E1E",
            fg="#E0E0E0",
            insertbackground="#FFFFFF",
            font=("Consolas", 9),
            wrap=tk.WORD,
            height=10
        )
        log_scroll = ttk.Scrollbar(log_group, orient=tk.VERTICAL, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=log_scroll.set)

        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Tag colors for logs
        self.txt_log.tag_config("info", foreground="#61AFEF")
        self.txt_log.tag_config("success", foreground="#98C379")
        self.txt_log.tag_config("warning", foreground="#E5C07B")
        self.txt_log.tag_config("error", foreground="#E06C75")
        self.txt_log.tag_config("header", foreground="#C678DD", font=("Consolas", 9, "bold"))

        # 6. Status Bar
        status_bar = tk.Frame(self.root, bg="#ECEFF1", padx=8, pady=4)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        lbl_status = ttk.Label(status_bar, textvariable=self.var_status, font=("Segoe UI", 8), background="#ECEFF1")
        lbl_status.pack(side=tk.LEFT)

    def log(self, text: str, tag: str = "info"):
        """Thread-safe log printer into the Text widget."""
        def _append():
            self.txt_log.insert(tk.END, text + "\n", tag)
            self.txt_log.see(tk.END)
        self.root.after(0, _append)

    def _on_format_changed(self):
        """Update button label and format hint when format checkboxes change."""
        need_dxf = self.var_need_dxf.get()
        need_dwg = self.var_need_dwg.get()

        if need_dxf and need_dwg:
            label = "▶ Generate Nesting (.DXF & .DWG)"
            hint = "(Both .DXF and .DWG will be saved)"
            hint_color = "#27AE60"
        elif need_dwg:
            label = "▶ Generate Nesting (.DWG)"
            hint = "(Only .DWG will be saved)"
            hint_color = "#E67E22"
        elif need_dxf:
            label = "▶ Generate Nesting (.DXF)"
            hint = "(Only .DXF will be saved)"
            hint_color = "#2980B9"
        else:
            label = "▶ Generate Nesting"
            hint = "(⚠ Select at least one format)"
            hint_color = "#E06C75"

        if not self.is_running:
            self.btn_run.config(text=label)
        self.lbl_format_hint.config(text=hint, foreground=hint_color)

    def reload_sheets(self, auto_select_first: bool = False):
        """Inspect the current Excel file and populate the sheet dropdown."""
        excel_path = self.var_excel.get().strip()
        if not excel_path or not pathlib.Path(excel_path).exists():
            self.available_sheets = []
            self.cbo_sheet["values"] = []
            self.var_sheet.set("")
            return

        sheets = get_excel_sheet_names(excel_path)
        self.available_sheets = sheets

        if sheets:
            combo_options = sheets + [ALL_SHEETS_OPTION]
            self.cbo_sheet["values"] = combo_options

            current_val = self.var_sheet.get()
            if current_val in combo_options:
                pass
            elif auto_select_first or not current_val:
                self.var_sheet.set(sheets[0])
                self._update_output_name_and_title(sheets[0])

            self.log(f"Detected {len(sheets)} sheet(s) in '{pathlib.Path(excel_path).name}': {', '.join(sheets)}", "info")
        else:
            self.cbo_sheet["values"] = []
            self.var_sheet.set("")
            self.log(f"Warning: Could not read sheet names from '{excel_path}'", "warning")

    def _on_sheet_selected(self, event=None):
        """Handle user changing sheet in dropdown."""
        chosen = self.var_sheet.get()
        if chosen:
            self._update_output_name_and_title(chosen)
            self.log(f"Selected Worksheet: '{chosen}'", "info")

    def _update_output_name_and_title(self, sheet_name: str):
        """Adapt output name and suggest in-drawing title when switching sheets."""
        excel_path = self.var_excel.get().strip()
        proj_name = pathlib.Path(excel_path).parent.name if excel_path else "PROJECT"

        if sheet_name == ALL_SHEETS_OPTION:
            self.var_out_name.set(DEFAULT_OUTPUT_NAME)
            if not self.var_title.get() or self.var_title.get().startswith("PROJECT:"):
                self.var_title.set(f"PROJECT: {proj_name.upper()}")
        elif sheet_name:
            sanitized = sanitize_filename(sheet_name)
            current = self.var_out_name.get()
            if current.startswith("combined_nesting"):
                self.var_out_name.set(f"combined_nesting_{sanitized}")
            if not self.var_title.get() or "SET" in self.var_title.get() or self.var_title.get().startswith("PROJECT:"):
                self.var_title.set(f"{proj_name.upper()} - {sheet_name.upper()}")

    def _auto_detect_defaults(self):
        """Auto-detect workspace files if present."""
        cwd = pathlib.Path.cwd()
        default_ex = cwd / DEFAULT_EXCEL
        if default_ex.exists():
            self.var_excel.set(str(default_ex))
        elif (cwd / FALLBACK_EXCEL).exists():
            self.var_excel.set(str(cwd / FALLBACK_EXCEL))

        default_src = cwd / DEFAULT_SOURCE_DIR
        if default_src.exists():
            self.var_source.set(str(default_src))
        else:
            self.var_source.set(str(cwd))

        if (cwd / FALLBACK_EXCEL).exists():
            self.var_fallback.set(str(cwd / FALLBACK_EXCEL))

        self.var_out_dir.set(str(cwd))
        self.last_out_dir = cwd

        # Default drawing title
        self.var_title.set(f"{cwd.name.upper()} - NESTING LAYOUT")

        # Load sheets for default Excel
        self.reload_sheets(auto_select_first=True)

    def browse_excel(self):
        """Open file dialog to pick BOM Excel file."""
        initial = self.var_excel.get() or self.var_out_dir.get() or str(pathlib.Path.cwd())
        init_dir = str(pathlib.Path(initial).parent) if pathlib.Path(initial).is_file() else str(pathlib.Path(initial))
        path = filedialog.askopenfilename(
            title="Select BOM Excel File",
            initialdir=init_dir,
            filetypes=[("Excel Files", "*.xlsx;*.xls;*.xlsm"), ("All Files", "*.*")]
        )
        if path:
            self.var_excel.set(path)
            excel_p = pathlib.Path(path)

            # Refresh sheets immediately
            self.reload_sheets(auto_select_first=True)

            # Auto-suggest DXF source & Output folder if not set or default
            if not self.var_source.get() or self.var_source.get() == str(pathlib.Path.cwd()):
                candidate_dxf = excel_p.parent / "dxf_source"
                if candidate_dxf.exists():
                    self.var_source.set(str(candidate_dxf))
                else:
                    self.var_source.set(str(excel_p.parent))
            if not self.var_out_dir.get() or self.var_out_dir.get() == str(pathlib.Path.cwd()):
                self.var_out_dir.set(str(excel_p.parent))
            self.log(f"Selected BOM Excel: {path}", "info")

    def browse_source(self):
        """Open directory dialog to pick DXF parts source folder."""
        initial = self.var_source.get() or self.var_out_dir.get() or str(pathlib.Path.cwd())
        path = filedialog.askdirectory(
            title="Select DXF Parts Source Folder",
            initialdir=initial
        )
        if path:
            self.var_source.set(path)
            self.log(f"Selected DXF Source Folder: {path}", "info")

    def browse_fallback(self):
        """Open file dialog to pick template / fallback Excel file."""
        initial = self.var_fallback.get() or self.var_out_dir.get() or str(pathlib.Path.cwd())
        init_dir = str(pathlib.Path(initial).parent) if pathlib.Path(initial).is_file() else str(pathlib.Path(initial))
        path = filedialog.askopenfilename(
            title="Select Template / Fallback Excel File",
            initialdir=init_dir,
            filetypes=[("Excel Files", "*.xlsx;*.xls;*.xlsm"), ("All Files", "*.*")]
        )
        if path:
            self.var_fallback.set(path)
            self.log(f"Selected Fallback Excel: {path}", "info")

    def browse_out_dir(self):
        """Open directory dialog to pick output save folder."""
        initial = self.var_out_dir.get() or str(pathlib.Path.cwd())
        path = filedialog.askdirectory(
            title="Select Output Save Folder",
            initialdir=initial
        )
        if path:
            self.var_out_dir.set(path)
            self.last_out_dir = pathlib.Path(path)
            self.log(f"Selected Output Folder: {path}", "info")

    def browse_save_as(self):
        """Open Save As dialog to choose folder and base output name together."""
        initial_dir = self.var_out_dir.get() or str(pathlib.Path.cwd())
        initial_file = f"{self.var_out_name.get() or DEFAULT_OUTPUT_NAME}.dxf"
        path = filedialog.asksaveasfilename(
            title="Select Destination & Output Name",
            initialdir=initial_dir,
            initialfile=initial_file,
            defaultextension=".dxf",
            filetypes=[("DXF Nesting File", "*.dxf"), ("All Files", "*.*")]
        )
        if path:
            p = pathlib.Path(path)
            self.var_out_dir.set(str(p.parent))
            self.var_out_name.set(p.stem)
            self.last_out_dir = p.parent
            self.log(f"Target Save Location: Folder='{p.parent}', BaseName='{p.stem}'", "info")

    def run_wizard_dialogs(self):
        """Sequential step-by-step dialog picker wizard with Sheet & Title selection."""
        # Step 1: Excel File
        excel_path = filedialog.askopenfilename(
            title="[Step 1 of 4] Select BOM Excel File (*.xlsx)",
            filetypes=[("Excel Files", "*.xlsx;*.xls;*.xlsm"), ("All Files", "*.*")]
        )
        if not excel_path:
            return
        self.var_excel.set(excel_path)
        p_excel = pathlib.Path(excel_path)

        # Step 2: Choose Sheet / Set
        sheets = get_excel_sheet_names(excel_path)
        chosen_sheet = None
        if len(sheets) > 1:
            sheet_options = sheets + [ALL_SHEETS_OPTION]
            dlg = tk.Toplevel(self.root)
            dlg.title("Select Sheet / Vendor Set")
            dlg.geometry("420x240")
            dlg.resizable(False, False)
            dlg.transient(self.root)
            dlg.grab_set()

            tk.Label(
                dlg,
                text=f"The Excel file '{p_excel.name}' contains {len(sheets)} sheets.\nChoose which sheet / vendor set to convert:",
                font=("Segoe UI", 9),
                wraplength=380,
                justify=tk.LEFT
            ).pack(padx=16, pady=(16, 8), anchor=tk.W)

            var_dlg_sheet = tk.StringVar(value=sheets[0])
            cbo_dlg = ttk.Combobox(dlg, textvariable=var_dlg_sheet, values=sheet_options, state="readonly", font=("Segoe UI", 10))
            cbo_dlg.pack(padx=16, pady=6, fill=tk.X)

            tk.Label(dlg, text="Drawing Title Heading (Centered at top of CAD):", font=("Segoe UI", 9)).pack(padx=16, pady=(6, 2), anchor=tk.W)
            var_dlg_title = tk.StringVar(value=f"{p_excel.parent.name.upper()} - {sheets[0].upper()}")
            ent_dlg_title = ttk.Entry(dlg, textvariable=var_dlg_title, font=("Segoe UI", 9))
            ent_dlg_title.pack(padx=16, pady=4, fill=tk.X)

            def _on_dlg_sheet_change(e):
                s = var_dlg_sheet.get()
                var_dlg_title.set(f"{p_excel.parent.name.upper()} - {s.upper()}")
            cbo_dlg.bind("<<ComboboxSelected>>", _on_dlg_sheet_change)

            selected_ref = {"sheet": None, "title": None}

            def _on_ok():
                selected_ref["sheet"] = var_dlg_sheet.get()
                selected_ref["title"] = var_dlg_title.get()
                dlg.destroy()

            def _on_cancel():
                dlg.destroy()

            btn_box = ttk.Frame(dlg)
            btn_box.pack(pady=(12, 16))
            ttk.Button(btn_box, text="Proceed", command=_on_ok, width=12).pack(side=tk.LEFT, padx=6)
            ttk.Button(btn_box, text="Cancel", command=_on_cancel, width=10).pack(side=tk.LEFT, padx=6)

            dlg.wait_window()
            chosen_sheet = selected_ref["sheet"]
            if not chosen_sheet:
                return
            if selected_ref["title"]:
                self.var_title.set(selected_ref["title"])
        elif len(sheets) == 1:
            chosen_sheet = sheets[0]
            self.var_title.set(f"{p_excel.parent.name.upper()} - {sheets[0].upper()}")

        if chosen_sheet:
            self.var_sheet.set(chosen_sheet)
            self._update_output_name_and_title(chosen_sheet)

        self.reload_sheets()

        # Step 3: DXF Source Folder
        def_source = p_excel.parent / "dxf_source"
        init_src = str(def_source) if def_source.exists() else str(p_excel.parent)
        dxf_dir = filedialog.askdirectory(
            title="[Step 3 of 4] Select DXF Parts Folder",
            initialdir=init_src
        )
        if not dxf_dir:
            return
        self.var_source.set(dxf_dir)

        # Step 4: Save Output Folder
        save_dir = filedialog.askdirectory(
            title="[Step 4 of 4] Select Destination Folder to Save DXF & DWG",
            initialdir=str(p_excel.parent)
        )
        if not save_dir:
            return
        self.var_out_dir.set(save_dir)
        self.last_out_dir = pathlib.Path(save_dir)

        # Confirm & Run
        sheet_desc = f"• Sheet / Set: {self.var_sheet.get() or '(Default)'}\n"
        title_desc = f"• Drawing Title: {self.var_title.get()}\n" if self.var_title.get() else ""
        need_dxf_v = self.var_need_dxf.get()
        need_dwg_v = self.var_need_dwg.get()
        if need_dxf_v and need_dwg_v:
            fmt_desc = "• Output Format: .DXF & .DWG\n"
            confirm_action = "generate nesting (.DXF & .DWG) now!"
        elif need_dwg_v:
            fmt_desc = "• Output Format: .DWG only\n"
            confirm_action = "generate nesting (.DWG) now!"
        else:
            fmt_desc = "• Output Format: .DXF only\n"
            confirm_action = "generate nesting (.DXF) now!"

        ans = messagebox.askyesno(
            "Ready to Build Nesting",
            f"Run nesting with selected options?\n\n"
            f"• Excel File: {p_excel.name}\n"
            f"{sheet_desc}"
            f"{title_desc}"
            f"• DXF Folder: {pathlib.Path(dxf_dir).name}\n"
            f"• Output Folder: {save_dir}\n"
            f"{fmt_desc}\n"
            f"Click Yes to {confirm_action}"
        )
        if ans:
            self.start_build_thread()

    def reset_fields(self):
        """Reset fields to defaults."""
        self.var_excel.set("")
        self.var_sheet.set("")
        self.var_title.set("")
        self.var_source.set("")
        self.var_fallback.set("")
        self.var_out_dir.set("")
        self.var_out_name.set(DEFAULT_OUTPUT_NAME)
        self.var_need_dxf.set(True)
        self.var_need_dwg.set(False)
        self._on_format_changed()
        self._auto_detect_defaults()
        self.log("Fields reset to workspace defaults.", "info")

    def open_output_folder(self):
        """Open the output destination in Windows Explorer."""
        folder = self.var_out_dir.get() or str(pathlib.Path.cwd())
        p = pathlib.Path(folder)
        if p.exists():
            os.startfile(str(p))
        else:
            messagebox.showwarning("Folder Not Found", f"Directory does not exist:\n{folder}")

    def start_build_thread(self):
        """Validates inputs and starts nesting in background thread."""
        if self.is_running:
            return

        excel_val = self.var_excel.get().strip()
        sheet_val = self.var_sheet.get().strip() or None
        title_val = self.var_title.get().strip() or None
        source_val = self.var_source.get().strip()
        out_dir_val = self.var_out_dir.get().strip()
        out_name_val = self.var_out_name.get().strip() or DEFAULT_OUTPUT_NAME
        fallback_val = self.var_fallback.get().strip() or None
        need_dxf_val = self.var_need_dxf.get()
        need_dwg_val = self.var_need_dwg.get()

        if not excel_val:
            messagebox.showerror("Missing Input", "Please select a BOM Excel file (*.xlsx) first.")
            return

        if not source_val:
            messagebox.showerror("Missing Input", "Please select a DXF Parts source folder.")
            return

        if not need_dxf_val and not need_dwg_val:
            messagebox.showerror(
                "No Output Format Selected",
                "Please select at least one output format:\n  ☑ .dxf needed  or  ☑ .dwg needed"
            )
            return

        if not out_dir_val:
            out_dir_val = str(pathlib.Path(excel_val).parent)
            self.var_out_dir.set(out_dir_val)

        self.last_out_dir = pathlib.Path(out_dir_val)

        self.is_running = True
        self.btn_run.config(state=tk.DISABLED, text="⏳ Processing Nesting...")
        self.var_status.set("Processing nesting... please wait.")
        self.txt_log.delete("1.0", tk.END)

        def worker():
            try:
                result = run_nesting_process(
                    excel_path=excel_val,
                    sheet_name=sheet_val,
                    drawing_title=title_val,
                    source_dir=source_val,
                    output_dir=out_dir_val,
                    output_name=out_name_val,
                    need_dxf=need_dxf_val,
                    need_dwg=need_dwg_val,
                    fallback_excel=fallback_val,
                    log_func=self.log
                )
                self.root.after(0, lambda: self._on_success(result))
            except Exception as e:
                self.root.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, result: dict):
        self.is_running = False
        # Restore btn label based on format selection
        self._on_format_changed()
        self.btn_run.config(state=tk.NORMAL)

        placed = result["placed"]
        failed = result["failed"]

        if result.get("is_batch"):
            count_sheets = result.get("sheets_processed", 1)
            status_msg = f"Batch Complete: {count_sheets} sheets processed ({placed} parts placed, {failed} issues)."
            self.var_status.set(status_msg)
            self.log(f"✓ SUCCESS: {status_msg}", "success")

            messagebox.showinfo(
                "Batch Nesting Completed Successfully!",
                f"Batch nesting for all sheets completed!\n\n"
                f"• Sheets Processed: {count_sheets}\n"
                f"• Total Parts Placed: {placed}\n"
                f"• Total Issues: {failed}\n\n"
                f"Files Saved in:\n{pathlib.Path(self.var_out_dir.get()).resolve()}"
            )
        else:
            out_dxf = result.get("out_dxf")
            out_dwg = result.get("out_dwg")
            sheet_info = f" (Sheet: '{result.get('sheet_name', 'BOM')}')"
            status_msg = f"Finished{sheet_info}: {placed} parts placed, {failed} issues."
            self.var_status.set(status_msg)
            self.log(f"✓ SUCCESS: {status_msg}", "success")

            # Build file summary lines only for generated files
            file_lines = []
            save_folder = None
            if out_dxf and out_dxf.exists():
                file_lines.append(f"• DXF: {out_dxf.name}")
                save_folder = out_dxf.parent
            if out_dwg and out_dwg.exists():
                file_lines.append(f"• DWG: {out_dwg.name}")
                if not save_folder:
                    save_folder = out_dwg.parent
            elif result.get("need_dwg") and not (out_dwg and out_dwg.exists()):
                file_lines.append("(DWG conversion skipped — AutoCAD accoreconsole not found)")

            if not save_folder:
                save_folder = pathlib.Path(self.var_out_dir.get()).resolve()

            files_summary = "\n".join(file_lines) if file_lines else "(No output files)"

            messagebox.showinfo(
                "Nesting Completed Successfully!",
                f"Nesting generation completed for sheet '{result.get('sheet_name', 'BOM')}'!\n\n"
                f"• Placed Parts: {placed}\n"
                f"• Issues: {failed}\n\n"
                f"Output Files Saved to:\n{save_folder}\n\n"
                f"{files_summary}"
            )

    def _on_error(self, err_msg: str):
        self.is_running = False
        # Restore btn label based on format selection
        self._on_format_changed()
        self.btn_run.config(state=tk.NORMAL)
        self.var_status.set(f"Error: {err_msg}")
        self.log(f"[ERROR] {err_msg}", "error")
        messagebox.showerror("Error During Nesting Generation", f"An error occurred:\n\n{err_msg}")


# ---------------------------------------------------------------------------
def launch_gui():
    """Launch the interactive dialog-based Tkinter GUI window."""
    root = tk.Tk()
    app = NestingAppGUI(root)
    root.mainloop()


# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="AutoCAD Nesting Sheet Builder (GUI Dialogs & CLI with Title & Multi-Sheet support)."
    )
    parser.add_argument(
        "--gui", "-g",
        action="store_true",
        help="Launch the interactive UI dialog window (default when run without arguments)."
    )
    parser.add_argument(
        "--wizard", "-w",
        action="store_true",
        help="Launch the step-by-step dialog picker wizard directly."
    )
    parser.add_argument(
        "--cli",
        action="store_true",
        help="Run directly in command-line mode without GUI."
    )
    parser.add_argument(
        "--project-dir", "-p",
        type=str,
        default=".",
        help="Target project directory path (default: current folder)."
    )
    parser.add_argument(
        "--batch", "-b",
        nargs="+",
        type=str,
        help="List of multiple project directories to process sequentially in batch."
    )
    parser.add_argument(
        "--excel", "-e",
        type=str,
        help="Custom Excel BOM filename/path."
    )
    parser.add_argument(
        "--sheet", "-sh",
        type=str,
        help="Worksheet / Vendor Set name to convert (use '__ALL__' for batch all sheets)."
    )
    parser.add_argument(
        "--title", "-t",
        type=str,
        help="Drawing Sheet Header / Title stamped at top of CAD drawing."
    )
    parser.add_argument(
        "--fallback-excel", "-f",
        type=str,
        help="Custom fallback Excel BOM template filename/path."
    )
    parser.add_argument(
        "--source", "-s",
        type=str,
        help="Custom source DXF folder name/path."
    )
    parser.add_argument(
        "--out-dir",
        type=str,
        help="Destination directory to save output DXF and DWG files."
    )
    parser.add_argument(
        "--out-name", "-o",
        type=str,
        help=f"Base name for output files (default: {DEFAULT_OUTPUT_NAME}). Auto-increments if exists."
    )
    parser.add_argument(
        "--format",
        type=str,
        choices=["dxf", "dwg", "both"],
        default=None,
        help="Output format: 'dxf' (default), 'dwg', or 'both'. Overrides --dxf/--dwg flags."
    )
    parser.add_argument(
        "--dxf",
        action="store_true",
        default=False,
        help="Save output as .DXF file (can be combined with --dwg)."
    )
    parser.add_argument(
        "--dwg",
        action="store_true",
        default=False,
        help="Save output as .DWG file via AutoCAD accoreconsole (can be combined with --dxf)."
    )

    args = parser.parse_args()

    # Resolve format flags
    if args.format == "both":
        cli_need_dxf, cli_need_dwg = True, True
    elif args.format == "dwg":
        cli_need_dxf, cli_need_dwg = False, True
    elif args.format == "dxf":
        cli_need_dxf, cli_need_dwg = True, False
    elif args.dwg and args.dxf:
        cli_need_dxf, cli_need_dwg = True, True
    elif args.dwg:
        cli_need_dxf, cli_need_dwg = False, True
    elif args.dxf:
        cli_need_dxf, cli_need_dwg = True, False
    else:
        cli_need_dxf, cli_need_dwg = True, False  # Default: DXF only

    # If --wizard is requested, launch root and wizard
    if args.wizard:
        root = tk.Tk()
        app = NestingAppGUI(root)
        root.after(100, app.run_wizard_dialogs)
        root.mainloop()
        return

    # If CLI flags were given, run in CLI / Batch mode
    if args.cli or args.batch or args.excel or args.source or args.out_dir or args.sheet or args.title:
        if args.batch:
            print(f"Batch Processing {len(args.batch)} Project Directories...")
            for proj_path in args.batch:
                p = pathlib.Path(proj_path)
                if p.exists() and p.is_dir():
                    process_project(
                        p,
                        excel_file=args.excel,
                        sheet_name=args.sheet,
                        drawing_title=args.title,
                        fallback_excel=args.fallback_excel,
                        source_dir_name=args.source,
                        output_name=args.out_name,
                        need_dxf=cli_need_dxf,
                        need_dwg=cli_need_dwg
                    )
                else:
                    print(f"Warning: Directory not found: {proj_path}")
        else:
            target_dir = pathlib.Path(args.project_dir)
            if args.out_dir:
                run_nesting_process(
                    excel_path=args.excel or (target_dir / DEFAULT_EXCEL),
                    sheet_name=args.sheet,
                    drawing_title=args.title,
                    source_dir=args.source or (target_dir / DEFAULT_SOURCE_DIR),
                    output_dir=args.out_dir,
                    output_name=args.out_name or DEFAULT_OUTPUT_NAME,
                    need_dxf=cli_need_dxf,
                    need_dwg=cli_need_dwg,
                    fallback_excel=args.fallback_excel
                )
            else:
                process_project(
                    target_dir,
                    excel_file=args.excel,
                    sheet_name=args.sheet,
                    drawing_title=args.title,
                    fallback_excel=args.fallback_excel,
                    source_dir_name=args.source,
                    output_name=args.out_name,
                    need_dxf=cli_need_dxf,
                    need_dwg=cli_need_dwg
                )
        return

    # Default: Launch the Interactive UI Window!
    launch_gui()


if __name__ == "__main__":
    main()
