"""
Export Updated Format - Generates a single-sheet Excel with the new layout:

Header Row 1: "New Formula List" (cols A-K) | "Available Material Details" (cols L-U)
Header Row 2: SR.NO | PANEL NAME | WL | WR | CODE | LL | LR | ELEMENT TYPE | LOCATION | QTY/AREA/ITEM | AREA | Condition | Width | Code | Length | Type | Quantity | Remainder | Remark | Additional Remark | Panel Used

Green rows = AS IT IS (reuse)
White rows  = MODIFY
No color    = NEW (unmatched)
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO


def generate_updated_export(processed_path: str, output_path: str):
    """
    Read the existing processed Excel and generate the Updated Format export.
    """
    # Load all needed sheets from the processed file
    try:
        df_reuse = pd.read_excel(processed_path, sheet_name="REUSED")
    except Exception:
        df_reuse = pd.DataFrame()

    try:
        df_modify = pd.read_excel(processed_path, sheet_name="MODIFY")
    except Exception:
        df_modify = pd.DataFrame()

    try:
        df_new = pd.read_excel(processed_path, sheet_name="NEW")
    except Exception:
        df_new = pd.DataFrame()

    # Build unified rows
    # Target columns for the updated format:
    # A: SR.NO
    # B: PANEL NAME   -> New Spec
    # C: WL           -> New WL
    # D: WR           -> New WR
    # E: CODE         -> New CODE
    # F: LL           -> New LL
    # G: LR           -> New LR
    # H: ELEMENT TYPE -> New Flat
    # I: LOCATION     -> New Location
    # J: QTY/AREA/ITEM-> New QTY
    # K: AREA         -> calculated area
    # L: Condition    -> "AS IT IS" or "MODIFY" or empty
    # M: Width        -> Old WR
    # N: Code         -> Old CODE
    # O: Length       -> Old LL
    # P: Type         -> Old Location (Cleaned/Uncleaned)
    # Q: Quantity     -> Old QTY
    # R: Remainder    -> Remainder spec
    # S: Remark       -> Remark
    # T: Additional Remark -> Source_Detail or extra remark info
    # U: Panel Used   -> Remainder_Used

    rows = []
    row_types = []  # Track 'REUSE', 'MODIFY', 'NEW' for coloring

    def safe_val(df_row, col, default=""):
        if isinstance(df_row, dict):
            v = df_row.get(col)
            if v is None or pd.isna(v):
                return default
            return v
        if hasattr(df_row, 'index') and col in df_row.index:
            v = df_row[col]
            if pd.isna(v):
                return default
            return v
        return default

    def calc_area(wl, wr, ll, lr):
        """Calculate area in m^2"""
        try:
            w = max(int(wl or 0), int(wr or 0))
            l = max(int(ll or 0), int(lr or 0))
            return round((w * l) / 1_000_000, 4)
        except (ValueError, TypeError):
            return 0.0

    def format_dim(a, b):
        """Format dimension: if both non-zero show 'a+b', else show whichever is non-zero."""
        try:
            va = int(a) if a and str(a).strip() and not pd.isna(a) else 0
            vb = int(b) if b and str(b).strip() and not pd.isna(b) else 0
        except (ValueError, TypeError):
            va, vb = 0, 0
        if va > 0 and vb > 0:
            return f"{va}+{vb}"
        elif va > 0:
            return va
        elif vb > 0:
            return vb
        return ""

    sr_no = 1

    # --- REUSE rows ---
    if not df_reuse.empty:
        for row in df_reuse.to_dict('records'):
            new_spec = safe_val(row, "New Spec")
            new_wl = safe_val(row, "New WL", 0)
            new_wr = safe_val(row, "New WR", 0)
            new_code = safe_val(row, "New CODE")
            new_ll = safe_val(row, "New LL", 0)
            new_lr = safe_val(row, "New LR", 0)
            new_flat = safe_val(row, "New Flat")
            new_loc = safe_val(row, "New Location")
            new_qty = safe_val(row, "New QTY", 1)
            area_val = calc_area(new_wl, new_wr, new_ll, new_lr)

            old_wl = safe_val(row, "Old WL", 0)
            old_wr = safe_val(row, "Old WR", 0)
            old_code = safe_val(row, "Old CODE", "")
            old_ll = safe_val(row, "Old LL", 0)
            old_lr = safe_val(row, "Old LR", 0)
            old_loc = safe_val(row, "Old Location", "")
            old_qty = safe_val(row, "Old QTY", 1)

            remark = safe_val(row, "Remark", "")

            # Format width and length with WL+WR / LL+LR
            display_width = format_dim(old_wl, old_wr)
            display_length = format_dim(old_ll, old_lr)

            rows.append([
                sr_no, new_spec, new_wl, new_wr, new_code, new_ll, new_lr,
                new_flat, new_loc, new_qty, area_val,
                "AS IT IS",
                display_width, old_code, display_length, old_loc, old_qty,
                "",  # Remainder
                remark,  # Remark
                "",  # Additional Remark
                ""   # Panel Used
            ])
            row_types.append("REUSE")
            sr_no += 1

    # --- MODIFY rows ---
    if not df_modify.empty:
        for row in df_modify.to_dict('records'):
            new_spec = safe_val(row, "New Spec")
            new_wl = safe_val(row, "New WL", 0)
            new_wr = safe_val(row, "New WR", 0)
            new_code = safe_val(row, "New CODE")
            new_ll = safe_val(row, "New LL", 0)
            new_lr = safe_val(row, "New LR", 0)
            new_flat = safe_val(row, "New Flat")
            new_loc = safe_val(row, "New Location")
            new_qty = safe_val(row, "New QTY", 1)
            area_val = calc_area(new_wl, new_wr, new_ll, new_lr)

            old_wl = safe_val(row, "Old WL", 0)
            old_wr = safe_val(row, "Old WR", 0)
            old_code = safe_val(row, "Old CODE", "")
            old_ll = safe_val(row, "Old LL", 0)
            old_lr = safe_val(row, "Old LR", 0)
            old_loc = safe_val(row, "Old Location", "")
            old_qty = safe_val(row, "Old QTY", 1)

            remainder = safe_val(row, "Remainder", "")
            remainder_used = safe_val(row, "Remainder_Used", "")
            remark = safe_val(row, "Remark", "")
            source = safe_val(row, "Source", "")
            source_detail = safe_val(row, "Source_Detail", "")

            # Format width and length with WL+WR / LL+LR
            display_width = format_dim(old_wl, old_wr)
            display_length = format_dim(old_ll, old_lr)

            # Build additional remark from source info
            additional_remark = ""
            if source == "FROM REMAINDER":
                additional_remark = f"MODIFIED FROM REMAINDER"
                if source_detail:
                    additional_remark += f"\n{source_detail}"
            elif source_detail:
                additional_remark = str(source_detail)

            rows.append([
                sr_no, new_spec, new_wl, new_wr, new_code, new_ll, new_lr,
                new_flat, new_loc, new_qty, area_val,
                "MODIFY",
                display_width, old_code, display_length, old_loc, old_qty,
                remainder,
                remark,
                additional_remark,
                remainder_used
            ])
            row_types.append("MODIFY")
            sr_no += 1

    # --- NEW (unmatched) rows ---
    if not df_new.empty:
        for row in df_new.to_dict('records'):
            new_spec = safe_val(row, "New Spec")
            new_wl = safe_val(row, "New WL", 0)
            new_wr = safe_val(row, "New WR", 0)
            new_code = safe_val(row, "New CODE")
            new_ll = safe_val(row, "New LL", 0)
            new_lr = safe_val(row, "New LR", 0)
            new_flat = safe_val(row, "New Flat")
            new_loc = safe_val(row, "New Location")
            new_qty = safe_val(row, "New QTY", 1)
            area_val = calc_area(new_wl, new_wr, new_ll, new_lr)

            rows.append([
                sr_no, new_spec, new_wl, new_wr, new_code, new_ll, new_lr,
                new_flat, new_loc, new_qty, area_val,
                "",  # Condition - blank for NEW
                "", "", "", "", "",
                "",  # Remainder
                "",  # Remark
                "",  # Additional Remark
                ""   # Panel Used
            ])
            row_types.append("NEW")
            sr_no += 1

    # Column headers for data rows
    col_headers = [
        "SR.NO", "PANEL NAME", "WL", "WR", "CODE", "LL", "LR",
        "ELEMENT TYPE", "LOCATION", "QTY/AREA/ITEM", "AREA",
        "Condition", "Width", "Code", "Length", "Type", "Quantity",
        "Remainder", "Remark", "Additional Remark", "Panel Used"
    ]

    # Create DataFrame and write to Excel
    df_out = pd.DataFrame(rows, columns=col_headers)

    # Write using openpyxl for full formatting control
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "NEW"

    # --- Styles ---
    green_header = PatternFill(start_color="4CAF50", end_color="4CAF50", fill_type="solid")
    yellow_header = PatternFill(start_color="FFEB3B", end_color="FFEB3B", fill_type="solid")
    green_row = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    light_yellow_row = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    header_font = Font(name='Calibri', size=11, bold=True, color="000000")
    group_header_font = Font(name='Calibri', size=12, bold=True, color="FFFFFF")
    data_font = Font(name='Calibri', size=10)
    green_code_font = Font(name='Calibri', size=10, bold=True, color="006100")

    thin = Side(border_style='thin', color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=False)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # --- Row 1: Group Headers ---
    # "New Formula List" spans A1:K1
    ws.merge_cells('A1:K1')
    cell_nfl = ws.cell(1, 1, "New Formula List")
    cell_nfl.font = group_header_font
    cell_nfl.fill = yellow_header
    cell_nfl.alignment = center_align
    cell_nfl.border = border

    # "Available Material Details" spans L1:U1
    ws.merge_cells('L1:U1')
    cell_amd = ws.cell(1, 12, "Available Material Details")
    cell_amd.font = group_header_font
    cell_amd.fill = green_header
    cell_amd.alignment = center_align
    cell_amd.border = border

    # Apply borders to all merged cells in row 1
    for c in range(1, 22):
        cell = ws.cell(1, c)
        cell.border = border
        if c <= 11:
            cell.fill = yellow_header
        else:
            cell.fill = green_header

    # --- Row 2: Column Headers ---
    for c_idx, header in enumerate(col_headers, start=1):
        cell = ws.cell(2, c_idx, header)
        cell.font = header_font
        cell.alignment = center_align
        cell.border = border
        if c_idx <= 11:
            cell.fill = PatternFill(start_color="FFF176", end_color="FFF176", fill_type="solid")
        else:
            cell.fill = PatternFill(start_color="81C784", end_color="81C784", fill_type="solid")

    # --- Data Rows (start at row 3) ---
    for r_idx, (data_row, row_type) in enumerate(zip(rows, row_types), start=3):
        for c_idx, val in enumerate(data_row, start=1):
            cell = ws.cell(r_idx, c_idx, val)
            cell.font = data_font
            cell.alignment = center_align
            cell.border = border

            if row_type == "REUSE":
                cell.fill = green_row
            elif row_type == "MODIFY":
                cell.fill = white_fill
            else:
                cell.fill = white_fill

        # Highlight matched code/width/length cells in green for MODIFY rows
        if row_type == "MODIFY":
            # Highlight columns M (Width), N (Code), O (Length) if they have values
            for highlight_col in [13, 14, 15]:  # M, N, O
                c = ws.cell(r_idx, highlight_col)
                if c.value and str(c.value).strip():
                    c.fill = PatternFill(start_color="A5D6A7", end_color="A5D6A7", fill_type="solid")
                    c.font = green_code_font

    # --- Column Widths ---
    col_widths = {
        'A': 7,   # SR.NO
        'B': 16,  # PANEL NAME
        'C': 6,   # WL
        'D': 6,   # WR
        'E': 7,   # CODE
        'F': 6,   # LL
        'G': 5,   # LR
        'H': 14,  # ELEMENT TYPE
        'I': 10,  # LOCATION
        'J': 14,  # QTY/AREA/ITEM
        'K': 8,   # AREA
        'L': 12,  # Condition
        'M': 8,   # Width
        'N': 7,   # Code
        'O': 8,   # Length
        'P': 12,  # Type
        'Q': 10,  # Quantity
        'R': 16,  # Remainder
        'S': 20,  # Remark
        'T': 28,  # Additional Remark
        'U': 12,  # Panel Used
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # Freeze panes (freeze row 1 and 2)
    ws.freeze_panes = "A3"

    # --- Add OLD sheet from processed file ---
    try:
        wb_processed = load_workbook(BytesIO(open(processed_path, "rb").read()), data_only=True)

        # Copy UNUSED STOCK sheet
        if "UNUSED STOCK" in wb_processed.sheetnames:
            ws_unused_src = wb_processed["UNUSED STOCK"]
            ws_unused = wb.create_sheet("OLD")
            for r_idx, row in enumerate(ws_unused_src.iter_rows(values_only=False), start=1):
                for c_idx, cell in enumerate(row, start=1):
                    new_cell = ws_unused.cell(r_idx, c_idx, cell.value)
                    new_cell.font = data_font
                    new_cell.alignment = center_align
                    new_cell.border = border
            # Format header row
            if ws_unused.max_row and ws_unused.max_row >= 1:
                for c in range(1, ws_unused.max_column + 1):
                    h_cell = ws_unused.cell(1, c)
                    h_cell.font = header_font
                    h_cell.fill = PatternFill(start_color="81C784", end_color="81C784", fill_type="solid")

        # Copy SUMMARY sheet
        if "SUMMARY" in wb_processed.sheetnames:
            ws_summ_src = wb_processed["SUMMARY"]
            ws_summ = wb.create_sheet("SUMMARY")
            for r_idx, row in enumerate(ws_summ_src.iter_rows(values_only=False), start=1):
                for c_idx, cell in enumerate(row, start=1):
                    new_cell = ws_summ.cell(r_idx, c_idx, cell.value)
                    new_cell.font = data_font
                    new_cell.alignment = center_align
                    new_cell.border = border
            # Format header row
            if ws_summ.max_row and ws_summ.max_row >= 1:
                for c in range(1, ws_summ.max_column + 1):
                    h_cell = ws_summ.cell(1, c)
                    h_cell.font = header_font
                    h_cell.fill = PatternFill(start_color="4FC3F7", end_color="4FC3F7", fill_type="solid")

        wb_processed.close()
    except Exception as e:
        print(f"(Could not copy additional sheets: {e})")

    wb.save(output_path)
    return output_path
