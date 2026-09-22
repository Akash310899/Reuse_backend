"""
Reuse Modify Tool - Automation Script

This script processes Excel files with "New" and "Old" sheets containing panel specifications.
It matches new panel requirements with old inventory for reuse, modification, or new purchase decisions.

Features:
- AS-IS reuse matching (exact dimensions and base code)
- Alternative base code matching (ALT)
- Vertical/Horizontal swap matching (VV)
- RK tolerance matching (length adjustments)
- Modify/Cut operations with remainder tracking
- Location-based summary breakdown
- Highlighted output sheets showing matched items

Business-rule notes encoded in this version (kept here so future edits don't
silently drop them):

  1) WX  -> T,B,D,PC,WS,WX (cut). Cutting only "pays off" if the resulting
     scrap piece is small (<= MIN_USABLE_REMAINDER_MM) or, when larger, can
     itself be consumed by another New row (see can_consume_remainder()).
  2) WE  -> WE, WSE (remark "ATTACH RK"), WS/W/T/B/D/PC (remark "ATTACH EC").
  3) WS  -> WS, W (remark "REMOVE RK", RK_MAX=50mm tolerance), WX/T/B,
     D/PC (remark "NOTCHING").
  4) W   -> W as-is; WS/WX/T/B/D/PC all get remark "ATTACH RK".
  5) T   -> WX,B,WS,D,PC,T,TX,WT.
  6) B   -> B,T,W,WX,WS,D,PC,TX,WT.
  7) EC family (any code starting "EC") ignores WL/WR entirely for BOTH the
     reuse stage and the modify stage - only CODE-starts-with-EC and LL are
     compared.
  8-10) ICB/ICK/ICT: WL & WR must match exactly (never cut), only LL is cut.
     ICB can only be cut from inventory code "ICB" itself; ICK/ICT can be
     cut from ICB/ICK/ICT/BIC.
  11) WXE -> WS,T,B,WE,WSE,WXE,TE,W (W gets "ATTACH EC & REMOVE RK";
     WS/T/B get "ATTACH EC").
  12) BIC -> same rules as ICB/ICK/ICT, plus ICX and BC are also valid
     inventory source codes.
  13-14) D  can be modified from D, PC, DG. DG can be modified from D, PC, DG.
  14) SC*  families (SCJ/SCU/SCA/SCF/SCR ...): exact WL/WR/CODE match, LL
     must exceed tolerance.
  15) SI*  families: exact WL/WR/CODE match, LL and LR must both exceed
     tolerance.
  16) SO*  families: same as SI*.
  17) DE  -> DE, D/PC (remark "ATTACH EC").
  18) PH  -> CP, PH, PPH.
  19) PLB -> PLB only.
  20) K   -> can be cut from inventory code K or B (WL/WR identical, LL must
     exceed tolerance).
  21) KC  -> can be cut from inventory code KC or BC (WL/WR identical, LL
     and LR must both exceed tolerance).
  22) KCE -> can only be cut from KCE (WL/WR identical, LL and LR must both
     exceed tolerance).

  Cutting policy (applies globally, see MIN_USABLE_REMAINDER_MM /
  MAX_ACCEPTABLE_SCRAP_FRACTION below): prefer cuts that leave a remainder
  under ~200mm (effectively no usable scrap). A larger remainder is only
  acceptable if it stays under ~20% of the panel's own area, OR if a later
  New row can actually consume that remainder (see can_consume_remainder).

  Cut-mode restriction: 'Length' and 'Width' single-axis cutting are only
  meaningful for the base codes W, WS, T, B, WE, WSE, TE, D, WXE. For every
  other base code, a 'Length' or 'Width' selection silently falls back to
  'Both' (cut on whichever axis is actually oversized).

Author: Automation Tool
Updated: Business-rule alignment pass
"""
import sys
import time
from collections import defaultdict
import pandas as pd

from backend.core.constants import *
from backend.core.utils import *
from backend.core.excel_io import *

def process_excel(in_path, out_path, progress_callback=None, tolerance=10, cut_mode='Both', operation_mode='Both', swap_modify=True):
    """
    Main processing function for reuse/modify automation.
    
    Args:
        in_path: Path to input Excel file (must have "New" and "Old" sheets)
        out_path: Path to output Excel file
        progress_callback: Optional callback function(percent, message) for progress updates
        tolerance: Tolerance in mm for modify operations (default: 10)
        cut_mode: Cut mode - 'Length', 'Width', or 'Both' (default: 'Both')
        operation_mode: Operation mode - 'Both', 'REUSE Only', or 'MODIFY Only' (default: 'Both')
    
    Returns:
        None (writes results to out_path)
    """
    def update_progress(p, msg=""):
        if progress_callback:
            progress_callback(p, msg)

    t0 = time.time()
    update_progress(0, "Reading Excel sheets...")

    # Read sheet names and find case-insensitive matches
    xl = pd.ExcelFile(in_path)
    sheet_names = xl.sheet_names
    new_sheet = next((s for s in sheet_names if s.lower() == "new"), None)
    old_sheet = next((s for s in sheet_names if s.lower() == "old"), None)
    
    if not new_sheet or not old_sheet:
        raise ValueError("Excel file must contain sheets named 'New' and 'Old' (case-insensitive).")
    
    new = pd.read_excel(in_path, sheet_name=new_sheet, dtype=str)
    old = pd.read_excel(in_path, sheet_name=old_sheet, dtype=str)

    update_progress(10, "Preparing data...")

    cmap_new = build_col_map(new)
    cmap_old = build_col_map(old)

    new_spec_col = find_col(cmap_new, ["spec", "specification", "description"])
    new_flat_col = find_col(cmap_new, ["flat", "flat ", "flatname"])
    new_part_col = find_col(cmap_new, ["part", "parttype", "part type", "part_name"])
    new_qty_col = find_col(cmap_new, ["qty", "quantity", "qnty"])
    new_unit_col = find_col(cmap_new, ["unit", "unitname", "unit area", "unit_area"])
    # Location / flat name column (if present)
    new_loc_col = find_col(cmap_new, ["location", "loc", "locn"])

    old_spec_col = find_col(cmap_old, ["spec", "specification", "description"])
    old_qty_col = find_col(cmap_old, ["qty", "quantity", "qnty"])
    old_unit_area_col = find_col(cmap_old, ["unit area", "unitarea", "unit_area", "unit"])
    old_loc_col = find_col(cmap_old, ["location", "loc", "locn"])

    wl_col_new = find_col(cmap_new, ["wl", "widthleft", "width_l", "w_l"])
    wr_col_new = find_col(cmap_new, ["wr", "widthright", "width_r", "w_r"])
    ll_col_new = find_col(cmap_new, ["ll", "lengthleft", "length_l", "l_l"])
    lr_col_new = find_col(cmap_new, ["lr", "lengthright", "length_r", "l_r"])
    code_col_new = find_col(cmap_new, ["code", "code ", "type", "code_name", "c_o_d_e"])

    wl_col_old = find_col(cmap_old, ["wl", "widthleft", "width_l", "w_l"])
    wr_col_old = find_col(cmap_old, ["wr", "widthright", "width_r", "w_r"])
    ll_col_old = find_col(cmap_old, ["ll", "lengthleft", "length_l", "l_l"])
    lr_col_old = find_col(cmap_old, ["lr", "lengthright", "length_r", "l_r"])
    code_col_old = find_col(cmap_old, ["code", "code ", "type", "code_name", "c_o_d_e"])

    # Handle missing columns by creating synthetic columns and updating references
    synthetic_counter = 0
    col_mappings = [
        ("wl_col_new", wl_col_new, "0", new), ("wr_col_new", wr_col_new, "0", new),
        ("ll_col_new", ll_col_new, "0", new), ("lr_col_new", lr_col_new, "0", new),
        ("code_col_new", code_col_new, "", new),
        ("wl_col_old", wl_col_old, "0", old), ("wr_col_old", wr_col_old, "0", old),
        ("ll_col_old", ll_col_old, "0", old), ("lr_col_old", lr_col_old, "0", old),
        ("code_col_old", code_col_old, "", old)
    ]
    
    for var_name, col_key, default_val, df in col_mappings:
        if col_key is None:
            if default_val == "":
                synthetic = "CODE_SYNTHETIC"
            else:
                synthetic_counter += 1
                synthetic = f"NUM_SYNTH_{int(time.time()*1000)}_{synthetic_counter}"
            df[synthetic] = default_val
            # Update the variable to point to the synthetic column
            if var_name == "wl_col_new":
                wl_col_new = synthetic
            elif var_name == "wr_col_new":
                wr_col_new = synthetic
            elif var_name == "ll_col_new":
                ll_col_new = synthetic
            elif var_name == "lr_col_new":
                lr_col_new = synthetic
            elif var_name == "code_col_new":
                code_col_new = synthetic
            elif var_name == "wl_col_old":
                wl_col_old = synthetic
            elif var_name == "wr_col_old":
                wr_col_old = synthetic
            elif var_name == "ll_col_old":
                ll_col_old = synthetic
            elif var_name == "lr_col_old":
                lr_col_old = synthetic
            elif var_name == "code_col_old":
                code_col_old = synthetic

    # Fallback to default column names if still None
    wl_col_new = wl_col_new or "WL"
    wr_col_new = wr_col_new or "WR"
    ll_col_new = ll_col_new or "LL"
    lr_col_new = lr_col_new or "LR"
    code_col_new = code_col_new or "CODE"

    wl_col_old = wl_col_old or "WL"
    wr_col_old = wr_col_old or "WR"
    ll_col_old = ll_col_old or "LL"
    lr_col_old = lr_col_old or "LR"
    code_col_old = code_col_old or "CODE"

    for col in [wl_col_new, wr_col_new, ll_col_new, lr_col_new]:
        if col not in new.columns:
            new[col] = "0"
        new[col] = read_dims(new[col])

    for col in [wl_col_old, wr_col_old, ll_col_old, lr_col_old]:
        if col not in old.columns:
            old[col] = "0"
        old[col] = read_dims(old[col])

    if code_col_new not in new.columns:
        new[code_col_new] = ""
    new[code_col_new] = new[code_col_new].fillna("").astype(str).apply(clean_code_raw)

    if code_col_old not in old.columns:
        old[code_col_old] = ""
    old[code_col_old] = old[code_col_old].fillna("").astype(str).apply(clean_code_raw)

    new["BASE"] = new[code_col_new].map(base_code)
    old["BASE"] = old[code_col_old].map(base_code)

    # ── BIG WIN #1: First-Fit Decreasing (FFD) sort ──────────────────────────
    # Sort New requirements by panel area largest-first before ANY matching.
    # This ensures large panels are matched first so their cut remainders are
    # available in the scrap pool for smaller rows downstream.
    # Industry research shows FFD alone improves bin-packing utilisation 8-15%.
    new["_ffd_area"] = (
        new[wl_col_new].astype(int) * new[wr_col_new].astype(int) +
        new[ll_col_new].astype(int) * new[lr_col_new].astype(int)
    )
    new = new.sort_values("_ffd_area", ascending=False).reset_index(drop=True)
    new = new.drop(columns=["_ffd_area"])
    # ─────────────────────────────────────────────────────────────────────────

    old_records = old.to_dict('records')
    reuse_idx = defaultdict(list)
    mod_idx = defaultdict(list)
    swap_rotate_idx = defaultdict(list)
    swap_flip_idx = defaultdict(list)
    inv_list = []
    inv_by_base = defaultdict(list)
    inv_by_code = defaultdict(list)
    old_meta_list = []

    for idx, r in enumerate(old_records):
        wl = int(r.get(wl_col_old, 0))
        wr = int(r.get(wr_col_old, 0))
        LL = int(r.get(ll_col_old, 0))
        LR = int(r.get(lr_col_old, 0))
        base = str(r.get("BASE", ""))
        code = str(r.get(code_col_old, "")) if pd.notna(r.get(code_col_old)) else ""

        reuse_idx[(wl, wr, base, LL, LR)].append(idx)
        mod_idx[(wl, wr, base)].append(idx)
        swap_rotate_idx[(wl, LL, wr, LR, base)].append(idx)
        swap_flip_idx[(wr, wl, LL, LR, base)].append(idx)

        spec = r.get(old_spec_col, "") if old_spec_col else ""
        qty = r.get(old_qty_col, "") if old_qty_col else ""
        unit_area = r.get(old_unit_area_col, "") if old_unit_area_col else ""
        loc = r.get(old_loc_col, "") if old_loc_col else ""
        old_meta_list.append((
            (str(spec) if pd.notna(spec) else ""),
            wl, wr, code, LL, LR,
            (str(loc) if pd.notna(loc) else ""),
            (str(qty) if pd.notna(qty) else ""),
            (str(unit_area) if pd.notna(unit_area) else "")
        ))

        inv_item = {
            "old_idx": idx,
            "WL": wl,
            "WR": wr,
            "LL": LL,
            "LR": LR,
            "CODE": code,
            "BASE": base,
            "USED": False
        }
        inv_list.append(inv_item)
        inv_by_base[base].append(inv_item)
        inv_by_code[code].append(inv_item)

    used_old = [False] * len(old)

    reuse_rows, modify_rows, none_rows, unused_rows = [], [], [], []
    rk_reuse_rows = []

    reuse_new_indices = set()
    reuse_old_indices = set()
    rk_new_indices = set()
    rk_old_indices = set()
    swap_reuse_new_indices = set()
    swap_reuse_old_indices = set()
    modify_new_indices = set()
    modify_old_indices = set()

    def get_old_meta_by_idx(idx):
        return old_meta_list[idx]

    new_records = new.to_dict('records')
    new_meta_list = []
    new_data_tuples = []
    for i, r in enumerate(new_records):
        spec = r.get(new_spec_col, "") if new_spec_col else ""
        flat = r.get(new_flat_col, "") if new_flat_col else r.get("FLAT", "") if "FLAT" in r else ""
        loc = r.get(new_loc_col, "") if new_loc_col else ""
        part = r.get(new_part_col, "") if new_part_col else ""
        qty = r.get(new_qty_col, "") if new_qty_col else ""
        unit = r.get(new_unit_col, "") if new_unit_col else ""
        wl_n = int(r.get(wl_col_new, 0))
        wr_n = int(r.get(wr_col_new, 0))
        ll_n = int(r.get(ll_col_new, 0))
        lr_n = int(r.get(lr_col_new, 0))
        code_n = str(r.get(code_col_new, "")) if pd.notna(r.get(code_col_new)) else ""
        base_n = str(r.get("BASE", ""))

        new_meta_list.append((
            (str(spec) if pd.notna(spec) else ""),
            wl_n, wr_n, code_n, ll_n, lr_n,
            (str(flat) if pd.notna(flat) else ""),
            (str(loc) if pd.notna(loc) else ""),
            (str(part) if pd.notna(part) else ""),
            (str(qty) if pd.notna(qty) else ""),
            (str(unit) if pd.notna(unit) else "")
        ))
        new_data_tuples.append((
            wl_n, wr_n, ll_n, lr_n, base_n, code_n
        ))

    def get_new_meta(i):
        return new_meta_list[i]

    # -------- STAGE 1: REUSE SWEEP (AS-IS -> ALT -> RK) -----------
    # Skip REUSE stage if operation_mode is 'MODIFY Only'
    if operation_mode != 'MODIFY Only':
        for i in range(len(new_data_tuples)):
            wlN, wrN, LLN, LRN, baseN, codeN = new_data_tuples[i]
            matched = False

            new_meta = get_new_meta(i)

            # --- STEP 1: AS-IS exact match (WL, WR, BASE, LL, LR) ---
            for idx_old in reuse_idx.get((wlN, wrN, baseN, LLN, LRN), []):
                if not used_old[idx_old]:
                    old_meta = get_old_meta_by_idx(idx_old)
                    reuse_rows.append(
                        list(new_meta) + list(old_meta) + [
                            area(wlN, wrN, LLN, LRN),
                            area(old_meta[1], old_meta[2], old_meta[4], old_meta[5]),
                            f"AS-IS ({baseN})",
                        ]
                    )
                    used_old[idx_old] = True
                    reuse_new_indices.add(i)
                    reuse_old_indices.add(idx_old)
                    inv_list[idx_old]["USED"] = True
                    matched = True
                    break
            if matched:
                continue

            # --- STEP 2: ALT match (same WL,WR,LL,LR, different base via REUSE_ALT_RULES) ---
            if baseN in REUSE_ALT_RULES:
                for alt_base, remark in REUSE_ALT_RULES[baseN]:
                    for idx_old in reuse_idx.get((wlN, wrN, alt_base, LLN, LRN), []):
                        if not used_old[idx_old]:
                            old_meta = get_old_meta_by_idx(idx_old)
                            display_remark = f"{remark} ({baseN}\u2194{alt_base})" if remark != "ALT" else f"ALT ({baseN}\u2194{alt_base})"
                            reuse_rows.append(
                                list(new_meta) + list(old_meta) + [
                                    area(wlN, wrN, LLN, LRN),
                                    area(old_meta[1], old_meta[2], old_meta[4], old_meta[5]),
                                    display_remark,
                                ]
                            )
                            used_old[idx_old] = True
                            reuse_new_indices.add(i)
                            reuse_old_indices.add(idx_old)
                            inv_list[idx_old]["USED"] = True
                            matched = True
                            break
                    if matched:
                        break
            if matched:
                continue

            # --- STEP 3: RK match (same WL,WR; LL adjusted by ±50 via REUSE_RK_RULES) ---
            if baseN in REUSE_RK_RULES:
                for (old_base, diff, remark) in REUSE_RK_RULES[baseN]:
                    for idx_old in mod_idx.get((wlN, wrN, old_base), []):
                        if used_old[idx_old]:
                            continue
                        LLO = inv_list[idx_old]["LL"]
                        if LLO == LLN + diff:
                            old_meta = get_old_meta_by_idx(idx_old)
                            rk_reuse_rows.append(
                                list(new_meta) + list(old_meta) + [
                                    area(wlN, wrN, LLN, LRN),
                                    area(old_meta[1], old_meta[2], old_meta[4], old_meta[5]),
                                    f"RK ({remark}) {baseN}\u2194{old_base}",
                                ]
                            )
                            used_old[idx_old] = True
                            rk_new_indices.add(i)
                            rk_old_indices.add(idx_old)
                            inv_list[idx_old]["USED"] = True
                            matched = True
                            break
                    if matched:
                        break

            # --- STEP 4: SWAP REUSE (rotate/flip panel and reuse AS-IS, no cutting) ---
            if not matched:
                if baseN in SWAP_REUSE_ROTATE_BASES:
                    for idx_old in swap_rotate_idx.get((wlN, wrN, LLN, LRN, baseN), []):
                        if not used_old[idx_old]:
                            old_meta = get_old_meta_by_idx(idx_old)
                            reuse_rows.append(
                                list(new_meta) + list(old_meta) + [
                                    area(wlN, wrN, LLN, LRN),
                                    area(old_meta[1], old_meta[2], old_meta[4], old_meta[5]),
                                    f"SWAP REUSE (WR\u2194LL) ({baseN})",
                                ]
                            )
                            inv_list[idx_old]["USED"] = True
                            used_old[idx_old] = True
                            swap_reuse_new_indices.add(i)
                            swap_reuse_old_indices.add(idx_old)
                            reuse_new_indices.add(i)
                            reuse_old_indices.add(idx_old)
                            matched = True
                            break

                if not matched and baseN in SWAP_REUSE_FLIP_BASES:
                    for idx_old in swap_flip_idx.get((wlN, wrN, LLN, LRN, baseN), []):
                        if not used_old[idx_old]:
                            old_meta = get_old_meta_by_idx(idx_old)
                            reuse_rows.append(
                                list(new_meta) + list(old_meta) + [
                                    area(wlN, wrN, LLN, LRN),
                                    area(old_meta[1], old_meta[2], old_meta[4], old_meta[5]),
                                    f"SWAP REUSE (WL\u2194WR) ({baseN})",
                                ]
                            )
                            inv_list[idx_old]["USED"] = True
                            used_old[idx_old] = True
                            swap_reuse_new_indices.add(i)
                            swap_reuse_old_indices.add(idx_old)
                            reuse_new_indices.add(i)
                            reuse_old_indices.add(idx_old)
                            matched = True
                            break

            if (i + 1) % 500 == 0 or (i + 1) == len(new):
                percent = 10 + int(((i + 1) / len(new)) * 70)
                update_progress(percent, f"Processing REUSE {i + 1}/{len(new)}")

    # If we are running in REUSE Only mode, capture all remaining NEW rows
    # (i.e. not matched in REUSE / RK) into the Balance (Not Reused/Modified)
    # bucket so that SUMMARY and the "Not Reuse & Modify" sheet reflect the
    # full New sheet area.
    if operation_mode == 'REUSE Only':
        for i in range(len(new_data_tuples)):
            if i in reuse_new_indices or i in rk_new_indices or i in swap_reuse_new_indices:
                continue
            new_meta = get_new_meta(i)
            none_rows.append(
                list(new_meta) + [
                    area(new_meta[1], new_meta[2], new_meta[4], new_meta[5])
                ]
            )

    if operation_mode == 'MODIFY Only':
        for inv in inv_list:
            inv["USED"] = False
        used_old = [False] * len(old)
    scrap_pool = []
    scrap_by_base = defaultdict(list)
    scrap_by_code = defaultdict(list)

    def add_to_scrap_pool(entry):
        scrap_pool.append(entry)
        s_b = entry.get("BASE", "") or base_code(entry.get("CODE", ""))
        s_c = entry.get("CODE", "")
        scrap_by_base[s_b].append(entry)
        scrap_by_code[s_c].append(entry)

    # Precompute remaining size counts for ultra-fast O(1) remainder check
    future_size_counts = defaultdict(lambda: defaultdict(int))
    for j in range(len(new_data_tuples)):
        if j in reuse_new_indices or j in rk_new_indices or j in swap_reuse_new_indices:
            continue
        wl_n, wr_n, ll_n, lr_n, _, c = new_data_tuples[j]
        future_size_counts[c][(wl_n, wr_n, ll_n)] += 1

    def can_consume_remainder(rem_code, rem_wl, rem_wr, rem_ll, _start_idx=0):
        """Check if any remaining requirement can use this remainder piece in O(1) time."""
        sizes = future_size_counts.get(rem_code)
        if not sizes:
            return False
        for (u_wl, u_wr, u_ll), count in sizes.items():
            if count > 0 and rem_wl >= u_wl and rem_wr >= u_wr and rem_ll >= u_ll:
                return True
        return False

    def check_constraint(constraint, inv_wl, inv_wr, inv_ll, inv_lr, wlN, wrN, LLN, LRN, tol, eff_cut_mode):
        """Check if an inventory panel satisfies a constraint against required dims."""
        if constraint == "FULL":
            if eff_cut_mode == 'Length':
                if inv_wl == wlN and inv_wr == wrN and inv_ll >= LLN + tol and inv_lr >= LRN:
                    return True
                return False
            elif eff_cut_mode == 'Width':
                if inv_ll == LLN and inv_lr == LRN and inv_wl >= wlN and inv_wr >= wrN + tol:
                    return True
                return False
            else:
                if inv_wl >= wlN and inv_wr >= wrN and inv_ll >= LLN and inv_lr >= LRN:
                    # Must exceed by at least tolerance in at least one dimension
                    if (inv_wr >= wrN + tol) or (inv_ll >= LLN + tol):
                        return True
                return False
        elif constraint == "WL_WR_EXACT_LL":
            return inv_wl == wlN and inv_wr == wrN and inv_ll >= LLN + tol
        elif constraint == "WL_WR_EXACT_LL_GE":
            # Cross-family match (e.g. SC → ICK): WL+WR exact, LL just >= req (no +tol gap).
            # SC panels don't appear in REUSE so exact-LL is fine here.
            return inv_wl == wlN and inv_wr == wrN and inv_ll >= LLN
        elif constraint == "FULL+50LL":
            if eff_cut_mode == 'Length':
                if inv_wl == wlN and inv_wr == wrN and inv_ll >= LLN + 50 + tol and inv_lr >= LRN:
                    return True
                return False
            elif eff_cut_mode == 'Width':
                return False
            else:
                return (inv_wl >= wlN and inv_wr >= wrN and
                        inv_ll >= LLN + 50 + tol and inv_lr >= LRN)
        return False

    def consume_inventory(inv_entry, req_idx, req_tuple, source_label, modify_list_idx, matched_remark="ALT"):
        """Cut a panel from inventory and calculate remainder."""
        inv_entry["USED"] = True
        used_old_idx = inv_entry.get("old_idx", -1)
        if 0 <= used_old_idx < len(used_old):
            used_old[used_old_idx] = True

        inv_wl = inv_entry["WL"]
        inv_wr = inv_entry["WR"]
        inv_ll = inv_entry["LL"]
        inv_lr = inv_entry["LR"]
        req_wl, req_wr, req_ll, req_lr, req_base, req_code = req_tuple

        scrap_w = 0
        scrap_l = 0

        # Determine if this is a WL/WR-exact code (only cut LL)
        is_ll_only = (req_base in LL_ONLY_CUT_BASES or
                      req_code.startswith("SC") or req_code.startswith("SI") or req_code.startswith("SO"))

        if is_ll_only:
            # WL/WR stay exact, only cut LL
            if inv_ll > req_ll:
                scrap_l = inv_ll - req_ll - tolerance
        else:
            # Check cut mode (only for W, WS)
            eff_cut_mode = cut_mode if req_base in CUT_MODE_BASES else 'Both'

            if eff_cut_mode == 'Length':
                # WL same, cut only LL
                if inv_ll > req_ll:
                    scrap_l = inv_ll - req_ll - tolerance
            elif eff_cut_mode == 'Width':
                # LL same, cut only WL (actually WR dimension)
                if inv_wr > req_wr:
                    scrap_w = inv_wr - req_wr - tolerance
            else:  # Both
                if inv_ll > req_ll:
                    scrap_l = inv_ll - req_ll - tolerance
                if inv_wr > req_wr:
                    scrap_w = inv_wr - req_wr - tolerance

        # Add remainder to scrap pool if it's usable
        if max(scrap_w, 0) >= MIN_USABLE_REMAINDER_MM or max(scrap_l, 0) >= MIN_USABLE_REMAINDER_MM:
            add_to_scrap_pool({
                'WL': inv_wl,
                'WR': int(scrap_w) if scrap_w > 0 else inv_wr,
                'CODE': inv_entry["CODE"],
                'LL': int(scrap_l) if scrap_l > 0 else inv_ll,
                'LR': inv_lr,
                'Balance': 1,
                'Origin_Row': req_idx + 1,
                'BASE': inv_entry.get("BASE", ""),
                'Modify_List_Idx': modify_list_idx
            })

        is_scrap = (source_label == 'SCRAP')
        rem_wl = inv_wl
        rem_wr = int(scrap_w) if scrap_w > 0 else inv_wr
        rem_ll = int(scrap_l) if scrap_l > 0 else inv_ll
        width_str = f"{rem_wl}+{rem_wr}" if rem_wl > 0 else f"{rem_wr}"
        rem_str = f"{width_str} {req_code} {rem_ll}"

        if is_scrap and (scrap_w > 0 or scrap_l > 0):
            gen2_wr = int(scrap_w) if scrap_w > 0 else inv_wr
            gen2_ll = int(scrap_l) if scrap_l > 0 else inv_ll
            if max(gen2_wr - req_wr, 0) >= MIN_USABLE_REMAINDER_MM or max(gen2_ll - req_ll, 0) >= MIN_USABLE_REMAINDER_MM:
                add_to_scrap_pool({
                    'WL': inv_wl,
                    'WR': gen2_wr,
                    'CODE': inv_entry["CODE"],
                    'LL': gen2_ll,
                    'LR': inv_lr,
                    'Balance': 1,
                    'Origin_Row': req_idx + 1,
                    'BASE': inv_entry.get("BASE", ""),
                    'Modify_List_Idx': modify_list_idx,
                    '_gen': 2
                })
        details = {
            'Status': 'MODIFY',
            'Source': 'FROM REMAINDER' if is_scrap else 'Inventory',
            'Source_Detail': f"{inv_wr} {req_code} {inv_ll}" if is_scrap else '',
            'Inv_WL': inv_wl,
            'Inv_WR': inv_wr,
            'Inv_CODE': inv_entry['CODE'],
            'Inv_LL': inv_ll,
            'Inv_LR': inv_lr,
            'Inv_QTY_USED': 1,
            'Inv_old_idx': inv_entry.get('old_idx'),
            'Remainder': rem_str if (scrap_w > 0 or scrap_l > 0) else None,
            'Remainder_Used': '',
            'Modify_Remark': matched_remark if not is_scrap else 'FROM REMAINDER'
        }
        return details

    if operation_mode != 'REUSE Only':
        for i in range(len(new_data_tuples)):
            if i in reuse_new_indices or i in rk_new_indices or i in swap_reuse_new_indices:
                continue

            req_tuple = new_data_tuples[i]
            wlN, wrN, LLN, LRN, baseN, codeN = req_tuple
            new_meta = get_new_meta(i)

            # Decrement future requirements count for this exact profile
            if future_size_counts[codeN][(wlN, wrN, LLN)] > 0:
                future_size_counts[codeN][(wlN, wrN, LLN)] -= 1

            # Resolve modify rules for this base code
            modify_rules = resolve_modify_rules(baseN, codeN)

            # NO_MODIFY: DP, Jointbar — skip to none_rows
            if modify_rules == "NO_MODIFY":
                none_rows.append(
                    list(new_meta) + [
                        area(new_meta[1], new_meta[2], new_meta[4], new_meta[5])
                    ]
                )
                if (i + 1) % 500 == 0 or (i + 1) == len(new_data_tuples):
                    percent = 80 + int(((i + 1) / len(new_data_tuples)) * 15)
                    update_progress(percent, f"Processing MODIFY {i + 1}/{len(new_data_tuples)}")
                continue

            # No rules found — add to none_rows
            if modify_rules is None:
                none_rows.append(
                    list(new_meta) + [
                        area(new_meta[1], new_meta[2], new_meta[4], new_meta[5])
                    ]
                )
                if (i + 1) % 500 == 0 or (i + 1) == len(new_data_tuples):
                    percent = 80 + int(((i + 1) / len(new_data_tuples)) * 15)
                    update_progress(percent, f"Processing MODIFY {i + 1}/{len(new_data_tuples)}")
                continue

            details = None
            matched_remark_final = "ALT"

            # --- STEP A: Try scrap pool first (remainder from previous cuts) ---
            # SI* and SO* codes CANNOT be matched from remainder — skip scrap pool for them
            best_scrap = None
            best_scrap_score = -1

            is_max_rem = (baseN in MAX_REMAINDER_BASES or
                          codeN.startswith("SC") or codeN.startswith("SI") or codeN.startswith("SO"))

            # Check cut mode (only for W, WS)
            eff_cut_mode = cut_mode if baseN in CUT_MODE_BASES else 'Both'

            if not (codeN.startswith("SI") or codeN.startswith("SO") or codeN.startswith("KC") or codeN.startswith("KCE")):
                scrap_candidates = []
                seen_s_id = set()
                for (rule_code, remark, constraint) in modify_rules:
                    # Clean scrap_by_base
                    s_base_list = scrap_by_base.get(rule_code, [])
                    if s_base_list:
                        active_s_base = []
                        for s in s_base_list:
                            if s.get('Balance', 0) > 0:
                                active_s_base.append(s)
                                if id(s) not in seen_s_id:
                                    seen_s_id.add(id(s))
                                    scrap_candidates.append(s)
                        scrap_by_base[rule_code] = active_s_base

                    # Clean scrap_by_code
                    s_code_list = scrap_by_code.get(rule_code, [])
                    if s_code_list:
                        active_s_code = []
                        for s in s_code_list:
                            if s.get('Balance', 0) > 0:
                                active_s_code.append(s)
                                if id(s) not in seen_s_id:
                                    seen_s_id.add(id(s))
                                    scrap_candidates.append(s)
                        scrap_by_code[rule_code] = active_s_code

                for s in scrap_candidates:
                    scrap_base = s.get('BASE', base_code(s.get('CODE', '')))
                    scrap_code = s.get('CODE', '')
                    scrap_ok = False
                    for (rule_code, remark, constraint) in modify_rules:
                        if rule_code == scrap_base or rule_code == scrap_code:
                            if check_constraint(constraint,
                                                s['WL'], s['WR'], s['LL'], s.get('LR', 0),
                                                wlN, wrN, LLN, LRN, tolerance, eff_cut_mode):
                                scrap_ok = True
                                break
                    if not scrap_ok:
                        continue

                    # Score: best-fit
                    sw = max(s['WR'] - wrN, 0)
                    sl = max(s['LL'] - LLN, 0)
                    waste = sw + sl
                    score = waste if is_max_rem else (1000000 - waste)
                    if score > best_scrap_score:
                        best_scrap_score = score
                        best_scrap = s

            if best_scrap is not None:
                s = best_scrap
                s['Balance'] = 0
                details = consume_inventory(s, i, req_tuple, source_label='SCRAP', modify_list_idx=len(modify_rows), matched_remark="FROM REMAINDER")
                modify_new_indices.add(i)

                origin_modify_idx = s.get('Modify_List_Idx')
                if origin_modify_idx is not None and origin_modify_idx < len(modify_rows):
                    current_sr_no = len(modify_rows) + 2  # Excel row number (header is row 1, data starts at 2)
                    existing_val = modify_rows[origin_modify_idx][21]
                    new_val = f"{current_sr_no}"
                    if existing_val:
                        modify_rows[origin_modify_idx][21] = existing_val + ", " + new_val
                    else:
                        modify_rows[origin_modify_idx][21] = new_val

            # --- Gather filtered candidate inventory matching rule codes ---
            inv_candidates = []
            seen_cand_idx = set()
            for (rule_code, remark, constraint) in modify_rules:
                # Clean inv_by_base
                base_list = inv_by_base.get(rule_code, [])
                if base_list:
                    active_base = []
                    for inv in base_list:
                        if not inv["USED"]:
                            active_base.append(inv)
                            if inv["old_idx"] not in seen_cand_idx:
                                seen_cand_idx.add(inv["old_idx"])
                                inv_candidates.append(inv)
                    inv_by_base[rule_code] = active_base

                # Clean inv_by_code
                code_list = inv_by_code.get(rule_code, [])
                if code_list:
                    active_code = []
                    for inv in code_list:
                        if not inv["USED"]:
                            active_code.append(inv)
                            if inv["old_idx"] not in seen_cand_idx:
                                seen_cand_idx.add(inv["old_idx"])
                                inv_candidates.append(inv)
                    inv_by_code[rule_code] = active_code

            # --- STEP B: Try inventory ---
            if details is None:
                candidate = None
                best_score = -1
                candidate_remark = "ALT"

                for inv in inv_candidates:
                    if inv["USED"]:
                        continue

                    inv_wl, inv_wr, inv_ll, inv_lr = inv["WL"], inv["WR"], inv["LL"], inv["LR"]
                    inv_code = inv["CODE"]
                    inv_base = inv["BASE"]

                    # Check each rule in order
                    matched = False
                    matched_remark_candidate = "ALT"
                    for (rule_code, remark, constraint) in modify_rules:
                        if rule_code == inv_base or rule_code == inv_code:
                            if check_constraint(constraint,
                                                inv_wl, inv_wr, inv_ll, inv_lr,
                                                wlN, wrN, LLN, LRN, tolerance, eff_cut_mode):
                                matched = True
                                matched_remark_candidate = remark
                                break

                    if not matched:
                        continue

                    # Calculate waste for scoring
                    sw = 0
                    sl = 0
                    is_ll_only = (baseN in LL_ONLY_CUT_BASES or
                                  codeN.startswith("SC") or codeN.startswith("SI") or codeN.startswith("SO"))

                    if is_ll_only:
                        if inv_ll > LLN:
                            sl = inv_ll - LLN - tolerance
                    else:
                        if inv_ll > LLN:
                            sl = inv_ll - LLN - tolerance
                        if inv_wr > wrN:
                            sw = inv_wr - wrN - tolerance

                    waste = max(sw, 0) + max(sl, 0)

                    if is_max_rem:
                        score = 3000000 + waste  # Prefer larger panels (max remainder)
                    elif matched_remark_candidate == "USE SC":
                        score = 1500000 - waste
                    else:
                        if max(sl, 0) >= MIN_USABLE_REMAINDER_MM or max(sw, 0) >= MIN_USABLE_REMAINDER_MM:
                            rem_wr_val = max(sw, 0) if sw > 0 else inv_wr
                            rem_ll_val = max(sl, 0) if sl > 0 else inv_ll
                            if can_consume_remainder(inv_code, inv_wl, rem_wr_val, rem_ll_val, i):
                                score = 2500000 - waste
                            else:
                                continue  # REJECT: large remainder not consumable by anyone
                        else:
                            score = 1000000 - waste  # Best-fit: smallest waste

                    if score > best_score:
                        best_score = score
                        candidate = inv
                        candidate_remark = matched_remark_candidate

                if candidate is not None:
                    details = consume_inventory(candidate, i, req_tuple, source_label='INVENTORY', modify_list_idx=len(modify_rows), matched_remark=candidate_remark)
                    modify_new_indices.add(i)
                    modify_old_indices.add(candidate["old_idx"])

            # --- STEP C: Code-family-aware Swap MODIFY ---
            if details is None and swap_modify:
                swap_candidate = None
                swap_best_score = -1
                swap_remark = "ALT"
                swapped_dims = {}   # will hold effective dims of winner

                for inv in inv_candidates:
                    if inv["USED"]:
                        continue

                    inv_base = inv["BASE"]
                    inv_code = inv["CODE"]

                    # --- Determine effective dims after swap based on required code family ---
                    if baseN in SWAP_REUSE_FLIP_BASES:
                        eff_wl = inv["WR"]
                        eff_wr = inv["WL"]
                        eff_ll = inv["LL"]
                        eff_lr = inv["LR"]
                        swap_label = "WL↔WR"
                    elif baseN in SWAP_REUSE_ROTATE_BASES:
                        eff_wl = inv["WL"]
                        eff_wr = inv["LL"]
                        eff_ll = inv["WR"]
                        eff_lr = inv["LR"]
                        swap_label = "WR↔LL"
                    else:
                        eff_wl = inv["LR"]
                        eff_wr = inv["LL"]
                        eff_ll = inv["WR"]
                        eff_lr = inv["WL"]
                        swap_label = "SWAP"

                    # Check modify rules with effective (swapped) dims
                    matched_swap = False
                    matched_swap_remark = "ALT"
                    for (rule_code, remark, constraint) in modify_rules:
                        if rule_code == inv_base or rule_code == inv_code:
                            if check_constraint(constraint,
                                                eff_wl, eff_wr, eff_ll, eff_lr,
                                                wlN, wrN, LLN, LRN, tolerance, eff_cut_mode):
                                matched_swap = True
                                matched_swap_remark = remark
                                break

                    if not matched_swap:
                        continue

                    # Score based on waste after swap
                    sw = 0
                    sl = 0
                    is_ll_only_swap = (baseN in LL_ONLY_CUT_BASES or
                                       codeN.startswith("SC") or codeN.startswith("SI") or codeN.startswith("SO"))
                    if is_ll_only_swap:
                        if eff_ll > LLN:
                            sl = eff_ll - LLN - tolerance
                    else:
                        if eff_ll > LLN:
                            sl = eff_ll - LLN - tolerance
                        if eff_wr > wrN:
                            sw = eff_wr - wrN - tolerance

                    waste = max(sw, 0) + max(sl, 0)

                    if is_max_rem:
                        score = 3000000 + waste
                    else:
                        if max(sl, 0) >= MIN_USABLE_REMAINDER_MM or max(sw, 0) >= MIN_USABLE_REMAINDER_MM:
                            if can_consume_remainder(inv_code, eff_wl, max(sw, 0) if sw > 0 else eff_wr, max(sl, 0) if sl > 0 else eff_ll, i):
                                score = 2000000 - (eff_wl * eff_wr + eff_ll * eff_lr)
                            else:
                                continue
                        else:
                            score = 1000000 - waste

                    if score > swap_best_score:
                        swap_best_score = score
                        swap_candidate = inv
                        swap_remark = matched_swap_remark
                        swapped_dims = {"WL": eff_wl, "WR": eff_wr, "LL": eff_ll, "LR": eff_lr,
                                        "CODE": inv_code, "BASE": inv_base,
                                        "old_idx": inv["old_idx"], "label": swap_label}

                if swap_candidate is not None:
                    swapped_inv = dict(swap_candidate)
                    swapped_inv["WL"] = swapped_dims["WL"]
                    swapped_inv["WR"] = swapped_dims["WR"]
                    swapped_inv["LL"] = swapped_dims["LL"]
                    swapped_inv["LR"] = swapped_dims["LR"]
                    swap_candidate["USED"] = True
                    used_old_idx = swap_candidate.get("old_idx", -1)
                    if 0 <= used_old_idx < len(used_old):
                        used_old[used_old_idx] = True

                    details = consume_inventory(
                        swapped_inv, i, req_tuple,
                        source_label='INVENTORY',
                        modify_list_idx=len(modify_rows),
                        matched_remark=f"{swap_remark} (SWAP {swapped_dims['label']})"
                    )
                    modify_new_indices.add(i)
                    modify_old_indices.add(swap_candidate["old_idx"])

            # --- Build output row ---
            if details is not None:
                if details.get('Source') == 'FROM REMAINDER':
                    old_meta = ("", "", "", "", "", "", "", "", "")
                elif details.get('Inv_old_idx') is not None:
                    old_meta = get_old_meta_by_idx(details['Inv_old_idx'])
                elif details.get('Inv_WL') is not None and details.get('Inv_WR') is not None:
                    old_meta = ("", details.get('Inv_WL'), details.get('Inv_WR'), details.get('Inv_CODE'), details.get('Inv_LL'), details.get('Inv_LR'), "", "", "")
                else:
                    old_meta = ("", 0, 0, "", 0, 0, "", "", "")

                inv_w_val = int(details.get('Inv_WR')) if details.get('Inv_WR') is not None else 0
                inv_l_val = int(details.get('Inv_LL')) if details.get('Inv_LL') is not None else 0
                area_new_val = area(new_meta[1], new_meta[2], new_meta[4], new_meta[5])
                
                if details.get('Source') == 'FROM REMAINDER':
                    area_old_val = area(details.get('Inv_WL'), inv_w_val, inv_l_val, details.get('Inv_LR'))
                else:
                    area_old_val = area(old_meta[1] or 0, inv_w_val or 0, inv_l_val or 0, old_meta[5] or 0)
                    
                # Build a rich remark showing exactly which panel / remainder was used
                mod_remark_tag = details.get('Modify_Remark', 'ALT')
                inv_wl_d = details.get('Inv_WL', 0)
                inv_wr_d = details.get('Inv_WR', 0)
                inv_ll_d = details.get('Inv_LL', 0)
                inv_code_d = details.get('Inv_CODE', '')
                panel_dim_str = f"{inv_wl_d}+{inv_wr_d} {inv_code_d} {inv_ll_d}" if inv_wl_d else f"{inv_wr_d} {inv_code_d} {inv_ll_d}"

                if details.get('Source') == 'FROM REMAINDER':
                    src_detail = details.get('Source_Detail', '') or panel_dim_str
                    remark_str = f"MODIFY - Remainder Used: {src_detail}"
                elif '(SWAP)' in mod_remark_tag:
                    remark_str = f"MODIFY - Panel Used (SWAP): {panel_dim_str}"
                    if mod_remark_tag.replace('(SWAP)', '').strip() not in ('', 'ALT'):
                        remark_str += f" | {mod_remark_tag.replace('(SWAP)', '').strip()}"
                else:
                    remark_str = f"MODIFY - Panel Used: {panel_dim_str}"
                    if mod_remark_tag not in ('', 'ALT'):
                        remark_str += f" | {mod_remark_tag}"

                # Always fill Source_Detail with original panel dims
                if not details.get('Source_Detail'):
                    details['Source_Detail'] = panel_dim_str

                modify_rows.append(
                    list(new_meta) + list(old_meta) + [
                        details.get('Remainder'),
                        details.get('Remainder_Used', ''),
                        details.get('Source'),
                        details.get('Source_Detail', ''),
                        area_new_val,
                        area_old_val,
                        remark_str,
                    ]
                )
                if (i + 1) % 500 == 0 or (i + 1) == len(new):
                    percent = 80 + int(((i + 1) / len(new)) * 15)
                    update_progress(percent, f"Processing MODIFY {i + 1}/{len(new)}")
            else:
                none_rows.append(
                    list(new_meta) + [
                        area(new_meta[1], new_meta[2], new_meta[4], new_meta[5])
                    ]
                )

    for idx, ro in enumerate(old_records):
        if not used_old[idx]:
            wlO = inv_list[idx]["WL"]
            wrO = inv_list[idx]["WR"]
            LLO = inv_list[idx]["LL"]
            LRO = inv_list[idx]["LR"]
            codeO = inv_list[idx]["CODE"]
            unused_rows.append([wlO, wrO, codeO, LLO, LRO, area(wlO, wrO, LLO, LRO)])

    new_headers = ["New Spec", "New WL", "New WR", "New CODE", "New LL", "New LR", "New Flat", "New Location", "New Part", "New QTY", "New Unit"]
    old_headers = ["Old Spec", "Old WL", "Old WR", "Old CODE", "Old LL", "Old LR", "Old Location", "Old QTY", "Old Unit Area"]
    tail_headers_reuse = ["Area New (m^2)", "Area Old (m^2)", "Remark"]

    df_reuse = pd.DataFrame(reuse_rows, columns=new_headers + old_headers + tail_headers_reuse) if reuse_rows else pd.DataFrame(columns=new_headers + old_headers + tail_headers_reuse)
    df_rk = pd.DataFrame(rk_reuse_rows, columns=new_headers + old_headers + tail_headers_reuse) if rk_reuse_rows else pd.DataFrame(columns=new_headers + old_headers + tail_headers_reuse)

    df_modify = pd.DataFrame(modify_rows, columns=
                        new_headers + old_headers +
                        ["Remainder", "Remainder_Used", "Source", "Source_Detail",
                        "Area New", "Area Old", "Remark"]) if modify_rows else pd.DataFrame(columns=
                        new_headers + old_headers +
                        ["Remainder", "Remainder_Used", "Source", "Source_Detail",
                        "Area New", "Area Old", "Remark"])

    df_unused = pd.DataFrame(unused_rows, columns=["WL","WR","CODE","LL","LR","Area m^2"])
    df_none = pd.DataFrame(none_rows, columns=["New Spec","New WL","New WR","New CODE","New LL","New LR","New Flat","New Location","New Part","New QTY","New Unit","Area m^2"])
    df_scrap = pd.DataFrame(scrap_pool)

    # Convert New QTY to numeric for proper summing
    df_none['New QTY'] = pd.to_numeric(df_none['New QTY'], errors='coerce').fillna(0).astype(int)

    # Ensure numeric types for area columns to avoid TypeErrors when summing
    if 'Area New' in df_modify.columns:
        df_modify['Area New'] = pd.to_numeric(df_modify['Area New'], errors='coerce').fillna(0.0)
    if 'Area Old' in df_modify.columns:
        df_modify['Area Old'] = pd.to_numeric(df_modify['Area Old'], errors='coerce').fillna(0.0)

    # --- NEW: Build MERGED sheet using MODIFY sheet layout as master ---
    # Master layout (as you confirmed): New headers + Old headers + remainder/source + area + remark
    merged_cols = new_headers + old_headers + ["Remainder", "Remainder_Used", "Status", "Source_Detail", "Area New (m^2)", "Area Old (m^2)", "Remark"]
    merged_rows = []

    def pick(df_row, col_name):
        if col_name in df_row:
            return df_row.get(col_name, "")
        return ""

    # Helper to convert reuse/rk rows (they already match new_headers + old_headers + tail headers)
    def convert_reuse_row(row):
        out = []
        for c in new_headers:
            out.append(pick(row, c))
        for c in old_headers:
            out.append(pick(row, c))
        # Remainder related columns: none for reuse
        out += ["", "", pick(row, "Remark"), ""]  # Remainder, Remainder_Used, Source (we'll use Remark as Source), Source_Detail placeholder
        # Area New, Area Old, Remark - reuse already has Area New (m^2), Area Old (m^2), Remark at the end
        an = pick(row, "Area New (m^2)")
        ao = pick(row, "Area Old (m^2)")
        rem = pick(row, "Remark")
        # Replace the earlier 'Remark' placed into Source; we want Source derived from Remark text and Remark as actual remark
        source = ""
        if isinstance(rem, str):
            if "AS-IS" in rem.upper():
                source = "AS-IS"
            elif "RK" in rem.upper():
                source = "RK REUSE"
            elif "ALT" in rem.upper() or "VV" in rem.upper():
                source = "ALT/AS-IS"
            else:
                source = rem
        else:
            source = rem
        out[-4] = ""  # Remainder
        out[-3] = ""  # Remainder_Used
        out[-2] = source  # Source
        out[-1] = ""  # Source_Detail
        out += [an, ao, rem]
        return out

    # Convert MODIFY rows using the same master layout (new + old + remainder/source + areas + remark)
    def convert_modify_row(row):
        out = []
        for c in new_headers:
            out.append(pick(row, c))
        for c in old_headers:
            out.append(pick(row, c))
        rem_val = pick(row, "Remainder")
        rem_used = pick(row, "Remainder_Used")
        src = pick(row, "Source")
        src_det = pick(row, "Source_Detail")
        an = pick(row, "Area New") or pick(row, "Area New (m^2)")
        ao = pick(row, "Area Old") or pick(row, "Area Old (m^2)")
        remark = pick(row, "Remark") or ""
        out += [rem_val, rem_used, src, src_det, an, ao, remark]
        return out

    # Convert NOT REUSE / NOT MODIFY rows: these have only New... and Area m^2. We will set Old... empty and Remark = NEW
    def convert_none_row(row):
        out = []
        # row is series with new headers plus 'Area m^2'
        for c in new_headers:
            out.append(pick(row, c))
        for c in old_headers:
            out.append("")  # empty old columns
        out += ["", "", "NEW", "", pick(row, "Area m^2"), "", "NEW"]
        return out

    # Append df_reuse rows
    if not df_reuse.empty:
        for r in df_reuse.to_dict('records'):
            merged_rows.append(convert_reuse_row(r))
    # Append df_rk rows
    if not df_rk.empty:
        for r in df_rk.to_dict('records'):
            merged_rows.append(convert_reuse_row(r))
    # Append df_modify rows
    if not df_modify.empty:
        for r in df_modify.to_dict('records'):
            merged_rows.append(convert_modify_row(r))
    # Append df_none (Not Reuse & Modify) rows - mark as NEW
    if not df_none.empty:
        for r in df_none.to_dict('records'):
            merged_rows.append(convert_none_row(r))

    df_merged = pd.DataFrame(merged_rows, columns=merged_cols)

    total_new_count = len(new)
    total_new_area = 0.0
    total_new_area += df_reuse['Area New (m^2)'].sum() if not df_reuse.empty and 'Area New (m^2)' in df_reuse.columns else 0.0
    total_new_area += df_rk['Area New (m^2)'].sum() if not df_rk.empty and 'Area New (m^2)' in df_rk.columns else 0.0
    if 'Area New' in df_modify.columns:
        total_new_area += df_modify['Area New'].sum()
    total_new_area += df_none['Area m^2'].sum() if not df_none.empty else 0.0

    total_old_count = len(old)
    total_old_area = 0.0
    total_old_area += df_reuse['Area Old (m^2)'].sum() if not df_reuse.empty and 'Area Old (m^2)' in df_reuse.columns else 0.0
    total_old_area += df_rk['Area Old (m^2)'].sum() if not df_rk.empty and 'Area Old (m^2)' in df_rk.columns else 0.0
    if 'Area Old' in df_modify.columns:
        total_old_area += df_modify['Area Old'].sum()
    total_old_area += df_unused['Area m^2'].sum() if not df_unused.empty else 0.0

    # Calculate combined reuse values
    reuse_count = len(df_reuse) + len(df_rk)
    reuse_area_new = (df_reuse['Area New (m^2)'].sum() if not df_reuse.empty and 'Area New (m^2)' in df_reuse.columns else 0.0) + (df_rk['Area New (m^2)'].sum() if not df_rk.empty and 'Area New (m^2)' in df_rk.columns else 0.0)
    reuse_area_old = (df_reuse['Area Old (m^2)'].sum() if not df_reuse.empty and 'Area Old (m^2)' in df_reuse.columns else 0.0) + (df_rk['Area Old (m^2)'].sum() if not df_rk.empty and 'Area Old (m^2)' in df_rk.columns else 0.0)

    df_summary = pd.DataFrame([
        ["Total Material List Panels", total_new_count, round(total_new_area,2), None],
        ["  Reused", reuse_count, round(reuse_area_new,2), round(reuse_area_old,2)],
        ["  Modified", len(df_modify), round(df_modify['Area New'].sum(),2) if 'Area New' in df_modify.columns else 0.0, round(df_modify['Area Old'].sum(),2) if 'Area Old' in df_modify.columns else 0.0],
        ["  New", len(df_none), round(df_none['Area m^2'].sum(),2) if not df_none.empty else 0.0, None],
        ["", "", "", ""],
        ["Total INV Panels", total_old_count, None, round(total_old_area,2)],
        ["  Used for Reuse", reuse_count, round(reuse_area_new,2), round(reuse_area_old,2)],
        ["  Used for Modify", len(df_modify), round(df_modify['Area New'].sum(),2) if 'Area New' in df_modify.columns else 0.0, round(df_modify['Area Old'].sum(),2) if 'Area Old' in df_modify.columns else 0.0],
        ["  Balance (Unused Stock)", len(df_unused), None, round(df_unused['Area m^2'].sum(),2) if not df_unused.empty else 0.0]
    ], columns=["Category","Quantity","Material List Area(m²)","INV Area(m²)"])

    # --- NEW: Build Flat/Location breakdown (alphabetical) ---
    try:
        # Combine all reuse (REUSE + RK REUSE) for AS-IT-IS area
        df_all_reuse = pd.concat([df_reuse, df_rk], ignore_index=True) if (not df_reuse.empty or not df_rk.empty) else pd.DataFrame(columns=new_headers + old_headers + tail_headers_reuse)

        # group by New Flat
        reuse_group = df_all_reuse.groupby("New Flat")["Area New (m^2)"].sum() if "New Flat" in df_all_reuse.columns else pd.Series(dtype=float)
        modify_group = df_modify.groupby("New Flat")["Area New"].sum() if "New Flat" in df_modify.columns else pd.Series(dtype=float)
        new_group = df_none.groupby("New Flat")["Area m^2"].sum() if "New Flat" in df_none.columns else pd.Series(dtype=float)

        flats = sorted(set(list(reuse_group.index) + list(modify_group.index) + list(new_group.index)))

        loc_rows = []
        for i, fl in enumerate(flats, start=1):
            asi_area = float(reuse_group.get(fl, 0.0)) if not reuse_group.empty else 0.0
            mod_area = float(modify_group.get(fl, 0.0)) if not modify_group.empty else 0.0
            new_area = float(new_group.get(fl, 0.0)) if not new_group.empty else 0.0
            subtotal = asi_area + mod_area + new_area
            loc_rows.append([i, fl, round(asi_area,2), round(mod_area,2), round(new_area,2), round(subtotal,2)])

        total_sub = sum([r[5] for r in loc_rows]) if loc_rows else 0.0
        # Append Total row
        loc_rows.append(["", "Total", "", "", "", round(total_sub,2)])
        # Append Percentage row (percent of total for each column)
        perc_asi = []
        perc_mod = []
        perc_new = []
        for r in loc_rows[:-1]:
            if total_sub > 0:
                perc_asi.append(round((r[2] / total_sub) * 100, 2))
                perc_mod.append(round((r[3] / total_sub) * 100, 2))
                perc_new.append(round((r[4] / total_sub) * 100, 2))
            else:
                perc_asi.append(0.0)
                perc_mod.append(0.0)
                perc_new.append(0.0)
        # build percentage row as a label plus percentages concatenated (for clarity)
        # We'll place percentage for each category in separate columns beneath Sub-Total row
        # For simplicity, create a single Percentage summary cell per category
        if total_sub > 0:
            total_perc_asi = round((sum([r[2] for r in loc_rows[:-1]]) / total_sub) * 100, 2)
            total_perc_mod = round((sum([r[3] for r in loc_rows[:-1]]) / total_sub) * 100, 2)
            total_perc_new = round((sum([r[4] for r in loc_rows[:-1]]) / total_sub) * 100, 2)
        else:
            total_perc_asi = total_perc_mod = total_perc_new = 0.0

        # prepare DataFrame
        df_location = pd.DataFrame(loc_rows, columns=["Sr. No.", "Location", "AS-IT-IS Area", "MODIFY Area", "NEW Area", "Sub-Total"])
        # Add a blank row then a percentage summary row for neatness
        perc_row = ["", "Percentage", f"{total_perc_asi}%", f"{total_perc_mod}%", f"{total_perc_new}%", "100%"]
        # We'll append the percentage row as the last row in df_location
        df_location = pd.concat([df_location, pd.DataFrame([perc_row], columns=df_location.columns)], ignore_index=True)
    except Exception as e:
        df_location = pd.DataFrame(columns=["Sr. No.", "Location", "AS-IT-IS Area", "MODIFY Area", "NEW Area", "Sub-Total"])
        print(f"(Location summary generation failed: {e})")

    # Write all sheets with INFO as first sheet
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        # Create INFO sheet first
        info_data = {"": ["Old = Inventory List", "New = Material List"]}
        df_info = pd.DataFrame(info_data)
        df_info.to_excel(xw, sheet_name="INFO", index=False)
        
        # MERGED next
        df_merged.to_excel(xw, sheet_name="MERGED", index=False)
        df_all_reuse.to_excel(xw, sheet_name="REUSED", index=False)
        df_modify.to_excel(xw, sheet_name="MODIFY", index=False)
        df_unused.to_excel(xw, sheet_name="UNUSED STOCK", index=False)
        # Write summary and location side-by-side in SUMMARY sheet
        df_summary.to_excel(xw, sheet_name="SUMMARY", index=False, startrow=0, startcol=0)
        # compute startcol for location block (leave 2-column gap)
        startcol = df_summary.shape[1] + 2
        try:
            # pandas supports startcol parameter for to_excel
            df_location.to_excel(xw, sheet_name="SUMMARY", index=False, startrow=0, startcol=startcol)
        except Exception:
            # fallback: write in a new sheet if writing to same sheet fails
            df_location.to_excel(xw, sheet_name="LOCATION SUMMARY", index=False)
        df_none.to_excel(xw, sheet_name="NEW", index=False)
        df_scrap.to_excel(xw, sheet_name="SCRAP_POOL", index=False)

    # --- Copy highlighted "New" and "Old" sheets into output file using captured indices ---
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill
        # NEW imports for formatting
        from openpyxl.styles import Border, Side, Alignment
        from openpyxl.utils.exceptions import InvalidFileException
        import zipfile, io, tempfile, shutil

        green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

        from io import BytesIO

        def safe_load_workbook(path, read_only=False):
            for attempt in range(3):
                try:
                    return load_workbook(path, data_only=True, read_only=read_only)
                except (InvalidFileException, KeyError, zipfile.BadZipFile, OSError) as e:
                    if attempt < 2:
                        time.sleep(1.5)
                    else:
                        raise e

        import gc
        gc.collect()

        wb_in = safe_load_workbook(in_path, read_only=True)
        wb_out = safe_load_workbook(out_path, read_only=False)
        
        # Format INFO sheet with big bold font
        ws_info = wb_out["INFO"]
        from openpyxl.styles import Font
        
        # Clear default columns and format cells
        ws_info.delete_cols(1)  # Remove the default empty column
        
        # Add text in cells A2 and A3
        ws_info.cell(2, 1, "Old = Inventory List")
        ws_info.cell(3, 1, "New = Material List")
        
        # Format with big bold font
        bold_font = Font(name='Arial', size=18, bold=True)
        ws_info.cell(2, 1).font = bold_font
        ws_info.cell(3, 1).font = bold_font
        
        # Adjust column width
        ws_info.column_dimensions['A'].width = 30
        
        # Add LEGENDS
        ws_info.cell(5, 1, "LEGENDS")
        ws_info.cell(6, 1, "REUSED").fill = green
        ws_info.cell(7, 1, "MODIFY").fill = yellow
        ws_info.cell(8, 1, "NEW")
        
        # Format legend with borders
        thin = Side(border_style='thin', color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for r in range(5, 9):
            ws_info.cell(r, 1).border = border
            ws_info.cell(r, 1).font = Font(name='Arial', size=11, bold=True)

        # Copy "New" sheet and apply highlights using the captured new-row indices
        if new_sheet in wb_in.sheetnames:
            ws_in = wb_in[new_sheet]
            out_name = "New (Highlighted)"
            if out_name in wb_out.sheetnames:
                std = wb_out[out_name]
                wb_out.remove(std)
            ws_out = wb_out.create_sheet(out_name)
            for r_idx, row in enumerate(ws_in.iter_rows(values_only=False), start=1):
                fill_color = None
                if r_idx > 1:
                    new_df_index = r_idx - 2
                    if new_df_index in reuse_new_indices or new_df_index in rk_new_indices:
                        fill_color = green
                    elif new_df_index in modify_new_indices:
                        fill_color = yellow

                for c_idx, cell in enumerate(row, start=1):
                    out_cell = ws_out.cell(r_idx, c_idx, cell.value)
                    if fill_color:
                        out_cell.fill = fill_color
            
            # Add legend at column R, row 3
            ws_out.cell(3, 18, "LEGENDS")
            ws_out.cell(4, 18, "REUSED").fill = green
            ws_out.cell(5, 18, "MODIFY").fill = yellow
            ws_out.cell(6, 18, "NEW")
            
            # Apply border to legend cells
            thin = Side(border_style='thin', color="000000")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            for r in range(3, 7):
                ws_out.cell(r, 18).border = border

        # Copy "Old" sheet and highlight rows used (used_old list aligns to old.reset_index())
        if old_sheet in wb_in.sheetnames:
            ws_in2 = wb_in[old_sheet]
            out_name2 = "Old (Highlighted)"
            if out_name2 in wb_out.sheetnames:
                std2 = wb_out[out_name2]
                wb_out.remove(std2)
            ws_out2 = wb_out.create_sheet(out_name2)
            for r_idx, row in enumerate(ws_in2.iter_rows(values_only=False), start=1):
                fill_color = None
                if r_idx > 1:
                    old_df_index = r_idx - 2
                    if old_df_index in reuse_old_indices or old_df_index in rk_old_indices:
                        fill_color = green
                    elif old_df_index in modify_old_indices:
                        fill_color = yellow

                for c_idx, cell in enumerate(row, start=1):
                    out_cell = ws_out2.cell(r_idx, c_idx, cell.value)
                    if fill_color:
                        out_cell.fill = fill_color
            
            # Add legend at column R, row 3
            ws_out2.cell(3, 18, "LEGENDS")
            ws_out2.cell(4, 18, "REUSED").fill = green
            ws_out2.cell(5, 18, "MODIFY").fill = yellow
            ws_out2.cell(6, 18, "NEW")
            
            # Apply border to legend cells
            thin = Side(border_style='thin', color="000000")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            for r in range(3, 7):
                ws_out2.cell(r, 18).border = border

        # --- NEW: Apply thin borders + center/middle alignment efficiently ---
        thin = Side(border_style='thin', color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        align = Alignment(horizontal='center', vertical='center', wrap_text=False)

        for ws in wb_out.worksheets:
            max_r = ws.max_row or 1
            last_data_col = 1
            for c in range(1, (ws.max_column or 1) + 1):
                if ws.cell(1, c).value is not None and str(ws.cell(1, c).value).strip() != "":
                    last_data_col = c
            
            # Format all rows for summary sheets; for massive data sheets, format up to 500 rows to keep it ultra fast
            style_limit = max_r if (ws.title in {"INFO", "SUMMARY", "LOCATION SUMMARY", "SCRAP_POOL"} or max_r <= 500) else 500
            for row in ws.iter_rows(min_row=1, max_row=style_limit, min_col=1, max_col=last_data_col):
                for cell in row:
                    cell.alignment = align
                    cell.border = border
            
            # Always ensure header row has borders and alignment
            for row in ws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=last_data_col):
                for cell in row:
                    cell.alignment = align
                    cell.border = border

        wb_out.save(out_path)
        wb_in.close()
        wb_out.close()

    except Exception as e:
        print(f"(Highlight copy skipped: {e})")

    update_progress(100, f"Done in {time.time()-t0:.1f}s -> {out_path}")
    print(f"Done in {time.time()-t0:.2f}s -> {out_path}")

def main():
    """
    Main entry point for command-line execution.
    
    Usage:
        python reuse_modification.py <input.xlsx> [output.xlsx] [tolerance_mm] [cut_mode]
    
    Args (command line):
        input.xlsx: Required - Input Excel file with "New" and "Old" sheets
        output.xlsx: Optional - Output file path (default: reuse_modify_output_location_summary.xlsx)
        tolerance_mm: Optional - Tolerance in mm for modify operations (default: 10)
        cut_mode: Optional - 'Length', 'Width', or 'Both' (default: 'Both')
    """
    if len(sys.argv) < 2:
        print("Usage: python reuse_modification.py <input.xlsx> [output.xlsx] [tolerance_mm] [cut_mode]")
        print("\nExample:")
        print("  python reuse_modification.py input.xlsx output.xlsx 10 Both")
        sys.exit(1)
    
    in_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else "reuse_modify_output_location_summary.xlsx"
    
    try:
        tol = int(sys.argv[3]) if len(sys.argv) > 3 else 10
    except ValueError:
        print(f"Warning: Invalid tolerance value '{sys.argv[3]}', using default 10mm")
        tol = 10
    
    cut_mode = sys.argv[4] if len(sys.argv) > 4 else 'Both'
    if cut_mode not in ['Length', 'Width', 'Both']:
        print(f"Warning: Invalid cut_mode '{cut_mode}', using default 'Both'")
        cut_mode = 'Both'
    
    operation_mode = sys.argv[5] if len(sys.argv) > 5 else 'Both'
    if operation_mode not in ['Both', 'REUSE Only', 'MODIFY Only']:
        print(f"Warning: Invalid operation_mode '{operation_mode}', using default 'Both'")
        operation_mode = 'Both'
    
    try:
        process_excel(in_path, out_path, tolerance=tol, cut_mode=cut_mode, operation_mode=operation_mode)
    except FileNotFoundError:
        print(f"Error: Input file not found: {in_path}")
        sys.exit(1)
    except Exception as e:
        print(f"Error processing file: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()