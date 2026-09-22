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

RK_MAX = 50  # mm (RK tolerance for length adjustments)

# --- Cutting / scrap policy constants -------------------------------------
# A remainder at or below this size is treated as "unusable scrap" and is
# always fine to leave behind (rule: "cut panels with maximum less than
# 200mm remainder").
MIN_USABLE_REMAINDER_MM = 200

# If the remainder exceeds MIN_USABLE_REMAINDER_MM on either axis, the cut
# is only allowed automatically when the wasted area is within this
# fraction of the source panel's own area (10-20% per the business rule);
# otherwise a later New row must be able to consume the remainder.
MAX_ACCEPTABLE_SCRAP_FRACTION = 0.20

# Base codes for which single-axis ('Length' only / 'Width' only) cutting
# is meaningful. Any other base code always cuts in 'Both' mode regardless
# of the user-selected cut_mode.
LENGTH_WIDTH_ONLY_BASES = {"W", "WS", "T", "B", "WE", "WSE", "TE", "D", "WXE"}


def clean_code_raw(s):
    if pd.isna(s):
        return ""
    t = str(s).strip().upper()
    for ch in ["-", "—", "_", ".", "-"]:
        t = t.replace(ch, "")
    if t in {"", "NA", "N/A", ".", "-"}:
        return ""
    if t.startswith("NSP"):
        t = t[3:]
    return t

def base_code(s):
    t = clean_code_raw(s)
    if t in {"CC", "CCL", "CCR"}:
        return "CC"
    if t in {"D", "SB"}:
        return "D"
    if t in {"DE", "SBE"}:
        return "DE"
    # EC family – any code starting with EC maps to base EC
    if t.startswith("EC") or t == "CA":
        return "EC"
    if t == "EB":
        return "EB"
    if t == "MB":
        return "MB"
    # IC sub-family – each variant keeps its own base for targeted modify matching
    if t == "ICB":
        return "ICB"
    if t == "ICK":
        return "ICK"
    if t == "ICT":
        return "ICT"
    if t == "BIC":
        return "BIC"
    if t in {"IC", "ICR", "NSP"}:
        return "IC"
    # K sub-family – own base codes for targeted modify matching
    if t == "K":
        return "K"
    if t == "KC":
        return "KC"
    if t == "KCE":
        return "KCE"
    if t == "KIC":
        return "KIC"
    if t in {"PC", "PCE"}:
        return "PC"
    if t in {"PH", "CP", "PPH", "CPP", "CPPP"}:
        return "PH"
    if t in {"SX", "SXC", "SXCE"}:
        return "SX"
    # SC / SI / SO families – return raw code so each code indexes uniquely in reuse_idx
    if t.startswith("SC"):
        return t
    if t.startswith("SI"):
        return t
    if t.startswith("SO"):
        return t
    if t in {"SL", "SLWL", "SLR"}:
        return "SL"
    if t == "LS":
        return "LS"
    if t == "T":
        return "T"
    if t == "TE":
        return "TE"
    if t == "B":
        return "B"
    if t in {"W", "WRB", "WRBS"}:
        return "W"
    if t in {"XBS", "LSC", "LSCY", "LSCZ"}:
        return "XBS"
    if t == "BHH":
        return "BHH"
    if t == "DP":
        return "DP"
    if t == "PLB":
        return "PLB"
    if t == "RK":
        return "RK"
    if t == "WE":
        return "WE"
    if t == "WSE":
        return "WSE"
    if t == "WS":
        return "WS"
    if t == "WX":
        return "WX"
    if t == "WXE":
        return "WXE"
    if t in {"HCC", "JOINTBAR", "SPD"}:
        return t
    return t

def is_alt_bridge(bn, bo):
    pairs = {
        ("WS", "WX"), ("WX", "WS"), ("WS", "T"), ("T", "WS"),
        ("WS", "B"), ("B", "WS"), ("WS", "D"), ("D", "WS"),
        ("WS", "PC"), ("PC", "WS"), ("WX", "T"), ("T", "WX"),
        ("WX", "B"), ("B", "WX"), ("WX", "D"), ("D", "WX"),
        ("WX", "PC"), ("PC", "WX"),
    }
    return (bn, bo) in pairs

MODIFICATION_RULES = {
    "WX": [("T", "ALT"), ("B", "ALT"), ("D", "ALT"), ("PC", "ALT"), ("WS", "ALT"), ("WX", "ALT")],
    "WE": [("WE", "ALT"), ("WSE", "ATTACH RK"), ("WS", "ATTACH EC"), ("W", "ATTACH EC"), ("T", "ATTACH EC"), ("B", "ATTACH EC"), ("D", "ATTACH EC"), ("PC", "ATTACH EC")],
    "WS": [("WS", "ALT"), ("W", "REMOVE RK"), ("WX", "ALT"), ("T", "ALT"), ("B", "ALT"), ("D", "NOTCHING"), ("PC", "NOTCHING")],
    "W": [("W", "ALT"), ("WS", "ATTACH RK"), ("WX", "ATTACH RK"), ("T", "ATTACH RK"), ("B", "ATTACH RK"), ("D", "ATTACH RK"), ("PC", "ATTACH RK")],
    "T": [("WX", "ALT"), ("B", "ALT"), ("WS", "ALT"), ("D", "ALT"), ("PC", "ALT"), ("T", "ALT"), ("TX", "ALT"), ("WT", "ALT")],
    "B": [("B", "ALT"), ("T", "ALT"), ("W", "ALT"), ("WX", "ALT"), ("WS", "ALT"), ("D", "ALT"), ("PC", "ALT"), ("TX", "ALT"), ("WT", "ALT")],
    "WXE": [("WXE", "ALT"), ("TE", "ALT"), ("WS", "ATTACH EC"), ("T", "ATTACH EC"), ("B", "ATTACH EC"), ("WE", "ATTACH EC"), ("WSE", "ATTACH EC"), ("W", "ATTACH EC & REMOVE RK")],
    "D": [("D", "ALT"), ("PC", "ALT"), ("DG", "ALT")],
    "DG": [("D", "ALT"), ("PC", "ALT"), ("DG", "ALT")],
    "DE": [("DE", "ALT"), ("D", "ATTACH EC"), ("PC", "ATTACH EC")],
    "PH": [("CP", "ALT"), ("PH", "ALT"), ("PPH", "ALT")],
    "PLB": [("PLB", "ALT")],
    
    # Existing standard fallbacks:
    "CC": [("SL", "ALT"), ("LS", "ALT")],
    "SC": [("XBS", "ALT")],
    "SCE": [("XBSE", "ALT")],
    "SX": [("JLX", "ALT"), ("JRX", "ALT")],
    "BHH": [("PCE", "ALT"), ("DE", "ALT"), ("D", "+RK"), ("PC", "+RK")],
    "LSS": [("CC", "ALT"), ("SL", "ALT"), ("LS", "ALT")],
    "SL": [("CC", "ALT"), ("LS", "ALT"), ("LSS", "ALT")],
    "LS": [("CC", "ALT"), ("SL", "ALT"), ("LSS", "ALT"), ("IC", "ALT")],
    "XBS": [("SC", "ALT")],
    "XBSE": [("SCE", "ALT")],
    "KIC": [("KC", "ALT")],
    "PC": [("D", "ALT")],
    "BIC": [("BIC", "ALT"), ("ICX", "ALT"), ("BC", "ALT")]
}

RK_MODIFICATION_RULES = {
    # Base: [(Old_Base, Diff, Remark)]
    "WS": [("W", 50, "REMOVE RK")],
    "W": [("WS", -50, "ATTACH RK"), ("WX", -50, "ATTACH RK"), ("T", -50, "ATTACH RK"), ("B", -50, "ATTACH RK"), ("D", -50, "ATTACH RK"), ("PC", -50, "ATTACH RK")],
    "WE": [("WSE", -50, "ATTACH RK")],
    "T": [("W", 50, "-RK")],
    "WSE": [("WE", 50, "-RK")],
    "WX": [("W", 50, "-RK")],
    "B": [("W", 50, "-RK")]
}

VV_MAP = {
    "T": {"T"},
    "WX": {"WX", "T"},
    "B": {"B"},
    "D": {"D"}
}

def get_modify_remark(new_base, old_base):
    for ob, remark in MODIFICATION_RULES.get(new_base, []):
        if ob == old_base:
            return remark
    if new_base == "K" and old_base in {"K", "B"}: return "ALT"
    if new_base == "KC" and old_base in {"KC", "BC"}: return "ALT"
    return "ALT"

def add_alt_bases_for_modify(b):
    """Return set of inventory BASE codes compatible for modify (cut) operations."""
    bases = {b}
    for ob, _ in MODIFICATION_RULES.get(b, []):
        bases.add(ob)
    return bases

def area(wl, wr, LL, LR):
    try:
        return (float(wl) + float(wr)) * (float(LL) + float(LR)) / 1_000_000.0
    except Exception:
        return 0.0

def read_dims(series):
    return pd.to_numeric(series, errors="coerce").fillna(0).round().astype(int)

def build_col_map(df):
    cmap = {}
    for c in df.columns:
        if c is None:
            continue
        key = str(c).strip().lower()
        cmap[key] = c
    return cmap

def find_col(cmap, candidates):
    for cand in candidates:
        k = cand.strip().lower()
        if k in cmap:
            return cmap[k]
    return None

def process_excel(in_path, out_path, progress_callback=None, tolerance=10, cut_mode='Both', operation_mode='Both'):
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

    old_records = old.to_dict('records')
    reuse_idx = defaultdict(list)
    mod_idx = defaultdict(list)
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
    modify_new_indices = set()
    modify_old_indices = set()

    alt_map_remark = MODIFICATION_RULES
    rk_map = RK_MODIFICATION_RULES
    vv_map = VV_MAP

    total = len(new)

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

    # Pre-filter EC inventory list for O(1) matching
    ec_inv_list = [inv for inv in inv_list if str(inv["CODE"]).startswith("EC") or inv["BASE"] == "EC"]

    # -------- STAGE 1: REUSE SWEEP (AS-IS -> ALT -> VV -> RK) -----------
    # Skip REUSE stage if operation_mode is 'MODIFY Only'
    if operation_mode != 'MODIFY Only':
        for i in range(len(new_data_tuples)):
            wlN, wrN, LLN, LRN, baseN, codeN = new_data_tuples[i]
            matched = False

            new_meta = get_new_meta(i)

            # --- Rule 7: EC family ignores WL/WR entirely, both for reuse
            # and modify.
            if baseN == "EC" or str(codeN).startswith("EC"):
                for inv in ec_inv_list:
                    if inv["USED"]:
                        continue
                    if not str(inv["CODE"]).startswith("EC"):
                        continue
                    if inv["LL"] == LLN:
                        idx_old = inv["old_idx"]
                        old_meta = get_old_meta_by_idx(idx_old)
                        reuse_rows.append(
                            list(new_meta) + list(old_meta) + [
                                area(wlN, wrN, LLN, LRN),
                                area(old_meta[1], old_meta[2], old_meta[4], old_meta[5]),
                                "AS-IS (EC, LL match only)",
                            ]
                        )
                        used_old[idx_old] = True
                        inv["USED"] = True
                        reuse_new_indices.add(i)
                        reuse_old_indices.add(idx_old)
                        matched = True
                        break
                if matched:
                    continue

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

            if baseN in alt_map_remark:
                for alt_base, remark in alt_map_remark[baseN]:
                    for idx_old in reuse_idx.get((wlN, wrN, alt_base, LLN, LRN), []):
                        if not used_old[idx_old]:
                            old_meta = get_old_meta_by_idx(idx_old)
                            reuse_rows.append(
                                list(new_meta) + list(old_meta) + [
                                    area(wlN, wrN, LLN, LRN),
                                    area(old_meta[1], old_meta[2], old_meta[4], old_meta[5]),
                                    f"{remark} ({baseN}↔{alt_base})" if remark != "ALT" else f"ALT ({baseN}↔{alt_base})",
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

            if baseN in vv_map:
                vv_candidates = vv_map[baseN]
                for cand_base in vv_candidates:
                    # VV: swap width and length dimensions (wl,wr,LL,LR) -> (LL,LR,wl,wr)
                    for idx_old in reuse_idx.get((LLN, LRN, cand_base, wlN, wrN), []):
                        if not used_old[idx_old]:
                            old_meta = get_old_meta_by_idx(idx_old)
                            reuse_rows.append(
                                list(new_meta) + list(old_meta) + [
                                    area(wlN, wrN, LLN, LRN),
                                    area(old_meta[1], old_meta[2], old_meta[4], old_meta[5]),
                                    f"VV ({baseN}↔{cand_base})",
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

            if baseN in rk_map:
                for (old_base, diff, remark) in rk_map[baseN]:
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

            if (i + 1) % 500 == 0 or (i + 1) == len(new_data_tuples):
                percent = 10 + int(((i + 1) / len(new_data_tuples)) * 70)
                update_progress(percent, f"Processing REUSE {i + 1}/{len(new_data_tuples)}")

    # If we are running in REUSE Only mode, capture all remaining NEW rows
    # (i.e. not matched in REUSE / RK) into the Balance (Not Reused/Modified)
    # bucket so that SUMMARY and the "Not Reuse & Modify" sheet reflect the
    # full New sheet area.
    # If we are running in REUSE Only mode, capture all remaining NEW rows
    # (i.e. not matched in REUSE / RK) into the Balance (Not Reused/Modified)
    # bucket so that SUMMARY and the "Not Reuse & Modify" sheet reflect the
    # full New sheet area.
    if operation_mode == 'REUSE Only':
        for i in range(len(new_data_tuples)):
            if i in reuse_new_indices or i in rk_new_indices:
                continue
            new_meta = get_new_meta(i)
            none_rows.append(
                list(new_meta) + [
                    area(new_meta[1], new_meta[2], new_meta[4], new_meta[5])
                ]
            )
    # 🔧 FIX: In MODIFY Only mode, ensure inventory is fully available
    if operation_mode == 'MODIFY Only':
        for inv in inv_list:
            inv["USED"] = False
        used_old = [False] * len(old)

    # -------- STAGE 2: MODIFY SWEEP (only on remaining New rows) ----------
    scrap_pool = []
    scrap_by_code = defaultdict(list)

    def add_to_scrap_pool(entry):
        scrap_pool.append(entry)
        scrap_by_code[entry.get("CODE", "")].append(entry)

    # Precompute remaining size counts for ultra-fast O(1) remainder check
    future_size_counts = defaultdict(lambda: defaultdict(int))
    for j in range(len(new_data_tuples)):
        if j in reuse_new_indices or j in rk_new_indices:
            continue
        wl_n, wr_n, ll_n, lr_n, base_n, code_n = new_data_tuples[j]
        future_size_counts[code_n][(wr_n, ll_n)] += 1

    def can_consume_remainder(rem_code, rem_w, rem_l, _start_idx=0):
        sizes = future_size_counts.get(rem_code)
        if not sizes:
            return False
        for (u_w, u_l), count in sizes.items():
            if count > 0 and rem_w >= u_w and rem_l >= u_l:
                return True
        return False

    def consume_inventory(inv_entry, req_idx, req_row_tuple, source_label):
        inv_entry["USED"] = True
        used_old_idx = inv_entry.get("old_idx", -1)
        if 0 <= used_old_idx < len(used_old):
            used_old[used_old_idx] = True

        req_wl, req_wr, req_ll, req_lr, req_base, req_code = req_row_tuple

        scrap_w = scrap_l = 0
        inv_wl = inv_entry['WL']
        inv_wr = inv_entry['WR']
        inv_ll = inv_entry['LL']
        inv_lr = inv_entry['LR']

        # Rule: 'Length' / 'Width' single-axis cutting only applies to the
        # whitelisted base codes; every other base always cuts in 'Both'.
        eff_cut_mode = cut_mode
        if eff_cut_mode in ('Length', 'Width') and req_base not in LENGTH_WIDTH_ONLY_BASES:
            eff_cut_mode = 'Both'

        if eff_cut_mode == 'Length':
            if inv_wl == req_wl and inv_wr == req_wr:
                diff_ll = inv_ll - req_ll
                if diff_ll >= tolerance:
                    scrap_w = inv_wr
                    scrap_l = diff_ll - tolerance
                else:
                    inv_entry["USED"] = False
                    if 0 <= used_old_idx < len(used_old):
                        used_old[used_old_idx] = False
                    return None
        elif eff_cut_mode == 'Width':
            if inv_ll == req_ll and inv_lr == req_lr:
                diff_wr = inv_wr - req_wr
                if diff_wr >= tolerance:
                    scrap_w = diff_wr - tolerance
                    scrap_l = inv_ll
                else:
                    inv_entry["USED"] = False
                    if 0 <= used_old_idx < len(used_old):
                        used_old[used_old_idx] = False
                    return None
        else:
            scrap_w = 0
            scrap_l = 0
            if inv_ll > req_ll:
                scrap_l = inv_ll - req_ll - tolerance
            if inv_wr > req_wr:
                scrap_w = inv_wr - req_wr - tolerance

        if (scrap_w >= MIN_USABLE_REMAINDER_MM or scrap_l >= MIN_USABLE_REMAINDER_MM):
            add_to_scrap_pool({
                'WL': inv_entry["WL"],
                'WR': int(scrap_w) if scrap_w > 0 else inv_entry["WR"],
                'CODE': inv_entry["CODE"],
                'LL': int(scrap_l) if scrap_l > 0 else inv_entry["LL"],
                'LR': inv_entry['LR'],
                'Balance': 1,
                'Origin_Row': req_idx + 1
            })

        is_scrap = (source_label == 'SCRAP')
        rem_wl = inv_entry['WL']
        rem_wr = int(scrap_w) if scrap_w > 0 else inv_entry['WR']
        rem_ll = int(scrap_l) if scrap_l > 0 else inv_entry['LL']
        width_str = f"{rem_wl}+{rem_wr}" if rem_wl > 0 else f"{rem_wr}"
        rem_str = f"{width_str} {req_code} {rem_ll}"
        details = {
            'Status': 'MODIFY',
            'Source': 'FROM REMAINDER' if is_scrap else 'Inventory',
            'Source_Detail': f"{inv_entry['WR']} {req_code} {inv_entry['LL']}" if is_scrap else '',
            'Inv_WL': inv_entry['WL'],
            'Inv_WR': inv_entry['WR'],
            'Inv_CODE': inv_entry['CODE'],
            'Inv_LL': inv_entry['LL'],
            'Inv_LR': inv_entry['LR'],
            'Inv_QTY_USED': 1,
            'Inv_old_idx': inv_entry.get('old_idx'),
            'Remainder': rem_str if (scrap_w > 0 or scrap_l > 0) else None,
            'Remainder_Used': '',
            'Modify_Remark': get_modify_remark(req_base, inv_entry.get("BASE", "")) if not is_scrap else 'FROM REMAINDER'
        }
        return details

    if operation_mode != 'REUSE Only':
        for i in range(len(new_data_tuples)):
            if i in reuse_new_indices or i in rk_new_indices:
                continue

            req_tuple = new_data_tuples[i]
            wlN, wrN, LLN, LRN, baseN, codeN = req_tuple
            new_meta = get_new_meta(i)

            if future_size_counts[codeN][(wrN, LLN)] > 0:
                future_size_counts[codeN][(wrN, LLN)] -= 1

            details = None
            best_scrap = None
            best_scrap_score = -1
            is_max_remainder_mode = baseN.startswith("SC") or baseN.startswith("IC") or baseN in {"K", "BIC"}
            
            for s in scrap_by_code.get(codeN, []):
                if s.get('Balance', 0) > 0:
                    if s['WL'] >= wlN and s['WR'] >= wrN and s['LL'] >= LLN and s['LR'] >= LRN:
                        sw = s['WR'] - wrN - tolerance if s['WR'] > wrN else 0
                        sl = s['LL'] - LLN - tolerance if s['LL'] > LLN else 0
                        waste = max(sw, 0) + max(sl, 0)
                        
                        score = waste if is_max_remainder_mode else (1000000 - waste)
                        if score > best_scrap_score:
                            best_scrap_score = score
                            best_scrap = s
                            
            if best_scrap is not None:
                s = best_scrap
                s['Balance'] = 0
                details = consume_inventory(s, i, req_tuple, source_label='SCRAP')
                modify_new_indices.add(i)

            if details is None:
                candidate = None
                best_score = -1
                candidate_bases = add_alt_bases_for_modify(baseN)
                is_special = (baseN in {"EC", "ICB", "ICK", "ICT", "BIC", "K", "KC", "KCE"} or codeN.startswith("EC") or codeN.startswith("SC") or codeN.startswith("SI") or codeN.startswith("SO"))
                
                # Filter candidates quickly using indexes
                inv_candidates = []
                seen_idx = set()
                for b in candidate_bases:
                    for inv in inv_by_base.get(b, []):
                        if not inv["USED"] and inv["old_idx"] not in seen_idx:
                            seen_idx.add(inv["old_idx"])
                            inv_candidates.append(inv)
                if is_special:
                    for inv in inv_by_code.get(codeN, []):
                        if not inv["USED"] and inv["old_idx"] not in seen_idx:
                            seen_idx.add(inv["old_idx"])
                            inv_candidates.append(inv)
                    if codeN.startswith("EC") or baseN == "EC":
                        for inv in inv_by_base.get("EC", []):
                            if not inv["USED"] and inv["old_idx"] not in seen_idx:
                                seen_idx.add(inv["old_idx"])
                                inv_candidates.append(inv)
                    elif baseN in {"ICK", "ICT", "BIC"}:
                        for alt_b in {"ICB", "ICK", "ICT", "BIC", "ICX", "BC"}:
                            for inv in inv_by_base.get(alt_b, []):
                                if not inv["USED"] and inv["old_idx"] not in seen_idx:
                                    seen_idx.add(inv["old_idx"])
                                    inv_candidates.append(inv)
                    elif baseN == "KC":
                        for alt_b in {"KC", "BC"}:
                            for inv in inv_by_base.get(alt_b, []):
                                if not inv["USED"] and inv["old_idx"] not in seen_idx:
                                    seen_idx.add(inv["old_idx"])
                                    inv_candidates.append(inv)
                    elif baseN == "K":
                        for alt_b in {"K", "B"}:
                            for inv in inv_by_base.get(alt_b, []):
                                if not inv["USED"] and inv["old_idx"] not in seen_idx:
                                    seen_idx.add(inv["old_idx"])
                                    inv_candidates.append(inv)

                for inv in inv_candidates:
                    if inv["USED"]:
                        continue
                        
                    inv_wl, inv_wr, inv_ll, inv_lr = inv["WL"], inv["WR"], inv["LL"], inv["LR"]
                    inv_code = inv["CODE"]
                    
                    matched_special = False
                    if is_special:
                        if codeN.startswith("EC") or baseN == "EC":
                            if inv_code.startswith("EC") and inv_ll >= LLN + tolerance:
                                matched_special = True
                        elif baseN == "ICB":
                            if inv_code == "ICB" and inv_wl == wlN and inv_wr == wrN and inv_ll >= LLN + tolerance:
                                matched_special = True
                        elif baseN in {"ICK", "ICT"}:
                            if inv_code in {"ICB", "ICK", "ICT", "BIC"} and inv_wl == wlN and inv_wr == wrN and inv_ll >= LLN + tolerance:
                                matched_special = True
                        elif baseN == "BIC":
                            if inv_code in {"ICB", "ICK", "ICT", "BIC", "ICX", "BC"} and inv_wl == wlN and inv_wr == wrN and inv_ll >= LLN + tolerance:
                                matched_special = True
                        elif codeN.startswith("SC"):
                            if inv_code == codeN and inv_wl == wlN and inv_wr == wrN and inv_ll >= LLN + tolerance:
                                matched_special = True
                        elif baseN == "K":
                            if inv_code in {"K", "B"} and inv_wl == wlN and inv_wr == wrN and inv_ll >= LLN + tolerance:
                                matched_special = True
                        elif codeN.startswith("SI") or codeN.startswith("SO"):
                            if inv_code == codeN and inv_wl == wlN and inv_wr == wrN and inv_ll >= LLN + tolerance and inv_lr >= LRN + tolerance:
                                matched_special = True
                        elif baseN == "KC":
                            if inv_code in {"KC", "BC"} and inv_wl == wlN and inv_wr == wrN and inv_ll >= LLN + tolerance and inv_lr >= LRN + tolerance:
                                matched_special = True
                        elif baseN == "KCE":
                            if inv_code == "KCE" and inv_wl == wlN and inv_wr == wrN and inv_ll >= LLN + tolerance and inv_lr >= LRN + tolerance:
                                matched_special = True
                        
                        if not matched_special:
                            continue  # Special codes do not fallback to standard checks
                            
                    if not matched_special:
                        if inv["BASE"] not in candidate_bases:
                            continue
                            
                        if baseN == "D" and inv["BASE"] == "DE" and inv_wr <= 230:
                            continue
                            
                        eff_cut_mode = cut_mode
                        if eff_cut_mode in ('Length', 'Width') and baseN not in LENGTH_WIDTH_ONLY_BASES:
                            eff_cut_mode = 'Both'
                            
                        is_valid = False
                        is_vv = False
                        
                        if eff_cut_mode == 'Length':
                            if (inv_wl == wlN and inv_wr == wrN and inv_ll >= LLN + tolerance and inv_lr >= LRN):
                                is_valid = True
                        elif eff_cut_mode == 'Width':
                            if (inv_ll == LLN and inv_lr == LRN and inv_wl >= wlN and inv_wr >= wrN + tolerance):
                                is_valid = True
                        else:
                            if (inv_wl >= wlN and inv_wr >= wrN and inv_ll >= LLN and inv_lr >= LRN):
                                if (inv_wr >= wrN + tolerance) or (inv_ll >= LLN + tolerance):
                                    is_valid = True
                                    
                        if not is_valid:
                            vv_bases = vv_map.get(baseN, set())
                            if inv["BASE"] in vv_bases or baseN in vv_bases:
                                if (inv_ll == wlN and inv_lr == wrN) and (inv_wl >= LLN + tolerance or inv_wr >= LRN + tolerance):
                                    is_valid = True
                                    is_vv = True
                                elif (inv_wl == LLN and inv_wr == LRN) and (inv_ll >= wlN + tolerance or inv_lr >= wrN + tolerance):
                                    is_valid = True
                                    is_vv = True
                                    
                        if not is_valid:
                            continue

                    # Calculate scrap to score this candidate
                    sw = 0
                    sl = 0
                    if not matched_special:
                        if inv_ll > LLN:
                            sl = inv_ll - LLN - tolerance
                        if inv_wr > wrN:
                            sw = inv_wr - wrN - tolerance
                    else:
                        if inv_ll > LLN:
                            sl = inv_ll - LLN - tolerance
                            
                    is_max_remainder_mode = baseN.startswith("SC") or baseN.startswith("IC") or baseN in {"K", "BIC"}
                    waste = max(sw, 0) + max(sl, 0)
                    
                    if is_max_remainder_mode:
                        score = 3000000 + waste # Highest score for max remainder
                    else:
                        is_consumable = False
                        if sl >= MIN_USABLE_REMAINDER_MM or sw >= MIN_USABLE_REMAINDER_MM:
                            is_consumable = can_consume_remainder(inv_code, max(sw, 0), max(sl, 0), i)
                            
                        score = 0
                        if sl >= MIN_USABLE_REMAINDER_MM or sw >= MIN_USABLE_REMAINDER_MM:
                            if is_consumable:
                                score = 2000000 - (inv_wl * inv_wr + inv_ll * inv_lr) # Tie-breaker: smallest inventory panel first
                            else:
                                continue # REJECT
                        else:
                            score = 1000000 - waste # Tie-breaker: smallest waste first
                        
                    if score > best_score:
                        best_score = score
                        candidate = inv
                if candidate is not None:
                    details = consume_inventory(candidate, i, req_tuple, source_label='INVENTORY')
                    modify_new_indices.add(i)
                    modify_old_indices.add(candidate["old_idx"])
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
                area_old_val = area(old_meta[1] or 0, inv_w_val or 0, inv_l_val or 0, old_meta[5] or 0)
                remark_str = f"MODIFY ({details.get('Modify_Remark', 'ALT')})" if details.get('Source') == 'Inventory' else f"MODIFY ({details.get('Source')})"
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
                if (i + 1) % 500 == 0 or (i + 1) == len(new_data_tuples):
                    percent = 80 + int(((i + 1) / len(new_data_tuples)) * 15)
                    update_progress(percent, f"Processing MODIFY {i + 1}/{len(new_data_tuples)}")
            else:
                none_rows.append(
                    list(new_meta) + [
                        area(new_meta[1], new_meta[2], new_meta[4], new_meta[5])
                    ]
                )

    for idx, ro in enumerate(old_records):
        if not used_old[idx]:
            wlO, wrO, LLO, LRO = int(ro.get(wl_col_old, 0)), int(ro.get(wr_col_old, 0)), int(ro.get(ll_col_old, 0)), int(ro.get(lr_col_old, 0))
            unused_rows.append([wlO, wrO, ro.get(code_col_old, ""), LLO, LRO, area(wlO, wrO, LLO, LRO)])

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

    # --- NEW: Build MERGED sheet using fast list building ---
    merged_cols = new_headers + old_headers + ["Remainder", "Remainder_Used", "Status", "Source_Detail", "Area New (m^2)", "Area Old (m^2)", "Remark"]
    merged_rows = []

    # Fast conversion of reuse / rk rows
    for r in reuse_rows + rk_reuse_rows:
        # r structure: new_meta (11) + old_meta (9) + [area_new, area_old, remark]
        rem = r[22] if len(r) > 22 else ""
        source = "AS-IS"
        if isinstance(rem, str):
            if "AS-IS" in rem.upper():
                source = "AS-IS"
            elif "RK" in rem.upper():
                source = "RK REUSE"
            elif "ALT" in rem.upper() or "VV" in rem.upper():
                source = "ALT/AS-IS"
            else:
                source = rem
        merged_rows.append(r[:20] + ["", "", source, "", r[20], r[21], rem])

    # Fast conversion of modify rows
    for r in modify_rows:
        # r structure: new_meta (11) + old_meta (9) + [rem, rem_used, src, src_det, an, ao, remk]
        merged_rows.append(r[:20] + [r[20], r[21], r[22], r[23], r[24], r[25], r[26]])

    # Fast conversion of none rows
    for r in none_rows:
        # r structure: new_meta (11) + [area_m2]
        merged_rows.append(r[:11] + [""] * 9 + ["", "", "NEW", "", r[11], "", "NEW"])

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
        df_location = pd.concat([df_location, pd.DataFrame([perc_row], columns=df_location.columns)], ignore_index=True)
    except Exception as e:
        df_location = pd.DataFrame(columns=["Sr. No.", "Location", "AS-IT-IS Area", "MODIFY Area", "NEW Area", "Sub-Total"])
        print(f"(Location summary generation failed: {e})")

    # Write all sheets with INFO as first sheet
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        info_data = {"": ["Old = Inventory List", "New = Material List"]}
        df_info = pd.DataFrame(info_data)
        df_info.to_excel(xw, sheet_name="INFO", index=False)
        
        df_merged.to_excel(xw, sheet_name="MERGED", index=False)
        df_all_reuse.to_excel(xw, sheet_name="REUSED", index=False)
        df_modify.to_excel(xw, sheet_name="MODIFY", index=False)
        df_unused.to_excel(xw, sheet_name="UNUSED STOCK", index=False)
        df_summary.to_excel(xw, sheet_name="SUMMARY", index=False, startrow=0, startcol=0)
        startcol = df_summary.shape[1] + 2
        try:
            df_location.to_excel(xw, sheet_name="SUMMARY", index=False, startrow=0, startcol=startcol)
        except Exception:
            df_location.to_excel(xw, sheet_name="LOCATION SUMMARY", index=False)
        df_none.to_excel(xw, sheet_name="NEW", index=False)
        df_scrap.to_excel(xw, sheet_name="SCRAP_POOL", index=False)

    # --- Copy highlighted "New" and "Old" sheets into output file using captured indices ---
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
        from openpyxl.utils.exceptions import InvalidFileException
        import zipfile, io, tempfile, shutil
        from io import BytesIO

        green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

        def safe_load_workbook(path):
            for attempt in range(3):
                try:
                    with open(path, "rb") as f:
                        data = f.read()
                    return load_workbook(BytesIO(data), data_only=True)
                except (InvalidFileException, KeyError, zipfile.BadZipFile, OSError) as e:
                    if attempt < 2:
                        time.sleep(1.5)
                    else:
                        raise e

        wb_in = safe_load_workbook(in_path)
        wb_out = safe_load_workbook(out_path)
        
        # Format INFO sheet
        ws_info = wb_out["INFO"]
        ws_info.delete_cols(1)
        ws_info.cell(2, 1, "Old = Inventory List")
        ws_info.cell(3, 1, "New = Material List")
        
        bold_font = Font(name='Arial', size=18, bold=True)
        ws_info.cell(2, 1).font = bold_font
        ws_info.cell(3, 1).font = bold_font
        ws_info.column_dimensions['A'].width = 30
        
        ws_info.cell(5, 1, "LEGENDS")
        ws_info.cell(6, 1, "REUSED").fill = green
        ws_info.cell(7, 1, "MODIFY").fill = yellow
        ws_info.cell(8, 1, "NEW")
        
        thin = Side(border_style='thin', color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for r in range(5, 9):
            ws_info.cell(r, 1).border = border
            ws_info.cell(r, 1).font = Font(name='Arial', size=11, bold=True)

        # Copy "New" sheet and apply highlights
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
            
            # Add legend
            ws_out.cell(3, 18, "LEGENDS")
            ws_out.cell(4, 18, "REUSED").fill = green
            ws_out.cell(5, 18, "MODIFY").fill = yellow
            ws_out.cell(6, 18, "NEW")
            for r in range(3, 7):
                ws_out.cell(r, 18).border = border

        # Copy "Old" sheet and highlight rows used
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
            
            # Add legend
            ws_out2.cell(3, 18, "LEGENDS")
            ws_out2.cell(4, 18, "REUSED").fill = green
            ws_out2.cell(5, 18, "MODIFY").fill = yellow
            ws_out2.cell(6, 18, "NEW")
            for r in range(3, 7):
                ws_out2.cell(r, 18).border = border

        # Apply formatting efficiently
        align = Alignment(horizontal='center', vertical='center', wrap_text=False)
        for ws in wb_out.worksheets:
            max_r = ws.max_row or 1
            last_data_col = 1
            for c in range(1, (ws.max_column or 1) + 1):
                if ws.cell(1, c).value is not None and str(ws.cell(1, c).value).strip() != "":
                    last_data_col = c
            
            style_limit = max_r if (ws.title in {"INFO", "SUMMARY", "LOCATION SUMMARY", "SCRAP_POOL"} or max_r <= 500) else 500
            for row in ws.iter_rows(min_row=1, max_row=style_limit, min_col=1, max_col=last_data_col):
                for cell in row:
                    cell.alignment = align
                    cell.border = border

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