import pandas as pd
import time

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

def prepare_dataframes(in_path):
    xl = pd.ExcelFile(in_path)
    sheet_names = xl.sheet_names
    new_sheet = next((s for s in sheet_names if s.lower() == "new"), None)
    old_sheet = next((s for s in sheet_names if s.lower() == "old"), None)
    
    if not new_sheet or not old_sheet:
        raise ValueError("Excel file must contain sheets named 'New' and 'Old' (case-insensitive).")
    
    new = pd.read_excel(in_path, sheet_name=new_sheet, dtype=str)
    old = pd.read_excel(in_path, sheet_name=old_sheet, dtype=str)

    cmap_new = build_col_map(new)
    cmap_old = build_col_map(old)

    cols = {
        'new_spec': find_col(cmap_new, ["spec", "specification", "description"]),
        'new_flat': find_col(cmap_new, ["flat", "flat ", "flatname"]),
        'new_part': find_col(cmap_new, ["part", "parttype", "part type", "part_name"]),
        'new_qty': find_col(cmap_new, ["qty", "quantity", "qnty"]),
        'new_unit': find_col(cmap_new, ["unit", "unitname", "unit area", "unit_area"]),
        'new_loc': find_col(cmap_new, ["location", "loc", "locn"]),
        
        'old_spec': find_col(cmap_old, ["spec", "specification", "description"]),
        'old_qty': find_col(cmap_old, ["qty", "quantity", "qnty"]),
        'old_unit_area': find_col(cmap_old, ["unit area", "unitarea", "unit_area", "unit"]),
        'old_loc': find_col(cmap_old, ["location", "loc", "locn"]),

        'wl_new': find_col(cmap_new, ["wl", "widthleft", "width_l", "w_l"]),
        'wr_new': find_col(cmap_new, ["wr", "widthright", "width_r", "w_r"]),
        'll_new': find_col(cmap_new, ["ll", "lengthleft", "length_l", "l_l"]),
        'lr_new': find_col(cmap_new, ["lr", "lengthright", "length_r", "l_r"]),
        'code_new': find_col(cmap_new, ["code", "code ", "type", "code_name", "c_o_d_e"]),

        'wl_old': find_col(cmap_old, ["wl", "widthleft", "width_l", "w_l"]),
        'wr_old': find_col(cmap_old, ["wr", "widthright", "width_r", "w_r"]),
        'll_old': find_col(cmap_old, ["ll", "lengthleft", "length_l", "l_l"]),
        'lr_old': find_col(cmap_old, ["lr", "lengthright", "length_r", "l_r"]),
        'code_old': find_col(cmap_old, ["code", "code ", "type", "code_name", "c_o_d_e"])
    }

    synthetic_counter = 0
    mappings = [
        ("wl_new", "0", new), ("wr_new", "0", new),
        ("ll_new", "0", new), ("lr_new", "0", new),
        ("code_new", "", new),
        ("wl_old", "0", old), ("wr_old", "0", old),
        ("ll_old", "0", old), ("lr_old", "0", old),
        ("code_old", "", old)
    ]
    
    for key, default_val, df in mappings:
        if cols[key] is None:
            if default_val == "":
                synthetic = "CODE_SYNTHETIC"
            else:
                synthetic_counter += 1
                synthetic = f"NUM_SYNTH_{int(time.time()*1000)}_{synthetic_counter}"
            df[synthetic] = default_val
            cols[key] = synthetic

    for key in ['wl_new', 'wr_new', 'll_new', 'lr_new', 'wl_old', 'wr_old', 'll_old', 'lr_old']:
        if not cols[key]: cols[key] = key.split('_')[0].upper()
    for key in ['code_new', 'code_old']:
        if not cols[key]: cols[key] = "CODE"

    for col in [cols['wl_new'], cols['wr_new'], cols['ll_new'], cols['lr_new']]:
        if col not in new.columns: new[col] = "0"
        new[col] = read_dims(new[col])
        
    for col in [cols['wl_old'], cols['wr_old'], cols['ll_old'], cols['lr_old']]:
        if col not in old.columns: old[col] = "0"
        old[col] = read_dims(old[col])

    return new, old, cols
