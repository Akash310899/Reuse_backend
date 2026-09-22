import pandas as pd
from backend.core.constants import MODIFY_CODE_RULES

def clean_code_raw(s):
    if pd.isna(s):
        return ""
    t = str(s).strip().upper()
    for ch in ["-", "\u2014", "_", ".", "-"]:
        t = t.replace(ch, "")
    if t in {"", "NA", "N/A", ".", "-"}:
        return ""
    if t.startswith("NSP"):
        t = t[3:]
    return t

def base_code(s):
    """
    Map a raw panel code to its canonical base code for indexing.
    
    Grouping rules:
    - PH, CP, PPH, CPP, CPPP -> "PH"  (Rule 16: interchangeable)
    - D, SB -> "D"                      (Rule 11)
    - DE, SBE -> "DE"                   (Rule 12)
    - CC, CCL, CCR -> "CC"
    - JB, BB, BEAMBAR, JOINTBAR -> "JB" (Rule 19: interchangeable)
    - SC/SI/SO variants -> raw code     (Rule 13-15: each matches only itself)
    - EC variants -> raw code           (Rule 10: each matches only itself)
    - ICB, ICK, ICT, ICX, BIC -> each keeps own base (Rule 8-9: ALT handled separately)
    - K, B -> each keeps own base       (Rule 22: ALT handled separately)
    - WT, TX -> own base codes          (Rule 3: same reuse logic as WS)
    """
    t = clean_code_raw(s)
    if not t:
        return ""
    
    # CC family
    if t in {"CC", "CCL", "CCR"}: return "CC"
    
    # D family
    if t in {"D", "SB"}: return "D"
    if t in {"DE", "SBE"}: return "DE"
    if t == "DG": return "DG"
    if t == "DP": return "DP"
    
    # EC family - each variant keeps its own code (Rule 10: exact match only)
    if t.startswith("EC"): return t
    if t == "CA": return "CA"
    
    # EB, MB - exact match only
    if t == "EB": return "EB"
    if t == "MB": return "MB"
    
    # IC sub-family - each keeps own base (Rule 8-9: ALT rules handle interchangeability)
    if t == "ICB": return "ICB"
    if t == "ICK": return "ICK"
    if t == "ICT": return "ICT"
    if t == "ICX": return "ICX"
    if t == "BIC": return "BIC"
    if t in {"IC", "ICR", "NSP"}: return "IC"
    
    # K sub-family
    if t == "K": return "K"
    if t == "KC": return "KC"
    if t == "KCE": return "KCE"
    if t == "KIC": return "KIC"
    
    # PC family
    if t in {"PC", "PCE"}: return "PC"
    
    # PH family (Rule 16: interchangeable)
    if t in {"PH", "CP", "PPH", "CPP", "CPPP"}: return "PH"
    
    # SX family
    if t in {"SX", "SXC", "SXCE"}: return "SX"
    
    # SC / SI / SO families - each variant is its own base (Rule 13-15: exact match)
    if t.startswith("SC"): return t
    if t.startswith("SI"): return t
    if t.startswith("SO"): return t
    
    # SL family
    if t in {"SL", "SLWL", "SLR"}: return "SL"
    if t == "LS": return "LS"
    if t == "LSS": return "LSS"
    
    # Panel types with RK relationships
    if t == "T": return "T"
    if t == "TE": return "TE"
    if t == "B": return "B"
    if t == "WT": return "WT"
    if t == "TX": return "TX"
    if t in {"W", "WRB", "WRBS"}: return "W"
    if t == "WE": return "WE"
    if t == "WSE": return "WSE"
    if t == "WS": return "WS"
    if t == "WX": return "WX"
    if t == "WXE": return "WXE"
    
    # Misc
    if t in {"XBS", "LSC", "LSCY", "LSCZ"}: return "XBS"
    if t == "BHH": return "BHH"
    if t == "PLB": return "PLB"
    if t == "RK": return "RK"
    
    # Joint Bar family (Rule 19: interchangeable)
    if t in {"JB", "BB", "BEAMBAR", "JOINTBAR"}: return "JB"
    
    if t in {"HCC", "SPD"}: return t
    
    return t

def get_modify_remark(new_base, old_base, old_code=""):
    """Look up the remark for a modify match from MODIFY_CODE_RULES."""
    rules = resolve_modify_rules(new_base, old_code)
    if rules is None or rules == "NO_MODIFY":
        return "ALT"
    for (code, remark, _constraint) in rules:
        if code == old_base or code == old_code:
            return remark
    return "ALT"

def resolve_modify_rules(base_code_val, raw_code=""):
    """Resolve the modify rules for a given base code.
    Handles SC*/SI*/SO* dynamic matching and NO_MODIFY.
    Returns a list of (old_code, remark, constraint) or 'NO_MODIFY' or None.
    """
    # Check direct lookup first
    rules = MODIFY_CODE_RULES.get(base_code_val)
    if rules is not None:
        if rules == "NO_MODIFY":
            return "NO_MODIFY"
        return rules

    # SC*/SI*/SO* dynamic: exact code match, WL/WR exact, cut LL
    if raw_code.startswith("SC") or raw_code.startswith("SI") or raw_code.startswith("SO"):
        return [(raw_code, "ALT", "WL_WR_EXACT_LL")]

    # EC variants (base_code returns raw code for EC)
    if base_code_val.startswith("EC"):
        ec_rules = MODIFY_CODE_RULES.get("EC", [])
        if ec_rules and ec_rules != "NO_MODIFY":
            return ec_rules

    return None  # No modify rules found

def area(wl, wr, LL, LR):
    try:
        return (float(wl) + float(wr)) * (float(LL) + float(LR)) / 1_000_000.0
    except Exception:
        return 0.0

def read_dims(series):
    return pd.to_numeric(series, errors="coerce").fillna(0).round().astype(int)
