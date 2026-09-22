RK_MAX = 50  # mm (RK tolerance for length adjustments)

MIN_USABLE_REMAINDER_MM = 200
MAX_ACCEPTABLE_SCRAP_FRACTION = 0.20

# Cut modes only apply to these base codes
CUT_MODE_BASES = {"W", "WS"}

LENGTH_WIDTH_ONLY_BASES = {"W", "WS", "T", "B", "WE", "WSE", "TE", "D", "WXE"}

# ============================================================================
# REUSE STAGE RULES (separate from modify)
# ============================================================================

REUSE_ALT_RULES = {
    "WSE": [("WS", "ADD EC"), ("T", "ADD EC"), ("B", "ADD EC"), ("WT", "ADD EC"), ("TX", "ADD EC")],
    "TE": [("WS", "ADD EC"), ("T", "ADD EC"), ("B", "ADD EC"), ("WT", "ADD EC"), ("TX", "ADD EC")],
    "WXE": [("WS", "ADD EC"), ("T", "ADD EC"), ("B", "ADD EC"), ("WT", "ADD EC"), ("TX", "ADD EC")],
    "ICT": [("ICK", "ALT"), ("ICX", "ALT"), ("BIC", "ALT")],
    "ICK": [("ICT", "ALT"), ("ICX", "ALT"), ("BIC", "ALT")],
    "ICX": [("ICT", "ALT"), ("ICK", "ALT"), ("BIC", "ALT")],
    "BIC": [("ICT", "ALT"), ("ICK", "ALT"), ("ICX", "ALT")],
    "D": [("PC", "ALT")],
    "DE": [("D", "ATTACH EC"), ("PC", "ATTACH EC")],
    "K": [("B", "ALT")],
}

REUSE_RK_RULES = {
    "W": [("WS", -50, "ADD RK"), ("T", -50, "ADD RK"), ("B", -50, "ADD RK"),
          ("WX", -50, "ADD RK"), ("WT", -50, "ADD RK"), ("TX", -50, "ADD RK")],
    "WS": [("W", 50, "REMOVE RK")],
    "T": [("W", 50, "REMOVE RK")],
    "B": [("W", 50, "REMOVE RK")],
    "WX": [("W", 50, "REMOVE RK")],
    "WT": [("W", 50, "REMOVE RK")],
    "TX": [("W", 50, "REMOVE RK")],
    "WE": [("WSE", -50, "ADD RK"), ("TE", -50, "ADD RK"), ("WXE", -50, "ADD RK")],
    "WSE": [("WE", 50, "REMOVE RK"), ("W", 50, "ADD EC & REMOVE RK")],
    "TE": [("WE", 50, "REMOVE RK"), ("W", 50, "ADD EC & REMOVE RK")],
    "WXE": [("WE", 50, "REMOVE RK"), ("W", 50, "ADD EC & REMOVE RK")],
}


# ============================================================================
# MODIFY STAGE RULES
# ============================================================================

# Constraint types:
#   "FULL"              = inv WL>=req WL, WR>=req WR, LL>=req LL, LR>=req LR
#                         AND at least one dim exceeds by tolerance
#   "WL_WR_EXACT_LL"    = inv WL==req WL, WR==req WR, LL >= req LL + tolerance
#   "FULL+50LL"         = same as FULL but LL must be >= req LL + 50 + tolerance
#   (SC/SI/SO handled dynamically by raw code match with WL_WR_EXACT_LL)

# Entries: list of (old_base_or_code, remark, constraint)
# Special string values: "SAME_AS_<base>" = alias, "NO_MODIFY" = skip

_WS_RULES = [
    ("WS", "ALT", "FULL"), ("B", "ALT", "FULL"), ("T", "ALT", "FULL"),
    ("WX", "ALT", "FULL"), ("WT", "ALT", "FULL"), ("TX", "ALT", "FULL"),
    ("W", "REMOVE RK", "FULL+50LL"),
]

_WSE_RULES = [
    ("WSE", "ALT", "FULL"), ("TE", "ALT", "FULL"), ("WXE", "ALT", "FULL"), ("DE", "ALT", "FULL"),
    ("WE", "REMOVE RK", "FULL+50LL"),
    ("W", "ADD EC", "FULL"), ("T", "ADD EC", "FULL"), ("B", "ADD EC", "FULL"),
    ("WX", "ADD EC", "FULL"), ("WT", "ADD EC", "FULL"), ("TX", "ADD EC", "FULL"),
]

_ICK_RULES = [
    # Priority 1: SC* panels — same WL/WR, LL >= req (cross-family, no +tol needed)
    ("SC", "USE SC", "WL_WR_EXACT_LL_GE"),
    # Priority 2: IC family interchangeable
    ("ICK", "ALT", "WL_WR_EXACT_LL"), ("ICT", "ALT", "WL_WR_EXACT_LL"),
    ("ICX", "ALT", "WL_WR_EXACT_LL"), ("BIC", "ALT", "WL_WR_EXACT_LL"),
]

MODIFY_CODE_RULES = {
    # W: use W, B, WS, T, WX, WT, TX — all dims greater
    "W": [
        ("W", "ALT", "FULL"), ("B", "ALT", "FULL"), ("WS", "ALT", "FULL"),
        ("T", "ALT", "FULL"), ("WX", "ALT", "FULL"), ("WT", "ALT", "FULL"), ("TX", "ALT", "FULL"),
    ],

    # WS, T, B, WX, WT, TX — share rules
    "WS": _WS_RULES,
    "T":  _WS_RULES,
    "B":  _WS_RULES,
    "WX": _WS_RULES,
    "WT": _WS_RULES,
    "TX": _WS_RULES,

    # WE: use WE, WSE/TE/WXE (ADD RK), W/T/B/WX/WT/TX (ADD EC) — all size greater, no +50
    "WE": [
        ("WE", "ALT", "FULL"), ("WSE", "ADD RK", "FULL"), ("TE", "ADD RK", "FULL"), ("WXE", "ADD RK", "FULL"),
        ("W", "ADD EC", "FULL"), ("T", "ADD EC", "FULL"), ("B", "ADD EC", "FULL"),
        ("WX", "ADD EC", "FULL"), ("WT", "ADD EC", "FULL"), ("TX", "ADD EC", "FULL"),
    ],

    # WSE, TE, WXE — share rules
    "WSE": _WSE_RULES,
    "TE":  _WSE_RULES,
    "WXE": _WSE_RULES,

    # ICB: only ICB, exact WL/WR, cut LL
    "ICB": [("ICB", "ALT", "WL_WR_EXACT_LL")],

    # ICK, ICT, ICX, BIC: interchangeable, exact WL/WR, cut LL
    "ICK": _ICK_RULES,
    "ICT": _ICK_RULES,
    "ICX": _ICK_RULES,
    "BIC": _ICK_RULES,

    # EC: use EC/ECX/ECV/ECB/ECH, exact WL/WR, cut LL only
    "EC": [
        ("EC", "ALT", "WL_WR_EXACT_LL"), ("ECX", "ALT", "WL_WR_EXACT_LL"),
        ("ECV", "ALT", "WL_WR_EXACT_LL"), ("ECB", "ALT", "WL_WR_EXACT_LL"),
        ("ECH", "ALT", "WL_WR_EXACT_LL"),
    ],

    # D: use D, PC — all dims greater
    "D": [("D", "ALT", "FULL"), ("PC", "ALT", "FULL")],

    # DE: use DE, D (ADD EC), TE, WSE, WXE — all dims greater
    "DE": [
        ("DE", "ALT", "FULL"), ("D", "ADD EC", "FULL"),
        ("TE", "ALT", "FULL"), ("WSE", "ALT", "FULL"), ("WXE", "ALT", "FULL"),
    ],

    # PH: use PH, CP, CPP, PPH — all dims greater
    "PH": [("PH", "ALT", "FULL"), ("CP", "ALT", "FULL"), ("CPP", "ALT", "FULL"), ("PPH", "ALT", "FULL")],

    # PLB: only PLB
    "PLB": [("PLB", "ALT", "FULL")],

    # DP & Jointbar: NO MODIFY
    "DP": "NO_MODIFY",
    "JB": "NO_MODIFY",

    # EB: exact WL, WR, code — cut LL
    "EB": [("EB", "ALT", "WL_WR_EXACT_LL")],

    # MB: exact WL, WR, code — cut LL
    "MB": [("MB", "ALT", "WL_WR_EXACT_LL")],

    # K: use K, B — all dims greater
    "K": [("K", "ALT", "FULL"), ("B", "ALT", "FULL")],

    # KC: use KC, BC — exact WL/WR, cut LL
    "KC": [("KC", "ALT", "WL_WR_EXACT_LL"), ("BC", "ALT", "WL_WR_EXACT_LL")],

    # KCE: use KCE, BCE — exact WL/WR, cut LL
    "KCE": [("KCE", "ALT", "WL_WR_EXACT_LL"), ("BCE", "ALT", "WL_WR_EXACT_LL")],

    # BC: only BC
    "BC": [("BC", "ALT", "FULL")],
}

# Codes where remainder is calculated from LL only (WL/WR stay exact)
LL_ONLY_CUT_BASES = {"ICB", "ICK", "ICT", "ICX", "BIC", "K", "KC", "KCE", "EB", "MB"}

# Codes that prefer maximum remainder (larger panel preferred)
MAX_REMAINDER_BASES = {"ICB", "ICK", "ICT", "ICX", "BIC", "K"}

# ============================================================================
# SWAP REUSE GROUPS
# ============================================================================

# Group 1: 90° rotation — WR↔LL, WL↔LR
# e.g. stored as WL=500, WR=300, W, LL=200  can fill  WL=200, WR=300, W, LL=500
SWAP_REUSE_ROTATE_BASES = {"W", "WS", "B", "T", "D", "K", "PH", "PLB"}

# Group 2: Horizontal flip — WL↔WR (LL/LR unchanged)
# e.g. stored as WL=200, WR=100, ICK, LL=2400  can fill  WL=100, WR=200, ICK, LL=2400
SWAP_REUSE_FLIP_BASES = {"ICB", "ICX", "BIC", "ICK", "ICT"}
