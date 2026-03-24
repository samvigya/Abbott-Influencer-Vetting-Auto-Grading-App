"""
Abbott Influencer Vetting — Auto-Grader  v5.0
=============================================
For each influencer sheet, reads every column BY NAME (never by position),
counts YES values, computes GREEN/AMBER/RED using confirmed rules,
cross-checks against pre-computed Code/Point row where available,
then builds the grading sheet.

Verified scoring rules (from TH_Adult_Nutrition reference file):
  RISK params     → % of YES videos vs threshold → GREEN / AMBER / RED
  AUTO-REJECT     → count > 0 → RED (flag), count = 0 → GREEN
  RELEVANCE cols  → count > 0 → GREEN (present = points), count = 0 → RED
  Pregnant        → count > 0 → RED (reject/penalty), count = 0 → GREEN

Formats handled (auto-detected per sheet, all by column name):
  Format A  — response_1 col row 0 = "HEADER",     row 3 = Code/Point (GREEN/AMBER/RED)
  Format B  — individual Yes/No video rows, no summary
  Format C  — response_1 col row 0 = "No of occurences", row 3 = Verdict (GREEN/AMBER/RED)
  Format IG — Audio transcription col row 0 = "HEADER", row 3 = Code/Point
  Captions  — header row at index 4, counts at rows 0/1
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO

# ─────────────────────────────────────────────────────────────────────────────
# SCORING RULES  (single source of truth — verified against reference file)
# ─────────────────────────────────────────────────────────────────────────────

# Rule types
RULE_PCT      = "pct"       # % threshold based
RULE_BINARY_R = "binary_r"  # count > 0 → RED   (auto-reject / pregnant)
RULE_BINARY_G = "binary_g"  # count > 0 → GREEN (relevance: topics/kids/media)

# (ai_column_name, grading_label, rule_type, green_max_pct, amber_max_pct, points_if_positive)
COLUMN_RULES = [
    # ── AUTO-REJECT GATES (binary_r: any count → RED → Risk=10) ─────────────
    ("Substance Use Discussion",         "Substance Use",         RULE_BINARY_R, None,  None,  None),
    ("Breastfeeding",                    "Anti-breastfeeding",    RULE_BINARY_R, None,  None,  None),
    ("Vaccination",                      "Anti-vaccination",      RULE_BINARY_R, None,  None,  None),
    ("Health care stance",               "Anti-healthcare",       RULE_BINARY_R, None,  None,  None),
    # ── 8 RISK PARAMETERS (pct: % threshold → GREEN/AMBER/RED → 0/0.5/1 pts) ─
    ("Profanity Usage",                  "Profanity",             RULE_PCT,      0.15,  0.30,  None),
    ("Alcohol Use Discussion",           "Alcohol",               RULE_PCT,      0.05,  0.15,  None),
    ("Sensitive Visual Content",         "Sensitive content",     RULE_PCT,      0.20,  None,  None),
    ("Stereotypes or Bias",              "Stereotype & Bias",     RULE_PCT,      0.05,  0.15,  None),
    ("Violence Advocacy",                "Violence Content",      RULE_PCT,      0.001, 0.05,  None),  # 0% exact = GREEN
    ("Political Stance",                 "Political",             RULE_PCT,      0.10,  0.25,  None),
    ("Unscientific Claims",              "Unscientific",          RULE_PCT,      0.15,  None,  None),
    ("Ultra-processed food",             "Ultra-processed Food",  RULE_PCT,      0.15,  0.30,  None),
    # ── RELEVANCE — PEDIATRIC (binary_g: any count → GREEN → pts) ────────────
    ("Kids Presenence",                  "Kids Presence",         RULE_BINARY_G, None,  None,  3.5),
    # ── RELEVANCE — ADULT NUTRITION ───────────────────────────────────────────
    ("Topics around Adult Health",       "Topics Adult Health",   RULE_BINARY_G, None,  None,  5.0),
    ("Topics on Adult Healthy Nutrition","Topics Adult Nutrition",RULE_BINARY_G, None,  None,  2.0),
    # ── OTHER BINARY (used for relevance scoring) ─────────────────────────────
    ("Media Presence & Awards",          "Awards/Media Presence", RULE_BINARY_G, None,  None,  1.0),
    ("Brand Partnership Presence",       "Brand Partnerships",    RULE_BINARY_G, None,  None,  1.0),
    # ── PREGNANT (binary_r: any count → RED → flag) ───────────────────────────
    ("Pregnant",                         "Pregnant",              RULE_BINARY_R, None,  None,  None),
]

# Build lookup dict: ai_col → rule entry
RULE_MAP = {r[0]: r for r in COLUMN_RULES}

# Auto-reject gate column names
AUTO_REJECT_COLS = [r[0] for r in COLUMN_RULES if r[2] == RULE_BINARY_R and r[0] != "Pregnant"]
# 8 risk parameter column names (in order)
RISK_PARAM_COLS  = [r[0] for r in COLUMN_RULES if r[2] == RULE_PCT]

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1: COUNT YES VALUES  (always by column name)
# ─────────────────────────────────────────────────────────────────────────────

def count_yes_format_a(df, col_name):
    """
    Format A/C/IG: aggregated sheet.
    count = row 0[col_name], total = row 1[col_name].
    Returns (count, total, pct).
    """
    if col_name not in df.columns:
        return 0.0, 0.0, 0.0
    try:
        count = float(df[col_name].iloc[0]) if pd.notna(df[col_name].iloc[0]) else 0.0
        total = float(df[col_name].iloc[1]) if len(df) > 1 and pd.notna(df[col_name].iloc[1]) else 0.0
        pct   = count / total if total > 0 else 0.0
        return count, total, pct
    except Exception:
        return 0.0, 0.0, 0.0


def count_yes_format_b(df, col_name):
    """
    Format B: individual Yes/No rows.
    Count YES occurrences in named column.
    Returns (count, total, pct).
    """
    if col_name not in df.columns:
        return 0.0, 0.0, 0.0
    series = df[col_name].astype(str).str.strip().str.upper()
    count  = float((series == "YES").sum())
    total  = float(series.isin(["YES", "NO"]).sum())
    pct    = count / total if total > 0 else 0.0
    return count, total, pct


def count_yes_captions(df, col_name, cmap):
    """
    Captions: headers at row 4, counts at rows 0/1.
    Returns (count, total, pct).
    """
    pandas_col = cmap.get(col_name)
    if pandas_col is None:
        return 0.0, 0.0, 0.0
    try:
        count = float(df.iloc[0][pandas_col]) if pd.notna(df.iloc[0][pandas_col]) else 0.0
        total = float(df.iloc[1][pandas_col]) if len(df) > 1 and pd.notna(df.iloc[1][pandas_col]) else 0.0
        pct   = count / total if total > 0 else 0.0
        return count, total, pct
    except Exception:
        return 0.0, 0.0, 0.0


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2: APPLY RULE → GREEN / AMBER / RED  (pure function, no side effects)
# ─────────────────────────────────────────────────────────────────────────────

def apply_rule(count, pct, rule_type, green_max, amber_max):
    """
    Apply the scoring rule to return GREEN / AMBER / RED.

    RULE_BINARY_R: count > 0 → RED,   count = 0 → GREEN
    RULE_BINARY_G: count > 0 → GREEN, count = 0 → RED
    RULE_PCT:      pct < green_max → GREEN
                   pct < amber_max → AMBER  (if amber_max set)
                   else            → RED
    """
    if rule_type == RULE_BINARY_R:
        return "RED" if count > 0 else "GREEN"
    if rule_type == RULE_BINARY_G:
        return "GREEN" if count > 0 else "RED"
    # RULE_PCT
    if pct < green_max:
        return "GREEN"
    if amber_max is not None and pct < amber_max:
        return "AMBER"
    return "RED"


# ─────────────────────────────────────────────────────────────────────────────
# STEP 3: CROSS-CHECK  (for pre-computed formats)
# ─────────────────────────────────────────────────────────────────────────────

def read_precomputed(df, col_name, fmt):
    """
    Read the pre-computed GREEN/AMBER/RED from the verdict row (row 3)
    for Format A / C / IG sheets. Returns None if not present.
    """
    if col_name not in df.columns or len(df) <= 3:
        return None
    val = str(df[col_name].iloc[3]).strip().upper()
    if val in ("GREEN", "AMBER", "RED"):
        return val
    return None


def cross_check(precomp, computed, col_name, sheet_name):
    """
    Compare pre-computed vs count-based result.
    Returns (final_code, mismatch_flag).
    If they match → use it.
    If they disagree → trust computed (from raw counts) and flag it.
    """
    if precomp is None:
        return computed, False
    if precomp == computed:
        return precomp, False
    # Mismatch — flag it, use count-based as ground truth
    return computed, True


# ─────────────────────────────────────────────────────────────────────────────
# FORMAT DETECTION  (by column name)
# ─────────────────────────────────────────────────────────────────────────────

def detect_format(df):
    """
    Auto-detect sheet format by reading specific named columns.
    Returns: 'A', 'B', 'C', 'IG', or 'Captions'.
    """
    # Check response_1 col (Format A, B, C)
    if "response_1" in df.columns and len(df) > 0:
        r0 = str(df["response_1"].iloc[0]).strip()
        if "HEADER" in r0.upper():
            return "A"
        if "no of" in r0.lower() or "occurences" in r0.lower():
            return "C"

    # Check Audio transcription col (Format IG)
    if "Audio transcription" in df.columns and len(df) > 0:
        r0 = str(df["Audio transcription"].iloc[0]).strip()
        if "HEADER" in r0.upper():
            return "IG"

    # Check if individual Yes/No rows (Format B)
    for col_name in ("Profanity Usage", "Kids Presenence", "Breastfeeding"):
        if col_name in df.columns and len(df) > 0:
            sample = df[col_name].astype(str).str.strip().str.upper()
            if sample.isin(["YES", "NO"]).any():
                return "B"

    # Captions: no response_1, but has flag columns
    if "Profanity Usage" in df.columns:
        return "Captions"

    return "B"


def captions_col_map(df):
    """Captions: real column names are in row 4."""
    mapping = {}
    if len(df) > 4:
        for col, val in df.iloc[4].items():
            if pd.notna(val):
                mapping[str(val).strip()] = col
    return mapping


# ─────────────────────────────────────────────────────────────────────────────
# COMPLETE COLUMN SCORER  (combines steps 1-3)
# ─────────────────────────────────────────────────────────────────────────────

def score_column(df, col_name, fmt, cmap=None):
    """
    For a single named column in a sheet, return:
      (count, total, pct, final_code, mismatch_flag)

    Process:
      1. Count YES by column name (appropriate method per format)
      2. Apply rule → GREEN/AMBER/RED
      3. For Format A/C/IG: read pre-computed, cross-check, flag if mismatch
    """
    rule = RULE_MAP.get(col_name)
    if rule is None:
        return 0.0, 0.0, 0.0, "GREEN", False

    _, label, rule_type, green_max, amber_max, _ = rule

    # Step 1: count YES by format
    if fmt == "Captions":
        count, total, pct = count_yes_captions(df, col_name, cmap or {})
    elif fmt == "B":
        count, total, pct = count_yes_format_b(df, col_name)
    else:  # A, C, IG
        count, total, pct = count_yes_format_a(df, col_name)

    # Step 2: apply rule
    computed = apply_rule(count, pct, rule_type, green_max, amber_max)

    # Step 3: cross-check precomputed (Format A/C/IG only)
    if fmt in ("A", "C", "IG"):
        precomp = read_precomputed(df, col_name, fmt)
        final_code, mismatch = cross_check(precomp, computed, col_name, "")
    else:
        final_code, mismatch = computed, False

    return count, total, pct, final_code, mismatch


# ─────────────────────────────────────────────────────────────────────────────
# FULL SHEET SCORER
# ─────────────────────────────────────────────────────────────────────────────

def score_sheet(df, handle, use_case):
    fmt  = detect_format(df)
    cmap = captions_col_map(df) if fmt == "Captions" else None

    mismatches = []  # track any cross-check mismatches

    def sc(col_name):
        count, total, pct, code, mismatch = score_column(df, col_name, fmt, cmap)
        if mismatch:
            mismatches.append(f"{col_name}: precomp≠computed (using count-based)")
        return count, total, pct, code

    # ── AUTO-REJECT GATES ────────────────────────────────────────────────────
    auto_flags = {}
    for col_name in AUTO_REJECT_COLS:
        _, label, *_ = RULE_MAP[col_name]
        count, _, _, code = sc(col_name)
        auto_flags[label] = 1 if code == "RED" else 0

    auto_reject = sum(auto_flags.values()) > 0

    # ── 8 RISK PARAMETERS ────────────────────────────────────────────────────
    risk_params = {}
    total_pts = 0.0
    for col_name in RISK_PARAM_COLS:
        _, label, _, _, _, _ = RULE_MAP[col_name]
        _, _, _, code = sc(col_name)
        pts = 1.0 if code == "RED" else (0.5 if code == "AMBER" else 0.0)
        risk_params[label] = {"code": code, "pts": pts}
        total_pts += pts

    risk_score = 10.0 if auto_reject else round((total_pts / 8) * 10, 4)

    # ── PREGNANT FLAG ─────────────────────────────────────────────────────────
    preg_count, _, _, preg_code = sc("Pregnant")
    preg_flag = 1 if preg_code == "RED" else 0

    # ── RELEVANCE ─────────────────────────────────────────────────────────────
    if use_case == "Pediatric":
        kids_count, _, _, kids_code = sc("Kids Presenence")
        kids_pts = 3.5 if kids_code == "GREEN" else 0.0
        auto_base = kids_pts

        # Kids age summary
        kids_age = _kids_age_summary(df, fmt)

        relevance = {
            "Kids Presence":               {"pts": kids_pts,  "code": kids_code, "auto": True},
            "Topics around kids":          {"pts": None,      "manual": True,    "max_pts": 3.5},
            "Abbott brands":               {"pts": None,      "manual": True,    "max_pts": 1.0},
            "Awards/Media Presence":       {"pts": None,      "manual": True,    "max_pts": 1.0},
            "Relevant Brand Partnerships": {"pts": None,      "manual": True,    "max_pts": 1.0},
        }
        preg_note = "Pregnant — Reject" if preg_flag else ""

    else:  # Adult Nutrition
        h_count, _, _, h_code = sc("Topics around Adult Health")
        n_count, _, _, n_code = sc("Topics on Adult Healthy Nutrition")
        h_pts = 5.0 if h_code == "GREEN" else 0.0
        n_pts = 2.0 if n_code == "GREEN" else 0.0
        auto_base = max(0.0, h_pts + n_pts - (2.0 if preg_flag else 0.0))
        kids_age  = ""

        relevance = {
            "Topics Adult Health":         {"pts": h_pts, "code": h_code, "auto": True},
            "Topics Adult Nutrition":      {"pts": n_pts, "code": n_code, "auto": True},
            "Abbott brands":               {"pts": None,  "manual": True, "max_pts": 1.0},
            "Awards/Media Presence":       {"pts": None,  "manual": True, "max_pts": 1.0},
            "Relevant Brand Partnerships": {"pts": None,  "manual": True, "max_pts": 1.0},
        }
        preg_note = "Pregnant — Penalty -2pts" if preg_flag else ""

    return {
        "handle":        handle,
        "use_case":      use_case,
        "fmt":           fmt,
        "total_videos":  _total_videos(df, fmt, cmap),
        "auto_flags":    auto_flags,
        "auto_reject":   auto_reject,
        "risk_params":   risk_params,
        "risk_score":    risk_score,
        "relevance":     relevance,
        "auto_base":     auto_base,
        "preg_flag":     preg_flag,
        "preg_note":     preg_note,
        "kids_age":      kids_age,
        "mismatches":    mismatches,
    }


def _kids_age_summary(df, fmt):
    """Extract unique Kids Age Group values from videos where kids are present."""
    try:
        age_col, pres_col = "Kids Age Group", "Kids Presenence"
        if age_col not in df.columns or pres_col not in df.columns:
            return ""
        data = df.iloc[4:] if fmt in ("A", "C", "IG") else df
        mask = data[pres_col].astype(str).str.strip().str.upper() == "YES"
        ages = data.loc[mask, age_col].dropna().astype(str).str.strip()
        skip = {"nan", "none", "none of these", ""}
        unique = sorted({a for a in ages if a.lower() not in skip})
        return ", ".join(unique)
    except Exception:
        return ""


def _total_videos(df, fmt, cmap):
    """Get total video count for display."""
    if fmt == "Captions":
        if cmap:
            col = cmap.get("Profanity Usage")
            if col and len(df) > 1:
                try: return int(float(df.iloc[1][col]))
                except: pass
        return 0
    if fmt in ("A", "C", "IG"):
        for col in ("Profanity Usage", "Alcohol Use Discussion", "Kids Presenence",
                    "Topics around Adult Health"):
            if col in df.columns:
                try:
                    v = float(df[col].iloc[1])
                    if v > 0: return int(v)
                except: pass
        return 0
    else:  # Format B
        for col in ("Profanity Usage", "Kids Presenence", "Breastfeeding"):
            if col in df.columns:
                n = df[col].astype(str).str.strip().str.upper().isin(["YES","NO"]).sum()
                if n > 0: return int(n)
        return len(df)


# ─────────────────────────────────────────────────────────────────────────────
# USE CASE AUTO-DETECTION  (by column name + count)
# ─────────────────────────────────────────────────────────────────────────────

def detect_use_case(xl):
    """
    Adult Nutrition if any TT/FB/IG sheet has Topics around Adult Health count > 0.
    Pediatric otherwise.
    """
    for sheet_name, df in xl.items():
        if not any(sheet_name.startswith(p) for p in ("TT ", "FB ", "IG ", "IG + YT ")):
            continue
        if "Topics around Adult Health" not in df.columns:
            continue
        try:
            val = df["Topics around Adult Health"].iloc[0]
            if str(val).strip().lower() not in ("nan", "none", "no", ""):
                if float(val) > 0:
                    return "Adult"
        except: pass
    return "Pediatric"


# ─────────────────────────────────────────────────────────────────────────────
# SHEET DETECTION & PROFILES
# ─────────────────────────────────────────────────────────────────────────────

PLATFORM_PREFIXES = {
    "TT ": "TT", "FB ": "FB",
    "IG + YT ": "IG_YT", "IG ": "IG", "Captions ": "Captions",
}
SKIP_SHEETS = {
    "master grading sheet", "tt grading sheet", "ig + yt grading sheet",
    "fb grading sheet", "captions grading sheet", "grading", "profiles",
    "unknown",
}

def detect_sheets(xl):
    """Detect influencer sheets in the main AI file (sheets have platform prefixes)."""
    grouped = {p: [] for p in PLATFORM_PREFIXES.values()}
    for sheet in xl:
        if sheet.lower().strip() in SKIP_SHEETS:
            continue
        for prefix, platform in PLATFORM_PREFIXES.items():
            if sheet.startswith(prefix):
                grouped[platform].append((sheet, sheet[len(prefix):].strip()))
                break
    return grouped


def detect_captions_sheets(xl_cap):
    """
    Detect influencer sheets in a Captions-only file.
    These files have one sheet per influencer named directly by handle
    (no platform prefix). Skip 'Grading' and 'Unknown' sheets.
    Returns list of (sheet_name, handle) tuples.
    """
    sheets = []
    for sheet in xl_cap:
        if sheet.lower().strip() in SKIP_SHEETS:
            continue
        sheets.append((sheet, sheet.strip()))
    return sheets


def build_profiles(xl):
    for name, df in xl.items():
        if name.lower() != "profiles":
            continue
        mapping = {}
        for _, row in df.iterrows():
            handle, full_name = None, None
            for col in df.columns:
                val = str(row[col]).strip() if pd.notna(row[col]) else ""
                if "username" in str(col).lower() and val:
                    handle = val
                if "influencer name" in str(col).lower() and val:
                    full_name = val
            if handle:
                mapping[handle] = full_name or handle
        return mapping
    return {}


SKIP_HANDLE = {
    "nan", "none", "header", "total videos processed", "% prevelance", "% prevalence",
    "% of occurence", "code/point", "verdict", "no of occurences",
    "platform_username", "username", "",
}

def extract_handle(df):
    """
    Extract influencer handle from sheet. Reads by column name, skips header strings.
    Returns None if no clean handle found (caller uses raw sheet-name handle as fallback).
    """
    for col in ("platform_username", "Username"):
        if col not in df.columns:
            continue
        for val in df[col].dropna().astype(str).str.strip():
            if val.lower() not in SKIP_HANDLE and not val.startswith("```"):
                return val
    return None  # Caller will use sheet-name handle (correct for FB Format C sheets)


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL STYLES
# ─────────────────────────────────────────────────────────────────────────────

FN = "Arial"
BG = {
    "navy":    "1F3864", "dkred":   "C00000", "brtred":  "FF0000",
    "dkgrn":   "375623", "blue":    "2E75B6", "manual":  "FFF2CC",
    "green":   "C6EFCE", "amber":   "FFEB9C", "red":     "FFC7CE",
    "grey":    "F2F2F2", "name":    "D9E1F2", "white":   "FFFFFF",
}

def fp(k): return PatternFill("solid", fgColor=BG.get(k, k))
def bdr():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)
def hf(color="FFFFFF", bold=True, sz=9):  return Font(name=FN, bold=bold, color=color, size=sz)
def df_(color="000000", bold=False, sz=9): return Font(name=FN, bold=bold, color=color, size=sz)
def ac(wrap=True): return Alignment(horizontal="center", vertical="center", wrap_text=wrap)
def al():          return Alignment(horizontal="left",   vertical="center", wrap_text=True)


# ─────────────────────────────────────────────────────────────────────────────
# GRADING SHEET WRITER
# ─────────────────────────────────────────────────────────────────────────────

def write_grading_sheet(ws, scores, label, use_case, profiles):
    ws.row_dimensions[1].height = 22
    t = ws.cell(1, 1, f"Abbott Influencer Vetting — {label} Grading Sheet ({use_case})")
    t.font = Font(name=FN, bold=True, size=12, color=BG["navy"]); t.alignment = al()

    # Relevance columns vary by use case
    if use_case == "Pediatric":
        rel_cols = [
            ("Kids Presence\n(3.5 pts — AI auto)",              "dkgrn",  16, False),
            ("Topics around kids\n(3.5 pts — ⬛ Manual)",        "manual", 18, True),
            ("Abbott brands\n(1 pt — ⬛ Manual)",                "manual", 16, True),
            ("Awards/Media Presence\n(1 pt — ⬛ Manual)",        "manual", 18, True),
            ("Relevant Brand Partnerships\n(1 pt — ⬛ Manual)",  "manual", 20, True),
        ]
        notes_hdr = "Kids Age Groups (from AI)"
    else:
        rel_cols = [
            ("Topics Adult Health\n(5 pts — AI auto)",          "dkgrn",  18, False),
            ("Topics Adult Nutrition\n(2 pts — AI auto)",       "dkgrn",  18, False),
            ("Abbott brands\n(1 pt — ⬛ Manual)",                "manual", 16, True),
            ("Awards/Media Presence\n(1 pt — ⬛ Manual)",        "manual", 18, True),
            ("Relevant Brand Partnerships\n(1 pt — ⬛ Manual)",  "manual", 20, True),
        ]
        notes_hdr = "Notes"

    COLS = [
        ("Influencer Name",                                "navy",   20, False),
        ("Handle / Username",                              "navy",   18, False),
        ("Substance Use\n(Auto-reject — AI)",              "brtred", 14, False),
        ("Worked With Competition\n(⬛ Manual—last 90d)",   "manual", 20, True),
        ("Anti-breastfeeding\n(Auto-reject — AI)",         "brtred", 16, False),
        ("Anti-vaccination\n(Auto-reject — AI)",           "brtred", 15, False),
        ("Anti-healthcare\n(Auto-reject — AI)",            "brtred", 15, False),
        ("Auto-reject Check\n(>0 → Risk Score = 10)",      "dkred",  16, False),
        ("Profanity",                                      "blue",   12, False),
        ("Alcohol",                                        "blue",   12, False),
        ("Sensitive Content",                              "blue",   14, False),
        ("Stereotype & Bias",                              "blue",   14, False),
        ("Violence Content",                               "blue",   13, False),
        ("Political",                                      "blue",   12, False),
        ("Unscientific",                                   "blue",   13, False),
        ("Ultra-processed Food",                           "blue",   14, False),
        ("RISK Score\n(0=Clean → 10=Reject)",              "navy",   14, False),
    ] + rel_cols + [
        ("Total Relevance Score\n(max 10)",                "dkgrn",  15, False),
        ("Pregnant Or Not",                                "dkgrn",  13, False),
        ("Notes / Flags\n(⬛ Manual)",                      "manual", 24, True),
        (notes_hdr,                                        "595959", 28, False),
    ]

    for ci, (_, _, w, _) in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ROW_H, ROW_D = 2, 3
    ws.row_dimensions[ROW_H].height = 52
    for ci, (lbl, bg, _, is_m) in enumerate(COLS, 1):
        c = ws.cell(ROW_H, ci, lbl)
        c.fill = fp(bg); c.border = bdr(); c.alignment = ac()
        c.font = hf(color="7B4F00" if is_m else "FFFFFF")

    # Column index constants
    CI_NAME=1; CI_HAND=2; CI_SUB=3; CI_COMP=4; CI_BF=5; CI_VAX=6; CI_HC=7
    CI_AR=8; CI_PROF=9; CI_ALC=10; CI_SEN=11; CI_STE=12; CI_VIO=13
    CI_POL=14; CI_UNS=15; CI_ULP=16; CI_RISK=17
    CI_R1=18; CI_R2=19; CI_R3=20; CI_R4=21; CI_R5=22
    CI_TOT=23; CI_PREG=24; CI_N1=25; CI_N2=26

    RISK_LABELS = ["Profanity","Alcohol","Sensitive content","Stereotype & Bias",
                   "Violence Content","Political","Unscientific","Ultra-processed Food"]

    for off, s in enumerate(scores):
        row = ROW_D + off; rbg = "white" if off % 2 == 0 else "grey"
        ws.row_dimensions[row].height = 16

        def wc(col, val, bg=None, bold=False, left=False):
            c = ws.cell(row, col, val)
            c.fill = fp(bg or rbg); c.font = df_(bold=bold)
            c.border = bdr(); c.alignment = al() if left else ac()
            return c

        def manual(col, hint="↑ fill manually"):
            c = ws.cell(row, col, hint)
            c.fill = fp("manual")
            c.font = Font(name=FN, italic=True, color="7B4F00", size=8)
            c.border = bdr(); c.alignment = ac()

        # Name and handle
        full = profiles.get(s["handle"], s["handle"])
        wc(CI_NAME, full, bg="name", bold=True, left=True)
        wc(CI_HAND, s["handle"], bg="name", left=True)

        # Auto-reject gates (AI-computed)
        for ci, gate in [(CI_SUB,"Substance Use"),(CI_BF,"Anti-breastfeeding"),
                          (CI_VAX,"Anti-vaccination"),(CI_HC,"Anti-healthcare")]:
            f = s["auto_flags"].get(gate, 0)
            wc(ci, f, bg="red" if f else "green")

        # Competition — always manual
        manual(CI_COMP)

        # Auto-reject sum
        ar = sum(s["auto_flags"].values())
        if ar > 0:
            c = ws.cell(row, CI_AR, ar)
            c.fill=fp("brtred"); c.border=bdr(); c.alignment=ac()
            c.font=Font(name=FN, bold=True, color="FFFFFF", size=9)
        else:
            wc(CI_AR, 0, bg="green")

        # 8 risk params
        for ci_off, lbl in enumerate(RISK_LABELS):
            ci = CI_PROF + ci_off
            info = s["risk_params"].get(lbl, {"code":"GREEN","pts":0})
            bg = "red" if info["code"]=="RED" else ("amber" if info["code"]=="AMBER" else "green")
            wc(ci, info["pts"], bg=bg)

        # Risk score
        rs = s["risk_score"]
        if rs >= 5:
            c = ws.cell(row, CI_RISK, rs)
            c.fill=fp("brtred"); c.border=bdr(); c.alignment=ac()
            c.font=Font(name=FN, bold=True, color="FFFFFF", size=9)
        elif rs > 0:
            wc(CI_RISK, rs, bg="amber", bold=True)
        else:
            wc(CI_RISK, rs, bg="green", bold=True)

        # Relevance params
        for i, rk in enumerate(list(s["relevance"].keys())):
            ci = CI_R1 + i
            info = s["relevance"][rk]
            if info.get("manual"):
                manual(ci)
            else:
                pts  = info["pts"] or 0
                bg   = "green" if pts > 0 else rbg
                wc(ci, pts, bg=bg)

        # Total relevance — live SUM formula
        rl = [get_column_letter(CI_R1+i) for i in range(5)]
        pr = get_column_letter(CI_PREG)
        if use_case == "Pediatric":
            formula = f"=MIN(10,MAX(0,{'+'.join(f'IFERROR({l}{row},0)' for l in rl)}))"
        else:
            formula = f"=MIN(10,MAX(0,{'+'.join(f'IFERROR({l}{row},0)' for l in rl)}-IF({pr}{row}=1,2,0)))"

        rb_base = s["auto_base"]
        rbg_rel = "green" if rb_base >= 7 else ("amber" if rb_base >= 4 else rbg)
        c = ws.cell(row, CI_TOT, formula)
        c.fill=fp(rbg_rel); c.font=df_(bold=True); c.border=bdr(); c.alignment=ac()

        # Pregnant
        pf = s["preg_flag"]
        if pf:
            c = wc(CI_PREG, "⚠ Pregnant", bg="red", bold=True)
            c.font = Font(name=FN, bold=True, color="C00000", size=8)
        else:
            wc(CI_PREG, 0, bg=rbg)

        # Notes
        manual(CI_N1, "")
        # Flag mismatches in notes column
        note = "; ".join(s.get("mismatches", [])) if s.get("mismatches") else ""
        wc(CI_N2, note or s.get("kids_age",""), bg=rbg, left=True)

    # Legend
    lg = ROW_D + len(scores) + 2
    ws.cell(lg, 1, "LEGEND").font = Font(name=FN, bold=True, size=9)
    for i,(clr,txt) in enumerate([
        ("manual", "⬛ Soft Yellow = Manual entry required — fill before client submission"),
        ("green",  "🟢 Green  = Pass / Clean (auto-scored from AI data)"),
        ("amber",  "🟠 Amber  = Review (auto-scored)"),
        ("red",    "🔴 Pink   = Flagged / Reject (auto-scored)"),
        ("brtred", "🔴 Bright Red = Auto-reject triggered or Risk Score ≥ 5"),
    ]):
        ws.cell(lg+1+i,1,"").fill=fp(clr); ws.cell(lg+1+i,1,"").border=bdr()
        ws.cell(lg+1+i,2,txt).font=Font(name=FN,size=8)

    ws.freeze_panes = f"C{ROW_D}"


# ─────────────────────────────────────────────────────────────────────────────
# MASTER SHEET
# ─────────────────────────────────────────────────────────────────────────────

def write_master_sheet(ws, all_scores, use_case, profiles):
    ws.row_dimensions[1].height = 22
    ws.cell(1,1,f"Abbott Influencer Vetting — Master Summary ({use_case})").font = \
        Font(name=FN,bold=True,size=12,color=BG["navy"])

    COLS = [("Influencer Name","navy",22),("Handle","navy",18),
            ("Risk MAX\n(all platforms)","dkred",16),("Relevance AVG\n(auto-base)","dkgrn",18),
            ("Risk Zone","navy",16),("Relevance Zone","navy",16),
            ("Recommendation","navy",20),("Platforms Scored","navy",22)]
    for i,(_,_,w) in enumerate(COLS,1): ws.column_dimensions[get_column_letter(i)].width=w

    ROW_H,ROW_D=3,4; ws.row_dimensions[ROW_H].height=40
    for ci,(lbl,bg,_) in enumerate(COLS,1):
        c=ws.cell(ROW_H,ci,lbl); c.fill=fp(bg); c.font=hf(); c.border=bdr(); c.alignment=ac()

    by_h={}
    for plat,sl in all_scores.items():
        for s in sl:
            h=s["handle"]
            if h not in by_h: by_h[h]={"risks":[],"rels":[],"plats":[]}
            by_h[h]["risks"].append(s["risk_score"])
            by_h[h]["rels"].append(s["auto_base"])
            by_h[h]["plats"].append(plat)

    for off,(handle,data) in enumerate(sorted(by_h.items())):
        row=ROW_D+off; rbg="white" if off%2==0 else "grey"
        ws.row_dimensions[row].height=16
        mr=max(data["risks"]); al2=round(sum(data["rels"])/len(data["rels"]),2)
        plats=", ".join(sorted(set(data["plats"])))
        r_bg="red" if mr>=5 else ("amber" if mr>0 else "green")
        l_bg="green" if al2>=8 else ("amber" if al2>=4 else "red")
        rec="❌ Reject" if mr>=5 else ("✅ Approve" if mr==0 and al2>=7 else "🟡 Manual Review")
        rc="red" if "Reject" in rec else ("green" if "Approve" in rec else "amber")
        def mc(col,val,bg,bold=False,left=False):
            c=ws.cell(row,col,val); c.fill=fp(bg); c.font=df_(bold=bold)
            c.border=bdr(); c.alignment=al() if left else ac()
        mc(1,profiles.get(handle,handle),"name",bold=True,left=True)
        mc(2,handle,"name",left=True); mc(3,mr,r_bg,bold=True); mc(4,al2,l_bg,bold=True)
        mc(5,"🔴 Reject" if mr>=5 else "🟠 Review" if mr>0 else "🟢 Clean",r_bg)
        mc(6,"🟢 High" if al2>=8 else "🟠 Moderate" if al2>=4 else "🔴 Low",l_bg)
        mc(7,rec,rc,bold=True); mc(8,plats,rbg)

    nr=ROW_D+len(by_h)+2
    ws.cell(nr,1,"NOTES").font=Font(name=FN,bold=True,size=9)
    for i,n in enumerate([
        "Risk = MAX across platforms.  Relevance = AVERAGE across platforms (auto-base only).",
        "⬛ Yellow cells in grading sheets = manual input required before client submission.",
        "Relevance auto-base excludes manual fields (Abbott brands, Awards, Partnerships).",
    ]): ws.cell(nr+1+i,1,n).font=Font(name=FN,size=8,italic=True,color="595959")
    ws.freeze_panes=f"C{ROW_D}"


# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT BUILDER
# ─────────────────────────────────────────────────────────────────────────────

PLAT_LABELS={"TT":"TikTok","FB":"Facebook","IG_YT":"IG + YT","IG":"Instagram","Captions":"Captions"}

def build_excel(all_scores, use_case, profiles, captions_scores=None):
    """
    all_scores:      dict platform → list of score dicts  (main AI file)
    captions_scores: list of score dicts (Captions file, optional)
    """
    wb = openpyxl.Workbook()

    # Merge captions scores into all_scores for master sheet
    all_for_master = dict(all_scores)
    if captions_scores:
        all_for_master["Captions"] = (all_for_master.get("Captions", []) + captions_scores)

    ws_m = wb.active; ws_m.title = "Master Grading Sheet"
    write_master_sheet(ws_m, all_for_master, use_case, profiles)

    # Per-platform grading sheets from main file
    for plat in ["TT", "IG_YT", "IG", "FB", "Captions"]:
        sl = all_scores.get(plat, [])
        if not sl: continue
        ws = wb.create_sheet(title=f"{PLAT_LABELS[plat]} Grading Sheet")
        write_grading_sheet(ws, sl, PLAT_LABELS[plat], use_case, profiles)

    # Captions file grading sheet (separate tab)
    if captions_scores:
        ws_cap = wb.create_sheet(title="Captions File Grading Sheet")
        write_grading_sheet(ws_cap, captions_scores, "Captions (Caption File)", use_case, profiles)

    buf = BytesIO(); wb.save(buf); buf.seek(0); return buf


# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Abbott Influencer Vetting", page_icon="🏥", layout="wide")
st.markdown("""<style>
.stApp{font-family:Arial,sans-serif;}h1{color:#1F3864;}h2{color:#2E75B6;}
.g{color:#375623;font-weight:bold;}.a{color:#806000;font-weight:bold;}.r{color:#C00000;font-weight:bold;}
</style>""", unsafe_allow_html=True)

st.title("🏥 Abbott Influencer Vetting — Auto-Grader v5.0")
st.caption("Framework V4 | Auto-detects Pediatric vs Adult | All columns read by name | Cross-checks pre-computed values")
st.divider()

# ── File uploaders ─────────────────────────────────────────────────────────
col_a, col_b = st.columns(2)

with col_a:
    st.subheader("📂 Main AI File")
    st.caption("TT / FB / IG sheets with platform prefix (e.g. TT woanying_24)")
    uploaded_main = st.file_uploader(
        "Drop your ConvoTrack AI export here",
        type=["xlsx"],
        key="main_file",
        help="Sheet names start with platform prefix: TT, FB, IG, Captions etc."
    )

with col_b:
    st.subheader("📎 Captions File  _(optional)_")
    st.caption("One sheet per influencer, named by handle (e.g. woanying24)")
    uploaded_cap = st.file_uploader(
        "Drop your Captions file here (e.g. Abbott SG Similac_All Captions)",
        type=["xlsx"],
        key="captions_file",
        help="Sheet names are handle names directly, no platform prefix. Grading and Unknown sheets are skipped."
    )

st.divider()

# ── Only proceed if at least the main file is uploaded ────────────────────
if not uploaded_main:
    st.info("👆 Upload the Main AI file to begin. The Captions file is optional.")
    with st.expander("ℹ️ Scoring methodology"):
        st.markdown("""
**Step 1 — Per influencer sheet, read each column BY NAME and count YES:**
- Format A/C/IG (aggregated): YES count already in row 0, total in row 1
- Format B (Yes/No rows): count "Yes" values across all video rows by column name
- Cross-check: if pre-computed Code/Point disagrees with raw count → use count-based result, flag in notes

**Step 2 — Apply verified rules:**
| Column group | Rule | Details |
|---|---|---|
| 8 Risk params | % threshold | See thresholds below |
| Auto-reject gates (Substance/BF/Vax/HC) | count > 0 = RED | Any occurrence = auto-reject |
| Kids Presence / Topics / Media | **count > 0 = GREEN** | Any occurrence = full points |
| Pregnant | count > 0 = RED | Any occurrence = flag |

**Risk thresholds:**
Profanity <15%→🟢, 15-30%→🟠, >30%→🔴 | Alcohol <5%→🟢, 5-15%→🟠, >15%→🔴 |
Sensitive <20%→🟢, ≥20%→🔴 | Stereotype <5%→🟢, 5-15%→🟠, >15%→🔴 |
Violence 0%→🟢, <5%→🟠, ≥5%→🔴 | Political <10%→🟢, 10-25%→🟠, >25%→🔴 |
Unscientific <15%→🟢, ≥15%→🔴 | Ultra-processed <15%→🟢, 15-30%→🟠, >30%→🔴

**⬛ Yellow cells (manual):** Competition (last 90d), Abbott brands, Awards, Brand Partnerships, Topics around kids (Pediatric only)
        """)

else:
    # ── Read main file ─────────────────────────────────────────────────────
    with st.spinner("Reading main AI file…"):
        xl      = pd.read_excel(uploaded_main, sheet_name=None)
        grp     = detect_sheets(xl)
        prf     = build_profiles(xl)
        uc      = detect_use_case(xl)

    total_main = sum(len(v) for v in grp.values())
    st.success(f"✅ Main file loaded | **{uc}** use case detected | {total_main} influencer sheets")

    cols = st.columns(5)
    for cw, (p, lb) in zip(cols, [("TT","TikTok"),("FB","Facebook"),
                                    ("IG_YT","IG+YT"),("IG","Instagram"),("Captions","Captions")]):
        cw.metric(lb, len(grp[p]))

    # ── Read captions file (if uploaded) ──────────────────────────────────
    cap_sheets = []
    if uploaded_cap:
        with st.spinner("Reading captions file…"):
            xl_cap     = pd.read_excel(uploaded_cap, sheet_name=None)
            cap_sheets = detect_captions_sheets(xl_cap)
            # Use same use_case as main file
        st.success(f"✅ Captions file loaded | {len(cap_sheets)} influencer caption sheets found")

    st.divider()

    # ── Score main file ────────────────────────────────────────────────────
    with st.spinner("Scoring all influencer sheets…"):
        all_sc, warns = {}, []
        for plat, sl in grp.items():
            if not sl: continue
            psc = []
            for sn, h in sl:
                df = xl[sn]
                try:
                    s = score_sheet(df, h, uc)
                    actual = extract_handle(df)
                    if actual: s["handle"] = actual
                    psc.append(s)
                except Exception as e:
                    warns.append(f"{sn}: {e}")
            all_sc[plat] = psc

    # ── Score captions file ────────────────────────────────────────────────
    cap_scores = []
    if cap_sheets:
        with st.spinner(f"Scoring {len(cap_sheets)} captions sheets…"):
            for sn, h in cap_sheets:
                df = xl_cap[sn]
                try:
                    s = score_sheet(df, h, uc)
                    # Use sheet name as handle (captions file has no platform_username col)
                    s["handle"] = h
                    s["fmt"] = s["fmt"] + " (Captions file)"
                    cap_scores.append(s)
                except Exception as e:
                    warns.append(f"Captions/{sn}: {e}")
        st.success(f"✅ {len(cap_scores)} captions sheets scored")

    if warns:
        with st.expander(f"⚠ {len(warns)} sheets had errors"):
            for w in warns: st.text(w)

    # ── Preview ────────────────────────────────────────────────────────────
    st.subheader("📊 Master Summary Preview")

    # Combine all scores for preview
    all_combined = dict(all_sc)
    if cap_scores:
        all_combined["Captions"] = all_combined.get("Captions", []) + cap_scores

    rows = []
    by_h = {}
    for p, sl in all_combined.items():
        for s in sl:
            h = s["handle"]
            if h not in by_h: by_h[h] = {"risks": [], "rels": [], "plats": []}
            by_h[h]["risks"].append(s["risk_score"])
            by_h[h]["rels"].append(s["auto_base"])
            by_h[h]["plats"].append(p)

    for h, d in sorted(by_h.items()):
        mr  = max(d["risks"])
        al2 = round(sum(d["rels"]) / len(d["rels"]), 2)
        rec = "❌ Reject" if mr >= 5 else ("✅ Approve" if mr == 0 and al2 >= 7 else "🟡 Review")
        rows.append({"Name": prf.get(h, h), "Handle": h, "Risk MAX": mr,
                     "Rel AVG (auto)": al2,
                     "Platforms": ", ".join(sorted(set(d["plats"]))), "Rec": rec})

    dfp = pd.DataFrame(rows)
    def sty(row):
        s = [""] * len(row)
        ri  = dfp.columns.get_loc("Risk MAX")
        li  = dfp.columns.get_loc("Rel AVG (auto)")
        rci = dfp.columns.get_loc("Rec")
        s[ri]  = f"background-color:{'#FFC7CE' if row['Risk MAX']>=5 else '#FFEB9C' if row['Risk MAX']>0 else '#C6EFCE'}"
        s[li]  = f"background-color:{'#C6EFCE' if row['Rel AVG (auto)']>=7 else '#FFEB9C' if row['Rel AVG (auto)']>=4 else '#FFC7CE'}"
        s[rci] = f"background-color:{'#C6EFCE' if 'Approve' in row['Rec'] else '#FFC7CE' if 'Reject' in row['Rec'] else '#FFEB9C'};font-weight:bold"
        return s

    st.dataframe(dfp.style.apply(sty, axis=1), use_container_width=True,
                 height=min(500, 50 + len(rows) * 36))

    appr = sum(1 for r in rows if "Approve" in r["Rec"])
    rev  = sum(1 for r in rows if "Review"  in r["Rec"])
    rej  = sum(1 for r in rows if "Reject"  in r["Rec"])
    st.markdown(
        f"**{len(rows)} influencers** &nbsp;|&nbsp; "
        f"<span class='g'>✅ Approve: {appr}</span> &nbsp;|&nbsp; "
        f"<span class='a'>🟡 Review: {rev}</span> &nbsp;|&nbsp; "
        f"<span class='r'>❌ Reject: {rej}</span>",
        unsafe_allow_html=True
    )

    st.divider()
    st.subheader("📥 Download Graded Excel")
    with st.spinner("Building formatted Excel…"):
        buf = build_excel(all_sc, uc, prf, captions_scores=cap_scores if cap_scores else None)

    fname = uploaded_main.name.replace(".xlsx", "") + f"_GRADED_{uc}.xlsx"
    st.download_button(
        "⬇️ Download Graded Excel", data=buf, file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary"
    )

    # Output sheet list
    sheet_list = ["Master Grading Sheet"]
    for plat in ["TT","IG_YT","IG","FB","Captions"]:
        if all_sc.get(plat): sheet_list.append(f"{PLAT_LABELS[plat]} Grading Sheet")
    if cap_scores: sheet_list.append("Captions File Grading Sheet")

    st.caption(
        f"Output contains: {', '.join(sheet_list)}  |  "
        "⬛ Yellow cells = manual input: Competition check, Abbott brands, Awards, Relevant partnerships"
        + (", Topics around kids." if uc == "Pediatric" else ".")
    )
