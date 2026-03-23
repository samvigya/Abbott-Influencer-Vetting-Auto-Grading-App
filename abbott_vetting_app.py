"""
Abbott Influencer Vetting — Auto-Grading App  v2.0
===================================================
Upload the AI-processed Excel file → get fully scored grading sheets.

Supports TWO use cases (auto-detected, manual override available):
  • Pediatric Nutrition  (Similac / PediaSure) — Kids Presence scoring
  • Adult Nutrition      (Ensure / Glucerna)   — Adult Health topics scoring

Scoring logic: Abbott Influencer Vetting Framework V4 (26 Feb 2026)
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO

# ─────────────────────────────────────────────────────────────────────────────
# RISK SCORING CONFIG  (same for both use cases)
# ─────────────────────────────────────────────────────────────────────────────
# (column_in_AI_file, label_in_grading, yellow_thresh, red_thresh)
RISK_PARAMS = [
    ("Profanity Usage",         "Profanity",         0.15,   0.30),
    ("Alcohol Use Discussion",  "Alcohol",           0.05,   0.15),
    ("Sensitive Visual Content","Sensitive content", None,   0.20),
    ("Stereotypes or Bias",     "Stereotype & Bias", 0.05,   0.15),
    ("Violence Advocacy",       "Violence Content",  0.0001, 0.05),
    ("Political Stance",        "Political",         0.10,   0.25),
    ("Unscientific Claims",     "Unscientific",      None,   0.15),
    ("Ultra-processed food",    "Ultra-processed Food", 0.15, 0.30),
]

# Auto-reject gates: (AI_column, grading_label)
AUTO_REJECT_GATES = [
    ("Substance Use Discussion", "Substance Use"),
    ("Breastfeeding",            "Anti-breastfeeding"),
    ("Vaccination",              "Anti-vaccination"),
    ("Health care stance",       "Anti-healthcare"),
]

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def risk_points(pct, y_thresh, r_thresh):
    if pct is None or (isinstance(pct, float) and pd.isna(pct)):
        return 0.0
    pct = float(pct)
    if pct >= r_thresh:
        return 1.0
    if y_thresh is not None and pct >= y_thresh:
        return 0.5
    return 0.0

def safe_val(df, row_idx, col):
    try:
        if col in df.columns:
            v = df.iloc[row_idx][col]
            return float(v) if not pd.isna(v) else 0.0
    except Exception:
        pass
    return 0.0

def get_count(df, col):   return safe_val(df, 0, col)
def get_total(df, col):   return safe_val(df, 1, col)

def safe_pct(count, total):
    try:
        c, t = float(count), float(total)
        return c / t if t > 0 else 0.0
    except Exception:
        return 0.0

def captions_col_map(df):
    """Captions sheets: real headers are in row index 4."""
    mapping = {}
    if len(df) > 4:
        for col, val in df.iloc[4].items():
            if pd.notna(val):
                mapping[str(val).strip()] = col
    return mapping

def captions_get(df, field, col_map, row_idx):
    col = col_map.get(field)
    if col is None:
        return 0.0
    try:
        v = df.iloc[row_idx][col]
        return float(v) if not pd.isna(v) else 0.0
    except Exception:
        return 0.0

def extract_kids_age_summary(df):
    """Collect unique kids age groups from raw data rows."""
    try:
        col = "Kids Age Group"
        if col not in df.columns:
            return ""
        skip = {"HEADER", "Total Videos Processed", "% Prevelance", "nan", "None", ""}
        ages = df[col].dropna().astype(str).tolist()
        unique = sorted({a.strip() for a in ages if a.strip() not in skip})
        return ", ".join(unique) if unique else ""
    except Exception:
        return ""

def detect_use_case(xl_sheets):
    """
    Auto-detect use case.
    Primary signal: if any content sheet has 'Topics around Adult Health' with
    a numeric count > 0 → Adult Nutrition.
    Otherwise → Pediatric.

    Rationale: Both use cases may have Kids Presenence column, but only Adult
    Nutrition sheets have actual non-zero counts in the Topics around Adult Health
    column. Pediatric sheets have NaN / string 'No' in that column.
    """
    for sheet_name, df in xl_sheets.items():
        if not (sheet_name.startswith("TT ") or sheet_name.startswith("FB ") or
                sheet_name.startswith("IG ")):
            continue
        if "Topics around Adult Health" in df.columns:
            try:
                val = df.iloc[0]["Topics around Adult Health"]
                if val is not None and str(val) not in ("nan", "None", "No", ""):
                    if float(val) > 0:
                        return "Adult"
            except Exception:
                pass
    return "Pediatric"

def build_profiles_map(xl_sheets):
    """
    If a Profiles sheet exists, build handle → full_name mapping.
    """
    mapping = {}
    for sheet_name, df in xl_sheets.items():
        if sheet_name.lower() == "profiles":
            for _, row in df.iterrows():
                # look for Username and Influencer Name columns
                handle = None
                name = None
                for col in df.columns:
                    col_l = str(col).lower()
                    val = str(row[col]).strip() if pd.notna(row[col]) else ""
                    if "username" in col_l and val:
                        handle = val
                    if "influencer name" in col_l and val:
                        name = val
                if handle:
                    mapping[handle] = name or handle
    return mapping

# ─────────────────────────────────────────────────────────────────────────────
# SHEET DETECTION
# ─────────────────────────────────────────────────────────────────────────────

PLATFORM_PREFIXES = {
    "TT ":       "TT",
    "FB ":       "FB",
    "IG + YT ":  "IG_YT",
    "IG ":       "IG",
    "Captions ": "Captions",
}

SKIP_SHEETS = {
    "master grading sheet", "tt grading sheet", "ig + yt grading sheet",
    "fb grading sheet", "captions grading sheet", "grading", "profiles",
}

def detect_influencer_sheets(xl_sheets):
    grouped = {"TT": [], "FB": [], "IG_YT": [], "IG": [], "Captions": []}
    for sheet in xl_sheets:
        if sheet.lower().strip() in SKIP_SHEETS:
            continue
        for prefix, platform in PLATFORM_PREFIXES.items():
            if sheet.startswith(prefix):
                handle = sheet[len(prefix):].strip()
                grouped[platform].append((sheet, handle))
                break
    return grouped

def extract_username(df):
    for col in ("platform_username", "Username"):
        if col in df.columns:
            vals = df[col].dropna().astype(str)
            vals = vals[~vals.isin(["nan", "HEADER", "Total Videos Processed", "% Prevelance"])]
            if not vals.empty:
                return vals.iloc[0]
    return None

# ─────────────────────────────────────────────────────────────────────────────
# SCORING ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def score_risk(df, is_captions=False, col_map=None):
    """
    Returns:
      auto_reject_flags: dict label → count (from AI)
      auto_reject_sum:   sum of all flags (>0 means manual review needed)
      risk_param_scores: dict label → 0/0.5/1.0
      risk_score:        final 0-10 risk score
    """
    auto_reject_flags = {}
    auto_reject_sum = 0

    for ai_col, label in AUTO_REJECT_GATES:
        if is_captions:
            count = captions_get(df, ai_col, col_map, 0)
            total = captions_get(df, ai_col, col_map, 1)
        else:
            count = get_count(df, ai_col)
            total = get_total(df, ai_col)
        flag = 1 if count > 0 else 0
        auto_reject_flags[label] = flag
        auto_reject_sum += flag

    auto_reject_triggered = auto_reject_sum > 0

    risk_param_scores = {}
    total_pts = 0.0
    for ai_col, label, y_thresh, r_thresh in RISK_PARAMS:
        if is_captions:
            count = captions_get(df, ai_col, col_map, 0)
            total = captions_get(df, ai_col, col_map, 1)
        else:
            count = get_count(df, ai_col)
            total = get_total(df, ai_col)
        pct = safe_pct(count, total)
        pts = risk_points(pct, y_thresh, r_thresh)
        risk_param_scores[label] = pts
        total_pts += pts

    risk_score = 10.0 if auto_reject_triggered else round((total_pts / 8) * 10, 4)

    return auto_reject_flags, auto_reject_sum, risk_param_scores, risk_score


def score_relevance_pediatric(df, is_captions=False, col_map=None):
    """
    Pediatric relevance:
      Kids Presence  (AI)    → 3.5 pts
      Topics around kids     → MANUAL (None)
      Abbott brands          → MANUAL (None)
      Awards/Media           → MANUAL (None)
      Brand Partnerships     → MANUAL (None)
      Pregnant               → AI flag
    """
    if is_captions:
        kids_count = captions_get(df, "Kids Presenence", col_map, 0)
        preg_count = captions_get(df, "Pregnant", col_map, 0)
    else:
        kids_count = get_count(df, "Kids Presenence")
        preg_count = get_count(df, "Pregnant")

    kids_pts = 3.5 if kids_count > 0 else 0.0
    pregnant_flag = 1 if preg_count > 0 else 0

    return {
        "Kids Presence":             kids_pts,          # AI auto
        "Topics around kids":        None,              # MANUAL
        "Abbott brands":             None,              # MANUAL
        "Awards/Media Presence":     None,              # MANUAL
        "Relevant Brand Partnerships": None,            # MANUAL
        "Pregnant Or Not":           pregnant_flag,
        "_auto_base":                kids_pts,          # base before manual
        "_kids_age_summary":         extract_kids_age_summary(df) if not is_captions else "",
    }


def score_relevance_adult(df, is_captions=False, col_map=None):
    """
    Adult Nutrition relevance:
      Topics Adult Health    (AI) → 5 pts
      Topics Adult Nutrition (AI) → 2 pts
      Abbott brands          → MANUAL (None)
      Awards/Media           → MANUAL (None)
      Brand Partnerships     → MANUAL (None)
      Pregnant               → AI flag
    """
    if is_captions:
        h_count = captions_get(df, "Topics around Adult Health", col_map, 0)
        n_count = captions_get(df, "Topics on Adult Healthy Nutrition", col_map, 0)
        preg_count = captions_get(df, "Pregnant", col_map, 0)
    else:
        h_count = get_count(df, "Topics around Adult Health")
        n_count = get_count(df, "Topics on Adult Healthy Nutrition")
        preg_count = get_count(df, "Pregnant")

    h_pts = 5.0 if h_count > 0 else 0.0
    n_pts = 2.0 if n_count > 0 else 0.0
    pregnant_flag = 1 if preg_count > 0 else 0
    auto_base = h_pts + n_pts - (2.0 if pregnant_flag else 0.0)
    auto_base = max(0.0, auto_base)

    return {
        "Topics Adult Health":       h_pts,
        "Topics Adult Nutrition":    n_pts,
        "Abbott brands":             None,
        "Awards/Media Presence":     None,
        "Relevant Brand Partnerships": None,
        "Pregnant Or Not":           pregnant_flag,
        "_auto_base":                auto_base,
    }


def score_sheet(df, handle, use_case, is_captions=False):
    """Score one influencer sheet. Returns a clean result dict."""
    col_map = captions_col_map(df) if is_captions else None

    auto_reject_flags, auto_reject_sum, risk_params, risk_score = score_risk(
        df, is_captions, col_map
    )

    if use_case == "Pediatric":
        rel = score_relevance_pediatric(df, is_captions, col_map)
    else:
        rel = score_relevance_adult(df, is_captions, col_map)

    total_videos = (
        get_total(df, "Profanity Usage") or
        get_total(df, "Alcohol Use Discussion") or 0
    ) if not is_captions else (
        captions_get(df, "Profanity Usage", col_map, 1) or 0
    )

    return {
        "handle":               handle,
        "total_videos":         int(total_videos),
        "auto_reject_flags":    auto_reject_flags,
        "auto_reject_sum":      auto_reject_sum,
        "risk_params":          risk_params,
        "risk_score":           risk_score,
        "relevance":            rel,
        "use_case":             use_case,
    }

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL STYLES
# ─────────────────────────────────────────────────────────────────────────────

FN = "Arial"

# Colours
C_NAV    = "1F3864"   # dark navy
C_DKRED  = "C00000"   # dark red (auto-reject headers)
C_RED_H  = "FF0000"   # bright red (risk headers / triggered)
C_DKGRN  = "375623"   # dark green (relevance headers)
C_BLUE   = "2E75B6"   # mid blue (risk param headers)
C_MANUAL = "FFEB9C"   # yellow (manual entry)
C_GREEN  = "C6EFCE"   # green (pass)
C_AMBER  = "FFEB9C"   # amber (review)
C_RED    = "FFC7CE"   # red (fail)
C_GREY   = "F2F2F2"   # alternating row
C_NAME   = "D9E1F2"   # name cell background

def fp(hex_c):  return PatternFill("solid", fgColor=hex_c)

def bdr():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def hfont(color="FFFFFF", bold=True, sz=9):
    return Font(name=FN, bold=bold, color=color, size=sz)

def dfont(color="000000", bold=False, sz=9):
    return Font(name=FN, bold=bold, color=color, size=sz)

def ac(wrap=True):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def al():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

# ─────────────────────────────────────────────────────────────────────────────
# GRADING SHEET WRITER
# ─────────────────────────────────────────────────────────────────────────────

def write_grading_sheet(ws, scores, platform_label, use_case, profiles_map):
    """
    Write the grading sheet that exactly matches the reference file format.
    Columns are identical to the Abbott reference grading sheet.
    """

    # ── Title ─────────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 22
    t = ws.cell(1, 1, f"Abbott Influencer Vetting — {platform_label} Grading Sheet ({use_case})")
    t.font = Font(name=FN, bold=True, size=12, color=C_NAV)
    t.alignment = al()

    # ── Column definitions ─────────────────────────────────────────────────
    # Format: (header_text, bg_color, width)
    if use_case == "Pediatric":
        relevance_cols = [
            ("Kids Presence\n(3.5 pts — AI auto)",             C_DKGRN, 16),
            ("Topics around kids\n(3.5 pts — Manual)",         C_DKGRN, 18),
            ("Abbott brands\n(1 pt — Manual)",                 C_DKGRN, 16),
            ("Awards/Media Presence\n(1 pt — Manual)",         C_DKGRN, 18),
            ("Relevant Brand Partnerships\n(1 pt — Manual)",   C_DKGRN, 20),
        ]
        notes_col_label = "Kids Age Group (from AI)"
    else:
        relevance_cols = [
            ("Topics Adult Health\n(5 pts — AI auto)",         C_DKGRN, 18),
            ("Topics Adult Nutrition\n(2 pts — AI auto)",      C_DKGRN, 18),
            ("Abbott brands\n(1 pt — Manual)",                 C_DKGRN, 16),
            ("Awards/Media Presence\n(1 pt — Manual)",         C_DKGRN, 18),
            ("Relevant Brand Partnerships\n(1 pt — Manual)",   C_DKGRN, 20),
        ]
        notes_col_label = "Notes"

    COLS = [
        # (header, bg_color, width)
        ("Influencer Name",                                     C_NAV,    20),
        ("Handle / Username",                                   C_NAV,    18),
        ("Substance Use\n(Auto-reject gate)",                   C_RED_H,  14),
        ("Worked With Competition\n(Manual Check — last 90d)",  C_RED_H,  20),
        ("Anti-breastfeeding\n(Auto-reject gate)",              C_RED_H,  16),
        ("Anti-vaccination\n(Auto-reject gate)",                C_RED_H,  15),
        ("Anti-healthcare\n(Auto-reject gate)",                 C_RED_H,  15),
        ("Auto-reject Check\n(>0 = Risk must be 10)",           C_DKRED,  16),
        ("Profanity",                                           C_BLUE,   12),
        ("Alcohol",                                             C_BLUE,   12),
        ("Sensitive Content",                                   C_BLUE,   14),
        ("Stereotype & Bias",                                   C_BLUE,   14),
        ("Violence Content",                                    C_BLUE,   13),
        ("Political",                                           C_BLUE,   12),
        ("Unscientific",                                        C_BLUE,   13),
        ("Ultra-processed Food",                                C_BLUE,   14),
        ("RISK Score\n(0=Clean → 10=Reject)",                   C_NAV,    14),
    ] + relevance_cols + [
        ("Total Relevance Score\n(max 10)",                     C_DKGRN,  15),
        ("Pregnant Or Not",                                     C_DKGRN,  13),
        ("Notes / Flags",                                       "595959", 24),
        (notes_col_label,                                       "595959", 28),
    ]

    # Set widths
    for ci, (_, _, w) in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    # ── Header row ─────────────────────────────────────────────────────────
    ROW_H = 2
    ROW_D = 3
    ws.row_dimensions[ROW_H].height = 52

    for ci, (label, bg, _) in enumerate(COLS, 1):
        c = ws.cell(ROW_H, ci, label)
        c.fill      = fp(bg)
        c.font      = hfont()
        c.border    = bdr()
        c.alignment = ac()

    # Column index reference (1-based)
    CI_NAME    = 1
    CI_HANDLE  = 2
    CI_SUBST   = 3
    CI_COMP    = 4    # MANUAL
    CI_BF      = 5
    CI_VAX     = 6
    CI_HC      = 7
    CI_ARCHECK = 8
    CI_PROF    = 9
    CI_ALCO    = 10
    CI_SENS    = 11
    CI_STEREO  = 12
    CI_VIOL    = 13
    CI_POL     = 14
    CI_UNSCI   = 15
    CI_ULTRA   = 16
    CI_RISK    = 17
    CI_REL1    = 18   # Kids Presence or Topics Adult Health
    CI_REL2    = 19   # Topics around kids (manual) or Topics Adult Nutrition
    CI_REL3    = 20   # Abbott brands (MANUAL)
    CI_REL4    = 21   # Awards (MANUAL)
    CI_REL5    = 22   # Brand Partnerships (MANUAL)
    CI_TOTAL   = 23   # Total Relevance (SUM formula)
    CI_PREG    = 24
    CI_NOTE1   = 25
    CI_NOTE2   = 26

    RISK_PARAM_COLS = [CI_PROF, CI_ALCO, CI_SENS, CI_STEREO,
                       CI_VIOL, CI_POL, CI_UNSCI, CI_ULTRA]
    RISK_PARAM_LABELS = [
        "Profanity", "Alcohol", "Sensitive content", "Stereotype & Bias",
        "Violence Content", "Political", "Unscientific", "Ultra-processed Food",
    ]

    # ── Data rows ──────────────────────────────────────────────────────────
    for row_offset, s in enumerate(scores):
        row = ROW_D + row_offset
        ws.row_dimensions[row].height = 16
        row_bg = "FFFFFF" if row_offset % 2 == 0 else C_GREY
        rel = s["relevance"]

        def wc(col, value, bg=None, bold=False, left=False):
            c = ws.cell(row, col, value)
            c.fill      = fp(bg or row_bg)
            c.font      = dfont(bold=bold)
            c.border    = bdr()
            c.alignment = al() if left else ac()
            return c

        def manual_cell(col):
            """Yellow cell for manual entry."""
            c = ws.cell(row, col, "")
            c.fill      = fp(C_MANUAL)
            c.font      = dfont(color="7B4F00")
            c.border    = bdr()
            c.alignment = ac()
            return c

        # Col 1: Full name (from profiles map, or same as handle)
        full_name = profiles_map.get(s["handle"], s["handle"])
        wc(CI_NAME, full_name, bg=C_NAME, bold=True, left=True)

        # Col 2: Handle
        wc(CI_HANDLE, s["handle"], bg=C_NAME, left=True)

        # Cols 3, 5, 6, 7: Auto-reject gates (AI-detected, show flag 0/1)
        flags = s["auto_reject_flags"]
        for col_idx, label in [
            (CI_SUBST, "Substance Use"),
            (CI_BF,    "Anti-breastfeeding"),
            (CI_VAX,   "Anti-vaccination"),
            (CI_HC,    "Anti-healthcare"),
        ]:
            val = flags.get(label, 0)
            bg  = C_RED if val else C_GREEN
            wc(col_idx, val, bg=bg)

        # Col 4: Competition — MANUAL
        manual_cell(CI_COMP)

        # Col 8: Auto-reject check (SUM of flags)
        ar_sum = s["auto_reject_sum"]
        if ar_sum > 0:
            c8 = wc(CI_ARCHECK, ar_sum, bg=C_RED, bold=True)
            c8.font = Font(name=FN, bold=True, color="FFFFFF", size=9)
        else:
            wc(CI_ARCHECK, 0, bg=C_GREEN)

        # Cols 9-16: 8 risk parameters
        rp = s["risk_params"]
        for col_idx, label in zip(RISK_PARAM_COLS, RISK_PARAM_LABELS):
            val = rp.get(label, 0)
            bg  = C_RED if val == 1.0 else (C_AMBER if val == 0.5 else C_GREEN)
            wc(col_idx, val, bg=bg)

        # Col 17: Risk Score
        rs = s["risk_score"]
        if rs >= 5:
            c17 = ws.cell(row, CI_RISK, rs)
            c17.fill   = fp(C_RED_H)
            c17.font   = Font(name=FN, bold=True, color="FFFFFF", size=9)
            c17.border = bdr()
            c17.alignment = ac()
        elif rs > 0:
            wc(CI_RISK, rs, bg=C_AMBER, bold=True)
        else:
            wc(CI_RISK, rs, bg=C_GREEN, bold=True)

        # Col 18: Relevance param 1
        rel1_val = rel.get(list(rel.keys())[0])   # Kids Presence or Topics Health
        if rel1_val is None:
            manual_cell(CI_REL1)
        else:
            bg = C_GREEN if rel1_val > 0 else row_bg
            wc(CI_REL1, rel1_val, bg=bg)

        # Col 19: Relevance param 2 — always MANUAL for Pediatric (Topics around kids)
        # For Adult: Topics Adult Nutrition is AI-auto
        rel2_key = list(rel.keys())[1]
        rel2_val = rel.get(rel2_key)
        if rel2_val is None:
            manual_cell(CI_REL2)
        else:
            bg = C_GREEN if rel2_val > 0 else row_bg
            wc(CI_REL2, rel2_val, bg=bg)

        # Cols 20, 21, 22: Abbott / Awards / Partnerships — always MANUAL
        manual_cell(CI_REL3)
        manual_cell(CI_REL4)
        manual_cell(CI_REL5)

        # Col 23: Total Relevance — live SUM formula
        # Adjusts for pregnancy: if Pregnant=1, subtract 2
        r2_formula_part = (
            f"IFERROR({get_column_letter(CI_REL2)}{row},0)+"
        )
        # For pediatric, also subtract 2 if pregnant (but Pediatric = reject, not penalty)
        if use_case == "Pediatric":
            formula = (
                f"=MIN(10,MAX(0,"
                f"IFERROR({get_column_letter(CI_REL1)}{row},0)+"
                f"IFERROR({get_column_letter(CI_REL2)}{row},0)+"
                f"IFERROR({get_column_letter(CI_REL3)}{row},0)+"
                f"IFERROR({get_column_letter(CI_REL4)}{row},0)+"
                f"IFERROR({get_column_letter(CI_REL5)}{row},0)))"
            )
        else:
            formula = (
                f"=MIN(10,MAX(0,"
                f"IFERROR({get_column_letter(CI_REL1)}{row},0)+"
                f"IFERROR({get_column_letter(CI_REL2)}{row},0)+"
                f"IFERROR({get_column_letter(CI_REL3)}{row},0)+"
                f"IFERROR({get_column_letter(CI_REL4)}{row},0)+"
                f"IFERROR({get_column_letter(CI_REL5)}{row},0)-"
                f"IF({get_column_letter(CI_PREG)}{row}=1,2,0)))"
            )

        auto_base = rel.get("_auto_base", 0) or 0
        rel_bg = C_GREEN if auto_base >= 7 else (C_AMBER if auto_base >= 4 else row_bg)
        c23 = ws.cell(row, CI_TOTAL, formula)
        c23.fill      = fp(rel_bg)
        c23.font      = dfont(bold=True)
        c23.border    = bdr()
        c23.alignment = ac()

        # Col 24: Pregnant
        preg = rel.get("Pregnant Or Not", 0)
        if preg:
            c24 = wc(CI_PREG, "Pregnant — Reject", bg=C_RED, bold=True)
            c24.font = Font(name=FN, bold=True, color="C00000", size=9)
        else:
            wc(CI_PREG, 0, bg=row_bg)

        # Col 25: Notes (blank, manual)
        wc(CI_NOTE1, "", bg=row_bg, left=True)

        # Col 26: Kids age summary or notes
        age_summary = rel.get("_kids_age_summary", "")
        wc(CI_NOTE2, age_summary, bg=row_bg, left=True)

    # ── Legend ─────────────────────────────────────────────────────────────
    leg = ROW_D + len(scores) + 2
    ws.cell(leg, 1, "KEY").font = Font(name=FN, bold=True, size=9)
    for i, (clr, txt) in enumerate([
        (C_MANUAL, "🟡 Yellow = Manual entry required — fill before submitting"),
        (C_GREEN,  "🟢 Green  = Clean / Pass"),
        (C_AMBER,  "🟠 Amber  = Review required"),
        (C_RED,    "🔴 Red    = Flagged / Reject"),
        (C_RED_H,  "🔴 Bright Red = Auto-reject triggered or Risk ≥5"),
    ]):
        ws.cell(leg+1+i, 1, "").fill   = fp(clr)
        ws.cell(leg+1+i, 1, "").border = bdr()
        ws.cell(leg+1+i, 2, txt).font  = Font(name=FN, size=9)

    ws.freeze_panes = f"C{ROW_D}"


# ─────────────────────────────────────────────────────────────────────────────
# MASTER SHEET WRITER
# ─────────────────────────────────────────────────────────────────────────────

def write_master_sheet(ws, all_scores, use_case, profiles_map):
    ws.row_dimensions[1].height = 22
    ws.cell(1, 1, f"Abbott Influencer Vetting — Master Summary ({use_case})").font = \
        Font(name=FN, bold=True, size=12, color=C_NAV)

    headers = [
        ("Influencer Name",          C_NAV, 22),
        ("Handle",                   C_NAV, 18),
        ("Risk MAX\n(all platforms)", C_DKRED, 16),
        ("Relevance AVG\n(all platforms)", C_DKGRN, 18),
        ("Risk Zone",                C_NAV, 16),
        ("Relevance Zone",           C_NAV, 16),
        ("Recommendation",           C_NAV, 20),
        ("Platforms Scored",         C_NAV, 22),
    ]
    for i, (_, _, w) in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ROW_H = 3
    ROW_D = 4
    ws.row_dimensions[ROW_H].height = 40
    for ci, (lbl, bg, _) in enumerate(headers, 1):
        c = ws.cell(ROW_H, ci, lbl)
        c.fill = fp(bg); c.font = hfont(); c.border = bdr(); c.alignment = ac()

    # Consolidate across platforms
    by_handle = {}
    for platform, scores_list in all_scores.items():
        for s in scores_list:
            h = s["handle"]
            if h not in by_handle:
                by_handle[h] = {"risks": [], "rels": [], "platforms": []}
            by_handle[h]["risks"].append(s["risk_score"])
            by_handle[h]["rels"].append(s["relevance"].get("_auto_base", 0) or 0)
            by_handle[h]["platforms"].append(platform)

    for row_offset, (handle, data) in enumerate(sorted(by_handle.items())):
        row = ROW_D + row_offset
        ws.row_dimensions[row].height = 16
        row_bg = "FFFFFF" if row_offset % 2 == 0 else C_GREY

        max_risk = max(data["risks"])
        avg_rel  = round(sum(data["rels"]) / len(data["rels"]), 2)
        plats    = ", ".join(sorted(set(data["platforms"])))

        risk_zone = "🔴 Red — Reject" if max_risk >= 5 else ("🟠 Amber — Review" if max_risk > 0 else "🟢 Green — Clean")
        risk_bg   = C_RED if max_risk >= 5 else (C_AMBER if max_risk > 0 else C_GREEN)
        rel_zone  = "🟢 High (8–10)" if avg_rel >= 8 else ("🟠 Moderate (4–7)" if avg_rel >= 4 else "🔴 Low (0–3)")
        rel_bg    = C_GREEN if avg_rel >= 8 else (C_AMBER if avg_rel >= 4 else C_RED)

        if max_risk >= 5:
            rec, rec_bg = "❌ Reject", C_RED
        elif max_risk > 0:
            rec, rec_bg = "🟡 Manual Review", C_AMBER
        elif avg_rel >= 8:
            rec, rec_bg = "✅ Approve", C_GREEN
        elif avg_rel >= 4:
            rec, rec_bg = "🟡 Manual Review", C_AMBER
        else:
            rec, rec_bg = "⚠ Low Relevance", C_AMBER

        full_name = profiles_map.get(handle, handle)

        def mc(col, val, bg, bold=False, left=False):
            c = ws.cell(row, col, val)
            c.fill = fp(bg); c.font = dfont(bold=bold)
            c.border = bdr()
            c.alignment = al() if left else ac()

        mc(1, full_name, C_NAME, bold=True, left=True)
        mc(2, handle, C_NAME, left=True)
        mc(3, max_risk, risk_bg, bold=True)
        mc(4, avg_rel, rel_bg, bold=True)
        mc(5, risk_zone, risk_bg)
        mc(6, rel_zone, rel_bg)
        mc(7, rec, rec_bg, bold=True)
        mc(8, plats, row_bg)

    note_row = ROW_D + len(by_handle) + 2
    ws.cell(note_row, 1, "NOTES").font = Font(name=FN, bold=True, size=9)
    for i, note in enumerate([
        "Risk = MAX across all platforms (worst-case brand safety signal wins).",
        "Relevance = AVERAGE across platforms (consistent content identity matters).",
        "🟡 Yellow cells in grading sheets = manual input required before submitting to client.",
        "NOTE: Relevance shown here is auto-detectable base only. Add manual scores for final total.",
    ]):
        ws.cell(note_row+1+i, 1, note).font = Font(name=FN, size=8, italic=True, color="595959")

    ws.freeze_panes = f"C{ROW_D}"


# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT BUILDER
# ─────────────────────────────────────────────────────────────────────────────

PLATFORM_LABELS = {
    "TT":       "TikTok",
    "FB":       "Facebook",
    "IG_YT":    "IG + YT",
    "IG":       "Instagram",
    "Captions": "Captions",
}

def build_output_excel(all_scores, use_case, profiles_map):
    wb = openpyxl.Workbook()

    # 1. Master sheet
    ws_master = wb.active
    ws_master.title = "Master Grading Sheet"
    write_master_sheet(ws_master, all_scores, use_case, profiles_map)

    # 2. Per-platform grading sheets
    for platform in ["TT", "IG_YT", "IG", "FB", "Captions"]:
        scores = all_scores.get(platform, [])
        if not scores:
            continue
        label = PLATFORM_LABELS[platform]
        ws = wb.create_sheet(title=f"{label} Grading Sheet")
        write_grading_sheet(ws, scores, label, use_case, profiles_map)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Abbott Influencer Vetting — Auto-Grader",
    page_icon="🏥",
    layout="wide",
)

st.markdown("""
<style>
.stApp { font-family: Arial, sans-serif; }
h1 { color: #1F3864; }
h2 { color: #2E75B6; }
.green  { color: #375623; font-weight: bold; }
.amber  { color: #806000; font-weight: bold; }
.red    { color: #C00000; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

st.title("🏥 Abbott Influencer Vetting — Auto-Grader")
st.caption("Framework V4 | Supports Pediatric (Similac/PediaSure) & Adult Nutrition (Ensure/Glucerna)")

st.divider()

col_up, col_cfg = st.columns([2, 1])

with col_up:
    st.subheader("📂 Upload AI-Processed Excel File")
    st.caption("The Excel file exported from ConvoTrack / AI analysis pipeline.")
    uploaded = st.file_uploader("Drop file here", type=["xlsx"])

with col_cfg:
    st.subheader("⚙️ Settings")
    use_case_override = st.selectbox(
        "Use Case (auto-detected, override if needed)",
        ["Auto-detect", "Pediatric (Similac / PediaSure)", "Adult Nutrition (Ensure / Glucerna)"],
        index=0,
    )

st.divider()

if uploaded:
    with st.spinner("Reading file…"):
        xl = pd.read_excel(uploaded, sheet_name=None)
        grouped = detect_influencer_sheets(xl)
        profiles_map = build_profiles_map(xl)

    # Use case
    if use_case_override == "Auto-detect":
        use_case = detect_use_case(xl)
        st.info(f"🔍 Auto-detected use case: **{use_case}**")
    elif "Pediatric" in use_case_override:
        use_case = "Pediatric"
    else:
        use_case = "Adult"

    total_sheets = sum(len(v) for v in grouped.values())
    st.success(f"✅ {len(xl)} sheets found — {total_sheets} influencer content sheets detected")

    c1, c2, c3, c4, c5 = st.columns(5)
    for col_w, (plat, lbl) in zip([c1,c2,c3,c4,c5],
        [("TT","TikTok"),("FB","Facebook"),("IG_YT","IG+YT"),("IG","Instagram"),("Captions","Captions")]):
        col_w.metric(lbl, len(grouped[plat]))

    st.divider()

    # Score
    with st.spinner("Scoring all influencers…"):
        all_scores = {}
        warnings = []
        for platform, sheet_list in grouped.items():
            if not sheet_list:
                continue
            pl_scores = []
            for sheet_name, handle in sheet_list:
                df = xl[sheet_name]
                is_cap = platform == "Captions"
                try:
                    s = score_sheet(df, handle, use_case, is_cap)
                    actual = extract_username(df)
                    if actual and actual not in ("nan", "HEADER"):
                        s["handle"] = actual
                    pl_scores.append(s)
                except Exception as e:
                    warnings.append(f"{sheet_name}: {e}")
            all_scores[platform] = pl_scores

    if warnings:
        with st.expander(f"⚠ {len(warnings)} sheets had warnings"):
            for w in warnings:
                st.text(w)

    # Preview
    st.subheader("📊 Master Summary Preview")
    preview_rows = []
    by_handle = {}
    for plat, sl in all_scores.items():
        for s in sl:
            h = s["handle"]
            if h not in by_handle:
                by_handle[h] = {"risks":[], "rels":[], "plats":[]}
            by_handle[h]["risks"].append(s["risk_score"])
            by_handle[h]["rels"].append(s["relevance"].get("_auto_base",0) or 0)
            by_handle[h]["plats"].append(plat)

    for handle, data in sorted(by_handle.items()):
        max_r = max(data["risks"])
        avg_l = round(sum(data["rels"])/len(data["rels"]),2)
        rec = ("❌ Reject" if max_r>=5 else
               "✅ Approve" if max_r==0 and avg_l>=8 else "🟡 Review")
        preview_rows.append({
            "Name": profiles_map.get(handle, handle),
            "Handle": handle,
            "Risk MAX": max_r,
            "Relevance AVG (auto-base)": avg_l,
            "Platforms": ", ".join(sorted(set(data["plats"]))),
            "Recommendation": rec,
        })

    df_prev = pd.DataFrame(preview_rows)

    def style_df(row):
        styles = [""] * len(row)
        ri = df_prev.columns.get_loc("Risk MAX")
        li = df_prev.columns.get_loc("Relevance AVG (auto-base)")
        reci = df_prev.columns.get_loc("Recommendation")
        r, l, rec = row["Risk MAX"], row["Relevance AVG (auto-base)"], row["Recommendation"]
        styles[ri]  = f"background-color: {'#FFC7CE' if r>=5 else '#FFEB9C' if r>0 else '#C6EFCE'}"
        styles[li]  = f"background-color: {'#C6EFCE' if l>=7 else '#FFEB9C' if l>=4 else '#FFC7CE'}"
        styles[reci]= f"background-color: {'#C6EFCE' if 'Approve' in rec else '#FFC7CE' if 'Reject' in rec else '#FFEB9C'}; font-weight: bold"
        return styles

    st.dataframe(df_prev.style.apply(style_df, axis=1),
                 use_container_width=True, height=min(500, 50+len(preview_rows)*36))

    approve = sum(1 for r in preview_rows if "Approve" in r["Recommendation"])
    review  = sum(1 for r in preview_rows if "Review"  in r["Recommendation"])
    reject  = sum(1 for r in preview_rows if "Reject"  in r["Recommendation"])
    st.markdown(
        f"**{len(preview_rows)} influencers scored** &nbsp;|&nbsp; "
        f"<span class='green'>✅ Approve: {approve}</span> &nbsp;|&nbsp; "
        f"<span class='amber'>🟡 Review: {review}</span> &nbsp;|&nbsp; "
        f"<span class='red'>❌ Reject: {reject}</span>",
        unsafe_allow_html=True,
    )

    st.divider()
    st.subheader("📥 Download Graded Excel")

    with st.spinner("Building output Excel…"):
        buf = build_output_excel(all_scores, use_case, profiles_map)

    fname = uploaded.name.replace(".xlsx","") + f"_GRADED_{use_case}.xlsx"
    st.download_button(
        "⬇️ Download Graded Excel",
        data=buf, file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary",
    )

    st.caption(
        "⚠️ **Yellow cells require manual input before submitting to client:** "
        "Competition check (last 90 days), Abbott brand association, "
        "Awards/Media presence, Relevant brand partnerships" +
        (" and Topics around kids." if use_case == "Pediatric" else ".")
    )

else:
    st.info("👆 Upload your AI-processed Excel file to begin.")
    with st.expander("ℹ️ How it works"):
        st.markdown("""
**Input:** AI-processed Excel from ConvoTrack — one sheet per influencer per platform.

**Auto-detected use case:**
- **Pediatric** (Similac/PediaSure): Uses Kids Presence (3.5 pts auto)
- **Adult Nutrition** (Ensure/Glucerna): Uses Topics Adult Health (5 pts auto) + Topics Nutrition (2 pts auto)

**Auto-calculated (zero manual work):**
| Parameter | Detail |
|---|---|
| All 8 risk scores | Profanity, Alcohol, Sensitive, Stereotype, Violence, Political, Unscientific, Ultra-processed |
| Auto-reject gate flags | Substance Use, Anti-BF, Anti-Vax, Anti-HC — AI count shown, human confirms |
| Risk Score | (Sum of 8 params ÷ 8) × 10, or 10 if gate triggered |
| Kids Presence (Pediatric) | Any kids in videos → 3.5 pts |
| Topics Adult Health/Nutrition (Adult) | Any videos on those topics → 5 + 2 pts |
| Pregnancy flag | AI-detected |

**Manual (yellow cells — requires human review):**
- Competition check (last 90 days)
- Topics around kids (Pediatric only)
- Abbott brand association
- Awards / Media presence
- Relevant brand partnerships

**Output sheets:** Master Summary + per-platform Grading Sheets (TikTok, Facebook, IG+YT, Captions)
        """)
