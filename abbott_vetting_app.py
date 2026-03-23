"""
Abbott Influencer Vetting — Auto-Grading App
============================================
Upload the AI-processed Excel file → get fully scored grading sheets + master summary.

Scoring logic based on: Abbott Influencer Vetting Framework V4 (26 Feb 2026)
Use case: ADULT NUTRITION (Ensure / Glucerna)
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import re

# ─────────────────────────────────────────────────────────────────────────────
# SCORING CONFIGURATION — edit here if thresholds change
# ─────────────────────────────────────────────────────────────────────────────

RISK_PARAMS = [
    # (column_in_AI_file,            label_in_output,        yellow_thresh, red_thresh)
    ("Profanity Usage",              "Profanity Usage",        0.15,          0.30),
    ("Alcohol Use Discussion",       "Alcohol",                0.05,          0.15),
    ("Sensitive Visual Content",     "Sensitive Content",      None,          0.20),   # no yellow
    ("Stereotypes or Bias",          "Stereotype & Bias",      0.05,          0.15),
    ("Violence Advocacy",            "Violence Content",       0.0001,        0.05),   # any = yellow
    ("Political Stance",             "Political",              0.10,          0.25),
    ("Unscientific Claims",          "Unscientific",           None,          0.15),   # no yellow
    ("Ultra-processed food",         "Ultra-processed Food",   0.15,          0.30),
]

AUTO_REJECT_MAP = {
    "Substance Use Discussion": "Substance Use",
    "Breastfeeding":            "Anti-breastfeeding",
    "Vaccination":              "Anti-vaccination",
    "Health care stance":       "Anti-healthcare",
}

RELEVANCE_PARAMS = [
    # (column_in_AI_file,                label,                       points, any_count_triggers)
    ("Topics around Adult Health",       "Topics Adult Health",        5.0,   True),
    ("Topics on Adult Healthy Nutrition","Topics Adult Nutrition",     2.0,   True),
    ("Media Presence & Awards",          "Awards/Media Presence",      1.0,   True),
    ("Brand Partnership Presence",       "Relevant Brand Partnerships",1.0,   True),
]



# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def risk_points(pct, yellow_thresh, red_thresh):
    """Return 0 / 0.5 / 1.0 based on % prevalence."""
    if pct is None or (isinstance(pct, float) and pd.isna(pct)):
        return 0.0
    pct = float(pct)
    if pct >= red_thresh:
        return 1.0
    if yellow_thresh is not None and pct >= yellow_thresh:
        return 0.5
    return 0.0


def safe_pct(count, total):
    """Calculate % safely."""
    try:
        c, t = float(count), float(total)
        return c / t if t > 0 else 0.0
    except Exception:
        return 0.0


def get_count(df, col):
    """Get YES count from row 0 for standard AI sheets."""
    try:
        if col in df.columns:
            val = df.iloc[0][col]
            return float(val) if not pd.isna(val) else 0.0
    except Exception:
        pass
    return 0.0


def get_total(df, col):
    """Get total videos from row 1 for standard AI sheets."""
    try:
        if col in df.columns:
            val = df.iloc[1][col]
            return float(val) if not pd.isna(val) else 0.0
    except Exception:
        pass
    return 0.0


def captions_col_map(df):
    """
    Captions sheets have actual column names in row 4 (index 4).
    Build a mapping: field_name → pandas_column.
    """
    mapping = {}
    if len(df) > 4:
        header_row = df.iloc[4]
        for pandas_col, value in header_row.items():
            if pd.notna(value):
                mapping[str(value).strip()] = pandas_col
    return mapping


def captions_count(df, field_name, col_map):
    """Get count from captions sheet using the header-in-row-4 pattern."""
    pandas_col = col_map.get(field_name)
    if pandas_col is None:
        return 0.0, 0.0
    try:
        count = float(df.iloc[0][pandas_col]) if not pd.isna(df.iloc[0][pandas_col]) else 0.0
        total = float(df.iloc[1][pandas_col]) if not pd.isna(df.iloc[1][pandas_col]) else 0.0
        return count, total
    except Exception:
        return 0.0, 0.0




# ─────────────────────────────────────────────────────────────────────────────
# SCORING ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def _build_relevance(topics_health_count, topics_nutrition_count,
                     pregnant_flag):
    """
    Compute the auto-detectable portion of relevance.

    Auto (from AI data):
      Topics Around Adult Health       → 5 pts  (any video flagged)
      Topics on Adult Healthy Nutrition→ 2 pts  (any video flagged)
      Pregnancy penalty                → -2 pts (any video flagged)

    Manual (None = yellow cell in Excel, human fills in):
      Abbott Brand Association         → 1 pt
      Awards / Media Presence          → 1 pt
      Relevant Brand Partnerships      → 1 pt

    Total auto base = 0–7.  Full max with manual = 0–10 (capped).
    """
    auto_topics_health    = 5.0 if topics_health_count > 0 else 0.0
    auto_topics_nutrition = 2.0 if topics_nutrition_count > 0 else 0.0
    base = auto_topics_health + auto_topics_nutrition
    if pregnant_flag:
        base = max(0.0, base - 2.0)
    return {
        "Topics Adult Health":          auto_topics_health,
        "Topics Adult Nutrition":       auto_topics_nutrition,
        # Manual fields — None renders as yellow empty cell in Excel
        "Abbott Brand Association":     None,
        "Awards/Media Presence":        None,
        "Relevant Brand Partnerships":  None,
        # Auto base (manual fields not yet added)
        "_auto_base":                   base,
    }


def score_standard_sheet(df, influencer_handle):
    """
    Score a standard AI sheet (TT / FB / IG).
    Risk:      fully auto (per-column totals, per Framework V4).
    Relevance: Topics auto; Abbott / Awards / Brand Partnerships = manual.
    """
    # ── Auto-reject gates ──────────────────────────────────────────────────
    auto_reject = {}
    auto_reject_triggered = False
    for col, label in AUTO_REJECT_MAP.items():
        count = get_count(df, col)
        total = get_total(df, col)
        flagged = 1 if safe_pct(count, total) > 0 else 0
        auto_reject[label] = flagged
        if flagged:
            auto_reject_triggered = True

    # ── 8 risk parameters ─────────────────────────────────────────────────
    risk_scores = {}
    total_risk_pts = 0.0
    for col, label, y_thresh, r_thresh in RISK_PARAMS:
        count = get_count(df, col)
        total = get_total(df, col)
        pts   = risk_points(safe_pct(count, total), y_thresh, r_thresh)
        risk_scores[label] = pts
        total_risk_pts += pts

    final_risk = 10.0 if auto_reject_triggered else round((total_risk_pts / 8) * 10, 4)

    # ── Relevance ─────────────────────────────────────────────────────────
    health_count     = get_count(df, "Topics around Adult Health")
    nutrition_count  = get_count(df, "Topics on Adult Healthy Nutrition")
    pregnant_flag    = 1 if get_count(df, "Pregnant") > 0 else 0
    rel              = _build_relevance(health_count, nutrition_count, pregnant_flag)
    auto_base        = rel.pop("_auto_base")

    total_videos = get_total(df, "Profanity Usage") or get_total(df, "Alcohol Use Discussion") or 0

    return {
        "influencer":            influencer_handle,
        "total_videos":          int(total_videos),
        **auto_reject,
        "auto_reject_triggered": 1 if auto_reject_triggered else 0,
        **risk_scores,
        "RISK Score":            final_risk,
        **rel,
        "Total Relevance Score": round(auto_base, 4),   # updated by Excel SUM formula
        "Pregnant Or Not":       pregnant_flag,
    }


def score_captions_sheet(df, influencer_handle):
    """
    Score a Captions sheet (header row at index 4, counts at row 0/1).
    Same scoring rules as standard sheets.
    """
    col_map = captions_col_map(df)
    _, total_videos = captions_count(df, "Profanity Usage", col_map)

    # ── Auto-reject gates ──────────────────────────────────────────────────
    auto_reject = {}
    auto_reject_triggered = False
    for col, label in AUTO_REJECT_MAP.items():
        count, total = captions_count(df, col, col_map)
        flagged = 1 if safe_pct(count, total) > 0 else 0
        auto_reject[label] = flagged
        if flagged:
            auto_reject_triggered = True

    # ── 8 risk parameters ─────────────────────────────────────────────────
    risk_scores = {}
    total_risk_pts = 0.0
    for col, label, y_thresh, r_thresh in RISK_PARAMS:
        count, total = captions_count(df, col, col_map)
        pts = risk_points(safe_pct(count, total), y_thresh, r_thresh)
        risk_scores[label] = pts
        total_risk_pts += pts

    final_risk = 10.0 if auto_reject_triggered else round((total_risk_pts / 8) * 10, 4)

    # ── Relevance ─────────────────────────────────────────────────────────
    h_count, _  = captions_count(df, "Topics around Adult Health",        col_map)
    n_count, _  = captions_count(df, "Topics on Adult Healthy Nutrition",  col_map)
    p_count, _  = captions_count(df, "Pregnant",                           col_map)
    pregnant_flag = 1 if p_count > 0 else 0
    rel           = _build_relevance(h_count, n_count, pregnant_flag)
    auto_base     = rel.pop("_auto_base")

    return {
        "influencer":            influencer_handle,
        "total_videos":          int(total_videos),
        **auto_reject,
        "auto_reject_triggered": 1 if auto_reject_triggered else 0,
        **risk_scores,
        "RISK Score":            final_risk,
        **rel,
        "Total Relevance Score": round(auto_base, 4),
        "Pregnant Or Not":       pregnant_flag,
    }


# ─────────────────────────────────────────────────────────────────────────────
# SHEET DETECTION
# ─────────────────────────────────────────────────────────────────────────────

PLATFORM_PREFIXES = {
    "TT ":      "TT",
    "FB ":      "FB",
    "IG + YT ": "IG_YT",
    "IG ":      "IG",
    "Captions ":"Captions",
}


def detect_influencer_sheets(xl_sheets):
    """
    Parse sheet names → returns dict:
      platform → list of (sheet_name, influencer_handle)
    """
    grouped = {"TT": [], "FB": [], "IG_YT": [], "IG": [], "Captions": []}
    skip = {"master grading", "tt grading", "ig + yt grading",
            "fb grading", "captions grading", "grading", "profiles"}

    for sheet in xl_sheets:
        if sheet.lower().strip() in skip:
            continue
        for prefix, platform in PLATFORM_PREFIXES.items():
            if sheet.startswith(prefix):
                handle = sheet[len(prefix):].strip()
                grouped[platform].append((sheet, handle))
                break

    return grouped


def extract_username(df):
    """Try to pull platform_username from the data."""
    try:
        col = "platform_username"
        if col in df.columns:
            vals = df[col].dropna()
            if not vals.empty:
                return str(vals.iloc[0])
    except Exception:
        pass
    return None


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL OUTPUT BUILDER
# ─────────────────────────────────────────────────────────────────────────────

# Colors
C_HEADER_DARK  = "1F3864"   # dark navy
C_HEADER_MID   = "2E75B6"   # mid blue
C_HEADER_LIGHT = "D6E4F0"   # light blue
C_GREEN_BG     = "C6EFCE"   # risk green
C_YELLOW_BG    = "FFEB9C"   # manual / amber
C_RED_BG       = "FFC7CE"   # risk red
C_WHITE        = "FFFFFF"
C_GREY_BG      = "F2F2F2"
C_ORANGE_BG    = "FCE4D6"   # auto-reject warning
C_MANUAL_BG    = "FFF2CC"   # yellow for manual-entry cells

FONT_NAME = "Arial"


def hdr_font(bold=True, color="FFFFFF", size=10):
    return Font(name=FONT_NAME, bold=bold, color=color, size=size)


def cell_font(bold=False, color="000000", size=10):
    return Font(name=FONT_NAME, bold=bold, color=color, size=size)


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def thin_border():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)


def center_align(wrap=True):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)


def write_grading_sheet(ws, scores_list, platform_label, is_captions=False):
    """Write a formatted grading sheet into a worksheet."""

    # ── Title row ──────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 30
    title_cell = ws.cell(1, 1, f"Abbott Influencer Vetting — {platform_label} Grading Sheet")
    title_cell.font = Font(name=FONT_NAME, bold=True, size=13, color=C_HEADER_DARK)
    title_cell.alignment = Alignment(horizontal="left", vertical="center")

    # ── Column headers ─────────────────────────────────────────────────────
    headers = [
        "Influencer Handle",
        # Auto-reject
        "Substance Use",
        "Worked With Competition\n(Manual Check)",
        "Anti-breastfeeding",
        "Anti-vaccination",
        "Anti-healthcare",
        "Auto-Reject?\n(Risk=10 if any flag)",
        # 8 risk params
        "Profanity Usage",
        "Alcohol",
        "Sensitive Content",
        "Stereotype & Bias",
        "Violence Content",
        "Political",
        "Unscientific",
        "Ultra-processed Food",
        "RISK Score\n(0=Clean, 10=Reject)",
        # Relevance
        "Topics Adult Health\n(5 pts)",
        "Topics Adult Nutrition\n(2 pts)",
        "Abbott Brand Association\n(1 pt)",
        "Awards / Media Presence\n(1 pt)",
        "Relevant Brand Partnerships\n(1 pt)",
        "Total Relevance Score\n(max 10)",
        "Pregnant Flag",
    ]

    # Section header spans
    section_headers = [
        (1, 1,  ""),
        (2, 7,  "RISK — AUTO-REJECT GATES"),
        (8, 15, "RISK — PARAMETER SCORING"),
        (16, 16,""),
        (17, 23,"RELEVANCE SCORING"),
    ]

    ROW_SECTION = 2
    ROW_HEADER  = 3
    ROW_DATA    = 4

    # Section row
    ws.row_dimensions[ROW_SECTION].height = 20
    for start_col, end_col, label in section_headers:
        if label:
            c = ws.cell(ROW_SECTION, start_col, label)
            c.font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
            if start_col <= 7:
                c.fill = fill("C00000")   # dark red for auto-reject
            else:
                c.fill = fill(C_HEADER_MID)
            c.alignment = center_align()
            if end_col > start_col:
                ws.merge_cells(
                    start_row=ROW_SECTION, start_column=start_col,
                    end_row=ROW_SECTION, end_column=end_col
                )

    # Column header row
    ws.row_dimensions[ROW_HEADER].height = 50
    for col_idx, hdr in enumerate(headers, 1):
        c = ws.cell(ROW_HEADER, col_idx, hdr)
        if col_idx <= 7:
            c.fill = fill("FF0000") if col_idx > 1 else fill(C_HEADER_DARK)
            c.font = hdr_font(color="FFFFFF")
        elif col_idx <= 15:
            c.fill = fill(C_HEADER_MID)
            c.font = hdr_font(color="FFFFFF")
        elif col_idx == 16:
            c.fill = fill(C_HEADER_DARK)
            c.font = hdr_font(color="FFFFFF")
        else:
            c.fill = fill("375623")
            c.font = hdr_font(color="FFFFFF")
        c.alignment = center_align()
        c.border = thin_border()

    # Column widths
    col_widths = [22, 14, 22, 18, 16, 16, 18, 14, 12, 15, 15, 14, 12, 14, 18, 14,
                  16, 18, 20, 18, 22, 14, 14]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── Data rows ──────────────────────────────────────────────────────────
    for row_offset, score in enumerate(scores_list):
        row = ROW_DATA + row_offset
        ws.row_dimensions[row].height = 18

        auto_rej = score.get("auto_reject_triggered", 0)
        risk_val = score.get("RISK Score", 0)
        relevance_val = score.get("Total Relevance Score", 0)

        def write_cell(col, value, bg=C_WHITE, bold=False, center=True):
            c = ws.cell(row, col, value)
            c.fill = fill(bg)
            c.font = cell_font(bold=bold)
            c.border = thin_border()
            c.alignment = center_align() if center else Alignment(horizontal="left",
                                                                    vertical="center",
                                                                    wrap_text=True)
            return c

        # Col 1: Influencer name
        write_cell(1, score["influencer"], C_GREY_BG, bold=True, center=False)

        # Cols 2-6: Auto-reject gates (AI-detected)
        for col_idx, key in enumerate(
            ["Substance Use", "Anti-breastfeeding", "Anti-vaccination", "Anti-healthcare"],
            2
        ):
            val = score.get(key, 0)
            bg = "FFC7CE" if val else C_GREEN_BG
            write_cell(col_idx, val, bg)

        # Col 3: Competition (MANUAL — yellow)
        c = ws.cell(row, 3, "")
        c.fill = fill(C_MANUAL_BG)
        c.border = thin_border()
        c.alignment = center_align()
        c.font = cell_font(color="8B4513")

        # Col 7: Auto-reject summary
        bg7 = "FF0000" if auto_rej else C_GREEN_BG
        txt7 = "⚠ AUTO-REJECT" if auto_rej else "Pass"
        c7 = write_cell(7, txt7, bg7, bold=auto_rej)
        if auto_rej:
            c7.font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)

        # Cols 8-15: 8 risk parameters
        risk_param_keys = [
            "Profanity Usage", "Alcohol", "Sensitive Content",
            "Stereotype & Bias", "Violence Content", "Political",
            "Unscientific", "Ultra-processed Food"
        ]
        for col_idx, key in enumerate(risk_param_keys, 8):
            val = score.get(key, 0)
            if val == 1.0:
                bg = C_RED_BG
            elif val == 0.5:
                bg = C_YELLOW_BG
            else:
                bg = C_GREEN_BG
            write_cell(col_idx, val, bg)

        # Col 16: RISK Score
        if risk_val == 10:
            risk_bg = "FF0000"
            risk_font_color = "FFFFFF"
        elif risk_val >= 5:
            risk_bg = "FF0000"
            risk_font_color = "FFFFFF"
        elif risk_val > 0:
            risk_bg = C_YELLOW_BG
            risk_font_color = "000000"
        else:
            risk_bg = C_GREEN_BG
            risk_font_color = "000000"
        c16 = write_cell(16, risk_val, risk_bg, bold=True)
        c16.font = Font(name=FONT_NAME, bold=True, color=risk_font_color, size=10)

        # Cols 17-21: Relevance (AI auto-detected; None = manual for captions)
        rel_keys = [
            "Topics Adult Health", "Topics Adult Nutrition",
            "Abbott Brand Association", "Awards/Media Presence",
            "Relevant Brand Partnerships"
        ]
        for col_idx, key in enumerate(rel_keys, 17):
            val = score.get(key)
            if val is None:
                # Manual field — yellow highlight, empty
                c = ws.cell(row, col_idx, "")
                c.fill = fill(C_MANUAL_BG)
                c.border = thin_border()
                c.alignment = center_align()
                c.font = cell_font(color="8B4513")
            else:
                bg = C_GREEN_BG if val > 0 else C_WHITE
                write_cell(col_idx, val, bg)

        # Col 22: Total Relevance — live formula so it recalculates when manual cells filled
        # Cols: Topics Health=Q(17), Topics Nutrition=R(18), Abbott=S(19), Awards=T(20), Brand=U(21)
        # Pregnancy penalty is in col 23
        rel_sum_cols = f"{get_column_letter(17)}{row}:{get_column_letter(21)}{row}"
        preg_col     = get_column_letter(23)
        formula = f"=MIN(10,MAX(0,SUM({rel_sum_cols})-IF({preg_col}{row}=1,2,0)))"
        c22 = ws.cell(row, 22, formula)
        # Colour based on static auto_base (changes when human fills in manual cells)
        auto_base = (score.get("Topics Adult Health") or 0) + (score.get("Topics Adult Nutrition") or 0)
        if score.get("Pregnant Or Not"):
            auto_base = max(0, auto_base - 2)
        if auto_base >= 8:
            rel_bg = C_GREEN_BG
        elif auto_base >= 4:
            rel_bg = C_YELLOW_BG
        else:
            rel_bg = C_RED_BG
        c22.fill  = fill(rel_bg)
        c22.font  = cell_font(bold=True)
        c22.border = thin_border()
        c22.alignment = center_align()

        # Col 23: Pregnant flag
        preg = score.get("Pregnant Or Not", 0)
        write_cell(23, "⚠ Pregnant" if preg else 0, "FFC7CE" if preg else C_GREEN_BG)

    # ── Legend ────────────────────────────────────────────────────────────
    legend_row = ROW_DATA + len(scores_list) + 2
    ws.cell(legend_row, 1, "LEGEND").font = Font(name=FONT_NAME, bold=True, size=10)
    items = [
        (C_MANUAL_BG, "Manual check required — fill in before finalising"),
        (C_GREEN_BG,  "Green = Pass / Clean"),
        (C_YELLOW_BG, "Amber = Review / Moderate risk"),
        (C_RED_BG,    "Red = Fail / High risk"),
    ]
    for i, (color, text) in enumerate(items):
        c = ws.cell(legend_row + 1 + i, 1, "")
        c.fill = fill(color)
        c.border = thin_border()
        ws.cell(legend_row + 1 + i, 2, text).font = cell_font(size=9)

    ws.freeze_panes = f"B{ROW_DATA}"


def write_master_sheet(ws, all_scores_by_platform):
    """
    Master sheet: Risk = MAX, Relevance = AVERAGE across all platforms per influencer.
    """
    ws.row_dimensions[1].height = 30
    ws.cell(1, 1, "Abbott Influencer Vetting — Master Grading Summary").font = Font(
        name=FONT_NAME, bold=True, size=13, color=C_HEADER_DARK
    )

    # Collect all unique influencers
    all_influencers = set()
    for platform, scores_list in all_scores_by_platform.items():
        for s in scores_list:
            all_influencers.add(s["influencer"])
    all_influencers = sorted(all_influencers)

    headers = [
        "Influencer Handle",
        "Risk MAX Score\n(across all platforms)",
        "Relevance AVG Score\n(across all platforms)",
        "Risk Zone",
        "Relevance Zone",
        "Recommendation",
        "Platforms Assessed",
    ]

    ROW_HEADER = 3
    ROW_DATA   = 4

    ws.row_dimensions[ROW_HEADER].height = 45
    for ci, h in enumerate(headers, 1):
        c = ws.cell(ROW_HEADER, ci, h)
        c.fill = fill(C_HEADER_DARK)
        c.font = hdr_font()
        c.alignment = center_align()
        c.border = thin_border()

    col_widths = [24, 20, 22, 16, 16, 22, 24]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for row_offset, influencer in enumerate(all_influencers):
        row = ROW_DATA + row_offset
        ws.row_dimensions[row].height = 18

        risk_vals = []
        rel_vals  = []
        platforms_found = []

        for platform, scores_list in all_scores_by_platform.items():
            for s in scores_list:
                if s["influencer"] == influencer:
                    risk_vals.append(s["RISK Score"])
                    rel_vals.append(s["Total Relevance Score"])
                    platforms_found.append(platform)

        if not risk_vals:
            continue

        max_risk   = max(risk_vals)
        avg_rel    = round(sum(rel_vals) / len(rel_vals), 2)
        platforms_str = ", ".join(sorted(set(platforms_found)))

        # Risk zone
        if max_risk == 0:
            risk_zone = "Green — Clean"
            risk_bg   = C_GREEN_BG
        elif max_risk <= 5:
            risk_zone = "Amber — Review"
            risk_bg   = C_YELLOW_BG
        else:
            risk_zone = "Red — Reject"
            risk_bg   = C_RED_BG

        # Relevance zone
        if avg_rel >= 8:
            rel_zone = "High (8–10)"
            rel_bg   = C_GREEN_BG
        elif avg_rel >= 4:
            rel_zone = "Moderate (4–7)"
            rel_bg   = C_YELLOW_BG
        else:
            rel_zone = "Low (0–3)"
            rel_bg   = C_RED_BG

        # Recommendation
        if max_risk > 5:
            rec = "❌ Reject"
            rec_bg = C_RED_BG
        elif max_risk > 0 and avg_rel >= 7:
            rec = "🟡 Manual Review"
            rec_bg = C_YELLOW_BG
        elif avg_rel >= 8 and max_risk == 0:
            rec = "✅ Approve"
            rec_bg = C_GREEN_BG
        elif avg_rel >= 4:
            rec = "🟡 Manual Review"
            rec_bg = C_YELLOW_BG
        else:
            rec = "⚠ Low Relevance"
            rec_bg = C_ORANGE_BG

        def mc(col, value, bg, bold=False):
            c = ws.cell(row, col, value)
            c.fill = fill(bg)
            c.font = cell_font(bold=bold)
            c.border = thin_border()
            c.alignment = center_align()

        mc(1, influencer, C_GREY_BG, bold=True)
        mc(2, max_risk,   risk_bg,   bold=True)
        mc(3, avg_rel,    rel_bg,    bold=True)
        mc(4, risk_zone,  risk_bg)
        mc(5, rel_zone,   rel_bg)
        mc(6, rec,        rec_bg,    bold=True)
        mc(7, platforms_str, C_WHITE)

    # Notes
    note_row = ROW_DATA + len(all_influencers) + 2
    notes = [
        "Risk = MAX across all platforms — a single harmful instance anywhere is a brand safety risk.",
        "Relevance = AVERAGE across all platforms — content identity should be consistent, not a peak moment.",
        "⚠ Yellow cells in platform sheets = Competition check (manual). Fill before finalising Risk Score.",
        "Abbott Brand Association & Brand Partnerships auto-detected from AI video data (review flagged brand names).",
    ]
    ws.cell(note_row, 1, "METHODOLOGY NOTES").font = Font(name=FONT_NAME, bold=True, size=10)
    for i, note in enumerate(notes):
        c = ws.cell(note_row + 1 + i, 1, note)
        c.font = Font(name=FONT_NAME, size=9, italic=True, color="595959")

    ws.freeze_panes = f"B{ROW_DATA}"


def build_output_excel(all_scores_by_platform):
    """Build the complete output Excel file in memory."""
    wb = openpyxl.Workbook()

    # Sheet order: Master → TT → IG_YT → FB → Captions
    platform_labels = {
        "TT":       "TikTok",
        "IG_YT":    "Instagram + YouTube",
        "IG":       "Instagram",
        "FB":       "Facebook",
        "Captions": "Captions",
    }

    # Master sheet
    ws_master = wb.active
    ws_master.title = "Master Grading Sheet"
    write_master_sheet(ws_master, all_scores_by_platform)

    # Per-platform sheets
    for platform, scores_list in all_scores_by_platform.items():
        if not scores_list:
            continue
        label = platform_labels.get(platform, platform)
        ws = wb.create_sheet(title=f"{label} Grading")
        write_grading_sheet(ws, scores_list, label)

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

# CSS
st.markdown("""
<style>
.stApp { font-family: Arial, sans-serif; }
.metric-card {
    background: #f8f9fa; border-radius: 8px; padding: 16px;
    border-left: 4px solid #2E75B6; margin-bottom: 12px;
}
h1 { color: #1F3864; }
h2 { color: #2E75B6; }
.green  { color: #375623; font-weight: bold; }
.yellow { color: #806000; font-weight: bold; }
.red    { color: #C00000; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

st.title("🏥 Abbott Influencer Vetting — Auto-Grader")
st.caption("Adult Nutrition Use Case (Ensure / Glucerna) | Framework V4 | March 2026")

st.divider()

col1, col2 = st.columns([2, 1])

with col1:
    st.subheader("📂 Upload AI-Processed Excel File")
    st.caption("Upload the Excel file exported from ConvoTrack / AI analysis pipeline.")
    uploaded = st.file_uploader(
        "Drop your file here",
        type=["xlsx"],
        help="The file should contain raw AI analysis sheets per influencer per platform (TT, FB, IG, Captions)."
    )

with col2:
    st.subheader("⚙️ Settings")
    brand_context = st.selectbox(
        "Brand / Use Case",
        ["Adult Nutrition (Ensure / Glucerna)", "Pediatric (Similac / PediaSure)"],
        index=0,
        help="Affects relevance scoring rules."
    )
    show_details = st.checkbox("Show per-influencer breakdown preview", value=True)

st.divider()

if uploaded:
    with st.spinner("Reading file and detecting influencer sheets…"):
        xl = pd.read_excel(uploaded, sheet_name=None)
        sheet_names = list(xl.keys())
        grouped = detect_influencer_sheets(sheet_names)

    # ── Detection summary ──────────────────────────────────────────────────
    total_sheets = sum(len(v) for v in grouped.values())
    st.success(f"✅ File loaded — {len(sheet_names)} total sheets detected, {total_sheets} influencer content sheets found.")

    c1, c2, c3, c4, c5 = st.columns(5)
    for col_widget, (platform, label) in zip(
        [c1, c2, c3, c4, c5],
        [("TT","TikTok"), ("FB","Facebook"), ("IG_YT","IG+YT"), ("IG","Instagram"), ("Captions","Captions")]
    ):
        col_widget.metric(label, len(grouped[platform]), "sheets")

    st.divider()

    # ── Run scoring ────────────────────────────────────────────────────────
    with st.spinner("Scoring all influencers…"):
        all_scores = {}

        for platform, sheet_list in grouped.items():
            if not sheet_list:
                continue
            platform_scores = []
            for sheet_name, handle in sheet_list:
                df = xl[sheet_name]
                try:
                    if platform == "Captions":
                        score = score_captions_sheet(df, handle)
                    else:
                        score = score_standard_sheet(df, handle)
                    # Try to use actual username from data
                    actual_handle = extract_username(df)
                    if actual_handle and actual_handle != "nan":
                        score["influencer"] = actual_handle
                    platform_scores.append(score)
                except Exception as e:
                    st.warning(f"Could not score {sheet_name}: {e}")
            all_scores[platform] = platform_scores

    # ── Preview ────────────────────────────────────────────────────────────
    if show_details:
        st.subheader("📊 Scoring Preview")

        # Consolidate master scores
        all_influencers_set = set()
        for pl_scores in all_scores.values():
            for s in pl_scores:
                all_influencers_set.add(s["influencer"])

        master_rows = []
        for inf in sorted(all_influencers_set):
            risk_vals, rel_vals, platforms = [], [], []
            for platform, pl_scores in all_scores.items():
                for s in pl_scores:
                    if s["influencer"] == inf:
                        risk_vals.append(s["RISK Score"])
                        rel_vals.append(s["Total Relevance Score"])
                        platforms.append(platform)
            if risk_vals:
                max_risk = max(risk_vals)
                avg_rel  = round(sum(rel_vals) / len(rel_vals), 2)
                if max_risk > 5:
                    rec = "❌ Reject"
                elif max_risk > 0 and avg_rel >= 7:
                    rec = "🟡 Review"
                elif avg_rel >= 8 and max_risk == 0:
                    rec = "✅ Approve"
                elif avg_rel >= 4:
                    rec = "🟡 Review"
                else:
                    rec = "⚠ Low Rel."
                master_rows.append({
                    "Influencer": inf,
                    "Risk MAX": max_risk,
                    "Relevance AVG": avg_rel,
                    "Platforms": ", ".join(sorted(set(platforms))),
                    "Recommendation": rec
                })

        df_master = pd.DataFrame(master_rows)

        def style_row(row):
            styles = [""] * len(row)
            idx_risk = df_master.columns.get_loc("Risk MAX")
            idx_rel  = df_master.columns.get_loc("Relevance AVG")
            idx_rec  = df_master.columns.get_loc("Recommendation")

            risk = row["Risk MAX"]
            rel  = row["Relevance AVG"]
            rec  = row["Recommendation"]

            if risk > 5:
                styles[idx_risk] = "background-color: #FFC7CE; font-weight: bold"
            elif risk > 0:
                styles[idx_risk] = "background-color: #FFEB9C"
            else:
                styles[idx_risk] = "background-color: #C6EFCE"

            if rel >= 8:
                styles[idx_rel] = "background-color: #C6EFCE; font-weight: bold"
            elif rel >= 4:
                styles[idx_rel] = "background-color: #FFEB9C"
            else:
                styles[idx_rel] = "background-color: #FFC7CE"

            if "Approve" in rec:
                styles[idx_rec] = "background-color: #C6EFCE; font-weight: bold"
            elif "Reject" in rec:
                styles[idx_rec] = "background-color: #FFC7CE; font-weight: bold; color: #C00000"
            else:
                styles[idx_rec] = "background-color: #FFEB9C"
            return styles

        st.dataframe(
            df_master.style.apply(style_row, axis=1),
            use_container_width=True,
            height=min(450, 40 + len(master_rows) * 36)
        )

        # Quick stats
        approved = len([r for r in master_rows if "Approve" in r["Recommendation"]])
        review   = len([r for r in master_rows if "Review" in r["Recommendation"]])
        rejected = len([r for r in master_rows if "Reject" in r["Recommendation"]])

        st.markdown(f"""
        **Summary:** {len(master_rows)} influencers scored &nbsp;|&nbsp;
        <span class='green'>✅ Approve: {approved}</span> &nbsp;|&nbsp;
        <span class='yellow'>🟡 Review: {review}</span> &nbsp;|&nbsp;
        <span class='red'>❌ Reject: {rejected}</span>
        """, unsafe_allow_html=True)

    st.divider()

    # ── Generate and download ──────────────────────────────────────────────
    st.subheader("📥 Download Grading Sheets")

    with st.spinner("Building formatted Excel output…"):
        output_buffer = build_output_excel(all_scores)

    st.success("✅ Output ready! Click below to download.")

    filename = uploaded.name.replace(".xlsx", "") + "_GRADED.xlsx"
    st.download_button(
        label="⬇️ Download Graded Excel",
        data=output_buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )

    st.caption(
        "⚠️ **Manual checks still required:** (1) Competition partnerships (last 90 days) — "
        "highlighted yellow in each grading sheet. "
        "(2) Review AI-detected brand names for relevance accuracy before submitting to client."
    )

else:
    st.info("👆 Upload your AI-processed Excel file to begin automated grading.")

    with st.expander("ℹ️ How this works"):
        st.markdown("""
        **Input:** The Excel file from your ConvoTrack / AI pipeline — one sheet per influencer per platform.

        **Auto-calculated (zero manual work):**
        - All 8 risk parameter scores (Profanity, Alcohol, Sensitive, Stereotype, Violence, Political, Unscientific, Ultra-processed)
        - Auto-reject gate checks (Substance Use, Anti-BF, Anti-Vax, Anti-HC)
        - Risk Score per platform: `(sum of 8 params / 8) × 10`
        - Topics Adult Health (5 pts) + Topics Adult Nutrition (2 pts)
        - Awards / Media Presence (detected from AI data)
        - Brand Partnerships — health/nutrition brand keyword detection
        - Abbott Brand Association (keyword match on brand names)
        - Pregnancy flag penalty (−2 pts on relevance)
        - Master sheet: Risk = MAX, Relevance = AVERAGE across all platforms

        **Still manual (flagged in yellow):**
        - Competition check (last 90 days) — requires human Google search

        **Output:** Formatted Excel with:
        - TikTok Grading Sheet
        - Facebook Grading Sheet
        - Instagram / YouTube Grading Sheet
        - Captions Grading Sheet
        - Master Summary (all influencers, final recommendation)
        """)
