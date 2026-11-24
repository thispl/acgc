import os
import re
import openpyxl
import frappe
from datetime import datetime
from frappe.utils import get_files_path, getdate
from xml.sax.saxutils import escape as xml_escape

_MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11,
    "december": 12, "jan": 1, "feb": 2, "mar": 3, "apr": 4, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
}
_MONTH_RE = re.compile(
    r"(?i)(?<![a-z])(january|february|march|april|may|june|july|august|september|"
    r"october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\b"
)

def extract_reporting_period(file_path):
    """
    Detect the reporting month/year from inside the workbook (a "MONTH: August 2026"
    style header, a date cell, or a month-name cell), falling back to the file name.
    Returns (year, month) or None.
    """
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        for name in wb.sheetnames[:6]:
            ws = wb[name]
            for row in ws.iter_rows(min_row=1, max_row=15, values_only=True):
                cells = [c for c in row if c is not None]
                for c in cells:
                    s = str(c)
                    m = _MONTH_RE.search(s)
                    if m:
                        mon = _MONTHS[m.group(1).lower()]
                        yr = re.search(r"20\d\d", s)
                        if yr:
                            wb.close()
                            return int(yr.group()), mon
                        for c2 in cells:
                            yr2 = re.search(r"20\d\d", str(c2))
                            if yr2:
                                wb.close()
                                return int(yr2.group()), mon
                for c in cells:
                    if isinstance(c, datetime):
                        wb.close()
                        return c.year, c.month
        wb.close()
    except Exception as e:
        frappe.log_error(f"Period detection failed: {e}")

    base = str(file_path)
    m = _MONTH_RE.search(base)
    yr = re.search(r"20\d\d", base)
    if m and yr:
        return int(yr.group()), _MONTHS[m.group(1).lower()]
    return None

SITE_ALIASES = {
    "PS001": "PS1", "PS01": "PS1", "PS1": "PS1", "PS 01": "PS1", "PS-01": "PS1",
    "PS3": "PS03", "PS 03": "PS03", "PS-03": "PS03",
    "PS5": "PS05", "PS 05": "PS05", "PS-05": "PS05",
    "PS6": "PS06", "PS 06": "PS06", "PS-06": "PS06",
    "PS 10": "PS10", "PS-10": "PS10",
    "PRS 1": "PRS1", "PRS-1": "PRS1", "PS R1": "PSR1",
}

def normalize_site(site):
    s = (site or "").strip().upper()
    if s in SITE_ALIASES:
        return SITE_ALIASES[s]
    compact = re.sub(r"[^A-Z0-9]", "", s)
    compact = re.sub(r"0+(?=\d)", "", compact)
    for alias, target in SITE_ALIASES.items():
        if re.sub(r"[^A-Z0-9]", "", alias) == compact:
            return target
    return compact

# ---------------------------------------------------------------------------
# Dynamic Adaptive Hierarchy Extractor (Supports Any Site, Month & Structure)
# ---------------------------------------------------------------------------

def extract_excel_hierarchy(file_path, site_name="", posting_date=None):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = None
    for name in wb.sheetnames:
        if "org" in name.lower() or "chart" in name.lower():
            ws = wb[name]
            break
    if not ws:
        ws = wb.active

    # 1. Detect Compound / Site Title from header cells
    main_title = ""
    for r in range(1, 15):
        for c in range(1, 35):
            v = ws.cell(row=r, column=c).value
            if v and any(k in str(v).lower() for k in ["compound", "pipeline", "residential", "site management"]):
                s = str(v).strip().replace("\n", " - ")
                if len(s) > 10 and not s.lower().startswith("month") and not s.lower().startswith("operation"):
                    main_title = s
                    break
        if main_title:
            break

    if not main_title:
        main_title = f"EAST WEST PIPELINE RESIDENTIAL COMPOUND - {site_name or 'PS'}"

    # 2. Extract Month & Date context
    period_str = ""
    if posting_date:
        try:
            pd = getdate(posting_date)
            period_str = pd.strftime("%B %Y")
        except Exception:
            period_str = str(posting_date)

    # 3. Footer Metadata
    footer_meta = {
        "prepared_by": "",
        "reviewed_by": "",
        "approved_by": "",
        "approved_by_label": "APPROVED BY:",
        "revision": "",
        "date": "",
        "total_count": ""
    }

    for r in range(max(1, ws.max_row - 45), ws.max_row + 1):
        for c in range(1, min(ws.max_column + 1, 450)):
            v = ws.cell(row=r, column=c).value
            if not v:
                continue
            s = str(v).strip()
            l = s.lower()
            if "prepared by" in l:
                for roff in range(1, 8):
                    nv = ws.cell(row=r+roff, column=c).value
                    if nv and str(nv).strip() and not any(k in str(nv).lower() for k in ["reviewed", "approved", "authorized", "date"]):
                        footer_meta["prepared_by"] = str(nv).strip()
                        break
            elif "reviewed by" in l:
                for roff in range(1, 8):
                    nv = ws.cell(row=r+roff, column=c).value
                    if nv and str(nv).strip() and not any(k in str(nv).lower() for k in ["prepared", "approved", "authorized", "date"]):
                        footer_meta["reviewed_by"] = str(nv).strip()
                        break
            elif "authorized by" in l:
                footer_meta["approved_by_label"] = "AUTHORIZED BY:"
                for roff in range(1, 8):
                    nv = ws.cell(row=r+roff, column=c).value
                    if nv and str(nv).strip() and not any(k in str(nv).lower() for k in ["prepared", "reviewed", "approved", "authorized", "date"]):
                        footer_meta["approved_by"] = str(nv).strip()
                        break
            elif "approved by" in l:
                footer_meta["approved_by_label"] = "APPROVED BY:"
                for roff in range(1, 8):
                    nv = ws.cell(row=r+roff, column=c).value
                    if nv and str(nv).strip() and not any(k in str(nv).lower() for k in ["prepared", "reviewed", "approved", "authorized", "date"]):
                        footer_meta["approved_by"] = str(nv).strip()
                        break
            elif "revision" in l and not footer_meta["revision"]:
                footer_meta["revision"] = s
            elif (l.startswith("date") or "date:" in l) and not footer_meta["date"]:
                footer_meta["date"] = re.sub(r"(?i)^date\s*:?\s*", "", s).strip()
            elif "as per contract" in l:
                for coff in range(1, 100):
                    cv = ws.cell(row=r, column=c+coff).value
                    if cv is not None and (isinstance(cv, (int, float)) or str(cv).strip().isdigit()):
                        footer_meta["total_count"] = str(int(cv) if isinstance(cv, (int, float)) else str(cv).strip())
                        break

    if not footer_meta["date"] and posting_date:
        try:
            footer_meta["date"] = getdate(posting_date).strftime("%d/%m/%Y")
        except Exception:
            footer_meta["date"] = str(posting_date)

    # 4. Extract all role and position cards
    cards = []
    keywords = [
        "manager", "admin", "engineer", "unit", "services", "dept", "supervisor",
        "officer", "technician", "leader", "operator", "plumber", "gardener",
        "janitor", "janiror", "driver", "imam", "security", "catering", "nesma",
        "qudrat", "rentokil", "pest", "reception", "fitness", "barber", "handyman",
        "coordinator", "camp", "boss", "clerk", "attendant", "garbage", "sewage",
        "gas", "ace"
    ]

    total_cnt = 0
    for r in range(1, ws.max_row + 1):
        for c in range(1, min(ws.max_column + 1, 450)):
            v = ws.cell(row=r, column=c).value
            if not v or not isinstance(v, str) or not v.strip():
                continue
            text = v.strip()
            lower = text.lower()
            if any(k in lower for k in ["operation &", "organisation chart - site", "east west pipeline", "prepared by", "reviewed by", "approved by", "revision -", "as per contract"]):
                continue

            count = None
            for offset in range(1, 45):
                rv = ws.cell(row=r, column=c + offset).value
                if rv is not None and (isinstance(rv, (int, float)) or str(rv).strip().isdigit()):
                    count = int(rv) if isinstance(rv, (int, float)) else int(str(rv).strip())
                    break

            lines = [l.strip() for l in text.split("\n") if l.strip()]
            if lines:
                title = lines[0]
                name = lines[1] if len(lines) > 1 else ""
                if count is not None or any(w in title.lower() for w in keywords):
                    cards.append({"row": r, "col": c, "title": title, "name": name, "count": count})
                    if count is not None and not any(h in title.lower() for h in ["unit", "dept", "services", "adminstration", "administration"]):
                        total_cnt += count

    if not footer_meta["total_count"] and total_cnt > 0:
        footer_meta["total_count"] = str(total_cnt)
    wb.close()

    # 5. Build dynamic tree mapping
    H = {
        "fm": None,
        "ga": {"header": None, "roles": []},
        "maint": {"header": None, "supervisor": None, "roles": []},
        "fe": {"header": None, "roles": []},
        "hk": {
            "header": None,
            "supervisor": None,
            "fo": {"header": None, "roles": []},
            "res": {"header": None, "roles": []},
            "ext": {"header": None, "roles": []},
            "rec": {"header": None, "roles": []}
        },
        "hse": {"header": None, "roles": []},
        "sub": {"roles": []},
        "other": {"roles": []}
    }

    for c in cards:
        t = c["title"].lower()
        if any(m in t for m in ["facility manager", "project manager", "compound manager", "site manager"]) and not H["fm"]:
            H["fm"] = c
        elif "facility engineer" in t and not H["fe"]["header"]:
            H["fe"]["header"] = c
        elif any(k in t for k in ["warehouse admin", "cmms", "helpdesk"]):
            H["fe"]["roles"].append(c)
        elif "general admin" in t and not H["ga"]["header"]:
            H["ga"]["header"] = c
        elif any(k in t for k in ["site administrator", "purchaser", "imam", "muezzin", "driver"]) and not any(k in t for k in ["maintenance", "house", "hse"]):
            H["ga"]["roles"].append(c)
        elif "maintenance unit" in t and not H["maint"]["header"]:
            H["maint"]["header"] = c
        elif "maintenance supervisor" in t and not H["maint"]["supervisor"]:
            H["maint"]["supervisor"] = c
        elif any(k in t for k in ["mechanical", "hvac", "chiller", "electrical", "fire alarm", "cctv", "plumber", "stp", "ro plant", "handyman", "gas attendant"]):
            H["maint"]["roles"].append(c)
        elif ("house keeping" in t or "housekeeping" in t) and not H["hk"]["header"]:
            H["hk"]["header"] = c
        elif "supervisor" in t and "maintenance" not in t and not H["hk"]["supervisor"]:
            H["hk"]["supervisor"] = c
        elif "front office" in t and not H["hk"]["fo"]["header"]:
            H["hk"]["fo"]["header"] = c
        elif "hous." in t or "front off" in t:
            H["hk"]["fo"]["roles"].append(c)
        elif "residential" in t and not H["hk"]["res"]["header"]:
            H["hk"]["res"]["header"] = c
        elif any(k in t for k in ["team leader", "room attendant", "tea boy"]):
            H["hk"]["res"]["roles"].append(c)
        elif "external" in t and not H["hk"]["ext"]["header"]:
            H["hk"]["ext"]["header"] = c
        elif any(k in t for k in ["gardener", "janitor", "janiror", "pest", "garbage", "sewage", "ace"]):
            H["hk"]["ext"]["roles"].append(c)
        elif ("recrational" in t or "recreational" in t) and not H["hk"]["rec"]["header"]:
            H["hk"]["rec"]["header"] = c
        elif any(k in t for k in ["fitness", "barber"]):
            H["hk"]["rec"]["roles"].append(c)
        elif "hse" in t and not H["hse"]["header"]:
            H["hse"]["header"] = c
        elif "safety officer" in t or "hse inspector" in t:
            H["hse"]["roles"].append(c)
        elif any(k in t for k in ["security", "nesma", "catering", "qudrat"]):
            H["sub"]["roles"].append(c)
        else:
            H["other"]["roles"].append(c)

    return H, main_title, footer_meta, period_str


# ---------------------------------------------------------------------------
# Flowchart Helper: Directional Arrows & Centered Cards
# ---------------------------------------------------------------------------

def draw_arrow_down(x, y1, y2, color="#5285be", stroke_width=1.4):
    """
    Renders a vertical connector line with a clean downward arrowhead polygon (▼).
    """
    if y2 <= y1:
        return ""
    line = f'<line x1="{x:.1f}" y1="{y1:.1f}" x2="{x:.1f}" y2="{y2 - 5:.1f}" stroke="{color}" stroke-width="{stroke_width}"/>'
    arrow = f'<polygon points="{x - 3.5:.1f},{y2 - 5:.1f} {x + 3.5:.1f},{y2 - 5:.1f} {x:.1f},{y2:.1f}" fill="{color}"/>'
    return f"{line}\n{arrow}"


def render_box(x, y, w, h, title, name="", count=None, box_type="standard"):
    """
    Renders an organization chart card with title and employee name
    CENTERED inside the box, using Illustrator colors and fonts.
    """
    cnt_str = str(count) if count is not None else ""
    t_clean = xml_escape(title.strip())
    n_clean = xml_escape(name.strip())

    cx = x + w / 2

    # Responsive font size based on title length
    t_len = len(t_clean)
    if t_len > 25:
        font_sz = "9.5"
    elif t_len > 19:
        font_sz = "10.5"
    else:
        font_sz = "11.5"

    has_name = bool(n_clean)
    t_y = y + 20 if has_name else y + 29
    n_y = y + 36

    if box_type == "primary_header":
        fill_col = "#386898"
        stroke_col = "#003366"
        t_col = "#ffffff"
        n_col = "#f7e97e"
        badge_bg = "#003366"
        badge_fg = "#ffffff"
        stroke_w = "1.2"
    elif box_type == "secondary_header":
        fill_col = "#c8d8e8"
        stroke_col = "#5285be"
        t_col = "#003366"
        n_col = "#b07b4e"
        badge_bg = "#386898"
        badge_fg = "#ffffff"
        stroke_w = "1.2"
    else:
        fill_col = "#e8e8f8"
        stroke_col = "#5285be"
        t_col = "#003366"
        n_col = "#b07b4e"
        badge_bg = "#286ca3"
        badge_fg = "#ffffff"
        stroke_w = "1"

    # Count badge in top-right corner
    badge_markup = ""
    if cnt_str:
        bx = x + w - 26
        by = y + 3
        badge_markup = f"""<rect x="{bx}" y="{by}" width="22" height="18" rx="3" fill="{badge_bg}" stroke="#ffffff" stroke-width="0.8"/>
            <text x="{bx + 11}" y="{by + 13}" font-size="10" font-weight="bold" fill="{badge_fg}" text-anchor="middle">{cnt_str}</text>"""

    # Center-aligned text inside the card
    name_markup = ""
    if has_name:
        name_markup = f'<text x="{cx}" y="{n_y}" font-size="10.5" font-weight="600" fill="{n_col}" text-anchor="middle">{n_clean}</text>'

    return f"""
        <g>
            <rect x="{x}" y="{y}" width="{w}" height="{h}" rx="4" fill="{fill_col}" stroke="{stroke_col}" stroke-width="{stroke_w}"/>
            <text x="{cx}" y="{t_y}" font-size="{font_sz}" font-weight="bold" fill="{t_col}" text-anchor="middle">{t_clean}</text>
            {name_markup}
            {badge_markup}
        </g>
    """


def build_exact_excel_branch_svg(H, main_title, footer_meta, site_name="", period_str="", posting_date=None):
    """
    Renders the exact hierarchical branching from the Excel workbook:
    1. Facility Manager at the apex.
    2. Main horizontal bus connects 4 primary pillars with downward arrows:
       - General Administration (Col 8)
       - Facility Engineer (Col 110) branching into:
           * Maintenance Unit (Col 57) on the left (Maintenance Supervisor + cascading trades)
           * Warehouse Admin & CMMS / HelpDesk (Col 110) on the right
       - House Keeping & Recreational Services (Col 224) branching into 4 sub-columns:
           * Front Office & Reception
           * Residential Services
           * External Services
           * Recreational Services
       - HSE DEPT. (Col 285) -> Safety Officer
    3. Standalone Separate Column (Col 293 in Excel):
       - Security Management -> Nesma -> Catering -> Qudrat
       - Kept in separate, standalone with NO connecting line to HSE,
         and NO "External Contractors" container or label.
    4. Text centered inside every card.
    5. Clean corporate Illustrator colors and fonts.
    6. Sign-off footer displaying only names present in the Excel file.
    """
    card_h = 48
    card_gap = 10
    total_w = 1740

    # Pillar dimensions matching Excel layout
    w_ga = 205
    w_maint = 205
    w_cmms = 200
    fe_gap = 16
    w_fe_group = w_maint + fe_gap + w_cmms  # 421

    w_fo = 185
    w_res = 185
    w_ext = 205
    w_rec = 185
    hk_sub_gap = 12
    w_hk_group = w_fo + w_res + w_ext + w_rec + 3 * hk_sub_gap  # 796

    w_hse = 205

    # Coordinates
    margin_x = 22
    pillar_gap = 18

    x_ga = margin_x                         # 22
    x_fe_group = x_ga + w_ga + pillar_gap   # 245
    x_hk_group = x_fe_group + w_fe_group + pillar_gap  # 684
    x_hse = x_hk_group + w_hk_group + pillar_gap       # 1498

    cx_ga = x_ga + w_ga / 2                 # 124.5
    cx_fe = x_fe_group + w_fe_group / 2     # 455.5
    cx_hk = x_hk_group + w_hk_group / 2     # 1082.0
    cx_hse = x_hse + w_hse / 2              # 1600.5

    # Sub-centers inside FE
    cx_maint = x_fe_group + w_maint / 2     # 347.5
    cx_cmms = x_fe_group + w_maint + fe_gap + w_cmms / 2  # 566.0

    # Sub-columns inside HK Group
    sub_branches_def = [
        ("Front Office & Reception", w_fo, H.get("hk", {}).get("fo", {}).get("roles", []), H.get("hk", {}).get("fo", {}).get("header", {}).get("count") if H.get("hk", {}).get("fo", {}).get("header") else None),
        ("Residential Services", w_res, H.get("hk", {}).get("res", {}).get("roles", []), H.get("hk", {}).get("res", {}).get("header", {}).get("count") if H.get("hk", {}).get("res", {}).get("header") else None),
        ("External Services", w_ext, H.get("hk", {}).get("ext", {}).get("roles", []), H.get("hk", {}).get("ext", {}).get("header", {}).get("count") if H.get("hk", {}).get("ext", {}).get("header") else None),
        ("Recrational Services", w_rec, H.get("hk", {}).get("rec", {}).get("roles", []), H.get("hk", {}).get("rec", {}).get("header", {}).get("count") if H.get("hk", {}).get("rec", {}).get("header") else None),
    ]

    sub_xs = []
    curr_hk_x = x_hk_group
    for label, w, roles, cnt in sub_branches_def:
        sub_xs.append(curr_hk_x)
        curr_hk_x += w + hk_sub_gap

    # Max depth calculation
    maint_roles_cnt = len(H.get("maint", {}).get("roles", []))
    maint_bottom_y = 276 + 48 + (1 + maint_roles_cnt) * (card_h + card_gap)

    ext_roles_cnt = len(H.get("hk", {}).get("ext", {}).get("roles", []))
    ext_bottom_y = 338 + 48 + ext_roles_cnt * (card_h + card_gap)

    sub_roles_cnt = len(H.get("sub", {}).get("roles", []))
    sub_bottom_y = 338 + sub_roles_cnt * (card_h + card_gap)

    max_content_y = max(maint_bottom_y, ext_bottom_y, sub_bottom_y, 880)
    footer_h = 75
    total_h = max_content_y + footer_h + 45

    svg = []
    svg.append(f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {total_w} {total_h}" width="100%" height="100%" style="background:#ffffff; font-family:Arial, Helvetica, sans-serif;">')
    svg.append(f'<rect x="10" y="10" width="{total_w - 20}" height="{total_h - 20}" rx="6" fill="#ffffff" stroke="#003366" stroke-width="1.2"/>')

    # Top Illustrator Header Banner
    sub_title_text = "ORGANISATION CHART - SITE MANAGEMENT"
    if period_str:
        sub_title_text += f" ({period_str.upper()})"

    svg.append(f"""
        <rect x="18" y="18" width="{total_w - 36}" height="70" rx="4" fill="#003366"/>
        <text x="{total_w / 2}" y="47" font-size="19" font-weight="bold" fill="#ffffff" text-anchor="middle" letter-spacing="0.5">{xml_escape(main_title.upper())}</text>
        <text x="{total_w / 2}" y="70" font-size="12" font-weight="600" fill="#f7e97e" text-anchor="middle" letter-spacing="0.8">{xml_escape(sub_title_text)}</text>
    """)

    # Root Node: Facility Manager
    fm = H.get("fm") or {"title": "Facility Manager", "name": "", "count": 1}
    fm_w, fm_h = 320, 56
    fm_x = (total_w - fm_w) / 2
    fm_y = 104
    top_bus_y = 180
    tree_top_y = 196

    svg.append(f"""
        <g>
            <rect x="{fm_x}" y="{fm_y}" width="{fm_w}" height="{fm_h}" rx="4" fill="#5888b8" stroke="#003366" stroke-width="1.5"/>
            <rect x="{fm_x}" y="{fm_y}" width="{fm_w}" height="24" rx="4" fill="#003366"/>
            <rect x="{fm_x}" y="{fm_y + 16}" width="{fm_w}" height="8" fill="#003366"/>
            <text x="{fm_x + fm_w / 2}" y="{fm_y + 17}" font-size="12.5" font-weight="bold" fill="#ffffff" text-anchor="middle">{xml_escape(fm['title'])}</text>
            <text x="{fm_x + fm_w / 2}" y="{fm_y + 43}" font-size="13" font-weight="bold" fill="#f7e97e" text-anchor="middle">{xml_escape(fm.get('name', ''))}</text>
            <rect x="{fm_x + fm_w - 38}" y="{fm_y + 16}" width="30" height="24" rx="3" fill="#003366" stroke="#ffffff" stroke-width="1"/>
            <text x="{fm_x + fm_w - 23}" y="{fm_y + 33}" font-size="13" font-weight="bold" fill="#ffffff" text-anchor="middle">{xml_escape(str(fm.get('count', 1)))}</text>
            <line x1="{total_w / 2}" y1="{fm_y + fm_h}" x2="{total_w / 2}" y2="{top_bus_y}" stroke="#003366" stroke-width="1.8"/>
        </g>
    """)

    # Main Horizontal Bus Line
    svg.append(f'<line x1="{cx_ga}" y1="{top_bus_y}" x2="{cx_hse}" y2="{top_bus_y}" stroke="#003366" stroke-width="1.8"/>')

    # Directional Arrows into the 4 Primary Department Pillars
    for cx in [cx_ga, cx_fe, cx_hk, cx_hse]:
        svg.append(draw_arrow_down(cx, top_bus_y, tree_top_y, color="#003366", stroke_width=1.8))

    # --- 1. GENERAL ADMINISTRATION (Left Column) ---
    ga = H.get("ga", {"header": None, "roles": []})
    ga_hdr = ga.get("header") or {"title": "General Adminstration", "name": "", "count": None}
    svg.append(render_box(x_ga, tree_top_y, w_ga, card_h, ga_hdr["title"], ga_hdr.get("name", ""), ga_hdr.get("count"), box_type="primary_header"))
    cy = tree_top_y + card_h
    for r in ga.get("roles", []):
        svg.append(draw_arrow_down(cx_ga, cy, cy + card_gap, color="#5285be", stroke_width=1.4))
        cy += card_gap
        svg.append(render_box(x_ga, cy, w_ga, card_h, r["title"], r.get("name", ""), r.get("count"), box_type="standard"))
        cy += card_h

    # --- 2. FACILITY ENGINEER GROUP ---
    fe = H.get("fe", {"header": None, "roles": []})
    fe_hdr = fe.get("header") or {"title": "Facility Engineer", "name": "", "count": 1}
    fe_w = 260
    fe_x = cx_fe - fe_w / 2
    svg.append(render_box(fe_x, tree_top_y, fe_w, card_h, fe_hdr["title"], fe_hdr.get("name", ""), fe_hdr.get("count"), box_type="primary_header"))

    # FE sub-branching bus (splits into Maintenance Unit on left & Warehouse/CMMS on right)
    fe_bus_y = tree_top_y + card_h + 16
    fe_sub_y = fe_bus_y + 16
    svg.append(f'<line x1="{cx_fe}" y1="{tree_top_y + card_h}" x2="{cx_fe}" y2="{fe_bus_y}" stroke="#003366" stroke-width="1.6"/>')
    svg.append(f'<line x1="{cx_maint}" y1="{fe_bus_y}" x2="{cx_cmms}" y2="{fe_bus_y}" stroke="#003366" stroke-width="1.6"/>')
    svg.append(draw_arrow_down(cx_maint, fe_bus_y, fe_sub_y, color="#003366", stroke_width=1.6))
    svg.append(draw_arrow_down(cx_cmms, fe_bus_y, fe_sub_y, color="#003366", stroke_width=1.6))

    # Branch 2A: Maintenance Unit (Cascading Technical Trades)
    maint = H.get("maint", {"header": None, "supervisor": None, "roles": []})
    maint_hdr = maint.get("header") or {"title": "Maintenance Unit", "name": "", "count": None}
    x_maint = cx_maint - w_maint / 2
    svg.append(render_box(x_maint, fe_sub_y, w_maint, card_h, maint_hdr["title"], maint_hdr.get("name", ""), maint_hdr.get("count"), box_type="secondary_header"))
    cy = fe_sub_y + card_h
    if maint.get("supervisor"):
        sup = maint["supervisor"]
        svg.append(draw_arrow_down(cx_maint, cy, cy + card_gap, color="#5285be", stroke_width=1.4))
        cy += card_gap
        svg.append(render_box(x_maint, cy, w_maint, card_h, sup["title"], sup.get("name", ""), sup.get("count"), box_type="secondary_header"))
        cy += card_h

    for r in maint.get("roles", []):
        svg.append(draw_arrow_down(cx_maint, cy, cy + card_gap, color="#5285be", stroke_width=1.4))
        cy += card_gap
        svg.append(render_box(x_maint, cy, w_maint, card_h, r["title"], r.get("name", ""), r.get("count"), box_type="standard"))
        cy += card_h

    # Branch 2B: Warehouse Admin & CMMS
    x_cmms = cx_cmms - w_cmms / 2
    wh_role = next((r for r in fe.get("roles", []) if "warehouse" in r["title"].lower()), None)
    if not wh_role:
        wh_role = {"title": "Warehouse Admin.", "name": "", "count": 1}
    svg.append(render_box(x_cmms, fe_sub_y, w_cmms, card_h, wh_role["title"], wh_role.get("name", ""), wh_role.get("count"), box_type="standard"))

    cmms_role = next((r for r in fe.get("roles", []) if any(k in r["title"].lower() for k in ["cmms", "help"])), None)
    if not cmms_role:
        cmms_role = {"title": "CMMS / HelpDesk", "name": "", "count": 2}
    cmms_y = fe_sub_y + card_h + card_gap
    svg.append(draw_arrow_down(cx_cmms, fe_sub_y + card_h, cmms_y, color="#5285be", stroke_width=1.4))
    svg.append(render_box(x_cmms, cmms_y, w_cmms, card_h, cmms_role["title"], cmms_role.get("name", ""), cmms_role.get("count"), box_type="standard"))

    # --- 3. HOUSE KEEPING & RECREATIONAL SERVICES GROUP ---
    hk = H.get("hk", {})
    hk_hdr = hk.get("header") or {"title": "House Keeping & Recreational Services", "name": "", "count": None}
    hk_w = 360
    hk_x = cx_hk - hk_w / 2
    svg.append(render_box(hk_x, tree_top_y, hk_w, card_h, hk_hdr["title"], hk_hdr.get("name", ""), hk_hdr.get("count"), box_type="primary_header"))

    hk_sup = hk.get("supervisor") or {"title": "Supervisor", "name": "", "count": 1}
    sup_y = tree_top_y + card_h + card_gap
    svg.append(draw_arrow_down(cx_hk, tree_top_y + card_h, sup_y, color="#003366", stroke_width=1.6))
    svg.append(render_box(hk_x, sup_y, hk_w, card_h, hk_sup["title"], hk_sup.get("name", ""), hk_sup.get("count"), box_type="secondary_header"))

    # Sub-branches under Supervisor: Front Office, Residential, External, Recreational
    hk_bus_y = sup_y + card_h + 16
    hk_sub_y = hk_bus_y + 16

    first_sub_cx = sub_xs[0] + sub_branches_def[0][1] / 2
    last_sub_cx = sub_xs[-1] + sub_branches_def[-1][1] / 2

    svg.append(f'<line x1="{cx_hk}" y1="{sup_y + card_h}" x2="{cx_hk}" y2="{hk_bus_y}" stroke="#003366" stroke-width="1.6"/>')
    svg.append(f'<line x1="{first_sub_cx}" y1="{hk_bus_y}" x2="{last_sub_cx}" y2="{hk_bus_y}" stroke="#003366" stroke-width="1.6"/>')

    for i, (label, w, roles, cnt) in enumerate(sub_branches_def):
        sx = sub_xs[i]
        scx = sx + w / 2
        svg.append(draw_arrow_down(scx, hk_bus_y, hk_sub_y, color="#003366", stroke_width=1.6))
        svg.append(render_box(sx, hk_sub_y, w, card_h, label, "", cnt, "secondary_header"))

        cy = hk_sub_y + card_h
        for r in roles:
            svg.append(draw_arrow_down(scx, cy, cy + card_gap, color="#5285be", stroke_width=1.4))
            cy += card_gap
            svg.append(render_box(sx, cy, w, card_h, r["title"], r.get("name", ""), r.get("count"), box_type="standard"))
            cy += card_h

    # --- 4. HSE DEPT. ---
    hse = H.get("hse", {"header": None, "roles": []})
    hse_hdr = hse.get("header") or {"title": "HSE DEPT.", "name": "", "count": 1}
    svg.append(render_box(x_hse, tree_top_y, w_hse, card_h, hse_hdr["title"], hse_hdr.get("name", ""), hse_hdr.get("count"), box_type="primary_header"))
    cy = tree_top_y + card_h
    for r in hse.get("roles", []):
        svg.append(draw_arrow_down(cx_hse, cy, cy + card_gap, color="#5285be", stroke_width=1.4))
        cy += card_gap
        svg.append(render_box(x_hse, cy, w_hse, card_h, r["title"], r.get("name", ""), r.get("count"), box_type="standard"))
        cy += card_h

    # --- 5. STANDALONE SEPARATE BRANCH (SECURITY, NESMA, CATERING, QUDRAT) ---
    # As in Excel Col 293 Row 136: Standalone branch alone in separate,
    # with NO connecting line to HSE DEPT, and NO "External Contractors" container or label.
    sub_roles = H.get("sub", {}).get("roles", [])
    if sub_roles:
        cy_sub = 338
        for i, r in enumerate(sub_roles):
            if i > 0:
                svg.append(draw_arrow_down(cx_hse, cy_sub, cy_sub + card_gap, color="#5285be", stroke_width=1.4))
                cy_sub += card_gap
            svg.append(render_box(x_hse, cy_sub, w_hse, card_h, r["title"], r.get("name", ""), r.get("count"), box_type="standard"))
            cy_sub += card_h

    # Custom / unclassified roles fallback
    other_roles = H.get("other", {}).get("roles", [])
    if other_roles:
        cy = 338
        for r in other_roles:
            svg.append(render_box(x_ga, cy + 300, w_ga, card_h, r["title"], r.get("name", ""), r.get("count"), box_type="standard"))
            cy += card_h + card_gap

    # --- 6. CORPORATE SIGN-OFF FOOTER ---
    fy = total_h - footer_h - 18
    fw = total_w - 36
    fx = 18

    col_prep_w = fw * 0.22
    col_rev_w = fw * 0.22
    col_app_w = fw * 0.22
    col_meta_w = fw * 0.34

    x_prep = fx
    x_rev = fx + col_prep_w
    x_app = fx + col_prep_w + col_rev_w
    x_meta = fx + col_prep_w + col_rev_w + col_app_w

    prep_val = footer_meta.get("prepared_by") or ""
    rev_val = footer_meta.get("reviewed_by") or ""
    app_val = footer_meta.get("approved_by") or ""
    app_label = footer_meta.get("approved_by_label") or "APPROVED BY:"

    svg.append(f"""
        <!-- Corporate Sign-Off Footer -->
        <g id="footer_signoff">
            <rect x="{fx}" y="{fy}" width="{fw}" height="{footer_h}" rx="3" fill="#f4f7fa" stroke="#003366" stroke-width="1.2"/>

            <!-- Prepared By -->
            <rect x="{x_prep}" y="{fy}" width="{col_prep_w}" height="24" fill="#003366"/>
            <text x="{x_prep + col_prep_w / 2}" y="{fy + 16}" font-size="10.5" font-weight="bold" fill="#ffffff" text-anchor="middle">PREPARED BY:</text>
            <text x="{x_prep + col_prep_w / 2}" y="{fy + 52}" font-size="12" font-weight="600" fill="#003366" text-anchor="middle">{xml_escape(prep_val)}</text>
            <line x1="{x_rev}" y1="{fy}" x2="{x_rev}" y2="{fy + footer_h}" stroke="#003366" stroke-width="1"/>

            <!-- Reviewed By -->
            <rect x="{x_rev}" y="{fy}" width="{col_rev_w}" height="24" fill="#003366"/>
            <text x="{x_rev + col_rev_w / 2}" y="{fy + 16}" font-size="10.5" font-weight="bold" fill="#ffffff" text-anchor="middle">REVIEWED BY:</text>
            <text x="{x_rev + col_rev_w / 2}" y="{fy + 52}" font-size="12" font-weight="600" fill="#003366" text-anchor="middle">{xml_escape(rev_val)}</text>
            <line x1="{x_app}" y1="{fy}" x2="{x_app}" y2="{fy + footer_h}" stroke="#003366" stroke-width="1"/>

            <!-- Approved / Authorized By -->
            <rect x="{x_app}" y="{fy}" width="{col_app_w}" height="24" fill="#003366"/>
            <text x="{x_app + col_app_w / 2}" y="{fy + 16}" font-size="10.5" font-weight="bold" fill="#ffffff" text-anchor="middle">{xml_escape(app_label)}</text>
            <text x="{x_app + col_app_w / 2}" y="{fy + 52}" font-size="12" font-weight="600" fill="#003366" text-anchor="middle">{xml_escape(app_val)}</text>
            <line x1="{x_meta}" y1="{fy}" x2="{x_meta}" y2="{fy + footer_h}" stroke="#003366" stroke-width="1"/>

            <!-- Revision, Date & Headcount Total -->
            <rect x="{x_meta}" y="{fy}" width="{col_meta_w}" height="24" fill="#386898"/>
            <text x="{x_meta + col_meta_w / 2}" y="{fy + 16}" font-size="10.5" font-weight="bold" fill="#ffffff" text-anchor="middle">PROJECT METADATA</text>

            <g font-size="11" font-weight="600" fill="#003366">
                <text x="{x_meta + 14}" y="{fy + 44}">Revision: <tspan fill="#b07b4e">{xml_escape(footer_meta.get("revision") or "Rev. 01")}</tspan></text>
                <text x="{x_meta + 14}" y="{fy + 62}">Date: <tspan fill="#b07b4e">{xml_escape(footer_meta.get("date") or "")}</tspan></text>

                <text x="{fx + fw - 95}" y="{fy + 53}" text-anchor="end">AS PER CONTRACT:</text>
                <rect x="{fx + fw - 85}" y="{fy + 36}" width="70" height="26" rx="4" fill="#003366"/>
                <text x="{fx + fw - 50}" y="{fy + 54}" font-size="13" font-weight="bold" fill="#ffffff" text-anchor="middle">{xml_escape(footer_meta.get("total_count") or "Total")}</text>
            </g>
        </g>
    """)

    svg.append("</svg>")
    return "\n".join(svg)


def generate_org_chart_svg(site, excel_file_path, posting_date=None):
    """
    Template-free dynamic Org Chart generator replicating the exact Excel branching,
    classic Illustrator color palette, centered text inside cards, and directional arrows.
    """
    try:
        H, main_title, footer_meta, period_str = extract_excel_hierarchy(excel_file_path, site_name=site, posting_date=posting_date)
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), f"Org chart hierarchy extraction error: {e}")
        H = {"fm": None, "ga": {"header": None, "roles": []}, "maint": {"header": None, "supervisor": None, "roles": []}, "fe": {"header": None, "roles": []}, "hk": {"header": None, "supervisor": None, "fo": {"header": None, "roles": []}, "res": {"header": None, "roles": []}, "ext": {"header": None, "roles": []}, "rec": {"header": None, "roles": []}}, "hse": {"header": None, "roles": []}, "sub": {"roles": []}, "other": {"roles": []}}
        main_title = f"EAST WEST PIPELINE RESIDENTIAL COMPOUND - {site}"
        footer_meta = {"prepared_by": "", "reviewed_by": "", "approved_by": "", "revision": "", "date": "", "total_count": ""}
        period_str = ""

    svg_content = build_exact_excel_branch_svg(H, main_title, footer_meta, site_name=site, period_str=period_str, posting_date=posting_date)

    date_suffix = ""
    if posting_date:
        p_str = str(posting_date).replace("-", "_")
        date_suffix = f"_{p_str}"

    filename = f"{site}_Org_Chart{date_suffix}.svg".replace(" ", "_")
    files_dir = get_files_path()
    output_path = os.path.join(files_dir, filename)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(svg_content)

    file_url = f"/files/{filename}"

    if not frappe.db.exists("File", {"file_url": file_url}):
        file_doc = frappe.new_doc("File")
        file_doc.file_name = filename
        file_doc.file_url = file_url
        file_doc.is_private = 0
        file_doc.save(ignore_permissions=True)

    return file_url
