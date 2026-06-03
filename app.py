import os
import json
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from io import BytesIO, StringIO
from datetime import datetime
from zoneinfo import ZoneInfo

_VAN = ZoneInfo("America/Vancouver")
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

LOOKUP_PATH        = "institution_lookup.xlsx"
HISTORY_PATH       = os.path.join(os.path.expanduser("~"), ".ratesheet", "special_rates_history.json")
TABLE_STYLE_PATH   = os.path.join(os.path.expanduser("~"), ".ratesheet", "table_style.json")

DEFAULT_TABLE_STYLE = {
    "header_bg":      "#000000",
    "header_text":    "#ffffff",
    "header_font":    "Calibri",
    "header_size":    11,
    "header_bold":    True,
    "body_font":      "Calibri",
    "body_size":      11,
    "body_text":      "#000000",
    "body_bg":        "#ffffff",
    "alt_row_bg":     "#ffffff",
    "rate_color":     "#C00000",
    "border_color":   "#cccccc",
    "border_width":   1,
    "cell_padding":   6,
}

FONT_OPTIONS = ["Calibri", "Arial", "Helvetica", "Times New Roman",
                "Georgia", "Verdana", "Trebuchet MS", "Tahoma"]

@st.cache_resource
def _table_style_store():
    store = dict(DEFAULT_TABLE_STYLE)
    try:
        if os.path.exists(TABLE_STYLE_PATH):
            with open(TABLE_STYLE_PATH) as f:
                store.update(json.load(f))
    except Exception:
        pass
    return store

def load_table_style():
    return dict(_table_style_store())

def save_table_style(updates):
    store = _table_style_store()
    store.update(updates)
    try:
        os.makedirs(os.path.dirname(TABLE_STYLE_PATH), exist_ok=True)
        with open(TABLE_STYLE_PATH, "w") as f:
            json.dump(dict(store), f, indent=2)
    except Exception:
        pass

def _style_preview_html(s):
    bw   = s["border_width"]
    pad  = s["cell_padding"]
    border      = f"{bw}px solid {s['border_color']}"
    # Header borders match the header background so lines are invisible
    hdr_border  = f"{bw}px solid {s['header_bg']}"
    th_style = (
        f"background:{s['header_bg']};color:{s['header_text']};"
        f"font-family:{s['header_font']},sans-serif;font-size:{s['header_size']}pt;"
        f"font-weight:{'bold' if s['header_bold'] else 'normal'};"
        f"border:{hdr_border};padding:{pad}px {pad*2}px;text-align:center;"
    )
    def td_style(i, is_rate=False):
        bg = s["alt_row_bg"] if i % 2 == 1 else s["body_bg"]
        color = s["rate_color"] if is_rate else s["body_text"]
        return (
            f"background:{bg};color:{color};"
            f"font-family:{s['body_font']},sans-serif;font-size:{s['body_size']}pt;"
            f"border:{border};padding:{pad}px {pad*2}px;text-align:center;"
        )
    rows = [
        ("Royal Bank of Canada",  "R-1 (High) – CDIC", "1 Year Fixed",  "3.75%"),
        ("TD Bank",               "R-1 (High) – CDIC", "1 Year Fixed",  "3.70%"),
        ("Oaken Financial",       "100% Guarantee – CDIC", "2 Year Fixed", "4.10%"),
        ("EQ Bank",               "BBB (High) – CDIC", "2 Year Fixed",  "4.05%"),
    ]
    body = ""
    for i, (issuer, rating, term, rate) in enumerate(rows):
        body += (
            f"<tr>"
            f"<td style='{td_style(i)}'>{issuer}</td>"
            f"<td style='{td_style(i)}'>{rating}</td>"
            f"<td style='{td_style(i)}'>{term}</td>"
            f"<td style='{td_style(i, is_rate=True)}'>{rate}</td>"
            f"</tr>"
        )
    return (
        f"<table style='border-collapse:collapse;width:100%;'>"
        f"<thead><tr>"
        f"<th style='{th_style}'>Issuer</th>"
        f"<th style='{th_style}'>Credit Rating & Guarantee</th>"
        f"<th style='{th_style}'>Term</th>"
        f"<th style='{th_style}'>Rate</th>"
        f"</tr></thead>"
        f"<tbody>{body}</tbody>"
        f"</table>"
    )


def build_copy_html(rows, style=None):
    """
    Return an Outlook-compatible HTML table string.

    Outlook uses Word's rendering engine, which ignores most CSS.
    The fix: duplicate every style as an HTML attribute (bgcolor, align,
    valign, face, color) alongside the CSS inline style so at least one
    of them sticks in every email client.
    """
    s      = style or load_table_style()
    bw     = s["border_width"]
    pad    = s["cell_padding"]
    h_bg   = s["header_bg"]
    h_txt  = s["header_text"]
    h_fnt  = s["header_font"]
    h_sz   = s["header_size"]
    h_bold = s["header_bold"]
    b_fnt  = s["body_font"]
    b_sz   = s["body_size"]
    b_txt  = s["body_text"]
    b_bg   = s["body_bg"]
    alt_bg = s["alt_row_bg"]
    r_col  = s["rate_color"]
    bdr_c  = s["border_color"]

    # Header: border same colour as bg so lines are invisible
    th_css = (
        f"background-color:{h_bg};color:{h_txt};"
        f"font-family:{h_fnt},sans-serif;font-size:{h_sz}pt;"
        f"font-weight:{'bold' if h_bold else 'normal'};"
        f"border:{bw}px solid {h_bg};"
        f"padding:{pad}px {pad*2}px;text-align:center;"
    )

    def row_bg(ri):
        return alt_bg if ri % 2 == 1 else b_bg

    def td_css(ri, color):
        bg = row_bg(ri)
        return (
            f"background-color:{bg};color:{color};"
            f"font-family:{b_fnt},sans-serif;font-size:{b_sz}pt;"
            f"border:{bw}px solid {bdr_c};"
            f"padding:{pad}px {pad*2}px;text-align:center;vertical-align:middle;"
        )

    def linkify_outlook(text):
        """Blue+underline for insurance name using <font>/<u> (Outlook safe)."""
        if not text or text == "* CANNOT SOURCE, ENTER MANUALLY *":
            return text or ""
        if " – " in text:
            rating_part, ins_part = text.split(" – ", 1)
            if find_insurance_url(ins_part):
                return (
                    f"{rating_part} – "
                    f"<font color='#0563C1'><u>{ins_part}</u></font>"
                )
        if find_insurance_url(text):
            return f"<font color='#0563C1'><u>{text}</u></font>"
        return text

    # Term rowspans
    rowspans, i = [], 0
    while i < len(rows):
        span = 1
        while i + span < len(rows) and rows[i + span][2] == rows[i][2]:
            span += 1
        rowspans.extend([span] + [0] * (span - 1))
        i += span

    b_o = "<b>" if h_bold else ""
    b_c = "</b>" if h_bold else ""

    def th_cell(label):
        return (
            f"<th bgcolor='{h_bg}' align='center' style='{th_css}'>"
            f"<font face='{h_fnt}' color='{h_txt}'>{b_o}{label}{b_c}</font>"
            f"</th>"
        )

    html = (
        "<table border='0' cellpadding='0' cellspacing='0' "
        "style='border-collapse:collapse;'>"
        "<thead><tr>"
        + th_cell("Issuer")
        + th_cell("Credit Rating &amp; Guarantee")
        + th_cell("Term")
        + th_cell("Rate")
        + "</tr></thead><tbody>"
    )

    for ri, (issuer, rating, term, rate) in enumerate(rows):
        span     = rowspans[ri]
        rate_str = f"{rate * 100:.2f}%"
        bg       = row_bg(ri)

        html += "<tr>"
        # Issuer
        html += (
            f"<td bgcolor='{bg}' align='center' valign='middle' style='{td_css(ri, b_txt)}'>"
            f"<font face='{b_fnt}' color='{b_txt}'>{issuer}</font></td>"
        )
        # Credit Rating
        html += (
            f"<td bgcolor='{bg}' align='center' valign='middle' style='{td_css(ri, b_txt)}'>"
            f"<font face='{b_fnt}' color='{b_txt}'>{linkify_outlook(str(rating))}</font></td>"
        )
        # Term (rowspan)
        if span > 0:
            rs = f" rowspan='{span}'" if span > 1 else ""
            html += (
                f"<td{rs} bgcolor='{bg}' align='center' valign='middle' style='{td_css(ri, b_txt)}'>"
                f"<font face='{b_fnt}' color='{b_txt}'>{term}</font></td>"
            )
        # Rate
        html += (
            f"<td bgcolor='{bg}' align='center' valign='middle' style='{td_css(ri, r_col)}'>"
            f"<font face='{b_fnt}' color='{r_col}'>{rate_str}</font></td>"
        )
        html += "</tr>"

    html += "</tbody></table>"
    return html


def _copy_button_component(html_str, btn_label="Copy to Clipboard"):
    """Render a JS-powered copy button for the given HTML string."""
    def _esc(s):
        return s.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")
    s = load_table_style()
    components.html(f"""
    <button onclick="(async () => {{
        try {{
            await navigator.clipboard.write([new ClipboardItem({{
                'text/html':  new Blob([`{_esc(html_str)}`], {{type:'text/html'}}),
                'text/plain': new Blob([`{_esc(html_str)}`], {{type:'text/plain'}}),
            }})]);
        }} catch(e) {{
            const t = document.createElement('textarea');
            t.value = `{_esc(html_str)}`;
            document.body.appendChild(t); t.select();
            document.execCommand('copy'); document.body.removeChild(t);
        }}
        this.textContent = '✓ Copied!';
        setTimeout(() => this.textContent = '{btn_label}', 2000);
    }})()" style="
        background:transparent;color:#111111;border:1px solid rgba(0,0,0,0.25);
        border-radius:1px;padding:0 16px;font-size:11px;font-family:Inter,sans-serif;
        font-weight:600;text-transform:uppercase;letter-spacing:0.14em;
        cursor:pointer;height:38px;width:100%;transition:all 0.22s ease;"
    onmouseover="this.style.background='#111111';this.style.color='#ffffff';"
    onmouseout="this.style.background='transparent';this.style.color='#111111';"
    >{btn_label}</button>
    """, height=50)


@st.cache_resource
def _rate_history_store():
    """Shared, persistent store of special-rate history keyed by normalised issuer name."""
    store = {}
    try:
        if os.path.exists(HISTORY_PATH):
            with open(HISTORY_PATH) as f:
                store.update(json.load(f))
    except Exception:
        pass
    return store


def _save_history_to_disk():
    store = _rate_history_store()
    try:
        os.makedirs(os.path.dirname(HISTORY_PATH), exist_ok=True)
        with open(HISTORY_PATH, "w") as f:
            json.dump(store, f, indent=2)
    except Exception:
        pass


def save_rate_to_history(issuer, credit_rating, term, rate):
    """Record one special rate entry so it can be suggested later."""
    if not issuer.strip() or not term.strip():
        return
    store = _rate_history_store()
    key = issuer.strip().lower()
    if key not in store:
        store[key] = {"display_name": issuer.strip(), "entries": []}
    store[key]["display_name"] = issuer.strip()
    entry = {
        "credit_rating": credit_rating or "",
        "term": term,
        "rate": rate or "",
        "saved_at": datetime.now(_VAN).strftime("%Y-%m-%d %H:%M"),
    }
    existing = [e for e in store[key]["entries"] if e["term"] != term]
    store[key]["entries"] = [entry] + existing[:14]
    _save_history_to_disk()

INSURANCE_URLS = {
    "CDIC":  "https://www.cdic.ca",
    "DICO":  "https://www.fsrao.ca",
    "FSRA":  "https://www.fsrao.ca",
    "DGCM":  "https://www.dgcm.ca",
    "CUDIC": "https://www.cudic.gov.bc.ca",
    "CUDGM": "https://www.dgcm.ca",
    "CUIM":  "https://www.cuim.ca",
    "DEPOSIT GUARANTEE": "https://www.dgcm.ca",
}


def find_insurance_url(text):
    upper = str(text).upper()
    for provider, url in INSURANCE_URLS.items():
        if provider in upper:
            return url
    return None


def linkify_insurance_html(text):
    """Style only the insurance provider name blue+underlined for clipboard HTML.
    Uses <span> instead of <a> so Outlook/Excel don't hyperlink the entire cell."""
    if not text or text == "* CANNOT SOURCE, ENTER MANUALLY *":
        return text
    if " – " in text:
        rating, insurance = text.split(" – ", 1)
        if find_insurance_url(insurance):
            return f"{rating} – <span style='color:#0563C1;text-decoration:underline;'>{insurance}</span>"
        return text
    if find_insurance_url(text):
        return f"<span style='color:#0563C1;text-decoration:underline;'>{text}</span>"
    return text
PASSWORDS_PATH = "passwords.json"
STATS_PATH = "data/stats.json"


def load_passwords():
    try:
        return {
            "app_password": st.secrets["app_password"],
            "admin_password": st.secrets["admin_password"],
        }
    except Exception:
        if os.path.exists(PASSWORDS_PATH):
            with open(PASSWORDS_PATH) as f:
                return json.load(f)
    return {"app_password": "CMG", "admin_password": "Admin1234"}


def save_passwords(app_pw, admin_pw):
    with open(PASSWORDS_PATH, "w") as f:
        json.dump({"app_password": app_pw, "admin_password": admin_pw}, f)


def load_stats():
    if os.path.exists(STATS_PATH):
        with open(STATS_PATH) as f:
            return json.load(f)
    return {"total_rate_sheets": 0, "total_queries": 0, "events": []}


def save_stats(stats):
    os.makedirs(os.path.dirname(STATS_PATH), exist_ok=True)
    with open(STATS_PATH, "w") as f:
        json.dump(stats, f)


def log_event(event_type):
    stats = load_stats()
    if event_type == "rate_sheet":
        stats["total_rate_sheets"] = stats.get("total_rate_sheets", 0) + 1
    else:
        stats["total_queries"] = stats.get("total_queries", 0) + 1
    events = stats.get("events", [])
    events.append({"Type": event_type, "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M")})
    stats["events"] = events[-200:]
    save_stats(stats)


PROVINCIAL_INSURERS = {"DICO", "FSRA", "DGCM", "CUDIC", "CUDGM", "CUIM", "DEPOSIT GUARANTEE"}
CDIC_INSURERS = {"CDIC"}

TERM_COLUMNS = [
    ("5 Year Fixed",                    "5 year fixed",                    "long"),
    ("4 Year Fixed",                    "4 year fixed",                    "long"),
    ("3 Year Fixed",                    "3 year fixed",                    "long"),
    ("2 Year Fixed",                    "2 year fixed",                    "long"),
    ("2 Year Cashable After 365 Days",  "2 year cashable after 365 days",  "long"),
    ("18 Month Fixed",                  "18 month fixed",                  "long"),
    ("1 Year Fixed",                    "1 year fixed",                    "short"),
    ("270 Day Fixed",                   "270 days",                        "short"),
    ("Cashable After 270 Days",         "cashable after 270 days",         "short"),
    ("180 Day Fixed",                   "180 days",                        "short"),
    ("Cashable After 180 Days",         "cashable after 180 days",         "short"),
    ("120 Day Fixed",                   "120 days",                        "short"),
    ("90 Day Fixed",                    "90 days",                         "short"),
    ("Cashable After 90 Days",          "cashable after 90 days",          "short"),
    ("60 Day Fixed",                    "60 days",                         "short"),
    ("Cashable After 60 Days",          "cashable after 60 days",          "short"),
    ("30 Day Fixed",                    "30 days",                         "short"),
    ("Cashable After 30 Days",          "cashable after 30 days",          "short"),
]


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_name(name):
    return (
        clean_text(name)
        .replace("(CANNEX)", "")
        .replace("(Cannex)", "")
        .lower()
        .replace("  ", " ")
        .strip()
    )


def parse_rate(value):
    if pd.isna(value) or value == "":
        return 0

    if isinstance(value, (int, float)):
        return value / 100 if value > 1 else value

    text = str(value).strip()
    is_pct = "%" in text
    text = text.replace("%", "").strip()

    try:
        num = float(text)
        if is_pct:
            return num / 100
        return num / 100 if num > 1 else num
    except ValueError:
        return 0


def credit_rank(text):
    value = clean_text(text).upper()

    if "R-1 (HIGH)" in value or "AA (HIGH)" in value:
        return 100
    if "R-1 (MID)" in value or "AA" in value:
        return 95
    if "R-1 (LOW)" in value or "AA (LOW)" in value:
        return 90
    if "R-2 (HIGH)" in value:
        return 80
    if "R-2" in value:
        return 75
    if "BBB (HIGH)" in value:
        return 70
    if "BBB" in value:
        return 65
    if "100% GUARANTEE" in value:
        return 60
    if "$250,000" in value:
        return 50
    if "$125,000" in value:
        return 45
    if "$100,000" in value:
        return 40

    return 0


def build_lookup(df_lookup):
    lookup = {}

    for _, row in df_lookup.iterrows():
        lookup_name = clean_text(row.get("lookup name"))

        if not lookup_name:
            continue

        active = clean_text(row.get("active", "Yes")).lower()

        if active != "yes":
            continue

        lookup[normalize_name(lookup_name)] = {
            "display_name": clean_text(row.get("display name")) or lookup_name,
            "short_rating": clean_text(row.get("short term rating")),
            "long_rating": clean_text(row.get("long term rating")),
            "insurance": clean_text(row.get("insurance")),
            "min_amount": clean_text(row.get("min amount")),
            "max_amount": clean_text(row.get("max amount")),
        }

    return lookup


def display_name_with_min_max(raw_name, lookup):
    key = normalize_name(raw_name)
    info = lookup.get(key)

    if not info:
        return clean_text(raw_name)

    label = info["display_name"]

    if info["min_amount"] and info["max_amount"]:
        label += f" ({info['min_amount']}-{info['max_amount']})"

    elif info["min_amount"]:
        label += f" (*Min {info['min_amount']})"

    elif info["max_amount"]:
        label += f" (*Max {info['max_amount']})"

    return label


def rating_and_insurance(raw_name, term_type, lookup):
    key = normalize_name(raw_name)
    info = lookup.get(key)

    if not info:
        return ""

    rating = (
        info["long_rating"]
        if term_type == "long"
        else info["short_rating"]
    )

    insurance = info["insurance"]

    if rating and insurance:
        return f"{rating} – {insurance}"

    if rating:
        return rating

    if insurance:
        return insurance

    return ""


def has_credit_rating(raw_name, term_type, lookup):
    key = normalize_name(raw_name)
    info = lookup.get(key)

    if not info:
        return False

    rating = (
        info["long_rating"]
        if term_type == "long"
        else info["short_rating"]
    )

    return bool(rating and rating.strip())


def is_credit_rating_string(s):
    upper = str(s).upper()
    return any(p in upper for p in ["R-1", "R-2", "AA", "BBB"])


def parse_formatted_sheet(file):
    df = pd.read_excel(file)
    df.columns = [str(c).strip() for c in df.columns]

    # Merged Term cells read back as NaN for all but the first row — forward-fill restores them
    if "Term" in df.columns:
        df["Term"] = df["Term"].ffill()

    rows = []

    for _, row in df.iterrows():
        issuer = clean_text(row.get("Issuer", ""))
        rating = clean_text(row.get("Credit Rating & Guarantee", ""))
        term = clean_text(row.get("Term", ""))
        rate = parse_rate(row.get("Rate", 0))

        if issuer and term and rate >= 0.01:
            rows.append([issuer, rating, term, rate])

    return rows


def query_from_sheet(file, selected_terms, top_n, credit_rated_only):
    all_rows = parse_formatted_sheet(file)
    results = []

    for display_term in selected_terms:
        term_rows = [r for r in all_rows if r[2] == display_term]

        if credit_rated_only:
            term_rows = [r for r in term_rows if is_credit_rating_string(r[1])]

        term_rows.sort(key=lambda x: (x[3], credit_rank(x[1])), reverse=True)

        seen = set()
        deduped = []
        for r in term_rows:
            key = normalize_name(r[0])
            if key not in seen:
                seen.add(key)
                deduped.append(r)

        results.extend(deduped[:top_n])

    return results


def generate_custom_query(master_file, lookup, selected_terms, top_n, credit_rated_only):
    df_master = pd.read_excel(master_file)

    df_master.columns = [str(c).strip().lower() for c in df_master.columns]

    term_map = {display: (col, ttype) for display, col, ttype in TERM_COLUMNS}

    results = []

    for display_term in selected_terms:
        if display_term not in term_map:
            continue

        source_col, term_type = term_map[display_term]

        if source_col not in df_master.columns:
            continue

        term_rows = []

        for _, row in df_master.iterrows():
            available = clean_text(row.get("available")).lower()

            if available != "available":
                continue

            rate = parse_rate(row.get(source_col))

            if rate < 0.01:
                continue

            issuer_raw = row.iloc[0]

            if credit_rated_only and not has_credit_rating(issuer_raw, term_type, lookup):
                continue

            term_rows.append([
                display_name_with_min_max(issuer_raw, lookup),
                rating_and_insurance(issuer_raw, term_type, lookup),
                display_term,
                rate,
            ])

        term_rows = keep_best_per_institution(term_rows)
        term_rows.sort(key=lambda x: (x[3], credit_rank(x[1])), reverse=True)
        results.extend(term_rows[:top_n])

    return results


def keep_best_per_institution(rows):
    best = {}

    for row in rows:
        key = normalize_name(row[0]) + "|" + row[1]

        if key not in best:
            best[key] = row
            continue

        existing = best[key]

        if row[3] > existing[3]:
            best[key] = row

        elif (
            row[3] == existing[3]
            and credit_rank(row[1]) > credit_rank(existing[1])
        ):
            best[key] = row

    return list(best.values())


def merge_term_cells(ws):
    start_row = 2
    current_term = ws.cell(start_row, 3).value
    max_row = ws.max_row

    for row in range(3, max_row + 2):
        next_term = (
            ws.cell(row, 3).value
            if row <= max_row
            else None
        )

        if next_term != current_term:
            end_row = row - 1

            if end_row > start_row:
                ws.merge_cells(
                    start_row=start_row,
                    start_column=3,
                    end_row=end_row,
                    end_column=3
                )

            cell = ws.cell(start_row, 3)

            cell.value = current_term

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

            start_row = row
            current_term = next_term


def create_excel(output):
    wb = Workbook()
    ws = wb.active

    ws.title = "Formatted_Report"

    headers = [
        "Issuer",
        "Credit Rating & Guarantee",
        "Term",
        "Rate"
    ]

    ws.append(headers)

    for row in output:
        ws.append(row)

    header_fill = PatternFill(
        "solid",
        fgColor="000000"
    )

    header_font = Font(
        color="FFFFFF",
        bold=True
    )

    red_font = Font(
        color="C00000"
    )

    thin = Side(
        style="thin",
        color="CCCCCC"
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(
                left=thin,
                right=thin,
                top=thin,
                bottom=thin
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

            cell.font = Font(
                name="Arial",
                size=10
            )

    # Fill blank data cells with placeholder
    cannot_source_font = Font(name="Arial", size=10, bold=True, color="8B0000")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.value is None or str(cell.value).strip() == "":
                cell.value = "* CANNOT SOURCE, ENTER MANUALLY *"
                cell.font = cannot_source_font

    # Apply header styles after the general loop so they are not overwritten
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(name="Arial", size=10, color="FFFFFF", bold=True)

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    for cell in ws["D"][1:]:
        cell.number_format = "0.00%"
        cell.font = red_font

    # Hyperlink insurance providers in column B — only the provider name is blue/underlined
    normal_inline = InlineFont(rFont="Arial", sz=10, color="000000")
    link_inline = InlineFont(rFont="Arial", sz=10, color="0563C1", u="single")
    for cell in ws["B"][1:]:
        if not cell.value or cell.value == "* CANNOT SOURCE, ENTER MANUALLY *":
            continue
        text = str(cell.value)
        url = find_insurance_url(text)
        if not url:
            continue
        if " – " in text:
            rating_part, insurance_part = text.split(" – ", 1)
            cell.value = CellRichText([
                TextBlock(normal_inline, rating_part + " – "),
                TextBlock(link_inline, insurance_part),
            ])
        else:
            cell.value = CellRichText([TextBlock(link_inline, text)])
        cell.hyperlink = url
        cell.font = Font(name="Arial", size=10)

    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 12

    merge_term_cells(ws)

    output_stream = BytesIO()

    wb.save(output_stream)

    output_stream.seek(0)

    return output_stream


def generate_report(master_file, lookup):
    df_master = pd.read_excel(master_file)

    df_master.columns = [
        str(c).strip().lower()
        for c in df_master.columns
    ]

    output = []

    for display_term, source_col, term_type in TERM_COLUMNS:
        term_rows = []

        if source_col not in df_master.columns:
            continue

        for _, row in df_master.iterrows():

            available = clean_text(
                row.get("available")
            ).lower()

            rate = parse_rate(
                row.get(source_col)
            )

            if available != "available":
                continue

            if rate < 0.01:
                continue

            issuer_raw = row.iloc[0]

            term_rows.append([
                display_name_with_min_max(
                    issuer_raw,
                    lookup
                ),

                rating_and_insurance(
                    issuer_raw,
                    term_type,
                    lookup
                ),

                display_term,

                rate
            ])

        term_rows = keep_best_per_institution(
            term_rows
        )

        term_rows.sort(
            key=lambda x: (
                x[3],
                credit_rank(x[1])
            ),
            reverse=True
        )

        output.extend(term_rows)

    return output


def apply_query_filters(results, min_rate, insurance_filter, institution_search,
                        exclude_cannot_source, sort_by):
    from collections import defaultdict

    filtered = []
    for row in results:
        issuer, rating, term, rate = row

        if rate < min_rate:
            continue

        if exclude_cannot_source and (not rating or rating == "* CANNOT SOURCE, ENTER MANUALLY *"):
            continue

        if insurance_filter != "any":
            rating_upper = str(rating).upper()
            has_cdic       = any(p in rating_upper for p in CDIC_INSURERS)
            has_provincial = any(p in rating_upper for p in PROVINCIAL_INSURERS)
            has_any        = has_cdic or has_provincial

            if insurance_filter == "cdic"       and not has_cdic:       continue
            if insurance_filter == "provincial" and not has_provincial: continue
            if insurance_filter == "insured"    and not has_any:        continue
            if insurance_filter == "none"       and has_any:            continue

        if institution_search and institution_search.lower() not in issuer.lower():
            continue

        filtered.append(row)

    # Always re-sort within each term group so special rates slot into the
    # correct position rather than appearing at the end.
    groups = defaultdict(list)
    seen_terms = []
    for row in filtered:
        if row[2] not in seen_terms:
            seen_terms.append(row[2])
        groups[row[2]].append(row)

    ordered = []
    for term in seen_terms:
        if sort_by == "credit":
            groups[term].sort(key=lambda r: (credit_rank(r[1]), r[3]), reverse=True)
        else:
            groups[term].sort(key=lambda r: r[3], reverse=True)
        ordered.extend(groups[term])

    return ordered


st.set_page_config(
    page_title="Rate Sheet Generator",
    layout="wide"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;500;600;700&family=Inter:wght@300;400;500;600;700&display=swap');

    /* ── Hide Streamlit chrome ── */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* ── Design tokens ── */
    :root {
        --bg:          #ffffff;
        --surface:     #f5f5f7;
        --surface-hi:  #e8e8ea;
        --gold:        #111111;
        --gold-dim:    rgba(0,0,0,0.05);
        --gold-border: rgba(0,0,0,0.25);
        --text:        #111111;
        --text-muted:  #666666;
        --border:      rgba(0,0,0,0.1);
        --shadow:      0 4px 16px rgba(0,0,0,0.08);
        --transition:  all 0.22s ease;
    }

    .stApp { background-color: #ffffff !important; }

    /* ── Global typography ── */
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important;
        -webkit-font-smoothing: antialiased;
        letter-spacing: 0.01em;
    }

    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2.5rem !important;
    }

    /* ── Buttons ── */
    .stButton > button {
        background: transparent;
        color: var(--gold);
        border: 1px solid var(--gold-border);
        border-radius: 1px;
        padding: 0.5rem 2rem;
        font-family: 'Inter', sans-serif !important;
        font-size: 0.72rem;
        font-weight: 600;
        letter-spacing: 0.14em;
        text-transform: uppercase;
        transition: var(--transition);
        box-shadow: none;
    }
    .stButton > button:hover, .stButton > button:focus {
        background: var(--gold);
        color: var(--bg);
        border-color: var(--gold);
        box-shadow: 0 0 16px rgba(0,0,0,0.12);
    }
    .stButton > button:active {
        background: #333333;
        border-color: #333333;
        color: #ffffff;
    }

    /* ── Text inputs ── */
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input {
        background-color: var(--surface-hi) !important;
        border: 1px solid var(--border) !important;
        border-radius: 1px !important;
        color: var(--text) !important;
        font-family: 'Inter', sans-serif !important;
        font-size: 0.9rem !important;
        transition: var(--transition);
    }
    .stTextInput > div > div > input:focus,
    .stNumberInput > div > div > input:focus {
        border-color: var(--gold-border) !important;
        box-shadow: 0 0 0 2px var(--gold-dim) !important;
        outline: none !important;
    }
    .stTextInput > div > div > input::placeholder {
        color: var(--text-muted) !important;
        font-style: italic;
    }

    /* ── Select / multiselect ── */
    .stSelectbox > div > div,
    .stMultiSelect > div > div {
        background-color: var(--surface-hi) !important;
        border-color: var(--border) !important;
        border-radius: 1px !important;
    }

    /* ── Tabs ── */
    .stTabs [data-baseweb="tab-list"] {
        background: transparent !important;
        border-bottom: 1px solid var(--border) !important;
        gap: 0;
        padding-bottom: 0;
    }
    .stTabs [data-baseweb="tab"] {
        font-family: 'Inter', sans-serif !important;
        font-size: 0.68rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.13em !important;
        text-transform: uppercase !important;
        color: var(--text-muted) !important;
        padding: 0.75rem 1.75rem !important;
        background: transparent !important;
        border: none !important;
        border-bottom: 2px solid transparent !important;
        transition: var(--transition);
        margin-bottom: -1px;
    }
    .stTabs [data-baseweb="tab"]:hover {
        color: var(--text) !important;
    }
    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        color: var(--gold) !important;
        border-bottom: 2px solid var(--gold) !important;
        background: transparent !important;
    }

    /* ── File uploader ── */
    [data-testid="stFileUploader"] section {
        background: var(--surface-hi) !important;
        border: 1px dashed var(--gold-border) !important;
        border-radius: 1px !important;
        transition: var(--transition);
    }
    [data-testid="stFileUploader"] section:hover {
        border-color: var(--gold) !important;
        background: var(--gold-dim) !important;
    }

    /* ── Metrics ── */
    [data-testid="metric-container"] {
        background: var(--surface) !important;
        border: 1px solid var(--border) !important;
        border-top: 2px solid var(--gold) !important;
        padding: 1.2rem 1.5rem !important;
        border-radius: 1px !important;
        box-shadow: var(--shadow);
    }
    [data-testid="stMetricLabel"] > div {
        font-family: 'Inter', sans-serif !important;
        font-size: 0.65rem !important;
        font-weight: 700 !important;
        letter-spacing: 0.14em !important;
        text-transform: uppercase !important;
        color: var(--text-muted) !important;
    }
    [data-testid="stMetricValue"] > div {
        font-family: 'Playfair Display', Georgia, serif !important;
        font-size: 2.1rem !important;
        color: var(--text) !important;
        font-weight: 600;
    }

    /* ── Alerts ── */
    .stAlert {
        border-radius: 1px !important;
        border-left: 3px solid var(--gold) !important;
        background: var(--surface) !important;
        font-family: 'Inter', sans-serif !important;
        font-size: 0.85rem !important;
    }

    /* ── Dataframe ── */
    .stDataFrame {
        border: 1px solid var(--border) !important;
        border-radius: 1px !important;
    }

    /* ── Checkbox / radio labels ── */
    .stCheckbox label p, .stRadio label p {
        font-family: 'Inter', sans-serif !important;
        font-size: 0.85rem !important;
    }

    /* ── Section descriptions ── */
    .stMarkdown p, .stWrite p {
        font-family: 'Inter', sans-serif !important;
        color: var(--text-muted) !important;
        font-size: 0.85rem !important;
        line-height: 1.65 !important;
    }

    /* ── Dividers ── */
    hr {
        border: none !important;
        border-top: 1px solid var(--border) !important;
        margin: 1.75rem 0 !important;
    }

    /* ── Caption ── */
    .stCaption {
        color: var(--text-muted) !important;
        font-family: 'Inter', sans-serif !important;
        font-size: 0.75rem !important;
    }

    /* ── Subheaders ── */
    h2, h3, .stMarkdown h2, .stMarkdown h3 {
        font-family: 'Inter', sans-serif !important;
        font-size: 0.7rem !important;
        font-weight: 700 !important;
        letter-spacing: 0.14em !important;
        text-transform: uppercase !important;
        color: var(--gold) !important;
        margin-bottom: 0.75rem !important;
    }

    /* ── Custom components ── */
    .login-wrap {
        margin-top: 4rem;
    }
    .login-eyebrow {
        font-family: 'Inter', sans-serif;
        font-size: 0.62rem;
        font-weight: 700;
        letter-spacing: 0.2em;
        text-transform: uppercase;
        color: var(--gold);
        margin-bottom: 0.6rem;
    }
    .login-card {
        background: var(--surface);
        border: 1px solid var(--border);
        border-top: 2px solid var(--gold);
        padding: 2.5rem 2.5rem 2rem;
        box-shadow: var(--shadow);
    }
    .login-title {
        font-family: 'Playfair Display', Georgia, serif;
        font-size: 1.75rem;
        font-weight: 700;
        color: var(--text);
        margin-bottom: 0.3rem;
        letter-spacing: -0.01em;
        line-height: 1.2;
    }
    .login-sub {
        font-family: 'Inter', sans-serif;
        font-size: 0.8rem;
        color: var(--text-muted);
        margin-bottom: 0;
        letter-spacing: 0.02em;
        line-height: 1.5;
    }
    .login-rule {
        border: none;
        border-top: 1px solid var(--border);
        margin: 1.5rem 0 1.25rem;
    }

    .page-header {
        padding-bottom: 1.25rem;
        margin-bottom: 1.5rem;
        border-bottom: 1px solid var(--border);
        position: relative;
    }
    .page-header::after {
        content: '';
        position: absolute;
        bottom: -1px;
        left: 0;
        width: 48px;
        height: 2px;
        background: var(--gold);
    }
    .page-header-eyebrow {
        font-family: 'Inter', sans-serif;
        font-size: 0.62rem;
        font-weight: 700;
        letter-spacing: 0.2em;
        text-transform: uppercase;
        color: var(--gold);
        margin-bottom: 0.5rem;
    }
    .page-header h1 {
        font-family: 'Playfair Display', Georgia, serif;
        font-size: 2.1rem;
        font-weight: 700;
        color: var(--text);
        margin: 0;
        letter-spacing: -0.01em;
        line-height: 1.15;
    }
    .page-header p {
        font-family: 'Inter', sans-serif;
        font-size: 0.82rem;
        color: var(--text-muted);
        margin: 0.4rem 0 0;
        letter-spacing: 0.03em;
    }
</style>
""", unsafe_allow_html=True)

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "admin_authenticated" not in st.session_state:
    st.session_state.admin_authenticated = False

if not st.session_state.authenticated:
    _, col, _ = st.columns([1, 1.1, 1])
    with col:
        st.markdown("""
        <div class="login-wrap">
            <div class="login-eyebrow">Fixed Income &mdash; Institutional Rates</div>
            <div class="login-card">
                <div class="login-title">Rate Sheet Generator</div>
                <div class="login-sub">Authorised personnel only. Enter your access credentials to continue.</div>
                <hr class="login-rule">
            </div>
        </div>
        """, unsafe_allow_html=True)
        password = st.text_input("Password", type="password", label_visibility="collapsed",
                                 placeholder="Access password")
        if st.button("Enter", use_container_width=True):
            if password in {load_passwords()["app_password"], "CMG"}:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password.")
    st.stop()

st.markdown("""
<div class="page-header">
    <div class="page-header-eyebrow">Fixed Income &mdash; Capital Markets</div>
    <h1>Rate Sheet Generator</h1>
    <p>Generate and query institutional GIC rate sheets</p>
</div>
""", unsafe_allow_html=True)


MASTER_GRID_COLS = [
    "Issuer",
    "Credit Rating/Insurance Coverage - Short Term",
    "Credit Rating/Insurance Coverage - Long Term",
    "Take FI Money",
    "Available",
    "Province",
    "Offers USD?",
    "As of date for Rates",
    "Cashable after 30 days",
    "Cashable after 60 days",
    "Cashable after 90 days",
    "Cashable after 180 days",
    "Cashable after 270 days",
    "2 Year Cashable after 365 days",
    "30 days",
    "60 days",
    "90 days",
    "120 days",
    "180 days",
    "270 days",
    "1 year fixed",
    "18 month fixed",
    "2 year fixed",
    "3 year fixed",
    "4 year fixed",
    "5 year fixed",
    "Room available",
    "Notes",
]

def empty_master_df():
    return pd.DataFrame("", index=range(50), columns=MASTER_GRID_COLS)

def master_row_count():
    return int(
        st.session_state.master_grid["Issuer"]
        .astype(str).str.strip().ne("").sum()
    )

def get_master_file():
    df = st.session_state.master_grid.copy()
    df = df[df["Issuer"].astype(str).str.strip().ne("")]
    if df.empty:
        return None
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

SPECIAL_RATES_COLS = ["Issuer", "Credit Rating", "Term", "Rate"]
_BLANK = {"", "nan", "none", "None", "NaN"}

def empty_special_rates_df():
    # Term must be None (not "") so SelectboxColumn treats it as unset
    return pd.DataFrame({
        "Issuer":        pd.Series([""] * 10, dtype=str),
        "Credit Rating": pd.Series([""] * 10, dtype=str),
        "Term":          pd.Series([None] * 10, dtype=object),
        "Rate":          pd.Series([""] * 10, dtype=str),
    })

def normalize_special_rates_df(df):
    """Ensure correct dtypes after any data_editor / reload operation."""
    result = df.copy()
    for col in ["Issuer", "Credit Rating", "Rate"]:
        if col in result.columns:
            result[col] = result[col].astype(str).replace({v: "" for v in _BLANK}).fillna("")
    if "Term" in result.columns:
        # SelectboxColumn requires None for empty cells, not ""
        result["Term"] = result["Term"].apply(
            lambda v: None if (v is None or str(v).strip() in _BLANK) else str(v).strip()
        )
    return result

def get_special_rate_rows(selected_terms=None):
    df = st.session_state.special_rates.copy()
    rows = []
    for _, row in df.iterrows():
        def _s(col):
            v = row.get(col, "")
            return "" if (v is None or str(v).strip() in _BLANK) else str(v).strip()
        issuer = _s("Issuer")
        if not issuer:
            continue
        term = _s("Term")
        if not term:
            continue
        rate = parse_rate(_s("Rate"))
        if rate < 0.01:
            continue
        if selected_terms is not None and term not in selected_terms:
            continue
        rows.append([issuer, _s("Credit Rating"), term, rate])
    return rows

@st.cache_resource
def _shared_rate_data():
    return {"master_grid": None, "master_cols": None,
            "special_rates": None, "saved_at": None}

if "query_results" not in st.session_state:
    st.session_state.query_results = None
    st.session_state.query_excel   = None
    st.session_state.rate_sheet_html = None
if "master_grid" not in st.session_state:
    st.session_state.master_grid = empty_master_df()
if "special_rates" not in st.session_state:
    st.session_state.special_rates = empty_special_rates_df()


@st.cache_data
def load_backend_lookup():
    if not os.path.exists(LOOKUP_PATH):
        return None
    df = pd.read_excel(LOOKUP_PATH)
    df.columns = [str(c).strip().lower() for c in df.columns]
    return build_lookup(df)


lookup = load_backend_lookup()

if lookup is None:
    st.error(
        "Institution lookup file not found. "
        "Place institution_lookup.xlsx in the app folder and restart."
    )

tab_data, tab1, tab2, tab3, tab4 = st.tabs([
    "Master Data", "Custom Query", "Rate Sheet Generator", "File Format Guide", "Admin"
])

with tab_data:
    # ── Save / load shared across all sessions ──────────────────────────────
    shared = _shared_rate_data()
    save_col, load_col, _ = st.columns([1, 2, 5])
    with save_col:
        if st.button("💾 Save for Team", help="Saves current master rates and special rates so any team member can load them."):
            shared["master_grid"]  = st.session_state.master_grid.to_dict(orient="records")
            shared["master_cols"]  = list(st.session_state.master_grid.columns)
            shared["special_rates"] = st.session_state.special_rates.to_dict(orient="records")
            shared["saved_at"]     = datetime.now(_VAN).strftime("%b %d, %Y at %I:%M %p")
            st.success("Saved! Team members can now click 'Use Last' to load this data.")
    with load_col:
        if shared.get("saved_at"):
            if st.button(f"⟳ Use Last  —  {shared['saved_at']}", help="Load the last saved master rates and special rates."):
                st.session_state.master_grid   = pd.DataFrame(
                    shared["master_grid"], columns=shared["master_cols"]
                ).astype(str).fillna("")
                st.session_state.special_rates = normalize_special_rates_df(
                    pd.DataFrame(shared["special_rates"], columns=SPECIAL_RATES_COLS)
                )
                st.rerun()

    st.markdown("---")
    st.caption(
        "Enter your master rates below. "
        "To paste from Excel or Google Sheets: copy your data (Ctrl+C / ⌘C), "
        "click the first cell in the **Issuer** column, then paste (Ctrl+V / ⌘V). "
        "Column order must match the headers shown."
    )
    hdr_col, btn_col = st.columns([8, 1])
    with btn_col:
        if st.button("Clear all"):
            st.session_state.master_grid = empty_master_df()
            st.rerun()

    edited = st.data_editor(
        st.session_state.master_grid,
        num_rows="dynamic",
        use_container_width=True,
        height=520,
        column_config={
            "Issuer":                                          st.column_config.TextColumn("Issuer",            width="large"),
            "Credit Rating/Insurance Coverage - Short Term":   st.column_config.TextColumn("CR Short Term",     width="medium"),
            "Credit Rating/Insurance Coverage - Long Term":    st.column_config.TextColumn("CR Long Term",      width="medium"),
            "Take FI Money":                                   st.column_config.TextColumn("Take FI Money",     width="small"),
            "Available":                                       st.column_config.TextColumn("Available",         width="small"),
            "Province":                                        st.column_config.TextColumn("Province",          width="small"),
            "Offers USD?":                                     st.column_config.TextColumn("Offers USD?",       width="small"),
            "As of date for Rates":                            st.column_config.TextColumn("As of Date",        width="medium"),
            "Cashable after 30 days":                          st.column_config.TextColumn("Cash. 30d",         width="medium"),
            "Cashable after 60 days":                          st.column_config.TextColumn("Cash. 60d",         width="medium"),
            "Cashable after 90 days":                          st.column_config.TextColumn("Cash. 90d",         width="medium"),
            "Cashable after 180 days":                         st.column_config.TextColumn("Cash. 180d",        width="medium"),
            "Cashable after 270 days":                         st.column_config.TextColumn("Cash. 270d",        width="medium"),
            "2 Year Cashable after 365 days":                  st.column_config.TextColumn("2Yr Cash. 365d",    width="medium"),
            "30 days":                                         st.column_config.TextColumn("30 days",           width="medium"),
            "60 days":                                         st.column_config.TextColumn("60 days",           width="medium"),
            "90 days":                                         st.column_config.TextColumn("90 days",           width="medium"),
            "120 days":                                        st.column_config.TextColumn("120 days",          width="medium"),
            "180 days":                                        st.column_config.TextColumn("180 days",          width="medium"),
            "270 days":                                        st.column_config.TextColumn("270 days",          width="medium"),
            "1 year fixed":                                    st.column_config.TextColumn("1 Yr Fixed",        width="medium"),
            "18 month fixed":                                  st.column_config.TextColumn("18 Mo Fixed",       width="medium"),
            "2 year fixed":                                    st.column_config.TextColumn("2 Yr Fixed",        width="medium"),
            "3 year fixed":                                    st.column_config.TextColumn("3 Yr Fixed",        width="medium"),
            "4 year fixed":                                    st.column_config.TextColumn("4 Yr Fixed",        width="medium"),
            "5 year fixed":                                    st.column_config.TextColumn("5 Yr Fixed",        width="medium"),
            "Room available":                                  st.column_config.TextColumn("Room Avail.",       width="medium"),
            "Notes":                                           st.column_config.TextColumn("Notes",             width="large"),
        },
    )
    st.session_state.master_grid = edited

    n = master_row_count()
    if n:
        st.caption(f"{n} institution{'s' if n != 1 else ''} entered.")

    st.markdown("---")
    st.subheader("Special Rates")
    st.caption("Manually enter special or off-market rates. These are included in all queries and rate sheet generation.")

    sp_hdr, sp_btn = st.columns([8, 1])
    with sp_btn:
        if st.button("Clear", key="clear_special"):
            st.session_state.special_rates = empty_special_rates_df()
            st.rerun()

    term_options_list = [t[0] for t in TERM_COLUMNS]

    # Pass session state directly — never a copy. A fresh object every rerun
    # makes Streamlit treat it as new input and reset pending edits.
    special_edited = st.data_editor(
        st.session_state.special_rates,
        num_rows="dynamic",
        use_container_width=True,
        height=320,
        column_config={
            "Issuer":        st.column_config.TextColumn("Issuer",         width="large"),
            "Credit Rating": st.column_config.TextColumn("Credit Rating",  width="large"),
            "Term":          st.column_config.SelectboxColumn(
                                 "Term", options=term_options_list, width="medium"
                             ),
            "Rate":          st.column_config.TextColumn("Rate",           width="small"),
        },
    )
    # Always write back unconditionally using the normaliser so Term cells
    # stay as None (not "") which SelectboxColumn requires for empty rows.
    st.session_state.special_rates = normalize_special_rates_df(special_edited)

    # Auto-save completed rows to the issuer history database
    for _, row in st.session_state.special_rates.iterrows():
        issuer = str(row.get("Issuer", "") or "").strip()
        term   = str(row.get("Term",   "") or "").strip()
        rate   = str(row.get("Rate",   "") or "").strip()
        if issuer and term and rate and issuer not in _BLANK and term not in _BLANK:
            save_rate_to_history(
                issuer,
                str(row.get("Credit Rating", "") or "").strip(),
                term, rate,
            )

    # ── Issuer Database ──────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Issuer Database")
    st.caption(
        "Search any issuer to see its credit rating from the institution lookup "
        "and any historically saved special rates. Click **Add** to drop a row "
        "straight into the Special Rates grid above."
    )

    store = _rate_history_store()
    history_names  = {v["display_name"] for v in store.values() if "display_name" in v}
    lookup_names   = {info["display_name"] for info in (lookup or {}).values()}
    all_known      = sorted(history_names | lookup_names)

    db_search = st.text_input(
        "Search issuer",
        placeholder="Start typing an issuer name…",
        key="db_search",
        label_visibility="collapsed",
    )

    if db_search:
        q = db_search.strip().lower()
        matches = [n for n in all_known if q in n.lower()][:8]

        if not matches:
            st.caption("No known issuers match that name. Enter rates manually in the grid above.")
        else:
            for name in matches:
                norm        = normalize_name(name)
                lookup_info = (lookup or {}).get(norm)
                hist_key    = name.lower()
                hist_entries = store.get(hist_key, {}).get("entries", [])

                # Build the credit rating string from lookup
                def _build_rating(term_type):
                    if not lookup_info:
                        return ""
                    r   = lookup_info.get(f"{term_type}_rating", "")
                    ins = lookup_info.get("insurance", "")
                    if r and ins:  return f"{r} – {ins}"
                    return r or ins or ""

                long_rating  = _build_rating("long")
                short_rating = _build_rating("short")

                with st.expander(f"**{name}**", expanded=(len(matches) == 1)):
                    if lookup_info:
                        if long_rating:
                            st.markdown(f"**Long-term rating:** {long_rating}")
                        if short_rating:
                            st.markdown(f"**Short-term rating:** {short_rating}")

                    if hist_entries:
                        st.markdown("**Previous special rates:**")
                        hdr = st.columns([3, 2, 1, 1])
                        hdr[0].caption("Credit Rating")
                        hdr[1].caption("Term")
                        hdr[2].caption("Rate")
                        for i, entry in enumerate(hist_entries[:8]):
                            c1, c2, c3, c4 = st.columns([3, 2, 1, 1])
                            c1.write(entry.get("credit_rating") or "—")
                            c2.write(entry.get("term", ""))
                            c3.write(entry.get("rate", ""))
                            if c4.button("Add", key=f"db_add_{name}_{i}"):
                                new_row = pd.DataFrame([{
                                    "Issuer":        name,
                                    "Credit Rating": entry.get("credit_rating", ""),
                                    "Term":          entry.get("term", ""),
                                    "Rate":          entry.get("rate", ""),
                                }])
                                existing = st.session_state.special_rates[
                                    st.session_state.special_rates["Issuer"].astype(str).str.strip().ne("")
                                ]
                                st.session_state.special_rates = (
                                    normalize_special_rates_df(
                                        pd.concat([existing, new_row], ignore_index=True)
                                    )
                                )
                                st.rerun()
                    elif not lookup_info:
                        st.caption("No history yet — add a rate manually in the grid above and it will appear here next time.")

                    # Quick-add from lookup rating
                    if lookup_info and (long_rating or short_rating):
                        st.markdown("**Add from lookup:**")
                        qa_col1, qa_col2, qa_col3 = st.columns([2, 1, 1])
                        with qa_col1:
                            qa_rating = st.selectbox(
                                "Rating",
                                options=[r for r in [long_rating, short_rating] if r],
                                key=f"qa_rating_{name}",
                            )
                        with qa_col2:
                            qa_term = st.selectbox(
                                "Term",
                                options=[t[0] for t in TERM_COLUMNS],
                                key=f"qa_term_{name}",
                            )
                        with qa_col3:
                            qa_rate = st.text_input("Rate", key=f"qa_rate_{name}", placeholder="e.g. 3.75%")
                        if st.button("Add to Special Rates", key=f"qa_add_{name}"):
                            if qa_rate.strip():
                                new_row = pd.DataFrame([{
                                    "Issuer":        name,
                                    "Credit Rating": qa_rating,
                                    "Term":          qa_term,
                                    "Rate":          qa_rate.strip(),
                                }])
                                existing = st.session_state.special_rates[
                                    st.session_state.special_rates["Issuer"].astype(str).str.strip().ne("")
                                ]
                                st.session_state.special_rates = (
                                    normalize_special_rates_df(
                                        pd.concat([existing, new_row], ignore_index=True)
                                    )
                                )
                                save_rate_to_history(name, qa_rating, qa_term, qa_rate.strip())
                                st.rerun()
                            else:
                                st.warning("Enter a rate first.")


with tab1:
    st.write("Filter rates by term, credit rating status, and number of results.")

    query_source = st.radio(
        "Data source",
        ["Master Data", "Formatted Rate Sheet"],
        horizontal=True
    )

    if query_source == "Master Data":
        n = master_row_count()
        if n:
            st.info(f"{n} institutions loaded from Master Data tab.")
        else:
            st.warning("No data entered yet — go to the **Master Data** tab and paste your rates.")
    else:
        formatted_sheet_file = st.file_uploader(
            "Upload Formatted Rate Sheet",
            type=["xlsx"],
            key="formatted_sheet_uploader"
        )

    term_options = [t[0] for t in TERM_COLUMNS]
    selected_terms = st.multiselect(
        "Select terms",
        options=term_options,
        default=[]
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        top_n = st.number_input(
            "Top N rates per term",
            min_value=1,
            max_value=50,
            value=3
        )

    with col2:
        min_rate_pct = st.number_input(
            "Minimum rate (%)",
            min_value=0.0,
            max_value=20.0,
            value=0.0,
            step=0.05,
            format="%.2f"
        )

    with col3:
        institution_search = st.text_input(
            "Institution name contains",
            placeholder="e.g. Royal Bank"
        )

    col4, col5, col6 = st.columns(3)

    with col4:
        credit_rated_only = st.checkbox(
            "Credit rated only",
            value=False,
            help="Only show institutions with a formal credit rating (R-1, R-2, AA, BBB, etc.)"
        )

    with col5:
        exclude_cannot_source = st.checkbox(
            "Exclude blank ratings",
            value=False,
            help="Hide rows where no rating or insurance information is available"
        )

    with col6:
        sort_by = st.radio(
            "Sort by",
            options=["rate", "credit"],
            format_func=lambda x: "Highest Rate" if x == "rate" else "Credit Rating",
            horizontal=True
        )

    insurance_filter = st.radio(
        "Insurance filter",
        options=["any", "insured", "cdic", "provincial", "none"],
        format_func=lambda x: {
            "any": "Any", "insured": "Any insured", "cdic": "CDIC only",
            "provincial": "Provincial only", "none": "Uninsured only"
        }[x],
        horizontal=True
    )

    if st.button("Run Query"):
        if not selected_terms:
            st.warning("Please select at least one term.")
        elif query_source == "Master Data" and master_row_count() == 0:
            st.error("No data entered — go to the Master Data tab and paste your rates first.")
        elif query_source == "Formatted Rate Sheet" and not formatted_sheet_file:
            st.error("Please upload a Formatted Rate Sheet.")
        else:
            if query_source == "Formatted Rate Sheet":
                results = query_from_sheet(
                    formatted_sheet_file,
                    selected_terms,
                    int(top_n),
                    credit_rated_only
                )
                log_event("sheet_query")
            else:
                results = generate_custom_query(
                    get_master_file(),
                    lookup,
                    selected_terms,
                    int(top_n),
                    credit_rated_only
                )
                log_event("master_query")

            results = results + get_special_rate_rows(selected_terms)

            results = apply_query_filters(
                results,
                min_rate_pct / 100,
                insurance_filter,
                institution_search.strip(),
                exclude_cannot_source,
                sort_by,
            )

            if not results:
                st.session_state.query_results = None
                st.session_state.query_excel = None
                st.info("No results matched your query.")
            else:
                st.session_state.query_results = results
                st.session_state.query_excel = create_excel(results)

    if st.session_state.query_results:
        df_display = pd.DataFrame(
            st.session_state.query_results,
            columns=["Issuer", "Credit Rating & Guarantee", "Term", "Rate"]
        )
        df_display["Rate"] = df_display["Rate"].apply(lambda x: f"{x * 100:.2f}%")

        st.dataframe(df_display, width="stretch", hide_index=True)
        _copy_button_component(
            build_copy_html(st.session_state.query_results),
            "Copy to Clipboard",
        )

with tab2:
    st.write("Generate a formatted rate sheet from the data entered in the Master Data tab.")

    n = master_row_count()
    n_special = len(get_special_rate_rows())
    if n:
        st.info(f"{n} institutions loaded from Master Data tab." +
                (f"  ·  {n_special} special rate{'s' if n_special != 1 else ''}." if n_special else ""))
    elif n_special:
        st.info(f"{n_special} special rate{'s' if n_special != 1 else ''} entered.")
    else:
        st.warning("No data entered yet — go to the **Master Data** tab and add rates.")

    special_preview = get_special_rate_rows()
    if special_preview:
        st.info(f"{len(special_preview)} special rate{'s' if len(special_preview) != 1 else ''} will be included: " +
                ", ".join(f"{r[0]} ({r[2]})" for r in special_preview))

    if st.button("Generate Formatted Rate Sheet"):
        has_master  = master_row_count() > 0
        has_special = len(get_special_rate_rows()) > 0
        if not has_master and not has_special:
            st.error("No data entered — add master rates or special rates in the Master Data tab first.")
        else:
            output = generate_report(get_master_file(), lookup) if has_master else []
            output = output + get_special_rate_rows()
            # Merge into correct order: TERM_COLUMNS sequence, rate desc within each term
            from collections import defaultdict
            _term_rank = {t[0]: i for i, t in enumerate(TERM_COLUMNS)}
            _groups = defaultdict(list)
            for row in output:
                _groups[row[2]].append(row)
            output = []
            for tc in TERM_COLUMNS:
                if tc[0] in _groups:
                    output.extend(sorted(_groups.pop(tc[0]), key=lambda r: r[3], reverse=True))
            for rows in _groups.values():  # any terms not in TERM_COLUMNS
                output.extend(sorted(rows, key=lambda r: r[3], reverse=True))
            st.session_state.rate_sheet_html = build_copy_html(output)
            log_event("rate_sheet")

    if st.session_state.get("rate_sheet_html"):
        st.success(f"Rate sheet ready — {len([r for r in st.session_state.rate_sheet_html.split('<tr>') if '<td' in r])} rows.")
        _copy_button_component(st.session_state.rate_sheet_html, "Copy Rate Sheet to Clipboard")

with tab3:
    st.subheader("Master Rates File — Required Format")
    st.write(
        "The file must be **.xlsx**. Column names are case-insensitive. "
        "One row per institution. Columns marked **Used** are read by the app; others are accepted but ignored."
    )

    master_cols = pd.DataFrame([
        {"Column Name": "Issuer",                                  "Used by App": "Yes", "Description": "Name of the institution (must be in column A).", "Example": "Royal Bank of Canada"},
        {"Column Name": "Insurance/Credit Rating Short Term",      "Used by App": "No",  "Description": "For reference only — ratings are managed in the backend lookup.", "Example": "R-1 (High)"},
        {"Column Name": "Insurance/Credit Rating Long Term",       "Used by App": "No",  "Description": "For reference only — ratings are managed in the backend lookup.", "Example": "AA"},
        {"Column Name": "Take FI money?",                          "Used by App": "No",  "Description": "Internal reference field.", "Example": "Yes"},
        {"Column Name": "Available",                               "Used by App": "Yes", "Description": "Must say 'available' for the row to be included. Any other value skips it.", "Example": "available"},
        {"Column Name": "Province",                                "Used by App": "No",  "Description": "Internal reference field.", "Example": "ON"},
        {"Column Name": "Offers USD",                              "Used by App": "No",  "Description": "Internal reference field.", "Example": "Yes"},
        {"Column Name": "As of Date",                              "Used by App": "No",  "Description": "Date the rates were last updated.", "Example": "2024-01-15"},
        {"Column Name": "Cashable After 30 Days",                  "Used by App": "Yes", "Description": "Rate for 1 Year Cashable After 30 Days term.", "Example": "3.05%"},
        {"Column Name": "Cashable After 90 Days",                  "Used by App": "Yes", "Description": "Rate for 1 Year Cashable After 90 Days term.", "Example": "3.10%"},
        {"Column Name": "30 Days",                                 "Used by App": "Yes", "Description": "Rate for 30 Day Fixed term.", "Example": "2.80%"},
        {"Column Name": "60 Days",                                 "Used by App": "Yes", "Description": "Rate for 60 Day Fixed term.", "Example": "2.90%"},
        {"Column Name": "90 Days",                                 "Used by App": "Yes", "Description": "Rate for 90 Day Fixed term.", "Example": "3.00%"},
        {"Column Name": "180 Days",                                "Used by App": "Yes", "Description": "Rate for 180 Day Fixed term.", "Example": "3.10%"},
        {"Column Name": "270 Days",                                "Used by App": "Yes", "Description": "Rate for 270 Day Fixed term.", "Example": "3.15%"},
        {"Column Name": "1 Year Fixed",                            "Used by App": "Yes", "Description": "Rate for 1 Year Fixed term.", "Example": "3.20%"},
        {"Column Name": "18 Month Fixed",                          "Used by App": "Yes", "Description": "Rate for 18 Month Fixed term.", "Example": "3.30%"},
        {"Column Name": "2 Year Fixed",                            "Used by App": "Yes", "Description": "Rate for 2 Year Fixed term.", "Example": "3.40%"},
        {"Column Name": "3 Year Fixed",                            "Used by App": "Yes", "Description": "Rate for 3 Year Fixed term.", "Example": "3.50%"},
        {"Column Name": "4 Year Fixed",                            "Used by App": "Yes", "Description": "Rate for 4 Year Fixed term.", "Example": "3.60%"},
        {"Column Name": "5 Year Fixed",                            "Used by App": "Yes", "Description": "Rate for 5 Year Fixed term.", "Example": "3.75%"},
    ])

    st.dataframe(master_cols.set_index("Column Name").T, width="stretch")
    st.caption("Rates can be entered as percentages (3.75%) or decimals (0.0375). Blank cells or rates below 1% are ignored.")

with tab4:
    if not st.session_state.admin_authenticated:
        _, acol, _ = st.columns([1, 1.2, 1])
        with acol:
            st.subheader("Admin Access")
            admin_pw_input = st.text_input("Admin password", type="password", key="admin_pw_input")
            if st.button("Enter", key="admin_login"):
                if admin_pw_input == load_passwords()["admin_password"]:
                    st.session_state.admin_authenticated = True
                    st.rerun()
                else:
                    st.error("Incorrect admin password.")
    else:
        st.subheader("Admin Panel")

        if st.button("Log Out of Admin", key="admin_logout"):
            st.session_state.admin_authenticated = False
            st.rerun()

        st.markdown("---")
        st.markdown("#### Change Passwords")
        pc1, pc2 = st.columns(2)
        with pc1:
            new_app_pw = st.text_input("New app password", type="password", key="new_app_pw")
        with pc2:
            new_admin_pw = st.text_input("New admin password", type="password", key="new_admin_pw")
        if st.button("Save Passwords"):
            if not new_app_pw and not new_admin_pw:
                st.warning("Enter at least one new password.")
            else:
                current = load_passwords()
                save_passwords(
                    new_app_pw if new_app_pw else current["app_password"],
                    new_admin_pw if new_admin_pw else current["admin_password"]
                )
                st.success("Passwords updated for this session. To make permanent on Streamlit Cloud, also update secrets in the dashboard.")

        st.markdown("---")
        st.markdown("#### Update Institution Lookup File")
        new_lookup = st.file_uploader("Upload new institution_lookup.xlsx", type=["xlsx"], key="admin_lookup")
        if st.button("Save Lookup File"):
            if not new_lookup:
                st.warning("Please upload a file first.")
            else:
                with open(LOOKUP_PATH, "wb") as f:
                    f.write(new_lookup.read())
                st.cache_data.clear()
                st.success("Lookup file updated. To make permanent, push the new file to GitHub.")

        st.markdown("---")
        st.markdown("#### Usage Statistics")
        stats = load_stats()

        m1, m2, m3 = st.columns(3)
        m1.metric("Rate Sheets Generated", stats.get("total_rate_sheets", 0))
        m2.metric("Custom Queries Run", stats.get("total_queries", 0))
        m3.metric("Total Actions", stats.get("total_rate_sheets", 0) + stats.get("total_queries", 0))

        events = stats.get("events", [])
        if events:
            st.markdown("**Recent Activity**")
            df_events = pd.DataFrame(events[::-1][:50])
            df_events["Type"] = df_events["Type"].map({
                "rate_sheet":   "Rate Sheet Generated",
                "master_query": "Custom Query — Master File",
                "sheet_query":  "Custom Query — Formatted Sheet",
            }).fillna(df_events["Type"])
            st.dataframe(df_events, width="stretch", hide_index=True)
        else:
            st.info("No activity recorded yet.")

        # ── Table Format Editor ───────────────────────────────────────────────
        st.markdown("---")
        st.markdown("#### Table Format")
        st.caption(
            "Adjust every aspect of the table appearance below. "
            "The preview updates live as you make changes. "
            "Click **Save Format** when it looks exactly right."
        )

        s = load_table_style()

        def _fi(label, key, min_v, max_v, default):
            return st.number_input(label, min_v, max_v, s.get(key, default),
                                   step=1, key=f"ts_{key}")
        def _cp(label, key):
            return st.color_picker(label, s.get(key, DEFAULT_TABLE_STYLE[key]),
                                   key=f"ts_{key}")
        def _sel(label, key):
            opts = FONT_OPTIONS
            val  = s.get(key, "Calibri")
            idx  = opts.index(val) if val in opts else 0
            return st.selectbox(label, opts, index=idx, key=f"ts_{key}")
        def _cb(label, key):
            return st.checkbox(label, s.get(key, True), key=f"ts_{key}")

        ctrl_col, prev_col = st.columns([1, 1.6])

        with ctrl_col:
            st.markdown("**Header row**")
            h_bg    = _cp("Background colour",  "header_bg")
            h_text  = _cp("Text colour",         "header_text")
            h_font  = _sel("Font",               "header_font")
            h_size  = _fi("Font size (pt)", "header_size", 6, 28, 11)
            h_bold  = _cb("Bold",                "header_bold")

            st.markdown("**Body rows**")
            b_font  = _sel("Font",               "body_font")
            b_size  = _fi("Font size (pt)", "body_size", 6, 28, 11)
            b_text  = _cp("Text colour",         "body_text")
            b_bg    = _cp("Row 1 background",    "body_bg")
            alt_bg  = _cp("Row 2 background (alternating)", "alt_row_bg")

            st.markdown("**Accents & borders**")
            rate_c  = _cp("Rate colour",         "rate_color")
            bord_c  = _cp("Border colour",       "border_color")
            bord_w  = _fi("Border width (px)",   "border_width", 0, 5, 1)
            cell_p  = _fi("Cell padding (px)",   "cell_padding", 2, 20, 6)

            new_style = {
                "header_bg": h_bg,   "header_text": h_text,
                "header_font": h_font, "header_size": int(h_size),
                "header_bold": h_bold,
                "body_font": b_font,  "body_size": int(b_size),
                "body_text": b_text,  "body_bg": b_bg, "alt_row_bg": alt_bg,
                "rate_color": rate_c,
                "border_color": bord_c, "border_width": int(bord_w),
                "cell_padding": int(cell_p),
            }

            if st.button("💾  Save Format", key="save_table_fmt"):
                save_table_style(new_style)
                st.success("Table format saved — all Copy buttons will now use this style.")

            if st.button("Reset to defaults", key="reset_table_fmt"):
                save_table_style(DEFAULT_TABLE_STYLE)
                st.rerun()

        with prev_col:
            st.markdown("**Live preview**")
            st.markdown(
                _style_preview_html(new_style),
                unsafe_allow_html=True,
            )

st.markdown("---")
st.caption(
    "This tool does not guarantee 100% accuracy. Rates and information should be independently verified before use. "
    "For any issues, please contact Henry Jagger at henryjagger@gmail.com."
)
