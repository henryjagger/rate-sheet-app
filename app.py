import os
import json
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from io import BytesIO, StringIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

LOOKUP_PATH = "institution_lookup.xlsx"

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
    ("5 Year Fixed", "5 year fixed", "long"),
    ("4 Year Fixed", "4 year fixed", "long"),
    ("3 Year Fixed", "3 year fixed", "long"),
    ("2 Year Fixed", "2 year fixed", "long"),
    ("18 Month Fixed", "18 month fixed", "long"),
    ("1 Year Fixed", "1 year fixed", "short"),
    ("270 Day Fixed", "270 days", "short"),
    ("180 Day Fixed", "180 days", "short"),
    ("90 Day Fixed", "90 days", "short"),
    ("60 Day Fixed", "60 days", "short"),
    ("30 Day Fixed", "30 days", "short"),
    ("1 Year Cashable After 90 Days", "cashable after 90 days", "short"),
    ("1 Year Cashable After 30 Days", "cashable after 30 days", "short"),
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

    text = str(value).replace("%", "").strip()

    try:
        num = float(text)
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

    if sort_by == "credit":
        from collections import defaultdict
        groups = defaultdict(list)
        for row in filtered:
            groups[row[2]].append(row)
        seen_terms, ordered = [], []
        for row in filtered:
            if row[2] not in seen_terms:
                seen_terms.append(row[2])
        for term in seen_terms:
            groups[term].sort(key=lambda r: (credit_rank(r[1]), r[3]), reverse=True)
            ordered.extend(groups[term])
        filtered = ordered

    return filtered


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
    "Issuer", "Available",
    "Cashable After 30 Days", "Cashable After 90 Days",
    "30 Days", "60 Days", "90 Days", "180 Days", "270 Days",
    "1 Year Fixed", "18 Month Fixed",
    "2 Year Fixed", "3 Year Fixed", "4 Year Fixed", "5 Year Fixed",
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

if "query_results" not in st.session_state:
    st.session_state.query_results = None
    st.session_state.query_excel = None
if "master_grid" not in st.session_state:
    st.session_state.master_grid = empty_master_df()


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
            "Issuer":                  st.column_config.TextColumn("Issuer",            width="large"),
            "Available":               st.column_config.TextColumn("Available",         width="small"),
            "Cashable After 30 Days":  st.column_config.TextColumn("Cash. 30 Days",     width="medium"),
            "Cashable After 90 Days":  st.column_config.TextColumn("Cash. 90 Days",     width="medium"),
            "30 Days":                 st.column_config.TextColumn("30 Days",            width="medium"),
            "60 Days":                 st.column_config.TextColumn("60 Days",            width="medium"),
            "90 Days":                 st.column_config.TextColumn("90 Days",            width="medium"),
            "180 Days":                st.column_config.TextColumn("180 Days",           width="medium"),
            "270 Days":                st.column_config.TextColumn("270 Days",           width="medium"),
            "1 Year Fixed":            st.column_config.TextColumn("1 Yr Fixed",         width="medium"),
            "18 Month Fixed":          st.column_config.TextColumn("18 Mo Fixed",        width="medium"),
            "2 Year Fixed":            st.column_config.TextColumn("2 Yr Fixed",         width="medium"),
            "3 Year Fixed":            st.column_config.TextColumn("3 Yr Fixed",         width="medium"),
            "4 Year Fixed":            st.column_config.TextColumn("4 Yr Fixed",         width="medium"),
            "5 Year Fixed":            st.column_config.TextColumn("5 Yr Fixed",         width="medium"),
        },
    )
    st.session_state.master_grid = edited

    n = master_row_count()
    if n:
        st.caption(f"{n} institution{'s' if n != 1 else ''} entered.")


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

        btn_col1, btn_col2 = st.columns([1, 1])

        with btn_col1:
            st.download_button(
                label="Download as Excel",
                data=st.session_state.query_excel,
                file_name="custom_query.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with btn_col2:
            # Build styled HTML table for rich paste into Excel/Word/Outlook
            # <font> tags used inside headers so Outlook's Word renderer shows white text
            headers = df_display.columns.tolist()
            th_style = "border:1px solid #ccc;padding:6px 12px;background-color:#000000;font-family:Calibri,sans-serif;font-size:11pt;"
            td_style = "border:1px solid #ccc;padding:6px 12px;font-family:Calibri,sans-serif;font-size:11pt;text-align:center;"
            td_rate_style = td_style + "color:#C00000;"
            header_row = "<tr>" + "".join(
                f"<th style='{th_style}' bgcolor='#000000'><font color='#ffffff' face='Calibri'><b>{h}</b></font></th>"
                for h in headers
            ) + "</tr>"
            # Compute rowspans for the Term column so consecutive identical terms merge
            rows_list = list(df_display.itertuples(index=False, name=None))
            term_col_idx = headers.index("Term") if "Term" in headers else None
            rowspans = []
            if term_col_idx is not None:
                i = 0
                while i < len(rows_list):
                    span = 1
                    while i + span < len(rows_list) and rows_list[i + span][term_col_idx] == rows_list[i][term_col_idx]:
                        span += 1
                    rowspans.extend([span] + [0] * (span - 1))
                    i += span
            else:
                rowspans = [1] * len(rows_list)

            td_merge_style = td_style + "vertical-align:middle;"
            data_rows = ""
            for row_idx, row_vals in enumerate(rows_list):
                cells = ""
                for col_idx, (col, val) in enumerate(zip(headers, row_vals)):
                    if term_col_idx is not None and col_idx == term_col_idx:
                        span = rowspans[row_idx]
                        if span == 0:
                            continue
                        span_attr = f" rowspan='{span}'" if span > 1 else ""
                        cells += f"<td{span_attr} style='{td_merge_style}'><font face='Calibri'>{val}</font></td>"
                    elif col == "Rate":
                        cells += f"<td style='{td_rate_style}'><font face='Calibri' color='#C00000'>{val}</font></td>"
                    elif col == "Credit Rating & Guarantee":
                        linked = linkify_insurance_html(str(val))
                        cells += f"<td style='{td_style}'><font face='Calibri'>{linked}</font></td>"
                    else:
                        cells += f"<td style='{td_style}'><font face='Calibri'>{val}</font></td>"
                data_rows += f"<tr>{cells}</tr>"
            html_table = f"<table style='border-collapse:collapse;'>{header_row}{data_rows}</table>"

            def js_escape(s):
                return s.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")

            html_js = js_escape(html_table)
            tsv_js = js_escape(df_display.to_csv(sep="\t", index=False))

            components.html(f"""
            <button onclick="(async () => {{
                try {{
                    const html = `{html_js}`;
                    const plain = `{tsv_js}`;
                    const item = new ClipboardItem({{
                        'text/html':  new Blob([html],  {{type: 'text/html'}}),
                        'text/plain': new Blob([plain], {{type: 'text/plain'}})
                    }});
                    await navigator.clipboard.write([item]);
                }} catch(e) {{
                    await navigator.clipboard.writeText(`{tsv_js}`);
                }}
                this.textContent = '✓ Copied!';
                setTimeout(() => this.textContent = 'Copy to Clipboard', 2000);
            }})()" style="
                background: transparent;
                color: #111111;
                border: 1px solid rgba(0,0,0,0.25);
                border-radius: 1px;
                padding: 0 16px;
                font-size: 11px;
                font-family: 'Inter', sans-serif;
                font-weight: 600;
                text-transform: uppercase;
                letter-spacing: 0.14em;
                cursor: pointer;
                height: 38px;
                width: 100%;
                transition: all 0.22s ease;
            " onmouseover="this.style.background='#111111';this.style.color='#ffffff';"
               onmouseout="this.style.background='transparent';this.style.color='#111111';"
            >Copy to Clipboard</button>
            """, height=50)

with tab2:
    st.write("Generate a formatted rate sheet from the data entered in the Master Data tab.")

    n = master_row_count()
    if n:
        st.info(f"{n} institutions loaded from Master Data tab.")
    else:
        st.warning("No data entered yet — go to the **Master Data** tab and paste your rates.")

    if st.button("Generate Formatted Rate Sheet"):
        if master_row_count() == 0:
            st.error("No data entered — go to the Master Data tab and paste your rates first.")
        else:
            output = generate_report(get_master_file(), lookup)
            excel_file = create_excel(output)
            log_event("rate_sheet")

            st.success("Formatted rate sheet generated.")

            st.download_button(
                label="Download Formatted Rate Sheet",
                data=excel_file,
                file_name="formatted_rate_sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

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

st.markdown("---")
st.caption(
    "This tool does not guarantee 100% accuracy. Rates and information should be independently verified before use. "
    "For any issues, please contact Henry Jagger at henryjagger@gmail.com."
)
