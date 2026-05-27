import os
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

LOOKUP_PATH = "institution_lookup.xlsx"


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
        fgColor="333333"
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


st.set_page_config(
    page_title="Rate Sheet Generator",
    layout="wide"
)

st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Times New Roman everywhere */
    html, body, [class*="css"], p, div, span, label, input,
    h1, h2, h3, h4, h5, h6, button, td, th, textarea, select {
        font-family: 'Times New Roman', Times, serif !important;
    }

    /* Tighter page padding */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 1rem !important;
    }

    /* Buttons */
    .stButton > button {
        background-color: #1A3A5C;
        color: white;
        border: none;
        border-radius: 6px;
        padding: 0.4rem 1.5rem;
        font-weight: 600;
        font-family: 'Times New Roman', Times, serif !important;
        transition: background-color 0.2s;
    }
    .stButton > button:hover {
        background-color: #142E4A;
        color: white;
    }

    /* Login card */
    .login-card {
        background: #FAF7F2;
        border-radius: 10px;
        padding: 2rem 2rem 1.5rem 2rem;
        box-shadow: 0 2px 12px rgba(0,0,0,0.08);
        margin-top: 2rem;
    }
    .login-title {
        color: #1A202C;
        font-family: 'Times New Roman', Times, serif;
        font-size: 1.5rem;
        font-weight: 700;
        margin-bottom: 0.2rem;
    }
    .login-sub {
        color: #718096;
        font-family: 'Times New Roman', Times, serif;
        font-size: 0.88rem;
        margin-bottom: 1.25rem;
    }

    /* Page header */
    .page-header {
        border-bottom: 2px solid #D9D0C4;
        padding-bottom: 0.6rem;
        margin-bottom: 1rem;
    }
    .page-header h1 {
        color: #1A202C;
        font-family: 'Times New Roman', Times, serif;
        margin: 0;
        font-size: 1.6rem;
        font-weight: 700;
    }
    .page-header p {
        color: #718096;
        font-family: 'Times New Roman', Times, serif;
        margin: 0.15rem 0 0 0;
        font-size: 0.88rem;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        color: #1A3A5C;
        font-weight: 700;
        border-bottom-color: #1A3A5C;
    }
</style>
""", unsafe_allow_html=True)

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    _, col, _ = st.columns([1, 1.2, 1])
    with col:
        st.markdown("""
        <div class="login-card">
            <div class="login-title">Rate Sheet Generator</div>
            <div class="login-sub">Enter your password to continue</div>
        </div>
        """, unsafe_allow_html=True)
        password = st.text_input("Password", type="password", label_visibility="collapsed",
                                 placeholder="Password")
        if st.button("Enter", use_container_width=True):
            if password == "Locarno":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password.")
    st.stop()

st.markdown("""
<div class="page-header">
    <h1>Rate Sheet Generator</h1>
    <p>Generate and query institutional rate sheets</p>
</div>
""", unsafe_allow_html=True)


if "query_results" not in st.session_state:
    st.session_state.query_results = None
    st.session_state.query_excel = None


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

tab1, tab2, tab3 = st.tabs(["Custom Query", "Rate Sheet Generator", "File Format Guide"])

with tab1:
    st.write("Filter rates by term, credit rating status, and number of results.")

    query_source = st.radio(
        "Data source",
        ["Master Rates file", "Formatted Rate Sheet"],
        horizontal=True
    )

    if query_source == "Master Rates file":
        query_master_file = st.file_uploader(
            "Upload Master Rates Excel file",
            type=["xlsx"],
            key="query_master_uploader"
        )
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

    col1, col2 = st.columns(2)

    with col1:
        top_n = st.number_input(
            "Top N rates per term",
            min_value=1,
            max_value=20,
            value=3
        )

    with col2:
        credit_rated_only = st.checkbox(
            "Credit rated only",
            value=False,
            help="Only show institutions with a formal credit rating (R-1, R-2, AA, BBB, etc.)"
        )

    if st.button("Run Query"):
        if not selected_terms:
            st.warning("Please select at least one term.")
        elif query_source == "Master Rates file" and not query_master_file:
            st.error("Please upload a Master Rates file.")
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
            else:
                results = generate_custom_query(
                    query_master_file,
                    lookup,
                    selected_terms,
                    int(top_n),
                    credit_rated_only
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

        st.download_button(
            label="Download Custom Query as Excel",
            data=st.session_state.query_excel,
            file_name="custom_query.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

with tab2:
    st.write("Upload your Master Rates file to generate the full formatted rate sheet.")

    master_file = st.file_uploader(
        "Upload Master Rates Excel file",
        type=["xlsx"],
        key="report_master_uploader"
    )

    if st.button("Generate Formatted Rate Sheet"):
        if not master_file:
            st.error("Please upload a Master Rates file.")
        else:
            output = generate_report(master_file, lookup)
            excel_file = create_excel(output)

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

st.markdown("---")
st.caption(
    "This tool does not guarantee 100% accuracy. Rates and information should be independently verified before use. "
    "For any issues, please contact Henry Jagger at henryjagger@gmail.com."
)
