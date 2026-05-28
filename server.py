import os
import uuid
from io import BytesIO
from flask import (
    Flask, render_template, request, redirect,
    url_for, session, send_file
)
import pandas as pd
from processing import (
    TERM_COLUMNS, load_passwords, save_passwords,
    load_stats, log_event, load_backend_lookup, clear_lookup_cache,
    generate_custom_query, query_from_sheet,
    generate_report, create_excel, linkify_insurance_html,
    LOOKUP_PATH,
)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "rsg-dev-secret-change-in-prod")

# In-process cache for generated Excel files keyed by a random UUID stored in the session.
# Fine for single-server deployments; replace with Redis/S3 for multi-instance.
_file_cache: dict[str, bytes] = {}

TERM_OPTIONS = [t[0] for t in TERM_COLUMNS]

# ── Helpers ──────────────────────────────────────────────────────────────────

def require_auth():
    if not session.get("authenticated"):
        return redirect(url_for("login"))

def require_admin():
    if not session.get("admin_authenticated"):
        return redirect(url_for("admin"))


# ── Auth ──────────────────────────────────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        pw = request.form.get("password", "")
        if pw == load_passwords()["app_password"]:
            session["authenticated"] = True
            return redirect(url_for("index"))
        error = "Incorrect password. Please try again."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ── Main app ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    if not session.get("authenticated"):
        return redirect(url_for("login"))
    return render_template("main.html",
                           active_tab="query",
                           term_options=TERM_OPTIONS)


# ── Custom Query ──────────────────────────────────────────────────────────────

@app.route("/query", methods=["POST"])
def query():
    if not session.get("authenticated"):
        return redirect(url_for("login"))

    source           = request.form.get("source", "master")
    selected_terms   = request.form.getlist("terms")
    top_n            = max(1, min(20, int(request.form.get("top_n", 3) or 3)))
    credit_rated_only = "credit_rated_only" in request.form

    def err(msg):
        return render_template("main.html", active_tab="query",
                               term_options=TERM_OPTIONS,
                               query_error=msg,
                               prev_source=source,
                               prev_terms=selected_terms,
                               prev_top_n=top_n,
                               prev_credit=credit_rated_only)

    if not selected_terms:
        return err("Please select at least one term.")

    lookup = load_backend_lookup()

    if source == "master":
        f = request.files.get("master_file")
        if not f or not f.filename:
            return err("Please upload a Master Rates file.")
        results = generate_custom_query(f, lookup, selected_terms, top_n, credit_rated_only)
        log_event("master_query")
    else:
        f = request.files.get("formatted_file")
        if not f or not f.filename:
            return err("Please upload a Formatted Rate Sheet.")
        results = query_from_sheet(f, selected_terms, top_n, credit_rated_only)
        log_event("sheet_query")

    if not results:
        return err("No results matched your query.")

    excel_bytes = create_excel(results).getvalue()
    cache_key = str(uuid.uuid4())
    _file_cache[cache_key] = excel_bytes
    session["query_cache_key"] = cache_key

    df = pd.DataFrame(results, columns=["Issuer", "Credit Rating & Guarantee", "Term", "Rate"])
    df["Rate"] = df["Rate"].apply(lambda x: f"{x * 100:.2f}%")

    # Build HTML table rows with rowspan on Term
    rows_list  = list(df.itertuples(index=False, name=None))
    headers    = df.columns.tolist()
    term_idx   = headers.index("Term")
    rowspans   = []
    i = 0
    while i < len(rows_list):
        span = 1
        while (i + span < len(rows_list) and
               rows_list[i + span][term_idx] == rows_list[i][term_idx]):
            span += 1
        rowspans.extend([span] + [0] * (span - 1))
        i += span

    table_rows = []
    for ri, row_vals in enumerate(rows_list):
        cells = []
        for ci, (col, val) in enumerate(zip(headers, row_vals)):
            if ci == term_idx:
                span = rowspans[ri]
                if span == 0:
                    continue
                cells.append({"type": "term", "val": val, "span": span})
            elif col == "Rate":
                cells.append({"type": "rate", "val": val})
            elif col == "Credit Rating & Guarantee":
                cells.append({"type": "rating", "val": linkify_insurance_html(str(val))})
            else:
                cells.append({"type": "plain", "val": val})
        table_rows.append(cells)

    return render_template("main.html",
                           active_tab="query",
                           term_options=TERM_OPTIONS,
                           query_results=table_rows,
                           query_headers=headers,
                           cache_key=cache_key,
                           prev_source=source,
                           prev_terms=selected_terms,
                           prev_top_n=top_n,
                           prev_credit=credit_rated_only)


@app.route("/download-query")
def download_query():
    if not session.get("authenticated"):
        return redirect(url_for("login"))
    key = session.get("query_cache_key")
    if not key or key not in _file_cache:
        return redirect(url_for("index"))
    return send_file(
        BytesIO(_file_cache[key]),
        download_name="custom_query.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Rate Sheet Generator ──────────────────────────────────────────────────────

@app.route("/generate", methods=["POST"])
def generate():
    if not session.get("authenticated"):
        return redirect(url_for("login"))

    f = request.files.get("master_file")
    if not f or not f.filename:
        return render_template("main.html", active_tab="generate",
                               term_options=TERM_OPTIONS,
                               gen_error="Please upload a Master Rates file.")

    lookup = load_backend_lookup()
    output = generate_report(f, lookup)
    buf    = create_excel(output)
    log_event("rate_sheet")

    return send_file(
        buf,
        download_name="formatted_rate_sheet.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Admin ─────────────────────────────────────────────────────────────────────

@app.route("/admin", methods=["GET"])
def admin():
    if not session.get("authenticated"):
        return redirect(url_for("login"))
    if not session.get("admin_authenticated"):
        return render_template("main.html", active_tab="admin",
                               term_options=TERM_OPTIONS,
                               show_admin_login=True)
    stats = load_stats()
    events = list(reversed(stats.get("events", [])[:50]))
    label_map = {
        "rate_sheet":   "Rate Sheet Generated",
        "master_query": "Custom Query — Master File",
        "sheet_query":  "Custom Query — Formatted Sheet",
    }
    for e in events:
        e["label"] = label_map.get(e.get("type", ""), e.get("type", ""))
    return render_template("main.html", active_tab="admin",
                           term_options=TERM_OPTIONS,
                           stats=stats, events=events)


@app.route("/admin/login", methods=["POST"])
def admin_login():
    if not session.get("authenticated"):
        return redirect(url_for("login"))
    pw = request.form.get("admin_password", "")
    if pw == load_passwords()["admin_password"]:
        session["admin_authenticated"] = True
        return redirect(url_for("admin"))
    return render_template("main.html", active_tab="admin",
                           term_options=TERM_OPTIONS,
                           show_admin_login=True,
                           admin_error="Incorrect admin password.")


@app.route("/admin/logout", methods=["POST"])
def admin_logout():
    session.pop("admin_authenticated", None)
    return redirect(url_for("index"))


@app.route("/admin/save-passwords", methods=["POST"])
def save_passwords_route():
    if not session.get("admin_authenticated"):
        return redirect(url_for("login"))
    new_app   = request.form.get("new_app_pw", "").strip()
    new_admin = request.form.get("new_admin_pw", "").strip()
    current   = load_passwords()
    save_passwords(
        new_app   if new_app   else current["app_password"],
        new_admin if new_admin else current["admin_password"],
    )
    return redirect(url_for("admin") + "?msg=passwords_saved")


@app.route("/admin/upload-lookup", methods=["POST"])
def upload_lookup():
    if not session.get("admin_authenticated"):
        return redirect(url_for("login"))
    f = request.files.get("lookup_file")
    if f and f.filename:
        f.save(LOOKUP_PATH)
        clear_lookup_cache()
    return redirect(url_for("admin") + "?msg=lookup_saved")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
