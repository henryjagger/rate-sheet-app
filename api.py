"""
FastAPI-based REST API for Rate Sheet generation.
Allows Power Automate and other tools to generate rate sheets programmatically.
"""

import os
import json
import pandas as pd
from typing import Optional, List
from datetime import datetime
from zoneinfo import ZoneInfo
from io import BytesIO
from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import uvicorn
import re

_VAN = ZoneInfo("America/Vancouver")

# ── Configuration ──────────────────────────────────────────────────────────
PRIMARY_LOOKUP_PATH = "institution_lookup_primary.xlsx"
LOOKUP_PATH = "institution_lookup.xlsx"

# ── Core utility functions ─────────────────────────────────────────────────

def clean_text(value):
    """Clean text value."""
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()

def parse_rate(value):
    """Parse rate from percentage or decimal format."""
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
    """Rank credit rating for sorting."""
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

def is_credit_or_guarantee(rating):
    """Check if rating is a formal credit rating or 100% guarantee."""
    _CREDIT_KEYWORDS = ("R-1", "R-2", "AA", "BBB", "A (HIGH)", "A (MID)", "A (LOW)")
    upper = str(rating).upper()
    return any(k in upper for k in _CREDIT_KEYWORDS) or "100%" in upper

# ── Term definitions ──────────────────────────────────────────────────────
TERM_COLUMNS = [
    ("Cashable After 30 Days", "cashable after 30 days", "HISA"),
    ("Cashable After 90 Days", "cashable after 90 days", "HISA"),
    ("30 Days", "30 days", "GIC"),
    ("60 Days", "60 days", "GIC"),
    ("90 Days", "90 days", "GIC"),
    ("180 Days", "180 days", "GIC"),
    ("270 Days", "270 days", "GIC"),
    ("1 Year Fixed", "1 year fixed", "GIC"),
    ("18 Month Fixed", "18 month fixed", "GIC"),
    ("2 Year Fixed", "2 year fixed", "GIC"),
    ("3 Year Fixed", "3 year fixed", "GIC"),
    ("4 Year Fixed", "4 year fixed", "GIC"),
    ("5 Year Fixed", "5 year fixed", "GIC"),
]

MASTER_GRID_COLS_USD = [
    "Issuer",
    "Cashable after 30",
    "Cashable after 90",
    "Cashable after 180",
    "30",
    "60",
    "90",
    "120",
    "180",
    "270",
    "1",
    "18 months",
    "2",
    "Available",
    "As of date for Rates",
    "DBRS",
    "S&P"
]

# ── Request/Response Models ────────────────────────────────────────────────
class RateEntry(BaseModel):
    """Single rate entry from master data."""
    issuer: str
    term: str
    rate: float
    rating: Optional[str] = None

class MasterData(BaseModel):
    """Master data for rate sheet generation."""
    currency: str  # CAD or USD
    institutions: List[dict]  # List of dicts with institution data
    special_rates: Optional[List[RateEntry]] = None

class RateSheetRequest(BaseModel):
    """Request to generate a rate sheet."""
    currency: str = "CAD"  # CAD or USD
    data: MasterData
    output_type: str = "all_in"  # all_in, credit_only, email
    filter_terms: Optional[List[str]] = None
    top_n: Optional[int] = None
    min_rate: Optional[float] = None
    credit_rated_only: Optional[bool] = False

class RateSheetResponse(BaseModel):
    """Response from rate sheet generation."""
    success: bool
    message: str
    html: Optional[str] = None
    error: Optional[str] = None
    generated_at: str

# ── Core Generation Functions ──────────────────────────────────────────────

def generate_report_from_data(institutions_data, currency="CAD", fi_only=False):
    """
    Generate report from raw institution data.
    """
    output = []

    for inst in institutions_data:
        issuer_raw = inst.get("issuer", "")
        if not issuer_raw or not str(issuer_raw).strip():
            continue

        available = clean_text(inst.get("available", "")).lower()
        if available != "available":
            continue

        # For each term, extract rate
        for display_term, source_col, term_type in TERM_COLUMNS:
            # Try both lowercase and original case
            rate_value = inst.get(source_col) or inst.get(source_col.lower()) or inst.get(display_term) or inst.get(display_term.lower())
            if not rate_value:
                continue

            rate = parse_rate(rate_value)
            if rate < 0.01:
                continue

            # Check FI-only filter if needed
            if fi_only:
                fi_val = clean_text(inst.get("take fi money", "")).lower()
                if fi_val not in ("yes", "y"):
                    continue

            # Get rating
            rating = inst.get("rating", inst.get("insurance/credit rating short term", "") or "")

            output.append([
                str(issuer_raw).strip(),
                str(rating).strip(),
                display_term,
                float(rate)
            ])

    return output

def generate_report_usd(institutions_data):
    """Generate report for USD data."""
    output = []

    for inst in institutions_data:
        issuer = inst.get("issuer", "").strip()
        if not issuer:
            continue

        available = clean_text(inst.get("available", "")).lower()
        if available != "available":
            continue

        # Extract all term columns
        for col_name in MASTER_GRID_COLS_USD[1:]:  # Skip Issuer
            if col_name in ["Available", "As of date for Rates", "DBRS", "S&P"]:
                continue

            rate_value = inst.get(col_name) or inst.get(col_name.lower())
            if not rate_value:
                continue

            rate = parse_rate(rate_value)
            if rate < 0.01:
                continue

            output.append([
                issuer,
                "",  # USD has no rating in output
                col_name,
                float(rate)
            ])

    return output

def sort_output(output):
    """Sort rows into TERM_COLUMNS order, rate descending within each term."""
    from collections import defaultdict
    groups = defaultdict(list)
    for row in output:
        groups[row[2]].append(row)
    result = []
    for tc in TERM_COLUMNS:
        if tc[0] in groups:
            result.extend(sorted(groups.pop(tc[0]), key=lambda r: r[3], reverse=True))
    for rows in groups.values():
        result.extend(sorted(rows, key=lambda r: r[3], reverse=True))
    return result

def build_html_table(rows):
    """Build a simple HTML table from rate rows."""
    html = '<table style="border-collapse:collapse;width:100%;font-family:Calibri,sans-serif;font-size:11pt;">'
    html += '<thead><tr style="background-color:#000;color:#fff;"><th style="border:1px solid #ccc;padding:6px;">Issuer</th><th style="border:1px solid #ccc;padding:6px;">Rating</th><th style="border:1px solid #ccc;padding:6px;">Term</th><th style="border:1px solid #ccc;padding:6px;color:#C00000;">Rate</th></tr></thead>'
    html += '<tbody>'
    for i, row in enumerate(rows):
        bg = "#fff" if i % 2 == 0 else "#f5f5f5"
        html += f'<tr style="background-color:{bg};"><td style="border:1px solid #ccc;padding:6px;">{row[0]}</td><td style="border:1px solid #ccc;padding:6px;">{row[1]}</td><td style="border:1px solid #ccc;padding:6px;">{row[2]}</td><td style="border:1px solid #ccc;padding:6px;color:#C00000;text-align:center;">{row[3]*100:.2f}%</td></tr>'
    html += '</tbody></table>'
    return html

# ── FastAPI App ────────────────────────────────────────────────────────────
app = FastAPI(
    title="Rate Sheet Generator API",
    description="REST API for generating custom rate sheets for Power Automate",
    version="1.0.0"
)

@app.get("/")
async def root():
    """Health check and API info."""
    return {
        "status": "ok",
        "name": "Rate Sheet Generator API",
        "endpoints": {
            "POST /generate-rate-sheet": "Generate all-in GIC rates",
            "GET /health": "Health check",
            "GET /docs": "Interactive API documentation"
        }
    }

@app.get("/health")
async def health():
    """Health check endpoint."""
    return {"status": "healthy", "timestamp": datetime.now(_VAN).isoformat()}

@app.post("/generate-rate-sheet", response_model=RateSheetResponse)
async def generate_rate_sheet(request: RateSheetRequest):
    """
    Generate a rate sheet from master data.

    Supports two output types:
    - all_in: Every rate sorted by term then rate
    - credit_only: Only credit-rated & 100% guarantees (CAD only)

    Example request:
    {
        "currency": "CAD",
        "data": {
            "currency": "CAD",
            "institutions": [
                {
                    "issuer": "Royal Bank of Canada",
                    "available": "available",
                    "rating": "R-1 (High)",
                    "1 Year Fixed": "3.75%"
                }
            ]
        },
        "output_type": "all_in"
    }
    """
    try:
        # Validate inputs
        if request.currency not in ["CAD", "USD"]:
            raise ValueError("Currency must be CAD or USD")

        if request.output_type not in ["all_in", "credit_only"]:
            raise ValueError("output_type must be: all_in or credit_only")

        if request.output_type == "credit_only" and request.currency == "USD":
            raise ValueError("credit_only is only supported for CAD")

        # Generate base report
        if request.currency == "CAD":
            base = generate_report_from_data(
                request.data.institutions,
                fi_only=False
            )
        else:
            base = generate_report_usd(request.data.institutions)

        # Add special rates if provided
        special = request.data.special_rates or []
        special_formatted = [
            [s.issuer, s.rating or "", s.term, s.rate]
            for s in special
        ]

        # Combine and sort
        output = sort_output(base + special_formatted)

        # Apply filters if credit_only
        if request.output_type == "credit_only":
            output = [r for r in output if is_credit_or_guarantee(r[1])]

        # Build HTML output
        html = build_html_table(output)

        return RateSheetResponse(
            success=True,
            message=f"Generated {request.output_type} rate sheet ({request.currency})",
            html=html,
            generated_at=datetime.now(_VAN).isoformat()
        )

    except Exception as e:
        return RateSheetResponse(
            success=False,
            message="Error generating rate sheet",
            error=str(e),
            generated_at=datetime.now(_VAN).isoformat()
        )

# ── Run server ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # Run: python api.py
    # Access at http://localhost:8001
    # Docs at http://localhost:8001/docs
    port = int(os.environ.get("PORT", 8001))
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=port
    )
