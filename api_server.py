"""
Public API server for Rate Sheet generation.
Deploy this separately from the Streamlit app.
"""

import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional, List
from datetime import datetime
from zoneinfo import ZoneInfo

# Import the API code
from api import (
    app as api_app,
    RateSheetRequest,
    RateSheetResponse,
    generate_report_from_data,
    generate_report_usd,
    sort_output,
    build_html_table,
    is_credit_or_guarantee,
)

_VAN = ZoneInfo("America/Vancouver")

# Enable CORS so Power Automate can call it from anywhere
api_app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Export the app for deployment
app = api_app

if __name__ == "__main__":
    import uvicorn

    # Get port from environment or default to 8001
    port = int(os.getenv("PORT", 8001))

    uvicorn.run(
        app,
        host="0.0.0.0",
        port=port
    )
