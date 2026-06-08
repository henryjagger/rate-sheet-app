# Rate Sheet Generator API Guide

This FastAPI application allows you to generate rate sheets programmatically, perfect for Power Automate automation.

## Getting Started

### 1. Start the API Server

```bash
python api.py
```

The API will be available at `http://localhost:8000`

### 2. Interactive Documentation

Once running, visit:
- **Swagger UI**: http://localhost:8000/docs
- **ReDoc**: http://localhost:8000/redoc

## Endpoints

### POST `/generate-rate-sheet`

Generate a formatted rate sheet from master data.

**Parameters:**
- `currency`: "CAD" or "USD"
- `output_type`: "all_in", "credit_only" (CAD only), or "email"
- `data`: Master data object
  - `currency`: "CAD" or "USD"
  - `institutions`: Array of institution objects
  - `special_rates`: Optional array of special rate entries

**Example Request:**

```json
{
  "currency": "CAD",
  "output_type": "all_in",
  "data": {
    "currency": "CAD",
    "institutions": [
      {
        "issuer": "Royal Bank of Canada",
        "available": "available",
        "rating": "R-1 (High) – CDIC",
        "1 Year Fixed": "3.75%",
        "2 Year Fixed": "3.80%",
        "take fi money": "yes"
      },
      {
        "issuer": "TD Bank",
        "available": "available",
        "rating": "R-1 (High) – CDIC",
        "1 Year Fixed": "3.70%",
        "2 Year Fixed": "3.75%",
        "take fi money": "yes"
      },
      {
        "issuer": "EQ Bank",
        "available": "available",
        "rating": "BBB (High) – CDIC",
        "1 Year Fixed": "4.05%",
        "2 Year Fixed": "4.10%",
        "take fi money": "no"
      }
    ]
  }
}
```

**Response:**

```json
{
  "success": true,
  "message": "Generated all_in rate sheet (CAD)",
  "html": "<table>...</table>",
  "error": null,
  "generated_at": "2026-06-08T12:30:45.123456-07:00"
}
```

### POST `/generate-email`

Generate an email-ready format combining CAD and USD GIC data.

**Example Request:**

```json
{
  "currency": "CAD",
  "output_type": "email",
  "data": {
    "currency": "CAD",
    "institutions": [...]
  }
}
```

**Response:** Returns HTML ready to paste into Outlook or email client.

### GET `/health`

Quick health check to verify the API is running.

```bash
curl http://localhost:8000/health
```

Response:
```json
{
  "status": "healthy",
  "timestamp": "2026-06-08T12:30:45.123456-07:00"
}
```

## Power Automate Integration

### Step 1: Start the API

Run the API in the background on your machine or server:
```bash
python api.py
```

### Step 2: Create a Power Automate Cloud Flow

1. Go to **Power Automate** > **Create** > **Cloud flow** > **Automated cloud flow**
2. Choose a trigger (e.g., "When an email arrives")
3. Add an action: **HTTP**

### Step 3: Configure the HTTP Action

Set up the HTTP action to call the API:

**Method:** POST

**URI:** `http://localhost:8000/generate-rate-sheet`

(Or your server's IP/domain if running remotely)

**Headers:**
```
Content-Type: application/json
```

**Body:**
```json
{
  "currency": "CAD",
  "output_type": "all_in",
  "data": {
    "currency": "CAD",
    "institutions": [
      {
        "issuer": "Royal Bank of Canada",
        "available": "available",
        "rating": "R-1 (High) – CDIC",
        "1 Year Fixed": "3.75%"
      }
    ]
  }
}
```

### Step 4: Process the Response

The API returns JSON with an `html` field containing the formatted rate sheet table.

You can:
- **Send in email:** Copy the `html` field to email body
- **Save to file:** Use "Create file" action with the HTML
- **Display:** Use the HTML in a Teams message or other output

### Example: Email Rate Sheet

1. In Power Automate, add **Send an email** action
2. In the email body, use `@body('HTTP').html` to insert the generated table

## Supported Institution Fields

All fields are optional except `issuer` and `available`:

### CAD Fields:
- `issuer` (required)
- `available` (required, must be "available")
- `rating` - Credit/insurance rating (e.g., "R-1 (High) – CDIC")
- `take fi money` - "yes" or "no" for FI-only filters
- Terms: `Cashable After 30 Days`, `Cashable After 90 Days`, `30 Days`, `60 Days`, `90 Days`, `180 Days`, `270 Days`, `1 Year Fixed`, `18 Month Fixed`, `2 Year Fixed`, `3 Year Fixed`, `4 Year Fixed`, `5 Year Fixed`

### USD Fields:
- `issuer` (required)
- `available` (required, must be "available")
- `DBRS` - DBRS rating
- `S&P` - S&P rating
- Terms: `Cashable after 30`, `Cashable after 90`, `Cashable after 180`, `30`, `60`, `90`, `120`, `180`, `270`, `1`, `18 months`, `2`

## Output Types

| Type | Description | Currency | Notes |
|------|-------------|----------|-------|
| `all_in` | All rates sorted by term then rate | CAD/USD | Every available rate |
| `credit_only` | Credit-rated & 100% guarantees only | CAD only | Filters out unrated institutions |
| `email` | Email-ready format | CAD/USD | Combines CAD + USD with placeholders |

## Rate Format

Rates can be entered as:
- Percentages: `3.75%`
- Decimals: `0.0375`

Both will be parsed correctly.

## Troubleshooting

**API won't start:**
- Ensure FastAPI is installed: `pip install fastapi uvicorn`
- Check if port 8000 is already in use: `lsof -i :8000`

**Power Automate can't connect:**
- If running locally, use `http://localhost:8000` (won't work on cloud flows)
- For cloud flows, deploy API to Azure, AWS, or similar service

**Rates not showing:**
- Ensure `available` field is exactly "available" (case-insensitive)
- Check that rates are ≥ 1% (0.01 as decimal)

## Deployment Options

### Local Machine
```bash
python api.py
```
Works for desktop/RPA flows, not cloud flows.

### Azure App Service
```bash
pip install -r requirements.txt
gunicorn api:app
```

### Docker
```dockerfile
FROM python:3.11
WORKDIR /app
COPY . .
RUN pip install -r requirements.txt
CMD ["python", "api.py"]
```

## API Rate Limits

Currently no rate limiting. Implement as needed for production use.

## Security Notes

- The API currently has no authentication
- For production use, add:
  - API key authentication
  - CORS restrictions
  - Rate limiting
  - Input validation

Example with API key:

```python
@app.post("/generate-rate-sheet")
async def generate_rate_sheet(request: RateSheetRequest, x_api_key: str = Header(None)):
    if x_api_key != os.getenv("API_KEY"):
        raise HTTPException(status_code=401, detail="Invalid API key")
    # ... rest of function
```

## Questions?

Contact: henryjagger@gmail.com
