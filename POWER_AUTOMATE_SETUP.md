# Power Automate Integration Guide

Your Rate Sheet API is ready to use with Power Automate! Here's how to set it up.

## Quick Start

### 1. Start the API Server

```bash
python api.py
```

This starts the server at `http://localhost:8000`

### 2. Test in Browser

Visit **http://localhost:8000/docs** to see the interactive API documentation with live testing.

## Power Automate Setup

### For Desktop/RPA Flows (Local Machine)

If you're using Power Automate Desktop or a Windows RPA flow on your machine:

**1. Create a Desktop Flow**
- Open Power Automate Desktop
- Create a new desktop flow
- Add action: **Invoke Power Automate action**

**2. Use the HTTP action**

In your Power Automate Desktop flow, use **Invoke Web API** or **HTTP Request**:

```
Method: POST
URI: http://localhost:8000/generate-rate-sheet
Headers:
  Content-Type: application/json

Body:
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
        "2 Year Fixed": "3.80%"
      }
    ]
  }
}
```

### For Cloud Flows (Azure/Online)

If you want to use cloud-based Power Automate flows, you need to deploy the API to a cloud service:

**Option A: Azure App Service** (Recommended)
```bash
# Create an app service and deploy the API
az webapp up --name rate-sheet-api --resource-group my-group
```

**Option B: Docker Container**
```bash
docker build -t rate-sheet-api .
docker run -p 8000:8000 rate-sheet-api
```

Once deployed, use the cloud URL in your Power Automate flow:
```
http://your-api-name.azurewebsites.net/generate-rate-sheet
```

## Example Power Automate Flows

### Flow 1: Generate Rate Sheet on Demand

**Trigger:** Button
**Action 1:** Call API
```
Method: POST
URI: http://localhost:8000/generate-rate-sheet
Body: [your request JSON]
```

**Action 2:** Send Email
- To: recipient@example.com
- Subject: Daily Rate Sheet
- Body: `@body('HTTP').html` (the generated table)

### Flow 2: Automatic Daily Rate Sheet Email

**Trigger:** Schedule (Daily at 8 AM)
**Action 1:** Get your master data from Excel/SharePoint
**Action 2:** Call API with the data
**Action 3:** Send email with the result

### Flow 3: Add to Teams Channel

**Trigger:** Button click in Teams
**Action 1:** Call API
**Action 2:** Send message to Teams
```
Message: "Here's your rate sheet:"
Attachments: [Generate an HTML file from @body('HTTP').html]
```

## API Endpoints

### POST `/generate-rate-sheet`

**Parameters:**
- `currency`: "CAD" or "USD"
- `output_type`: "all_in" or "credit_only"
- `data`: Your master data

**Response:**
```json
{
  "success": true,
  "message": "Generated all_in rate sheet (CAD)",
  "html": "<table>...</table>",
  "error": null,
  "generated_at": "2026-06-08T09:51:24.874247-07:00"
}
```

Copy the `html` field into email or document.

## Example Payloads

### CAD Institutions

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
        "cashable after 30 days": "2.50%",
        "1 Year Fixed": "3.75%",
        "2 Year Fixed": "3.80%",
        "3 Year Fixed": "3.90%",
        "5 Year Fixed": "4.05%",
        "take fi money": "yes"
      },
      {
        "issuer": "TD Bank",
        "available": "available",
        "rating": "R-1 (High) – CDIC",
        "1 Year Fixed": "3.70%",
        "2 Year Fixed": "3.75%",
        "take fi money": "yes"
      }
    ]
  }
}
```

### USD Institutions

```json
{
  "currency": "USD",
  "output_type": "all_in",
  "data": {
    "currency": "USD",
    "institutions": [
      {
        "issuer": "Oaken Financial",
        "available": "available",
        "DBRS": "BBB",
        "S&P": "BBB",
        "1": "4.25%",
        "2": "4.35%"
      }
    ]
  }
}
```

### Credit-Rated Only (CAD)

```json
{
  "currency": "CAD",
  "output_type": "credit_only",
  "data": {
    "currency": "CAD",
    "institutions": [...]
  }
}
```

## Troubleshooting

**"Connection refused" error**
- Make sure API is running: `python api.py`
- Check if port 8000 is available

**API running but Power Automate can't reach it**
- For cloud flows: Deploy API to Azure/cloud service, not localhost
- For desktop flows: API must be on same machine or accessible network

**Rates not showing up**
- Ensure `available` field is exactly "available"
- Check rate is at least 1% (0.01)
- Verify field names match (case-insensitive, but must have correct term names)

**Invalid JSON error**
- Validate your JSON in http://localhost:8000/docs
- Check that all quotes are proper JSON quotes, not smart quotes

## Next Steps

1. ✅ API is running and tested
2. 📝 Create your Power Automate flow
3. 🧪 Test with sample data
4. 📧 Automate your rate sheet emails!

## Need Help?

- **API Docs:** http://localhost:8000/docs
- **Email:** henryjagger@gmail.com

---

**Remember:** Keep the API running (`python api.py`) while using Power Automate!
