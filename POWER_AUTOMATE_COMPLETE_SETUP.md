# Complete Power Automate Setup Guide

Full step-by-step instructions to set up Power Automate with your Rate Sheet API.

---

## Prerequisites

✅ API server is running: `python api.py`
✅ You have a Power Automate account (free or paid)
✅ For desktop flows: Power Automate Desktop installed on your Windows/Mac machine

---

## Option 1: Desktop Flow (For Local Automation)

Use this if you want to run automation on your computer (Windows/Mac).

### Step 1: Open Power Automate Desktop

1. Search for **Power Automate** on your computer
2. Click **Power Automate Desktop** app
3. Sign in with your Microsoft account (or create one - it's free)

### Step 2: Create a New Flow

1. Click **+ Create** (top left)
2. Select **Cloud flow** → **Automated cloud flow**
3. Give it a name: `Generate Rate Sheet`
4. Choose a trigger:
   - **Button** (if you want to run it manually)
   - **Schedule** (if you want it daily)
   - **When an email arrives** (if triggered by email)
5. Click **Create**

### Step 3: Add the HTTP Action

1. Click **+ Add an action** (or **+ New step**)
2. Search for **HTTP**
3. Select **HTTP** action
4. Fill in the fields:

| Field | Value |
|-------|-------|
| **Method** | POST |
| **URI** | `http://localhost:8001/generate-rate-sheet` |

### Step 4: Add Request Headers

1. Click the **...** menu on the HTTP action
2. Select **Settings**
3. Toggle **Secure inputs** to ON (optional, for security)
4. Go back to the HTTP action
5. In the **Headers** field, click to add:

```
Content-Type: application/json
```

### Step 5: Add Request Body

In the **Body** field, paste this JSON:

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
        "3 Year Fixed": "3.90%",
        "5 Year Fixed": "4.05%"
      },
      {
        "issuer": "TD Bank",
        "available": "available",
        "rating": "R-1 (High) – CDIC",
        "1 Year Fixed": "3.70%",
        "2 Year Fixed": "3.75%",
        "3 Year Fixed": "3.85%",
        "5 Year Fixed": "4.00%"
      },
      {
        "issuer": "Scotiabank",
        "available": "available",
        "rating": "R-1 (High) – CDIC",
        "1 Year Fixed": "3.65%",
        "2 Year Fixed": "3.70%",
        "3 Year Fixed": "3.80%",
        "5 Year Fixed": "3.95%"
      }
    ]
  }
}
```

### Step 6: Send the Result via Email

1. Click **+ Add an action**
2. Search for **Send an email** (Outlook)
3. Select **Send an email (V2)**
4. Fill in:
   - **To:** your email address
   - **Subject:** `Daily Rate Sheet`
   - **Body:** Click in the body field and select **Expression** tab
   - Paste: `@body('HTTP').html`
   - Click **OK**

Your email body now shows the generated rate table!

### Step 7: Test Your Flow

1. Click **Save** (top right)
2. Click **Test** (top right)
3. Choose **I'll perform the trigger action**
4. Click **Test**
5. Click the trigger button (e.g., **Manually trigger**) if using button trigger
6. Wait for it to complete (should take 5-10 seconds)
7. Check your email - you should receive the rate sheet table!

---

## Option 2: Cloud Flow (For Always-On Automation)

Use this if you want Power Automate to run in the cloud automatically.

> **Note:** Cloud flows can't reach `localhost`. You'll need to deploy the API to Azure first.

### Step 2A: Deploy API to Azure (Required for Cloud Flows)

#### Option A: Using Azure App Service

**1. Create an Azure account** (free tier available)
   - Go to https://azure.microsoft.com/en-us/free
   - Sign up

**2. Deploy the API**

```bash
# Install Azure CLI
# Mac: brew install azure-cli
# Windows: choco install azure-cli

# Login to Azure
az login

# Create resource group
az group create --name rate-sheet --location eastus

# Deploy the app
cd /Users/henry/Desktop/Rate-sheet-app
az webapp up --name rate-sheet-api-YOURNAME --resource-group rate-sheet --runtime "python|3.11" --sku B1
```

Your API is now live at: `https://rate-sheet-api-YOURNAME.azurewebsites.net`

#### Option B: Using Docker on a Server

```bash
# In your Rate-sheet-app folder, create Dockerfile:
# (already provided if you have one)

docker build -t rate-sheet-api .
docker run -p 8001:8001 rate-sheet-api
```

Then access via your server's IP: `http://your-server-ip:8001`

### Step 2B: Create Cloud Flow

1. Go to **Power Automate** in your browser: https://powerautomate.microsoft.com
2. Click **+ Create** → **Automated cloud flow**
3. Name it: `Generate Rate Sheet Daily`
4. Choose trigger: **Schedule - Recurrence**
5. Set it to **Daily** at **8:00 AM**
6. Click **Create**

### Step 3: Add HTTP Action

1. Click **+ New step**
2. Search and select **HTTP**
3. Fill in:

| Field | Value |
|-------|-------|
| **Method** | POST |
| **URI** | `https://rate-sheet-api-YOURNAME.azurewebsites.net/generate-rate-sheet` |
| **Headers** | `Content-Type: application/json` |
| **Body** | (same JSON as Step 5 above) |

### Step 4: Send Email with Result

1. Click **+ New step**
2. Search **Send an email (V2)** (Outlook)
3. Fill in:
   - **To:** your email
   - **Subject:** Rate Sheet - @{utcNow('MM/dd/yyyy')}
   - **Body:** Click **Expression** and paste: `@body('HTTP').html`

### Step 5: Save and Test

1. Click **Save**
2. Wait for it to run at 8 AM, or click **Test** → **I'll perform the trigger action**

---

## Advanced Example: Send to Teams Channel

Instead of email, post the rate sheet to Microsoft Teams.

### Setup

1. In your Power Automate flow, remove the **Send an email** action
2. Click **+ New step**
3. Search for **Post message in a chat or channel** (Teams)
4. Select it
5. Fill in:
   - **Team:** Choose your team
   - **Channel:** Choose a channel
   - **Message:** Click **HTML** tab and paste: `@body('HTTP').html`
6. Click **Save**

Now the rate sheet posts to Teams daily!

---

## Different Request Payloads

### Example 1: CAD All Institutions

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
        "2 Year Fixed": "3.80%"
      },
      {
        "issuer": "TD Bank",
        "available": "available",
        "rating": "R-1 (High) – CDIC",
        "1 Year Fixed": "3.70%",
        "2 Year Fixed": "3.75%"
      }
    ]
  }
}
```

### Example 2: CAD Credit-Rated Only

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

Filters out unrated institutions, shows only:
- R-1, R-2, AA, BBB ratings
- 100% guarantees

### Example 3: USD Institutions

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
        "2": "4.35%",
        "Cashable after 30": "3.50%"
      }
    ]
  }
}
```

---

## Dynamic Request Body (Pull from Excel/SharePoint)

Instead of hardcoding the institution data, pull it from Excel online:

### Step 1: Store Master Data in Excel Online

1. Open Excel Online in OneDrive
2. Create a sheet called "MasterRates"
3. Columns:
   - Issuer
   - Available
   - Rating
   - 1 Year Fixed
   - 2 Year Fixed
   - etc.

### Step 2: Use in Power Automate

1. In your flow, add action: **Excel Online (Business)**
2. Select **List rows present in a table**
3. Choose your workbook and table
4. In the HTTP Body, reference the Excel data:

```json
{
  "currency": "CAD",
  "output_type": "all_in",
  "data": {
    "currency": "CAD",
    "institutions": @body('List_rows_present_in_a_table')
  }
}
```

Power Automate automatically converts Excel rows to JSON!

---

## Troubleshooting

### "Connection refused" / "Cannot reach API"

**For Desktop Flow:**
- Is `python api.py` running? (Check terminal)
- Is API on port 8001? (Check output: `http://0.0.0.0:8001`)
- Are you using `http://localhost:8001`? (Should be - it's local)

**For Cloud Flow:**
- Are you using the Azure URL? (Should start with `https://`)
- Is the Azure deployment successful? (Check Azure portal)
- Do you have internet? (Cloud flows need internet)

### "Invalid JSON" Error

- Copy your JSON to http://jsonlint.com and validate
- Check for smart quotes (copy from Notepad instead)
- Make sure all field names are in quotes
- Make sure all string values are in quotes (rates too: "3.75%")

### Email not sending

- Check **Outlook** junk folder
- Verify email address in the **To** field
- Try test mode: click **Test** → check for error message
- Check Power Automate run history for errors

### Rate sheet table looks wrong

- Check institution `"available"` is exactly `"available"` (case-insensitive but must exist)
- Check rates are at least 1% (0.01)
- Make sure term names match exactly (e.g., "1 Year Fixed" not "1yr")

### How to Debug

1. In Power Automate, after your HTTP action, add action: **Compose**
2. In the **Inputs** field, paste: `@body('HTTP')`
3. Save and test
4. In the test results, expand the **Compose** step to see the exact response
5. Check the `html` field - it contains the generated table

---

## API Field Reference

### Supported CAD Terms
```
Cashable After 30 Days
Cashable After 90 Days
30 Days
60 Days
90 Days
180 Days
270 Days
1 Year Fixed
18 Month Fixed
2 Year Fixed
3 Year Fixed
4 Year Fixed
5 Year Fixed
```

### Supported USD Terms
```
Cashable after 30
Cashable after 90
Cashable after 180
30
60
90
120
180
270
1
18 months
2
```

### Field Names (Case-Insensitive)
- `issuer` (required)
- `available` (required, must be "available")
- `rating` (optional, for CAD)
- `dbrs` (optional, for USD)
- `s&p` (optional, for USD)
- `take fi money` (optional, "yes" or "no")

---

## Complete Working Example

Here's a complete flow you can copy and paste:

**Trigger:** Button

**Action 1: HTTP Request**
```
Method: POST
URI: http://localhost:8001/generate-rate-sheet
Headers: Content-Type: application/json
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
        "2 Year Fixed": "3.80%",
        "3 Year Fixed": "3.90%",
        "5 Year Fixed": "4.05%"
      },
      {
        "issuer": "TD Bank",
        "available": "available",
        "rating": "R-1 (High) – CDIC",
        "1 Year Fixed": "3.70%",
        "2 Year Fixed": "3.75%",
        "3 Year Fixed": "3.85%",
        "5 Year Fixed": "4.00%"
      }
    ]
  }
}
```

**Action 2: Send an email (V2)**
```
To: your-email@example.com
Subject: Daily Rate Sheet
Body: (Click HTML, paste: @body('HTTP').html)
```

**Action 3: Save**

Done! Click the button to test.

---

## Questions?

- **API Documentation:** http://localhost:8001/docs
- **Email:** henryjagger@gmail.com
- **Power Automate Help:** https://learn.microsoft.com/en-us/power-automate/

