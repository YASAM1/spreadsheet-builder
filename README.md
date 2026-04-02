# Spreadsheet Builder API

A microservice that generates filled Excel spreadsheets from JSON data using a template. Designed to be called from n8n workflows.

## Features

- Accepts JSON sections data and fills an Excel template
- Preserves existing formatting, styles, and formulas
- Rehydrates row formulas when a section needs more rows than the template originally provided
- Suppresses heading-like rows that leak through from summary-report documents
- Deterministically reroutes obvious miscategorized rows such as credit cards, tax liabilities, and business interests
- Normalizes debt rows into column B as negative values when the payload clearly represents a liability
- Handles various input formats (array, object, nested)
- Dynamically inserts rows when needed
- Returns binary XLSX file

## Prerequisites

- Node.js 18+
- npm

## Installation

```bash
npm install
```

## Running Locally

### Start the server

```bash
npm start
```

The server will start on `http://localhost:3000` (or `PORT` environment variable).

### Development mode (auto-reload)

```bash
npm run dev
```

## API Endpoints

### Health Check

```
GET /health
```

Response:
```json
{ "ok": true }
```

### Generate Mediation XLSX

```
POST /generate-mediation-xlsx
Content-Type: application/json
```

**Request Body Formats:**

Format A - Direct object:
```json
{
  "sections": {
    "Real Property": [...],
    "Bank Accounts": [...]
  }
}
```

Format B - Array (from n8n):
```json
[
  {
    "sections": {
      "Real Property": [...],
      "Bank Accounts": [...]
    }
  }
]
```

Format C - Nested json:
```json
{
  "json": {
    "sections": {
      "Real Property": [...]
    }
  }
}
```

**Row Object Structure:**
```json
{
  "description": "Property description",
  "market_value": 350000,
  "debt_balance": -120000,
  "notes": "Additional notes"
}
```

**Available Section Keys:**
- `Real Property`
- `Bank Accounts`
- `Business Interests`
- `Retirement Benefits`
- `Stock Options`
- `Personal Property`
- `Life Insurance / Annuities`
- `Vehicles`
- `Credit Cards`
- `Bonuses`
- `Other Debts & Liabilities`
- `Federal Income Taxes`
- `Attorney's & Expert Fees`
- `Spouse 1's Separate Property`
- `Spouse 2's Separate Property`
- `Children's Property`

The service also accepts a small set of alias section names from the sample
summary reports, including `Real Properties`, `Cash and Accounts with
Financial Institutions`, `Insurance And Annuities`, and `Motor Vehicles,
Boats, Airplanes, Cycles, etc`.

## Defensive Normalization

Before writing any rows into the template, the service now applies a narrow set
of deterministic safeguards:

- Drops rows whose descriptions are clearly document headings such as `Real
  Properties`, `Credit Cards`, `Community Assets`, `Defined Contribution
  Plans`, and similar non-data labels seen in the sample reports.
- Reroutes obvious credit-card rows into `Credit Cards`.
- Reroutes obvious tax rows into `Federal Income Taxes`.
- Reroutes obvious business-interest rows into `Business Interests`.
- Moves debt amounts into column B and forces them negative for rows that are
  clearly liabilities.

**Response:**
- Success: Binary XLSX file with appropriate headers
- Error 400: Invalid/missing sections
- Error 500: Processing errors with details

## Local Testing

Run the automated workbook regression tests:

```bash
npm test
```

### Create a test payload file

Create `payload.json`:
```json
[
  {
    "sections": {
      "Real Property": [
        {
          "description": "4526 Treasure Trail, Sugar Land, Texas, 77479",
          "market_value": 350000,
          "debt_balance": -120000,
          "notes": "County: Fort Bend | Mortgage Holder: Wells Fargo"
        }
      ],
      "Bank Accounts": [
        {
          "description": "Chase Bank (Checking Account)",
          "market_value": 5324,
          "debt_balance": 0,
          "notes": "Acct (redacted): 012456869 | Balance: 5324"
        }
      ]
    }
  }
]
```

### Test with curl

```bash
curl -X POST http://localhost:3000/generate-mediation-xlsx \
  -H "Content-Type: application/json" \
  --data @payload.json \
  --output output.xlsx
```

### Verify health endpoint

```bash
curl http://localhost:3000/health
```

## Deploy to Render

### 1. Push to GitHub

```bash
# Initialize git (if not already)
git init

# Add files
git add .

# Commit
git commit -m "Initial commit: Spreadsheet Builder API"

# Add your GitHub remote
git remote add origin https://github.com/YOUR_USERNAME/spreadsheet-builder.git

# Push
git push -u origin main
```

### 2. Create Render Web Service

1. Go to [Render Dashboard](https://dashboard.render.com)
2. Click **New** > **Web Service**
3. Connect your GitHub repository
4. Configure the service:
   - **Name:** `spreadsheet-builder` (or your choice)
   - **Region:** Choose closest to your location
   - **Branch:** `main`
   - **Runtime:** `Node`
   - **Build Command:** `npm install`
   - **Start Command:** `npm start`
   - **Instance Type:** Free (or paid for better performance)

5. Click **Create Web Service**

### 3. Verify Deployment

Once deployed, Render will provide a URL like `https://spreadsheet-builder-xxxx.onrender.com`

Test the health endpoint:
```bash
curl https://spreadsheet-builder-xxxx.onrender.com/health
```

Expected response:
```json
{ "ok": true }
```

## n8n Integration

### Configure HTTP Request Node

1. Add an **HTTP Request** node in your n8n workflow

2. Configure the node:
   - **Method:** `POST`
   - **URL:** `https://your-render-url.onrender.com/generate-mediation-xlsx`
   - **Authentication:** None (or add if you implement auth)
   - **Send Headers:** Yes
     - **Header Name:** `Content-Type`
     - **Header Value:** `application/json`
   - **Send Body:** Yes
   - **Body Content Type:** `JSON`
   - **Body:** Use expression `{{ $json }}` (or `{{ $json.sections }}` wrapped in the expected format)
   - **Response Format:** `File`
   - **Put Output in Field:** `data`

3. If your Code node outputs the sections directly, format the body as:
   ```json
   {
     "sections": {{ $json }}
   }
   ```

   Or if your Code node already outputs the full structure:
   ```json
   {{ $json }}
   ```

### Upload to Google Drive

After the HTTP Request node, add a **Google Drive** node:

1. Configure Google Drive node:
   - **Operation:** `Upload`
   - **Binary Data:** Yes
   - **Input Binary Field:** `data`
   - **File Name:** `MediationSpreadsheet-{{ $now.format('yyyy-MM-dd-HHmm') }}.xlsx`
   - **Parent Folder:** Select your target folder

### Complete n8n Workflow Example

```
[Your Data Source]
    ↓
[Code Node - Format sections object]
    ↓
[HTTP Request Node - POST to /generate-mediation-xlsx]
    ↓
[Google Drive Node - Upload binary file]
```

### Example Code Node (JavaScript)

If you need to format your data into the sections structure:

```javascript
// Assuming $input contains your parsed data
const sections = {
  "Real Property": [],
  "Bank Accounts": [],
  // ... other sections
};

// Process your input data and populate sections
// ...

return {
  sections: sections
};
```

## Troubleshooting

### Common Issues

**"Section header not found" error:**
- The section key must match exactly (case-sensitive)
- Check the section mapping in the code

**Empty or corrupted output:**
- Verify the template file exists at `src/template/2025.07.30.Inventory-MediationSpreadsheetMaster.xlsx`
- Check that the worksheet name matches exactly: `Prop Distr - KB Preferred - Use`

**n8n shows binary garbled text:**
- Ensure **Response Format** is set to `File`, not `Text` or `JSON`

**400 Bad Request:**
- Check that your request body contains a valid `sections` object
- Verify JSON syntax is correct

### Debugging

Check Render logs for detailed error messages:
1. Go to your Render service dashboard
2. Click **Logs** tab
3. Look for error messages during request processing

## Column Mapping

Data is written to these columns:
- **Column A:** `market_value`
- **Column B:** `debt_balance`
- **Column D:** `description`
- **Column G:** `notes`

## License

ISC
