# Tiller AI AutoCat

Automated transaction categorization engine for **Tiller Community Solutions** and Google Sheets powered by Google's **Gemini AI** (`gemini-flash-latest`).

## About
Tiller AI AutoCat automatically cleans up messy bank descriptions and assigns accurate categories to your ledger:
- Only touches transactions lacking a Category or having generic/raw platform descriptions (e.g. `AMZN Mktp US*...`, `PAYPAL INST XFER`, `Venmo Cashtag`).
- Correlates with Gmail receipts (Amazon, Venmo, PayPal, eBay) to find what was actually bought.
- Fixes Amazon Subscribe & Save transactions where order emails omit pricing by matching delivery dates.
- Cleans Venmo emojis and casual memos into professional expense descriptions (e.g., "Chil crisps" → "Chili Crisp" categorized as `Groceries`).
- Replaces raw platform descriptions with clear, clean merchant/product names (banning prefixes like `Amazon:` or `PayPal:`).
- Marks processed rows with `TRUE` if you include an `AI AutoCat` column in your Transactions sheet.
- Includes an optional local macOS Chrome bridge (`amazon_resolver.py`) to resolve exact Amazon line items when Amazon omits product details from order emails.

## Key Features
- **Powered by `gemini-flash-latest`**: Always runs on Google's latest, fastest Flash model with structured JSON enforcement.
- **Zero-Secret Script Properties**: API keys are stored securely in Apps Script `Script Properties` rather than hardcoded in source code.
- **Gmail Receipt Intelligence**: Hunts for order confirmations across Amazon, Venmo, PayPal, and eBay.
- **Safe Historical Memory**: Protects against platform prefix poisoning so recurring vendors don't misattribute Amazon charges.
- **Resilient Batch Processing**: Built-in exponential backoff for rate limits, processing batches within Google Apps Script execution limits.
- **Optional Chrome Amazon Resolver**: Reads Amazon invoices silently via your logged-in browser session to get itemized titles without manual CSV exports.

## Installation Instructions

### 1. Get a Gemini API Key
Get a free Gemini API key from [Google AI Studio](https://aistudio.google.com/).

### 2. Set Up Google Apps Script
1. In your Tiller Google Sheet, open **Extensions** → **Apps Script**.
2. Click the **Project Settings** (gear icon on the left).
3. Scroll down to **Script Properties** and click **Add script property**:
   - **Property**: `GEMINI_API_KEY`
   - **Value**: *(Paste your Gemini API key)*
   - Click **Save script properties**.
4. In the left panel, click the **Editor** (`< >` icon).
5. In the existing `Code.gs` file (or a new file named `ai_autocat.gs`), paste the entire contents of [`ai_autocat.gs`](ai_autocat.gs).
6. In `code.gs`, paste the contents of [`code.gs`](code.gs) to enable the sheet menu.
7. Click **Save** (disk icon).

### 3. Usage
1. Refresh your Tiller Google Sheet.
2. After a few seconds, an **AI AutoCat** menu will appear in the top toolbar.
3. Click **AI AutoCat** → **Run AutoCat**.
4. *First Run Authorization*: Google will prompt you to authorize permissions for Gmail (read-only search for receipts) and Sheets.
5. (Optional) To run automatically on a schedule, click the clock icon in Apps Script (**Triggers**) and add a time-driven trigger for `categorizeUncategorizedTransactions` (e.g. nightly).

---

## Optional: Amazon Chrome Invoice Resolver

Amazon recently stopped including specific product names in order confirmation emails (often stating only `Ordered 1 item: Adhesive Tape`).

If you want exact product titles for Amazon orders, you can use the companion Python tool [`amazon_resolver.py`](amazon_resolver.py):

### Requirements
- macOS with Google Chrome installed and logged in to your Amazon account.
- Enable Apple Events in Chrome: **View** → **Developer** → **Allow JavaScript from Apple Events**.
- Python 3 (`urllib` and standard libraries only; no pip dependencies required).

### Quick Lookup
Resolve and categorize a single Amazon order:
```bash
export GEMINI_API_KEY="your-api-key"
./amazon_resolver.py --order 113-8099586-4134612
```

### Automated Two-Way Sync with Google Sheets
1. In Apps Script, click **Deploy** → **New deployment**.
2. Select type **Web app**.
   - **Execute as**: `Me`
   - **Who has access**: `Anyone` (or within your organization).
3. Copy the **Web App URL**.
4. Run the sync command:
```bash
./amazon_resolver.py --sync "https://script.google.com/macros/s/.../exec"
```
The script will fetch all pending Amazon transactions, silently load their print invoices in Chrome, categorize the exact items with Gemini, and write them directly back to your Tiller sheet.

---

## License
MIT License. See [LICENSE](LICENSE) for details.
