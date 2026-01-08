/**
 * TILLER AI AUTOCAT - ADVANCED TRANSACTION CATEGORIZATION ENGINE
 * 
 * This script uses Gemini AI to automatically categorize and clean up transaction 
 * descriptions by correlating data from Gmail receipts, Venmo memos, and historical lookups.
 * 
 * CORE FEATURES:
 * 1. Intelligent Gmail Integration: Hunts for receipts from Amazon, Venmo, PayPal, and eBay.
 * 2. Date-Match Fallback: Specifically handles Amazon Subscribe & Save by matching 
 *    delivery dates when prices aren't listed in the email.
 * 3. Venmo Memo Extraction: Pulls rough memos (e.g., "Chil crisps") and transforms 
 *    them into professional descriptions (e.g., "Chili Crisp").
 * 4. Zero Tolerance for Generic Labels: Strictly bans "Transfer", "Shopping", or 
 *    brand-prefixed labels like "Amazon: Item" to ensure a clean, useful ledger.
 * 5. Performance Batching: Processes transactions in groups of 50 to maximize 
 *    speed and reliability within Google Apps Script execution limits.
 */

const GEMINI_API_KEY = 'YOUR_KEY_HERE';
const TRANSACTION_SHEET_NAME = 'Transactions';
const CATEGORY_SHEET_NAME = 'Categories';
const TRANSACTION_ID_COL_NAME = 'Transaction ID';
const ORIGINAL_DESCRIPTION_COL_NAME = 'Full Description';
const DESCRIPTION_COL_NAME = 'Description';
const CATEGORY_COL_NAME = 'Category';
const AI_AUTOCAT_COL_NAME = 'AI AutoCat';
const DATE_COL_NAME = 'Date';
const AMOUNT_COL_NAME = 'Amount'; 
const FALLBACK_CATEGORY = "To Be Categorized";
const MAX_BATCH_SIZE = 50; 

/**
 * Main entry point: Scans for uncategorized or generically labeled transactions
 * and uses AI to refine their descriptions and categories.
 */
function categorizeUncategorizedTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txnSheet = ss.getSheetByName(TRANSACTION_SHEET_NAME);
  if (!txnSheet) return;

  const fullData = txnSheet.getDataRange().getValues();
  if (fullData.length < 2) return;
  
  const headers = fullData[0];
  const rows = fullData.slice(1);

  const idx = {
    id: headers.indexOf(TRANSACTION_ID_COL_NAME),
    origDesc: headers.indexOf(ORIGINAL_DESCRIPTION_COL_NAME),
    desc: headers.indexOf(DESCRIPTION_COL_NAME),
    category: headers.indexOf(CATEGORY_COL_NAME),
    aiFlag: headers.indexOf(AI_AUTOCAT_COL_NAME),
    date: headers.indexOf(DATE_COL_NAME),
    amount: headers.indexOf(AMOUNT_COL_NAME)
  };

  const categoryList = getAllowedCategories();
  
  // Create a memory cache of already categorized rows for historical lookup
  const categorizedRows = rows.filter(r => r[idx.category] && r[idx.origDesc]);
  
  const allToProcess = rows
    .map((r, i) => ({ data: r, rowIndex: i + 1 }))
    .filter(item => {
      const rowData = item.data;
      const origDesc = String(rowData[idx.origDesc] || "").toLowerCase();
      const currentDesc = String(rowData[idx.desc] || "").toLowerCase();
      const currentCat = String(rowData[idx.category] || "").toLowerCase();

      // Process if missing category
      if (!currentCat || currentCat === FALLBACK_CATEGORY.toLowerCase()) return true;

      // Also target transactions that are currently generic platform noise
      const isPlatform = origDesc.includes("amazon") || origDesc.includes("amzn") || 
                         origDesc.includes("paypal") || origDesc.includes("ebay") ||
                         origDesc.includes("venmo") || origDesc.includes("xfer");
      
      if (isPlatform) {
        const isGeneric = currentDesc === "" || 
                          ["amazon", "paypal", "ebay", "venmo", "shopping", "transfer"].includes(currentDesc.trim()) ||
                          currentDesc.includes("*") || 
                          currentDesc.match(/[A-Z0-9]{8,}/) || // Contains long IDs but no product
                          currentDesc.includes("transfer") ||
                          currentDesc.includes("instant xfer");
        
        if (isGeneric) return true;
      }

      return false;
    });

  if (allToProcess.length === 0) {
    Logger.log("No transactions need processing.");
    return;
  }

  // Process in batches to prevent API timeout and hit rate limits gracefully
  for (let i = 0; i < allToProcess.length; i += MAX_BATCH_SIZE) {
    const chunk = allToProcess.slice(i, i + MAX_BATCH_SIZE);
    const transactionList = chunk.map(item => {
      const fullOrigDesc = String(item.data[idx.origDesc]);
      const amount = idx.amount !== -1 ? item.data[idx.amount] : null;
      const date = idx.date !== -1 ? item.data[idx.date] : null;
      
      let platformContext = "";
      const lowerDesc = fullOrigDesc.toLowerCase();
      
      // Attempt to retrieve email context for known platforms
      if (lowerDesc.includes("amazon") || lowerDesc.includes("amzn")) {
        platformContext = fetchPlatformEmail(amount, date, "amazon.com", true);
      } else if (lowerDesc.includes("paypal") || lowerDesc.includes("pp*")) {
        platformContext = fetchPlatformEmail(amount, date, "paypal.com", false);
      } else if (lowerDesc.includes("ebay")) {
        platformContext = fetchPlatformEmail(amount, date, "ebay.com", false);
      } else if (lowerDesc.includes("venmo")) {
        platformContext = fetchPlatformEmail(amount, date, "venmo.com", false);
      }

      return {
        transaction_id: item.data[idx.id],
        transaction_date: date ? Utilities.formatDate(new Date(date), "GMT", "yyyy-MM-dd") : "Unknown",
        original_description: fullOrigDesc,
        platform_order_details: platformContext,
        previous_transactions: findSimilarInMemory(fullOrigDesc, categorizedRows, idx)
      };
    });

    const updatedResults = lookupDescAndCategory(transactionList, categoryList);
    if (!updatedResults) continue;

    const resultsMap = new Map();
    updatedResults.forEach(res => resultsMap.set(String(res.transaction_id), res));

    // Write results back to the sheet
    chunk.forEach(item => {
      const update = resultsMap.get(String(item.data[idx.id]));
      if (update) {
        const rowNum = item.rowIndex + 1;
        const rowData = item.data;
        
        rowData[idx.category] = categoryList.includes(update.category) ? update.category : FALLBACK_CATEGORY;
        if (idx.desc !== -1) rowData[idx.desc] = update.updated_description;
        if (idx.aiFlag !== -1) rowData[idx.aiFlag] = "TRUE";
        txnSheet.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);
      }
    });
  }
}

/**
 * Searches Gmail for order confirmations based on price or date (for subscriptions).
 */
function fetchPlatformEmail(amount, transactionDate, domain, tryDateFallback) {
  if (!transactionDate) return "";
  
  const absAmount = amount ? Math.abs(amount).toFixed(2) : null;
  const dateObj = new Date(transactionDate);
  // Expand search window to handle processing delays between merchant and bank
  const afterDate = new Date(dateObj.getTime() - (7 * 24 * 60 * 60 * 1000));
  const beforeDate = new Date(dateObj.getTime() + (3 * 24 * 60 * 60 * 1000));
  const formatDate = (d) => `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
  
  // Try exact price search first
  if (absAmount) {
    const queries = [
      `from:${domain} "$${absAmount}" after:${formatDate(afterDate)} before:${formatDate(beforeDate)}`,
      `from:${domain} "${absAmount}" after:${formatDate(afterDate)} before:${formatDate(beforeDate)}`
    ];
    for (let query of queries) {
      const body = getFirstEmailBody(query);
      if (body) return `EMAIL DATA (${domain} PRICE MATCH): \n` + body;
    }
  }

  // Fallback for Amazon Subscribe & Save which lacks explicit prices in many notification emails
  if (tryDateFallback && domain === "amazon.com") {
    const dateQuery = `from:${domain} "Arriving" after:${formatDate(afterDate)} before:${formatDate(beforeDate)}`;
    const body = getFirstEmailBody(dateQuery);
    if (body) return `EMAIL DATA (${domain} DATE MATCH): \n` + body;
  }

  return "";
}

function getFirstEmailBody(query) {
  try {
    const threads = GmailApp.search(query, 0, 1);
    if (threads.length > 0) {
      const msg = threads[0].getMessages()[0];
      return msg.getPlainBody().substring(0, 4500).replace(/\s\s+/g, ' ');
    }
  } catch (e) { }
  return null;
}

/**
 * Historical lookup to see how similar descriptions were handled in the past.
 */
function findSimilarInMemory(originalDescription, historicalRows, idx) {
  const matchString = originalDescription.toLowerCase().substring(0, 12);
  return historicalRows
    .filter(row => String(row[idx.origDesc]).toLowerCase().includes(matchString))
    .filter(row => !["transfer", "paypal", "venmo", "amazon", "shopping"].includes(String(row[idx.desc]).toLowerCase()))
    .slice(0, 1)
    .map(row => ({ 
      original_description: row[idx.origDesc], 
      updated_description: row[idx.desc], 
      category: row[idx.category] 
    }));
}

function getAllowedCategories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CATEGORY_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const catIdx = data[0].indexOf(CATEGORY_COL_NAME);
  return data.slice(1).map(row => String(row[catIdx])).filter(Boolean);
}

/**
 * Calls the Gemini API to analyze the transaction data and provide clean results.
 */
function lookupDescAndCategory(transactionList, categoryList) {
  const systemPrompt = `You are an elite financial forensics agent. 

CRITICAL CORRELATION RULE:
- For Amazon "Subscribe & Save", the email might NOT have a price. 
- Look for delivery dates in the EMAIL DATA and compare to "transaction_date".
- Use the item name from matching emails as the description.

STRICT OUTPUT RULES:
- NEVER prefix descriptions with "Amazon:", "PayPal:", "Venmo:", or "eBay:".
- NEVER use generic labels like "Shopping", "Transfer", or "Order".
- For Venmo, clean up rough memos (e.g., "Chil crisps" -> "Chili Crisp").
- For PayPal, extract the actual merchant (e.g., "PayPal * UBER" -> "Uber").

CATEGORIES:
- Must be a verbatim match from: ${JSON.stringify(categoryList)}.
- Avoid the "Transfer" category for any platform-based purchase.

Return JSON: {"suggested_transactions": [{"transaction_id": "...", "updated_description": "...", "category": "..."}]}`;

  const payload = {
    contents: [{ parts: [{ text: JSON.stringify({ transactions: transactionList }) }] }],
    system_instruction: { parts: [{ text: systemPrompt }] },
    generationConfig: { response_mime_type: "application/json" }
  };

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=' + GEMINI_API_KEY;
  try {
    const response = UrlFetchApp.fetch(url, { method: 'POST', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
    const content = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
    return JSON.parse(content).suggested_transactions;
  } catch (e) { return null; }
}
