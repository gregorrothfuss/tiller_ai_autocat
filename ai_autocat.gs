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

// Store your Gemini API key in Project Settings > Script Properties > GEMINI_API_KEY
// Alternatively, set it directly below if running privately.
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || 'YOUR_KEY_HERE';
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
const MAX_BATCH_SIZE = 25; 

function getSpreadsheetSafely() {
  let attempts = 0;
  let ss = null;
  
  while (attempts < 3) {
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) return ss;
    } catch (e) {
      attempts++;
      Utilities.sleep(1500); // Wait 1.5 seconds before retrying
    }
  }
  
  throw new Error("Spreadsheet service is persistently unavailable after 3 attempts.");
}

/**
 * Main entry point: Scans for uncategorized or generically labeled transactions
 * and uses AI to refine their descriptions and categories.
 */
function categorizeUncategorizedTransactions() {
  const startTime = new Date().getTime();

  if (!GEMINI_API_KEY || GEMINI_API_KEY === "YOUR_KEY_HERE") {
    const msg = "GEMINI_API_KEY is not configured. Please set it in Project Settings > Script Properties, or directly in ai_autocat.gs.";
    Logger.log(msg);
    try {
      SpreadsheetApp.getUi().alert("Configuration Required", msg, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {}
    return;
  }

  const ss = getSpreadsheetSafely();
  const txnSheet = ss.getSheetByName(TRANSACTION_SHEET_NAME);
  if (!txnSheet) {
    Logger.log(`Sheet "${TRANSACTION_SHEET_NAME}" not found.`);
    return;
  }

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
    try { ss.toast("No transactions need categorization.", "AI AutoCat", 4); } catch (e) {}
    return;
  }

  Logger.log(`Found ${allToProcess.length} transaction(s) requiring categorization.`);
  try { ss.toast(`Scanning transactions & Gmail receipts... Found ${allToProcess.length} pending.`, "AI AutoCat", 5); } catch (e) {}

  let totalUpdated = 0;

  // Process in batches to prevent API timeout and hit rate limits gracefully
  for (let i = 0; i < allToProcess.length; i += MAX_BATCH_SIZE) {
    // 6-minute Google Apps Script execution timeout guard (exit safely at 4.5 minutes)
    if (new Date().getTime() - startTime > 270000) {
      const timeoutMsg = `Batch paused to avoid execution timeout. Categorized ${totalUpdated}/${allToProcess.length}. Run again to continue.`;
      Logger.log(timeoutMsg);
      try { ss.toast(timeoutMsg, "AI AutoCat", 8); } catch (e) {}
      return;
    }

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

    // Targeted cell-level writes: protects all untouched columns and formulas from overwriting
    chunk.forEach(item => {
      const update = resultsMap.get(String(item.data[idx.id]));
      if (update) {
        const rowNum = item.rowIndex + 1;
        const assignedCat = categoryList.includes(update.category) ? update.category : FALLBACK_CATEGORY;

        if (idx.category !== -1) txnSheet.getRange(rowNum, idx.category + 1).setValue(assignedCat);
        if (idx.desc !== -1 && update.updated_description) txnSheet.getRange(rowNum, idx.desc + 1).setValue(update.updated_description);
        if (idx.aiFlag !== -1) txnSheet.getRange(rowNum, idx.aiFlag + 1).setValue("TRUE");
        totalUpdated++;
      }
    });

    try { ss.toast(`Processed ${Math.min(i + MAX_BATCH_SIZE, allToProcess.length)} of ${allToProcess.length}...`, "AI AutoCat", 3); } catch (e) {}
  }

  Logger.log(`AutoCat complete. Updated ${totalUpdated} transaction(s).`);
  try { ss.toast(`Successfully categorized ${totalUpdated} transaction(s)!`, "AI AutoCat", 5); } catch (e) {}
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
      `from:${domain} "${absAmount}" after:${formatDate(afterDate)} before:${formatDate(beforeDate)}`,
      `from:${domain} "${absAmount} USD" after:${formatDate(afterDate)} before:${formatDate(beforeDate)}`
    ];
    for (let query of queries) {
      const details = extractPlatformDetails(query, domain);
      if (details) return `EMAIL DATA (${domain} PRICE MATCH):\n` + details;
    }
  }

  // Fallback for Amazon Subscribe & Save or split orders
  if (tryDateFallback && domain === "amazon.com") {
    const dateQueries = [
      `from:${domain} "Ordered" after:${formatDate(afterDate)} before:${formatDate(beforeDate)}`,
      `from:${domain} "Arriving" after:${formatDate(afterDate)} before:${formatDate(beforeDate)}`,
      `from:${domain} "shipped" after:${formatDate(afterDate)} before:${formatDate(beforeDate)}`
    ];
    for (let query of dateQueries) {
      const details = extractPlatformDetails(query, domain);
      if (details) return `EMAIL DATA (${domain} DATE/STATUS MATCH):\n` + details;
    }
  }

  return "";
}

/**
 * Extracts deep structured details (subject, order ID, product categories, and shipment items)
 * from both plain text and HTML bodies.
 */
function extractPlatformDetails(query, domain) {
  try {
    const threads = GmailApp.search(query, 0, 1);
    if (!threads || threads.length === 0) return null;
    
    const messages = threads[0].getMessages();
    const latestMsg = messages[messages.length - 1];
    const subject = latestMsg.getSubject() || "";
    const plainBody = latestMsg.getPlainBody() || "";
    const htmlBody = latestMsg.getBody() || "";

    if (domain === "amazon.com") {
      return parseAmazonMessage(subject, plainBody, htmlBody);
    }

    // Default for PayPal, Venmo, eBay
    return `Subject: ${subject}\nBody: ${plainBody.substring(0, 2500).replace(/\s\s+/g, ' ')}`;
  } catch (e) {
    return null;
  }
}

/**
 * Specialized parser for modern Amazon emails:
 * - Extracts Order ID (###-#######-#######)
 * - Extracts item types from Subject ("Ordered 1 item: Adhesive Tape")
 * - Extracts item types from Confirmation Header ("your Adhesive Tape item is confirmed!")
 * - Checks if related shipment email exists with full product title ("order of '...' has shipped")
 * - Extracts HTML image alt attributes
 */
function parseAmazonMessage(subject, plainBody, htmlBody) {
  const combinedText = subject + "\n" + plainBody;
  
  // 1. Extract Order ID
  const orderIdMatch = combinedText.match(/\b\d{3}-\d{7}-\d{7}\b/);
  const orderId = orderIdMatch ? orderIdMatch[0] : null;

  const detectedItems = [];

  // 2. Check subject patterns (e.g. "Ordered 1 item: Adhesive Tape")
  const mSubj = subject.match(/Ordered \d+ items?:\s*(.+)/i);
  if (mSubj) {
    const val = mSubj[1].trim();
    if (val && !val.toLowerCase().startsWith("item")) detectedItems.push(val);
  }
  
  // Check shipment subject pattern: 'Your Amazon.com order of "..." has shipped!'
  const mShip = subject.match(/order of ["\u201c](.+?)["\u201d]/i);
  if (mShip) {
    detectedItems.push(mShip[1].trim());
  }

  // 3. Check plain text body confirmation headers (e.g. "Gregor, your Adhesive Tape item is confirmed!")
  const mConfirmed = plainBody.match(/your\s+([A-Za-z0-9\s&,.-]+?)\s+item is confirmed/i);
  if (mConfirmed) {
    const val = mConfirmed[1].trim();
    if (val && !detectedItems.some(item => item.toLowerCase() === val.toLowerCase())) {
      detectedItems.push(val);
    }
  }

  // Example: "1 Adhesive Tape item"
  const mCount = plainBody.match(/\b\d+\s+([A-Za-z0-9\s&,.-]+?)\s+item\b/i);
  if (mCount) {
    const val = mCount[1].trim();
    if (val && !["ordered", "confirmed", "total"].includes(val.toLowerCase()) && 
        !detectedItems.some(item => item.toLowerCase() === val.toLowerCase())) {
      detectedItems.push(val);
    }
  }

  // 4. Secondary lookup: If we have an Order ID, check if a shipment email exists with specific product title!
  if (orderId) {
    try {
      const relatedThreads = GmailApp.search(`from:amazon.com "${orderId}"`, 0, 3);
      for (let t of relatedThreads) {
        for (let m of t.getMessages()) {
          const s = m.getSubject() || "";
          const shipMatch = s.match(/order of ["\u201c](.+?)["\u201d]/i);
          if (shipMatch) {
            const specificTitle = shipMatch[1].trim();
            if (!detectedItems.includes(specificTitle)) {
              detectedItems.unshift(specificTitle); // Prioritize specific product title!
            }
          }
        }
      }
    } catch (err) { }
  }

  // 5. HTML alt-tag extraction for thumbnail images
  const imgRegex = /<img[^>]+alt=["']([^"']+)["'][^>]*>/gi;
  const ignoredAlts = [
    "amazon", "prime", "stars", "rating", "order", "delivery", "box", "package", 
    "logo", "return", "view", "track", "icon", "arrow", "cart", "completed", "pending", "amazon-order-confirmation-gif"
  ];
  let imgMatch;
  while ((imgMatch = imgRegex.exec(htmlBody)) !== null) {
    const alt = imgMatch[1].trim();
    const lower = alt.toLowerCase();
    if (alt.length > 3 && !ignoredAlts.some(ign => lower === ign || lower.startsWith(ign + " "))) {
      if (!detectedItems.some(item => item.toLowerCase() === lower)) {
        detectedItems.push(alt);
      }
    }
  }

  // Assemble summary
  let result = `Subject: ${subject}\n`;
  if (orderId) result += `Order ID: ${orderId}\n`;
  if (detectedItems.length > 0) {
    result += `DETECTED PRODUCTS / ITEM TYPES: ${JSON.stringify(detectedItems)}
`;
  }
  result += `Body Snippet: ${plainBody.substring(0, 1500).replace(/\s\s+/g, ' ')}`;
  return result;
}

/**
 * Historical lookup to see how similar descriptions were handled in the past.
 */
function findSimilarInMemory(originalDescription, historicalRows, idx) {
  const lowerOrig = originalDescription.toLowerCase();
  
  // Exclude platforms where the first 12 characters are identical across all purchases
  const isPlatform = ["amazon", "amzn", "paypal", "venmo", "ebay"].some(p => lowerOrig.includes(p));
  if (isPlatform) {
    return [];
  }

  const matchString = lowerOrig.substring(0, 12);
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

CRITICAL AMAZON RULES:
- Inspect "DETECTED PRODUCTS / ITEM TYPES", "Subject", and "Body Snippet" in platform_order_details.
- Modern Amazon emails often specify the item type in the subject (e.g. "Ordered 1 item: Adhesive Tape") or header ("your Adhesive Tape item is confirmed!").
- If DETECTED PRODUCTS has an item (e.g. "Adhesive Tape" or a specific product name), USE THAT as updated_description!
- Do NOT reject descriptions like "Adhesive Tape" as generic — they are valid item types.
- Only ban platform names like "Amazon", "Shopping", "Order", or raw transaction numbers.
- For Subscribe & Save, match delivery dates when price is absent.

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

  const GEMINI_MODEL = 'gemini-flash-latest';
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;
  
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      const code = response.getResponseCode();
      if (code === 200) {
        const parsed = JSON.parse(response.getContentText());
        const text = parsed.candidates[0].content.parts[0].text;
        return JSON.parse(text).suggested_transactions;
      }
      Logger.log(`Gemini API returned status ${code} on attempt ${attempt}/3: ${response.getContentText()}`);
      if (code === 429 || code >= 500) {
        Utilities.sleep(attempt * 2000);
        continue;
      }
      break;
    } catch (err) {
      Logger.log(`Gemini API call failed on attempt ${attempt}/3: ${err}`);
      if (attempt < 3) Utilities.sleep(attempt * 2000);
    }
  }
  return null;
}



/**
 * WEB APP ENDPOINTS FOR LOCAL AMAZON RESOLVER BRIDGE
 * Allows the local python resolver on your Mac to fetch pending Amazon orders and write back results.
 */
function doGet(e) {
  try {
    const ss = getSpreadsheetSafely();
    const txnSheet = ss.getSheetByName(TRANSACTION_SHEET_NAME);
    if (!txnSheet) return ContentService.createTextOutput(JSON.stringify({ error: "Transactions sheet not found" })).setMimeType(ContentService.MimeType.JSON);

    const fullData = txnSheet.getDataRange().getValues();
    if (fullData.length < 2) return ContentService.createTextOutput(JSON.stringify({ pending: [] })).setMimeType(ContentService.MimeType.JSON);

    const headers = fullData[0];
    const rows = fullData.slice(1);

    const idx = {
      id: headers.indexOf(TRANSACTION_ID_COL_NAME),
      origDesc: headers.indexOf(ORIGINAL_DESCRIPTION_COL_NAME),
      desc: headers.indexOf(DESCRIPTION_COL_NAME),
      category: headers.indexOf(CATEGORY_COL_NAME),
      date: headers.indexOf(DATE_COL_NAME),
      amount: headers.indexOf(AMOUNT_COL_NAME)
    };

    const pending = [];
    rows.forEach((r, i) => {
      const origDesc = String(r[idx.origDesc] || "").toLowerCase();
      const currentDesc = String(r[idx.desc] || "").toLowerCase();
      const currentCat = String(r[idx.category] || "").toLowerCase();

      const isAmazon = origDesc.includes("amazon") || origDesc.includes("amzn");
      if (!isAmazon) return;

      const needsCat = !currentCat || currentCat === FALLBACK_CATEGORY.toLowerCase();
      const isGeneric = currentDesc === "" || ["amazon", "amzn", "shopping"].includes(currentDesc.trim()) || currentDesc.match(/^[A-Z0-9*\s]{8,}$/);

      if (needsCat || isGeneric) {
        const amount = idx.amount !== -1 ? r[idx.amount] : null;
        const date = idx.date !== -1 ? r[idx.date] : null;
        
        // Find Order ID in description or search Gmail
        let orderId = null;
        const orderMatch = String(r[idx.origDesc]).match(/\b\d{3}-\d{7}-\d{7}\b/);
        if (orderMatch) {
          orderId = orderMatch[0];
        } else if (amount && date) {
          const absAmount = Math.abs(amount).toFixed(2);
          const dateObj = new Date(date);
          const afterDate = new Date(dateObj.getTime() - (7 * 24 * 60 * 60 * 1000));
          const beforeDate = new Date(dateObj.getTime() + (3 * 24 * 60 * 60 * 1000));
          const formatDate = (d) => `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
          const q = `from:amazon.com "${absAmount}" after:${formatDate(afterDate)} before:${formatDate(beforeDate)}`;
          try {
            const threads = GmailApp.search(q, 0, 1);
            if (threads.length > 0) {
              const body = threads[0].getMessages()[0].getPlainBody();
              const m = body.match(/\b\d{3}-\d{7}-\d{7}\b/);
              if (m) orderId = m[0];
            }
          } catch(err) {}
        }

        pending.push({
          row_index: i + 2,
          transaction_id: r[idx.id],
          date: date ? Utilities.formatDate(new Date(date), "GMT", "yyyy-MM-dd") : "",
          amount: amount,
          original_description: r[idx.origDesc],
          current_description: r[idx.desc],
          order_id: orderId
        });
      }
    });

    const categoryList = getAllowedCategories();
    return ContentService.createTextOutput(JSON.stringify({
      pending_transactions: pending,
      categories: categoryList
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const ss = getSpreadsheetSafely();
    const txnSheet = ss.getSheetByName(TRANSACTION_SHEET_NAME);
    if (!txnSheet) return ContentService.createTextOutput(JSON.stringify({ error: "Transactions sheet not found" })).setMimeType(ContentService.MimeType.JSON);

    const postData = JSON.parse(e.postData.contents);
    const updates = postData.updates || [];

    const fullData = txnSheet.getDataRange().getValues();
    const headers = fullData[0];
    const idx = {
      id: headers.indexOf(TRANSACTION_ID_COL_NAME),
      desc: headers.indexOf(DESCRIPTION_COL_NAME),
      category: headers.indexOf(CATEGORY_COL_NAME),
      aiFlag: headers.indexOf(AI_AUTOCAT_COL_NAME)
    };

    const idToRowMap = new Map();
    for (let i = 1; i < fullData.length; i++) {
      idToRowMap.set(String(fullData[i][idx.id]), i + 1);
    }

    let updatedCount = 0;
    updates.forEach(u => {
      const rowNum = idToRowMap.get(String(u.transaction_id));
      if (rowNum) {
        if (idx.desc !== -1 && u.updated_description) {
          txnSheet.getRange(rowNum, idx.desc + 1).setValue(u.updated_description);
        }
        if (idx.category !== -1 && u.category) {
          txnSheet.getRange(rowNum, idx.category + 1).setValue(u.category);
        }
        if (idx.aiFlag !== -1) {
          txnSheet.getRange(rowNum, idx.aiFlag + 1).setValue("TRUE");
        }
        updatedCount++;
      }
    });

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      updated: updatedCount
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}
