// API Keys
const GEMINI_API_KEY = 'YOUR_API_KEY_HERE';

// Sheet Names
const TRANSACTION_SHEET_NAME = 'Transactions';
const CATEGORY_SHEET_NAME = 'Categories';

// Column Names
const TRANSACTION_ID_COL_NAME = 'Transaction ID';
const ORIGINAL_DESCRIPTION_COL_NAME = 'Full Description';
const DESCRIPTION_COL_NAME = 'Description';
const CATEGORY_COL_NAME = 'Category';
const AI_AUTOCAT_COL_NAME = 'AI AutoCat'
const DATE_COL_NAME = 'Date';

// Fallback Transaction Category (to be used when we don't know how to categorize a transaction)
const FALLBACK_CATEGORY = "To Be Categorized";

// Other Misc Paramaters
const MAX_BATCH_SIZE = 50;

function categorizeUncategorizedTransactions() {
  var uncategorizedTransactions = getTransactionsToCategorize();

  var numTxnsToCategorize = uncategorizedTransactions.length;
  if (numTxnsToCategorize == 0) {
    Logger.log("No uncategorized transactions found");
    return;
  }

  Logger.log("Found " + numTxnsToCategorize + " transactions to categorize");
  Logger.log("Looking for historical similar transactions...");

  var transactionList = []
  for (var i = 0; i < uncategorizedTransactions.length; i++) {
    var similarTransactions = findSimilarTransactions(uncategorizedTransactions[i][1]);

    transactionList.push({
      'transaction_id': uncategorizedTransactions[i][0],
      'original_description': uncategorizedTransactions[i][1],
      'previous_transactions': similarTransactions
    });
  }

  Logger.log("Processing this set of transactions and similar transactions with Gemini AI:");
  Logger.log(transactionList);

  var categoryList = getAllowedCategories();

  var updatedTransactions = lookupDescAndCategory(transactionList, categoryList);

  if (updatedTransactions != null) {
    Logger.log("Gemini AI returned the following sugested categories and descriptions:");
    Logger.log(updatedTransactions);
    Logger.log("Writing updated transactions into your sheet...");
    writeUpdatedTransactions(updatedTransactions, categoryList);
    Logger.log("Finished updating your sheet!");
  }
}

// Gets all transactions that have an original description but no category set
function getTransactionsToCategorize() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSACTION_SHEET_NAME);
  var headers = sheet.getRange("1:1").getValues()[0];

  var txnIDColLetter = getColumnLetterFromColumnHeader(headers, TRANSACTION_ID_COL_NAME);
  var origDescColLetter = getColumnLetterFromColumnHeader(headers, ORIGINAL_DESCRIPTION_COL_NAME);
  var categoryColLetter = getColumnLetterFromColumnHeader(headers, CATEGORY_COL_NAME);
  var lastColLetter = getColumnLetterFromColumnHeader(headers, headers[headers.length - 1]);

  var queryString = "SELECT " + txnIDColLetter + ", " + origDescColLetter + " WHERE " + origDescColLetter +
                    " is not null AND " + categoryColLetter + " is null LIMIT " + MAX_BATCH_SIZE;

  var uncategorizedTransactions = Utils.gvizQuery(
      SpreadsheetApp.getActiveSpreadsheet().getId(), 
      queryString, 
      TRANSACTION_SHEET_NAME,
      "A:" + lastColLetter
    );

  return uncategorizedTransactions;
}

function findSimilarTransactions(originalDescription) {
  // Normalize to lowercase
  var matchString = originalDescription.toLowerCase();

  // Remove phone number placeholder
  matchString = matchString.replace('xx', '#');

  // Strip numbers at end
  var descriptionParts = matchString.split('#');
  matchString = descriptionParts[0];

  // Remove unimportant words
  matchString = matchString.replace('direct debit ', '');
  matchString = matchString.replace('direct deposit ', '');
  matchString = matchString.replace('zelle payment from ', '');
  matchString = matchString.replace('bill payment ', '');
  matchString = matchString.replace('dividend received ', '');
  matchString = matchString.replace('debit card purchase ', '');
  matchString = matchString.replace('sq *', '');
  matchString = matchString.replace('sq*', '');
  matchString = matchString.replace('tst *', '');
  matchString = matchString.replace('tst*', '');
  matchString = matchString.replace('in *', '');
  matchString = matchString.replace('in*', '');
  matchString = matchString.replace('tcb *', '');
  matchString = matchString.replace('tcb*', '');
  matchString = matchString.replace('dd *', '');
  matchString = matchString.replace('dd*', '');
  matchString = matchString.replace('py *', '');
  matchString = matchString.replace('py*', '');
  matchString = matchString.replace('p *', '');
  matchString = matchString.replace('pp*', '');
  matchString = matchString.replace('rx *', '');
  matchString = matchString.replace('rx*', '');
  matchString = matchString.replace('intuit *', '');
  matchString = matchString.replace('intuit*', '');
  matchString = matchString.replace('microsoft *', '');
  matchString = matchString.replace('microsoft*', '');

  matchString = matchString.replace('*', ' ');

  // Trim leading & trailing spaces
  matchString = matchString.trim();

  // Trim double spaces
  matchString = matchString.replace(/\s+/g, ' ');

  // Grab first 3 words
  descriptionParts = matchString.split(' ');
  descriptionParts = descriptionParts.slice(0, Math.min(3, descriptionParts.length))
  matchString = descriptionParts.join(' ');

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSACTION_SHEET_NAME);
  var headers = sheet.getRange("1:1").getValues()[0];

  var descColLetter = getColumnLetterFromColumnHeader(headers, DESCRIPTION_COL_NAME);
  var origDescColLetter = getColumnLetterFromColumnHeader(headers, ORIGINAL_DESCRIPTION_COL_NAME);
  var categoryColLetter = getColumnLetterFromColumnHeader(headers, CATEGORY_COL_NAME);
  var dateColLetter = getColumnLetterFromColumnHeader(headers, DATE_COL_NAME);
  var lastColLetter = getColumnLetterFromColumnHeader(headers, headers[headers.length - 1]);

  var queryString = "SELECT " + descColLetter + ", " + categoryColLetter + ", " + origDescColLetter + 
                    " WHERE " + categoryColLetter + " is not null AND (lower(" + 
                    origDescColLetter + ") contains \"" + matchString + "\" OR lower(" + descColLetter +
                    ") contains \"" + matchString + "\") ORDER BY " + dateColLetter +" DESC LIMIT 3";

  Logger.log("Looking for previous transactions with query: " + queryString);
  
  var result = Utils.gvizQuery(
      SpreadsheetApp.getActiveSpreadsheet().getId(), 
      queryString, 
      TRANSACTION_SHEET_NAME,
      "A:" + lastColLetter
    );

  var previousTransactionList = []
  for (var i = 0; i < result.length; i++) {
    previousTransactionList.push({
      'original_description': result[i][2],
      'updated_description': result[i][0],
      'category': result[i][1]
    });
  }

  return previousTransactionList;
}

function writeUpdatedTransactions(transactionList, categoryList) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions");

  // Get Column Numbers
  var headers = sheet.getRange("1:1").getValues()[0];

  var descriptionColumnLetter = getColumnLetterFromColumnHeader(headers, DESCRIPTION_COL_NAME);
  var categoryColumnLetter = getColumnLetterFromColumnHeader(headers, CATEGORY_COL_NAME);
  var transactionIDColumnLetter = getColumnLetterFromColumnHeader(headers, TRANSACTION_ID_COL_NAME);
  var geminiFlagColLetter = getColumnLetterFromColumnHeader(headers, AI_AUTOCAT_COL_NAME);

  for (var i = 0; i < transactionList.length; i++) {
    // Find Row of transaction
    var transactionIDRange = sheet.getRange(transactionIDColumnLetter + ":" + transactionIDColumnLetter);
    var textFinder = transactionIDRange.createTextFinder(transactionList[i]["transaction_id"]);
    var match = textFinder.findNext();
    if (match != null) {
      var transactionRow = match.getRowIndex();

      // Set Updated Category
      var categoryRangeString = categoryColumnLetter + transactionRow;

      try {
        var categoryRange = sheet.getRange(categoryRangeString);

        var updatedCategory = transactionList[i]["category"];
        if (!categoryList.includes(updatedCategory)) {
          updatedCategory = FALLBACK_CATEGORY;
        }
        
        categoryRange.setValue(updatedCategory);
      } catch (error) {
        Logger.log(error);
      }


      // Set Updated Description
      var descRangeString = descriptionColumnLetter + transactionRow;

      try {
        var descRange = sheet.getRange(descRangeString);
        descRange.setValue(transactionList[i]["updated_description"]);
      } catch (error) {
        Logger.log(error);
      }

      // Mark Gemini AI Flag
      if (geminiFlagColLetter != null) {
        var geminiFlagRangeString = geminiFlagColLetter + transactionRow;

        try {
          var geminiFlagRange = sheet.getRange(geminiFlagRangeString);
          geminiFlagRange.setValue("TRUE");
        } catch (error) {
          Logger.log(error);
        }
      }
    }
  }
}

function getAllowedCategories() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var categorySheet = spreadsheet.getSheetByName(CATEGORY_SHEET_NAME)
  var headers = categorySheet.getRange("1:1").getValues()[0];

  var categoryColLetter = getColumnLetterFromColumnHeader(headers, CATEGORY_COL_NAME);

  var categoryListRaw = categorySheet.getRange(categoryColLetter + "2:" + categoryColLetter).getValues();

  var categoryList = []
  for (var i = 0; i < categoryListRaw.length; i++) {
    categoryList.push(categoryListRaw[i][0]);
  }
  return categoryList;
}

function getColumnLetterFromColumnHeader(columnHeaders, columnName) {
  var columnIndex = columnHeaders.indexOf(columnName);
  var columnLetter = "";

    let base = 26;
    let letterCharCodeBase = 'A'.charCodeAt(0);

    while (columnIndex >= 0) {
        columnLetter = String.fromCharCode(columnIndex % base + letterCharCodeBase) + columnLetter;
        columnIndex = Math.floor(columnIndex / base) - 1;
    }

    return columnLetter;
}

function lookupDescAndCategory (transactionList, categoryList) {
  var transactionDict = {
    "transactions": transactionList
  };

  var tillerFormat = 'You will be given JSON input with a list of transaction descriptions and potentially related previously categorized transactions in the following format: \
            {"transactions": [\
              {\
                "transaction_id": "A unique ID for this transaction"\
                "original_description": "The original raw transaction description",\
                "previous_transactions": "(optional) Previously cleaned up transaction descriptions and the prior \
                category used that may be related to this transaction\
              }\
            ]}\n\
            For each transaction provided, follow these instructions:\n\
            (0) If previous_transactions were provided, see if the current transaction matches a previous one closely. \
                If it does, use the updated_description and category of the previous transaction exactly, \
                including capitalization and punctuation.\
            (1) If there is no matching previous_transaction, or none was provided suggest a better “updated_description” according to the following rules:\n\
            (a) Use all of your knowledge and information to propose a friendly, human readable updated_description for the \
              transaction given the original_description. The input often contains the name of a merchant name. \
              If you know of a merchant it might be referring to, use the name of that merchant for the suggested description.\n\
            (b) Keep the suggested description as simple as possible. Remove punctuation, extraneous \
              numbers, location information, abbreviations such as "Inc." or "LLC", IDs and account numbers.\n\
            (2) For each original_description, suggest a “category” for the transaction from the allowed_categories list that was provided.\n\
            (3) If you are not confident in the suggested category after using your own knowledge and the previous transactions provided, use the cateogry "' + FALLBACK_CATEGORY + '"\n\n\
            (4) Your response should be a JSON object and no other text.  The response object should be of the form:\n\
            {"suggested_transactions": [\
              {\
                "transaction_id": "The unique ID previously provided for this transaction",\
                "updated_description": "The cleaned up version of the description",\
                "category": "A category selected from the allowed_categories list"\
              }\
            ]}';

  const request = {
    system_instruction: {
      parts: [
        {
         text: "Act as an API that categorizes and cleans up bank transaction descriptions for a personal finance app. Reference the following list of allowed_categories:\n" + JSON.stringify(categoryList) + "\n" + tillerFormat
        }
      ],
    },
    contents: [{
      parts: [
        {
          text: JSON.stringify(transactionDict),
        }
      ],
    }],
    generationConfig: {
        "response_mime_type": "application/json",
    }
};

  const jsonRequest = JSON.stringify(request);

  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: jsonRequest,
    muteHttpExceptions: true,
  };

  var geminiURL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-exp:generateContent?key=' + GEMINI_API_KEY;
  var response = UrlFetchApp.fetch(geminiURL, options).getContentText();
  var parsedResponse = JSON.parse(response);

  if ("error" in parsedResponse) {
    Logger.log("Error from Gemini AI: " + parsedResponse["error"]["message"]);

    return null;
  } else {
    var apiResponse = JSON.parse(parsedResponse["candidates"][0]["content"]["parts"][0]["text"]);
    return apiResponse["suggested_transactions"];
  }
}
