
# Tiller AI AutoCat
Apps Script code to use Gemini AI to automatically categorize financial transactions (designed to work with Tiller Finance Feeds and Google Sheets)

## About
- This is a script that is designed to work with the Tiller finance product to automatically categorize and clean up the Description column of your transactions (so you don't have to do it all manually!).
- It will only touch transactions that don't have a Category set, or those that have generic/unhelpful descriptions (like raw "AMAZON RETA*" or "PAYPAL INST XFER").
- It works by trying to find how you've previously categorized transactions like the one it's working on, correlating with receipts in your Gmail, sending that context to Gemini AI, and asking it to do its magic.
- It will set the Category and Description field based on what comes back.
- It will pick the best valid category from your Category list, or fall back to a category you specify if it gets confused.
- If you want to mark transactions that have been modified by this code, add a column to your Transactions sheet called "AI AutoCat" - it will mark transactions it's modified by writing TRUE into this column.
- This works for me, and I've tried to make it somewhat generic so it works for others -- but I DISCLAIM ALL RESPONSIBILITY IF IT MESSES ANYTHING UP IN YOUR SHEET. You can always undo or revert to a previous version.
- Given how sensitive this is to data, any and all feedback about how it's working (or not) is greatly appreciated.

## Key Features
- **Gmail Receipt Hunting:** Pulls data from Amazon, Venmo, PayPal, and eBay.
- **Amazon Subscribe & Save Fix:** Correlates delivery dates when prices are hidden in emails.
- **Venmo Memo Cleanup:** Turns "Chil crisps" into "Chili Crisp" and assigns "Groceries".
- **Prefix Removal:** No more "Amazon: " or "Venmo: " prefixes; just clean item names.
- **Batch Processing:** High-performance logic designed to handle dozens of rows without crashing.

## Demo Video
- You can see this working with some sample data here: [Demo Link](https://drive.google.com/file/d/16ROtqWboSOaNfgKGs0hUSjc3heGqFPBD/view?usp=drive_link)

## Installation Instructions
- First, you need to get a Gemini AI API Key to use. Sign up as a developer with Google and get a secret key.
- From your Tiller connected Google Sheet, go to Extensions --> Apps Script
- If you don't have any existing Apps Script, you should just see `Code.gs` in the Files section on the left.
- Use the + button to add two new files called `gviz.gs` and `ai_autocat.gs`.
- Copy and paste the contents of the files here into those files.
- Add (or change if you have one already) an `onOpen` function to your `Code.gs` file. This adds a menu item to call the AI AutoCat code.
- Modify `ai_autocat.gs` to use your Gemini AI API Key.
- Modify `ai_autocat.gs` to use the `FALLBACK_CATEGORY` you want to use (this must be a valid category, or the empty string).

## Usage Instructions
- After installing the script, refresh your Tiller sheet. You should see a new menu item appear called "AI AutoCat" after a few seconds. You can run the AI autocat code manually from this menu item.
- If you want, you can also add a trigger to automatically run the AI AutoCat code nightly. See instructions here: [Apps Script Triggers](https://developers.google.com/apps-script/guides/triggers/installable). The function you want to run is `categorizeUncategorizedTransactions`.
