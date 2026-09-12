#!/usr/bin/env python3
"""
Tiller AI AutoCat - Amazon Order Invoice Resolver (Mac / Chrome Bridge)

Bypasses Amazon's omission of specific product names in confirmation emails by
fetching the printable order invoice directly via an existing active Google Chrome
session using Apple Events, extracting exact item names, and categorizing them with Gemini.

Requirements:
1. Google Chrome on macOS with your Amazon account logged in.
2. In Chrome: View > Developer > Allow JavaScript from Apple Events (enabled).
3. GEMINI_API_KEY set in environment or passed via --api-key.
"""

import argparse
import json
import os
import subprocess
import sys
import urllib.request
import urllib.error

DEFAULT_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = "gemini-flash-latest"

def resolve_amazon_invoice(order_id):
    """
    Opens the printable invoice URL in a background Chrome tab, waits for DOM load,
    runs a JavaScript extractor to get product titles, and closes the tab.
    """
    url = f"https://www.amazon.com/gp/css/summary/print.html?orderID={order_id}"

    js_code = """
    (function() {
        var title = document.title;
        var bodyText = document.body.innerText;
        var items = [];
        
        // Find product names in printable invoice tables
        var rows = document.querySelectorAll('table tr');
        for (var r of rows) {
            var text = r.innerText.trim();
            if (text.includes('$') && (text.includes('Sold by') || text.includes('Condition') || text.includes('Supplied by'))) {
                var lines = text.split('\\n').map(l => l.trim()).filter(l => l.length > 0);
                if (lines.length > 0) {
                    items.push(lines[0]);
                }
            }
        }
        
        // Fallback for alternate invoice structures
        if (items.length === 0) {
            var boldCells = document.querySelectorAll('table td b');
            for (var b of boldCells) {
                var t = b.innerText.trim();
                if (t.length > 5 && !t.includes('Order') && !t.includes('Total') && !t.includes('Ship')) {
                    items.push(t);
                }
            }
        }
        
        // Fallback for standard order details view
        if (items.length === 0) {
            var links = document.querySelectorAll('a[href*="/dp/"], a[href*="/gp/product/"]');
            for (var l of links) {
                var t = l.innerText.trim();
                if (t.length > 5 && !items.includes(t)) items.push(t);
            }
        }

        return JSON.stringify({
            title: title,
            url: window.location.href,
            items: items,
            orderId: '""" + order_id + """',
            snippet: bodyText.substring(0, 1200)
        });
    })()
    """

    applescript = f"""
    tell application "Google Chrome"
        tell front window
            set newTab to make new tab with properties {{URL:"{url}"}}
            set maxWait to 30
            repeat while (loading of newTab) and maxWait > 0
                delay 0.2
                set maxWait to maxWait - 1
            end repeat
            delay 1.0
            
            set res to execute newTab javascript {json.dumps(js_code)}
            close newTab
            return res
        end tell
    end tell
    """

    res = subprocess.run(["osascript", "-e", applescript], capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(
            f"Chrome AppleScript execution failed: {res.stderr.strip()}.\n"
            "Make sure Google Chrome is open and 'View > Developer > Allow JavaScript from Apple Events' is checked."
        )
    
    try:
        return json.loads(res.stdout.strip())
    except Exception:
        raise RuntimeError(f"Unexpected response from Chrome: {res.stdout}")

def categorize_with_gemini(order_data, categories=None, api_key=None):
    """
    Sends resolved product titles to Gemini (gemini-flash-latest) to produce a clean description and category.
    """
    api_key = api_key or DEFAULT_API_KEY
    if not api_key:
        print("[!] Error: No Gemini API key provided.")
        print("    Set it via environment: export GEMINI_API_KEY='your-key'")
        print("    Or pass it via: --api-key <YOUR_KEY>")
        sys.exit(1)

    if not categories:
        categories = ["Home Supplies", "Hardware", "Groceries", "Electronics", "Office", "Home Maintenance", "Personal", "Miscellaneous"]

    items_str = ", ".join(order_data.get("items", []))
    if not items_str:
        items_str = order_data.get("snippet", "")[:300]

    system_prompt = f"""You are an elite financial forensics agent.
Clean up the raw Amazon product item into a concise, professional ledger description (e.g. "Duck Indoor/Outdoor Carpet Tape" or "Anker USB-C Charger").
If multiple items were ordered, combine them cleanly (e.g. "Duck Carpet Tape & Anker Cable").

STRICT RULES:
- NEVER prefix with "Amazon:".
- NEVER use generic labels like "Shopping", "Transfer", or "Order".
- Assign the best matching category from this VERBATIM list: {json.dumps(categories)}.

Return JSON format:
{{"updated_description": "Clean Name", "category": "Exact Category From List"}}"""

    payload = {
        "contents": [{"parts": [{"text": f"Amazon Order ID: {order_data.get('orderId')}\nRaw Products: {items_str}\nOrder Snippet: {order_data.get('snippet', '')[:400]}"}]}],
        "system_instruction": {"parts": [{"text": system_prompt}]},
        "generationConfig": {"response_mime_type": "application/json"}
    }

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={api_key}"
    req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers={"Content-Type": "application/json"})
    
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read().decode("utf-8"))
        text = data["candidates"][0]["content"]["parts"][0]["text"]
        return json.loads(text)

def sync_with_sheet(webapp_url, api_key=None):
    """
    Pulls pending Amazon transactions from the Tiller Sheet Web App, resolves them via Chrome,
    categorizes them with Gemini, and posts updates back to the Sheet.
    """
    print(f"[*] Querying pending transactions from: {webapp_url}")
    req = urllib.request.Request(webapp_url)
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read().decode("utf-8"))

    pending = data.get("pending_transactions", [])
    categories = data.get("categories", [])
    print(f"[+] Found {len(pending)} pending Amazon transaction(s).")
    
    if not pending:
        print("[+] Nothing to update. Sheet is already clean!")
        return

    updates = []
    for txn in pending:
        order_id = txn.get("order_id")
        desc = txn.get("original_description")
        amt = txn.get("amount")
        
        print(f"\n--- Processing: {desc} (${amt}) ---")
        if not order_id:
            print("[!] No Order ID found for transaction. Skipping.")
            continue
            
        print(f"[*] Fetching invoice for Order #{order_id} via Chrome...")
        try:
            order_info = resolve_amazon_invoice(order_id)
            print(f"[+] Scraped item(s): {order_info.get('items')}")
            
            cat_result = categorize_with_gemini(order_info, categories=categories, api_key=api_key)
            print(f"[+] Gemini -> Description: '{cat_result.get('updated_description')}', Category: '{cat_result.get('category')}'")
            
            updates.append({
                "transaction_id": txn.get("transaction_id"),
                "updated_description": cat_result.get("updated_description"),
                "category": cat_result.get("category")
            })
        except Exception as err:
            print(f"[!] Error resolving order #{order_id}: {err}")

    if updates:
        print(f"\n[*] Posting {len(updates)} update(s) back to Tiller Sheet...")
        post_payload = json.dumps({"updates": updates}).encode("utf-8")
        post_req = urllib.request.Request(webapp_url, data=post_payload, headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(post_req) as resp:
            res_data = json.loads(resp.read().decode("utf-8"))
            print(f"[+] Server response: {res_data}")
            print("[✓] Sheet successfully updated!")

def main():
    parser = argparse.ArgumentParser(description="Resolve and categorize Amazon orders for Tiller using Chrome & Gemini")
    parser.add_argument("--order", help="Single Amazon Order ID to resolve and categorize (e.g. 113-8099586-4134612)")
    parser.add_argument("--sync", help="Tiller Google Apps Script Web App URL for automatic sync")
    parser.add_argument("--api-key", help="Gemini API Key (defaults to GEMINI_API_KEY environment variable)")
    args = parser.parse_args()

    api_key = args.api_key or DEFAULT_API_KEY

    if args.order:
        print(f"[*] Fetching invoice for Order #{args.order} via Chrome...")
        order_info = resolve_amazon_invoice(args.order)
        print(f"[+] Found {len(order_info.get('items', []))} item(s): {order_info.get('items')}")
        
        print(f"[*] Categorizing with Gemini ({GEMINI_MODEL})...")
        cat_result = categorize_with_gemini(order_info, api_key=api_key)
        print("\n=== Result ===")
        print(f"Description: {cat_result.get('updated_description')}")
        print(f"Category:    {cat_result.get('category')}")
        return

    if args.sync:
        sync_with_sheet(args.sync, api_key=api_key)
        return

    parser.print_help()

if __name__ == "__main__":
    main()
