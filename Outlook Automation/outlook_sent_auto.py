"""
outlook_sent_auto.py
───────────────────────────────────────────────────────────────────────────────
Production-ready Outlook “Sent Items” extractor
• Resolves shared/secondary mailboxes by SMTP or display name
• Manual date filtering (avoids fragile MAPI Restrict)
• Robust “Sent” folder discovery with fallback
• COM-safe extraction and export to XLSX (XlsxWriter)
Author : QM | Rev : 2025-05-15 (final prod patch)
"""

import sys
import re
import html
import pandas as pd
import win32com.client as win32
from datetime import datetime, timedelta
from pathlib import Path


# ─────────────────────────────────────────────── CLI input
acct_input = input("Outlook account SMTP or display name: ").strip().lower()
try:
    start_date = datetime.strptime(input("Start date  (YYYY-MM-DD): ").strip(), "%Y-%m-%d")
    end_date_in = datetime.strptime(input("End date    (YYYY-MM-DD): ").strip(), "%Y-%m-%d")
except ValueError:
    sys.exit("❌  Invalid date format. Use YYYY-MM-DD.")

if start_date > end_date_in:
    sys.exit("❌  Start date cannot be after end date.")

end_date = end_date_in + timedelta(days=1)  # inclusive filter logic


# ─────────────────────────────────────────────── Outlook connection
NS = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

# 1. Match account (store) by SMTP or display name
target_store = None
for store in NS.Folders:
    smtp_addr = getattr(store, "SMTPAddress", "").lower()
    if store.Name.lower() == acct_input or smtp_addr == acct_input:
        target_store = store
        break

if not target_store:
    sys.exit(f"❌  Mailbox ‘{acct_input}’ not found in current Outlook profile.")

# 2. Try default Sent Items folder, fallback to anything with “sent”
olFolderSentMail = 5
try:
    sent_folder = target_store.GetDefaultFolder(olFolderSentMail)
except Exception:
    sent_folder = None

if not sent_folder or sent_folder.Items.Count == 0:
    for fld in target_store.Folders:
        if "sent" in fld.Name.lower():
            sent_folder = fld
            break
    if not sent_folder:
        sys.exit("❌  No valid Sent folder found.")

print(f"📂 Using Sent folder: {sent_folder.Name} (store: {target_store.Name})")


# ─────────────────────────────────────────────── Manual date-based filtering
print("⏳ Filtering items manually by date (safe mode)...")
items = []
for item in sent_folder.Items:
    try:
        if getattr(item, "Class", None) != 43:  # Only MailItem
            continue
        sent = getattr(item, "SentOn", None)
        if sent is None:
            continue
        if sent.tzinfo:
            sent = sent.replace(tzinfo=None)
        if start_date <= sent < end_date:
            items.append(item)
    except Exception as e:
        print(f"⚠️  Skipped item due to error: {e}")

items.sort(key=lambda x: x.SentOn.replace(tzinfo=None), reverse=False)
print(f"🔎 Filtered items found: {len(items)}")
print(f"🔄 Sorted from oldest ({start_date.date()}) to newest ({end_date_in.date()})")



# ─────────────────────────────────────────────── Email content extraction
SEP_RX = re.compile(
    r"-----Original Message-----|^From:|^On .*? wrote:|^[-–]{8,}\s*Forwarded",
    re.I | re.M
)

records = []
for itm in items:
    try:
        sent_dt = itm.SentOn.replace(tzinfo=None)
        to_field = getattr(itm, "To", "") or ""
        subject = getattr(itm, "Subject", "") or ""

        body_txt = itm.Body or ""
        if not body_txt:
            raw_html = itm.HTMLBody or ""
            body_txt = html.unescape(re.sub("<[^>]+>", " ", raw_html))

        m = SEP_RX.search(body_txt)
        if m:
            sent_body = body_txt[:m.start()].strip()
            prev_body = body_txt[m.start():].strip()
            ps_match = re.search(r"From:\s*(.+)", prev_body, re.I)
            prev_sender = ps_match.group(1).strip() if ps_match else "Unknown"
        else:
            sent_body, prev_body, prev_sender = body_txt, "", "Not Found"

        records.append({
            "Date Sent": sent_dt,
            "Sent To": to_field,
            "Subject": subject,
            "Body of Sent Email": sent_body,
            "Previous Email Body": prev_body,
            "Previous Email Sender": prev_sender
        })
    except Exception as exc:
        print(f"⚠️  Skipping one item – {exc}")


# ─────────────────────────────────────────────── Excel export
if not records:
    sys.exit("⚠️  No matching emails in the specified range.")

df = pd.DataFrame(records)

# 🔍 Split 'Date Sent' into 'Date' and 'Time'
df["Date"] = df["Date Sent"].dt.strftime("%d-%m-%Y")
df["Time"] = df["Date Sent"].dt.strftime("%H:%M:%S")
df.drop(columns=["Date Sent"], inplace=True)

# Reorder columns for clarity
cols = ["Date", "Time", "Sent To", "Subject", "Body of Sent Email", "Previous Email Body", "Previous Email Sender"]
df = df[cols]

# 💾 Write to Excel
timestamp = datetime.now().strftime("%H%M%S")
outfile = (
    Path.cwd() /
    f"SentItems_{acct_input.replace('@', '_').replace('.', '_')}"
    f"_{start_date:%Y%m%d}_{end_date_in:%Y%m%d}_{timestamp}.xlsx"
)
df.to_excel(outfile, index=False, engine="xlsxwriter")
print(f"✅  Excel created → {outfile}")
