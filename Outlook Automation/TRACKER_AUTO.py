import sys
import re
import html
import pandas as pd
import win32com.client as win32
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
from tkinter.ttk import Button, Label, Entry, Style
import subprocess
import time
import psutil

# ───────────────────────────────────────────── GUI for Input ─────────────────────────────────────────────
def run_export(account_input, start_date_input, end_date_input):
    try:
        status_text.set("📅 Validating input dates...")
        app.update()
        start_date = datetime.strptime(start_date_input, "%Y-%m-%d")
        end_date_in = datetime.strptime(end_date_input, "%Y-%m-%d")
        if start_date > end_date_in:
            raise ValueError("Start date after end date")
    except:
        messagebox.showerror("Date Format Error", "Please use format YYYY-MM-DD for both dates.")
        return

    end_date = end_date_in + timedelta(days=1)

    try:
        outlook_running = any("OUTLOOK.EXE" in p.name().upper() for p in psutil.process_iter())
        if not outlook_running:
            status_text.set("🚀 Launching Outlook...")
            app.update()
            subprocess.Popen("start outlook", shell=True)
            time.sleep(6)  # wait for Outlook to start
        else:
            status_text.set("🔄 Outlook is already running.")
            app.update()
    except Exception as e:
        messagebox.showerror("Outlook Launch Error", f"Unable to launch Outlook.\n{str(e)}")
        return

    try:
        status_text.set("🔌 Connecting to Outlook COM interface...")
        app.update()
        outlook = win32.Dispatch("Outlook.Application")
        NS = outlook.GetNamespace("MAPI")
    except Exception as e:
        messagebox.showerror("Outlook COM Error", f"Cannot connect to Outlook.\n\n{str(e)}")
        return

    status_text.set("📡 Searching for mailbox in Outlook profile...")
    app.update()
    target_store = None
    for store in NS.Folders:
        smtp_addr = getattr(store, "SMTPAddress", "").lower()
        if store.Name.lower() == account_input.lower() or smtp_addr == account_input.lower():
            target_store = store
            break

    if not target_store:
        messagebox.showerror("Mailbox Error", f"Mailbox '{account_input}' not found in Outlook.")
        return

    status_text.set("📁 Locating Sent Items folder...")
    app.update()
    try:
        sent_folder = target_store.GetDefaultFolder(5)
    except:
        sent_folder = None

    if not sent_folder or sent_folder.Items.Count == 0:
        for fld in target_store.Folders:
            if "sent" in fld.Name.lower():
                sent_folder = fld
                break
        if not sent_folder:
            messagebox.showerror("Folder Error", "No valid Sent folder found.")
            return

    status_text.set("⚙️ Filtering emails...")
    app.update()

    items = []
    for item in sent_folder.Items:
        try:
            if getattr(item, "Class", None) != 43:
                continue
            sent = getattr(item, "SentOn", None)
            if sent is None:
                continue
            if sent.tzinfo:
                sent = sent.replace(tzinfo=None)
            if start_date <= sent < end_date:
                items.append(item)
        except:
            continue

    items.sort(key=lambda x: x.SentOn.replace(tzinfo=None))
    status_text.set(f"✅ Found {len(items)} emails. Extracting data...")
    app.update()

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
        except:
            continue

    if not records:
        messagebox.showinfo("No Data", "No matching emails found.")
        return

    df = pd.DataFrame(records)
    df["Date"] = df["Date Sent"].dt.strftime("%d-%m-%Y")
    df["Time"] = df["Date Sent"].dt.strftime("%H:%M:%S")
    df.drop(columns=["Date Sent"], inplace=True)

    cols = ["Date", "Time", "Sent To", "Subject", "Body of Sent Email", "Previous Email Body", "Previous Email Sender"]
    df = df[cols]

    status_text.set("💾 Writing to Excel...")
    app.update()

    timestamp = datetime.now().strftime("%H%M%S")
    outfile = (
        Path.cwd() /
        f"SentItems_{account_input.replace('@', '_').replace('.', '_')}"
        f"_{start_date:%Y%m%d}_{end_date_in:%Y%m%d}_{timestamp}.xlsx"
    )
    df.to_excel(outfile, index=False, engine="xlsxwriter")

    status_text.set("✅ Export completed")
    messagebox.showinfo("Success", f"Excel created:\n{outfile}")


# ───────────────────────────────────────────── GUI Layout ─────────────────────────────────────────────
app = tk.Tk()
app.title("🟢 SENT EXTRACTOR – OUTLOOK AUTOMATION")
app.geometry("480x380")
app.configure(bg="#0d0d0d")

style = Style()
style.theme_use('clam')
style.configure('TLabel', background="#0d0d0d", foreground="#00ffcc", font=("Consolas", 10))
style.configure('TButton', background="#1f1f1f", foreground="#00ffcc", font=("Consolas", 10))
style.configure('TEntry', fieldbackground="#1f1f1f", foreground="#00ffcc")

Label(app, text="Enter Outlook SMTP or Display Name:").pack(pady=4)
account_entry = Entry(app, width=45)
account_entry.pack()

Label(app, text="Start Date (YYYY-MM-DD):").pack(pady=4)
start_entry = Entry(app, width=25)
start_entry.pack()

Label(app, text="End Date (YYYY-MM-DD):").pack(pady=4)
end_entry = Entry(app, width=25)
end_entry.pack()

Button(app, text="💾 Run Export", command=lambda: run_export(
    account_entry.get(),
    start_entry.get(),
    end_entry.get()
)).pack(pady=12)

status_text = tk.StringVar()
status_label = Label(app, textvariable=status_text, font=("Consolas", 10, "italic"))
status_label.pack(pady=8)

Button(app, text="❌ Close", command=app.quit).pack(pady=10)

app.mainloop()
