📄 README — Outlook Sent Items Extractor (Dev Version)
────────────────────────────────────────────────────────────

🛠 TOOL NAME:
TRACKER_AUTO.py — GUI App to Export Outlook Sent Emails

👤 AUTHOR:
QM — Automation Engineer

📦 CONTENTS:
This folder contains source code + requirements to run a Python GUI tool that connects to Outlook, extracts “Sent Items” based on a date range, and exports the data to Excel.

────────────────────────────────────────────────────────────
📁 INCLUDED FILES:
- TRACKER_AUTO.py         → Main script
- requirements.txt        → All required packages
- TA.ico                  → App icon (optional)
- *.xlsx                  → Sample exported Excel files

────────────────────────────────────────────────────────────
🔧 SETUP (1 TIME — DEV SYSTEM ONLY):

Open Command Prompt in this folder and run:

1️⃣ Create a virtual environment:
    python -m venv outlookenv

2️⃣ Activate it:
    outlookenv\\Scripts\\activate

3️⃣ Install dependencies:
    pip install -r requirements.txt

4️⃣ Launch the tool:
    python TRACKER_AUTO.py

────────────────────────────────────────────────────────────
🧠 WHAT THE TOOL DOES:

- Launches a simple GUI
- Takes:
   • Outlook email (SMTP/display name)
   • Start + End date
- Extracts all matching Sent emails
- Outputs Excel with:
   • Date, Time, Recipients
   • Subject
   • Email body
   • Previous sender (if reply)

────────────────────────────────────────────────────────────
⚠ REQUIREMENTS:

- Windows with Microsoft Outlook installed (must be signed in)
- Python 3.8+ (recommended: 64-bit)
- Network access NOT required
- Excel not needed (just for opening the .xlsx)

────────────────────────────────────────────────────────────
🧼 CLEANUP (Optional):
To deactivate environment:
    deactivate

To remove:
    Delete the folder `outlookenv/`

────────────────────────────────────────────────────────────
🧪 TROUBLESHOOTING:

❌ Error: `Server execution failed`
→ Outlook must be open. Open Outlook first, then run the script.

❌ COM Error
→ Try running the script as administrator or ensure you're not blocked by security policies.

❌ Nothing exported?
→ Ensure correct email account and valid date range.

────────────────────────────────────────────────────────────
📧 SUPPORT:

QM | Slack/Teams | Ask if you want features like:
- Subject keyword filters
- Export to CSV instead of Excel
- Auto-email output

────────────────────────────────────────────────────────────
✔ BUILT WITH:
Python + Tkinter + Pandas + Win32com + XlsxWriter
