# ⚡ GOD MODE BOOSTER – Ultimate Windows Optimizer

![God Mode Booster Banner](https://img.shields.io/badge/PowerShell-Automation-blue?style=flat-square)  
A fully interactive, bulletproof PowerShell script that deep-cleans your Windows system without touching personal files. Optimized for **performance freaks**, **tinkerers**, and **power users**.

---

## 🧩 What This Script Does

| Category                | Action Description                                                                 |
|------------------------|-------------------------------------------------------------------------------------|
| 💻 System Cleanup       | Deletes temp files, logs, prefetch, cache, memory dumps                           |
| 🌐 Browser Cleanup      | Optionally clears **history & cookies** from Chrome, Edge, Firefox, etc.          |
| 🧠 RAM Optimization     | Flushes standby RAM (if `emptystandbylist.exe` is present)                         |
| 🧼 Disk Cleanup         | Runs Windows' own `cleanmgr.exe` with `/verylowdisk` silently                      |
| 🛑 Bloatware Termination | Kills background apps like OneDrive, Teams, Skype, Xbox, Widgets, etc.            |
| 📉 Performance Boost    | Stops noisy services like SysMain, WSearch, DiagTrack                             |
| 🗑️ Recycle Bin          | Empties system trash securely                                                      |
| 🔄 Post-cleanup Prompt  | Prompts for reboot to finalize optimizations                                       |
| 📊 Stats Display        | Shows **exact MB/GB** freed at the end                                             |

---

## 🔧 Requirements

- 🪟 **Windows 10/11**
- 🔐 **Run as Administrator**
- ⚙️ Optional: `emptystandbylist.exe` (for RAM flush)

---

## 📥 Installation & Usage

### ✅ Option 1: Run the `.ps1` Script
```powershell
Right-click > Run with PowerShell (as Administrator)
