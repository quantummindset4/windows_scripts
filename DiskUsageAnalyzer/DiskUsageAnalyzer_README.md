📊 File Size Analyzer PowerShell Script
This PowerShell script scans a given directory (and all its subdirectories), identifies all files, sorts them by size (descending), and exports the results to a CSV file for easy review and cleanup planning.

🔧 Features
Prompts the user to input a folder path

Defaults to scanning the entire C:\ drive if no path is provided

Lists files with:

📁 Full path

📅 Last modified date

📦 File size in MB

Sorts files by size (largest first)

Exports to a FileAnalysis.csv file in the script's own directory

▶️ How to Use
Save the script as DiskFileAnalyzer.ps1

Right-click PowerShell → Run as Administrator (recommended for full access)

Execute the script in your terminal:

powershell
Copy
Edit
.\DiskFileAnalyzer.ps1
Enter a folder path when prompted
(or press Enter to scan the entire C:\ drive)

The output CSV will be saved to the same folder where the script is located.

📝 Sample Output (CSV Columns)
SizeMB	LastModified	Path
1024.50	2024-12-10 14:32:00	C:\Users...\Downloads\bigfile.zip
987.12	2024-11-22 10:04:15	C:\Program Files...\some.log

📁 Output File
Filename: FileAnalysis.csv

Location: Same folder as the script

⚠️ Notes
This script ignores any permission-denied errors silently

Execution time depends on number of files and drive speed

For best performance, avoid scanning C:\ unless needed

