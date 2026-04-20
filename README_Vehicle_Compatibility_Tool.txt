VEHICLE COMPATIBILITY TOOL - STEP BY STEP GUIDE
===============================================

This package helps you:
1. Import or paste a fleet vehicle list
2. Match vehicles against supported compatibility lists
3. Generate a clean business report showing:
   - total vehicles
   - total brands
   - total unique models
   - unsupported vehicles
   - supported vehicles with no fuel reading
   - brand-wise gaps
   - rows that need review for accuracy


FILES INCLUDED IN THIS PACKAGE
==============================

1. Vehicle_Compatibility_UI.py
   The main application that you double-click or run.

2. Vehicle_Compatibility_Matcher.py
   The matching engine used by the UI.

3. Vehicle_Compatibility_Workbench.xlsx
   The workbook that stores support lists, input rows, and results.

4. README_Vehicle_Compatibility_Tool.txt
   This guide.


BEFORE YOU START
================

IMPORTANT:
- Keep all 3 main files in the SAME folder.
- Do not rename the files after extracting the zip.
- Do not move only one file by itself.
- The UI expects the workbook and matcher to be next to it.


WHAT YOUR COLLEAGUES NEED INSTALLED
===================================

Python is required.

Recommended Python version:
- Python 3.10 or 3.11

Required Python packages:
- pandas
- openpyxl
- rapidfuzz

If these are not installed, follow the installation steps below.


STEP 1 - EXTRACT THE ZIP
========================

1. Save the zip file to your computer.
2. Right-click the zip file.
3. Choose Extract / Unzip.
4. Open the extracted folder.
5. Confirm that you can see these files together in one folder:
   - Vehicle_Compatibility_UI.py
   - Vehicle_Compatibility_Matcher.py
   - Vehicle_Compatibility_Workbench.xlsx
   - README_Vehicle_Compatibility_Tool.txt


STEP 2 - INSTALL PYTHON (IF NEEDED)
===================================

WINDOWS
-------
1. Download Python from the official Python website.
2. Run the installer.
3. IMPORTANT: tick the checkbox that says:
   Add Python to PATH
4. Complete the installation.

MAC
---
Option A: Python already available
1. Open Terminal.
2. Type:
   python3 --version
3. If a version number appears, Python is installed.

Option B: Install Python
1. Install Python 3 from the official Python website, OR
2. Install using Homebrew if your team uses it.


STEP 3 - INSTALL REQUIRED PACKAGES
==================================

WINDOWS
-------
1. Open Command Prompt.
2. Change folder to the extracted folder.
   Example:
   cd C:\Users\YourName\Downloads\Vehicle_Compatibility_Tool
3. Run:
   pip install pandas openpyxl rapidfuzz

If pip does not work, try:
   py -m pip install pandas openpyxl rapidfuzz

MAC
---
1. Open Terminal.
2. Change folder to the extracted folder.
   Example:
   cd ~/Downloads/Vehicle_Compatibility_Tool
3. Run:
   python3 -m pip install pandas openpyxl rapidfuzz

If your Mac uses pip3 instead, run:
   pip3 install pandas openpyxl rapidfuzz


STEP 4 - OPEN THE TOOL FOLDER
=============================

Make sure all files remain together in the same folder.

Folder should contain:
- Vehicle_Compatibility_UI.py
- Vehicle_Compatibility_Matcher.py
- Vehicle_Compatibility_Workbench.xlsx


STEP 5 - RUN THE TOOL
=====================

WINDOWS
-------
Method 1:
1. Open Command Prompt.
2. Go to the extracted folder.
3. Run:
   py Vehicle_Compatibility_UI.py

Method 2:
1. Right-click Vehicle_Compatibility_UI.py
2. Open with Python, if your system is configured for that.

MAC
---
1. Open Terminal.
2. Go to the extracted folder.
3. Run:
   python3 Vehicle_Compatibility_UI.py

The UI window should open.


STEP 6 - IMPORT OR PASTE THE FLEET LIST
=======================================

You have 2 ways to provide vehicle data.

OPTION A - IMPORT A FILE
------------------------
1. Click the button: Paste vehicle rows
2. Choose your fleet file
   Supported file types usually include:
   - .xlsx
   - .csv
   - .txt
3. The tool will try to detect the main vehicle column automatically.
4. Review the imported text in the input area.

OPTION B - PASTE DIRECTLY
-------------------------
1. Copy vehicle entries from Excel or another file.
2. Click inside the input box.
3. Paste the entries.
4. Each vehicle should appear as one line.

Tip:
If your source file has separate brand/model/year columns, combine them before pasting or use the import function if the source file already contains enough detail.


STEP 7 - RUN MATCHING
=====================

1. After importing or pasting, click:
   Run Match
2. The tool will:
   - write the input rows into the workbook
   - analyze the vehicles
   - match them against supported lists
   - create final result sheets
   - generate a clean report workbook

Wait until the status message says the run is complete.


STEP 8 - OPEN THE REPORT
========================

After the run finishes, open:
- Final_Report_Clean.xlsx

This is the clean business report for sharing and review.

You can also open:
- Vehicle_Compatibility_Workbench.xlsx

The workbench contains the underlying detailed sheets and results.


WHAT THE CLEAN REPORT CONTAINS
==============================

1. Summary_KPI
   Main business metrics such as:
   - total vehicles in list
   - total unique vehicles
   - total brands in list
   - total unique models
   - vehicles supported but with no fuel reading
   - vehicles not included in any support list
   - support coverage
   - strong / possible / review-needed counts

2. Final_Report
   Main row-level clean report.

3. Unsupported_Vehicles
   Vehicles not found reliably in support lists.

4. No_Fuel_Vehicles
   Vehicles supported by one or more lists but without fuel reading.

5. Brand_Gaps
   Brand-wise models that appear missing from the support lists.

6. Needs_Review
   Rows that require manual review because confidence is lower or the input is ambiguous.


HOW TO INTERPRET RESULTS
========================

STRONG MATCH
- High confidence
- Good candidate for acceptance

POSSIBLE MATCH
- Reasonable candidate
- Check if the raw input was generic or incomplete

REVIEW NEEDED
- Needs manual confirmation
- Do not treat as final without checking

NO RELIABLE MATCH / UNSUPPORTED
- Vehicle could not be matched safely
- Treat as not supported unless reviewed further

NO FUEL VEHICLES
- Vehicle may be supported for some parameters
- But fuel level / fuel reading is not available in the matched support list(s)


IMPORTANT GOOD PRACTICES
========================

1. Always review:
   - Unsupported_Vehicles
   - No_Fuel_Vehicles
   - Needs_Review

2. Do not rely only on one top match if the source input is vague.

3. Inputs like these may need extra review:
   - generic names such as "Bus", "Truck", "Van"
   - local shorthand names
   - spelling variations
   - body-type descriptions instead of exact model names

4. If your company uses recurring shorthand vehicle names, update the workbook alias or rules sheets so future runs become more accurate.


RECOMMENDED REVIEW FLOW
=======================

For each new fleet upload:
1. Run the tool
2. Open Final_Report_Clean.xlsx
3. Check Summary_KPI first
4. Check Unsupported_Vehicles second
5. Check No_Fuel_Vehicles third
6. Check Needs_Review fourth
7. Share report with technical reviewer if needed
8. If repeated naming patterns are discovered, update workbook rules for better future accuracy


TROUBLESHOOTING
===============

PROBLEM: UI does not open
-------------------------
Cause:
- Python not installed correctly
- missing packages

Fix:
- confirm Python version
- install packages again
- run from Terminal / Command Prompt so you can see the error

PROBLEM: Error says module not found
------------------------------------
Fix:
Install packages again:
- pandas
- openpyxl
- rapidfuzz

WINDOWS:
   py -m pip install pandas openpyxl rapidfuzz

MAC:
   python3 -m pip install pandas openpyxl rapidfuzz

PROBLEM: Workbook or matcher file not found
-------------------------------------------
Cause:
- files are not in same folder
- file was renamed

Fix:
- keep all files together
- use the exact names included in this package

PROBLEM: Imported fleet looks messy
-----------------------------------
Fix:
- try Paste vehicle rows instead of copy-paste
- clean source file so the main vehicle detail appears in one column when possible

PROBLEM: Too many vehicles show as review needed
------------------------------------------------
Cause:
- input descriptions are too generic
- supported list naming differs from fleet naming

Fix:
- improve fleet input wording when possible
- update workbook alias / pattern sheets for recurring names
- review unsupported and ambiguous entries manually


SHARING WITH COLLEAGUES
=======================

To share this tool with colleagues:
1. Share the zip file as-is
2. Ask them to extract it fully
3. Tell them to read this README first
4. Remind them not to rename files after extraction


FINAL NOTE
==========

This tool is designed to be conservative.
That means it prefers:
- Review Needed
- No Reliable Match

instead of making a confident-looking wrong match.

That behavior is intentional and safer for fleet-scale compatibility analysis.
