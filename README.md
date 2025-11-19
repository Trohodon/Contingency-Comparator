# Contingency-Comparator
Contingency Comparison Tool - README

DESCRIPTION
This project is a Python-based tool for comparing contingency result tables between multiple sheets of an Excel workbook. It is designed for engineering workflows where contingency results shift positions between cases, such as “Before Opened vs After Opened” and “Before Closed vs After Closed.” The tool allows the user to visually compare differences in the ACCA Long Term, ACCA, and DCwAC tables across two sheets at a time. Two independent comparison windows are provided inside the GUI so the user can evaluate two scenarios simultaneously.

FEATURES
- Load an Excel workbook containing multiple sheets of contingency results.
- Automatically extract the three main tables: ACCA Long Term, ACCA, and DCwAC.
- Compare any two sheets selected by the user.
- Side-by-side comparison windows inside the GUI.
- Compare percent loading changes and identify whether contingency-event rows appear in both sheets or only one.
- Clean separation of logic, GUI, and entry point.

PROJECT STRUCTURE
main.py        - Entry point for the application. Launches the GUI.
gui.py         - All GUI code using tkinter. Contains the two comparison windows and controls for selecting sheets.
program.py     - All Excel parsing and comparison logic. No GUI code.

HOW IT WORKS
1. program.py loads the workbook, reads all sheets, detects the table blocks, and stores the extracted ACCA Long Term, ACCA, and DCwAC tables for each sheet.
2. gui.py provides two comparison panels. Each panel allows the user to choose two sheets and then display the comparison results in tabbed views.
3. main.py simply initializes and runs the GUI.

INSTALLATION REQUIREMENTS
- Python 3.9 or newer
- Required Python packages:
  pandas
  openpyxl

INSTALLATION STEPS
1. Install Python.
2. Install dependencies using:
   pip install pandas openpyxl
3. Run the application using:
   python main.py

USAGE
1. Open the program.
2. Click “Open Excel Workbook.”
3. Select the Excel file containing the contingency sheets.
4. Use the drop-down menus in each comparison window to choose which two sheets to compare.
5. Click “Compare” in each window to update the results.

PURPOSE
This tool is intended for engineers who need to quickly evaluate how contingencies shift between different planning or operations cases. It identifies how percent loading changes, whether any contingencies move up or down table sections, and whether contingencies disappear or appear between cases.

CONTRIBUTING
Users may modify or extend the parsing or GUI logic to meet specific workflow requirements. Pull requests and feature suggestions are welcome.

LICENSE
Open license may be used or defined based on the user’s preference.

END OF FILE
