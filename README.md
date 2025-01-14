CREDIT: Joshua van Niekerk (and ChatGPT)

### Description:
I have leveraged ChatGPT to code up several scripts for me (for my Workflow Coordinator role at Paragon).
These scripts serve the purpose of calculating and identifying what stock to order (from the warehouse) before the print reaches the machines in production.

### Specification:
**FilterDataAndCreateSummary** (The main method) does several things, namely:
1. Removes duplicate heading rows
2. Exports desired columns to a new worksheet called FilteredData
3. Gets outers based on *CORP_CD*
4. Highlights *WORK ORDERS* and *INSERTS* where inserts > 4
5. Highlights *REMAKES* (yellow)
6. Calculates a summary

### Dependencies:
To perform/run any of these scripts, you need:
1. To rename the active worksheet *Special*
2. The OUTERSKEY worksheet
    - Open up outers-key.xls
    - Copy the OUTERSKEY data (from the outers-key.xls)
    - Paste it in a new worksheet on the Paragon SLA excel spreadhseet
    - Rename the worksheet *OUTERSKEY*

### First Time Setup To Run Macros:
Enable the Developer Tab: If not already enabled, enable the Developer tab in Excel.

Go to File>Options>Customize Ribbon.
Check Developer in the right-hand list.

### How To Add A New Script:
1. Alt + F11 to open the VBA editor.
2. Insert a new module (Insert > Module).
3. Paste the code/script into the module.
4. Close the VBA editor and return to Excel.
5. Run the macro 
    - Alt + F8 
    - Select **FilterDataAndCreateSummary**
    - (Optional) Click the options button to add a shortcut key to run the script.
    - Click "Run".

### Send scripts via email
.vb scripts cannot be attached to an email because the filename flags up as potentially dangerous.
To get around this, we can rename each file extention to .txt
EXAMPLE: rename calc-sum-outers.vb to calc-sum-outers.txt
Then we can attach the file to an email