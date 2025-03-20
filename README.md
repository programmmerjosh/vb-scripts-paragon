CREDIT: Joshua van Niekerk (and ChatGPT)

### Description:
I have leveraged ChatGPT to code up several scripts for me (for my Workflow Coordinator role at Paragon).
These scripts serve the purpose of calculating and identifying what stock to order (from the warehouse) before the print reaches the machines in production.

### What The Scripts Do:
**MergeMySheets** takes all the relevent data from multiple sheets and places it together onto one worksheet called *special* and proceeds to call the next script.
**FilterDataAndCreateSummary** (The main method) does several things, namely:
1. Exports the desired columns to a new worksheet called FilteredData
2. Gets outers based on *CORP_CD*
3. Highlights *WORK ORDERS* (red) where inserts > 4
4. Highlights *WORK ORDERS* (orange) where we always need to order those particular outers
5. Highlights *REMAKES* (yellow)
6. Calculates a summary
7. Uses the *previous* worksheet to compare with FilteredData to find new entries (if *previous* exists)
8. Compares FilteredData (new-list) with *previous* to create an enclosed work order list
9. Deletes the *special* and *previous* worksheets

### How To Execute The Scripts And What You Need Before You Do:
To execute both script(s), follow these instructions:
1. Add the OUTERSKEY worksheet
    - Open up outers-key.xls
    - Copy the OUTERSKEY data (from the outers-key.xls)
    - Paste it in a new worksheet on the Paragon SLA excel spreadhseet
    - Rename the worksheet *OUTERSKEY*
2. Rename every worksheet(s) that you want to be included to "s1", "s2", "s3", and so forth. Up to "s8".
    - The order doesn't matter
    - It doesn't matter if the "S" is uppercase or lowercase.
    - So long as the sheetname has an "s" followed by a number 1-8.
3. Run the MergeMySheets script
    - You can go to *View* Tab and click on *View Macros* OR press *Alt + F8*
    - (Optional) setup a shorcut key by clicking on options
    - Click on *MergeMySheets* and click *Run* OR close the Macros window and use the shortcut key (if you have set one up)
    - **IMPORTANT NOTE** When running this script for the first time, the *previous* worksheet will not be present. This is absolutely fine. But bear in mind that the next time we want to run the script (to see the new/latest entries for the day), we should rename *FilteredData* to *previous* before we run our script.

### First Time Setup To Run Scripts in MS Excel:
Enable the Developer Tab: If not already enabled, enable the Developer tab in Excel.

Go to File>Options>Customize Ribbon.
Check Developer in the right-hand list.

### How To Add A New Script:
1. Alt + F11 to open the VBA editor.
2. Insert a new module (Insert > Module).
3. Paste the code/script into the module.
4. Close the VBA editor and return to Excel.

### Send Scripts Via Email
.vb scripts cannot be attached to an email because the filename flags up as potentially dangerous.
To get around this, we can rename each file extention to .txt
EXAMPLE: rename calc-sum-outers.vb to calc-sum-outers.txt
Then we can attach the file to an email
