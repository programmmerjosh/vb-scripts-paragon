CREDIT: Joshua van Niekerk (and ChatGPT)

### Description:
I have leveraged ChatGPT to code up several scripts for me (for my Workflow Coordinator role at Paragon).
These scripts serve the purpose of calculating and identifying what stock to order (from the warehouse) before the print reaches the machines in production.

### Application:
Typically, it would be best to run one of the SORT scripts first: FilterSortClientCorpPacks() OR FilterSortClientPacks().
HighlightRemakes() is not essential, but useful for visual purposes.
HighlightForInserts() is great to highlight work-orders (to order) as they require inserts to be staged.
GetOuters() identifies (based on CORP_CD) which outers each work-order will need.
CalcSumOuters() will create a summary of the total sum of each outer we will need to enclose all the jobs on the active worksheet.

### Dependencies:
To perform/run any of these scripts, you need to rename the active worksheet "Special1"

GetOuters() depends on the OUTERSKEY worksheet. To use OUTERSKEY, you need to open up outers-key.xls. Copy the OUTERSKEY data and paste it in a new worksheet on the Paragon SLA excel spreadhseet and rename the worksheet "OUTERSKEY".

CalcSumOuters() depends on the data that GetOuters() will add. So, you can only run CalcSumOuters() after GetOuters() has been run.

### First Time Setup To Run Macros:
Enable the Developer Tab: If not already enabled, enable the Developer tab in Excel.

Go to File > Options > Customize Ribbon.
Check Developer in the right-hand list.

### Steps to Use These Scripts
Open the VBA editor (Alt + F11).
Insert a new module (Insert > Module).
Paste the code/script into the module.
Close the VBA editor and return to Excel.
Run the macro (Alt + F8, select \[SCRIPT_NAME], then click "Run").

### Send scripts via email
.vb scripts cannot be attached to an email because the filename flags up as potentially dangerous.
To getaround this, we can rename each file extention to .txt
EXAMPLE: rename calc-sum-outers.vb to calc-sum-outers.txt
Then we can attach the file to an email