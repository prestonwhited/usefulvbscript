# usefulvbscript

VBS files that I have developed for work that do useful things.

------------------------------
SCRIPT Update File XLS to XLSX
------------------------------

SUMMARY
Changes the .xls files in a folder to .xlsx files.

CHALLENGE
Many companies have folders full of old format .xls files. When you need to use advanced Excel features, or copy and paste between new and old Excel files there are issues.

INSTRUCTIONS
1) Put this VBS file into a folder that has .xls files, run the script and click OK.
2) All .xls files will be opened one at a time, saved as a .xlsx file, then the .xls file is moved to a LegacyArchive folder.
3) At the end a log text file is created to show all changes made.

-------------------
SCRIPT Printer List
-------------------

SUMMARY
Displays a list of all printers available on a computer.

CHALLENGE
I was creating an HTA that has buttons to run external vbscripts to automate printing of various reports. Some of those reports print in black and white and others must print in color. I created a dropdown list for the user to select their color printer since each computer on the network listed our Xerox printer as something slightly different. I created this script to pull printer information from the user's computer as a test before implementing similar code inside the HTA to create a select list.

INSTRUCTIONS
1) Put this VBS file anywhere on the computer, run the script and it will show a list of all available printers.
