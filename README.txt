Summary:

What this script does
1: opens the excel file that it shares a folder with
2: starting at the 2nd row of the 1st column, compares each location with a list of active locations within HxGN
3: once each row has been checked, a report of all wrong locations will be printed to the console screen and written to a text file named after the current date and time (MM.DD.YYYY-HHMM)
If there are no wrong locations, the console window will say the same and no file will be created 

Requirements:

There MUST be ONE excel file (.xls or .xlsx) and the .txt file named "locationZone" in the same folder as the script file "locationCheckZone.py".
If any of these are not true, the script will not work.