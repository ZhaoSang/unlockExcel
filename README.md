# unlockExcel
Unlock Excel programmatically using Python win32com
The program is leveraging easygui for user input password and win32com to dispatch Excel program to unlock workbook protection.
The program will loop through the current program directory for files ending with "xlsx" and "xlsb", open each if condition met and save as a new file with "_unlocked" added at the end.
Note: Printing is beacuse windows console will be leveraged upon wrapping the .py into .exe
