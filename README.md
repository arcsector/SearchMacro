# SearchMacro
Uses Visual Basic to search through an excel spreadsheet

# How to use
In order to use this script as intended, one must first create a Searchbook in Excel.
### Preliminary Steps
1.	To create a searchbook in Excel, simply open up a new Excel document and click on the tab on the far right that says "View"
2.	Click on the "Macros" icon on the far right hand side.
3.	Enter the name "Search"; it is important that you type search in the same case as in the quotes, otherwise the function will not run properly (do not include the quotes).
4.	Click the "Create" button or press return/enter
5.	Delete anything in the window (sub search/end sub)
6.	Open the “Search Macro.vb” file in notepad, this will ensure that the text is not tampered with. 
7.	Copy + Paste the entire “Search Macro.vb” file into the window which popped up in step 4 and 5
8.	Save the pasted data (click the save button or press CTRL+S)
9.	Exit out of the window that popped up (click the “X” in the top right-hand corner)
### Main Steps
1. Copy all the data over to the Searchbook.
2. After this, a new line must be inserted at the top of the CSV. In this line, input the title of the column with the the data you want to search. You should also 
   * If you do not have a second sheet in your searchbook already, create a second sheet in it by clicking the “+” icon at the bottom of the window.
3. Next, a list of PRODID's to be compared should be pasted in a column list in only the first column of the second sheet.
4. After this, the macro can be run safely, and will copy the entire row of data from the first sheet that match the second sheet into a newly created third sheet.
   * In order to run the macro, click on "Macros" again under the "View" tab
   * Click on your "Search" Macro that we defined earlier, and click "Run"
   * If you are presented with any security prompts, allow the macro access
   * Wait. Depending on the volume of data, there may be a few seconds of run time.
