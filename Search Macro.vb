'==================================================================================================================================================================='
'                                                                                                                                                                   '
'           Author:                 George Haraksin                                                                                                                 '
'           Corporation:            Pacific Advisors                                                                                                                '
'           Last Revision Date:     4/5/17                                                                                                                      '
'           Supervisor:             Chris McMahan                                                                                                                   '
'           Description:                                                                                                                                            '
'This visual basic script is designed to be used by the accounting team to match PRODID's in order to get                                                           '
'the names and values that the team needs.                                                                                                                          '
'                                                                                                                                                                   '
'=============================================================================INSTRUCTIONS=========================================================================='
'                                                                                                                                                                   '
'   1. In order to use this script as intended, one must define it as a macro in Excel.
'       a. To define this script as a macro in Excel, simply open up a new workbook and click on the sheet on the far right tab that says "View"
'       b. Click on the "Macros" icon in the far right hand side.
'       c. Enter the name "Search"; IT IS IMPORTANT THAT YOU TYPE SEARCH IN THE SAME CASE AS IN THE QUOTES, OTHERWISE THE FUNCTION WILL NOT RUN PROPERLY
'       d. Click the "Create" button or press return/enter
'		e. Delete anything automatically generated in the window
'       f. Copy + Paste this ENTIRE file into the Window which pops up
'       g. Save the pasted data (click the save button or press CTRL+S)
'       h. Exit out of the Developer window that popped up
'   2. Then, one must open the CSV given to the team by Guardian on the first sheet.                                                                                '
'   3. After this, a new line must be inserted at the top of the CSV                                                                                                '
'   4. In this line, the column with the PRODID's should be typed in ALL CAPS, not with the quotes: "PRODID"                                          '
'   5. Next, a list of PRODID's should be pasted in a column list in the first column of the second sheet.                                                          '
'   6. After this, the macro can be run safely, and will copy the entire row of PRODID's of the first sheet that match                                              '
'      the second sheet into a newly created third sheet.                                                                                                           '
'       a. In order to run the macro, click on "Macros" again under the "View" tab
'       b. Click on your "Search" Macro that we defined earlier, and click "Run"
'       c. If you are presented with any security prompts, allow the macro access
'       d. Wait. Depending on the volume of data, there may be a few seconds of run time.
'   7. You may want to save this workbook as a Macro-enabled workbook for reuse, in order to mitigate steps 1.a through 1.g
'       a. To save this workbook as a Macro-enabled workbook, click on the "File" tab in the top left-hand corner in Excel
'       b. Click "Save As" in the window that opens up
'       c. From the drop-down menu select "Excel Macro-enabled Workbook (*.xlsm)" and click "Save"
'       d. Now your workbook will be able to be reused
'                                                                                                                                                                   '
'==================================================================================================================================================================='

Public Function IsAlpha(strValue As String) As Boolean
    IsAlpha = strValue Like WorksheetFunction.Rept("[a-zA-Z'.()/ \-]*", Len(strValue))
End Function

Public Function IsPRODID(strValue As String) As Boolean
    IsPRODID = strValue Like WorksheetFunction.Rept("[a-zA-Z0-9 ]*", Len(strValue))
End Function

Sub Search()
    
    Dim rngHeaders As Range
    Dim rngHdrFound As Range
    
    'Find PROD ID location in first sheet of CSV
    Set rngHeaders = Intersect(Sheets(1).UsedRange, Sheets(1).Rows(1))
    Set rngHdrFound = rngHeaders.Find("PRODID")
    If rngHdrFound Is Nothing Then
        MsgBox ("Cannot find Header ""PRODID"". Please name the Header this title.")
        Exit Sub
    End If
    
    'Set range as column
    Set rngHdrFound = rngHdrFound.EntireColumn.Cells
    
    'Get PROD IDs Pasted from second sheet as a range
    Dim rng As Range, cellAdd As Range, cellSearch As Range, prodID As Range
    Set rng = Sheets(2).UsedRange
    If rng Is Nothing Then
        MsgBox ("There is nothing in Sheet 2. Please put your items to compare in Sheet 2.")
        Exit Sub
    End If

    'Add *000 to PRODID's in second sheet
    For Each cellAdd In rng
        'Only add *000 if the cell contains numbers
        If IsPRODID(cellAdd.Value) = True Then
            'Only add *000 if the cell doesn't have it yet
            If cellAdd.Value <> "^[\*]" Then
                'Only add *000 if the cell isn't "Agent Code"
                If cellAdd.Value <> "Agent Code" Then
                    'Only Add *000 if the cell isn't "Group Debt"
                    If cellAdd.Value <> "GROUP DEBT" Then
                        'Only add *000 if the cell isn't empty
                        If IsEmpty(cellAdd) = False Then
                            cellAdd.Value = "*000" & cellAdd.Value
                        End If
                    End If
                End If
            End If
        End If
    Next cellAdd
    
    'Define new Worksheet to paste results
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    ws.Select

    'Begin Search algorithm and pasting
    For Each cellSearch In rng
        'If the cell value doesn't contain numbers, then paste it to the worksheet
        If IsAlpha(cellSearch.Value) = True Then
            'If the new worksheet, check if the cells are used, and go until it finds an empty cell
            For Each cell In ws.Columns(1).Cells
                If IsEmpty(cell) = True Then
                    cellSearch.EntireRow.Copy cell
                    Exit For
                End If
            Next cell
        End If
        'Cycle through all PRODID's in CSV
        For Each prodID In rngHdrFound
            'If the cell value is equal to the PROD ID value then paste the entire row to the new worksheet
            If (prodID.Value = cellSearch.Value) Then
                'If the new worksheet, check if the cells are used, and go until it finds an empty cell
                For Each cell In ws.Columns(1).Cells
                    If IsEmpty(cell) = True Then
                        'copy+paste
                        prodID.EntireRow.Copy cell
                        Exit For
                    End If
                Next cell
            End If
            'If the list ends, stop searching
            If IsEmpty(prodID) Then
                Exit For
            End If
        Next prodID
    Next cellSearch
         
    'Close the Function
End Sub
