Attribute VB_Name = "Module1"
Sub Button1_Click()
        
    'Define the workbook and worksheet
    Dim mainworkBook As Workbook
    Dim objNewWorksheet As Worksheet
    
    'Set the active worksheet
    Set mainworkBook = ActiveWorkbook
    Set objNewWorksheet = mainworkBook.Sheets("Directory")
    j = 3 'To start writing from the 3nd row
    
    'Find out all visible names and write here
    'i = sheet index, starts at 2 to ignore the Directory sheet

    For i = 2 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Visible = True Then
            objNewWorksheet.Cells(j, 1) = j - 2
            'Adds Number of the worksheet (i.e. the "Name" of the worksheet which is a numbere here=
            objNewWorksheet.Cells(j, 2) = ThisWorkbook.Sheets(i).Name
            'adds Worksheet Name (defined in cell A1 of each sheet)
            ThisWorkbook.Sheets(i).Range("A1").Copy
            objNewWorksheet.Cells(j, 3).PasteSpecial xlPasteValues
            'Adds the hyperlink
            objNewWorksheet.Hyperlinks.Add Cells(j, 4), Address:="", SubAddress:="'" & ThisWorkbook.Sheets(i).Name & "'" & "!A1", TextToDisplay:="Link"
            'adds description (Defined in cell M1 of each sheet)
            ThisWorkbook.Sheets(i).Range("M1").Copy
            objNewWorksheet.Cells(j, 5).Select
            ActiveSheet.Paste
            j = j + 1
        End If
    Next i
    
    'Set the Column Names and formats it
    With objNewWorksheet
         .Cells(2, 1) = "INDEX"
         .Cells(2, 1).Font.Bold = True
         .Cells(2, 2) = "Worksheet No."
         .Cells(2, 2).Font.Bold = True
         .Cells(2, 3) = "Worksheet Name"
         .Cells(2, 3).Font.Bold = True
         .Cells(2, 4) = "HYPERLINK"
         .Cells(2, 4).Font.Bold = True
         .Cells(2, 5) = "DESCRIPTION"
         .Cells(2, 5).Font.Bold = True
         .Columns("A:E").AutoFit
         .Range("A2").Select
    End With
    
End Sub
