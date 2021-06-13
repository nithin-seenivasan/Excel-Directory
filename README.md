## Introduction
This repo contins the VBA code used to make a simple Excel directory. The step-by-step instructions to actually make the directory are described my article on Medium titled ["Code a directory for a large Excel file in 6 easy steps!"](https://nithin-seenivasan.medium.com/how-to-code-a-directory-for-a-large-excel-file-in-6-easy-steps-a7ae42a19517 "Medium Article").  

## Section 1: Declaring Variables
~~~
    'Define the workbook and worksheet
    Dim mainworkBook As Workbook
    Dim objNewWorksheet As Worksheet
    
    'Set the active worksheet
    Set mainworkBook = ActiveWorkbook
    Set objNewWorksheet = mainworkBook.Sheets("Directory")
    j = 3 'To start writing from the 3nd row
~~~

The following operations are done here 
- Set the variable "mainworkBook" as the currently active WorkBook 
- Set the variable objNewWorksheet as the sheet "Directory"
- Set the vairable j as 3, in order to start from row 3 of the Directory sheet


## Section 2: Get data from each WorkSheet 
~~~
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
~~~

- Start from i=2 to skip the "Directory" sheet itself, and iterate to the number of sheets available 
- Write index value to the first column of directory (ref. line 26)
- Get the "name" of each Worksheet and write it to second column of directory (ref. line 28)
- Get the text value of cell A1 of each worksheet and write it to the third column of directory (ref. lines 30 and 31)
- Set the hyperlink of each worksheet in 4th column of directory (ref. line 33)
- Get the text value of cell M1 of each worksheet and write it to the 5th column of directory (ref. lines 35 and 36)

# Section 3: Format the columns 
~~~
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
~~~

- Add column names, set font weight to bold, and autosize the volumns 
- Select cell A2 to return focus of screen to the left most corner
