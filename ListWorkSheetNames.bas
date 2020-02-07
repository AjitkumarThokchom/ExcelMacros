Attribute VB_Name = "Module1"
'How to use this utility macro
'1. Add a new worksheet with name as "TOC"
'2. Add a button in the worksheet and right click the button => assign macros => New macros
'3. In Visual Basic Editor, copy and paste this code snippet

Sub ListWorksheets()
 
Dim ws As Worksheet
Dim x As Integer

'clear the sheet
Sheets("TOC").Range("A:A").Clear
 
'write header in the first row
Sheets("TOC").Cells(1, 1) = "SL"
Sheets("TOC").Cells(1, 2) = "Worksheet"
 
'Start listing worsheets from 2nd row
x = 2
 
For Each ws In Worksheets
    'Don't list if the worksheet name is "TOC"
     If ws.Name <> "TOC" Then
        Sheets("TOC").Cells(x, 1) = x
        Sheets("TOC").Cells(x, 2) = ws.Name
         x = x + 1
     End If
     
Next ws
 
End Sub
