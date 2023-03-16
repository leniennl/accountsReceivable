Attribute VB_Name = "Module1"
Sub Sort_CCA_Monthly_Statement()

'use the following website to convert CCA statemtn from PDF to xlsx
'https://www.adobe.com/au/acrobat/online/pdf-to-excel.html

'CHANGE DATE TO CURRENT MONTH OF STATEMENT

Dim sh As Shape
Dim finalrow As Long
Dim i As Integer


finalrow = Cells(Rows.Count, 1).End(xlUp).Row


'make a copy of itself before doing anyting
ActiveSheet.Copy before:=Sheets(1)
Sheets(1).Activate


'delete first 5 lines but retain headings
Rows("1:5").EntireRow.Delete


'unmerge cells
Cells.UnMerge
Range("a1:z1").EntireColumn.AutoFit
Range("a1:a" & finalrow).EntireRow.AutoFit


'delete shapes
For Each sh In Sheets(1).Shapes
    sh.Delete
Next sh

'delete empty rows, delete not-invoice-related rows
For i = finalrow To 2 Step -1
    If CStr(Left(Cells(i, 1), 6)) = "Outlet" Or CStr(Left(Cells(i, 1), 6)) = "" Or IsNumeric(Cells(i, 1)) = False Or Cells(i, 1).Text = "" Then
        Cells(i, 1).EntireRow.Delete
    End If
Next i
finalrow = Cells(Rows.Count, 1).End(xlUp).Row


'delete statement's subtotal lines
For i = finalrow To 2 Step -1
    If IsDate(Cells(i, 6)) = False And Cells(i, 6).Value <> "" Then
        Cells(i, 1).EntireRow.Delete
    End If
Next i
finalrow = Cells(Rows.Count, 1).End(xlUp).Row


'find current month invoices
For i = finalrow To 2 Step -1
    If Cells(i, 6).Value < #1/2/2021# And Cells(i, 6).Value <> "" Then
        Cells(i, 6).EntireRow.Delete
    End If
Next i
finalrow = Cells(Rows.Count, 1).End(xlUp).Row


Range("a1:a" & finalrow).EntireColumn.AutoFit
MsgBox "Current month invoices selected" & vbNewLine & "Continue work on the statement"

End Sub


