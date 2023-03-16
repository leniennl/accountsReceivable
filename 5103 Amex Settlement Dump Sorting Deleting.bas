Attribute VB_Name = "Module11"
Sub deleteAmexLines()

Dim FinalRow As Double

Dim i As Integer

Application.ScreenUpdating = False

Application.Calculation = xlCalculationManual

FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = FinalRow To 2 Step -1

    If Left(Cells(i, 4), 3) <> "979" Or "803" Then
    
        Cells(i, 4).EntireRow.Delete
        
    End If
    
Next i

MsgBox "completed"

Application.ScreenUpdating = True

Application.Calculation = xlAutomatic


End Sub


