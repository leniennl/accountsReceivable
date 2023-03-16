Attribute VB_Name = "Module1"


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'show quick analysis tool when a range of cells are selected
    Application.QuickAnalysis.Show (xlTotals)
    
End Sub

Function GetFilteredRangeBottomRow() As Long
  Dim HeaderRow As Long, LastFilterRow As Long, Addresses() As String
  On Error GoTo NoFilterOnSheet
  With ActiveSheet
    HeaderRow = .AutoFilter.Range(1).Row
    LastFilterRow = .Range(Split(.AutoFilter.Range.Address, ":")(1)).Row
    Addresses = Split(.Range((HeaderRow + 1) & ":" & LastFilterRow). _
                      SpecialCells(xlCellTypeVisible).Address, "$")
    GetFilteredRangeBottomRow = Addresses(UBound(Addresses))
  End With
NoFilterOnSheet:
End Function

Function GetFilteredRangeTopRow() As Long
  Dim HeaderRow As Long, LastFilterRow As Long
  On Error GoTo NoFilterOnSheet
  With ActiveSheet
    HeaderRow = .AutoFilter.Range(1).Row
    LastFilterRow = .Range(Split(.AutoFilter.Range.Address, ":")(1)).Row
    GetFilteredRangeTopRow = .Range(.Rows(HeaderRow + 1), .Rows(Rows.Count)). _
                                    SpecialCells(xlCellTypeVisible)(1).Row
    If GetFilteredRangeTopRow = LastFilterRow + 1 Then GetFilteredRangeTopRow = 0
  End With
NoFilterOnSheet:
End Function
 
Function CompatibilityCheck() As Boolean
'check Excel version and compatibility
    Dim blMode As Boolean
    Dim arrVersions()
    arrVersions = Array("12.0", "14,0", "15.0", "16.0")
    If Application.IsNumber(Application.Match(Application.Version, _
       arrVersions, 0)) Then
        blMode = ActiveWorkbook.Excel8CompatibilityMode
        If blMode = True Then
            CompatibilityCheck = True
        ElseIf blMode = False Then
            CompatibilityCheck = False
        End If
    End If
End Function

Sub CheckCompatibility()
'accompany UDF CompantibilityCheck()
    Dim xlCompatible As Boolean
    xlCompatible = CompatibilityCheck
    If xlCompatible = True Then
        MsgBox "You are attempting to use an Excel 2007 or newer function "
    End If
End Sub

Sub UniqueItems()
'find unique items in column A
Selection.AutoFilter
Range("a:a").AdvancedFilter Action:=xlFilterInPlace, Unique:=True
End Sub

Function FileThere(strFileName As String) As Boolean
'to determine if a Workbook exists
FileThere = Dir(strFileName) <> ""
End Function

Sub CheckIfThere()
Dim bIsThere As Boolean
'works with Fuction FileThere to determine if a workbook exists
'Add full path to file below
bIsThere = FileThere("c:\target.xlsx")
MsgBox (bIsThere)
End Sub
Function PathThere(strPathName As String) As Boolean
PathThere = Dir(strPathName, vbDirectory) <> ""
End Function
Sub CheckIfPathThere()
'works with Fuction FileThere to determine if a folder exists
'Put directory path in PathThere function call
bIsThere = PathThere("C:\Users\mli\Dropbox\Side Hussle\VBA\01. Basic\Excel VBA Managing Files and Data\Chapter02\")
MsgBox (bIsThere)
End Sub
Function CheckIfOpen(strFileName As String) As Boolean
'check if a workbook is open
Dim w As Workbook
On Error Resume Next
Set w = Workbooks(strFileName)
If Err Then CheckIfOpen = False Else CheckIfOpen = True
On Error GoTo 0
End Function
Sub CheckIfWkbkOpen()
'Add name of file you want to check
bIsOpen = CheckIfOpen("Target.xlsx")
MsgBox (bIsOpen)
End Sub
Sub CloseCurrentWorkbook()
'close current workbook
Application.DisplayAlerts = False
'ActiveWorkbook.Save
ActiveWindow.Close
Application.DisplayAlerts = True
End Sub

Sub CloseAnotherWorkbook()
'close a workbook not currently working on
On Error Resume Next
Workbooks("target.xlsx").Close
On Error GoTo 0
End Sub

Sub SaveasCSV()
'save as CSV
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs Filename:="c:\users\CSVout", FileFormat:=xlCSV
Application.DisplayAlerts = True
End Sub

Function wksExists(strSheetName) As Boolean
On Error Resume Next
wksExists = Sheets(strSheetName).Name <> ""
On Error GoTo 0
End Function
Sub CheckForSheet()
'use function wksExists to check if a worksheet exists
Dim bIsThere As Boolean
bIsThere = wksExists("April")
MsgBox (bIsThere)
End Sub

Sub Creat_Rename_worksheets()
'add new worksheet
Sheets.Add
Sheets.Add before:=Sheets("sheet1")
Sheets.Add after:=Sheets(Sheets.Count)
'rename existing worksheet
Sheets(1).Name = "april"
Sheets("april").Name = "2021April"
End Sub
Sub copy_paste_wks_winthin_workbook()
Sheets("january").Select
Sheets("january").Copy before:=Sheets(1) 'after:=sheets(sheets.count)
End Sub
Sub copy_worksheets_to_a_new_workbook()
Sheets("january").Select
Sheets("january").Copy 'copied to a new workbook
ActiveWorkbook.SaveAs Filename:="destination.xlsx" 'save the newworkbook in the same path.
End Sub
Sub CopyToExistingWorkBook()
Dim strCurrentWkbkName As String
strCurrentWkbkName = ActiveWorkbook.Name
If CheckIfOpen("C:\Users\mli\Dropbox\Side Hussle\VBA\01. Basic\Excel VBA Managing Files and Data\Chapter03\TargetForCopy.xlsx") = False Then
  Workbooks.Open Filename:="C:\Users\mli\Dropbox\Side Hussle\VBA\01. Basic\Excel VBA Managing Files and Data\Chapter03\TargetForCopy.xlsx"
  Workbooks(strCurrentWkbkName).Activate
End If
'Sheets(1).Select
Sheets("January").Copy before:=Workbooks("TargetForCopy.xlsx").Sheets(1)
End Sub
Sub GetFullPath()
'identify a file and get full path
Dim varFName As Variant
'Code to get the file's full path goes here
varFName = Application.GetOpenFilename
MsgBox ("The file's full path is " & varFName)
End Sub
Function CreateDateString() As String

'convert date into yyyymmdd format

Dim sYear As String, sMonth As String, sDay As String

sYear = CStr(Year(Date))

If Month(Date) < 10 Then

  sMonth = "0" & CStr(Month(Date))
  
Else: sMonth = CStr(Month(Date))

End If

If Day(Date) < 10 Then

  sDay = "0" & CStr(Day(Date))
  
Else: sDay = CStr(Day(Date))

End If

CreateDateString = sYear & sMonth & sDay

End Function

Sub CloseAllWorkbooks()
'Step 1: Declare your variables
Dim wb As Workbook
'Step 2: Loop through workbooks, save and close
For Each wb In Workbooks
wb.Close SaveChanges:=True
Next wb
End Sub

Sub SelectandFormatAllNamedRanges()
'Step 1: Declare your variables.
Dim RangeName As Name
Dim HighlightRange As Range
'Step 2: Tell Excel to Continue if Error.
On Error Resume Next
'Step 3: Loop through each Named Range.
For Each RangeName In ActiveWorkbook.Names
'Step 4: Capture the RefersToRange
Set HighlightRange = RangeName.RefersToRange
'Step 5: Color the Range
HighlightRange.Interior.ColorIndex = 36
'Step 6: Loop back around to get the next range
Next RangeName
End Sub

Sub DeletingBlankRows()
'Step1: Declare your variables.
Dim MyRange As Range
Dim iCounter As Long
'Step 2: Define the target Range.
Set MyRange = ActiveSheet.UsedRange
'Step 3: Start reverse looping through the range.
For iCounter = MyRange.Rows.Count To 1 Step -1
    'Step 4: If entire row is empty then delete it.
    If Application.CountA(Rows(iCounter).EntireRow) = 0 Then
    Rows(iCounter).Delete
    End If
'Step 5: Increment the counter down
Next iCounter
End Sub

Sub UnhideAllRowsandColumns()
Columns.EntireColumn.Hidden = False
Rows.EntireRow.Hidden = False
End Sub

Sub Format_All_Formulas_ina_Workbook()
'Step 1: Declare your Variables
Dim ws As Worksheet
'Step 2: Avoid Error if no formulas are found
On Error Resume Next
'Step 3: Start looping through worksheets
For Each ws In ActiveWorkbook.Worksheets
    'Step 4: Select cells and highlight them
    With ws.Cells.SpecialCells(xlCellTypeFormulas)
    .Interior.ColorIndex = 36
    End With
'Step 5: Get next worksheet
Next ws
End Sub


Sub Read_from_Static() 'in module 1
'place a variable both static and public, using 2 variables
Static stvar As String
LastDate = stvar
'toInput
stvar = LastDate
End Sub

Sub TransposeDataSetFromMatrixToTabular()

'Step 1: Declare your Variables
    Dim SourceRange As Range
    Dim GrandRowRange As Range
    Dim GrandColumnRange As Range


'Step 2:  Define your data source range
    Set SourceRange = Sheets("Sheet1").Range("A4:M87")
 

'Step 3: Build Multiple Consolidation Range Pivot Table
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlConsolidation, _
    SourceData:=SourceRange.Address(ReferenceStyle:=xlR1C1), _
    Version:=xlPivotTableVersion14).CreatePivotTable _
    TableDestination:="", _
    TableName:="Pvt2", _
    DefaultVersion:=xlPivotTableVersion14
    
    
'Step 4: Find the Column and Row Grand Totals
    ActiveSheet.PivotTables(1).PivotSelect "'Row Grand Total'"
    Set GrandRowRange = Range(Selection.Address)
    
    ActiveSheet.PivotTables(1).PivotSelect "'Column Grand Total'", xlDataAndLabel, True
    Set GrandColumnRange = Range(Selection.Address)


'Step 5:  Drill into the intersection of Row and Column
    Intersect(GrandRowRange, GrandColumnRange).ShowDetail = True

End Sub

Sub LabelFirstandLastChartPoints()

'Step 1: Declare your variables
Dim oChart As Chart
Dim MySeries As Series

'Step 2: Point to the active chart
On Error Resume Next
Set oChart = ActiveChart

'Step 3: Exit no chart has been selected
If oChart Is Nothing Then
    MsgBox "You select a chart first."
    Exit Sub
End If

'Step 4: Loop through the chart series
For Each MySeries In oChart.SeriesCollection

'Step 5: Clear ExistingData Labels
MySeries.ApplyDataLabels (xlDataLabelsShowNone)

'Step 6: Add labels to the first and last data point
MySeries.Points(1).ApplyDataLabels
MySeries.Points(MySeries.Points.Count).ApplyDataLabels
MySeries.DataLabels.Font.Bold = True

'Step 7: Move to the next series
Next MySeries

End Sub

Sub ColorChartSeriesToMatchSourceCellColor()

'Step 1:  Declare your variables
    Dim oChart As Chart
    Dim MySeries As Series
    Dim FormulaSplit As String
    Dim SourceRangeColor As Long


'Step 2: Point to the active chart
    On Error Resume Next
    Set oChart = ActiveChart


'Step 3:  Exit no chart has been selected
    If oChart Is Nothing Then
    MsgBox "You must select a chart first."
    Exit Sub
    End If


'Step 4: Loop through the chart series
    For Each MySeries In oChart.SeriesCollection

        
'Step 5: Get Source Data Range for the target series
    FormulaSplit = Split(MySeries.Formula, ",")(2)
        
        
'Step 6: Capture the color in the first cell
    SourceRangeColor = _
    Range(FormulaSplit).Item(1).Interior.Color


'Step 7: Apply Coloring
        On Error Resume Next
        MySeries.Format.Line.ForeColor.RGB = SourceRangeColor
        MySeries.Format.Line.BackColor.RGB = SourceRangeColor
        MySeries.Format.Fill.ForeColor.RGB = SourceRangeColor
   
        If Not MySeries.MarkerStyle = xlMarkerStyleNone Then
            MySeries.MarkerBackgroundColor = SourceRangeColor
            MySeries.MarkerForegroundColor = SourceRangeColor
        End If
   
'Step 8:  Move to the next series
    Next MySeries
    
End Sub

Sub MailActiveWBAsAnAttachment()

'set a reference to Microsoft Outlook XX Object Library
'Step 1:  Declare your variables
    Dim OLApp As Outlook.Application
    Dim OLMail As Object
    
    
'Step 2:  Open Outlook start a new mail item
    Set OLApp = New Outlook.Application
    Set OLMail = OLApp.CreateItem(0)
    OLApp.Session.Logon
    
    
'Step 3:  Build your mail item and send
    With OLMail
    .To = "admin@datapigtechnologies.com; mike@datapigtechnologies.com"
    .CC = ""
    .BCC = ""
    .Subject = "This is the Subject line"
    .Body = "Hi there"
    .Attachments.Add ActiveWorkbook.FullName
    .Display  'Change to .Send to send without reviewing
    End With
    
    
'Step 4:  Memory cleanup
    Set OLMail = Nothing
    Set OLApp = Nothing

End Sub
