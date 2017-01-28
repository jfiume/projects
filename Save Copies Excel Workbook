Sub SaveCopies()
Dim wb As Workbook
Dim ws As Worksheet
Dim wsList As Worksheet
Dim pt As PivotTable
Dim strSD As String
Dim strMsg As String
Dim pvtFld As PivotField
Dim colLabel As String
Dim LastRow As Long
Dim i As Long
Dim PtRng As Range
Dim BlRng1 As Range
Dim BlRng2 As Range
Dim coll As Collection
Dim nms As Long
Dim n As Long

On Error Resume Next
Application.DisplayAlerts = False
Application.EnableEvents = False


Set wb = ThisWorkbook
'make an array full of the named ranges to loop through them to make the copies


For nms = 1 To wb.names.Count
    'Debug.Print ActiveWorkbook.names(nms).name
  For Each ws In wb.Worksheets
    For Each pt In ws.PivotTables
    
    'changes the pivot table data source
      pt.ChangePivotCache _
        wb.PivotCaches.Create(SourceType:=xlDatabase, _
              SourceData:=(wb.names(nms).name))
              
              
        'adds month to the pivot table as a column
        pt.PivotFields("Month").Orientation = xlColumnField
                
        'adds year to the pivot table as a column
        pt.PivotFields("Year").Orientation = xlColumnField
               
        'named the fields "year" and "month"
        pt.PivotFields.CompactLayoutColumnHeader = "Month"
        pt.PivotFields.CompactLayoutColumnHeader = "Year"
        
        'hides erroneous data
        With pt.PivotFields("Month")
            .PivotItems("2015").Visible = False
            .PivotItems("2016").Visible = False
            .PivotItems("Jan").Visible = True
            .PivotItems("Feb").Visible = True
            .PivotItems("Mar").Visible = True
            .PivotItems("Apr").Visible = True
            .PivotItems("May").Visible = True
            .PivotItems("Jun").Visible = True
            .PivotItems("Jul").Visible = True
            .PivotItems("Aug").Visible = True
            .PivotItems("Sep").Visible = True
            .PivotItems("Oct").Visible = True
            .PivotItems("Nov").Visible = True
            .PivotItems("Dec").Visible = True
        End With
        
        'groups the columns by month and year
        pt.PivotFields("Month").AutoSort _
            xlAscending, "Month"
           
        'defines the year field to be like "2000"
        pt.PivotFields("Year").NumberFormat = "###0"
        
        'shows the columns with no data
        pt.PivotFields("Month").ShowAllItems = True
        
        'sorts the vendors by alpha
        pt.PivotFields("Vendor").AutoSort _
            xlAscending, "Vendor"
        
        'sorts the items by alpha
         pt.PivotFields("FullItem").AutoSort _
            xlAscending, "FullItem"

    Next pt
  Next ws



'highlight Sheet5 loop
LastRow = Sheets("Sheet5").Cells(Rows.Count, "A").End(xlUp).Row
i = 10
Do Until i = LastRow + 1
    Set PtRng = Sheets("Sheet5").Range(Sheets("Sheet5").Cells(i, 11), Sheets("Sheet5").Cells(i, 17))
    Set BlRng1 = Sheets("Sheet5").Range(Sheets("Sheet5").Cells(i, 1), Sheets("Sheet5").Cells(i, 10))
    Set BlRng2 = Sheets("Sheet5").Range(Sheets("Sheet5").Cells(i, 18), Sheets("Sheet5").Cells(i, 30))
    PtRng.Interior.Color = RGB(228, 223, 236)
    BlRng1.Interior.Color = RGB(255, 255, 255)
    BlRng2.Interior.Color = RGB(255, 255, 255)
    PtRng.Font.Color = RGB(22, 54, 92)
    i = i + 1
Loop


'highlight Sheet6 loop
LastRow = Sheets("Sheet6").Cells(Rows.Count, "A").End(xlUp).Row
i = 10
Do Until i = LastRow + 1
    Set PtRng = Sheets("Sheet6").Range(Sheets("Sheet6").Cells(i, 11), Sheets("Sheet6").Cells(i, 17))
    Set BlRng1 = Sheets("Sheet6").Range(Sheets("Sheet6").Cells(i, 1), Sheets("Sheet6").Cells(i, 10))
    Set BlRng2 = Sheets("Sheet6").Range(Sheets("Sheet6").Cells(i, 18), Sheets("Sheet6").Cells(i, 30))
    PtRng.Interior.Color = RGB(228, 223, 236)
    BlRng1.Interior.Color = RGB(255, 255, 255)
    BlRng2.Interior.Color = RGB(255, 255, 255)
    PtRng.Font.Color = RGB(22, 54, 92)
    i = i + 1
Loop


'Hides Columns u to y
If Sheets("Sheet5").Columns("u:y").EntireColumn.Hidden = True Then
Else
    Sheets("Sheet5").Columns("u:y").EntireColumn.Hidden = True
End If

'Hides Columns u to y
If Sheets("Sheet6").Columns("u:y").EntireColumn.Hidden = True Then
Else
    Sheets("Sheet6").Columns("u:y").EntireColumn.Hidden = True
End If


'Formats dates to display normally
Sheets("Sheet6").Range("b7").Style = "Normal"
Sheets("Sheet6").Range("n7").Style = "Normal"
Sheets("Sheet5").Range("b7").Style = "Normal"
Sheets("Sheet5").Range("n7").Style = "Normal"


'Formats the Logo
With Sheets("Sheet5").Shapes.Item("Round Same Side Corner Rectangle 10")
    .Left = 0
    .Top = 0
    .Width = Sheets("Sheet5").Range("a1:ac1").Width
    .Height = Sheets("Sheet5").Range("a1:a5").Height
End With
'Formats the 1st bullet
With Sheets("Sheet5").Shapes.Item("Rounded Rectangle 13")
    .Left = Sheets("Sheet5").Range("j1").Left
    .Top = Sheets("Sheet5").Range("j1").Top
    .Width = Sheets("Sheet5").Range("j1:ac1").Width
    .Height = Sheets("Sheet5").Range("a1:a2").Height
End With
'Formats the 2st bullet
With Sheets("Sheet5").Shapes.Item("Rounded Rectangle 6")
    .Left = Sheets("Sheet5").Range("j1").Left
    .Top = Sheets("Sheet5").Range("j2").Top
    .Width = Sheets("Sheet5").Range("j1:ac1").Width
    .Height = Sheets("Sheet5").Range("a1:a2").Height
End With
'Formats the 3st bullet
With Sheets("Sheet5").Shapes.Item("Rounded Rectangle 8")
    .Left = Sheets("Sheet5").Range("j1").Left
    .Top = Sheets("Sheet5").Range("j3").Top
    .Width = Sheets("Sheet5").Range("j1:ac1").Width
    .Height = Sheets("Sheet5").Range("a1:a2").Height
End With
'Formats the 4st bullet
With Sheets("Sheet5").Shapes.Item("Rounded Rectangle 7")
    .Left = Sheets("Sheet5").Range("j1").Left
    .Top = Sheets("Sheet5").Range("j4").Top
    .Width = Sheets("Sheet5").Range("j1:ac1").Width
    .Height = Sheets("Sheet5").Range("a1:a2").Height
End With


'Formats the Logo
With Sheets("Sheet6").Shapes.Item("Round Same Side Corner Rectangle 1")
    .Left = 0
    .Top = 0
    .Width = Sheets("Sheet6").Range("a1:ac1").Width
    .Height = Sheets("Sheet6").Range("a1:a5").Height
End With

'Formats the 1st bullet
With Sheets("Sheet6").Shapes.Item("Rounded Rectangle 2")
    .Left = Sheets("Sheet6").Range("j1").Left
    .Top = Sheets("Sheet6").Range("j1").Top
    .Width = Sheets("Sheet6").Range("j1:ac1").Width
    .Height = Sheets("Sheet6").Range("a1:a2").Height
End With
'Formats the 2st bullet
With Sheets("Sheet6").Shapes.Item("Rounded Rectangle 3")
    .Left = Sheets("Sheet6").Range("j1").Left
    .Top = Sheets("Sheet6").Range("j2").Top
    .Width = Sheets("Sheet6").Range("j1:ac1").Width
    .Height = Sheets("Sheet6").Range("a1:a2").Height
End With
'Formats the 3st bullet
With Sheets("Sheet6").Shapes.Item("Rounded Rectangle 5")
    .Left = Sheets("Sheet6").Range("j1").Left
    .Top = Sheets("Sheet6").Range("j3").Top
    .Width = Sheets("Sheet6").Range("j1:ac1").Width
    .Height = Sheets("Sheet6").Range("a1:a2").Height
End With
'Formats the 4st bullet
With Sheets("Sheet6").Shapes.Item("Rounded Rectangle 4")
    .Left = Sheets("Sheet6").Range("j1").Left
    .Top = Sheets("Sheet6").Range("j4").Top
    .Width = Sheets("Sheet6").Range("j1:ac1").Width
    .Height = Sheets("Sheet6").Range("a1:a2").Height
End With

CopyWorkBooks wb
Next

exit_Handler:
'  wsList.Delete
  Application.EnableEvents = True
  Application.DisplayAlerts = True
  Exit Sub


err_Handler:
  MsgBox "Could not update pivot table source data"
  Resume exit_Handler





End Sub
Function CopyWorkBooks()
'Copy every unhidden sheet from the workbook
Dim ws As Worksheet
Dim FileExtStr As String
Dim FileFormatNum As Long
Dim Masterwb As Workbook
Dim Destwb As Workbook
Dim DateString As String
Dim FileName As String


Application.ScreenUpdating = False

'sets the master workbook
Set Masterwb = ThisWorkbook

'Create new folder to save the new files in
DateString = Format(Now, "yyyymmdd")
FileName = Replace(Range("a9"), Left(Range("a9"), 10), "_")
FolderName = "FolderName"



Workbooks.Add
'Set Destwb to the new workbook
Set Destwb = ActiveWorkbook

On Error Resume Next

For Each ws In Masterwb.Worksheets
'If the sheet is visible then copy it to a new workbook
    If ws.Visible = -1 Then
        ws.Copy before:=Workbooks(Destwb.name).Sheets("Sheet1")
    End If
Next ws

'copy paste values
With Destwb.Sheets("Sheet5").UsedRange
    .Cells.Copy
    .Cells.PasteSpecial xlPasteValues
    .Cells(1).Select
End With

Sheets("Sheet5").Range("a6").Clear
Set ColorRng1 = Sheets("Sheet5").Range("a6:t8")
Set ColorRng2 = Sheets("Sheet5").Range("a9:t9")
ColorRng1.Interior.Color = RGB(247, 150, 70)
ColorRng1.Font.Color = RGB(255, 255, 255)
ColorRng2.Interior.Color = RGB(253, 233, 217)
ColorRng1.Font.Bold = True
ColorRng2.Font.Bold = True

'copy paste values
With Destwb.Sheets("Sheet6").UsedRange
    .Cells.Copy
    .Cells.PasteSpecial xlPasteValues
    .Cells(1).Select
End With

Sheets("Sheet6").Range("a6").Clear
Set ColorRng1 = Sheets("Sheet6").Range("a6:t8")
Set ColorRng2 = Sheets("Sheet6").Range("a9:t9")
ColorRng1.Interior.Color = RGB(247, 150, 70)
ColorRng1.Font.Color = RGB(255, 255, 255)
ColorRng2.Interior.Color = RGB(253, 233, 217)
ColorRng1.Font.Bold = True
ColorRng2.Font.Bold = True


'Delete unnecessary sheets
Application.DisplayAlerts = False
Destwb.Sheets("sheet1").Delete
Destwb.Sheets("sheet2").Delete
Destwb.Sheets("sheet3").Delete
Application.DisplayAlerts = True


'set workbook to binary
With Destwb
  FileExtStr = ".xlsb": FileFormatNum = 50
End With


'Save the new workbook and close it
With Destwb
    .SaveAs FolderName & _
    DateString & FileName & FileExtStr, _
            FileFormat:=FileFormatNum
    .Close False
End With

Application.ScreenUpdating = True


End Function
