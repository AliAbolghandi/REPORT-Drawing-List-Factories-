Attribute VB_Name = "DataUpdate"
Option Explicit

Sub Copy_Drawing_List()

    Dim wbSource As Workbook
    Dim wbTarget As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim sourcePath As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rngSource As Range
    Dim rngTarget As Range
    Dim tbl As ListObject
    Dim result As VbMsgBoxResult
    
    result = MsgBox("Do you want to continue?", vbOKCancel)

    If result = vbCancel Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    On Error Resume Next
    Sheets(Sheet1.Name).Select
    ActiveSheet.ShowAllData
    Sheets(Sheet2.Name).Select
    ActiveSheet.ShowAllData
    ActiveWorkbook.RefreshAll

    Set wbTarget = ThisWorkbook

    sourcePath = "Y:\Drawing Factories\Drawing List Factories-NEW.xlsm"

    Set wbSource = Workbooks.Open(sourcePath, ReadOnly:=True)

    Set wsSource = wbSource.Sheets(1)
    Set wsTarget = wbTarget.Sheets(1)

    lastRow = wsSource.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = wsSource.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    Set rngSource = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastCol))

    'Find existing table
    On Error Resume Next
    Set tbl = wsTarget.ListObjects("tblDrawing")
    On Error GoTo 0

    If Not tbl Is Nothing Then
        
        'Clear only data area, not table structure
        If Not tbl.DataBodyRange Is Nothing Then
            tbl.DataBodyRange.Clear
        End If
        
        Set rngTarget = tbl.HeaderRowRange.Cells(1, 1).Resize(lastRow, lastCol)

    Else
        
        Set rngTarget = wsTarget.Cells(1, 1).Resize(lastRow, lastCol)
        rngTarget.Clear

    End If

    rngSource.Copy

    rngTarget.PasteSpecial Paste:=xlPasteAll

    Application.CutCopyMode = False

    wbSource.Close SaveChanges:=False

    
    
    
    On Error Resume Next
    Sheets(Sheet1.Name).Select
    Sheets(Sheet1.Name).Range("A3:AJ3").Select
    Selection.Copy
    Range("tblDrawing[#Headers]").Select
    ActiveSheet.Paste
    Rows("3:3").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    
    Sheets(Sheet2.Name).Select
    MsgBox "Copy completed successfully."

End Sub

