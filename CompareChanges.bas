Attribute VB_Name = "CompareChanges"
'==================================================================================
' MODULE: CompareChanges
' PURPOSE: Compare local Sheet(1) data with network file Sheet(1) data.
'          Produces a read-only-safe report on a "Changes" sheet.
'          Does NOT modify the main sheet or the network file in any way.
'
' HOW TO USE:
'   1. Import this module into your workbook (VBA Editor > File > Import File)
'   2. Adjust the constants below (NETWORK_FILE_PATH, KEY columns) if needed
'   3. Run Sub CompareChanges_Run
'==================================================================================
Option Explicit

' ---------------------------------------------------------------------------------
' CONFIGURATION
' ---------------------------------------------------------------------------------
Private Const NETWORK_FILE_PATH As String = "Y:\Drawing Factories\Drawing List Factories-NEW.xlsm"
Private Const MAIN_SHEET_INDEX As Long = 1
Private Const NETWORK_SHEET_INDEX As Long = 1
Private Const CHANGES_SHEET_NAME As String = "Changes"

' Row where the actual column headers live (REQUEST, NUMBER, ...).
' In this workbook row 1 is not the header row (it holds stray values),
' the real header row is row 2. Change this if the layout changes.
Private Const HEADER_ROW As Long = 2

Private Const KEY_COL_NAME_1 As String = "REQUEST"
Private Const KEY_COL_NAME_2 As String = "NUMBER"

' Extra metadata columns appended after the data columns in the Changes report.
' Values come from the network file's document properties (who last saved it)
' and the file's last-modified timestamp (when it was last saved).
Private Const EXTRA_COL_COUNT As Long = 3
Private Const HEADER_CHANGED_BY As String = "Changed By"
Private Const HEADER_TIME As String = "Time"
Private Const HEADER_DATE As String = "Date"

' Colors (RGB)
Private Const COLOR_MODIFIED As Long = 65535    ' Yellow  (RGB 255,255,0)
Private Const COLOR_ADDED As Long = 5296274      ' Green   (RGB 178,255,80 -> using a pleasant green below instead)
Private Const COLOR_DELETED As Long = 255        ' Red     (RGB 255,0,0)

' ---------------------------------------------------------------------------------
' MAIN ENTRY POINT
' ---------------------------------------------------------------------------------
Sub CompareChanges_Run()

    Dim wbMain As Workbook
    Dim wbNetwork As Workbook
    Dim wsMain As Worksheet
    Dim wsNetwork As Worksheet
    Dim wsChanges As Worksheet

    Dim arrOld As Variant
    Dim arrNew As Variant

    Dim dictOld As Object
    Dim dictNew As Object

    Dim headers() As Variant
    Dim colCount As Long
    Dim keyColRequest As Long
    Dim keyColNumber As Long

    Dim countModified As Long
    Dim countAdded As Long
    Dim countDeleted As Long

    Dim changeAuthor As String
    Dim changeDateTime As Date

    Dim originalScreenUpdating As Boolean
    Dim originalEnableEvents As Boolean
    Dim originalCalculation As XlCalculation
    Dim originalDisplayAlerts As Boolean

    Dim networkOpenedHere As Boolean
    networkOpenedHere = False

    Dim CurrentStep As String
    CurrentStep = "Initializing"

    On Error GoTo ErrorHandler

    ' --- Save and set application state -----------------------------------------
    originalScreenUpdating = Application.ScreenUpdating
    originalEnableEvents = Application.EnableEvents
    originalCalculation = Application.Calculation
    originalDisplayAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    CurrentStep = "Getting main workbook and sheet"
    Set wbMain = ThisWorkbook
    Set wsMain = wbMain.Sheets(MAIN_SHEET_INDEX)

    ' --- Validate and check network file exists -------------------------------------
    CurrentStep = "Validating network path"
    Dim cleanPath As String
    cleanPath = Trim(NETWORK_FILE_PATH)

    ' Diagnostic: reveal any hidden/non-standard characters in the path
    If cleanPath <> NETWORK_FILE_PATH Then
        MsgBox "Warning: the network path constant has leading/trailing spaces." & vbCrLf & _
               "Original length: " & Len(NETWORK_FILE_PATH) & vbCrLf & _
               "Trimmed length: " & Len(cleanPath), vbExclamation, "Path Warning"
    End If

    CurrentStep = "Checking network file exists (Dir)"
    If Dir(cleanPath) = "" Then
        MsgBox "Network file not found or not accessible:" & vbCrLf & cleanPath & vbCrLf & vbCrLf & _
               "Check that:" & vbCrLf & _
               "- The network share is reachable (try opening this exact path in File Explorer)" & vbCrLf & _
               "- You have at least Read permission on the folder" & vbCrLf & _
               "- The path has no typo or hidden character", vbCritical, "File Not Found"
        GoTo CleanExit
    End If

    ' --- Open network file as ReadOnly ---------------------------------------------
    CurrentStep = "Opening network file (ReadOnly)"
    Set wbNetwork = Workbooks.Open(fileName:=cleanPath, _
                                    UpdateLinks:=0, _
                                    ReadOnly:=True, _
                                    IgnoreReadOnlyRecommended:=True, _
                                    Notify:=False)
    networkOpenedHere = True

    CurrentStep = "Getting network sheet"
    Set wsNetwork = wbNetwork.Sheets(NETWORK_SHEET_INDEX)

    ' --- Capture who last saved the network file, and when -------------------------
    CurrentStep = "Reading network file's last author and save time"
    On Error Resume Next
    changeAuthor = wbNetwork.BuiltinDocumentProperties("Last Author")
    On Error GoTo ErrorHandler
    If Len(Trim(changeAuthor)) = 0 Then changeAuthor = "Unknown"

    ' FileDateTime reads the file's OS last-modified timestamp (date + time)
    changeDateTime = FileDateTime(cleanPath)

    ' --- Read headers from main sheet ------------------------------------------
    CurrentStep = "Reading headers from main sheet"
    colCount = wsMain.Cells(HEADER_ROW, wsMain.Columns.Count).End(xlToLeft).Column
    headers = wsMain.Range(wsMain.Cells(HEADER_ROW, 1), wsMain.Cells(HEADER_ROW, colCount)).Value

    ' --- Find key columns (REQUEST / NUMBER) by header name -------------------
    CurrentStep = "Locating key columns (REQUEST / NUMBER)"
    keyColRequest = FindColumnByHeader(headers, KEY_COL_NAME_1)
    keyColNumber = FindColumnByHeader(headers, KEY_COL_NAME_2)

    If keyColRequest = 0 Or keyColNumber = 0 Then
        MsgBox "Key columns REQUEST and/or NUMBER were not found in the header row.", vbCritical, "Missing Key Columns"
        GoTo CleanExit
    End If

    ' --- Read full data blocks into arrays (fast, in-memory) -------------------
    CurrentStep = "Reading main sheet data into array"
    arrOld = ReadSheetToArray(wsMain)

    CurrentStep = "Reading network sheet data into array"
    arrNew = ReadSheetToArray(wsNetwork)

    ' --- Build dictionaries: key -> row index in array --------------------------
    CurrentStep = "Building dictionary for main data"
    Set dictOld = BuildKeyDictionary(arrOld, keyColRequest, keyColNumber)

    CurrentStep = "Building dictionary for network data"
    Set dictNew = BuildKeyDictionary(arrNew, keyColRequest, keyColNumber)

    ' --- Prepare Changes sheet ---------------------------------------------------
    CurrentStep = "Preparing Changes sheet"
    Set wsChanges = PrepareChangesSheet(wbMain, headers, colCount)

    ' --- Compare and write report -------------------------------------------------
    CurrentStep = "Comparing and writing report"
    WriteComparisonReport wsChanges, arrOld, arrNew, dictOld, dictNew, _
                           colCount, changeAuthor, changeDateTime, _
                           countModified, countAdded, countDeleted

    ' --- Close network file WITHOUT saving ----------------------------------------
    CurrentStep = "Closing network file"
    wbNetwork.Close SaveChanges:=False
    networkOpenedHere = False

    ' --- Final message ---------------------------------------------------------
    MsgBox "Comparison completed." & vbCrLf & vbCrLf & _
           "Modified: " & countModified & vbCrLf & _
           "Added: " & countAdded & vbCrLf & _
           "Deleted: " & countDeleted, vbInformation, "Comparison Result"

CleanExit:
    ' --- Restore application state -------------------------------------------------
    Application.ScreenUpdating = originalScreenUpdating
    Application.EnableEvents = originalEnableEvents
    Application.Calculation = originalCalculation
    Application.DisplayAlerts = originalDisplayAlerts
    
    Sheets(Sheet6.Name).Select
    Exit Sub

ErrorHandler:
    If networkOpenedHere Then
        On Error Resume Next
        wbNetwork.Close SaveChanges:=False
        On Error GoTo 0
    End If
    MsgBox "An error occurred at step: " & CurrentStep & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    Resume CleanExit

End Sub

' ---------------------------------------------------------------------------------
' HELPER: Find a column index in a headers array by its name (case-insensitive,
' trims whitespace). Returns 0 if not found.
' ---------------------------------------------------------------------------------
Private Function FindColumnByHeader(headers As Variant, headerName As String) As Long
    Dim i As Long
    FindColumnByHeader = 0
    For i = LBound(headers, 2) To UBound(headers, 2)
        If StrComp(Trim(CStr(headers(1, i))), headerName, vbTextCompare) = 0 Then
            FindColumnByHeader = i
            Exit Function
        End If
    Next i
End Function

' ---------------------------------------------------------------------------------
' HELPER: Read a worksheet's used range (data starts right after HEADER_ROW)
' into a 2D variant array for fast in-memory processing.
' Returns Empty array (dimension 0) if there is no data.
' ---------------------------------------------------------------------------------
Private Function ReadSheetToArray(ws As Worksheet) As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim firstDataRow As Long

    firstDataRow = HEADER_ROW + 1

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < firstDataRow Then
        ReadSheetToArray = Empty
        Exit Function
    End If

    ReadSheetToArray = ws.Range(ws.Cells(firstDataRow, 1), ws.Cells(lastRow, lastCol)).Value
End Function

' ---------------------------------------------------------------------------------
' HELPER: Build a Dictionary mapping "REQUEST|NUMBER" -> row index (in the array)
' ---------------------------------------------------------------------------------
Private Function BuildKeyDictionary(arr As Variant, keyColRequest As Long, keyColNumber As Long) As Object
    Dim dict As Object
    Dim i As Long
    Dim key As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    If IsEmpty(arr) Then
        Set BuildKeyDictionary = dict
        Exit Function
    End If

    For i = LBound(arr, 1) To UBound(arr, 1)
        key = BuildKey(arr(i, keyColRequest), arr(i, keyColNumber))
        If Len(key) > 1 Then ' avoid empty rows (key would just be "|")
            If Not dict.exists(key) Then
                dict.Add key, i
            End If
        End If
    Next i

    Set BuildKeyDictionary = dict
End Function

' ---------------------------------------------------------------------------------
' HELPER: Build the composite key string
' ---------------------------------------------------------------------------------
Private Function BuildKey(requestVal As Variant, numberVal As Variant) As String
    BuildKey = Trim(CStr(requestVal)) & "|" & Trim(CStr(numberVal))
End Function

' ---------------------------------------------------------------------------------
' HELPER: Prepare (clear or create) the Changes sheet and write headers
' ---------------------------------------------------------------------------------
Private Function PrepareChangesSheet(wb As Workbook, headers As Variant, colCount As Long) As Worksheet
    Dim ws As Worksheet
    Dim exists As Boolean

    exists = False
    On Error Resume Next
    Set ws = wb.Sheets(CHANGES_SHEET_NAME)
    On Error GoTo 0

    If Not ws Is Nothing Then
        exists = True
        ws.Cells.Clear
    Else
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = CHANGES_SHEET_NAME
    End If

    ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount)).Value = headers
    ws.Cells(1, colCount + 1).Value = HEADER_CHANGED_BY
    ws.Cells(1, colCount + 2).Value = HEADER_TIME
    ws.Cells(1, colCount + 3).Value = HEADER_DATE

    ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount + EXTRA_COL_COUNT)).Font.Bold = True

    Set PrepareChangesSheet = ws
End Function

' ---------------------------------------------------------------------------------
' HELPER: Perform the actual comparison and write results (Modified / Added / Deleted)
' ---------------------------------------------------------------------------------
Private Sub WriteComparisonReport(wsChanges As Worksheet, arrOld As Variant, arrNew As Variant, _
                                   dictOld As Object, dictNew As Object, colCount As Long, _
                                   changeAuthor As String, changeDateTime As Date, _
                                   ByRef countModified As Long, ByRef countAdded As Long, ByRef countDeleted As Long)

    Dim outRow As Long
    Dim key As Variant
    Dim rOld As Long, rNew As Long
    Dim c As Long
    Dim changedCols As Object
    Dim isDifferent As Boolean
    Dim timeText As String
    Dim dateText As String

    ' Buffer for writing rows in bulk (write per-row to keep it simple & fast enough,
    ' avoiding per-cell Select/Activate; only formatting touches specific cells)
    Dim rowBuffer() As Variant

    timeText = Format(changeDateTime, "hh:mm:ss")
    dateText = GregorianToJalali(Year(changeDateTime), Month(changeDateTime), Day(changeDateTime))

    countModified = 0
    countAdded = 0
    countDeleted = 0
    outRow = 2 ' row 1 = headers

    ' --- Pass 1: keys present in NEW file (covers Modified + Added) ---------------
    For Each key In dictNew.Keys

        rNew = dictNew(key)

        If dictOld.exists(key) Then
            ' ---- Possibly Modified ----
            rOld = dictOld(key)
            Set changedCols = CreateObject("Scripting.Dictionary")
            isDifferent = False

            For c = 1 To colCount
                If Not ValuesEqual(arrOld(rOld, c), arrNew(rNew, c)) Then
                    isDifferent = True
                    changedCols.Add c, True
                End If
            Next c

            If isDifferent Then
                ReDim rowBuffer(1 To 1, 1 To colCount)
                For c = 1 To colCount
                    rowBuffer(1, c) = arrNew(rNew, c)
                Next c
                wsChanges.Range(wsChanges.Cells(outRow, 1), wsChanges.Cells(outRow, colCount)).Value = rowBuffer

                ' Highlight only the changed cells (yellow) and attach a comment
                ' showing the previous value, visible on mouse hover
                For c = 1 To colCount
                    If changedCols.exists(c) Then
                        wsChanges.Cells(outRow, c).Interior.Color = COLOR_MODIFIED
                        AddOldValueComment wsChanges.Cells(outRow, c), arrOld(rOld, c)
                    End If
                Next c

                wsChanges.Cells(outRow, colCount + 1).Value = changeAuthor
                wsChanges.Cells(outRow, colCount + 2).Value = timeText
                wsChanges.Cells(outRow, colCount + 3).Value = dateText

                countModified = countModified + 1
                outRow = outRow + 1
            End If

        Else
            ' ---- Added (only in new file) ----
            ReDim rowBuffer(1 To 1, 1 To colCount)
            For c = 1 To colCount
                rowBuffer(1, c) = arrNew(rNew, c)
            Next c
            wsChanges.Range(wsChanges.Cells(outRow, 1), wsChanges.Cells(outRow, colCount)).Value = rowBuffer
            wsChanges.Range(wsChanges.Cells(outRow, 1), wsChanges.Cells(outRow, colCount)).Interior.Color = COLOR_ADDED

            wsChanges.Cells(outRow, colCount + 1).Value = changeAuthor
            wsChanges.Cells(outRow, colCount + 2).Value = timeText
            wsChanges.Cells(outRow, colCount + 3).Value = dateText

            countAdded = countAdded + 1
            outRow = outRow + 1
        End If

    Next key

    ' --- Pass 2: keys only in OLD file (Deleted) -----------------------------------
    For Each key In dictOld.Keys
        If Not dictNew.exists(key) Then
            rOld = dictOld(key)
            ReDim rowBuffer(1 To 1, 1 To colCount)
            For c = 1 To colCount
                rowBuffer(1, c) = arrOld(rOld, c)
            Next c
            wsChanges.Range(wsChanges.Cells(outRow, 1), wsChanges.Cells(outRow, colCount)).Value = rowBuffer
            wsChanges.Range(wsChanges.Cells(outRow, 1), wsChanges.Cells(outRow, colCount)).Interior.Color = COLOR_DELETED

            wsChanges.Cells(outRow, colCount + 1).Value = changeAuthor
            wsChanges.Cells(outRow, colCount + 2).Value = timeText
            wsChanges.Cells(outRow, colCount + 3).Value = dateText

            countDeleted = countDeleted + 1
            outRow = outRow + 1
        End If
    Next key

    wsChanges.Columns.AutoFit

End Sub

' ---------------------------------------------------------------------------------
' HELPER: Attach a comment (Note) to a cell showing its previous value.
' This is the classic hover-comment (not a threaded/modern comment), so it
' shows a tooltip on mouse-over exactly like a normal Excel cell comment.
' ---------------------------------------------------------------------------------
Private Sub AddOldValueComment(targetCell As Range, oldValue As Variant)
    Dim commentText As String
    Dim displayValue As String

    If IsEmpty(oldValue) Or IsNull(oldValue) Then
        displayValue = "(empty)"
    Else
        displayValue = CStr(oldValue)
        If Len(Trim(displayValue)) = 0 Then displayValue = "(empty)"
    End If

    commentText = "Previous value: " & displayValue

    ' Remove any pre-existing comment first (Changes sheet is rebuilt each run,
    ' but this guards against re-runs on the same live sheet without clearing)
    If Not targetCell.Comment Is Nothing Then
        targetCell.Comment.Delete
    End If

    targetCell.AddComment Text:=commentText
    targetCell.Comment.Shape.TextFrame.AutoSize = True
End Sub

' ---------------------------------------------------------------------------------
' HELPER: Convert a Gregorian date to a Jalali (Shamsi/Persian) date string
' in "yyyy/mm/dd" format. Standard, widely-used conversion algorithm.
' ---------------------------------------------------------------------------------
Private Function GregorianToJalali(gy As Long, gm As Long, gd As Long) As String
    Dim g_d_m As Variant
    g_d_m = Array(0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334)

    Dim jy As Long, jm As Long, jd As Long
    Dim gy2 As Long
    Dim gyTemp As Long
    Dim days As Long

    If gy > 1600 Then
        jy = 979
        gy2 = gy - 1600
    Else
        jy = 0
        gy2 = gy - 621
    End If

    If gm > 2 Then
        gyTemp = gy2 + 1
    Else
        gyTemp = gy2
    End If

    days = 365 * gy2 + (Int((gyTemp + 3) / 4)) - (Int((gyTemp + 99) / 100)) + _
           (Int((gyTemp + 399) / 400)) - 80 + gd + g_d_m(gm - 1)

    jy = jy + 33 * Int(days / 12053)
    days = days Mod 12053

    jy = jy + 4 * Int(days / 1461)
    days = days Mod 1461

    If days > 365 Then
        jy = jy + Int((days - 1) / 365)
        days = (days - 1) Mod 365
    End If

    If days < 186 Then
        jm = 1 + Int(days / 31)
        jd = 1 + (days Mod 31)
    Else
        jm = 7 + Int((days - 186) / 30)
        jd = 1 + ((days - 186) Mod 30)
    End If

    GregorianToJalali = CStr(jy) & "/" & Format(jm, "00") & "/" & Format(jd, "00")
End Function

' ---------------------------------------------------------------------------------
' HELPER: Compare two cell values safely (handles dates, numbers, text, blanks)
' ---------------------------------------------------------------------------------
Private Function ValuesEqual(v1 As Variant, v2 As Variant) As Boolean

    Dim s1 As String
    Dim s2 As String

    ' Treat Empty/Null as blank string for comparison purposes
    If IsError(v1) Or IsError(v2) Then
        ValuesEqual = False
        Exit Function
    End If

    s1 = Trim(CStr(v1))
    s2 = Trim(CStr(v2))

    ValuesEqual = (StrComp(s1, s2, vbTextCompare) = 0)

End Function
