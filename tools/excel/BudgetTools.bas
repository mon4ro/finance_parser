Attribute VB_Name = "BudgetTools"
Option Explicit

Private Const UNIFIED_SHEET_NAME As String = "UnifiedTransactions"
Private Const HEADER_ROW As Long = 1

Public Sub SplitSelectedTransaction()
    Dim ws As Worksheet
    Dim parentRow As Long
    Dim splitCountRaw As Variant
    Dim splitCount As Long
    Dim parentUnifiedID As String
    Dim parentRawID As String
    Dim firstSuffix As Long
    Dim i As Long
    Dim childRow As Long
    Dim childUnifiedID As String
    Dim msg As String

    On Error GoTo FailCleanly

    Set ws = ActiveSheet
    If ws.Name <> UNIFIED_SHEET_NAME Then
        MsgBox "Please select the transaction row on the '" & UNIFIED_SHEET_NAME & "' sheet first.", vbExclamation, "Split transaction"
        Exit Sub
    End If

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell on the transaction row you want to split.", vbExclamation, "Split transaction"
        Exit Sub
    End If

    parentRow = Selection.Row
    If parentRow <= HEADER_ROW Then
        MsgBox "Please select a transaction row, not the header row.", vbExclamation, "Split transaction"
        Exit Sub
    End If

    RequireHeader ws, "UnifiedID"
    RequireHeader ws, "RawID"
    RequireHeader ws, "TransactionType"
    RequireHeader ws, "Include"
    RequireHeader ws, "Amount"

    parentUnifiedID = Trim(CStr(GetCellByHeader(ws, parentRow, "UnifiedID").Value))
    parentRawID = Trim(CStr(GetCellByHeader(ws, parentRow, "RawID").Value))

    If parentUnifiedID = "" Then
        MsgBox "Selected row has blank UnifiedID. Cannot create split rows safely.", vbCritical, "Split transaction"
        Exit Sub
    End If

    splitCountRaw = Application.InputBox( _
        Prompt:="How many split rows do you want to create?", _
        Title:="Split transaction", _
        Default:=2, _
        Type:=1 _
    )

    If splitCountRaw = False Then Exit Sub

    splitCount = CLng(splitCountRaw)
    If splitCount < 2 Then
        MsgBox "Split count must be at least 2.", vbExclamation, "Split transaction"
        Exit Sub
    End If

    If splitCount > 25 Then
        MsgBox "Split count above 25 is probably a mistake. Aborting.", vbExclamation, "Split transaction"
        Exit Sub
    End If

    msg = "This will:" & vbCrLf & vbCrLf & _
          "- set the original row Include = NO" & vbCrLf & _
          "- insert " & splitCount & " split rows below it" & vbCrLf & _
          "- copy source/date/receiver fields from the original" & vbCrLf & _
          "- clear Amount and category fields on the split rows" & vbCrLf & vbCrLf & _
          "Continue?"

    If MsgBox(msg, vbQuestion + vbYesNo, "Split transaction") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    firstSuffix = NextSplitSuffix(ws, HeaderColumn(ws, "UnifiedID"), parentUnifiedID)

    ws.Rows(CStr(parentRow + 1) & ":" & CStr(parentRow + splitCount)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    For i = 1 To splitCount
        childRow = parentRow + i
        ws.Rows(parentRow).Copy Destination:=ws.Rows(childRow)

        childUnifiedID = parentUnifiedID & "-S" & Format(firstSuffix + i - 1, "00")

        SetIfHeaderExists ws, childRow, "UnifiedID", childUnifiedID
        SetIfHeaderExists ws, childRow, "RawID", parentRawID
        SetIfHeaderExists ws, childRow, "TransactionType", "Split"
        SetIfHeaderExists ws, childRow, "Include", "YES"

        ClearIfHeaderExists ws, childRow, "Amount"
        ClearIfHeaderExists ws, childRow, "Supercategory"
        ClearIfHeaderExists ws, childRow, "Category"
        ClearIfHeaderExists ws, childRow, "Subcategory"
        ClearIfHeaderExists ws, childRow, "Tag"

        SetIfHeaderExists ws, childRow, "Comments", "Split from " & parentUnifiedID
        SetIfHeaderExists ws, childRow, "ParentUnifiedID", parentUnifiedID
    Next i

    SetIfHeaderExists ws, parentRow, "Include", "NO"
    AppendComment ws, parentRow, "Split into " & splitCount & " rows"

    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    GetCellByHeader(ws, parentRow + 1, "Amount").Select

    MsgBox "Created " & splitCount & " split rows. Fill Amount and manual categorisation fields next.", vbInformation, "Split transaction"
    Exit Sub

FailCleanly:
    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Split failed: " & Err.Description, vbCritical, "Split transaction"
End Sub

Public Sub ValidateSelectedSplitGroup()
    Dim ws As Worksheet
    Dim selectedRow As Long
    Dim parentUnifiedID As String
    Dim parentRow As Long
    Dim amountCol As Long
    Dim unifiedIDCol As Long
    Dim parentAmount As Double
    Dim childSum As Double
    Dim lastRow As Long
    Dim r As Long
    Dim uid As String
    Dim diff As Double

    On Error GoTo FailCleanly

    Set ws = ActiveSheet
    If ws.Name <> UNIFIED_SHEET_NAME Then
        MsgBox "Please select a split parent or child row on the '" & UNIFIED_SHEET_NAME & "' sheet first.", vbExclamation, "Validate split"
        Exit Sub
    End If

    RequireHeader ws, "UnifiedID"
    RequireHeader ws, "Amount"

    selectedRow = Selection.Row
    If selectedRow <= HEADER_ROW Then
        MsgBox "Please select a transaction row, not the header row.", vbExclamation, "Validate split"
        Exit Sub
    End If

    unifiedIDCol = HeaderColumn(ws, "UnifiedID")
    amountCol = HeaderColumn(ws, "Amount")

    parentUnifiedID = CStr(ws.Cells(selectedRow, unifiedIDCol).Value)
    If InStr(1, parentUnifiedID, "-S", vbTextCompare) > 0 Then
        parentUnifiedID = Left(parentUnifiedID, InStrRev(parentUnifiedID, "-S") - 1)
    End If

    parentRow = FindRowByUnifiedID(ws, unifiedIDCol, parentUnifiedID)
    If parentRow = 0 Then
        MsgBox "Could not find split parent row: " & parentUnifiedID, vbCritical, "Validate split"
        Exit Sub
    End If

    parentAmount = ParseAmount(ws.Cells(parentRow, amountCol).Value)
    childSum = 0

    lastRow = ws.Cells(ws.Rows.Count, unifiedIDCol).End(xlUp).Row
    For r = HEADER_ROW + 1 To lastRow
        uid = CStr(ws.Cells(r, unifiedIDCol).Value)
        If Left(uid, Len(parentUnifiedID & "-S")) = parentUnifiedID & "-S" Then
            childSum = childSum + ParseAmount(ws.Cells(r, amountCol).Value)
        End If
    Next r

    diff = childSum - parentAmount

    If Abs(diff) < 0.005 Then
        MsgBox "Split is balanced." & vbCrLf & _
               "Parent amount: " & Format(parentAmount, "0.00") & vbCrLf & _
               "Split sum: " & Format(childSum, "0.00"), vbInformation, "Validate split"
    Else
        MsgBox "Split is NOT balanced." & vbCrLf & _
               "Parent amount: " & Format(parentAmount, "0.00") & vbCrLf & _
               "Split sum: " & Format(childSum, "0.00") & vbCrLf & _
               "Difference: " & Format(diff, "0.00"), vbExclamation, "Validate split"
    End If

    Exit Sub

FailCleanly:
    MsgBox "Validation failed: " & Err.Description, vbCritical, "Validate split"
End Sub

Private Function HeaderColumn(ws As Worksheet, headerName As String) As Long
    Dim lastCol As Long
    Dim col As Long
    Dim value As String

    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column

    For col = 1 To lastCol
        value = Trim(CStr(ws.Cells(HEADER_ROW, col).Value))
        If StrComp(value, headerName, vbTextCompare) = 0 Then
            HeaderColumn = col
            Exit Function
        End If
    Next col

    HeaderColumn = 0
End Function

Private Sub RequireHeader(ws As Worksheet, headerName As String)
    If HeaderColumn(ws, headerName) = 0 Then
        Err.Raise vbObjectError + 1000, "BudgetTools", "Missing required column: " & headerName
    End If
End Sub

Private Function GetCellByHeader(ws As Worksheet, rowNum As Long, headerName As String) As Range
    Dim col As Long

    col = HeaderColumn(ws, headerName)
    If col = 0 Then
        Err.Raise vbObjectError + 1001, "BudgetTools", "Missing required column: " & headerName
    End If

    Set GetCellByHeader = ws.Cells(rowNum, col)
End Function

Private Sub SetIfHeaderExists(ws As Worksheet, rowNum As Long, headerName As String, value As Variant)
    Dim col As Long

    col = HeaderColumn(ws, headerName)
    If col > 0 Then
        ws.Cells(rowNum, col).Value = value
    End If
End Sub

Private Sub ClearIfHeaderExists(ws As Worksheet, rowNum As Long, headerName As String)
    Dim col As Long

    col = HeaderColumn(ws, headerName)
    If col > 0 Then
        ws.Cells(rowNum, col).ClearContents
    End If
End Sub

Private Sub AppendComment(ws As Worksheet, rowNum As Long, note As String)
    Dim col As Long
    Dim existing As String

    col = HeaderColumn(ws, "Comments")
    If col = 0 Then Exit Sub

    existing = Trim(CStr(ws.Cells(rowNum, col).Value))
    If existing = "" Then
        ws.Cells(rowNum, col).Value = note
    ElseIf InStr(1, existing, note, vbTextCompare) = 0 Then
        ws.Cells(rowNum, col).Value = existing & "; " & note
    End If
End Sub

Private Function NextSplitSuffix(ws As Worksheet, unifiedIDCol As Long, parentUnifiedID As String) As Long
    Dim lastRow As Long
    Dim r As Long
    Dim uid As String
    Dim suffixText As String
    Dim suffixNumber As Long
    Dim maxSuffix As Long

    maxSuffix = 0
    lastRow = ws.Cells(ws.Rows.Count, unifiedIDCol).End(xlUp).Row

    For r = HEADER_ROW + 1 To lastRow
        uid = CStr(ws.Cells(r, unifiedIDCol).Value)
        If Left(uid, Len(parentUnifiedID & "-S")) = parentUnifiedID & "-S" Then
            suffixText = Mid(uid, Len(parentUnifiedID & "-S") + 1)
            If IsNumeric(suffixText) Then
                suffixNumber = CLng(suffixText)
                If suffixNumber > maxSuffix Then maxSuffix = suffixNumber
            End If
        End If
    Next r

    NextSplitSuffix = maxSuffix + 1
End Function

Private Function FindRowByUnifiedID(ws As Worksheet, unifiedIDCol As Long, unifiedID As String) As Long
    Dim lastRow As Long
    Dim r As Long

    lastRow = ws.Cells(ws.Rows.Count, unifiedIDCol).End(xlUp).Row

    For r = HEADER_ROW + 1 To lastRow
        If CStr(ws.Cells(r, unifiedIDCol).Value) = unifiedID Then
            FindRowByUnifiedID = r
            Exit Function
        End If
    Next r

    FindRowByUnifiedID = 0
End Function

Private Function ParseAmount(value As Variant) As Double
    If IsNumeric(value) Then
        ParseAmount = CDbl(value)
    Else
        ParseAmount = CDbl(Val(Replace(CStr(value), ",", ".")))
    End If
End Function
