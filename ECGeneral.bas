Attribute VB_Name = "ECGeneral"
Public ECSession As ECSession


Function ValidSession() As Boolean
    If (Not ECSession Is Nothing) Then
        If (ECSession.Validated = True) Then
            ValidSession = True
        Else
            MsgBox "Session Not Validated, Please Login Again"
            ValidSession = False
        End If
    Else
        MsgBox "No Current Session, Please Login"
        ValidSession = False
    End If
End Function

Function CurrentUser() As String
  CurrentUser = ECSession.Username
End Function

Sub RefreshTable(SheetName As String, TableName As String)
    Dim TableRange As Range
    Set TableRange = ThisWorkbook.Sheets(SheetName).ListObjects(TableName).Range
    TableRange.Dirty
    TableRange.Calculate
End Sub



Sub ClearTable(SheetName As String, TableName As String, Optional ClearContents As Boolean = False)
Dim loSource As Excel.ListObject
Set loSource = ThisWorkbook.Sheets(SheetName).ListObjects(TableName)
With loSource
    .Range.Interior.ColorIndex = -4142
    If (.DataBodyRange.Rows.Count > 1) Then
    .DataBodyRange.Offset(1).Resize(.DataBodyRange.Rows.Count - 1, .DataBodyRange.Columns.Count).Rows.Delete
    End If
    If ClearContents = True Then
        .DataBodyRange.ClearContents
    End If
End With
End Sub

Sub ClearValid(SheetName As String, TableName As String)
    ClearOnColor SheetName, TableName, RGB(124, 252, 0)
End Sub

Sub ClearInvalid(SheetName As String, TableName As String)
    ClearOnColor SheetName, TableName, RGB(255, 0, 0)
End Sub

Private Sub ClearOnColor(SheetName As String, TableName As String, RGBCheck As Long)
Start:
    Dim loSource As Excel.ListObject
    Set loSource = ThisWorkbook.Sheets(SheetName).ListObjects(TableName)
    Dim lr As Excel.ListRow

    For Each lr In loSource.ListRows
        If (Cells(lr.Range.Row, lr.Range.Column).Interior.Color = RGBCheck) Then
            lr.Delete
            ' The table has changed so we need to reevaluate
            ' the range based on the changed conditions
            GoTo Start
        End If
    Next lr
    loSource.Range.Interior.ColorIndex = -4142
End Sub




