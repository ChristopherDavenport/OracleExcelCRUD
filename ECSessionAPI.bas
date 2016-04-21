Attribute VB_Name = "ECSessionAPI"
Option Explicit
Public ECSession As ECSession

Private TableRange As Range

Private Function ValidSession() As Boolean
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
Private Function ValidateTable(SheetName As String, TableName As String) As Boolean
    On Error GoTo catch
    ValidateTable = False
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    
    Set wb = ThisWorkbook
    
    For Each ws In wb.Worksheets
        If ws.Name = SheetName Then
            For Each lo In ws.ListObjects
                If lo.Name = TableName Then
                    ValidateTable = True
                    Exit Function
                End If
            Next
        End If
    Next
    
    Exit Function
catch:
    ValidateTable = False
        
End Function
Private Sub ValidateSession()
    If Not ValidSession Then End
End Sub
Private Sub ValidateTable(SheetName As String, TableName As String)
    If Not ValidTable Then
        MsgBox ("Invalid Table")
        End
    End If
    Set TableRange = ThisWorkbook.Sheets(SheetName).ListObjects(TableName).Range
End Sub


Sub Insert(SheetName As String, TableName As String, OracleTableName As String, Optional ColorResults As Boolean = False)
        ValidateSession
        ValidateTable SheetName, TableName
        ECSession.Insert OracleTableName, TableRange, ColorResults
End Sub

