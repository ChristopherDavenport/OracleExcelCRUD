VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ECSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Notes

'Must Reference Microsoft ActiveX Data Objects Library

'Declarations
Private pUsername As String
Private pPassword As String
Private pDSN As String
Private pValidated As Boolean
Private pValidCommand As Boolean
Private pConnection As ADODB.Connection
Private pCommand As ADODB.Command
Private pDebug As Boolean


' Username Property
Public Property Get Username() As String
    Username = pUsername
End Property

'DSN Property
Public Property Get DSN() As String
    DSN = pDSN
End Property

'Validated Property
'Has Get but no Let
Public Property Get Validated() As Boolean
    Validated = pValidated
End Property

Public Property Let DebugClass(WhetherToDebug As Boolean)
    pDebug = WhetherToDebug
End Property

'Init Sub
'Allows User to Assign All Values With 1 Function
Sub Initialize(Username As String, Password As String, DSN As String)
    On Error GoTo Error
    pUsername = ParsedLoginCredentials(Username)
    pPassword = ParsedLoginCredentials(Password)
    pDSN = ParsedLoginCredentials(DSN)
    pValidated = False
    
    On Error GoTo ExitHere
    Validate
    GoTo ExitHere
Error:
    EckerdErrorHandler
ExitHere:
    
End Sub

Sub Reset_Password(txtNewPassword1 As String, txtNewPassword2 As String)
    Dim Valid As Boolean
    Dim NewPass As String
    Dim strSQL As String
    pValidCommand = True
    'Full New Password Validation Can Be Performed Here
    If (txtNewPassword1 <> txtNewPassword2) Then
        MsgBox "Your new passwords entered do not match"
        Valid = False
        GoTo ExitHere
    End If
    
    NewPass = ParsedLoginCredentials(txtNewPassword1)
    
    If pValidCommand = True Then
        Open_Session
        strSQL = "alter user " & pUsername & " identified by " & NewPass & " replace " & pPassword & ";"
        'strSQL = "alter user " & pUsername & " identified by " & pPassword & " replace " & txtNewPassword1 & ";"
        
        If pDebug = True Then
            MsgBox strSQL
        Else
            Execute_Command strSQL
        End If
    End If
    
    If pValidCommand = True Then
        pPassword = NewPass
    End If
    GoTo ExitHere
Error:
    EckerdErrorHandler
ExitHere:
    Close_Session
End Sub
' The insert command takes a table of variable size,
' uses the headers as the column names
' and inserts into the table indicated through the connection
Sub Insert(TableName As String, Selection As Range, Optional ColorResults As Boolean = False)
    On Error GoTo Error
    
    Dim strSQL As String
    Dim nCurrentCell As Range
    Dim nCurrentHeaderCell As Range
    
    Open_Session
    
    Selection.Dirty
    Selection.Calculate
    
    'Generate Variables for the Size of the Range
    nFirstRow = Selection.Row
    nLastRow = Selection.Rows.Count + Selection.Row - 1
    nCurrentRow = Selection.Row
        
    nFirstColumn = Selection.Column
    nLastColumn = Selection.Columns.Count + Selection.Column - 1
    nCurrentColumn = Selection.Column
             
    'Column Headers Assigned
    'First Row of the Range is used as the column identifier for the insert
    'This should fully generate strHeaderSQL
    strHeaderSQL = "("
    Do While nCurrentColumn <= nLastColumn
        Set nCurrentCell = Cells(nCurrentRow, nCurrentColumn)

        strHeaderSQL = strHeaderSQL & ParsedHeaderData(nCurrentCell)
                
        If (nCurrentColumn < nLastColumn) Then
            strHeaderSQL = strHeaderSQL & ", "
        End If
        If (nCurrentColumn = nLastColumn) Then
            strHeaderSQL = strHeaderSQL & ") "
        End If
                
        nCurrentColumn = nCurrentColumn + 1
    Loop
    'First - Reset Current Cells
    nCurrentColumn = nFirstColumn
    nCurrentRow = nCurrentRow + 1
    
    For R = nCurrentRow To nLastRow
        'Reset first column
        nCurrentColumn = nFirstColumn
        pValidCommand = True
        
        
        
        'Generate the body of the insert statement for each row
        strBodySQL = "VALUES ("
                
        Do While (nCurrentColumn <= nLastColumn)
            Set nCurrentCell = Cells(nCurrentRow, nCurrentColumn)
            
            If ColorResults = True Then
                Cells(nCurrentRow, nCurrentColumn).Interior.Color = RGB(255, 0, 0)
            End If
                    
            strBodySQL = strBodySQL & ParsedCellData(nCurrentCell)
            
            If (nCurrentColumn < nLastColumn) Then
                strBodySQL = strBodySQL & ", "
            End If
            If (nCurrentColumn = nLastColumn) Then
                strBodySQL = strBodySQL & ")"
            End If
            nCurrentColumn = nCurrentColumn + 1
                    
        Loop
        strStartSql = "INSERT INTO " & TableName & " "
        strFooterSql = ";"
                
        strSQL = strStartSql & strHeaderSQL & strBodySQL & strFooterSql
                                
        ' Does Not Execute If It Triggers One Of Our Built In Errors
        If (pValidCommand = True) Then
            If pDebug = True Then
                MsgBox strSQL
            Else
                Execute_Command strSQL
            End If
        End If
                
        ' Does Not Turn Green If It Fails Our Built in Errors Or During Execution
        If (pValidCommand = True And ColorResults = True) Then
            nCurrentColumn = nFirstColumn
            Do While nCurrentColumn <= nLastColumn
                Cells(nCurrentRow, nCurrentColumn).Interior.Color = RGB(124, 252, 0)
                nCurrentColumn = nCurrentColumn + 1
            Loop
        End If
        
        nCurrentRow = nCurrentRow + 1
    Next R
    GoTo ExitHere
    
Error:
    EckerdErrorHandler
    Resume Next
ExitHere:
    Close_Session
End Sub

' This Update Sub Takes TableName and a Variable Range and Updates the values using the Primary Key.
' The PRIMARY KEY MUST BE LOCATED IN THE LEFTMOST COLUMN
' Syntax of final command should be
'
'   Update TableName
'   SET column1=value1,column2=value2,...
'   WHERE column1=value1
'
'   IF YOU SCREW UP THE LEFT COLUMN YOU CAN DESTROY THE ENTIRE TABLE


Sub Update(TableName As String, Selection As Range, Optional PrimaryKeys As Integer = 1, Optional ColorResults As Boolean = False)
    Open_Session
    On Error GoTo Error
    Dim strSQL As String
    Dim nCurrentCell As Range
    Dim nCurrentHeader As Range
    Dim pkRow As Integer
    
    
    
    'Generate Variables for the Size of the Range
    nFirstRow = Selection.Row
    nLastRow = Selection.Rows.Count + Selection.Row - 1
    nCurrentRow = Selection.Row
        
    nFirstColumn = Selection.Column
    nLastColumn = Selection.Columns.Count + Selection.Column - 1
    nCurrentColumn = Selection.Column
       
    'First - Reset Current Cells
    nCurrentColumn = nFirstColumn
    nCurrentRow = nCurrentRow + 1
    
    For R = nCurrentRow To nLastRow
        'Reset to first column, reset valid, strWhereSqlEnd, strSetSqlEnd
        nCurrentColumn = nFirstColumn
        pValidCommand = True
        strWhereSqlEnd = ""
        strSetSqlEnd = ""
        pkRow = PrimaryKeys
        
        If ColorResults = True Then
            Do While nCurrentColumn <= nLastColumn
                Cells(nCurrentRow, nCurrentColumn).Interior.Color = RGB(255, 0, 0)
                nCurrentColumn = nCurrentColumn + 1
            Loop
            nCurrentColumn = nFirstColumn
        End If
                
        Do While (nCurrentColumn <= nLastColumn)
            Set nCurrentCell = Cells(nCurrentRow, nCurrentColumn)
            Set nCurrentHeader = Cells(nFirstRow, nCurrentColumn)
            
            
            
            ' Generates Where Clause Dependent On How Many Primary Keys Are Needed AND added except for last
            If pkRow > 0 Then
                strWhereSqlEnd = strWhereSqlEnd & ParsedHeaderData(nCurrentHeader) & _
                    "=" & ParsedCellData(nCurrentCell, True)
                If pkRow > 1 Then
                    strWhereSqlEnd = strWhereSqlEnd & " AND "
                End If
                pkRow = pkRow - 1
            End If
                
            'Generates Set Clause - Commas between each item except last
            strSetSqlEnd = strSetSqlEnd & ParsedHeaderData(nCurrentHeader) & "=" & ParsedCellData(nCurrentCell)
            If (nCurrentColumn < nLastColumn) Then
                strSetSqlEnd = strSetSqlEnd & ", "
            End If
            
            nCurrentColumn = nCurrentColumn + 1
                    
        Loop
        
        strStartSql = "UPDATE " & TableName & " "
        strSetSql = "SET " & strSetSqlEnd
        strWhereSql = "WHERE " & strWhereSqlEnd
        strFooterSql = ";"
    
        strSQL = strStartSql & " " & strSetSql & " " & strWhereSql & strFooterSql
    
        
        ' Does Not Execute If It Triggers One Of Our Built In Errors
        If (pValidCommand = True) Then
            If pDebug = True Then
                MsgBox strSQL
            Else
                Execute_Command strSQL
            End If
        End If
        
        If (pValidCommand = True And ColorResults = True) Then
            Do While nCurrentColumn <= nLastColumn
                Cells(nCurrentRow, nCurrentColumn).Interior.Color = RGB(124, 252, 0)
                nCurrentColumn = nCurrentColumn + 1
            Loop
            nCurrentColumn = nFirstColumn
        End If
        
    nCurrentRow = nCurrentRow + 1
    Next R
    GoTo ExitHere
    
Error:
    EckerdErrorHandler
    Resume Next
ExitHere:
    Close_Session
    
End Sub



'We Take Only The First Column As The Primary Key to Delete From The Table Specific Entries

Sub Delete(TableName As String, Selection As Range, Optional PrimaryKeys As Integer = 1, Optional ColorResults As Boolean = False)
    On Error GoTo Error
    Open_Session
    
    Dim nCurrentCell As Range
    Dim nCurrentHeader As Range
    Dim strSQL As String
    
    'Generate Variables for the Size of the Range
    nFirstRow = Selection.Row
    nLastRow = Selection.Rows.Count + Selection.Row - 1
    nCurrentRow = nFirstRow
        
    nFirstColumn = Selection.Column
    nLastColumn = Selection.Columns.Count + Selection.Column - 1
    nCurrentColumn = Selection.Column
    
    'First - Reset Current Cells
    nCurrentColumn = nFirstColumn
    nCurrentRow = nCurrentRow + 1
    
    For R = nCurrentRow To nLastRow
        pValidCommand = True
        strWhereSqlEnd = ""
        pkRow = PrimaryKeys
        nCurrentColumn = nFirstColumn
        
        If ColorResults = True Then
            nCurrentColumn = nFirstColumn
            Do While nCurrentColumn <= nLastColumn
                Cells(nCurrentRow, nCurrentColumn).Interior.Color = RGB(255, 0, 0)
                nCurrentColumn = nCurrentColumn + 1
            Loop
            nCurrentColumn = nFirstColumn
        End If
        
        Do While pkRow > 0
            'Current Cell and Header
            Set nCurrentCell = Cells(nCurrentRow, nCurrentColumn)
            Set nCurrentHeader = Cells(nFirstRow, nCurrentColumn)
            
            ' Generates Where Clause Dependent On How Many Primary Keys Are Needed AND added except for last
            If pkRow > 0 Then
                strWhereSqlEnd = strWhereSqlEnd & ParsedHeaderData(nCurrentHeader) & _
                    "=" & ParsedCellData(nCurrentCell, True)
                If pkRow > 1 Then
                    strWhereSqlEnd = strWhereSqlEnd & " AND "
                End If
                pkRow = pkRow - 1
            End If
            nCurrentColumn = nCurrentColumn + 1
        Loop
        strStartSql = "DELETE FROM " & TableName & " "
        strWhereSql = "WHERE " & strWhereSqlEnd
        strEndSql = ";"
        
        strSQL = strStartSql & strWhereSql & strEndSql
        
        'Does Not Execute If It Triggers One Of Our Built In Errors
        If (pValidCommand = True) Then
            If pDebug = True Then
                MsgBox strSQL
            Else
                Execute_Command strSQL
            End If
        End If
        
        If (pValidCommand = True And ColorResults = True) Then
            nCurrentColumn = nFirstColumn
            Do While nCurrentColumn <= nLastColumn
                Cells(nCurrentRow, nCurrentColumn).Interior.Color = RGB(124, 252, 0)
                nCurrentColumn = nCurrentColumn + 1
            Loop
            nCurrentColumn = nFirstColumn
        End If
        
        nCurrentRow = nCurrentRow + 1
    Next R
    GoTo ExitHere
    
Error:
    EckerdErrorHandler
    Resume Next
ExitHere:
    Close_Session
    
End Sub

Sub StoredProcedure(ProcedureName As String, Optional Arguments As String = "")
    On Error GoTo Error
    Dim strSQL As String
    pValidCommand = True
    Open_Session
    
    strSQL = "{CALL " & ProcedureName & "(" & Arguments & ")}"
    If pValidCommand = True Then
        If pDebug = True Then
            MsgBox strSQL
        Else
            Execute_Command strSQL
        End If
    End If
    GoTo ExitHere:

Error:
    EckerdErrorHandler
    Resume Next
ExitHere:
    Close_Session
        
End Sub

Sub UpdateHardCode(TableName As String, SetStatement As String, WhereStatement As String)
    On Error GoTo Error
    Dim strSQL As String
    pValidCommand = True
    Open_Session
    
    
    strSQL = "UPDATE " & TableName & " SET " & SetStatement & " WHERE " & WhereStatement & ";"
    
    If pValidCommand = True Then
        If pDebug = True Then
            MsgBox strSQL
        Else
            Execute_Command strSQL
        End If
    End If
    GoTo ExitHere
Error:
    EckerdErrorHandler
ExitHere:
    Close_Session
End Sub

Sub DeleteHardCode(TableName As String, WhereStatement As String)
    On Error GoTo Error
    Dim strSQL As String
    pValidCommand = True
    Open_Session
    
    strSQL = "DELETE " & TableName & " WHERE " & WhereStatement & ";"
    If pValidCommand = True Then
        If pDebug = True Then
            MsgBox strSQL
        Else
            Execute_Command strSQL
        End If
    End If
    
    GoTo ExitHere
Error:
    EckerdErrorHandler
ExitHere:
    Close_Session
End Sub



Sub SQLCommand(strSQL As String)
    On Error GoTo Error
    Open_Session
    Execute_Command strSQL
    GoTo ExitHere
Error:
    EckerdErrorHandler
    GoTo ExitHere
ExitHere:
    Close_Session
End Sub


'Private Subroutines
Private Sub Open_Session()
    Dim strConn As String
    Dim objErr As ADODB.Error
    strCon = "dsn=" & pDSN & ";uid=" & pUsername & ";pwd=" & pPassword
    If pDebug = True Then
        MsgBox strCon
    End If
    Set pConnection = New ADODB.Connection
    pConnection.Open strCon
End Sub

Private Sub Close_Session()
    On Error Resume Next
    pConnection.Close
    Set pConnection = Nothing
End Sub

'Validate Session Credentials
'If Credentials are valid then change
'persistent validated property to true

Private Sub Validate()
    On Error GoTo ConnectionError
    Open_Session
    If (pConnection.State = 1) Then
        pValidated = True
        MsgBox "Login Succesful"
    End If
    GoTo ExitHere

ConnectionError:
    MsgBox "Connection could not open as expected check username and password"
    Err.Clear
    GoTo ExitHere
ExitHere:
    Close_Session
End Sub

Private Sub Execute_Command(strSQL As String)
    'Create the Command Object to Be Utilized
    Dim pCommand As ADODB.Command
    Set pCommand = New ADODB.Command
    pCommand.ActiveConnection = pConnection
    pCommand.CommandType = adCmdText
    pCommand.CommandText = strSQL
    pCommand.Execute
    Set pCommand = Nothing
End Sub

Private Sub EckerdErrorHandler()
    MsgBox "Error" & " - " & Err.Description
    Err.Clear
    pValidCommand = False
End Sub

Private Function ParsedHeaderData(Cell As Range) As String
    Dim HeaderString As String
    If (IsDate(Cell) = True) Then
        Err.Raise 1100, "ECSession.ParsedHeaderData", "Header Cannot Be In Date Format"
    ElseIf (IsNumeric(Cell) = True) Then
        Err.Raise 1101, "ECSession.ParsedHeaderData", "Header Cannot Be In Numeric Format"
    Else
        HeaderString = Trim(Cell)
        If (InStr(HeaderString, "--") > 0 Or InStr(HeaderString, "'") > 0) Then
            Err.Raise 1102, "ECSession.ParsedHeaderData", "Header Cannot Contain SQL Comment Strings or Single Quotes"
        End If
    End If
    If HeaderString = "" Then
        Err.Raise 1103, "ECSession.ParsedHeaderData", "Header Cannot Have Blank Values"
    End If
    ParsedHeaderData = HeaderString
End Function

Private Function ParsedCellData(Cell As Range, Optional PK As Boolean = False) As String

    ' Error Handling For SQL Comments and Single Quotes
    If InStr(1, Cell.Text, "--") <> 0 Then
        Err.Raise 1000, "ECSession.ParsedCellData", "Cannot Insert SQL Comment Strings '--'"
    ElseIf InStr(1, Cell.Text, "'") <> 0 Then
        Err.Raise 1000, "ECSession.ParsedCellData", "Cannot Insert Single Quotes Into Database "
        
    ElseIf IsEmpty(Cell) Then
        ' Cell is recognized as empty
        If PK Then
            Err.Raise 1001, "ECSession.ParsedCellData", "Cannot Have Blank Primary Keys" & _
            vbNewLine & vbNewLine & "- Check Leftmost Columns For Blanks -"
        Else
            ParsedCellData = "null"
        End If
        
    ElseIf WorksheetFunction.IsText(Cell) Then
        ' Cell is recognized as Text
        ' Numbers need to be explicitly forced to Text Type
        ParsedCellData = "'" & Replace(Trim(Cell), "'", "''") & "'"
        
    ElseIf WorksheetFunction.IsLogical(Cell) Then
        ' Cell is True/False
        
    ElseIf WorksheetFunction.IsErr(Cell) Then
        ' Cell is Currently An Error
    
    ElseIf IsDate(Cell) Then
        ' Cell is Date Format
        ParsedCellData = "'" & Format(Cell, "dd-mmm-yyyy") & "'"
        
    ElseIf InStr(1, Cell.Text, ":") <> 0 Then
        ' Cell is Time Format
        
    ElseIf IsNumeric(Cell) Then
        ' Cell is recognized as a number
        ' If you have not forced the type to Text On a Number This Can Cause Errors
        ParsedCellData = Cell
        
    Else
        ParsedCellData = "'" & Replace(Trim(Cell), "'", "''") & "'"
    End If
    
    If ParsedCellData = "" Then
        If PK = True Then
            Err.Raise 1001, "ECSession.ParsedCellData", "Cannot Have Blank Primary Key" & _
            vbNewLine & vbNewLine & "- Check Leftmost Columns For Blanks -"
        Else
            ParsedCellData = "null"
        End If
    End If
    
End Function

Private Function ParsedLoginCredentials(Value As String) As String
    If Value = "" Then
        Err.Raise 1200, "ECSession.ParsedLoginCredentials", "Login Credentials Cannot Be Empty"
    ElseIf InStr(Value, "--") > 0 Then
        Err.Raise 1201, "ECSession.ParsedLoginCredentials", "Login Credentials Cannot Contain SQL Comment Strings"
    ElseIf InStr(Value, "'") > 0 Then
        Err.Raise 1202, "ECSession.ParsedLoginCredentials", "Login Credentials Cannot Contain SQL String Indicators"
    Else
        ParsedLoginCredentials = Value
    End If
End Function













