Attribute VB_Name = "modDatabaseCompat"
' ============================================================================
' modDatabaseCompat.bas - VBSQLite12 Adapter Layer (v3)
' ============================================================================

Option Explicit

Private m_Connection As VBSQLite12.SQLiteConnection
Private m_Connected As Boolean

' ============================================================================
' Logging Helper
' ============================================================================
Private Sub LogDB(ByVal msg As String)
    On Error Resume Next
    Dim fNum As Integer
    fNum = FreeFile
    Open App.Path & "\db_debug.log" For Append As #fNum
    Print #fNum, Format$(Now, "hh:mm:ss") & " - " & msg
    Close #fNum
End Sub

' ============================================================================
' Database Connection
' ============================================================================

Public Sub Database_Connect()
    Call LogDB("Database_Connect: called")
    Call Database_Connect_Native
End Sub

Public Sub Database_Connect_Async()
    Call LogDB("Database_Connect_Async: called")
    Call Database_Connect_Native
End Sub

Public Sub Database_Close()
    Call Database_Close_Native
End Sub

Public Sub Database_Connect_Native()
    On Error GoTo ConnectError
    
    Call LogDB("Database_Connect_Native: INICIO")
    
    Dim dbPath As String
    dbPath = App.Path & "/" & DatabaseFileName
    
    Call LogDB("Database_Connect_Native: dbPath = " & dbPath)
    
    If Dir(dbPath) = "" Then
        Call LogDB("Database_Connect_Native: ERROR - Archivo no existe")
        m_Connected = False
        Exit Sub
    End If
    
    Call LogDB("Database_Connect_Native: Creando conexion...")
    Set m_Connection = New VBSQLite12.SQLiteConnection
    
    Call LogDB("Database_Connect_Native: Abriendo BD...")
    m_Connection.OpenDB dbPath
    
    m_Connected = True
    ' Initialize Builder for SaveCharacterDB
    Set Builder = New cStringBuilder
    Call LogDB("Database_Connect_Native: CONECTADO OK - SQLite " & m_Connection.Version)
    Call LogDatabaseError("VBSQLite12: Connected (" & m_Connection.Version & ")")
    Exit Sub
    
ConnectError:
    m_Connected = False
    Call LogDB("Database_Connect_Native: ERROR " & Err.Number & " - " & Err.Description)
    Call LogDatabaseError("VBSQLite12 Error: " & Err.Number & " - " & Err.Description)
End Sub

Public Sub Database_Close_Native()
    On Error Resume Next
    Call LogDB("Database_Close_Native: Cerrando...")
    If m_Connected Then
        m_Connection.CloseDB
        Set m_Connection = Nothing
        m_Connected = False
    End If
End Sub

Public Function Database_IsConnected() As Boolean
    Database_IsConnected = m_Connected
End Function

' ============================================================================
' Query Function - Returns the raw SQLiteDataSet wrapped
' ============================================================================

Public Function Query(ByVal Text As String, ParamArray Arguments() As Variant) As clsSQLite3Recordset
    On Error GoTo QueryError
    
    Dim finalSQL As String
    Dim ds As VBSQLite12.SQLiteDataSet
    Dim rs As clsSQLite3Recordset
    Dim i As Long
    Dim args() As Variant
    
    Set rs = New clsSQLite3Recordset
    
    If Not m_Connected Then
        Call LogDB("Query: ERROR - No conectado")
        Set Query = rs
        Exit Function
    End If
    
    finalSQL = Text
    
    ' Handle case where first argument is an array
    If UBound(Arguments) >= 0 Then
        If IsArray(Arguments(0)) And UBound(Arguments) = 0 Then
            ' First arg is array, use it directly
            args = Arguments(0)
            For i = LBound(args) To UBound(args)
                finalSQL = ReplaceFirst(finalSQL, "?", FormatSQLValue(args(i)))
            Next i
        Else
            ' Individual arguments
            For i = LBound(Arguments) To UBound(Arguments)
                finalSQL = ReplaceFirst(finalSQL, "?", FormatSQLValue(Arguments(i)))
            Next i
        End If
    End If
    
    Call LogDB("Query: " & Left$(finalSQL, 100))
    
    Set ds = m_Connection.OpenDataSet(finalSQL)
    
    ' Store the DataSet directly in the recordset
    If Not ds Is Nothing Then
        Set rs.DataSet = ds
        Call LogDB("Query: OK - DataSet assigned")
    End If
    
    Set Query = rs
    Exit Function
    
QueryError:
    DBError = Err.Description
    Call LogDB("Query ERROR: " & Err.Number & " - " & Err.Description)
    Set Query = New clsSQLite3Recordset
End Function

' ============================================================================
' Execute Function
' ============================================================================

Public Function Execute(ByVal Text As String, ParamArray Arguments() As Variant) As Boolean
    On Error GoTo ExecuteError
    
    Dim finalSQL As String
    Dim i As Long
    Dim args() As Variant
    
    If Not m_Connected Then
        Call LogDB("Execute: ERROR - No conectado")
        Execute = False
        Exit Function
    End If
    
    finalSQL = Text
    
    ' Handle case where first argument is an array
    If UBound(Arguments) >= 0 Then
        If IsArray(Arguments(0)) And UBound(Arguments) = 0 Then
            ' First arg is array, use it directly
            args = Arguments(0)
            For i = LBound(args) To UBound(args)
                finalSQL = ReplaceFirst(finalSQL, "?", FormatSQLValue(args(i)))
            Next i
        Else
            ' Individual arguments
            For i = LBound(Arguments) To UBound(Arguments)
                finalSQL = ReplaceFirst(finalSQL, "?", FormatSQLValue(Arguments(i)))
            Next i
        End If
    End If
    
    Call LogDB("Execute: " & Left$(finalSQL, 100))
    
    m_Connection.Execute finalSQL
    
    Call LogDB("Execute: OK")
    Execute = True
    Exit Function
    
ExecuteError:
    DBError = Err.Description
    Call LogDB("Execute ERROR: " & Err.Number & " - " & Err.Description)
    Execute = False
End Function

' ============================================================================
' Helpers
' ============================================================================

Private Function ReplaceFirst(ByVal Text As String, ByVal Find As String, ByVal ReplaceWith As String) As String
    Dim pos As Long
    pos = InStr(1, Text, Find)
    If pos > 0 Then
        ReplaceFirst = Left$(Text, pos - 1) & ReplaceWith & Mid$(Text, pos + Len(Find))
    Else
        ReplaceFirst = Text
    End If
End Function

Private Function FormatSQLValue(ByVal value As Variant) As String
    If IsNull(value) Then
        FormatSQLValue = "NULL"
    ElseIf VarType(value) = vbString Then
        FormatSQLValue = "'" & Replace(CStr(value), "'", "''") & "'"
    ElseIf VarType(value) = vbBoolean Then
        FormatSQLValue = IIf(value, "1", "0")
    Else
        FormatSQLValue = CStr(value)
    End If
End Function

