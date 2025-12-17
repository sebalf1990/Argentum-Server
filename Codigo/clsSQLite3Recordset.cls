VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQLite3Recordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================================
' clsSQLite3Recordset - SQLite3 Recordset Class
' ============================================================================
' Replaces the SQLite3Recordset Type to allow passing as parameters
' ============================================================================

Option Explicit

' ============================================================================
' Private storage
' ============================================================================
Private Type RowData
    Values() As Variant
End Type

Private m_ColumnCount As Long
Private m_ColumnNames() As String
Private m_ColumnTypes() As Long
Private m_Rows() As RowData
Private m_RowCount As Long
Private m_CurrentRow As Long
Private m_BOF As Boolean
Private m_EOF As Boolean
Private m_DataSet As Object  ' VBSQLite12.SQLiteDataSet

' ============================================================================
' DataSet Property - Wraps VBSQLite12.SQLiteDataSet
' ============================================================================
Public Property Get DataSet() As Object
    Set DataSet = m_DataSet
End Property

Public Property Set DataSet(ByVal ds As Object)
    Set m_DataSet = ds
    ' Update EOF based on DataSet
    If Not ds Is Nothing Then
        m_EOF = ds.EOF
        If Not m_EOF Then
            m_CurrentRow = 0
            m_RowCount = 1  ' Approximate
            m_BOF = False
        End If
    End If
End Property

' Access field by name using DataSet
Public Property Get Item(ByVal fieldName As String) As Variant
    On Error Resume Next
    If Not m_DataSet Is Nothing Then
        Item = m_DataSet(fieldName)
    Else
        Item = Field(fieldName)
    End If
End Property

' ============================================================================
' Properties
' ============================================================================
Public Property Get ColumnCount() As Long
    ColumnCount = m_ColumnCount
End Property

Public Property Let ColumnCount(ByVal value As Long)
    m_ColumnCount = value
End Property

Public Property Get RowCount() As Long
    RowCount = m_RowCount
End Property

Public Property Let RowCount(ByVal value As Long)
    m_RowCount = value
End Property

Public Property Get CurrentRow() As Long
    CurrentRow = m_CurrentRow
End Property

Public Property Let CurrentRow(ByVal value As Long)
    m_CurrentRow = value
End Property

Public Property Get BOF() As Boolean
    BOF = m_BOF
End Property

Public Property Let BOF(ByVal value As Boolean)
    m_BOF = value
End Property

Public Property Get EOF() As Boolean
    EOF = m_EOF
End Property

Public Property Let EOF(ByVal value As Boolean)
    m_EOF = value
End Property

' ============================================================================
' Initialize
' ============================================================================
Private Sub Class_Initialize()
    m_RowCount = 0
    m_CurrentRow = -1
    m_BOF = True
    m_EOF = True
    m_ColumnCount = 0
End Sub

' ============================================================================
' Column Name/Type Management
' ============================================================================
Public Sub InitColumns(ByVal count As Long)
    m_ColumnCount = count
    If count > 0 Then
        ReDim m_ColumnNames(0 To count - 1)
        ReDim m_ColumnTypes(0 To count - 1)
    End If
End Sub

Public Sub SetColumnName(ByVal idx As Long, ByVal name As String)
    If idx >= 0 And idx < m_ColumnCount Then
        m_ColumnNames(idx) = name
    End If
End Sub

Public Function GetColumnName(ByVal idx As Long) As String
    If idx >= 0 And idx < m_ColumnCount Then
        GetColumnName = m_ColumnNames(idx)
    End If
End Function

Public Sub SetColumnType(ByVal idx As Long, ByVal colType As Long)
    If idx >= 0 And idx < m_ColumnCount Then
        m_ColumnTypes(idx) = colType
    End If
End Sub

Public Function GetColumnType(ByVal idx As Long) As Long
    If idx >= 0 And idx < m_ColumnCount Then
        GetColumnType = m_ColumnTypes(idx)
    End If
End Function

' ============================================================================
' Row Management
' ============================================================================
Public Sub InitRows(ByVal capacity As Long)
    If capacity > 0 Then
        ReDim m_Rows(0 To capacity - 1)
    End If
End Sub

Public Sub ExpandRows(ByVal newCapacity As Long)
    ReDim Preserve m_Rows(0 To newCapacity - 1)
End Sub

Public Sub SetRowValues(ByVal rowIndex As Long, ByRef values() As Variant)
    m_Rows(rowIndex).Values = values
End Sub

Public Function GetRowValue(ByVal rowIndex As Long, ByVal colIndex As Long) As Variant
    If rowIndex >= 0 And rowIndex < m_RowCount Then
        If colIndex >= 0 And colIndex < m_ColumnCount Then
            GetRowValue = m_Rows(rowIndex).Values(colIndex)
        Else
            GetRowValue = Null
        End If
    Else
        GetRowValue = Null
    End If
End Function

' ============================================================================
' Field Access
' ============================================================================
Public Function Field(ByVal fieldName As String) As Variant
    Dim i As Long
    For i = 0 To m_ColumnCount - 1
        If UCase$(m_ColumnNames(i)) = UCase$(fieldName) Then
            If m_CurrentRow >= 0 And m_CurrentRow < m_RowCount Then
                Field = m_Rows(m_CurrentRow).Values(i)
            Else
                Field = Null
            End If
            Exit Function
        End If
    Next i
    Field = Null
End Function

Public Function FieldByIndex(ByVal idx As Long) As Variant
    If idx >= 0 And idx < m_ColumnCount Then
        If m_CurrentRow >= 0 And m_CurrentRow < m_RowCount Then
            FieldByIndex = m_Rows(m_CurrentRow).Values(idx)
        Else
            FieldByIndex = Null
        End If
    Else
        FieldByIndex = Null
    End If
End Function

' ============================================================================
' Navigation
' ============================================================================

' ============================================================================
' Fields - ADODB Compatibility Layer
' ============================================================================

' Access field by index (for RS.Fields(0) compatibility)
Public Function Fields(ByVal indexOrName As Variant) As Variant
    On Error Resume Next
    If Not m_DataSet Is Nothing Then
        If VarType(indexOrName) = vbString Then
            Fields = m_DataSet(CStr(indexOrName))
        Else
            ' Numeric index - get field by position
            Fields = m_DataSet.FieldValue(CLng(indexOrName))
        End If
    Else
        ' Fallback to internal storage
        If VarType(indexOrName) = vbString Then
            Fields = Field(CStr(indexOrName))
        Else
            Fields = FieldByIndex(CLng(indexOrName))
        End If
    End If
End Function

Public Function MoveNext() As Boolean
    ' If using DataSet, delegate to it
    If Not m_DataSet Is Nothing Then
        m_DataSet.MoveNext
        m_EOF = m_DataSet.EOF
        MoveNext = Not m_EOF
        Exit Function
    End If
    ' Original implementation
    If m_CurrentRow < m_RowCount - 1 Then
        m_CurrentRow = m_CurrentRow + 1
        m_BOF = False
        m_EOF = False
        MoveNext = True
    Else
        m_EOF = True
        MoveNext = False
    End If
End Function

Public Function MoveFirst() As Boolean
    If m_RowCount > 0 Then
        m_CurrentRow = 0
        m_BOF = False
        m_EOF = False
        MoveFirst = True
    Else
        MoveFirst = False
    End If
End Function

' ============================================================================
' Utility
' ============================================================================
Public Function RecordCount() As Long
    RecordCount = m_RowCount
End Function





