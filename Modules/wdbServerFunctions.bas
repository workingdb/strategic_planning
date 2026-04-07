Option Compare Database
Option Explicit

Public Function OpenRecordsetReadOnly(conn As ADODB.Connection, sql As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set OpenRecordsetReadOnly = rs
End Function

Public Function OpenRecordsetReadWrite(conn As ADODB.Connection, sql As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenKeyset, adLockOptimistic, adCmdText
    Set OpenRecordsetReadWrite = rs
End Function

Public Function SqlMin(conn As ADODB.Connection, fieldName As String, tableName As String, _
    Optional whereClause As String = "", Optional defaultValue As Variant = Null) As Variant

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT MIN(" & fieldName & ") FROM " & tableName
    If Trim$(whereClause) <> "" Then sql = sql & " WHERE " & whereClause

    Set rs = OpenRecordsetReadOnly(conn, sql)

    If rs.EOF Or IsNull(rs.Fields(0).Value) Then
        SqlMin = defaultValue
    Else
        SqlMin = rs.Fields(0).Value
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function SqlCount(conn As ADODB.Connection, tableName As String, _
    Optional whereClause As String = "") As Long

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT COUNT(*) FROM " & tableName
    If Trim$(whereClause) <> "" Then sql = sql & " WHERE " & whereClause

    Set rs = OpenRecordsetReadOnly(conn, sql)
    SqlCount = CLng(Nz(rs.Fields(0).Value, 0))

    rs.Close
    Set rs = Nothing
End Function

Public Function SqlLookup(conn As ADODB.Connection, fieldName As String, tableName As String, _
    Optional whereClause As String = "", Optional defaultValue As Variant = Null) As Variant

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT " & fieldName & " FROM " & tableName
    If Trim$(whereClause) <> "" Then sql = sql & " WHERE " & whereClause

    Set rs = OpenRecordsetReadOnly(conn, sql)

    If rs.EOF Or IsNull(rs.Fields(0).Value) Then
        SqlLookup = defaultValue
    Else
        SqlLookup = rs.Fields(0).Value
    End If

    rs.Close
    Set rs = Nothing
End Function