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