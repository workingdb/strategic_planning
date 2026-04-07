Option Compare Database
Option Explicit

Function GetBestSQLDriver() As String
On Error GoTo Err_Handler

    Dim shell As Object
    Dim driverList As Variant, driver As Variant
    Dim regPath As String
    
    Set shell = CreateObject("WScript.Shell")
    ' List drivers in order of preference (newest first)
    driverList = Array("ODBC Driver 18 for SQL Server", _
                       "ODBC Driver 17 for SQL Server", _
                       "ODBC Driver 13 for SQL Server", _
                       "SQL Server Native Client 11.0", _
                       "SQL Server")
    
    On Error Resume Next
    For Each driver In driverList
        ' Check registry for the driver entry
        If RegKeyExists("HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\" & driver & "\Driver") Then
            GetBestSQLDriver = driver
            Exit Function
        End If
    Next driver
    On Error GoTo Err_Handler
    
    GetBestSQLDriver = "" ' No driver found
    
Exit Function
Err_Handler:
    Call handleError("wdbLinkage", "GetBestSQLDriver", err.Description, err.Number)
End Function

Function RegKeyExists(regPath As String) As Boolean
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    On Error Resume Next
    shell.RegRead regPath
    RegKeyExists = (err.Number = 0)
    On Error GoTo 0
End Function

Sub RelinkSQLTables()
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim qdf As DAO.QueryDef
    Dim strDriver As String
    Dim strConn As String
    
    strDriver = GetBestSQLDriver()
    If strDriver = "" Then Exit Sub
    
    Set db = CurrentDb
    ' Construct your base connection string (DSN-less)
    strConn = "ODBC;DRIVER=" & strDriver & ";SERVER=ITI-SQL\ITISQL;Trusted_Connection=Yes;APP=Microsoft Office;DATABASE=workingdb"
    
    ' Loop through all tables and update ODBC links
    For Each tdf In db.TableDefs
        ' Only relink tables that already have an ODBC connection string
        If InStr(1, tdf.Connect, "SERVER=ITI-SQL") Then
            tdf.Connect = strConn
            tdf.RefreshLink
        End If
    Next tdf
    
    For Each qdf In db.QueryDefs
    If qdf.Type = dbQSQLPassThrough Then
        If InStr(1, qdf.Connect, "SERVER=ITI-SQL", vbTextCompare) > 0 Then
            qdf.Connect = strConn
        End If
    End If
Next qdf
    
Exit Sub
Err_Handler:
    Call handleError("wdbLinkage", "RelinkSQLTables", err.Description, err.Number)
End Sub