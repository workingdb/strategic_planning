Option Compare Database
Option Explicit

Function userData(data, Optional specificUser As String = "") As String
On Error GoTo Err_Handler

If specificUser = "" Then specificUser = Environ("username")

userData = Nz(DLookup("[" & data & "]", "[tblPermissions]", "[User] = '" & specificUser & "'"))

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "replaceDriveLetters", Err.Description, Err.Number)
End Function

Function dbExecute(sql As String)
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

db.Execute sql

Set db = Nothing

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "dbExecute", Err.Description, Err.Number, sql)
End Function