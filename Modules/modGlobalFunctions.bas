Option Compare Database
Option Explicit

Function userData(data, Optional specificUser As String = "") As String
On Error GoTo Err_Handler

If specificUser = "" Then specificUser = Environ("username")

userData = Nz(DLookup("[" & data & "]", "[tblPermissions]", "[User] = '" & specificUser & "'"))

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "replaceDriveLetters", Err.DESCRIPTION, Err.number)
End Function

Function convertRecords()

Dim db As Database
Dim rs As Recordset

Set db = CurrentDb
Set rs = db.OpenRecordset("Capacity request Tracker")

Dim tempString As String
Dim newID As Long

Do While Not rs.EOF
    
    tempString = ""
    newID = 0
    
    If Nz(rs!Requestor, 0) = 0 Then GoTo nextrecord
    
    newID = CLng(DLookup("[Account Manager]", "Account Managers", "ID = " & rs!Requestor))
    
    rs.Edit
    rs!Requestor = newID
    rs.Update
    
nextrecord:
    rs.MoveNext
Loop

End Function