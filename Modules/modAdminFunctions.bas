Option Compare Database
Option Explicit

Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_RESTORE = 9

Private Type RECT
x1 As Long
y1 As Long
x2 As Long
y2 As Long
End Type

Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, r As RECT) As Long
Public Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function moveWindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As Long, ByVal x As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal fRepaint As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Dim AppX As Long, AppY As Long, AppTop As Long, AppLeft As Long, WinRECT As RECT

Sub maximizeAccess()
On Error GoTo Err_Handler

Dim h As Long
Dim r As RECT

On Error Resume Next

h = Application.hWndAccessApp
'If maximised, restore
If (IsZoomed(h) = False) Then ShowWindow h, SW_SHOWMAXIMIZED

Exit Sub
Err_Handler:
    Call handleError("modAdminFunctions", "maximizeAccess", err.Description, err.Number)
End Sub

Public Sub handleError(modName As String, activeCon As String, errDesc As String, errNum As Long, Optional dataTag As String = "")
On Error Resume Next

If (CurrentProject.Path <> "C:\workingdb") Then
    MsgBox errDesc, vbInformation, "Error Code: " & errNum
    Exit Sub
End If

Select Case errNum
    Case 70
        MsgBox "Permissions Error - Check if the file is already in use.", vbInformation, "Error Code: " & errNum
    Case 53
        MsgBox "File Not Found", vbInformation, "Error Code: " & errNum
        Exit Sub
    Case 3011
        MsgBox "Looks like I'm having issues connecting to SharePoint. Please reopen when you can", vbInformation, "Error Code: " & errNum
    Case 490, 52, 75
        MsgBox "I cannot open this file or location - check if it has been moved or deleted. Or - you do not have proper access to this location", vbInformation, "Error Code: " & errNum
        Exit Sub
    Case 3022
        MsgBox "A record with this key already exists. I cannot create another!", vbInformation, "Error Code: " & errNum
    Case 3167
        MsgBox "Looks like you already deleted that record", vbInformation, "Error Code: " & errNum
        Exit Sub
    Case 94
        MsgBox "Hmm. Looks like something is missing. Check for an empty field", vbInformation, "Error Code: " & errNum
    Case 3151
        MsgBox "You're not connected to Oracle. Just FYI, Oracle connection does not work outside of VMWare.", vbInformation, "Error Code: " & errNum
        Exit Sub
    Case 429
        If modName = "frmCatiaMacros" Then
            MsgBox "Looks like Catia isn't open", vbInformation, "Error Code: " & errNum
            Exit Sub
        Else
            MsgBox errDesc, vbInformation, "Error Code: " & errNum
        End If
    Case 3343
        MsgBox "Error. Please re-open WorkingDB to reset.", vbCritical, "Error Code: " & errNum
    Case Else
        MsgBox errDesc, vbInformation, "Error Code: " & errNum
End Select

Dim strSQL As String

modName = Replace(Nz(modName, ""), "'", "''")
errDesc = Replace(Nz(errDesc, ""), "'", "''")
errNum = Replace(Nz(errNum, ""), "'", "''")
dataTag = Replace(Nz(dataTag, ""), "'", "''")

strSQL = "INSERT INTO tblErrorLog([User],Form,Active_Control,Error_Date,Error_Description,Error_Number,databaseVersion,dataTag0) VALUES ('" & _
 Environ("username") & "','" & modName & "','" & Nz(activeCon, "") & "',#" & Now & "#,'" & errDesc & "'," & errNum & ",'SP:" & Nz(TempVars!dbVersion, "") & "','" & dataTag & "')"

Dim conn As ADODB.Connection
Set conn = CurrentProject.Connection

conn.Execute strSQL

Set conn = Nothing

End Sub

Sub SizeAccess(ByVal dx As Long, ByVal dy As Long)
On Error GoTo Err_Handler
'Set size of Access and center on Desktop

Dim h As Long
Dim r As RECT

On Error Resume Next

h = Application.hWndAccessApp
'If maximised, restore
If (IsZoomed(h)) Then ShowWindow h, SW_RESTORE
'
'Get available Desktop size
GetWindowRect GetDesktopWindow(), r
If ((r.x2 - r.x1) - dx) < 0 Or ((r.y2 - r.y1) - dy) < 0 Then
'Desktop smaller than requested size
'so size to Desktop
moveWindow h, r.x1, r.y1, r.x2, r.y2, True
Else
'Adjust to requested size and center
moveWindow h, _
r.x1 + ((r.x2 - r.x1) - dx) \ 2, _
r.y1 + ((r.y2 - r.y1) - dy) \ 2, _
dx, dy, True
End If

Exit Sub
Err_Handler:
    Call handleError("modAdminFunctions", "SizeAccess", err.Description, err.Number)
End Sub

Function grabVersion() As String
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT releaseVal FROM tblDBinfo WHERE recordId = 1", dbOpenSnapshot)
grabVersion = rs1!releaseVal
rs1.CLOSE: Set rs1 = Nothing
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError("modAdminFunctions", "grabVersion", err.Description, err.Number)
End Function