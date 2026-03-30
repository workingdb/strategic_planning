Option Compare Database
Option Explicit

Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal lpnShowCmd As Long) As Long

Public Sub openPath(Path)
On Error GoTo Err_Handler

CreateObject("Shell.Application").Open CVar(Path)

Exit Sub
Err_Handler:
    Call handleError("modDirectoryFunctions", "openPath", err.Description, err.Number)
End Sub

Function replaceDriveLetters(linkInput) As String
On Error GoTo Err_Handler

replaceDriveLetters = linkInput

replaceDriveLetters = Replace(replaceDriveLetters, "N:\", "\\ncm-fs2\data\Department\")
replaceDriveLetters = Replace(replaceDriveLetters, "T:\", "\\design\data\")
replaceDriveLetters = Replace(replaceDriveLetters, "S:\", "\\nas01\allshare\")

Exit Function
Err_Handler:
    Call handleError("modDirectoryFunctions", "replaceDriveLetters", err.Description, err.Number)
End Function

Function addLastSlash(linkString As String) As String
On Error GoTo Err_Handler

addLastSlash = linkString
If Right(addLastSlash, 1) <> "\" Then addLastSlash = addLastSlash & "\"

Exit Function
Err_Handler:
    Call handleError("modDirectoryFunctions", "addLastSlash", err.Description, err.Number)
End Function

Function FolderExists(sFile As Variant) As Boolean
On Error GoTo Err_Handler

FolderExists = False
If IsNull(sFile) Then Exit Function
If Dir(sFile, vbDirectory) <> "" Then FolderExists = True

Exit Function
Err_Handler:
    If err.Number = 52 Then Exit Function
    Call handleError("modDirectoryFunctions", "FolderExists", err.Description, err.Number)
End Function