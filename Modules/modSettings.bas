Option Compare Database
Option Explicit
 
'Returns setting as string (or defaultValue if missing)
Public Function GetSetting(ByVal settingName As String, Optional ByVal defaultValue As String = "") As String
    On Error GoTo ErrHandler
 
    Dim v As Variant
    v = DLookup("SettingValue", "tblAppSettings", "SettingName='" & Replace(settingName, "'", "''") & "'")
 
    If IsNull(v) Then
        GetSetting = defaultValue
    Else
        GetSetting = CStr(v)
    End If
    Exit Function
 
ErrHandler:
    GetSetting = defaultValue
End Function
 
'Returns True/False from setting
Public Function GetSettingBool(ByVal settingName As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim s As String
    s = Trim$(LCase$(GetSetting(settingName, IIf(defaultValue, "True", "False"))))
 
    GetSettingBool = (s = "true" Or s = "1" Or s = "yes" Or s = "on")
End Function
 
'Optional: enable only for certain users (semicolon list)
Public Function IsUserInBeta(Optional ByVal settingName As String = "BetaUsers") As Boolean
    Dim userList As String, u As String
    userList = LCase$(GetSetting(settingName, ""))
    u = LCase$(Environ$("Username"))
 
    IsUserInBeta = (InStr(1, ";" & userList & ";", ";" & u & ";", vbTextCompare) > 0)
End Function
 