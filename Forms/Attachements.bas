Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdOpenLink_Click()
 
    Dim url As String
    url = Trim(Nz(Me.directLink, ""))
 
    If Len(url) = 0 Then
        MsgBox "No link found.", vbInformation
        Exit Sub
    End If
 
    'Force Windows to open in default browser (Edge/Chrome/etc.)
    CreateObject("WScript.Shell").Run _
        "cmd /c start """" """ & url & """", 0, False
 
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

'THEME
Dim db As Database
Set db = CurrentDb()
Dim rsUserSettings As Recordset
Dim rsTheme As Recordset

Set rsUserSettings = db.OpenRecordset("tblUserSettings")
rsUserSettings.Filter = "[Username] = '" & Environ("username") & "'"
Set rsUserSettings = rsUserSettings.OpenRecordset

If Nz(rsUserSettings!themeId, 0) <> 0 Then
    Set rsTheme = db.OpenRecordset("SELECT * FROM tblTheme WHERE recordId = " & rsUserSettings!themeId)
    
    If rsTheme!darkMode Then
        TempVars.Add "themeMode", "Dark"
    Else
        TempVars.Add "themeMode", "Light"
    End If
    
    TempVars.Add "themePrimary", CStr(rsTheme!primaryColor)
    TempVars.Add "themeSecondary", CStr(rsTheme!secondaryColor)
    TempVars.Add "themeColorLevels", CStr(rsTheme!colorLevels)
    
    rsTheme.Close
    Set rsTheme = Nothing
End If

Call setTheme(Me)
'If CommandBars("Ribbon").Height > 100 Then CommandBars.ExecuteMso "MinimizeRibbon"
'DoCmd.ShowToolbar "Ribbon", acToolbarNo
'Call DoCmd.NavigateTo("acNavigationCategoryObjectType")
'Call DoCmd.RunCommand(acCmdWindowHide)

On Error Resume Next
rsUserSettings.Close: Set rsUserSettings = Nothing
rsTheme.Close: Set rsTheme = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.Description, Err.Numbe)
End Sub
