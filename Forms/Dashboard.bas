Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub allRequests_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmCapacityRequestTracker"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnOpenReportLauncher_Click()
    DoCmd.OpenForm "frmReportLauncher", acNormal
End Sub

Private Sub btnSettings_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserView"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cmdOpenSalesManagerReport_Click()

    DoCmd.OpenReport "rpt_SalesManager", acViewReport

End Sub

Private Sub Form_Load()
'On Error GoTo Err_Handler

'THEME
Dim db As Database
Set db = CurrentDb()
Dim rsUserSettings As Recordset
Dim rsTheme As Recordset

Set rsUserSettings = db.OpenRecordset("tblUserSettings")
rsUserSettings.Filter = "[User] = '" & Environ("username") & "'"
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
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.Numbe)
End Sub

Private Sub newRequest_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmCapacityRequestDetails", , , , acFormAdd

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
