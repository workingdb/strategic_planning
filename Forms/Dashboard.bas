Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub allRequests_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmCapacityRequestTracker"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub btnMaintenance_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmMaintanence", acNormal, "", "", , acNormal

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub btnOpenReportLauncher_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmReportLauncher", acNormal

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub btnSettings_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserView"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub cmdOpenSalesManagerReport_Click()
On Error GoTo Err_Handler

DoCmd.OpenReport "rpt_SalesManager", acViewReport

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub btnSurvery_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmSurveys", acNormal, "", "", , acNormal
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", err.Description, err.Numbe)
End Sub
