Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnCustomers_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmCustomers", acNormal, "", "", , acNormal

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub btnProductionType_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmProductionType", acNormal, "", "", , acNormal

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub btnQuoteAward_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmQuoteAward", acNormal, "", "", , acNormal

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub btnRequestType_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmRequestType", acNormal, "", "", , acNormal

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub btnResults_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmResults", acNormal, "", "", , acNormal

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub btnVolumeType_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmVolumeType", acNormal, "", "", , acNormal

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.Description, Err.Number)
End Sub
