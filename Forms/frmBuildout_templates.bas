Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub details_Click()
On Error GoTo Err_Handler

Me.Dirty = False
DoCmd.OpenForm "frmBuildout_template", acNormal, , "recordId = " & Me.RecordID

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub projectTitle_AfterUpdate()
On Error GoTo Err_Handler

Call registerStratPlanUpdates("tblBuildout_tasks_template", Me.RecordID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.RecordID, "frmBuildout_templates")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
