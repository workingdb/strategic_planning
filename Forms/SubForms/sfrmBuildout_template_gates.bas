Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()
On Error GoTo Err_Handler

Me.txtCF = Me.RecordID

If IsNull(Me.RecordID) = False Then
    Form_sfrmBuildout_template_tasks.Visible = True
    Form_sfrmBuildout_template_tasks.filter = "gateTemplateId = " & Me.RecordID
    Form_sfrmBuildout_template_tasks.gateTemplateId.defaultValue = Me.RecordID
    Form_sfrmBuildout_template_tasks.FilterOn = True
Else
    Form_sfrmBuildout_template_tasks.Visible = False
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Current", Err.Description, Err.Number)
End Sub

Private Sub gateDuration_AfterUpdate()
On Error GoTo Err_Handler

Call registerStratPlanUpdates("tblBuildout_tasks_template", Me.RecordID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmBuildout_template.RecordID, "frmBuildout_template")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub gateTitle_AfterUpdate()
On Error GoTo Err_Handler

Call registerStratPlanUpdates("tblBuildout_tasks_template", Me.RecordID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmBuildout_template.RecordID, "frmBuildout_template")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub
