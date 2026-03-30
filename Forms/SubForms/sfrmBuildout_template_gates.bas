Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()
On Error GoTo Err_Handler

Me.txtCF = Me.recordId

If IsNull(Me.recordId) = False Then
    Form_sfrmBuildout_template_tasks.Visible = True
    Form_sfrmBuildout_template_tasks.Filter = "gateTemplateId = " & Me.recordId
    Form_sfrmBuildout_template_tasks.gateTemplateId.defaultValue = Me.recordId
    Form_sfrmBuildout_template_tasks.FilterOn = True
Else
    Form_sfrmBuildout_template_tasks.Visible = False
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Current", err.Description, err.Number)
End Sub

Private Sub gateDuration_AfterUpdate()
On Error GoTo Err_Handler

Call registerStratPlanUpdates("tblBuildout_tasks_template", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmBuildout_template.recordId, "frmBuildout_template")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub gateTitle_AfterUpdate()
On Error GoTo Err_Handler

Call registerStratPlanUpdates("tblBuildout_tasks_template", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmBuildout_template.recordId, "frmBuildout_template")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
