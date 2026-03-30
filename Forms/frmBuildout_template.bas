Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim templateId
templateId = DMin("recordId", "tblBuildout_gates_template", "templateId = " & Me.recordId)

On Error GoTo invis
Me.sfrmBuildout_template_tasks.Form.Filter = "gateTemplateId = " & templateId
Me.sfrmBuildout_template_tasks.Form.gateTemplateId.defaultValue = templateId
Me.sfrmBuildout_template_tasks.Form.FilterOn = True

Exit Sub
invis:
Me.sfrmBuildout_template_tasks.Visible = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", err.Description, err.Number)
End Sub

Private Sub history_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", , , "referenceId = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub refresh_Click()
On Error GoTo Err_Handler

Me.Requery
Me.sfrmBuildout_template_tasks.Requery
Me.sfrmBuildout_template_gates.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub templateName_AfterUpdate()
On Error GoTo Err_Handler

Call registerStratPlanUpdates("tblBuildout_tasks_template", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmBuildout_template.recordId, "frmBuildout_template")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
