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
templateId = DMin("recordId", "tblBuildout_gates_template", "templateId = " & Me.RecordID)

On Error GoTo invis
Me.sfrmBuildout_template_tasks.Form.filter = "gateTemplateId = " & templateId
Me.sfrmBuildout_template_tasks.Form.gateTemplateId.defaultValue = templateId
Me.sfrmBuildout_template_tasks.Form.FilterOn = True

Exit Sub
invis:
Me.sfrmBuildout_template_tasks.Visible = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.Description, Err.Number)
End Sub

Private Sub history_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", , , "referenceId = " & Me.RecordID

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub refresh_Click()
On Error GoTo Err_Handler

Me.Requery
Me.sfrmBuildout_template_tasks.Requery
Me.sfrmBuildout_template_gates.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub templateName_AfterUpdate()
On Error GoTo Err_Handler

Call registerStratPlanUpdates("tblBuildout_tasks_template", Me.RecordID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmBuildout_template.RecordID, "frmBuildout_template")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub
