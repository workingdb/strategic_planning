Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub deletebtn_Click()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then
    MsgBox "This is an empty record.", vbInformation, "Can't do that"
    Exit Sub
End If

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") = vbYes Then
    Dim oldIndex
    oldIndex = Me.indexOrder
    If oldIndex < DMax("indexOrder", "tblBuildout_tasks_template", "gateTemplateId = " & Me.gateTemplateId) Then
        dbExecute "UPDATE tblBuildout_tasks_template SET indexOrder = indexOrder - 1 WHERE gateTemplateId = " & Me.gateTemplateId & " AND indexOrder > " & oldIndex
    End If
    
    Call registerStratPlanUpdates("tblBuildout_tasks_template", Me.recordId, "DELETE", Me.taskTitle, "DELETED", Form_frmBuildout_template.recordId, "frmBuildout_template")
    dbExecute "DELETE FROM tblBuildout_tasks_template WHERE recordId = " & Me.recordId
    Me.Requery
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
On Error GoTo Err_Handler

Me.indexOrder = Nz(DMax("indexOrder", "tblBuildout_tasks_template", "gateTemplateId = " & Me.gateTemplateId) + 1, 1)
Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub Form_Current()
On Error GoTo Err_Handler

Me.txtCF = Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Current", err.Description, err.Number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.OrderBy = "indexOrder Asc"
Me.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", err.Description, err.Number)
End Sub

Private Sub moveDown_Click()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then Exit Sub

If Me.indexOrder = DMax("indexOrder", "tblBuildout_tasks_template", "gateTemplateId = " & Me.gateTemplateId) Then Exit Sub

Dim oldIndex, newIndex
oldIndex = Me.indexOrder
newIndex = oldIndex + 1

dbExecute "UPDATE tblBuildout_tasks_template SET indexOrder = " & oldIndex & " WHERE gateTemplateId = " & Me.gateTemplateId & " AND indexOrder = " & newIndex
Me.indexOrder = newIndex
Me.Dirty = False
    
Me.Requery
Me.OrderBy = "indexOrder Asc"
Me.OrderByOn = True

Call registerStratPlanUpdates("tblBuildout_tasks_template", Me.recordId, "indexOrder", oldIndex, newIndex, Form_frmBuildout_template.recordId, "frmBuildout_template")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub moveUp_Click()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then Exit Sub
If Me.indexOrder = 1 Then Exit Sub
Dim oldIndex, newIndex
oldIndex = Me.indexOrder
newIndex = oldIndex - 1

dbExecute "UPDATE tblBuildout_tasks_template SET indexOrder = " & oldIndex & " WHERE gateTemplateId = " & Me.gateTemplateId & " AND indexOrder = " & newIndex
Me.indexOrder = newIndex
Me.Dirty = False
    
Me.Requery
Me.OrderBy = "indexOrder Asc"
Me.OrderByOn = True

Call registerStratPlanUpdates("tblBuildout_tasks_template", Me.recordId, "indexOrder", oldIndex, newIndex, Form_frmBuildout_template.recordId, "frmBuildout_template")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub newStep_Click()
On Error GoTo Err_Handler

dbExecute "INSERT INTO tblBuildout_tasks_template(gateTemplateId,indexOrder) VALUES (" & Form_sfrmBuildout_template_gates.recordId & "," & _
    Nz(DMax("indexOrder", "tblBuildout_tasks_template", "gateTemplateId = " & Form_sfrmBuildout_template_gates.recordId) + 1, 1) & ")"
    
Me.Requery

Call registerStratPlanUpdates("tblBuildout_tasks_template", Me.recordId, "New", "", "New Record", Form_frmBuildout_template.recordId, "frmBuildout_template")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub pillarTask_AfterUpdate()
On Error GoTo Err_Handler

Call registerStratPlanUpdates("tblBuildout_tasks_template", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmBuildout_template.recordId, "frmBuildout_template")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub responsibleDept_AfterUpdate()
On Error GoTo Err_Handler

Call registerStratPlanUpdates("tblBuildout_tasks_template", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmBuildout_template.recordId, "frmBuildout_template")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub taskTitle_AfterUpdate()
On Error GoTo Err_Handler

Call registerStratPlanUpdates("tblBuildout_tasks_template", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmBuildout_template.recordId, "frmBuildout_template")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
