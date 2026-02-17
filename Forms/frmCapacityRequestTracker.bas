Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Function GetRequestorEmail_(ByVal requestorId As Long) As String
    On Error GoTo ErrHandler
 
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
 
    sql = "SELECT Email FROM tblPermissions WHERE ID=" & requestorId & ";"
 
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
 
    If Not (rs.EOF And rs.BOF) Then
        GetRequestorEmail_ = Nz(rs!Email, "")
    Else
        GetRequestorEmail_ = ""
    End If
 
Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
 
ErrHandler:
    GetRequestorEmail_ = ""
    Resume Cleanup
End Function

Private Sub Capacity_Results_AfterUpdate()
    On Error GoTo ErrHandler
 
    ' Force save so table has the new value
    If Me.Dirty Then Me.Dirty = False
 
    ' Call the shared notifier
    Call NotifyCapacityResultIfNeeded(CLng(Me.RecordID))
 
    Exit Sub
 
ErrHandler:
    MsgBox "Capacity_Results_AfterUpdate error: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub



Private Sub Capacity_Results_Label_Click()
On Error GoTo Err_Handler

Me.Capacity_Results.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub cmdOpenAttachments_Click()
 
    DoCmd.OpenForm "fsubStratPlanAttachments", _
        , , "referenceTable='tblCapacityRequests' AND referenceId=" & Me.RecordID
 
End Sub

Private Sub Customer_Label_Click()
On Error GoTo Err_Handler

Me.Customer.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub EOP_Label_Click()
On Error GoTo Err_Handler

Me.EOP.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub




Private Sub NAM_Label_Click()
On Error GoTo Err_Handler

Me.NAM.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub newRequest_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmCapacityRequestDetails", , , , acFormAdd

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub openDetails_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmCapacityRequestDetails", , , "RecordID = " & Me.RecordID

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub Planner_Label_Click()
On Error GoTo Err_Handler

Me.Planner.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub PPV_Label_Click()
On Error GoTo Err_Handler

Me.PPV.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub Production_Type_Label_Click()
On Error GoTo Err_Handler

Me.Production_Type.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub Program_Label_Click()
On Error GoTo Err_Handler

Me.Program.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub Quote_Label_Click()
On Error GoTo Err_Handler

Me.Quote.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub RecordID_Label_Click()
On Error GoTo Err_Handler

Me.RecordID.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub Request_Date_Label_Click()
On Error GoTo Err_Handler

Me.Request_Date.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub Request_Type_Label_Click()
On Error GoTo Err_Handler

Me.Request_Type.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub Requestor_Label_Click()
On Error GoTo Err_Handler

Me.Requestor.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub Response_Date_Label_Click()
On Error GoTo Err_Handler

Me.Response_Date.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub SOP_Label_Click()
On Error GoTo Err_Handler

Me.SOP.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub Unit_Label_Click()
On Error GoTo Err_Handler

Me.Unit.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub Vehicle_Model_Label_Click()
On Error GoTo Err_Handler

Me.Vehicle_Model.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub Volume_Label_Click()
On Error GoTo Err_Handler

Me.Volume.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub Volume_Timing_Label_Click()
On Error GoTo Err_Handler

Me.Volume_Timing.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub Volume_Type_Label_Click()
On Error GoTo Err_Handler

Me.Volume_Type.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub
