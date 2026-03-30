Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
 
Private Sub Capacity_Results_AfterUpdate()
On Error GoTo Err_Handler
 
    'Stamp response date when Capactiy Results has value
    If Nz(Me.capacityResults, "") <> "" And IsNull(Me.responseDate) Then
        Me.responseDate = Date
    End If
    
    'Force save so table has the new value
    If Me.Dirty Then Me.Dirty = False
 
    'popup email for notification to
    Dim emailBody As String, subjectLine As String, strTo As String
    subjectLine = Me.partNumber & " Capacity Request"
    emailBody = generateHTML( _
            subjectLine, _
            "Capacity Result: " & Me.Capacity_Results.column(1), _
            "Regarding Capacity Request: " & Me.Request_Type.column(1) & " for " & Me.partNumber & " on program " & Me.Program, _
            "Notes: " & Me.Notes, _
            "Customer: " & Me.Customer.column(1), _
            "PPV: " & Me.PPV _
            )
    
    strTo = getEmail(Nz(Me.Requestor, ""))
    
    Call wdbEmail(strTo, "capacityrequests@us.nifco.com", "Capacity Request", emailBody)
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Capacity_Results_Label_Click()
On Error GoTo Err_Handler

    Me.Capacity_Results.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Customer_Label_Click()
    On Error GoTo Err_Handler
    Me.Customer.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub EOP_Label_Click()
    On Error GoTo Err_Handler
    Me.EOP.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub NAM_Label_Click()
    On Error GoTo Err_Handler
    Me.NAM.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Planner_Label_Click()
    On Error GoTo Err_Handler
    Me.Planner.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub PPV_Label_Click()
    On Error GoTo Err_Handler
    Me.PPV.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Production_Type_Label_Click()
    On Error GoTo Err_Handler
    Me.Production_Type.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Program_Label_Click()
    On Error GoTo Err_Handler
    Me.Program.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Quote_Label_Click()
    On Error GoTo Err_Handler
    Me.Quote.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub RecordID_Label_Click()
    On Error GoTo Err_Handler
    Me.RecordID.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Request_Date_Label_Click()
    On Error GoTo Err_Handler
    Me.Request_Date.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Request_Type_Label_Click()
    On Error GoTo Err_Handler
    Me.Request_Type.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Requestor_Label_Click()
    On Error GoTo Err_Handler
    Me.Requestor.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Response_Date_Label_Click()
    On Error GoTo Err_Handler
    Me.Response_Date.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub SOP_Label_Click()
    On Error GoTo Err_Handler
    Me.SOP.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Unit_Label_Click()
    On Error GoTo Err_Handler
    Me.Unit.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Vehicle_Model_Label_Click()
    On Error GoTo Err_Handler
    Me.Vehicle_Model.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Volume_Label_Click()
    On Error GoTo Err_Handler
    Me.Volume.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Volume_Timing_Label_Click()
    On Error GoTo Err_Handler
    Me.Volume_Timing.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Volume_Type_Label_Click()
    On Error GoTo Err_Handler
    Me.Volume_Type.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub newRequest_Click()
    On Error GoTo Err_Handler
    DoCmd.OpenForm "frmCapacityRequestDetails", , , , acFormAdd
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub openDetails_Click()
    On Error GoTo ErrHandler
 
    If IsNull(Me.RecordID) Then Exit Sub
 
    DoCmd.OpenForm "frmCapacityRequestDetails", acNormal, , "recordId = " & Me.RecordID
 
Exit Sub
ErrHandler:
    MsgBox "Open Details error " & err.Number & ":" & vbCrLf & err.Description, vbExclamation
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", err.Description, err.Numbe)
End Sub
