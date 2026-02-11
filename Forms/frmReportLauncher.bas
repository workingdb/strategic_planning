Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnGenerate_Click()
On Error GoTo Err_Handler
 
    Dim rt As String
    Dim rptName As String
    Dim whereClause As String
    Dim d1 As Date, d2 As Date
 
    'Use displayed text if cboReportType is multi-column; fallback to Value if not
    rt = Trim(Nz(Me.cboReportType.Column(1), Nz(Me.cboReportType.Value, "")))
 
    If rt = "" Then
        MsgBox "Please select a Report Type.", vbExclamation
        Exit Sub
    End If
 
    Select Case rt
 
        Case "Capacity by Date Range"
            rptName = "rpt_Capacity_DateRange"
 
            If IsNull(Me.txtStartDate) Or IsNull(Me.txtEndDate) Then
                MsgBox "Please enter both Start Date and End Date.", vbExclamation
                Exit Sub
            End If
            If Not IsDate(Me.txtStartDate) Or Not IsDate(Me.txtEndDate) Then
                MsgBox "Start Date and End Date must be valid dates.", vbExclamation
                Exit Sub
            End If
 
            d1 = DateValue(Me.txtStartDate)
            d2 = DateValue(Me.txtEndDate)
 
            If d2 < d1 Then
                MsgBox "End Date must be on or after Start Date.", vbExclamation
                Exit Sub
            End If
 
            whereClause = "[RequestDate] Between #" & Format(d1, "yyyy-mm-dd") & _
                          "# And #" & Format(d2, "yyyy-mm-dd") & "#"
 
        Case "KPI Report"
            rptName = "rpt_KPI_Dashboard"
 
             If IsNull(Me.txtStartDate) Or IsNull(Me.txtEndDate) Then
                 MsgBox "Please enter both Start Date and End Date.", vbExclamation
                 Exit Sub
             End If
             If Not IsDate(Me.txtStartDate) Or Not IsDate(Me.txtEndDate) Then
                 MsgBox "Start Date and End Date must be valid dates.", vbExclamation
                 Exit Sub
             End If
 
             d1 = DateValue(Me.txtStartDate)
             d2 = DateValue(Me.txtEndDate)
 
             If d2 < d1 Then
                MsgBox "End Date must be on or after Start Date.", vbExclamation
                Exit Sub
             End If
 
    'Set TempVars for KPI queries
    On Error Resume Next
    TempVars.Remove "StartDate"
    TempVars.Remove "EndDate"
    On Error GoTo 0
 
    TempVars.Add "StartDate", d1
    TempVars.Add "EndDate", d2
 
    'Open without filter
    DoCmd.OpenReport rptName, acViewPreview
    Exit Sub
            
 
        Case "Capacity by NAM"
            rptName = "rpt_Capacity_ByNAM"
 
            If IsNull(Me.txtNAM) Or Trim(Nz(Me.txtNAM, "")) = "" Then
                MsgBox "Please enter a NAM.", vbExclamation
                Exit Sub
            End If
 
            whereClause = "[partNumber] = """ & Replace(Trim(Me.txtNAM), """", """""") & """"
 
        Case "Requests by Sales Manager"
            rptName = "rpt_SalesManager"
 
            If IsNull(Me.cboSalesManager) Then
                MsgBox "Please select a Sales Manager.", vbExclamation
                Exit Sub
            End If
 
            whereClause = "[Requestor] = " & CLng(Me.cboSalesManager)
 
        Case "Capacity by Program Code"
            rptName = "rpt_Capacity_ByProgramCode"
 
            If IsNull(Me.cboProgramCode) Or Trim(Nz(Me.cboProgramCode, "")) = "" Then
                MsgBox "Please select a Program Code.", vbExclamation
                Exit Sub
            End If
 
            whereClause = "[Program] = """ & Replace(Trim(Me.cboProgramCode), """", """""") & """"
 
        Case "Past Due Report"
            rptName = "rpt_PastDue"
            
             If IsNull(Me.cboUnit) Then
                MsgBox "Please select a Unit.", vbExclamation
                Exit Sub
            End If
 
            whereClause = "[Unit] = " & CLng(Me.cboUnit)
 
        Case Else
            MsgBox "Report type not coded yet: " & rt, vbExclamation
            Exit Sub
    End Select
 
    If Len(whereClause) > 0 Then
        DoCmd.OpenReport rptName, acViewPreview, , whereClause
    Else
        DoCmd.OpenReport rptName, acViewPreview
    End If
 
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub
 

Private Sub cboReportType_AfterUpdate()
On Error GoTo Err_Handler
 
    '--- Hide everything first ---
    Me.txtStartDate.Visible = False
    Me.txtEndDate.Visible = False
    Me.txtNAM.Visible = False
    Me.cboSalesManager.Visible = False
    Me.cboProgramCode.Visible = False
    Me.cboUnit.Visible = False
 
    '--- Default: disable + no tabbing ---
    Me.txtStartDate.Enabled = False: Me.txtStartDate.TabStop = False
    Me.txtEndDate.Enabled = False: Me.txtEndDate.TabStop = False
    Me.txtNAM.Enabled = False: Me.txtNAM.TabStop = False
    Me.cboSalesManager.Enabled = False: Me.cboSalesManager.TabStop = False
    Me.cboProgramCode.Enabled = False: Me.cboProgramCode.TabStop = False
    Me.cboUnit.Enabled = False: Me.cboUnit.TabStop = False
 
    '--- Show only what the selected report needs ---
    Select Case Me.cboReportType.Value
 
        Case "Capacity by Date Range", "KPI Report"
            Me.txtStartDate.Visible = True
            Me.txtEndDate.Visible = True
            Me.txtStartDate.Enabled = True: Me.txtStartDate.TabStop = True
            Me.txtEndDate.Enabled = True: Me.txtEndDate.TabStop = True
 
        Case "Capacity by NAM"
            Me.txtNAM.Visible = True
            Me.txtNAM.Enabled = True: Me.txtNAM.TabStop = True
 
        Case "Requests by Sales Manager"
            Me.cboSalesManager.Visible = True
            Me.cboSalesManager.Enabled = True: Me.cboSalesManager.TabStop = True
 
        Case "Capacity by Program Code"
            Me.cboProgramCode.Visible = True
            Me.cboProgramCode.Enabled = True: Me.cboProgramCode.TabStop = True
 
        Case "Past Due Report"
            Me.cboUnit.Visible = True
            Me.cboUnit.Enabled = True: Me.cboUnit.TabStop = True
 
    End Select
 
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub
 
Private Sub Form_Load()
On Error GoTo Err_Handler
    'Clear all user selections/inputs when the launcher opens
 
 Call setTheme(Me)
 
    On Error Resume Next
 
    'Combos
    Me.cboReportType = Null
    Me.cboSalesManager = Null
    Me.cboProgramCode = Null
    Me.cboUnit = Null
 
    'Text boxes
    Me.txtStartDate = Null
    Me.txtEndDate = Null
    Me.txtNAM = Null
 
    'If you have any other inputs, add them here:
    'Me.txtSomething = Null
    'Me.cboSomething = Null
 
    'Optional: clear tempvars so nothing carries over
    TempVars.Remove "StartDate"
    TempVars.Remove "EndDate"
    
    Me.cboReportType.SetFocus
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.Description, Err.Number)
End Sub
