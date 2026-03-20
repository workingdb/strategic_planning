Option Compare Database
Option Explicit

Public Function createPartProject(projId As Long, Optional opT0 As Date) As Boolean
    On Error GoTo Err_Handler

    Dim conn As ADODB.Connection: Set conn = CurrentProject.Connection
    Dim rsProject As New ADODB.Recordset, rsGateTemplate As New ADODB.Recordset
    Dim rsStepTemplate As New ADODB.Recordset, rsSess As New ADODB.Recordset
    Dim strInsert As String, pNum As String, runningDate As Date, G3planned As Date
    Dim projTempId As Long, gateId As Long, stepId As Long, strSQL As String

    ' 1. Load Project Info
    rsProject.Open "SELECT projectTemplateId, partNumber, projectStartDate FROM tblPartProject WHERE recordId = " & projId, conn, adOpenForwardOnly, adLockReadOnly
    If rsProject.EOF Then GoTo CleanUp
    
    projTempId = rsProject!projectTemplateId
    pNum = Nz(rsProject!partNumber, "")
    runningDate = rsProject!projectStartDate
    rsProject.CLOSE

    If pNum = "" Then GoTo CleanUp

    ' 1. Check if the person is already on the team
    Dim rsCheck As ADODB.Recordset
    strSQL = "SELECT COUNT(*) FROM tblPartTeam WHERE partNumber = '" & pNum & "' AND person = '" & Environ("username") & "'"
    Set rsCheck = conn.Execute(strSQL)
    
    ' 2. Insert only if the count is 0
    If rsCheck(0) = 0 Then
        conn.Execute "INSERT INTO tblPartTeam (partNumber, person) VALUES ('" & pNum & "', '" & Environ("username") & "')"
    End If
    rsCheck.CLOSE


    ' 3. Open Templates
    rsGateTemplate.Open "SELECT * FROM tblPartGateTemplate WHERE [projectTemplateId] = " & projTempId, conn, adOpenForwardOnly, adLockReadOnly
    rsSess.Open "SELECT * FROM tblSessionVariables WHERE pillarTitle IS NOT NULL", conn, adOpenStatic, adLockReadOnly
    
    ' -- LOOP GATES --
    Do While Not rsGateTemplate.EOF
        runningDate = addWorkdays(runningDate, rsGateTemplate![gateDuration])
        
        conn.Execute "INSERT INTO tblPartGates(projectId, partNumber, gateTitle, plannedDate) VALUES (" & _
                     projId & ",'" & pNum & "','" & rsGateTemplate![gateTitle] & "',#" & Format(runningDate, "yyyy-mm-dd") & "#)"
        
        gateId = conn.Execute("SELECT @@IDENTITY")(0)

        ' -- LOOP STEPS --
        rsStepTemplate.Open "SELECT * from tblPartStepTemplate WHERE [gateTemplateId] = " & rsGateTemplate![RecordID] & " ORDER BY indexOrder Asc", conn, adOpenForwardOnly, adLockReadOnly
        
        Do While Not rsStepTemplate.EOF
            If Nz(rsStepTemplate![Title], "") = "" Then GoTo nextStep
            
            Dim stepDueDate As String: stepDueDate = "NULL"
            If rsStepTemplate!pillarStep Then
                rsSess.Filter = "pillarStepId = " & rsStepTemplate!RecordID
                If Not rsSess.EOF Then
                    stepDueDate = "#" & Format(rsSess!pillarDue, "yyyy-mm-dd") & "#"
                Else
                    GoTo nextStep
                End If
            End If
            
            ' Brackets [] on reserved words like [status], [duration]
            strInsert = "INSERT INTO tblPartSteps (partNumber, partProjectId, partGateId, stepType, openedBy, [status], openDate, lastUpdatedDate, lastUpdatedBy, stepActionId, documentType, responsible, indexOrder, [duration], dueDate) " & _
                        "VALUES ('" & pNum & "'," & projId & "," & gateId & ",'" & Replace(rsStepTemplate![Title], "'", "''") & "','" & _
                        Environ("username") & "','Not Started',Now(),Now(),'" & Environ("username") & "'," & _
                        Nz(rsStepTemplate![stepActionId], "NULL") & "," & Nz(rsStepTemplate![documentType], "NULL") & ",'" & _
                        Replace(Nz(rsStepTemplate![responsible], ""), "'", "''") & "'," & rsStepTemplate![indexOrder] & "," & Nz(rsStepTemplate![duration], 1) & "," & stepDueDate & ")"
            
            conn.Execute strInsert
            stepId = conn.Execute("SELECT @@IDENTITY")(0)
            
            ' -- BULK INSERT APPROVALS (Optimized) --
            conn.Execute "INSERT INTO tblPartTrackingApprovals (partNumber, requestedBy, requestedDate, dept, reqLevel, tableName, tableRecordId) " & _
                         "SELECT '" & pNum & "', '" & Environ("username") & "', Now(), dept, reqLevel, 'tblPartSteps', " & stepId & " " & _
                         "FROM tblPartStepTemplateApprovals WHERE stepTemplateId = " & rsStepTemplate!RecordID

nextStep:
            rsSess.Filter = adFilterNone
            rsStepTemplate.MoveNext
        Loop
        rsStepTemplate.CLOSE
        
        If Left(rsGateTemplate!gateTitle, 2) = "G3" Then G3planned = runningDate
        rsGateTemplate.MoveNext
    Loop

    ' 4. Assembled Parts Automation Logic
    If projTempId = 8 Then
        Dim totalDays As Long: totalDays = Nz(conn.Execute("SELECT SUM([duration]) FROM tblPartStepTemplate WHERE gateTemplateId = 43")(0), 0)
        Dim assyDate As Date: assyDate = addWorkdays(G3planned, (totalDays + 15) * -1)
        
        Dim rsAssy As New ADODB.Recordset
        rsAssy.Open "SELECT * FROM tblPartStepTemplate WHERE gateTemplateId = 43", conn, adOpenForwardOnly, adLockReadOnly
        
        Do While Not rsAssy.EOF
            assyDate = addWorkdays(assyDate, Nz(rsAssy![duration], 1))
            conn.Execute "INSERT INTO tblPartAssemblyGates(projectId, templateGateId, partNumber, gateStatus, plannedDate) VALUES (" & _
                         projId & "," & rsAssy!RecordID & ",'" & pNum & "',1,#" & Format(assyDate, "yyyy-mm-dd") & "#)"
            rsAssy.MoveNext
        Loop
        rsAssy.CLOSE
    End If

    createPartProject = True ' Signal success to the calling function

CleanUp:
    If rsSess.State = adStateOpen Then rsSess.CLOSE
    If rsGateTemplate.State = adStateOpen Then rsGateTemplate.CLOSE
    Exit Function

Err_Handler:
    createPartProject = False ' Signal failure to trigger Rollback in the calling sub
    Call handleError("wdbProjectE", "createPartProject", Err.Description, Err.Number)
    Resume CleanUp
End Function