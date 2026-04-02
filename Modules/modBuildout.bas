Option Compare Database
Option Explicit

Public Function createBuildoutProject(registerId As Long, n0Date As Date, templateId As Long) As Boolean
    On Error GoTo Err_Handler

    Dim connRead As adodb.Connection: Set connRead = CurrentProject.Connection
    Dim connWrite As adodb.Connection: Set connWrite = CurrentProject.Connection
    Dim rsProject As New adodb.Recordset
    Dim rsGateTemplate As New adodb.Recordset
    Dim rsTaskTemplate As New adodb.Recordset
    
    Dim strInsert As String, runningDate As Date
    Dim gateId As Long
    
    runningDate = n0Date
    
    rsGateTemplate.Open "SELECT * FROM tblBuildout_gates_template WHERE templateId = " & templateId & " ORDER BY indexOrder Asc", connRead, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rsGateTemplate.EOF
        runningDate = n0Date - rsGateTemplate!gateDuration
        
        connWrite.BeginTrans
        
        connWrite.Execute "INSERT INTO tblBuildout_gates(registerId, dueDate, indexOrder, gateTemplateId) VALUES (" & _
                     registerId & ",Date()," & rsGateTemplate!indexOrder & "," & rsGateTemplate!RecordID & ")"
        
        gateId = connWrite.Execute("SELECT @@IDENTITY")(0)
        
        connWrite.CommitTrans
        connWrite.BeginTrans

        ' -- LOOP STEPS --
        rsTaskTemplate.Open "SELECT * from tblBuildout_tasks_template WHERE [gateTemplateId] = " & rsGateTemplate![RecordID] & " ORDER BY indexOrder Asc", connRead, adOpenForwardOnly, adLockReadOnly
        
        Do While Not rsTaskTemplate.EOF
            strInsert = "INSERT INTO tblBuildout_tasks (gateId, taskStatus, templateTaskId, createdBy, createdDate, lastUpdatedDate, lastUpdatedBy, indexOrder) " & _
                        "VALUES (" & gateId & ",1," & rsTaskTemplate!RecordID & ",'" & Environ("username") & "','" & Format$(Now(), "yyyy-mm-dd\Thh:nn:ss") & "','" & Format$(Now(), "yyyy-mm-dd\Thh:nn:ss") & "','" & Environ("username") & _
                        "'," & rsTaskTemplate!indexOrder & ")"
            
            connWrite.Execute strInsert
            
nextStep:
            rsTaskTemplate.MoveNext
        Loop
        
        rsTaskTemplate.CLOSE
        Set rsTaskTemplate = Nothing
        
        connWrite.CommitTrans
        
        rsGateTemplate.MoveNext
    Loop
    
    

    createBuildoutProject = True

CleanUp:
    Set connWrite = Nothing
    Set connRead = Nothing
    If rsGateTemplate.State = adStateOpen Then rsGateTemplate.CLOSE
    If rsTaskTemplate.State = adStateOpen Then rsTaskTemplate.CLOSE
    Exit Function

Err_Handler:
    If Not connWrite Is Nothing Then
        If connWrite.State = adStateOpen Then connWrite.RollbackTrans
    End If
    Call handleError("modBuildout", "createBuildoutProject", err.Description, err.Number)
    Resume CleanUp
End Function