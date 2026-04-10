option compare database
option explicit

public function createbuildoutproject(registerid as long, n0date as date, templateid as long) as boolean
    on error goto err_handler

    dim connread as adodb.connection: set connread = currentproject.connection
    dim connwrite as adodb.connection: set connwrite = currentproject.connection
    dim rsproject as new adodb.recordset
    dim rsgatetemplate as new adodb.recordset
    dim rstasktemplate as new adodb.recordset
    
    dim strinsert as string, runningdate as date
    dim gateid as long
    
    runningdate = n0date
    
    rsgatetemplate.open "SELECT * FROM tblBuildout_gates_template WHERE templateId = " & templateid & " ORDER BY indexOrder Asc", connread, adopenforwardonly, adlockreadonly
    
    do while not rsgatetemplate.eof
        runningdate = n0date - rsgatetemplate!gateduration
        
        connwrite.begintrans
        
        connwrite.execute "INSERT INTO tblBuildout_gates(registerId, dueDate, indexOrder, gateTemplateId) VALUES (" & _
                     registerid & ",Date()," & rsgatetemplate!indexorder & "," & rsgatetemplate!recordid & ")"
        
        gateid = connwrite.execute("SELECT @@IDENTITY")(0)
        
        connwrite.committrans
        connwrite.begintrans

        ' -- loop steps --
        rstasktemplate.open "SELECT * from tblBuildout_tasks_template WHERE [gateTemplateId] = " & rsgatetemplate![recordid] & " ORDER BY indexOrder Asc", connread, adopenforwardonly, adlockreadonly
        
        do while not rstasktemplate.eof
            strinsert = "INSERT INTO tblBuildout_tasks (gateId, taskStatus, templateTaskId, createdBy, createdDate, lastUpdatedDate, lastUpdatedBy, indexOrder) " & _
                        "VALUES (" & gateid & ",1," & rstasktemplate!recordid & ",'" & environ("username") & "','" & format$(now(), "yyyy-mm-dd\Thh:nn:ss") & "','" & format$(now(), "yyyy-mm-dd\Thh:nn:ss") & "','" & environ("username") & _
                        "'," & rstasktemplate!indexorder & ")"
            
            connwrite.execute strinsert
            
nextstep:
            rstasktemplate.movenext
        loop
        
        rstasktemplate.close
        set rstasktemplate = nothing
        
        connwrite.committrans
        
        rsgatetemplate.movenext
    loop
    
    

    createbuildoutproject = true

cleanup:
    set connwrite = nothing
    set connread = nothing
    if rsgatetemplate.state = adstateopen then rsgatetemplate.close
    if rstasktemplate.state = adstateopen then rstasktemplate.close
    exit function

err_handler:
    if not connwrite is nothing then
        if connwrite.state = adstateopen then connwrite.rollbacktrans
    end if
    call handleerror("modBuildout", "createBuildoutProject", err.description, err.number)
    resume cleanup
end function
