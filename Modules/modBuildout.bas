option compare database
option explicit

public function createbuildoutproject(registerid as long, n0date as date, templateid as long) as boolean
    on error goto err_handler
    
    dim conn as new adodb.connection
    conn.open replace(relinksqltables(true), "ODBC;", "")

    dim rsproject as adodb.recordset
    dim rsgatetemplate as adodb.recordset
    dim rstasktemplate as adodb.recordset
    
    dim strinsert as string, runningdate as date
    dim gateid as long
    
    runningdate = n0date
    
    set rsgatetemplate = openrecordsetreadonly(conn, "SELECT * FROM tblBuildout_gates_template WHERE templateId = " & templateid & " ORDER BY indexOrder Asc")
    
    do while not rsgatetemplate.eof
        runningdate = n0date - rsgatetemplate!gateduration
        
        conn.begintrans
        
        conn.execute "INSERT INTO tblBuildout_gates(registerId, dueDate, indexOrder, gateTemplateId) VALUES (" & _
                     registerid & ",Date()," & rsgatetemplate!indexorder & "," & rsgatetemplate!recordid & ")"
        
        gateid = conn.execute("SELECT @@IDENTITY")(0)

        ' -- loop steps --
        set rstasktemplate = openrecordsetreadonly(conn, "SELECT * from tblBuildout_tasks_template WHERE [gateTemplateId] = " & rsgatetemplate![recordid] & " ORDER BY indexOrder Asc")
        
        do while not rstasktemplate.eof
            strinsert = "INSERT INTO tblBuildout_tasks (gateId, taskStatus, templateTaskId, createdBy, createdDate, lastUpdatedDate, lastUpdatedBy, indexOrder) " & _
                        "VALUES (" & gateid & ",1," & rstasktemplate!recordid & "," & userdata("ID") & ",'" & format$(now(), "yyyy-mm-dd\Thh:nn:ss") & "','" & format$(now(), "yyyy-mm-dd\Thh:nn:ss") & "'," & userdata("ID") & _
                        "," & rstasktemplate!indexorder & ")"
            
            conn.execute strinsert
            
nextstep:
            rstasktemplate.movenext
        loop
        
        rstasktemplate.close
        set rstasktemplate = nothing
        
        conn.committrans
        
        rsgatetemplate.movenext
    loop
    
    

    createbuildoutproject = true
    
cleanup:
    conn.close
    set conn = nothing
    if rsgatetemplate.state = adstateopen then rsgatetemplate.close
    if rstasktemplate.state = adstateopen then rstasktemplate.close
    exit function

err_handler:
    if not conn is nothing then
        if conn.state = adstateopen then conn.rollbacktrans
    end if
    call handleerror("modBuildout", "createBuildoutProject", err.description, err.number)
    resume cleanup
end function

function closeprojectstep(stepid as long, frmactive as string) as boolean
on error goto err_handler

    dim rsstep as adodb.recordset
    dim rsgate as adodb.recordset
    dim errortext as string
    dim currentdate as date

    closeprojectstep = false
    
    dim conn as new adodb.connection
    conn.open replace(relinksqltables(true), "ODBC;", "")
    
    set rsstep = openrecordsetreadwrite(conn, _
        "SELECT * FROM dbo.tblPartSteps WHERE recordId = " & stepid)

    if rsstep.eof then
        call snackbox("error", "Darn", "Step not found.", frmactive)
        goto cleanexit
    end if
    
    set rsgate = openrecordsetreadwrite(conn, _
        "SELECT * FROM dbo.tblPartGates WHERE recordId = " & clng(rsstep!partgateid))

    if rsgate.eof then
        call snackbox("error", "Darn", "Gate not found.", frmactive)
        goto cleanexit
    end if
    
    errortext = closeprojectstep_gatepillar(conn, rsstep)
    if errortext <> "" then goto validationerror

    errortext = closeprojectstep_auth(conn, rsstep)
    if errortext <> "" then goto validationerror

    errortext = validatestepreadytoclose(conn, rsstep)
    if errortext <> "" then goto validationerror

    if errortext = "__EXIT__" then goto cleanexit
    if errortext <> "" then goto validationerror

    currentdate = now()

    'call registerpartupdates("tblPartSteps", rsstep!recordid, "closeDate", "", currentdate, rsstep!partnumber, rsstep!steptype, rsstep!partprojectid)
    'call registerpartupdates("tblPartSteps", rsstep!recordid, "status", rsstep!status, "Closed", rsstep!partnumber, rsstep!steptype, rsstep!partprojectid)

    rsstep!closedate = currentdate
    rsstep!status = "Closed"
    rsstep.update

    'call notifype(rsstep!partnumber, "Closed", rsstep!steptype)

    if getopenstepcountforgate(conn, rsstep!partgateid) = 0 then
        'call registerpartupdates("tblPartGates", rsstep!partgateid, "actualDate", rsgate!actualdate, currentdate, rsstep!partnumber, rsgate!gatetitle, rsstep!partprojectid)
        rsgate!actualdate = currentdate
        rsgate.update

        if frmactive = "frmPartDashboard" then
            'form_frmpartdashboard.partdash_refresh_click
        end if
    end if

    closeprojectstep = true
    goto cleanexit

validationerror:
    call snackbox("error", "Darn", errortext, frmactive)

cleanexit:
    on error resume next
    if not rsgate is nothing then if rsgate.state = adstateopen then rsgate.close
    if not rsstep is nothing then if rsstep.state = adstateopen then rsstep.close
    set rsgate = nothing
    set rsstep = nothing
    set conn = nothing
    exit function

err_handler:
    call handleerror("modBuildout", "closeprojectstep", err.description, err.number)
    resume cleanexit
end function

private function getopenstepcountforgate(conn as adodb.connection, partgateid as long) as long
    getopenstepcountforgate = sqlcount(conn, "tblPartSteps", "[closeDate] IS NULL AND partGateId = " & partgateid)
end function

private function closeprojectstep_gatepillar(conn as adodb.connection, rsstep as adodb.recordset) as string
    dim bypass as boolean
    dim bypassinfo as string
    dim bypassorg as string
    dim gateid as long
    dim partnumbersafe as string

    closeprojectstep_gatepillar = ""
    partnumbersafe = replace(nz(rsstep!partnumber, ""), "'", "''")

    bypass = false

    'check parameter for temp bypass
    if nz(sqllookup(conn, "paramVal", "dbo.tblDBinfoBE", "parameter = 'allowGatePillarBypass'"), false) = true then
        bypassinfo = nz(sqllookup(conn, "Message", "dbo.tblDBinfoBE", "parameter = 'allowGatePillarBypass'"), "")
        bypassorg = nz(sqllookup(conn, "developingLocation", "dbo.tblPartInfo", "partNumber = '" & partnumbersafe & "'"), "SLB")

        if len(bypassinfo) = 3 then
            if bypassorg = "LVG" then bypassorg = "CNL"
            if bypassinfo = bypassorg then bypass = true
        else
            if bypassinfo = environ$("username") then bypass = true
        end if
    end if

    if bypass then exit function

    '---first, check if this step is in the current gate--- (you can only close it if true)
    gateid = nz( _
        sqlmin(conn, "partGateId", "dbo.tblPartSteps", "partProjectId = " & rsstep!partprojectid & " AND [status] <> 'Closed'"), _
        sqlmin(conn, "partGateId", "dbo.tblPartSteps", "partProjectId = " & rsstep!partprojectid))

    if gateid <> clng(rsstep!partgateid) then
        closeprojectstep_gatepillar = "This step is not in the current gate, you can't close it yet"
        exit function
    end if

    if not isnull(rsstep!duedate) then
        if sqlcount(conn, "dbo.tblPartSteps", _
                  "partGateId = " & rsstep!partgateid & _
                  " AND indexOrder < " & rsstep!indexorder & _
                  " AND [status] <> 'Closed'") > 0 then
            closeprojectstep_gatepillar = "This step is a pillar. All steps before this pillar must be closed before this step."
        end if
    end if
end function

function closeprojectstep_auth(conn as adodb.connection, rsstep as adodb.recordset) as string

    dim projectowner as string
    dim projecttemplateid as variant
    dim templatetype as variant
    dim partnumbersafe as string

    closeprojectstep_auth = ""
    partnumbersafe = replace(nz(rsstep!partnumber, ""), "'", "''")

    projecttemplateid = sqllookup(conn, "projectTemplateId", "dbo.tblPartProject", "recordId = " & rsstep!partprojectid)
    templatetype = sqllookup(conn, "templateType", "dbo.tblPartProjectTemplate", "recordId = " & nz(projecttemplateid, 0))

    select case nz(templatetype, 0)
        case 1: projectowner = "Project"
        case 2: projectowner = "Service"
        case else: projectowner = ""
    end select

    'temporary restriction override
    'project engineers can close steps for other departments until all departments are fully in
    if restrict(environ("username"), projectowner) = false then exit function

    if nz(rsstep!responsible, "") = userdata("Dept") and _
       sqlcount(conn, "dbo.tblPartTeam", "[person] = '" & environ("username") & "' AND partNumber = '" & partnumbersafe & "'") > 0 then
        exit function
    end if

    if restrict(environ("username"), projectowner, "Manager") = false then exit function
    if restrict(environ("username"), nz(rsstep!responsible, ""), "Manager") = false then exit function

    closeprojectstep_auth = "Only the 'Responsible' person, their manager, or a project/service Manager can close a step"

end function

private function validatestepreadytoclose(conn as adodb.connection, rsstep as adodb.recordset) as string
    validatestepreadytoclose = ""

    if not isnull(rsstep!closedate) then
        validatestepreadytoclose = "This is already closed - what's the point in closing again?"
        exit function
    end if
end function
