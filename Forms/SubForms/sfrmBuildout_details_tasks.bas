attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub closetask_click()
on error goto err_handler

'if me.dirty then me.dirty = false
'if closebostep(me.recordid, "frmBuildout_details") then
'    call snackbox("success", "Success!", "Step Closed", "frmPartDashboard")
'    me.requery
'    me.refresh
'    call updatelastupdate
'end if

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub nudgeresponsible_click()
on error goto err_handler
'
'dim sendto as string, notisent as boolean, steptitle as string
'if nz(me.responsible) = "" then
'    call snackbox("error", "No can do", "Need a 'responsible' department to nudge", "frmPartDashboard")
'    exit sub
'end if
'notisent = false
'
'dim db as database
'set db = currentdb()
'dim rs1 as recordset, rs2 as recordset
'set rs1 = db.openrecordset("select * from tblPartTeam where partNumber = '" & me.partnumber & "' AND [person] is not null", dbopensnapshot)
'do while not rs1.eof
'    set rs2 = db.openrecordset("select * from tblPermissions where user = '" & rs1!person & "'", dbopensnapshot)
'    if rs2!dept <> me.responsible then goto nextone 'if dept isn't the same, skip this user
'    sendto = rs2!user
'    if sendto = environ("username") then
'        call snackbox("error", "No can do", "You can't nudge yourself!", "frmPartDashboard")
'        goto nextone
'    end if
'
'    steptitle = me.steptype
'
'    dim body as string
'    body = emailcontentgen("You've been nudged...", "Nudge Notification", "You've been nudged by " & getfullname() & " to complete this step", steptitle, "Part Number: " & me.partnumber, "Requested By: " & environ("username"), "Requested On: " & cstr(date), appname:="Part Project", appid:=me.partnumber)
'    if sendnotification(sendto, 1, 2, "Please complete step " & me.steptype & " for " & me.partnumber, body, "Part Project", me.partnumber) = true then
'        call snackbox("success", "Well done.", "Notification sent to " & sendto & "!", "frmPartDashboard")
'        call registerpartupdates("tblPartSteps", me.recordid, "Nudge", "From: " & environ("username"), "To: " & sendto, me.partnumber, me.steptype, me.partprojectid)
'        notisent = true
'    end if
'nextone:
'    rs1.movenext
'loop
'
'if not notisent then call snackbox("error", "Woops", "No one found", "frmPartDashboard")
'
'on error resume next
'rs1.close
'set rs1 = nothing
'rs2.close
'set rs2 = nothing
'set db = nothing
'
exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub showclosedtoggle_click()
on error goto err_handler
'
'me.showmystepstoggle = false
'
'dim extfilt as string
'if me.activecontrol.value then
'    extfilt = "="
'    me.lbldue.caption = "Closed"
'else
'    extfilt = "<>"
'    me.lbldue.caption = "Due"
'end if
'
'me.duedate.visible = not me.activecontrol.value
'me.filter = "partGateId = " & form_sfrmpartdashboarddates.recordid & " AND [status] " & extfilt & " 'Closed'"
'me.filteron = true
'
exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

sub updatelastupdate()
on error goto err_handler
'
'if me.recordset.recordcount = 0 then exit sub
'
'me.lastupdateddate = now()
'me.lastupdatedby = environ("username")
'
'if me.status = "Not Started" then
'    me.status = "In Progress"
'    call registerpartupdates("tblPartSteps", me.recordid, "Status", "Not Started", "In Progress", me.partnumber, me.steptype, me.partprojectid)
'end if
'
exit sub
err_handler:
    call handleerror(me.name, "updateLastUpdate", err.description, err.number)
end sub

public sub tasks_lock(lockit as boolean)
on error goto err_handler

me.closetask.enabled = not lockit
me.nudgeresponsible.enabled = not lockit
me.allowedits = not lockit

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub taskhistory_click()
on error goto err_handler
'
'docmd.openform "frmHistory", acnormal, , "[tableName] = 'tblPartSteps' AND [tableRecordId] = " & me.recordid
'
exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub tasknotes_afterupdate()
on error goto err_handler
'
'if isnull(me.closedate) = false then
'    if restrict(environ("username"), tempvars!projectowner, "Manager") = true then
'        me.activecontrol = me.activecontrol.oldvalue
'        call snackbox("error", "You can't edit a closed step", "Only a project/service manager and edit a closed step - looks like that's not you, sorry.", "frmPartDashboard")
'        exit sub
'    end if
'end if
'
'call registerpartupdates("tblPartSteps", me.recordid, me.activecontrol.name, me.activecontrol.oldvalue, me.activecontrol, me.partnumber, me.steptype, me.partprojectid)
'call updatelastupdate
'
exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub taskstatus_afterupdate()
on error goto err_handler
'
'if me.responsible = userdata("Dept") and dcount("recordId", "tblPartTeam", "person = '" & environ("username") & "' AND partNumber = '" & me.partnumber & "'") > 0 then goto goodtogo 'responsible and on cf team
'if restrict(environ("username"), tempvars!projectowner) = false then goto goodtogo 'is the user an owner
'me.activecontrol = me.activecontrol.oldvalue
'call snackbox("error", "Not today.", "Only Responsible person or PE can edit status", "frmPartDashboard")
'exit sub
'
'goodtogo:
'
'if isnull(me.closedate) = false and restrict(environ("username"), tempvars!projectowner, "Manager") then
'    me.activecontrol = me.activecontrol.oldvalue
'    call snackbox("error", "You can't edit a closed step", "Only a project/service manager and edit a closed step - looks like that's not you, sorry.", "frmPartDashboard")
'    exit sub
'end if
'
'if me.activecontrol.oldvalue = "Closed" then
'    me.closedate = null
'    call registerpartupdates("tblPartSteps", me.recordid, me.closedate.name, me.closedate.oldvalue, me.closedate, me.partnumber, me.steptype, me.partprojectid)
'end if
'
'call registerpartupdates("tblPartSteps", me.recordid, me.activecontrol.name, me.activecontrol.oldvalue, me.activecontrol, me.partnumber, me.steptype, me.partprojectid)
'me.lastupdateddate = now()
'me.lastupdatedby = environ("username")
'
exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
