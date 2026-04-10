attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub deletebtn_click()
on error goto err_handler

if isnull(me.recordid) then
    msgbox "This is an empty record.", vbinformation, "Can't do that"
    exit sub
end if

if msgbox("Are you sure you want to delete this?", vbyesno, "Please confirm") = vbyes then
    dim oldindex
    oldindex = me.indexorder
    if oldindex < dmax("indexOrder", "tblBuildout_tasks_template", "gateTemplateId = " & me.gatetemplateid) then
        dbexecute "UPDATE tblBuildout_tasks_template SET indexOrder = indexOrder - 1 WHERE gateTemplateId = " & me.gatetemplateid & " AND indexOrder > " & oldindex
    end if
    
    call registerstratplanupdates("tblBuildout_tasks_template", me.recordid, "DELETE", me.tasktitle, "DELETED", form_frmbuildout_template.recordid, "frmBuildout_template")
    dbexecute "DELETE FROM tblBuildout_tasks_template WHERE recordId = " & me.recordid
    me.requery
end if

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub form_beforeinsert(cancel as integer)
on error goto err_handler

me.indexorder = nz(dmax("indexOrder", "tblBuildout_tasks_template", "gateTemplateId = " & me.gatetemplateid) + 1, 1)
me.dirty = false

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub form_current()
on error goto err_handler

me.txtcf = me.recordid

exit sub
err_handler:
    call handleerror(me.name, "Form_Current", err.description, err.number)
end sub

private sub form_load()
on error goto err_handler

call settheme(me)

me.orderby = "indexOrder Asc"
me.orderbyon = true

exit sub
err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.number)
end sub

private sub movedown_click()
on error goto err_handler

if isnull(me.recordid) then exit sub

if me.indexorder = dmax("indexOrder", "tblBuildout_tasks_template", "gateTemplateId = " & me.gatetemplateid) then exit sub

dim oldindex, newindex
oldindex = me.indexorder
newindex = oldindex + 1

dbexecute "UPDATE tblBuildout_tasks_template SET indexOrder = " & oldindex & " WHERE gateTemplateId = " & me.gatetemplateid & " AND indexOrder = " & newindex
me.indexorder = newindex
me.dirty = false
    
me.requery
me.orderby = "indexOrder Asc"
me.orderbyon = true

call registerstratplanupdates("tblBuildout_tasks_template", me.recordid, "indexOrder", oldindex, newindex, form_frmbuildout_template.recordid, "frmBuildout_template")

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub moveup_click()
on error goto err_handler

if isnull(me.recordid) then exit sub
if me.indexorder = 1 then exit sub
dim oldindex, newindex
oldindex = me.indexorder
newindex = oldindex - 1

dbexecute "UPDATE tblBuildout_tasks_template SET indexOrder = " & oldindex & " WHERE gateTemplateId = " & me.gatetemplateid & " AND indexOrder = " & newindex
me.indexorder = newindex
me.dirty = false
    
me.requery
me.orderby = "indexOrder Asc"
me.orderbyon = true

call registerstratplanupdates("tblBuildout_tasks_template", me.recordid, "indexOrder", oldindex, newindex, form_frmbuildout_template.recordid, "frmBuildout_template")

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub newstep_click()
on error goto err_handler

dbexecute "INSERT INTO tblBuildout_tasks_template(gateTemplateId,indexOrder) VALUES (" & form_sfrmbuildout_template_gates.recordid & "," & _
    nz(dmax("indexOrder", "tblBuildout_tasks_template", "gateTemplateId = " & form_sfrmbuildout_template_gates.recordid) + 1, 1) & ")"
    
me.requery

call registerstratplanupdates("tblBuildout_tasks_template", me.recordid, "New", "", "New Record", form_frmbuildout_template.recordid, "frmBuildout_template")

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub pillartask_afterupdate()
on error goto err_handler

call registerstratplanupdates("tblBuildout_tasks_template", me.recordid, me.activecontrol.name, me.activecontrol.oldvalue, me.activecontrol, form_frmbuildout_template.recordid, "frmBuildout_template")

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub responsibledept_afterupdate()
on error goto err_handler

call registerstratplanupdates("tblBuildout_tasks_template", me.recordid, me.activecontrol.name, me.activecontrol.oldvalue, me.activecontrol, form_frmbuildout_template.recordid, "frmBuildout_template")

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub tasktitle_afterupdate()
on error goto err_handler

call registerstratplanupdates("tblBuildout_tasks_template", me.recordid, me.activecontrol.name, me.activecontrol.oldvalue, me.activecontrol, form_frmbuildout_template.recordid, "frmBuildout_template")

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
