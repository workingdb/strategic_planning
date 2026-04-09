attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub details_click()
on error goto err_handler

me.dirty = false
docmd.openform "frmBuildout_template", acnormal, , "recordId = " & me.recordid

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub projecttitle_afterupdate()
on error goto err_handler

call registerstratplanupdates("tblBuildout_tasks_template", me.recordid, me.activecontrol.name, me.activecontrol.oldvalue, me.activecontrol, me.recordid, "frmBuildout_templates")

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
