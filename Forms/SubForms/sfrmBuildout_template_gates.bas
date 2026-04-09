attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub form_current()
on error goto err_handler

me.txtcf = me.recordid

if isnull(me.recordid) = false then
    form_sfrmbuildout_template_tasks.visible = true
    form_sfrmbuildout_template_tasks.filter = "gateTemplateId = " & me.recordid
    form_sfrmbuildout_template_tasks.gatetemplateid.defaultvalue = me.recordid
    form_sfrmbuildout_template_tasks.filteron = true
else
    form_sfrmbuildout_template_tasks.visible = false
end if

exit sub
err_handler:
    call handleerror(me.name, "Form_Current", err.description, err.number)
end sub

private sub gateduration_afterupdate()
on error goto err_handler

call registerstratplanupdates("tblBuildout_tasks_template", me.recordid, me.activecontrol.name, me.activecontrol.oldvalue, me.activecontrol, form_frmbuildout_template.recordid, "frmBuildout_template")

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub gatetitle_afterupdate()
on error goto err_handler

call registerstratplanupdates("tblBuildout_tasks_template", me.recordid, me.activecontrol.name, me.activecontrol.oldvalue, me.activecontrol, form_frmbuildout_template.recordid, "frmBuildout_template")

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
