attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub form_load()
on error goto err_handler

call settheme(me)

dim templateid
templateid = dmin("recordId", "tblBuildout_gates_template", "templateId = " & me.recordid)

on error goto invis
me.sfrmbuildout_template_tasks.form.filter = "gateTemplateId = " & templateid
me.sfrmbuildout_template_tasks.form.gatetemplateid.defaultvalue = templateid
me.sfrmbuildout_template_tasks.form.filteron = true

exit sub
invis:
me.sfrmbuildout_template_tasks.visible = false

exit sub
err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.number)
end sub

private sub history_click()
on error goto err_handler

docmd.openform "frmHistory", , , "referenceId = " & me.recordid

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub refresh_click()
on error goto err_handler

me.requery
me.sfrmbuildout_template_tasks.requery
me.sfrmbuildout_template_gates.requery

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub templatename_afterupdate()
on error goto err_handler

call registerstratplanupdates("tblBuildout_tasks_template", me.recordid, me.activecontrol.name, me.activecontrol.oldvalue, me.activecontrol, form_frmbuildout_template.recordid, "frmBuildout_template")

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
