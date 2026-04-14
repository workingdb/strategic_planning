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
    form_sfrmbuildout_details_tasks.visible = true
    form_sfrmbuildout_details_tasks.filter = "gateId = " & me.recordid & " AND taskStatus < 3"
    form_sfrmbuildout_details_tasks.filteron = true
    form_sfrmbuildout_details_tasks.showclosedtoggle.value = false
    call allowstepedit
else
    form.sfrmpartdashboard.visible = false
end if

exit sub
err_handler:
    call handleerror(me.name, "Form_Current", err.description, err.number)
end sub

public sub allowstepedit()
on error goto err_handler

dim msgcap as string
msgcap = "Steps: " & me.gatetitle
form_sfrmbuildout_details_tasks.form.allowedits = true
call form_sfrmbuildout_details_tasks.tasks_lock(false)

'are there steps in a previous gate that are open? must finish those first
if me.recordid > dmin("gateId", "tblBuildout_tasks", "registerId = " & me.registerid & " AND taskStatus < 3") then
    msgcap = "Tasks: " & me.gatetitle & " (LOCKED)"
    call form_sfrmbuildout_details_tasks.tasks_lock(true)
    goto setmsg
end if

setmsg:
form_sfrmbuildout_details_tasks.lbltasks.caption = msgcap

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
