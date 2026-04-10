attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub allrequests_click()
on error goto err_handler

docmd.openform "frmCapacityRequestTracker"

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub btnmaintenance_click()
on error goto err_handler

docmd.openform "frmMaintanence", acnormal, "", "", , acnormal

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub btnopenreportlauncher_click()
on error goto err_handler

docmd.openform "frmReportLauncher", acnormal

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub btnsettings_click()
on error goto err_handler

docmd.openform "frmUserView"

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub cmdopensalesmanagerreport_click()
on error goto err_handler

docmd.openreport "rpt_SalesManager", acviewreport

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub btnsurvery_click()
on error goto err_handler

docmd.openform "frmSurveys", acnormal, "", "", , acnormal
    
exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub form_load()
on error goto err_handler

call settheme(me)

exit sub
err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.numbe)
end sub
