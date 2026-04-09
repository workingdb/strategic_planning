attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub btncustomers_click()
on error goto err_handler

docmd.openform "frmCustomers", acnormal, "", "", , acnormal

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub btnproductiontype_click()
on error goto err_handler

docmd.openform "frmProductionType", acnormal, "", "", , acnormal

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub btnquoteaward_click()
on error goto err_handler

docmd.openform "frmQuoteAward", acnormal, "", "", , acnormal

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub btnrequesttype_click()
on error goto err_handler

docmd.openform "frmRequestType", acnormal, "", "", , acnormal

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub btnresults_click()
on error goto err_handler

docmd.openform "frmResults", acnormal, "", "", , acnormal

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub btnvolumetype_click()
on error goto err_handler

docmd.openform "frmVolumeType", acnormal, "", "", , acnormal

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub form_load()
on error goto err_handler

call settheme(me)

exit sub
err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.number)
end sub
