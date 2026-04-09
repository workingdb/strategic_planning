attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub imguser_click()
on error goto err_handler

docmd.openform "frmUserProfile", , , "user = '" & me.updatedby & "'"

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
