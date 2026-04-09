attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub form_load()
on error goto err_handler

call settheme(me)

docmd.applyfilter , "[tblPermissions].User = '" & environ("username") & "'"

exit sub
err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.number)
end sub
