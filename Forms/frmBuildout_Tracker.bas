attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub cmddetails_click()
    docmd.openform "frmBuildout_details", , , "recordId=" & me.[recordid]
end sub
