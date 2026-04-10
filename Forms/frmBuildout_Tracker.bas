attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub cmddetails_click()
    docmd.openform "frmBuildout_details", , , "[registerId]=" & me.[recordid]
end sub

private sub cmdexposureinput_click()
    docmd.openform "frmBuildout_exposure", , , "[registerId]=" & me.recordid
end sub
