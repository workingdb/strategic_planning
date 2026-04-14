attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub form_load()
on error goto err_handler

call settheme(me)

exit sub
err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.numbe)
end sub


private sub openattachment_click()
on error goto err_handler

if me.filestatus = "Uploaded" then
    application.followhyperlink me.directlink
    call registerstratplanupdates("tblStratPlanAttachmentsSP", me.id, "File Attachment", me.attachname, "Opened", me.referenceid, me.name)
end if

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.numbe)
end sub

private sub remove_click()
on error goto err_handler

if me.filestatus <> "Uploaded" then
    msgbox "File must be fully uploaded in order to delete.", vbinformation, "Wait a second..."
    exit sub
end if

if msgbox("Are you sure?", vbyesno, "Please confirm") <> vbyes then exit sub

call registerstratplanupdates("tblStratPlanAttachmentsSP", me.id, "File Attachment", me.attachname, "Deleted", me.referenceid, me.name)

me.filestatus = "Deleting"
if me.dirty then me.dirty = false

me.requery
me.refresh

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
