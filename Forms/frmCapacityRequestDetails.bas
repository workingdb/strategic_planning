attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

function trackupdate()
on error goto err_handler

if isnull(me.recordid) then exit function
call registerstratplanupdates("tblCapacityRequests", me.recordid, me.activecontrol.name, me.activecontrol.oldvalue, me.activecontrol, me.recordid, me.name)

exit function
err_handler:
    call handleerror(me.name, "trackUpdate", err.description, err.number)
end function

private sub addfile_click()
on error goto err_handler

if isnull(me.recordid) then
    msgbox "Please put more info in so there's a RecordID on the top, then try again.", vbinformation, "Woops"
    exit sub
end if

docmd.openform "frmDropFile"
'custom title stuff
form_frmdropfile.customname.visible = true
form_frmdropfile.customname.locked = false
form_frmdropfile.tdocumentlibary = "WDB_Capacity_Requests"
form_frmdropfile.label62.visible = true
form_frmdropfile.command63.visible = true

form_frmdropfile.lbldoccategory.caption = "Strategic Planning Document"
form_frmdropfile.tprojectid = me.recordid
form_frmdropfile.tpartnumber = "tblCapacityRequests"
form_frmdropfile.documenttype = 30
form_frmdropfile.documenttype.locked = true
form_frmdropfile.documenttype.visible = false
form_frmdropfile.doctypecard.visible = false
form_frmdropfile.label51.visible = false
form_frmdropfile.box57.visible = false
form_frmdropfile.command58.visible = false

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

private sub find_click()
on error goto err_handler

    on error resume next
    docmd.gotocontrol screen.previouscontrol.name
    err.clear
    docmd.runcommand accmdfind
    if (macroerror <> 0) then
        msgbox macroerror.description, vbokonly, ""
    end if

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub new_click()
on error goto err_handler

    on error resume next
    docmd.gotorecord , "", acnewrec
    if (macroerror <> 0) then
        msgbox macroerror.description, vbokonly, ""
    end if

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub requesttype_afterupdate()
on error goto err_handler

call trackupdate

me.surveypartcount.visible = me.requesttype = 2

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub trash_click()
on error goto err_handler

if msgbox("Are you sure you want to delete this request?", vbyesno, "Please confirm") = vbyes then
    if nz(me.recordid, 0) <> 0 then call registerstratplanupdates("tblCapacityRequestDetails", me.recordid, "Request", "", "Deleted", me.recordid, me.name)
    dbexecute ("DELETE FROM tblCapacityRequests WHERE [recordId] = " & me.recordid)
    tempvars.add "reqCapDelete", "True"
    docmd.close
    if currentproject.allforms("frmCapacityRequestTracker").isloaded then form_frmcapacityrequesttracker.requery
end if

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub copy_click()
on error goto err_handler

    on error resume next
    docmd.runcommand accmdselectrecord
    if (macroerror = 0) then
        docmd.runcommand accmdcopy
    end if
    if (macroerror = 0) then
        docmd.runcommand accmdrecordsgotonew
    end if
    if (macroerror = 0) then
        docmd.runcommand accmdselectrecord
    end if
    if (macroerror = 0) then
        docmd.runcommand accmdpaste
    end if
    if (macroerror <> 0) then
        msgbox macroerror.description, vbokonly, ""
    end if

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub mailreport_click()
on error goto err_handler

dim partnums as string
partnums = findcapreqpns(me.recordid, true)

dim pnsplit() as string, item, partnumfinal as string
pnsplit = split(partnums, ",")
partnumfinal = ""

for each item in pnsplit
    partnumfinal = partnumfinal & "PN: " & split(item, "|")(0) & " - Response: " & split(item, "|")(1)
next item

dim body as string
body = emailcontentgen("Capacity Request Results", _
    me.requesttype.column(1) & " Results", _
    "Notes: " & replace(me.notes, ",", ";"), _
     partnumfinal, _
    "Requested: " & cstr(date) & " by: " & me.requestor.column(1), _
    "Vehicle: " & me.program.column(1), _
    "Program: " & me.program.column(0))
call registerstratplanupdates("tblCapacityRequestDetails", me.recordid, "Results", "", "Results Sent to Requestor", me.recordid, me.name)
if sendnotification(me.requestor.column(2), 6, 2, "Capacity Request Results", body) then
    call snackbox("success", "Well Done!", me.requestor.column(2) & " Notified!", me.name)
end if

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
