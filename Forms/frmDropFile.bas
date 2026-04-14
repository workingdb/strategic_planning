attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

function savestratplandoc()

dim errortext as string
if nz(me.customname, "") = "" then errortext = "Please add a title"
if nz(me.dragdrop) = "" then errortext = "Please select a document to upload..."

if errortext <> "" then
    msgbox errortext, vbcritical, "Hold up"
    exit function
end if

dim fso as object
set fso = createobject("Scripting.FileSystemObject")

dim fileext as string, currentloc as string, fullpath as string, attchfullfilename as string, tempfold as string, newfile as string
currentloc = tempvars!dragdroplocation.value

'transfer file to temp location
tempfold = gettempfold
fileext = fso.getextensionname(currentloc)
newfile = tempfold & "tempUpload" & nowstring & "." & fileext

if folderexists(tempfold) = false then mkdir (tempfold)
call fso.copyfile(currentloc, newfile)

currentloc = newfile

me.attachname = me.customname & "-" & dmax("ID", "tblStratPlanAttachmentsSP") + 1

attchfullfilename = replace(me.attachname, " ", "_") & "." & fileext

dim db as dao.database
dim rsatt as dao.recordset
dim rsattchild as dao.recordset2
set db = currentdb
set rsatt = db.openrecordset("tblStratPlanAttachmentsSP", dbopendynaset)

rsatt.addnew
rsatt!filestatus = "Created"

rsatt.update
rsatt.movelast

rsatt.edit
set rsattchild = rsatt.fields("Attachments").value

rsattchild.addnew
dim fld as dao.field2
set fld = rsattchild.fields("FileData")
fld.loadfromfile (currentloc)
rsattchild.update

rsatt!uploadedby = environ("username")
rsatt!uploadeddate = now()
rsatt!attachname = me.attachname
rsatt!attachfullfilename = attchfullfilename
rsatt!filestatus = "Uploading"
rsatt!referenceid = me.tprojectid
rsatt!referencetable = me.tpartnumber
rsatt!documentlibrary = me.tdocumentlibary
rsatt.update

call registerstratplanupdates("tblStratPlanAttachmentsSP", me.tprojectid, "File Attachment", me.attachname, "Uploaded", me.tprojectid, me.name)

on error resume next
set fld = nothing
rsattchild.close: set rsattchild = nothing
rsatt.close: set rsatt = nothing
set db = nothing

docmd.close acform, "frmDropFile"
form_sfrmattachments.requery

end function

private sub btnsave_click()
on error goto err_handler

call savestratplandoc

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub dragdrop_afterupdate()
on error goto err_handler

me.refresh
dim strtext
dim dragdropval() as string, i as integer, item
dragdropval = split(me.dragdrop.value, "#")

i = 0
for each item in dragdropval
    if i = 0 then goto nextitem
    select case i
        case 0
            goto nextitem
        case 1
            strtext = item
        case else
            if item <> "" then strtext = strtext & "#" & item
    end select
nextitem:
    i = i + 1
next item

tempvars.add "dragDropLocation", strtext
me.dragdropview = strtext
me.getfocus.setfocus
if strtext = "" then
    msgbox "Didn't get that - please try again", vbinformation, "Oh no!"
    exit sub
end if
msgbox "Got it! Make sure details below are correct, then click Save + Close to Upload File", vbinformation, "Nice"

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub filepicker_click()
on error goto err_handler

dim strfile

with application.filedialog(msofiledialogopen)
    .title = "Choose a File"
    .allowmultiselect = false
    .show
    
    on error resume next
    strfile = .selecteditems(1)
end with

on error goto err_handler

if isnull(strfile) then exit sub

me.dragdrop = strfile
tempvars.add "dragDropLocation", strfile
me.dragdropview = strfile

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub form_load()
on error goto err_handler

call settheme(me)

me.dragdrop = ""
exit sub
err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.number)
end sub
