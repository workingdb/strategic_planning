attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub backbtn_click()
on error goto err_handler
docmd.close
exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub form_load()
on error goto err_handler
me.shortcutmenu = false

me.progbar.width = 6800

me.lbltitle.tag = tempvars!snacktype 'set icon
me.lbltitle.caption = tempvars!snacktitle 'set title

select case tempvars!snacktype 'set progress bar color
    case "success"
        me.progbar.bordercolor = rgb(140, 150, 100)
        me.bxoutline.bordercolor = rgb(140, 150, 100)
    case "error"
        me.progbar.bordercolor = rgb(150, 100, 100)
        me.bxoutline.bordercolor = rgb(150, 100, 100)
    case "info"
        me.progbar.bordercolor = rgb(110, 120, 130)
        me.bxoutline.bordercolor = rgb(110, 120, 130)
end select

me.lblmessage = tempvars!snackmessage
if len(me.lblmessage) < 48 then me.lblmessage.topmargin = 72 'if it's only one line, add top margin to the text box
if len(me.lblmessage) > 112 then me.lblmessage.fontsize = 8 'if it's longer than two lines (ish)

me.move tempvars!snackleft, tempvars!snacktop 'set position

dim arr() as string, subam
arr = vba.split(tempvars!snackmessage, " ")

subam = 2 * 274.44 / ((ubound(arr) - lbound(arr) + 1)) '250 wpm / 60 seconds

if subam < 70 then subam = 50
if subam > 200 then subam = 200

tempvars.add "snackSubtract", subam

exit sub
err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.number)
end sub

private sub form_timer()

if tempvars!snackautoclose = false then exit sub

if me.progbar.width < tempvars!snacksubtract then
    docmd.close
    exit sub
end if
me.progbar.width = me.progbar.width - nz(tempvars!snacksubtract, 0)

end sub
