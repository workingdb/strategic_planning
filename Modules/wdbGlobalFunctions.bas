option compare database
option explicit

public function addworkdays(dateinput as date, daystoadd as long) as date
on error goto err_handler

dim db as database
set db = currentdb()
dim i as long, testdate as date, daysleft as long, rsholidays as recordset, intdirection
testdate = dateinput
daysleft = abs(daystoadd)
intdirection = 1
if daystoadd < 0 then intdirection = -1

set rsholidays = db.openrecordset("tblHolidays", dbopensnapshot)

do while daysleft > 0
    testdate = testdate + intdirection
    if weekday(testdate) = 7 or weekday(testdate) = 1 then ' if weekend -> skip
        testdate = testdate + intdirection
        goto skipdate
    end if
    
    rsholidays.findfirst "holidayDate = #" & testdate & "#"
    if not rsholidays.nomatch then goto skipdate ' if holiday -> skip to next da

     daysleft = daysleft - 1
skipdate:
loop

addworkdays = testdate

on error resume next
rsholidays.close
set rsholidays = nothing
set db = nothing

exit function
err_handler:
    call handleerror("wdbGlobalFunctions", "addWorkdays", err.description, err.number)
end function

public function countworkdays(olddate as date, newdate as date) as long
on error goto err_handler

dim total, sunday, saturday, weekdays, holidays

total = datediff("d", [olddate], [newdate], vbsunday)
sunday = datediff("ww", [olddate], [newdate], 1)
saturday = datediff("ww", [olddate], [newdate], 7)
holidays = dcount("recordId", "tblHolidays", "holidayDate > #" & olddate - 1 & "# AND holidayDate < #" & newdate & "#")
countworkdays = total - sunday - saturday - holidays

exit function
err_handler:
    call handleerror("wdbGlobalFunctions", "countWorkdays", err.description, err.number)
end function

public function snackbox(stype as string, stitle as string, smessage as string, refform as string, optional centerbool as boolean = false, optional autoclose as boolean = true)
on error goto err_handler

tempvars.add "snackType", stype
tempvars.add "snackTitle", stitle
tempvars.add "snackMessage", smessage
tempvars.add "snackAutoClose", autoclose

if centerbool then
    tempvars.add "snackCenter", "True"
    tempvars.add "snackLeft", forms(refform).windowleft + forms(refform).windowwidth / 2 - 3393
    tempvars.add "snackTop", forms(refform).windowtop + forms(refform).windowheight / 2 - 500
else
    tempvars.add "snackCenter", "False"
    tempvars.add "snackLeft", forms(refform).windowleft + 200
    tempvars.add "snackTop", forms(refform).windowtop + forms(refform).windowheight - 1250
end if

docmd.openform "frmSnack"

exit function
err_handler:
    call handleerror("wdbGlobalFunctions", "snackBox", err.description, err.number)
end function

function structurechange()

dim conn as adodb.connection
dim rscap as adodb.recordset, rsparts as adodb.recordset
dim strsql as string

set conn = new adodb.connection
conn.connectionstring = "DRIVER=ODBC Driver 17 for SQL Server;SERVER=ITI-SQL\ITISQL;Trusted_Connection=Yes;APP=Microsoft Office;DATABASE=workingdb;"
conn.open

set rscap = openrecordsetreadonly(conn, "SELECT * FROM tblCapacityRequests")

do while not rscap.eof
    strsql = "INSERT INTO dbo.tblCapacityRequest_partnumbers (requestId,partNumber,unitId,productionType,tonnage,ppv,volumeType,volume,volumeTiming,capacityResults,responseDate,planner,quoteStatus) VALUES (" & _
            rscap!recordid & ",'" & _
            nz(rscap!partnumber, "Null") & "'," & _
            nz(rscap!unit, "Null") & "," & _
            nz(rscap!productiontype, "Null") & "," & _
            nz(rscap!tonnage, "Null") & "," & _
            nz(rscap!ppv, "Null") & "," & _
            nz(rscap!volumetype, "Null") & "," & _
            nz(rscap!volume, "Null") & "," & _
            nz(rscap!volumetiming, "Null") & "," & _
            nz(rscap!capacityresults, "Null") & ",'" & _
            format$(rscap!responsedate, "yyyy-mm-dd hh:nn:ss") & "'," & _
            nz(rscap!planner, "Null") & "," & _
            nz(rscap!quote, "Null") & ");"
    conn.execute strsql
    rscap.movenext
loop


rscap.close
set rscap = nothing
conn.close
set conn = nothing

end function

function emailcontentgen(subject as string, title as string, subtitle as string, primarymessage as string, detail1 as string, detail2 as string, detail3 as string, optional appname as string = "", optional appid as string = "") as string
on error goto err_handler

if appid <> "" then
    primarymessage = "<a href = ""\\data\mdbdata\WorkingDB\build\workingdb_commands\openNotification.vbs"">" & primarymessage & "</a>"
end if

emailcontentgen = subject & "," & title & "," & subtitle & "," & primarymessage & "," & detail1 & "," & detail2 & "," & detail3 & "," & appname & "," & appid

    exit function
err_handler:
    call handleerror("wdbGlobalFunctions", "emailContentGen", err.description, err.number)
end function

function findcapreqpns(requestid as long, optional returnresponse as boolean = false) as string
on error goto err_handler

findcapreqpns = ""

dim rs as adodb.recordset
dim conn as adodb.connection

set conn = currentproject.connection

if returnresponse then
    set rs = openrecordsetreadonly(conn, "SELECT partNumber, capacityResults FROM tblCapacityRequest_partnumbers WHERE requestId = " & requestid)
    
    do while not rs.eof
        if nz(rs!capacityresults, 0) = 0 then exit function
        findcapreqpns = findcapreqpns & rs!partnumber & "|" & sqllookup(conn, "results", "tblDropDowns_StrategicPlanning", "recordId = " & nz(rs!capacityresults, 0)) & ","
        rs.movenext
    loop
    
    if len(findcapreqpns) = 0 then goto clean_exit
    findcapreqpns = left(findcapreqpns, len(findcapreqpns) - 1)
else
    set rs = openrecordsetreadonly(conn, "SELECT partNumber FROM tblCapacityRequest_partnumbers WHERE requestId = " & requestid)
    
    do while not rs.eof
        findcapreqpns = findcapreqpns & rs!partnumber & ","
        rs.movenext
    loop
    
    if len(findcapreqpns) = 0 then goto clean_exit
    findcapreqpns = left(findcapreqpns, len(findcapreqpns) - 1)
end if



clean_exit:
on error resume next
if not rs is nothing then
    if rs.state = adstateopen then rs.close
end if
set rs = nothing
set conn = nothing

exit function
err_handler:
    call handleerror("wdbGlobalFunctions", "setSplashLoading", err.description, err.number)
    goto clean_exit
end function

function sendnotification(sendto as string, nottype as integer, notpriority as integer, desc as string, emailcontent as string, optional appname as string = "", optional appid as variant = "", optional multiemail as boolean = false, optional customemail as boolean = false) as boolean
sendnotification = true

on error goto err_handler

dim conn as adodb.connection
set conn = currentproject.connection

'has this person been notified about this thing today already?
dim rsnotifications as adodb.recordset
set rsnotifications = openrecordsetreadonly(conn, "SELECT * from tblNotificationsSP WHERE recipientUser = '" & replace(sendto, "'", "''") & "' AND notificationDescription = '" & strquotereplace(desc) & "' AND sentDate > #" & date - 1 & "#")
if not rsnotifications.eof then
    if rsnotifications!notificationtype = 1 then
        dim msgtxt as string
        if rsnotifications!senderuser = environ("username") then
            msgtxt = "You already nudged this person today"
        else
            msgtxt = sendto & " has already been nudged about this today by " & rsnotifications!senderuser & ". Let's wait until tomorrow to nudge them again."
        end if
        msgbox msgtxt, vbinformation, "Hold on a minute..."
        sendnotification = false
        rsnotifications.close
        set rsnotifications = nothing
        set conn = nothing
        exit function
    end if
end if
rsnotifications.close
set rsnotifications = nothing

dim stremail
if customemail = false then
    dim item, sendtoarr() as string
    if multiemail then
        sendtoarr = split(sendto, ",")
        stremail = ""
        for each item in sendtoarr
            if item = "" then goto nextitem
            stremail = stremail & getemail(cstr(item)) & ";"
nextitem:
        next item
        if stremail = "" then exit function
        stremail = left(stremail, len(stremail) - 1)
    else
        stremail = getemail(sendto)
    end if
else
    stremail = sendto
    sendto = split(sendto, "@")(0)
end if

dim strsql as string
strsql = "INSERT INTO tblNotificationsSP (recipientUser, recipientEmail, senderUser, senderEmail, sentDate, " & _
         "notificationType, notificationPriority, notificationDescription, appName, appId, emailContent) VALUES (" & _
         "'" & replace(sendto, "'", "''") & "', " & _
         "'" & replace(cstr(nz(stremail, "")), "'", "''") & "', " & _
         "'" & replace(environ("username"), "'", "''") & "', " & _
         "'" & replace(getemail(environ("username")), "'", "''") & "', " & _
         "#" & format$(now(), "yyyy-mm-dd hh:nn:ss") & "#, " & _
         nottype & ", " & notpriority & ", " & _
         "'" & strquotereplace(desc) & "', " & _
         "'" & replace(nz(appname, ""), "'", "''") & "', " & _
         "'" & replace(cstr(nz(appid, "")), "'", "''") & "', " & _
         "'" & strquotereplace(emailcontent) & "')"
conn.execute strsql

set conn = nothing

exit function
err_handler:
sendnotification = false
    call handleerror("wdbGlobalFunctions", "sendNotification", err.description, err.number)
end function

function getemail(username as string) as string
on error goto err_handler

getemail = ""
on error goto tryoracle
dim db as database
set db = currentdb()
dim rspermissions as recordset
set rspermissions = db.openrecordset("SELECT * from tblPermissions WHERE user = '" & username & "'", dbopensnapshot)
getemail = nz(rspermissions!useremail, "")
rspermissions.close
set rspermissions = nothing

goto exitfunc

tryoracle:
dim rsemployee as recordset
set rsemployee = db.openrecordset("SELECT FIRST_NAME, LAST_NAME, EMAIL_ADDRESS FROM APPS_XXCUS_USER_EMPLOYEES_V WHERE USER_NAME = '" & strconv(username, vbuppercase) & "'", dbopensnapshot)
getemail = nz(rsemployee!email_address, "")
rsemployee.close
set rsemployee = nothing

exitfunc:
set db = nothing

exit function
err_handler:
    call handleerror("wdbGlobalFunctions", "getEmail", err.description, err.number)
end function

function generatehtml(title as string, subtitle as string, primarymessage as string, _
        detail1 as string, detail2 as string, detail3 as string, _
        optional link as string = "", _
        optional addlines as boolean = false, _
        optional appname as string = "", _
        optional appid as string = "") as string
        
on error goto err_handler

dim tblheading as string, tblfooter as string, strhtmlbody as string

if link <> "" then
    primarymessage = "<a href = '" & link & "'>" & primarymessage & "</a>"
elseif appid <> "" then
    primarymessage = "<a href = ""\\data\mdbdata\WorkingDB\build\workingdb_commands\openNotification.vbs"">" & primarymessage & "</a>"
end if

tblheading = "<table style=""width: 100%; margin: 0 auto; padding: 2em 3em; text-align: center; background-color: #fafafa;"">" & _
                            "<tbody>" & _
                                "<tr><td><h2 style=""color: #414141; font-size: 28px; margin-top: 0;"">" & title & "</h2></td></tr>" & _
                                "<tr><td><p style=""color: rgb(73, 73, 73);"">" & subtitle & "</p></td></tr>" & _
                                 "<tr><td><table style=""padding: 1em; text-align: center;"">" & _
                                     "<tr><td style=""padding: 1em 1.5em; background: #FF6B00; "">" & primarymessage & "</td></tr>" & _
                                "</table></td></tr>" & _
                            "</tbody>" & _
                        "</table>"
                        
tblfooter = "<table style=""width: 100%; margin: 0 auto; padding: 3em; background: #2b2b2b; color: rgba(255,255,255,.5);"">" & _
                        "<tbody>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: 1em; color: #c9c9c9;"">Details</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;"">" & detail1 & "</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;"">" & detail2 & "</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em 1em 2em;"">" & detail3 & "</td></tr>" & _
                        "</tbody>" & _
                    "</table>"
                    
dim addstuff as string
addstuff = ""
if addlines then
    addstuff = "<table style=""max-width: 600px; margin: 0 auto; padding: 3em; background: #eaeaea; color: rgba(255,255,255,.5);"">" & _
        "<tr style=""border-collapse: collapse;""><td style=""padding: 1em;"">Extra Notes: type here...</td></tr></table>"
end if

strhtmlbody = "" & _
"<!DOCTYPE html><html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">" & _
    "<head><meta charset=""utf-8""><title>Working DB Notification</title></head>" & _
    "<body style=""margin: 0 auto; Font-family: 'Montserrat', sans-serif; font-weight: 400; font-size: 15px; line-height: 1.8;"">" & addstuff & _
        "<table style=""max-width: 600px; margin: 0 auto; text-align: center; "">" & _
            "<tbody>" & _
                "<tr><td>" & tblheading & "</td></tr>" & _
                "<tr><td>" & tblfooter & "</td></tr>" & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">AppName:[" & appname & "], AppId:[" & appid & "]</p></td></tr>" & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email was created by  &copy; workingDB</p></td></tr>" & _
            "</tbody>" & _
        "</table>" & _
    "</body>" & _
"</html>"

generatehtml = strhtmlbody

exit function
err_handler:
    call handleerror("wdbGlobalFunctions", "generateHTML", err.description, err.number)
end function

public function wdbemail(byval strto as string, byval strcc as string, byval strsubject as string, body as string) as boolean
on error goto err_handler
wdbemail = true
    
dim objemail as object

set objemail = createobject("outlook.Application")
set objemail = objemail.createitem(0)

with objemail
    .to = strto
    .cc = strcc
    .subject = strsubject
    .htmlbody = body
    .display
end with

set objemail = nothing
    
exit function
err_handler:
wdbemail = false
    call handleerror("wdbGlobalFunctions", "wdbEmail", err.description, err.number)
end function

function setsplashloading(label as string)
on error goto err_handler

if isnull(tempvars!loadamount) then exit function
tempvars.add "loadAmount", tempvars!loadamount + 1
form_frmsplash.lnloading.width = (tempvars!loadamount / 12) * tempvars!loadwd
form_frmsplash.lblloading.caption = label
form_frmsplash.repaint

exit function
err_handler:
    call handleerror("wdbGlobalFunctions", "setSplashLoading", err.description, err.number)
end function

function userdata(data as string, optional specificuser as string = "") as string
    on error goto err_handler

    if specificuser = "" then specificuser = environ("username")

    dim conn as adodb.connection
    dim rs as new adodb.recordset
    dim strsql as string
    
    set conn = currentproject.connection
    
    ' using brackets around the variable [data] and reserved word [user]
    strsql = "SELECT [" & data & "] FROM tblPermissions WHERE [User] = '" & replace(specificuser, "'", "''") & "'"
    
    rs.open strsql, conn, adopenforwardonly, adlockreadonly
    
    if not rs.eof then
        userdata = nz(rs.fields(0).value, "")
    else
        userdata = ""
    end if

cleanup:
    if rs.state = adstateopen then rs.close
    set rs = nothing
    set conn = nothing
    exit function

err_handler:
    call handleerror("wdbGlobalFunctions", "userData", err.description, err.number)
    resume cleanup
end function

function dbexecute(sql as string)
on error goto err_handler

dim conn as adodb.connection
set conn = currentproject.connection

conn.execute sql

set conn = nothing

exit function
err_handler:
    call handleerror("wdbGlobalFunctions", "dbExecute", err.description, err.number, sql)
end function

public sub registerstratplanupdates( _
    byval table as string, _
    byval id as variant, _
    byval column as string, _
    byval oldval as variant, _
    byval newval as variant, _
    byval referenceid as string, _
    byval formname as string, _
    optional byval tag0 as variant = "")

    on error goto err_handler

    dim cmd as adodb.command
    dim oldtext as string
    dim newtext as string
    dim tagtext as string

    ' normalize dates
    if vartype(oldval) = vbdate then oldval = format$(oldval, "mm/dd/yyyy")
    if vartype(newval) = vbdate then newval = format$(newval, "mm/dd/yyyy")

    ' normalize text values
    oldtext = left$(strquotereplace(cstr(nz(oldval, ""))), 255)
    newtext = left$(strquotereplace(cstr(nz(newval, ""))), 255)
    tagtext = left$(strquotereplace(cstr(nz(tag0, ""))), 255)

    ' normalize blank id to null
    if nz(id, "") = "" then id = null

    set cmd = new adodb.command

    with cmd
        .activeconnection = currentproject.connection
        .commandtype = adcmdtext
        .commandtext = _
            "INSERT INTO tblStratPlan_UpdateTracking (" & _
            "tableName, tableRecordId, updatedBy, updatedDate, columnName, " & _
            "previousData, newData, referenceId, formName, dataTag0) " & _
            "VALUES (?, ?, ?, '" & format$(now(), "yyyy-mm-dd\Thh:nn:ss") & "', ?, ?, ?, ?, ?, ?)"

        .parameters.append .createparameter("pTableName", advarchar, adparaminput, 100, table)

        if isnull(id) then
            .parameters.append .createparameter("pTableRecordId", adinteger, adparaminput, , null)
        else
            .parameters.append .createparameter("pTableRecordId", adinteger, adparaminput, , id)
        end if

        .parameters.append .createparameter("pUpdatedBy", advarchar, adparaminput, 55, environ$("username"))
        .parameters.append .createparameter("pColumnName", advarchar, adparaminput, 100, column)
        .parameters.append .createparameter("pPreviousData", advarchar, adparaminput, 255, oldtext)
        .parameters.append .createparameter("pNewData", advarchar, adparaminput, 255, newtext)
        .parameters.append .createparameter("pReferenceId", adinteger, adparaminput, , referenceid)
        .parameters.append .createparameter("pFormName", advarchar, adparaminput, 55, strquotereplace(formname))
        .parameters.append .createparameter("pDataTag0", advarchar, adparaminput, 55, tagtext)

        .execute , , adexecutenorecords
    end with

cleanexit:
    set cmd = nothing
    exit sub

err_handler:
    call handleerror("wdbGlobalFunctions", "registerStratPlanUpdates", err.description, err.number)
    resume cleanexit

end sub

function logclick(modname as string, formname as string, optional datatag0 = "")
    on error goto err_handler

    ' 1. check if analytics are enabled
    if nz(dlookup("paramVal", "tblDBinfoBE", "parameter = 'recordAnalytics'"), "False") = "False" then exit function

    dim conn as new adodb.connection
    conn.connectionstring = replace(relinksqltables(true), "ODBC;", "")
    conn.open
    
    dim strsql as string
    
    ' 2. build the sql string
    strsql = "INSERT INTO tblAnalytics ([module], [form], [username], [dateused], [datatag0], [datatag1]) " & _
             "VALUES (" & _
             "'" & replace(modname, "'", "''") & "', " & _
             "'" & replace(formname, "'", "''") & "', " & _
             "'" & environ("username") & "', " & _
             "'" & format(now(), "yyyy-mm-dd hh:nn:ss") & "', " & _
             "'" & replace(nz(datatag0, ""), "'", "''") & "', " & _
             "'" & nz(tempvars!wdbversion, "") & "')"


    ' 3. execute
    conn.execute strsql

cleanup:
    set conn = nothing
    exit function

err_handler:
    call handleerror("wdbGlobalFunctions", "logClick", err.description, err.number)
    resume cleanup
end function

public function strquotereplace(strvalue)
on error goto err_handler

strquotereplace = replace(nz(strvalue, ""), "'", "''")

exit function
err_handler:
    call handleerror("wdbGlobalFunctions", "StrQuoteReplace", err.description, err.number)
end function
