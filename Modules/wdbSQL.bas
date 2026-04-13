option compare database
option explicit

public function openrecordsetreadonly(conn as adodb.connection, sql as string) as adodb.recordset
    dim rs as adodb.recordset
    set rs = new adodb.recordset
    rs.open sql, conn, adopenforwardonly, adlockreadonly, adcmdtext
    set openrecordsetreadonly = rs
end function

public function openrecordsetreadwrite(conn as adodb.connection, sql as string) as adodb.recordset
    dim rs as adodb.recordset
    set rs = new adodb.recordset
    rs.open sql, conn, adopenkeyset, adlockoptimistic, adcmdtext
    set openrecordsetreadwrite = rs
end function

public function sqlmin(conn as adodb.connection, fieldname as string, tablename as string, _
    optional whereclause as string = "", optional defaultvalue as variant = null) as variant

    dim rs as adodb.recordset
    dim sql as string

    sql = "SELECT MIN(" & fieldname & ") FROM " & tablename
    if trim$(whereclause) <> "" then sql = sql & " WHERE " & whereclause

    set rs = openrecordsetreadonly(conn, sql)

    if rs.eof or isnull(rs.fields(0).value) then
        sqlmin = defaultvalue
    else
        sqlmin = rs.fields(0).value
    end if

    rs.close
    set rs = nothing
end function

public function sqlcount(conn as adodb.connection, tablename as string, _
    optional whereclause as string = "") as long

    dim rs as adodb.recordset
    dim sql as string

    sql = "SELECT COUNT(*) FROM " & tablename
    if trim$(whereclause) <> "" then sql = sql & " WHERE " & whereclause

    set rs = openrecordsetreadonly(conn, sql)
    sqlcount = clng(nz(rs.fields(0).value, 0))

    rs.close
    set rs = nothing
end function

public function sqllookup(conn as adodb.connection, fieldname as string, tablename as string, _
    optional whereclause as string = "", optional defaultvalue as variant = null) as variant

    dim rs as adodb.recordset
    dim sql as string

    sql = "SELECT " & fieldname & " FROM " & tablename
    if trim$(whereclause) <> "" then sql = sql & " WHERE " & whereclause

    set rs = openrecordsetreadonly(conn, sql)

    if rs.eof or isnull(rs.fields(0).value) then
        sqllookup = defaultvalue
    else
        sqllookup = rs.fields(0).value
    end if

    rs.close
    set rs = nothing
end function

function getbestsqldriver() as string
on error goto err_handler

    dim shell as object
    dim driverlist as variant, driver as variant
    dim regpath as string
    
    set shell = createobject("WScript.Shell")
    ' list drivers in order of preference (newest first)
    driverlist = array("ODBC Driver 18 for SQL Server", _
                       "ODBC Driver 17 for SQL Server", _
                       "ODBC Driver 13 for SQL Server", _
                       "SQL Server Native Client 11.0", _
                       "SQL Server")
    
    on error resume next
    for each driver in driverlist
        ' check registry for the driver entry
        if regkeyexists("HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\" & driver & "\Driver") then
            getbestsqldriver = driver
            exit function
        end if
    next driver
    on error goto err_handler
    
    getbestsqldriver = "" ' no driver found
    
exit function
err_handler:
    call handleerror("wdbLinkage", "GetBestSQLDriver", err.description, err.number)
end function

function regkeyexists(regpath as string) as boolean
    dim shell as object
    set shell = createobject("WScript.Shell")
    on error resume next
    shell.regread regpath
    regkeyexists = (err.number = 0)
    on error goto 0
end function

function relinksqltables(optional returnstringonly as boolean = false) as string
on error goto err_handler

    dim db as dao.database
    dim tdf as dao.tabledef
    dim qdf as dao.querydef
    dim strdriver as string
    dim strconn as string
    
    strdriver = getbestsqldriver()
    if strdriver = "" then exit function
    
    set db = currentdb
    ' base odbc connection string
    strconn = "ODBC;DRIVER=" & strdriver & _
              ";SERVER=ITI-SQL\ITISQL" & _
              ";Trusted_Connection=Yes" & _
              ";APP=Microsoft Office" & _
              ";DATABASE=workingdb;"
    
    if returnstringonly then
        relinksqltables = strconn
        exit function
    end if
    
    ' loop through all tables and update odbc links
    for each tdf in db.tabledefs
        ' only relink tables that already have an odbc connection string
        if instr(1, tdf.connect, "SERVER=ITI-SQL") then
            tdf.connect = strconn
            tdf.refreshlink
        end if
    next tdf
    
    ' relink pass-through queries
    for each qdf in db.querydefs
        if qdf.type = dbqsqlpassthrough then
            if instr(1, qdf.connect, "SERVER=ITI-SQL") > 0 then
                qdf.connect = strconn
            end if
        end if
    next qdf
    
relinksqltables = strconn
    
exit function
err_handler:
    call handleerror("wdbLinkage", "RelinkSQLTables", err.description, err.number)
end function
