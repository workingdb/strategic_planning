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
