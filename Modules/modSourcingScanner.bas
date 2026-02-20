Option Compare Database
Option Explicit
 
'How long a claim lasts before another session can re-claim it (minutes)
Private Const CLAIM_TTL_MIN As Long = 30
 
Public Sub ScanAndRoutePendingRequests()
    On Error GoTo ErrHandler
 
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
 
    Set db = CurrentDb
 
    'Pick up items that are NOT notified yet and NOT currently claimed (or claim expired)
    sql = _
        "SELECT RecordID " & _
        "FROM tblCapacityRequests " & _
        "WHERE (Nz(SourcingNotifiedFlag,0)=0) " & _
        "  AND (SourcingNotifiedOn Is Null) " & _
        "  AND (SourcingProcessingOn Is Null " & _
        "  AND requestDate >= #2/18/2026# " & _
        "       OR DateDiff('n', SourcingProcessingOn, Now()) >= " & CLAIM_TTL_MIN & ") " & _
        "ORDER BY RecordID;"
 
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
 
    Do While Not rs.EOF
        Call TryRouteOne(rs!RecordID)
        rs.MoveNext
    Loop
 
Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
 
ErrHandler:
    'Optional: Debug.Print Err.Number, Err.Description
    Resume Cleanup
End Sub
 
Private Sub TryRouteOne(ByVal recId As Long)
    On Error GoTo ErrHandler
 
    Dim db As DAO.Database
    Set db = CurrentDb
 
    '1) Claim the record atomically (prevents duplicates across multiple open sessions)
    '   Only claim if still un-notified and claim is open/expired.
    Dim sqlClaim As String
    sqlClaim = _
        "UPDATE tblCapacityRequests " & _
        "SET SourcingProcessingOn=Now(), " & _
        "    SourcingProcessingBy='" & Replace(GetSessionUser(), "'", "''") & "' " & _
        "WHERE RecordID=" & recId & " " & _
        "  AND Nz(SourcingNotifiedFlag,0)=0 " & _
        "  AND SourcingNotifiedOn Is Null " & _
        "  AND (SourcingProcessingOn Is Null " & _
        "       OR DateDiff('n', SourcingProcessingOn, Now()) >= " & CLAIM_TTL_MIN & ");"
 
    db.Execute sqlClaim, dbFailOnError
 
    'If no rows affected, someone else claimed it first or it’s already processed
    If db.RecordsAffected = 0 Then Exit Sub
 
    '2) Run YOUR existing gate logic + email + stamps
    '   IMPORTANT: This should NOT open/filter any visible forms.
    Call HandleSourcingRoutingById(recId)
 
    Exit Sub
 
ErrHandler:
    'If routing fails, release the claim so it can be retried
    On Error Resume Next
    CurrentDb.Execute _
        "UPDATE tblCapacityRequests " & _
        "SET SourcingProcessingOn=Null, SourcingProcessingBy=Null " & _
        "WHERE RecordID=" & recId & ";"
End Sub
 
Private Function GetSessionUser() As String
    'Pick your favorite identity source
    GetSessionUser = Environ$("USERNAME")
End Function
 