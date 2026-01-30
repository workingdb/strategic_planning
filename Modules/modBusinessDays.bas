Option Compare Database
Option Explicit
 
'Null-safe: returns Null if reqDate is Null.
Public Function BusinessDueDate(ByVal reqDate As Variant, ByVal Unit As Variant) As Variant
    Dim n As Long
    Dim d As Date
    Dim u As String
 
    'If no request date, no due date
    If IsNull(reqDate) Then
        BusinessDueDate = Null
        Exit Function
    End If
 
    'Unit as safe string (handles Null)
    u = Trim$(Nz(Unit, ""))
 
    'Map unit -> SLA business days (edit to match your rules)
    Select Case u
        Case "U6": n = 5
        Case "U1": n = 5   'example
        Case "U2": n = 5   'example
        Case Else: n = 5   'default
    End Select
 
    'Start counting from NEXT business day (exclusive of request date)
    d = NextBusinessDay(DateValue(reqDate))
 
    'If SLA is 1 day, due is that next business day (no extra adds)
    If n <= 1 Then
        BusinessDueDate = d
        Exit Function
    End If
 
    'Add remaining business days
    BusinessDueDate = AddBusinessDays(d, n - 1)
End Function
 
Private Function NextBusinessDay(ByVal d As Date) As Date
    d = DateAdd("d", 1, d)
    Do While Weekday(d, vbMonday) > 5
        d = DateAdd("d", 1, d)
    Loop
    NextBusinessDay = d
End Function
 
Private Function AddBusinessDays(ByVal d As Date, ByVal n As Long) As Date
    Dim i As Long
    For i = 1 To n
        d = DateAdd("d", 1, d)
        Do While Weekday(d, vbMonday) > 5
            d = DateAdd("d", 1, d)
        Loop
    Next i
    AddBusinessDays = d
End Function