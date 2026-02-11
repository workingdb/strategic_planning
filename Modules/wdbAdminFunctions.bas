Option Compare Database
Option Explicit

Public Function StrQuoteReplace(strValue)
On Error GoTo Err_Handler

StrQuoteReplace = Replace(Nz(strValue, ""), "'", "''")

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "StrQuoteReplace", Err.Description, Err.Number)
End Function