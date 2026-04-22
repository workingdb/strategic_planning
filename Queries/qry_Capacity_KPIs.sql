SELECT Count(*) AS TotalRequests, Avg(
        IIf(
            Not IsNull([RequestDateReal])
            And Not IsNull([ResponseDateReal])
            And [ResponseDateReal] >= [RequestDateReal],
            countWorkdays([RequestDateReal],[ResponseDateReal]),
            Null
        )
    ) AS AvgResponseTime, Sum(IIf(IsNull([ResponseDateReal]),1,0)) AS OpenRequests, Sum(IIf(Not IsNull([ResponseDateReal]),1,0)) AS RespondedRequests
FROM qry_Reports
WHERE [requestType]=1
    And [RequestDateReal] >= CDate(Forms!frmReportLauncher!txtStartDate)
    And [RequestDateReal] < DateAdd("d",1,CDate(Forms!frmReportLauncher!txtEndDate));

