SELECT Count(*) AS TotalSurveys, Avg(
        IIf(
            Not IsNull([RequestDateReal])
            And Not IsNull([ResponseDateReal])
            And [ResponseDateReal] >= [RequestDateReal],
            countWorkdays([RequestDateReal],[ResponseDateReal]),
            Null
        )
    ) AS AvgDaysToComplete, Sum(IIf(IsNull([ResponseDateReal]),1,0)) AS OpenSurveys, Sum(IIf(Not IsNull([ResponseDateReal]),1,0)) AS CompletedSurveys
FROM qry_Reports
WHERE [requestType]<>1
    And [RequestDateReal] Between Forms!frmReportLauncher!txtStartDate And Forms!frmReportLauncher!txtEndDate;

