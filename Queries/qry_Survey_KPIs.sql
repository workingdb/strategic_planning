SELECT Count([ID]) AS TotalSurveys, Avg(IIf([DateCompleted] Is Null, Null, DateDiff("d",[DateRequested],[DateCompleted]))) AS AvgDaysToComplete, Sum(IIf([DateCompleted] Is Null,1,0)) AS OpenSurveys, Sum(IIf([DateCompleted] Is Not Null,1,0)) AS CompletedSurveys
FROM tblSurveys
WHERE [DateRequested] >= TempVars!StartDate
  AND [DateRequested] < DateAdd("d", 1, TempVars!EndDate);

