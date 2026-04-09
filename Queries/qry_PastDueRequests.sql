SELECT CR.*, BusinessDueDate(CR.[RequestDate],CR.[Unit]) AS DueDate, DateDiff("d",BusinessDueDate(CR.[RequestDate],CR.[Unit]),Date()) AS DaysLate
FROM tblCapacityRequests AS CR
WHERE (((CR.RequestDate) Is Not Null) AND (([CR].[ResponseDate]) Is Null) AND ((UCase(Nz([CR].[Unit],"")))<>"U7") AND ((BusinessDueDate([CR].[RequestDate],[CR].[Unit]))<Date()));

