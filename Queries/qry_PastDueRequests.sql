SELECT CR.*, BusinessDueDate(CR.[RequestDate],CR.[Unit]) AS DueDate, DateDiff("d",BusinessDueDate(CR.[RequestDate],CR.[Unit]),Date()) AS DaysLate
FROM tblCapacityRequests AS CR
WHERE (((CR.RequestDate) Is Not Null) And ((CR.ResponseDate) Is Null) And ((UCase(Nz(CR.Unit,"")))<>"U7") And ((BusinessDueDate(CR.RequestDate,CR.Unit))<Date()));

