SELECT qryOnHand.PartNumber, Sum(qryOnHand.OnHandQty) AS TotalQty
FROM qryOnHand
GROUP BY qryOnHand.PartNumber;

