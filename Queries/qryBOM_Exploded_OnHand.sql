SELECT E.*, O.TotalQty AS OnHandQty
FROM tblBOM_Exploded AS E LEFT JOIN qryOnHand_Total AS O ON E.ComponentItem = O.PartNumber;

