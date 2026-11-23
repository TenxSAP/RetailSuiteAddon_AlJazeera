 
CREATE PROCEDURE "Check_GR_Price_Fluctuation"(IN DocEntry INT)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER
AS
BEGIN
Select *
,Case When "MarginDiff">="MarginTolerance" Then 'Price Increase' Else 'Price Reduction' END "FluctuationType"
FROM
(Select *
,Round("Price"-"SellingPrice",3) As "NewMarginAmount"
,CASE 
        WHEN IFNULL("SellingPrice", 0) = 0 THEN 0
        ELSE ROUND(
                (ROUND("Price" - "SellingPrice", 3) 
                / "SellingPrice") * 100
              , 3)
      END AS "NewMarginPercentage"
,Round("Price"-"SellingPrice",3) As "MarginDiff"
From (
SELECT         
H."DocEntry"
,H."DocNum"
,H."DocDate"
,L."LineNum"
,L."ItemCode"
,L."Dscription"
,L."UomCode"
,L."UomEntry"
,L."Price"
,L."LineNum"
,H."DocDate"
,H."UserSign"
,(SELECT U."USER_CODE" FROM "OUSR" U WHERE U."USERID" = H."UserSign") AS "UserCode"
,B."U_PricelistNo"
,H."BPLId"
,B."BPLName" AS "BranchCode"
,L."U_suppprice" as "SupplierCost"
,IFNULL((Select "TNX_SELLING_PRICE"(L."ItemCode",L."UomEntry",H."BPLId") From Dummy),0) As "SellingPrice"
,(Select "U_MTlce" From OADM) As "MarginTolerance"
FROM "OIGN" H
INNER JOIN "IGN1" L ON L."DocEntry" = H."DocEntry"
INNER JOIN "OBPL" B ON B."BPLId" = H."BPLId"
WHERE H."DocEntry" = :DocEntry
AND IFNULL(H."CANCELED", 'N') = 'N' 
))
;
--Call  "Check_GR_Price_Fluctuation"('115')
End;
 
 