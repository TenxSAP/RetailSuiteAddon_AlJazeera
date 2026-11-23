 
CREATE PROCEDURE "Check_GRPO_Price_Fluctuation_APCredit"(IN DocEntry INT)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER
AS
BEGIN
Select *
,Round("LastGRPOPrice"-"SellingPrice",3) As "OldMarginAmount"
,'Price Reduction' As "FluctuationType"
,CASE 
      WHEN IFNULL("SellingPrice", 0) = 0 THEN 0
      ELSE ROUND((ROUND("LastGRPOPrice" - "SellingPrice", 3)/ "SellingPrice") * 100, 3)
      END AS "OldMarginPercentage"
,Round("Price"-"SellingPrice",3) As "NewMarginAmount"
,CASE 
        WHEN IFNULL("SellingPrice", 0) = 0 THEN 0
        ELSE ROUND(
                (ROUND("Price" - "SellingPrice", 3) 
                / "SellingPrice") * 100
              , 3)
      END AS "NewMarginPercentage"
,Round(Round("Price"-"SellingPrice",3)-Round("LastGRPOPrice"-"SellingPrice",3),3) As "MarginDiff"
,"MarginTolerance"
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
,L."Quantity"
,L."LineNum"
,H."DocDate"
,H."UserSign"
,(SELECT U."USER_CODE" FROM "OUSR" U WHERE U."USERID" = H."UserSign") AS "UserCode"
,B."U_PricelistNo"
,H."BPLId"
,B."BPLName" AS "BranchCode"
,L."U_suppprice" as "SupplierCost"
,IFNULL((Select "TNX_LASTGRPO_PRICE"(L."ItemCode",L."UomEntry",H."DocDate",H."BPLId") From Dummy),0) AS "LastGRPOPrice"
,IFNULL((Select "TNX_SELLING_PRICE"(L."ItemCode",L."UomEntry",H."BPLId") From Dummy),0) As "SellingPrice"
,(Select "U_MTlce" From OADM) As "MarginTolerance"
FROM "ORPC" H
INNER JOIN "RPC1" L ON L."DocEntry" = H."DocEntry"
INNER JOIN "OBPL" B ON B."BPLId" = H."BPLId"
WHERE H."DocEntry" = :DocEntry
AND IFNULL(H."CANCELED", 'N') = 'N' 
AND IFNULL(L."NoInvtryMv",'N')='Y'
)Where Round(Round("Price"-"SellingPrice",3)-Round("LastGRPOPrice"-"SellingPrice",3),3)<"MarginTolerance"
;
--Call  "Check_GRPO_Price_Fluctuation_APCredit"('173')
End; 