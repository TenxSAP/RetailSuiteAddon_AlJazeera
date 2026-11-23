CREATE PROCEDURE "SP_TNX_PRICE_UPDATE"()

LANGUAGE SQLSCRIPT

SQL SECURITY INVOKER

AS

BEGIN

SELECT DISTINCT
    H."U_PriceLNum" AS "Pricelist",
    C."U_TNXICD" AS "ItemCode",
    C."U_TNXNM" AS "ItemName",
    H."U_TNXPUTP" AS "PriceChangeType",
    H."DocEntry" AS "DocEntry",
    (C."U_BCode") AS "BarCode",
    IFNULL((SELECT "UomEntry" FROM OUOM WHERE TO_VARCHAR("UomCode") = C."U_UOM"), '1') AS "UomEntry",
    C."LineId" AS "LineId",
    IFNULL((SELECT "UomEntry" FROM OUOM WHERE TO_VARCHAR("UomName") = 
        (SELECT "InvntryUom" FROM OITM WHERE "ItemCode" = C."U_TNXICD")), '1') AS "IUOM",
    IFNULL("U_NRP", 0) AS "Price",
    IFNULL("U_NRP", 0) AS "BasePrice",
0 AS "BaseAmount",
0 AS "ExTax",
0 AS "NewPrice"
FROM "@TNX_PCRT" H
INNER JOIN "@TNX_PCRT_C0" C ON H."DocEntry" = C."DocEntry"
WHERE IFNULL(C."U_PUpdate", '') <> 'Yes'
    AND IFNULL(C."U_PUS", '') = ''
    AND H."U_TNXPUTP" = 'Cost Fluctuation'
    AND IFNULL("U_NRP", 0) <> 0
    AND IFNULL((C."U_BCode"), '-2') <> '-2'
    AND C."U_SPTP" = 'Scheduled Update' and H."U_TNXSTUS" = 'A'
    and TO_TIMESTAMP(TO_VARCHAR(C."U_FDATE",'YYYY-MM-DD') || ' ' || LPAD(TO_VARCHAR(FLOOR(C."U_FTIME"/100)),2,'0') || ':' || LPAD(TO_VARCHAR(MOD(C."U_FTIME",100)),2,'0') || ':00','YYYY-MM-DD HH24:MI:SS') <= CURRENT_TIMESTAMP 
   
   
UNION
 
SELECT distinct
H."U_PriceLNum" AS "Pricelist",
C."U_TNXICD" AS "ItemCode",
C."U_TNXNM" AS "ItemName",H."U_TNXPUTP" AS "PriceChangeType",H."DocEntry" AS "DocEntry",
(C."U_BCode") AS "BarCode",
IFNULL((SELECT "UomEntry" FROM OUOM WHERE TO_VARCHAR("UomCode")=C."U_UOM"),'1') AS "UomEntry",
C."LineId" AS "LineId",
IFNULL((SELECT "UomEntry" FROM OUOM WHERE TO_VARCHAR("UomName")=
(Select "InvntryUom" from OITM where "ItemCode"=C."U_TNXICD")),'1') AS "IUOM",
IFNULL("U_NRP",0) AS "Price",
IFNULL("U_NRP",0)AS "BasePrice",
0 AS "BaseAmount",
0 AS "ExTax",
0 AS "NewPrice"
FROM "@TNX_PCRT" H
INNER JOIN "@TNX_PCRT_C0" C ON H."DocEntry"=C."DocEntry"
WHERE IFNULL(C."U_PUpdate",'')<>'Yes' and H."U_TNXSTUS" = 'A'
AND IFNULL(C."U_PUS",'')=''
AND H."U_TNXPUTP"='Cost Fluctuation' and IFNULL("U_NRP",0)<>0 
--and ifnull((C."U_BCode"),'-2')<>'-2'
and C."U_SPTP"='Immediate'

UNION
 
 
SELECT distinct
H."U_PriceLNum" AS "Pricelist",
C."U_ICode" AS "ItemCode",
C."U_IName" AS "ItemName",H."U_TNXPUTP"AS "PriceChangeType",H."DocEntry" AS "DocEntry",
C."U_BCode" AS "BarCode",
(Select "UomEntry" from "OBCD" where "BcdCode"=C."U_BCode") AS "UomEntry",
C."LineId" AS "LineId",
IFNULL((SELECT "UomEntry" FROM OUOM WHERE TO_VARCHAR("UomName")=
(Select "InvntryUom" from OITM where "ItemCode"=C."U_ICode")),'1') AS "IUOM",
IFNULL(C."U_NSlng",0) AS "Price",
IFNULL(C."U_NSlng",0) AS "BasePrice",
0 AS "BaseAmount",
0 AS "ExTax",
0 AS "NewPrice"
FROM "@TNX_PCRT" H
INNER JOIN "@TNX_PCRT_C5" C ON H."DocEntry"=C."DocEntry"
WHERE IFNULL(C."U_PUpdate",'')<>'Yes'and H."U_TNXSTUS" = 'A'
AND IFNULL(C."U_PUS",'')=''
AND H."U_TNXPUTP"='Every day Special Price' and IFNULL(C."U_NSlng",0)<>0 and ifnull(C."U_BCode",'')<>''


UNION
 
SELECT distinct
(SELECT "U_PriceLst" FROM "@ME_ACFG") AS "Pricelist",
C."U_ITCD" AS "ItemCode",
C."U_ITNM" AS "ItemName",H."U_TNXPUTP"AS "PriceChangeType",H."DocEntry" AS "DocEntry",
C."U_BCode" AS "BarCode",
(Select "UomEntry" from "OBCD" where "BcdCode"=C."U_BCode") AS "UomEntry",
C."LineId" AS "LineId",
IFNULL((SELECT "UomEntry" FROM OUOM WHERE TO_VARCHAR("UomName")=
(Select "InvntryUom" from OITM where "ItemCode"=C."U_ITCD")),'1') AS "IUOM",
IFNULL(C."U_NPRC",0) AS "Price",
IFNULL(C."U_NPRC",0) AS "BasePrice",
IFNULL(C."U_Base",0) AS "BaseAmount",
IFNULL(C."U_Tax",0) AS "ExTax",
IFNULL(C."U_NPRC",0) AS "NewPrice"
FROM "@TNX_PCRT" H
INNER JOIN "@TNX_PCRT_C6" C ON H."DocEntry"=C."DocEntry"
WHERE H."U_FDATEB" IS NOT NULL and IFNULL(C."U_PUpdate",'')<>'Yes'
AND IFNULL(C."U_PUS",'')=''
AND H."U_FDATEB" = CURRENT_DATE and H."U_TNXSTUS" = 'A'
AND H."U_TNXPUTP"='Bulk Cost Update' and IFNULL(C."U_NPRC",0)<>0 and ifnull(C."U_BCode",'')<>''

UNION
 
SELECT distinct
H."U_PriceLNum" AS "Pricelist",
C."U_ITCD" AS "ItemCode",
C."U_ITNM" AS "ItemName",H."U_TNXPUTP" AS "PriceChangeType",H."DocEntry" AS "DocEntry",
C."U_BCode" AS "BarCode",
(Select "UomEntry" from "OBCD" where "BcdCode"=C."U_BCode") AS "UomEntry",
C."LineId" AS "LineId",
IFNULL((SELECT "UomEntry" FROM OUOM WHERE TO_VARCHAR("UomName")=
(Select "InvntryUom" from OITM where "ItemCode"=C."U_ITCD")),'1') AS "IUOM",
IFNULL(C."U_NSP",0) AS "Price",
IFNULL(C."U_NSP",0) AS "BasePrice",
0 AS "BaseAmount",
0 AS "ExTax",
0 AS "NewPrice"
FROM "@TNX_PCRT" H
INNER JOIN "@TNX_PCRT_C7" C ON H."DocEntry"=C."DocEntry"
WHERE H."U_TNXDDT" IS NOT NULL
AND H."U_TNXDDT" = CURRENT_DATE and H."U_TNXSTUS" = 'A' and IFNULL(C."U_PUpdate",'')<>'Yes'
AND IFNULL(C."U_PUS",'')=''
AND H."U_TNXPUTP"='Cost Fluctuation Stock Transfer' and IFNULL(C."U_NSP",0)<>0 and ifnull(C."U_BCode",'')<>''

UNION 

SELECT distinct
H."U_PriceLNum" AS "Pricelist",
C."U_ITCD" AS "ItemCode",
'' AS "ItemName",H."U_TNXPUTP" AS "PriceChangeType",H."DocEntry" AS "DocEntry",
C."U_BCode" AS "BarCode",
(Select "UomEntry" from "OBCD" where "BcdCode"=C."U_BCode") AS "UomEntry",
C."LineId" AS "LineId",
IFNULL((SELECT "UomEntry" FROM OUOM WHERE TO_VARCHAR("UomName")=
(Select "InvntryUom" from OITM where "ItemCode"=C."U_ITCD")),'1') AS "IUOM",
IFNULL(C."U_NSLP",0) AS "Price",
IFNULL(C."U_NSLP",0) AS "BasePrice",
0 AS "BaseAmount",
0 AS "ExTax",
0 AS "NewPrice"
FROM "@TNX_PCRT" H
INNER JOIN "@TNX_PCRT_C8" C ON H."DocEntry"=C."DocEntry"
WHERE H."U_TNXDDT" IS NOT NULL and H."U_TNXSTUS" = 'A' and IFNULL(C."U_PUpdate",'')<>'Yes'
AND IFNULL(C."U_PUS",'')=''
AND H."U_TNXPUTP"='Bulk Price Update' and IFNULL(C."U_NSLP",0)<>0 and ifnull(C."U_BCode",'')<>''


ORDER BY "DocEntry" DESC;
 

END;