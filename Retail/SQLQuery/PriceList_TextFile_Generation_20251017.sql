CREATE PROCEDURE "PriceList_TextFile_Generation_20251017"
(
    IN BranchID INTEGER,
    IN ItemGroup NVARCHAR(20),
    IN DateFilter NVARCHAR(20),
    IN ScaleType NVARCHAR(20)
)
language SQLScript
sql security invoker

AS 

BEGIN 

IF :DateFilter ='2' THEN 

SELECT *, MOD(a."Rate", 1) AS "Rate1" 
FROM 
(
SELECT 
(CASE WHEN T2."UomCode"='KG1' THEN '1' ELSE '2' END) as "FLAG",

A."BcdCode" AS "EANCode",

(CASE WHEN LENGTH(A."BcdCode")='6' 
THEN RIGHT(A."BcdCode",4) 
ELSE '' END) as "PLU",
(CASE
    WHEN T2."UomCode" = 'KG1' THEN '1'
    ELSE '2'
END) AS "PLU Type",
V."ItemName" AS "ItemName",
V."FrgnName" AS "arabic name",
A."BcdCode" AS "ItemCode",

RPAD(T2."UomCode",3) AS "UomCode",

-- RATE LOGIC (UNCHANGED)
CASE 
WHEN IFNULL
(
(
SELECT MIN(A1."U_OFPRWVAT")
FROM "@TNX_ITEMDISC" A0
INNER JOIN "@TNX_ITEMDISC_C0" A1 ON A0."DocEntry" = A1."DocEntry"
WHERE UPPER(A0."U_STATUS") = 'ACTIVE'
AND A1."U_Active"='Yes'
AND A0."U_BPLID" = :BranchID
AND A1."U_ITMCODE" = A."ItemCode"
AND A1."U_UOM" = T2."UomCode"
),0) > 0

THEN
IFNULL
(
(
SELECT MIN(A1."U_OFPRWVAT")
FROM "@TNX_ITEMDISC" A0
INNER JOIN "@TNX_ITEMDISC_C0" A1 ON A0."DocEntry" = A1."DocEntry"
WHERE UPPER(A0."U_STATUS") = 'ACTIVE'
AND A1."U_Active"='Yes'
AND A0."U_BPLID" = :BranchID
AND A1."U_ITMCODE" = A."ItemCode"
AND A1."U_UOM" = T2."UomCode"
),0)

ELSE

IFNULL
(
CASE 
WHEN (SELECT "IUoMEntry" FROM OITM WHERE "ItemCode" = A."ItemCode") = A."UomEntry"
THEN
(
SELECT "Price"
FROM ITM1
WHERE "ItemCode" = A."ItemCode"
AND "PriceList" =
(
SELECT T0."U_PricelistNo"
FROM OBPL T0
WHERE T0."BPLId" = :BranchID
)
)
ELSE
(
SELECT "Price"
FROM ITM9
WHERE "ItemCode" = A."ItemCode"
AND "PriceList" =
(
SELECT T0."U_PricelistNo"
FROM OBPL T0
WHERE T0."BPLId" = :BranchID
)
AND "UomEntry" = A."UomEntry"
)
END
,0)

END AS "Rate",

/*(CASE WHEN LENGTH(A."BcdCode")='6' 
THEN RIGHT(LEFT(A."BcdCode",3),2) 
ELSE '' END) AS "Department",

IFNULL(V."U_ExpiryDays",0) AS "Expire Days",*/
'99' AS "Department",
7 AS "Expire Days",

-- OPTIONAL (for debugging / future use)
IFNULL(V."U_SCALET",'Internal Scale') AS "ScaleType"

FROM OBCD A
INNER JOIN OITM V ON A."ItemCode" = V."ItemCode"
INNER JOIN OUOM T2 ON A."UomEntry" = T2."UomEntry"

WHERE 

IFNULL(V."U_AcceptWMQty",'2') = '1'

AND
(
    CASE
        WHEN :ScaleType = 'All Scale' THEN 1
        WHEN :ScaleType = '' THEN 1
        WHEN IFNULL(V."U_SCALET",'Internal Scale') = :ScaleType THEN 1
        ELSE 0
    END = 1
)

AND IFNULL(A."BcdCode",'') <> ''
AND LENGTH(A."BcdCode") <= 6
AND (
:ItemGroup = 'All'
OR V."ItmsGrpCod" = CAST(:ItemGroup AS INTEGER)
)

) a;

--WHERE a."Rate" <> 0;


ELSEIF :DateFilter ='1' THEN 

SELECT *, MOD(a."Rate", 1) AS "Rate1" 
FROM 
(
SELECT 

(CASE WHEN T2."UomCode"='KG' THEN '1' ELSE '2' END) as "FLAG",

A."BcdCode" AS "EANCode",

(CASE WHEN LENGTH(A."BcdCode")='6' 
THEN RIGHT(A."BcdCode",4) 
ELSE '' END) AS "PLU",

V."ItemName" AS "ItemName",
V."FrgnName" AS "arabic name",
A."BcdCode" AS "ItemCode",
RPAD(T2."UomCode",3) AS "UomCode",

-- SAME RATE LOGIC

CASE 
WHEN IFNULL
(
(
SELECT MIN(A1."U_OFPRWVAT")
FROM "@TNX_ITEMDISC" A0
INNER JOIN "@TNX_ITEMDISC_C0" A1 ON A0."DocEntry" = A1."DocEntry"
WHERE UPPER(A0."U_STATUS") = 'ACTIVE'
AND A1."U_Active"='Yes'
AND A0."U_BPLID" = :BranchID
AND A1."U_ITMCODE" = A."ItemCode"
AND A1."U_UOM" = T2."UomCode"
),0) > 0

THEN
IFNULL
(
(
SELECT MIN(A1."U_OFPRWVAT")
FROM "@TNX_ITEMDISC" A0
INNER JOIN "@TNX_ITEMDISC_C0" A1 ON A0."DocEntry" = A1."DocEntry"
WHERE UPPER(A0."U_STATUS") = 'ACTIVE'
AND A1."U_Active"='Yes'
AND A0."U_BPLID" = :BranchID
AND A1."U_ITMCODE" = A."ItemCode"
AND A1."U_UOM" = T2."UomCode"
),0)

ELSE

IFNULL
(
CASE 
WHEN (SELECT "IUoMEntry" FROM OITM WHERE "ItemCode" = A."ItemCode") = A."UomEntry"
THEN
(
SELECT "Price"
FROM ITM1
WHERE "ItemCode" = A."ItemCode"
AND "PriceList" =
(
SELECT T0."U_PricelistNo"
FROM OBPL T0
WHERE T0."BPLId" = :BranchID
)
)
ELSE
(
SELECT "Price"
FROM ITM9
WHERE "ItemCode" = A."ItemCode"
AND "PriceList" =
(
SELECT T0."U_PricelistNo"
FROM OBPL T0
WHERE T0."BPLId" = :BranchID
)
AND "UomEntry" = A."UomEntry"
)
END
,0)

END AS "Rate",

/*(CASE WHEN LENGTH(A."BcdCode")='6'
THEN RIGHT(LEFT(A."BcdCode",3),2)
ELSE '' END) AS "Department",

IFNULL(V."U_ExpiryDays",0) AS "Expire Days",*/
'99' AS "Department",
7 AS "Expire Days",

-- OPTIONAL
IFNULL(V."U_SCALET",'Internal Scale') AS "ScaleType"

FROM OBCD A
INNER JOIN OITM V ON A."ItemCode" = V."ItemCode"
INNER JOIN OUOM T2 ON A."UomEntry" = T2."UomEntry"

WHERE 

IFNULL(V."U_AcceptWMQty",'2') = '1'

-- ✅ FINAL FIX (IMPORTANT)
AND (
    :ScaleType = '' 
    OR IFNULL(V."U_SCALET",'Internal Scale') = :ScaleType
)

AND (V."UpdateDate"=CURRENT_DATE OR V."CreateDate"=CURRENT_DATE)
AND IFNULL(A."BcdCode",'') <> ''
AND LENGTH(A."BcdCode") <= 6
AND (
:ItemGroup = 'All'
OR V."ItmsGrpCod" = CAST(:ItemGroup AS INTEGER)
)

) a;

--WHERE a."Rate" <> 0;

END IF;

END;