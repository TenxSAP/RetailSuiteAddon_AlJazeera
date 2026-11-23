CREATE PROCEDURE "PriceList_TextFile_Generation"(IN BranchID INTEGER,IN ItemGroup INTEGER)
language SQLScript
sql security invoker

AS 

BEGIN 

--if :ItemGroup ='111' then 

SELECT *, MOD(a."Rate", 1) AS "Rate1" 
    FROM (SELECT (CASE WHEN T2."UomCode"='KG' THEN '1' ELSE '2' END) as "FLAG",
    A."BcdCode" AS "EANCode", 
    (Case When LENGTH(A."BcdCode")='6' Then RIGHT(A."BcdCode",4) Else '' end) as "PLU", 
    V."ItemName" As "ItemName", RPAD(T2."UomCode",3,'S') as "UomCode"
,CASE WHEN 

IFNULL((Select min(T5."Price") from SPP2 T5 Inner join SPP1 T6 on T5."ItemCode"=T6."ItemCode" and T5."CardCode"=T6."CardCode" 
and T5."SPP1LNum"=T6."LINENUM" and T6."ListNum"=(SELECT T0."U_PricelistNo" FROM OBPL T0 WHERE T0."BPLId" = :BranchID)
where CURRENT_DATE>=T6."FromDate" and CURRENT_DATE<=T6."ToDate" and T5."ItemCode"=A."ItemCode" and T5."UomEntry"=A."UomEntry"),0)>0 then 

IFNULL((Select min(T5."Price") from SPP2 T5 Inner join SPP1 T6 on T5."ItemCode"=T6."ItemCode" and T5."CardCode"=T6."CardCode" 
and T5."SPP1LNum"=T6."LINENUM" and T6."ListNum"=(SELECT T0."U_PricelistNo" FROM OBPL T0 WHERE T0."BPLId" = :BranchID)
where CURRENT_DATE>=T6."FromDate" and CURRENT_DATE<=T6."ToDate" and T5."ItemCode"=A."ItemCode" and T5."UomEntry"=A."UomEntry"),0) 

ELSE 

IFNULL(CASE WHEN (SELECT "IUoMEntry" FROM OITM WHERE "ItemCode" = A."ItemCode") = A."UomEntry" THEN 
(SELECT "Price" FROM ITM1 WHERE "ItemCode" = A."ItemCode" AND "PriceList" = (SELECT T0."U_PricelistNo" FROM OBPL T0 WHERE T0."BPLId" = :BranchID)) 
ELSE (SELECT "Price" FROM ITM9 WHERE "ItemCode" = A."ItemCode" AND "PriceList" = 
(SELECT T0."U_PricelistNo" FROM OBPL T0 WHERE T0."BPLId" = :BranchID) AND "UomEntry" = A."UomEntry") END, 0) 

END AS "Rate"
,(Case When LENGTH(A."BcdCode")='6' Then RIGHT(LEFT(A."BcdCode",3),2) Else '' end) as "Department"
, IFNULL(V."U_Life", 0) AS "Expire Days" 
--V."U_ScaleItem" 
FROM OBCD A 
INNER JOIN OITM V ON A."ItemCode" = V."ItemCode" 
INNER JOIN OUOM T2 ON A."UomEntry" = T2."UomEntry" 
WHERE V."U_AcceptWMQty" = '1' AND IFNULL(A."BcdCode", '') <> '' and V."ItmsGrpCod"=ItemGroup) AS a 
    WHERE a."Rate" <> 0;

/*ELSE IF :ItemGroup = '116' then


SELECT *, MOD(a."Rate", 1) AS "Rate1" 
    FROM (SELECT (CASE WHEN T2."UomCode"='KG' THEN '1' ELSE '2' END || A."BcdCode") AS "PLU",
RIGHT(A."BcdCode",4) AS "EANCode", V."ItemName" As "ItemName", RPAD(T2."UomCode",3,'S') as "UomCode"
, CASE WHEN 

IFNULL((Select min(T5."Price") from SPP2 T5 Inner join SPP1 T6 on T5."ItemCode"=T6."ItemCode" and T5."CardCode"=T6."CardCode" 
and T5."SPP1LNum"=T6."LINENUM" and T6."ListNum"=(SELECT T0."U_PricelistNo" FROM OBPL T0 WHERE T0."BPLId" = :BranchID)
where CURRENT_DATE>=T6."FromDate" and CURRENT_DATE<=T6."ToDate" and T5."ItemCode"=A."ItemCode" and T5."UomEntry"=A."UomEntry"),0)>0 then 

IFNULL((Select min(T5."Price") from SPP2 T5 Inner join SPP1 T6 on T5."ItemCode"=T6."ItemCode" and T5."CardCode"=T6."CardCode" 
and T5."SPP1LNum"=T6."LINENUM" and T6."ListNum"=(SELECT T0."U_PricelistNo" FROM OBPL T0 WHERE T0."BPLId" = :BranchID)
where CURRENT_DATE>=T6."FromDate" and CURRENT_DATE<=T6."ToDate" and T5."ItemCode"=A."ItemCode" and T5."UomEntry"=A."UomEntry"),0) 

ELSE 

IFNULL(CASE WHEN (SELECT "IUoMEntry" FROM OITM WHERE "ItemCode" = A."ItemCode") = A."UomEntry" THEN 
(SELECT "Price" FROM ITM1 WHERE "ItemCode" = A."ItemCode" AND "PriceList" = (SELECT T0."U_PricelistNo" FROM OBPL T0 WHERE T0."BPLId" = :BranchID)) 
ELSE (SELECT "Price" FROM ITM9 WHERE "ItemCode" = A."ItemCode" AND "PriceList" = 
(SELECT T0."U_PricelistNo" FROM OBPL T0 WHERE T0."BPLId" = :BranchID) AND "UomEntry" = A."UomEntry") END, 0) 

END AS "Rate"

, IFNULL(V."U_Life", 0) AS "U_Life" 
--V."U_ScaleItem"
FROM OBCD A 
INNER JOIN OITM V ON A."ItemCode" = V."ItemCode" 
INNER JOIN OUOM T2 ON A."UomEntry" = T2."UomEntry" 
WHERE V."U_AcceptWMQty" = '1' AND IFNULL(A."BcdCode", '') <> '' ) AS a 
    WHERE a."Rate" <> 0;

/*ELSE 

Select 'ID' as "PLU",'ItemCode' as "EANCode",'DepartmentID' as "ItemName",'Name1' as "UomCode",'Price' as "Rate",'UnitID'||','||'BarcodeType1'||','||'Label1ID'||','||'FreshnessDate' as "U_Life",'0' as "Rate1" FROM DUMMY

UNION ALL

SELECT "PLU","EANCode","ItemName", "UomCode", TO_CHAR(TO_DECIMAL("Rate",6,2)) as "Rate","U_Life", MOD(a."Rate", 1) AS "Rate1" 
    FROM (SELECT RIGHT(A."BcdCode",4) AS "PLU", RIGHT(A."BcdCode",4) AS "EANCode", (CASE WHEN T2."UomCode"='KG' THEN '1' ELSE '2' END || LEFT(A."BcdCode",1))  As "ItemName", V."ItemName" as "UomCode"
, (CASE WHEN 

IFNULL((Select min(T5."Price") from SPP2 T5 Inner join SPP1 T6 on T5."ItemCode"=T6."ItemCode" and T5."CardCode"=T6."CardCode" 
and T5."SPP1LNum"=T6."LINENUM" and T6."ListNum"=(SELECT T0."U_PricelistNo" FROM OBPL T0 WHERE T0."BPLId" = :BranchID)
where CURRENT_DATE>=T6."FromDate" and CURRENT_DATE<=T6."ToDate" and T5."ItemCode"=A."ItemCode" and T5."UomEntry"=A."UomEntry"),0)>0 then 

IFNULL((Select min(T5."Price") from SPP2 T5 Inner join SPP1 T6 on T5."ItemCode"=T6."ItemCode" and T5."CardCode"=T6."CardCode" 
and T5."SPP1LNum"=T6."LINENUM" and T6."ListNum"=(SELECT T0."U_PricelistNo" FROM OBPL T0 WHERE T0."BPLId" = :BranchID)
where CURRENT_DATE>=T6."FromDate" and CURRENT_DATE<=T6."ToDate" and T5."ItemCode"=A."ItemCode" and T5."UomEntry"=A."UomEntry"),0) 

ELSE 

IFNULL(CASE WHEN (SELECT "IUoMEntry" FROM OITM WHERE "ItemCode" = A."ItemCode") = A."UomEntry" THEN 
(SELECT "Price" FROM ITM1 WHERE "ItemCode" = A."ItemCode" AND "PriceList" = (SELECT T0."U_PricelistNo" FROM OBPL T0 WHERE T0."BPLId" = :BranchID)) 
ELSE (SELECT "Price" FROM ITM9 WHERE "ItemCode" = A."ItemCode" AND "PriceList" = 
(SELECT T0."U_PricelistNo" FROM OBPL T0 WHERE T0."BPLId" = :BranchID) AND "UomEntry" = A."UomEntry") END, 0) 

END) AS "Rate"

, (CASE WHEN T2."UomCode"='KG' THEN '4' ELSE '10' END)||','||(CASE WHEN T2."UomCode"='KG' THEN '7' ELSE '9' END)||',1,'||
IFNULL(V."U_Life", 0) AS "U_Life" 
--V."U_ScaleItem"
FROM OBCD A 
INNER JOIN OITM V ON A."ItemCode" = V."ItemCode" 
INNER JOIN OUOM T2 ON A."UomEntry" = T2."UomEntry" 
WHERE V."U_AcceptWMQty" = '1' AND IFNULL(A."BcdCode", '') <> '' ) AS a 
    WHERE a."Rate" <> 0;

END IF;*/
--END IF;
END;