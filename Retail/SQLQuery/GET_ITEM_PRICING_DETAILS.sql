CREATE PROCEDURE "GET_ITEM_PRICING_DETAILS"
(
    IN P_ItemCode NVARCHAR(50)
)
LANGUAGE SQLSCRIPT
AS
BEGIN

DECLARE v_VatSale DECIMAL(19,6);
DECLARE v_VatPurchase DECIMAL(19,6);

SELECT TOP 1 (T1."Rate"+100)/100 INTO v_VatSale
FROM OITM T0
INNER JOIN OVTG T1 ON T0."VatGourpSa"=T1."Code"
WHERE T0."ItemCode" = :P_ItemCode;

SELECT TOP 1 (T1."Rate"+100)/100 INTO v_VatPurchase
FROM OITM T0
INNER JOIN OVTG T1 ON T0."VatGroupPu"=T1."Code"
WHERE T0."ItemCode" = :P_ItemCode;

WITH "LastPurchase" AS
(
    SELECT
        T0."BPLId" AS "BranchId",
        ((T1."INMPrice"/T1."NumPerMsr")-(IFNULL((Select SUM(IFNULL((IFNULL(A1."INMPrice", 0)/A1."NumPerMsr"),0)) 
from RPC1 A1 left join ORPC A2 on A2."DocEntry"=A1."DocEntry" where A2."CANCELED"='N' 
and A1."ItemCode"=T1."ItemCode" and A1."NoInvtryMv"='Y' and A2."BPLId" = T0."BPLId" and IFNULL(A2."U_CMType",'')='Normal'),0))) AS "LastPrice",
        ROW_NUMBER() OVER(PARTITION BY T0."BPLId" ORDER BY T0."DocDate" DESC,T0."DocEntry" DESC) RN
    FROM OPDN T0
    INNER JOIN PDN1 T1 ON T0."DocEntry"=T1."DocEntry"
    WHERE T1."ItemCode" = :P_ItemCode
      AND T0."CANCELED" = 'N' AND NOT (IFNULL(T1."U_normalqty",0) = 0 AND IFNULL(T1."U_FOCQTY",0) <> 0)
),

"CompleteQty" AS
(
    SELECT
        AA."ItemCode",
        T1."BPLId" AS "BranchId",
        T1."BPLName" AS "BranchName",
        (SELECT "ListName" FROM OPLN WHERE "ListNum"=T1."U_PricelistNo") AS "PriceList",
        (SELECT "IsGrossPrc" FROM OPLN WHERE "ListNum"=T1."U_PricelistNo") AS "PriceMode",
        AA."UomEntry",
        AA."BaseQty",
        AA."UomCode",
        0 AS "OldPrice"
    FROM 
    (
        SELECT
            T0."ItemCode",
            C."UomCode",
            C."UomEntry",
            B."BaseQty",
            ROW_NUMBER() OVER(PARTITION BY B."UomEntry" ORDER BY B."UomEntry") AS RN
        FROM OITM T0
        INNER JOIN OUGP A ON A."UgpEntry"=T0."UgpEntry"
        INNER JOIN UGP1 B ON B."UgpEntry"=A."UgpEntry"
        INNER JOIN OUOM C ON B."UomEntry"=C."UomEntry"
        WHERE T0."ItemCode" = :P_ItemCode
    ) AA
    INNER JOIN OBPL T1 ON AA.RN = 1

    UNION ALL

    SELECT
        T0."ItemCode",
        T2."BPLId",
        T2."BPLName",
        (SELECT "ListName" FROM OPLN WHERE "ListNum"=T2."U_PricelistNo"),
        (SELECT "IsGrossPrc" FROM OPLN WHERE "ListNum"=T2."U_PricelistNo"),
        B."UomEntry",
        B."BaseQty",
        C."UomCode",
        CASE 
            WHEN B."UomEntry"=T0."IUoMEntry" THEN T1."Price"
            WHEN B."UomEntry"=T3."UomEntry" AND T3."PriceList"=T2."U_PricelistNo" THEN T3."Price"
            ELSE 0
        END
    FROM OITM T0
    INNER JOIN OUGP A ON A."UgpEntry"=T0."UgpEntry"
    INNER JOIN UGP1 B ON B."UgpEntry"=A."UgpEntry"
    INNER JOIN OUOM C ON B."UomEntry"=C."UomEntry"
    INNER JOIN ITM1 T1 ON T1."ItemCode"=T0."ItemCode" AND T1."UomEntry"=T0."IUoMEntry"
    INNER JOIN OBPL T2 ON T2."U_PricelistNo"=T1."PriceList"
    LEFT JOIN ITM9 T3 ON T3."ItemCode"=T0."ItemCode"
        AND T3."UomEntry"=B."UomEntry"
        AND T3."PriceList"=T2."U_PricelistNo"
    WHERE T0."ItemCode" = :P_ItemCode
)

SELECT
    Q."BranchName" AS "Branch",
    Q."PriceList" AS "Price List",
    Q."UomCode" AS "UOM",
    
  /*  ROUND(IFNULL(Q."ItemCost",0),3) AS "Item Cost",

ROUND(IFNULL(LP."LastPrice",0) * Q."BaseQty",3) AS "Last Purchase Price",

ROUND((IFNULL(LP."LastPrice",0) * :v_VatPurchase) * Q."BaseQty",3) AS "Last Purchase Price With Vat",

ROUND(Q."RetailWithoutVat" * Q."BaseQty",3) AS "Retail Price Without Vat",

ROUND((Q."RetailWithoutVat" * :v_VatSale) * Q."BaseQty",3) AS "Retail Price With Vat", */

TO_VARCHAR(ROUND(IFNULL(Q."ItemCost",0),3), '0.000') AS "Item Cost",

TO_VARCHAR(ROUND(IFNULL(LP."LastPrice",0) * Q."BaseQty",3), '0.000') AS "Last Purchase Price",

TO_VARCHAR(ROUND((IFNULL(LP."LastPrice",0) * :v_VatPurchase) * Q."BaseQty",3), '0.000') AS "Last Purchase Price With Vat",

TO_VARCHAR(ROUND(Q."RetailWithoutVat",3), '0.000') AS "Retail Price Without Vat",

TO_VARCHAR(ROUND((Q."RetailWithoutVat" * :v_VatSale),3), '0.000') AS "Retail Price With Vat",

    /*IFNULL(Q."ItemCost",0) AS "Item Cost",

    IFNULL(LP."LastPrice",0) * Q."BaseQty" AS "Last Purchase Price",
    (IFNULL(LP."LastPrice",0) * :v_VatPurchase )* Q."BaseQty" AS "Last Purchase Price With Vat",

    Q."RetailWithoutVat" * Q."BaseQty" AS "Retail Price Without Vat",
    (Q."RetailWithoutVat" * :v_VatSale) * Q."BaseQty" AS "Retail Price With Vat",*/
    

    /*ROUND(CASE 
        WHEN (Q."RetailWithoutVat" * :v_VatSale) <> 0 
        THEN ((Q."RetailWithoutVat" * :v_VatSale - IFNULL(LP."LastPrice",0) * :v_VatPurchase)
             /(Q."RetailWithoutVat" * :v_VatSale)) * 100
        ELSE 0
    END,3) AS "GP",

    -- ✅ UPDATED MARKUP (Excel Logic)
    ROUND(CASE 
        WHEN (IFNULL(LP."LastPrice",0) * :v_VatPurchase) <> 0
        THEN ((Q."RetailWithoutVat" * :v_VatSale - IFNULL(LP."LastPrice",0) * :v_VatPurchase)
             /(IFNULL(LP."LastPrice",0) * :v_VatPurchase)) * 100
        ELSE 0
    END,3) AS "Markup %",*/
    
    TO_VARCHAR(ROUND(CASE
            WHEN (Q."RetailWithoutVat" * :v_VatSale) <> 0
            THEN ((Q."RetailWithoutVat" * :v_VatSale - (IFNULL(LP."LastPrice",0)* Q."BaseQty") * :v_VatPurchase)
                 / (Q."RetailWithoutVat" * :v_VatSale)) * 100
            ELSE 0 END,3),'0.000') AS "GP",

-- ✅ UPDATED MARKUP (Excel Logic)
TO_VARCHAR(ROUND(CASE WHEN (IFNULL(LP."LastPrice",0) * :v_VatPurchase) <> 0
            THEN ((Q."RetailWithoutVat" * :v_VatSale - (IFNULL(LP."LastPrice",0)* Q."BaseQty") * :v_VatPurchase)
                 / ((IFNULL(LP."LastPrice",0)* Q."BaseQty") * :v_VatPurchase)) * 100
            ELSE 0 END,3),'0.000') AS "Markup %",

    Q."InStock"

FROM
(
    SELECT
        CQ."BranchId",
        CQ."BranchName",
        CQ."PriceList",
        CQ."UomCode",
        CQ."BaseQty",

        SUM(CQ."OldPrice") AS "ItemPrice",
        (SUM(CQ."OldPrice")/:v_VatSale) AS "RetailWithoutVat",

        IFNULL(
            (SELECT "AvgPrice"
             FROM OITW
             WHERE "ItemCode"=:P_ItemCode
             LIMIT 1),0
        ) * CQ."BaseQty" AS "ItemCost",

        ROUND(
            IFNULL(
                (SELECT "OnHand"
                 FROM OITW
                 WHERE "ItemCode"=:P_ItemCode
                 LIMIT 1),0
            ) / CQ."BaseQty",2
        ) AS "InStock"

    FROM "CompleteQty" CQ
    GROUP BY
        CQ."BranchId", CQ."BranchName",
        CQ."PriceList", CQ."UomCode", CQ."BaseQty"
) Q

LEFT JOIN "LastPurchase" LP 
    ON Q."BranchId" = LP."BranchId" AND LP.RN = 1
ORDER BY Q."BranchName"; 
END;