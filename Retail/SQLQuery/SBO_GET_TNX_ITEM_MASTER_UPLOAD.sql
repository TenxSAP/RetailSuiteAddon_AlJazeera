CREATE PROCEDURE "SBO_GET_TNX_ITEM_MASTER_UPLOAD"
(
    IN AppID NVARCHAR(50)
)
LANGUAGE SQLSCRIPT
AS
BEGIN

SELECT DISTINCT
 --ROW_NUMBER() OVER (ORDER BY C0."U_KEYC") AS "Line No",

    C0."U_KEYC"        AS "key code",
    C2."U_BarCode"     AS "Barcode",
    C2."U_CNum"        AS "Vendor Item Ref No",
    C0."U_CARD"        AS "Vendor Code",
    OCRD."CardName"    AS "Vendor Name",
    C2."U_BDesc"       AS "Description",
    C0."U_FORIE"       AS "Arabic Description",
    C0."U_ITEMG"       AS "Category",
    C0."U_SUBGR"       AS "Sub Group",
    C0."U_SSGP"        AS "Sub Sub Group",

    C2."U_UOMCDE"      AS "Uom",
    (Select "U_UOMUNT" From "@TNX_ITEMUP_C1" Where "U_KEYCDE"=C0."U_KEYC" 
and "U_UOMCDE"=C2."U_UOMCDE" and "DocEntry"=H."DocEntry") AS "Cf",
   -- C1."U_UOMUNT"      AS "Cf",

    C0."U_BRAND"       AS "Brand",

   Case When C0."U_SugarTax" = 'Yes' then IFNULL(C2."U_Price",0)
When C0."U_SugarTax" = 'No' Then (Select "U_Cost" From "@TNX_ITEMUP_C3" Where "U_ItemC"=C0."U_KEYC" 
and "U_UOMCode"=C2."U_UOMCDE" and "DocEntry"=H."DocEntry")
END AS "Cost",
    
  CASE
    WHEN C0."U_PVAT" IS NULL OR C0."U_PVAT" = '' THEN 0

    WHEN C0."U_PVAT" LIKE '%5%' THEN 5
    WHEN C0."U_PVAT" LIKE '%12%' THEN 12
    WHEN C0."U_PVAT" LIKE '%18%' THEN 18
    WHEN C0."U_PVAT" LIKE '%28%' THEN 28
    ELSE 0
END AS "Vat",
CASE
    WHEN C0."U_SugarTax" = 'Yes' THEN
        IFNULL(C2."U_Price",0) +
        (
            IFNULL(C2."U_Price",0) *
            (
                CASE
                    WHEN C0."U_PVAT" IS NULL OR C0."U_PVAT" = '' THEN 0
                    WHEN C0."U_PVAT" LIKE '%5%'  THEN 5
                    WHEN C0."U_PVAT" LIKE '%12%' THEN 12
                    WHEN C0."U_PVAT" LIKE '%18%' THEN 18
                    WHEN C0."U_PVAT" LIKE '%28%' THEN 28
                    ELSE 0
                END
            ) / 100
        )

    WHEN C0."U_SugarTax" = 'No' THEN
        IFNULL(
            (Select "U_Cost"
             From "@TNX_ITEMUP_C3"
             Where "U_ItemC" = C0."U_KEYC"
             and "U_UOMCode" = C2."U_UOMCDE"
             and "DocEntry" = H."DocEntry"),0)
        +
        (
            IFNULL(
                (Select "U_Cost"
                 From "@TNX_ITEMUP_C3"
                 Where "U_ItemC" = C0."U_KEYC"
                 and "U_UOMCode" = C2."U_UOMCDE"
                 and "DocEntry" = H."DocEntry"),0)
            *
            (
                CASE
                    WHEN C0."U_PVAT" IS NULL OR C0."U_PVAT" = '' THEN 0
                    WHEN C0."U_PVAT" LIKE '%5%'  THEN 5
                    WHEN C0."U_PVAT" LIKE '%12%' THEN 12
                    WHEN C0."U_PVAT" LIKE '%18%' THEN 18
                    WHEN C0."U_PVAT" LIKE '%28%' THEN 28
                    ELSE 0
                END
            ) / 100
        )
END AS "Cost with VAT",

     C0."U_SugarTax" AS "Sugar tax",


    IFNULL(C2."U_ETAX",0)       AS "Sugar tax amt",
    IFNULL(C2."U_SUTAXAMT",0)   AS "Base cost",

    
  /*IFNULL(C2."U_SUTAXAMT",0)
+
(
    IFNULL(C2."U_SUTAXAMT",0)
    *
    (
        CASE
            WHEN C0."U_PVAT" IS NULL OR C0."U_PVAT" = '' THEN 0
            WHEN C0."U_PVAT" LIKE '%5%'  THEN 5
            WHEN C0."U_PVAT" LIKE '%12%' THEN 12
            WHEN C0."U_PVAT" LIKE '%18%' THEN 18
            WHEN C0."U_PVAT" LIKE '%28%' THEN 28
            ELSE 0
        END
    ) / 100
) AS "cost +vat",*/
    
    

    (Select "U_OldPrice" From "@TNX_ITEMUP_C3" Where "U_ItemC"=C0."U_KEYC" 
and "U_UOMCode"=C2."U_UOMCDE" and "DocEntry"=H."DocEntry") as "Selling",
    --IFNULL(C3."U_OldPrice",0)   AS "selling",
    IFNULL(C2."U_Margin",0)     AS "Margin",
    IFNULL(C2."U_Markup",0)     AS "Markup",

    C0."U_REBEGBLE"   AS "Rebate Eligibity",
    C0."U_LPO"        AS "Without PO",
    C0."U_INVENT"     AS "Inventory Item",
    C0."U_SCALE"      AS "Scale Item",
    C0."U_PURCH"      AS "Purchase Item",
    C6."U_BRANCH"     AS "Branch"

FROM "@TNX_ITEMUP" H

INNER JOIN "@TNX_ITEMUP_C0" C0
    ON H."DocEntry" = C0."DocEntry"

-- ✅ FIXED C1 (Pick correct UOM, not always LineId 1)
LEFT JOIN (
    SELECT *
    FROM (
        SELECT *,
               ROW_NUMBER() OVER (
                   PARTITION BY "DocEntry","U_KEYCDE"
                   ORDER BY "U_UOMUNT" DESC   -- 🔥 FIX
               ) RN
        FROM "@TNX_ITEMUP_C1"
    ) A WHERE RN = 1
) C1
ON C0."DocEntry" = C1."DocEntry"
AND C0."U_KEYC" = C1."U_KEYCDE"

-- ✅ C2 (same)
LEFT JOIN (
    SELECT *
    FROM (
        SELECT *,
               ROW_NUMBER() OVER (PARTITION BY "DocEntry","U_KEYC" ORDER BY "LineId") RN
        FROM "@TNX_ITEMUP_C2"
    ) A --WHERE RN = 1
) C2
ON H."DocEntry" = C2."DocEntry"
AND C0."U_KEYC" = C2."U_KEYC"

-- ✅ C3 (same)
LEFT JOIN (
    SELECT *
    FROM (
        SELECT *,
               ROW_NUMBER() OVER (PARTITION BY "DocEntry","U_ItemC" ORDER BY "LineId") RN
        FROM "@TNX_ITEMUP_C3"
    ) A WHERE RN = 1
) C3
ON H."DocEntry" = C3."DocEntry"
AND C0."U_KEYC" = C3."U_ItemC"

-- ✅ FIXED C6 (avoid duplicates)
LEFT JOIN (
    SELECT *
    FROM (
        SELECT *,
               ROW_NUMBER() OVER (PARTITION BY "DocEntry" ORDER BY "LineId") RN
        FROM "@TNX_ITEMUP_C6"
    ) A WHERE RN = 1
) C6
ON H."DocEntry" = C6."DocEntry"

LEFT JOIN "OCRD"
    ON OCRD."CardCode" = C0."U_CARD"

WHERE H."U_AppID" = :AppID
Order by  TO_INTEGER(C0."U_KEYC");
END;