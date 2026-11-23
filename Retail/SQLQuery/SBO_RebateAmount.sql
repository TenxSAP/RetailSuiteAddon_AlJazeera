CREATE FUNCTION "SBO_RebateAmount"(
    IN FromDate DATE,
    IN ToDate DATE,
    IN Branch VARCHAR(15)
)
RETURNS TABLE 
(
    "Department" NVARCHAR(30),
    "Section1" NVARCHAR(100),
    "Brand" NVARCHAR(50),
    "Family" NVARCHAR(50),
    "SubFamily" NVARCHAR(100),
    "SupplierCode" NVARCHAR(100),
    "SupplierName" NVARCHAR(254),
    "RebatePercentage" DOUBLE,
    "GLAccount" VARCHAR(215),
    "Currency" VARCHAR(215),
    "APInvoiceDocNum" NVARCHAR(50),
    "TotalInvoiceAmount" DOUBLE,
    "TotalApcreditAmount" DOUBLE,
    "NetPurchaseAmount" DOUBLE,
    "RebateAmount" DOUBLE,
    "APCreditMemo" NVARCHAR(250),
    "Status" NVARCHAR(5000)
)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER
AS
BEGIN

RETURN

SELECT 
    S."DEPARTMENT"        AS "Department",
    S."SECTION1"          AS "Section1",
    S."BRAND"             AS "Brand",
    S."FAMILY"            AS "Family",
    S."SUBFAMILY"         AS "SubFamily",
    S."SUPPLIERCODE"      AS "SupplierCode",
    S."SUPPLIERNAME"      AS "SupplierName",
    S."REBATEPERCENTAGE"  AS "RebatePercentage",
    S."GLACCOUNT"         AS "GLAccount",
    S."CURRENCY"          AS "Currency",
    S."APInvoiceDocNum"   AS "APInvoiceDocNum",
    SUM(S."TOTALINVOICEAMOUNT")  AS "TotalInvoiceAmount",
    SUM(S."TOTALAPCREDITAMOUNT") AS "TotalApcreditAmount",
    SUM(S."TOTALINVOICEAMOUNT") 
      - SUM(S."TOTALAPCREDITAMOUNT") AS "NetPurchaseAmount",
    0 AS "RebateAmount",
    '' AS "APCreditMemo",
    '' AS "Status"

FROM
(

/* ================= AP INVOICE ================= */

SELECT 
    T0."DocNum" AS "APInvoiceDocNum",
    T4."U_Depart" AS "DEPARTMENT",
    T4."U_Section" AS "SECTION1",
    T4."U_Brand" AS "BRAND",
    T4."U_Family" AS "FAMILY",
    T4."U_SFamily" AS "SUBFAMILY",
    T2."BpCode" AS "SUPPLIERCODE",
    T2."BpName" AS "SUPPLIERNAME",
    T4."U_RBTPRC" AS "REBATEPERCENTAGE",
    T4."U_GLact" AS "GLACCOUNT",
    T0."DocCur" AS "CURRENCY",
    SUM(
        CASE 
            WHEN T0."DocCur" = (SELECT "MainCurncy" FROM OADM)
                THEN IFNULL(T1."LineTotal",0)
            ELSE IFNULL(T1."TotalFrgn",0)
        END
    ) AS "TOTALINVOICEAMOUNT",
    0 AS "TOTALAPCREDITAMOUNT"

FROM OPCH T0
JOIN PCH1 T1 ON T0."DocEntry" = T1."DocEntry"
JOIN OITM I ON I."ItemCode" = T1."ItemCode"
JOIN OOAT T2 ON T1."AgrNo" = T2."AbsID"
JOIN "@SBO_SUPREBATE" T3 ON T3."U_BASENO" = T2."Number"
JOIN "@SBO_SUPREBATE_C0" T4 ON T4."DocEntry" = T3."DocEntry" AND (T4."U_Brand" = 'All' OR T4."U_Brand" = I."U_BRAND") --AND T4."U_Brand"= I."U_BRAND"

WHERE 
    T0."DocDate" BETWEEN :FromDate AND :ToDate
    AND T0."BPLId" = :Branch
    AND T2."Status" = 'A'
    AND T0."CANCELED" = 'N'
    AND IFNULL(T1."U_Apnum1",'') = ''
    AND IFNULL(I."U_REBEGBLE",'No') = 'Yes'

GROUP BY
    T0."DocNum",
    T4."U_Depart",
    T4."U_Section",
    T4."U_Brand",
    T4."U_Family",
    T4."U_SFamily",
    T2."BpCode",
    T2."BpName",
    T4."U_RBTPRC",
    T4."U_GLact",
    T0."DocCur"

UNION ALL

/* ================= AP CREDIT MEMO ================= */

SELECT 
    T0."DocNum" AS "APInvoiceDocNum",
    T4."U_Depart",
    T4."U_Section",
    T4."U_Brand",
    T4."U_Family",
    T4."U_SFamily",
    T2."BpCode",
    T2."BpName",
    T4."U_RBTPRC",
    T4."U_GLact",
    T0."DocCur",
    0,
    SUM(
        CASE 
            WHEN T0."DocCur" = (SELECT "MainCurncy" FROM OADM)
                THEN IFNULL(T1."LineTotal",0)
            ELSE IFNULL(T1."TotalFrgn",0)
        END
    )AS "TOTALAPCREDITAMOUNT"

FROM ORPC T0
JOIN RPC1 T1 ON T0."DocEntry" = T1."DocEntry"
JOIN OITM I ON I."ItemCode" = T1."ItemCode"
JOIN OOAT T2 ON T1."AgrNo" = T2."AbsID"
JOIN "@SBO_SUPREBATE" T3 ON T3."U_BASENO" = T2."Number"
JOIN "@SBO_SUPREBATE_C0" T4 ON T4."DocEntry" = T3."DocEntry" AND (T4."U_Brand" = 'All' OR T4."U_Brand" = I."U_BRAND") AND T4."U_Brand"= I."U_BRAND"

WHERE 
    T0."DocDate" BETWEEN :FromDate AND :ToDate
    AND T0."BPLId" = :Branch
    AND T2."Status" = 'A'
    AND T0."CANCELED" = 'N'
    AND IFNULL(T1."U_Apnum1",'') = ''
     AND IFNULL(I."U_REBEGBLE",'No') = 'Yes'

GROUP BY
    T0."DocNum",
    T4."U_Depart",
    T4."U_Section",
    T4."U_Brand",
    T4."U_Family",
    T4."U_SFamily",
    T2."BpCode",
    T2."BpName",
    T4."U_RBTPRC",
    T4."U_GLact",
    T0."DocCur"

) S

GROUP BY
    S."DEPARTMENT",
    S."SECTION1",
    S."BRAND",
    S."FAMILY",
    S."SUBFAMILY",
    S."SUPPLIERCODE",
    S."SUPPLIERNAME",
    S."REBATEPERCENTAGE",
    S."GLACCOUNT",
    S."CURRENCY",
    S."APInvoiceDocNum";

END;


