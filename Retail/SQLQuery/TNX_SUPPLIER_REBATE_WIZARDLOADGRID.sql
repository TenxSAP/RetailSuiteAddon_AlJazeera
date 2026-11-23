CREATE PROCEDURE "TNX_SUPPLIER_REBATE_WIZARDLOADGRID"
(
    IN P_FDATE      NVARCHAR(20),
    IN P_TDATE      NVARCHAR(20),
    IN P_BRANCH     NVARCHAR(50),
    IN P_SUPPLIER   NVARCHAR(50)
)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER
AS
BEGIN

    SELECT 
        'Y' AS "Check",

        T."SupplierCode"      AS "Supplier Code",
        T."SupplierName"      AS "Supplier Name",
        T."APInvoiceDocNum"   AS "AP Invoice Number",
        T."Brand"             AS "Brand",
        T."Department"        AS "Department",
        T."Section1"          AS "Section",
        T."Family"            AS "Family",
        T."SubFamily"         AS "Sub Family",
        T."RebatePercentage"  AS "Rebate Percentage",
        T."GLAccount"         AS "GL Account",
        T."Currency"          AS "Currency",
         T."VatGroup"          AS "VatGroup",

        T."TotalInvoiceAmount"    AS "Total Invoice Amount",
        T."TotalApcreditAmount"   AS "Total APCredit Amount",

        ( T."TotalInvoiceAmount" - T."TotalApcreditAmount" )
            AS "Net purchase Amount for rebate",

        ( T."RebatePercentage" / 100 ) *
        ( T."TotalInvoiceAmount" - T."TotalApcreditAmount" )
            AS "Rebate Amount",

        T."APCreditMemo"  AS "Ap Credit memo",
        T."Status"        AS "Status"

    FROM "SBO_RebateAmount"(
            TO_DATE(:P_FDATE), 
            TO_DATE(:P_TDATE), 
            :P_BRANCH
         ) T

    WHERE T."SupplierCode" = :P_SUPPLIER;

END;