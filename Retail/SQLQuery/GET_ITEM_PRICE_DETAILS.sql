CREATE PROCEDURE "GET_ITEM_PRICE_DETAILS"
(
    IN ItemCode NVARCHAR(50),
    IN BranchID NVARCHAR(500)
)
LANGUAGE SQLSCRIPT
AS
BEGIN
SELECT
    X."ItemCode",X."ItemName",X."BranchID",X."BranchName",X."PriceListID",
    X."PriceListName",X."UomName",X."Barcode",X."BarcodeName",X."Purchase Price Without VAT",
    X."Purchase Price With VAT",X."Sales Price Without VAT",X."Sales With VAT",
    X."Purchase VAT %",X."Sales VAT %"
FROM (SELECT T0."ItemCode",T0."ItemName",
        T2."ItmsGrpNam",T0."U_VSubGroup" AS "SubGroup",T0."U_VSubSubGroup" AS "SubSubGroup",
        T0."U_BRAND" AS "BrandName",B."BPLId" AS "BranchID",
        B."BPLName" AS "BranchName",P."ListNum" AS "PriceListID",
        P."ListName" AS "PriceListName",T1."UomName",
        T3."BcdCode" AS "Barcode",T3."BcdName" AS "BarcodeName",
        CASE WHEN P."ListNum" = 11 THEN ROUND(T5."Price" + (T5."Price" * IFNULL(VP."Rate",0) / 100), 3)
        ELSE NULL END AS "Purchase Price With VAT",
        CASE WHEN P."ListNum" = 11 THEN T5."Price"
        ELSE NULL END AS "Purchase Price Without VAT",
        CASE WHEN P."ListNum" <> 11 THEN T5."Price"
        ELSE NULL END AS "Sales With VAT",
        CASE WHEN P."ListNum" <> 11 THEN ROUND(T5."Price" / (1 + IFNULL(VS."Rate",0) / 100), 3)
        ELSE NULL END AS "Sales Price Without VAT",
        VP."Rate" AS "Purchase VAT %",
        VS."Rate" AS "Sales VAT %"
    FROM OITM T0
    INNER JOIN OUOM T1 ON T0."PUoMEntry" = T1."UomEntry"
    INNER JOIN OITB T2 ON T0."ItmsGrpCod" = T2."ItmsGrpCod"
    LEFT JOIN OBCD T3 ON T0."ItemCode" = T3."ItemCode" AND T0."PUoMEntry" = T3."UomEntry"
    INNER JOIN ITM1 T5 ON T0."ItemCode" = T5."ItemCode"
    INNER JOIN OPLN P ON T5."PriceList" = P."ListNum"
    LEFT JOIN OBPL B ON B."U_PricelistNo" = P."ListNum"
    LEFT JOIN OVTG VP ON T0."VatGroupPu" = VP."Code"
    LEFT JOIN OVTG VS ON T0."VatGourpSa" = VS."Code"
    WHERE T0."ItemCode" = :ItemCode

    UNION ALL

    SELECT T0."ItemCode",T0."ItemName",
        T2."ItmsGrpNam",T0."U_VSubGroup" AS "SubGroup",T0."U_VSubSubGroup" AS "SubSubGroup",
        T0."U_BRAND" AS "BrandName",B."BPLId" AS "BranchID",B."BPLName" AS "BranchName",
        P."ListNum" AS "PriceListID",P."ListName" AS "PriceListName",
        T6U."UomName",T3."BcdCode" AS "Barcode",T3."BcdName" AS "BarcodeName",
        CASE WHEN P."ListNum" = 11 THEN ROUND(T6."Price" + (T6."Price" * IFNULL(VP."Rate",0) / 100), 3)
        ELSE NULL END AS "Purchase Price With VAT",
        CASE WHEN P."ListNum" = 11 THEN T6."Price" ELSE NULL END AS "Purchase Price Without VAT",
        CASE WHEN P."ListNum" <> 11 THEN T6."Price" ELSE NULL END AS "Sales With VAT",
        CASE WHEN P."ListNum" <> 11 THEN ROUND(T6."Price" / (1 + IFNULL(VS."Rate",0) / 100), 3)
        ELSE NULL END AS "Sales Price Without VAT",
        VP."Rate" AS "Purchase VAT %",
        VS."Rate" AS "Sales VAT %"
    FROM OITM T0
    INNER JOIN OITB T2 ON T0."ItmsGrpCod" = T2."ItmsGrpCod"
    INNER JOIN ITM9 T6 ON T0."ItemCode" = T6."ItemCode"
    INNER JOIN OUOM T6U ON T6."UomEntry" = T6U."UomEntry"
    INNER JOIN OPLN P ON T6."PriceList" = P."ListNum"
    LEFT JOIN OBPL B ON B."U_PricelistNo" = P."ListNum"
    LEFT JOIN OBCD T3 ON T0."ItemCode" = T3."ItemCode" AND T6."UomEntry" = T3."UomEntry"
    LEFT JOIN OVTG VP ON T0."VatGroupPu" = VP."Code"
    LEFT JOIN OVTG VS ON T0."VatGourpSa" = VS."Code"
    WHERE T0."ItemCode" = :ItemCode) X
--WHERE (X."BranchID" IS NULL OR X."BranchID" = :BranchID)
WHERE 
(
    X."BranchID" IS NULL
    OR ',' || :BranchID || ',' 
       LIKE '%,' || TO_NVARCHAR(X."BranchID") || ',%'
)
ORDER BY X."BranchID", X."PriceListID",X."UomName";
END;