CREATE PROCEDURE "TNX_STOCKBYWAREHOUSE" (
    IN ITMGROUP NVARCHAR(100)
)
LANGUAGE SQLSCRIPT
AS
BEGIN

DECLARE v_columns NVARCHAR(40000);
DECLARE v_full_sql NVARCHAR(50000);

-- Step 1: Create interleaved stock + cost columns
SELECT STRING_AGG(
    'SUM(CASE WHEN W."WhsCode" = ''' || "WhsCode" || ''' THEN (W."OnHand" * T4."AltQty" / NULLIF(T4."BaseQty", 0)) ELSE 0 END) AS "' || "WhsCode" || '_Stock", ' ||
    'SUM(CASE WHEN W."WhsCode" = ''' || "WhsCode" || ''' THEN (W."AvgPrice"* T4."AltQty" / NULLIF(T4."BaseQty", 0)) ELSE 0 END) AS "' || "WhsCode" || '_Cost", '||
    '(
        CASE 
            WHEN IFNULL((
                SELECT "Price" 
                FROM ITM1 
                WHERE "ItemCode" = T1."ItemCode" 
                  AND "PriceList" = (
                      SELECT "U_PricelistNo" 
                      FROM OBPL 
                      WHERE "BPLId" = (SELECT "BPLid" FROM OWHS WHERE "WhsCode" = ''' || "WhsCode" || ''')
                  )
                  AND "UomEntry" = T0."UomEntry"
            ), 0) = 0 
            THEN IFNULL((
                SELECT "Price" 
                FROM ITM9 
                WHERE "ItemCode" = T1."ItemCode" 
                  AND "UomEntry" = T0."UomEntry" 
                  AND "PriceList" = (
                      SELECT "U_PricelistNo" 
                      FROM OBPL 
                      WHERE "BPLId" = (SELECT "BPLid" FROM OWHS WHERE "WhsCode" = ''' || "WhsCode" || ''')
                  )
            ), 0)
            ELSE IFNULL((
                SELECT "Price" 
                FROM ITM1 
                WHERE "ItemCode" = T1."ItemCode" 
                  AND "PriceList" = (
                      SELECT "U_PricelistNo" 
                      FROM OBPL 
                      WHERE "BPLId" = (SELECT "BPLid" FROM OWHS WHERE "WhsCode" = ''' || "WhsCode" || ''')
                  )
                  AND "UomEntry" = T0."UomEntry"
            ), 0)
        END
    ) AS "' || "WhsCode" || '_RSP", ' ||

    '(
    (
        CASE 
            WHEN IFNULL((
                SELECT "Price" 
                FROM ITM1 
                WHERE "ItemCode" = T1."ItemCode" 
                AND "PriceList" = (
                    SELECT "U_PricelistNo" 
                    FROM OBPL 
                    WHERE "BPLId" = (SELECT "BPLid" FROM OWHS WHERE "WhsCode" = ''' || "WhsCode" || ''')
                )
                AND "UomEntry" = T0."UomEntry"
            ), 0) = 0 
            THEN IFNULL((
                SELECT "Price" 
                FROM ITM9 
                WHERE "ItemCode" = T1."ItemCode" 
                AND "UomEntry" = T0."UomEntry" 
                AND "PriceList" = (
                    SELECT "U_PricelistNo" 
                    FROM OBPL 
                    WHERE "BPLId" = (SELECT "BPLid" FROM OWHS WHERE "WhsCode" = ''' || "WhsCode" || ''')
                )
            ), 0)
            ELSE IFNULL((
                SELECT "Price" 
                FROM ITM1 
                WHERE "ItemCode" = T1."ItemCode" 
                AND "PriceList" = (
                    SELECT "U_PricelistNo" 
                    FROM OBPL 
                    WHERE "BPLId" = (SELECT "BPLid" FROM OWHS WHERE "WhsCode" = ''' || "WhsCode" || ''')
                )
                AND "UomEntry" = T0."UomEntry"
            ), 0)
        END
    ) 
    - 
    SUM(CASE WHEN W."WhsCode" = ''' || "WhsCode" || ''' 
        THEN (W."AvgPrice" * T4."AltQty" / NULLIF(T4."BaseQty", 0)) 
        ELSE 0 
    END)
) AS "' || "WhsCode" || '_Margin"',

    ', '
) INTO v_columns
FROM (SELECT DISTINCT "WhsCode" FROM OITW) AS Warehouses;

-- Step 2: Build final SQL
v_full_sql := '
    SELECT 
    CASE WHEN T3."BaseUom" = T0."UomEntry" THEN ''Yes'' ELSE ''No'' END AS "BaseUom",
     T5."ItmsGrpNam" AS "ItemGroup",
      T1."U_VSubGroup",
        T1."U_VSubSubGroup",
        T0."BcdCode",  
        T1."ItemCode", 
        T1."ItemName",
       -- T5."ItmsGrpNam" AS "ItemGroup",
        --T1."U_VSubGroup",
       -- T1."U_VSubSubGroup", 
        T1."U_packing",
         T2."UomCode",
       -- CASE WHEN T3."BaseUom" = T0."UomEntry" THEN ''Yes'' ELSE ''No'' END AS "BaseUom",
        T4."BaseQty" / NULLIF(T4."AltQty", 0) AS "CF",
        ' || v_columns || ',
        SUM(W."OnHand" * T4."AltQty" / NULLIF(T4."BaseQty", 0)) AS "TotalStock"
    FROM OITM T1  
    INNER JOIN OUGP T3 ON T1."UgpEntry" = T3."UgpEntry" 
    INNER JOIN UGP1 T4 ON T3."UgpEntry" = T4."UgpEntry" 
    LEFT JOIN OBCD T0 ON T0."ItemCode" = T1."ItemCode" AND T4."UomEntry" = T0."UomEntry" 
    LEFT JOIN OUOM T2 ON T0."UomEntry" = T2."UomEntry"
    LEFT JOIN OITW W ON W."ItemCode" = T1."ItemCode"
    INNER JOIN OITB T5 ON T5."ItmsGrpCod" = T1."ItmsGrpCod"
    WHERE T5."ItmsGrpNam" = (CASE WHEN :ITMGROUP IS NULL OR :ITMGROUP = '''' THEN T5."ItmsGrpNam" ELSE :ITMGROUP END)
    GROUP BY 
        T0."BcdCode", T1."ItemCode", T1."ItemName", T5."ItmsGrpNam", T1."U_VSubGroup", T1."U_VSubSubGroup",
        T2."UomCode", T1."U_packing", T3."BaseUom", T0."UomEntry", T4."BaseQty", T4."AltQty"
    ORDER BY T1."ItemCode", "BaseUom" DESC
';

-- Step 3: Execute
EXECUTE IMMEDIATE :v_full_sql USING ITMGROUP;

END;
