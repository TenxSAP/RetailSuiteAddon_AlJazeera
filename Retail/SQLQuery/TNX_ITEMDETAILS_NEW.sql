CREATE PROCEDURE "TNX_ITEMDETAILS_NEW"(IN USERNAME VARCHAR(100))
AS
BEGIN

SELECT 
   T1."ItemCode",
   B."UomEntry",
   B."BcdName" "ItemName",
   T1."IUoMEntry",
   V."BARCODE",

   -- ? SAFE NUMERIC PRICELIST
   CAST(NULLIF(V."PRICELIST",'') AS INTEGER) AS "PRICELIST",

   V."PRICE",
   V."LINENUM",

   -- ? CURRENT MARGIN (SAFE)
   CASE 
     WHEN COALESCE(V."PRICE",0) = 0 THEN 0
     ELSE
     (
       COALESCE(V."PRICE",0)
       - COALESCE(
           (SELECT "GET_COST"(T1."ItemCode", B."UomEntry",
            CAST(NULLIF(V."PRICELIST",'') AS INTEGER)) FROM DUMMY), 0)
     )
     / NULLIF(V."PRICE",0) * 100
   END AS "margin",

   V."CREATEDBY",

   (SELECT "UomCode" 
    FROM OUOM 
    WHERE "UomEntry" = B."UomEntry") AS "UomName",

   (SELECT "ListName" 
    FROM OPLN 
    WHERE "ListNum" = CAST(NULLIF(V."PRICELIST",'') AS INTEGER)) AS "PriceListName",

   -- ? CURRENT COST
   COALESCE(
     (SELECT "GET_COST"(T1."ItemCode", B."UomEntry",
      CAST(NULLIF(V."PRICELIST",'') AS INTEGER)) FROM DUMMY), 0) AS "COST",

   -- ? CURRENT SELLING PRICE
   COALESCE(
     (SELECT "GET_RSP"(T1."ItemCode", B."UomEntry",
      CAST(NULLIF(V."PRICELIST",'') AS INTEGER)) FROM DUMMY), 0) AS "RSP",

   -- ? ? NEW COST
   COALESCE(V."PRICE",0)
     - COALESCE(
         (SELECT "GET_COST"(T1."ItemCode", B."UomEntry",
          CAST(NULLIF(V."PRICELIST",'') AS INTEGER)) FROM DUMMY), 0)
     AS "NCOST",

   -- ? ? NEW MARGIN
   CASE 
     WHEN COALESCE(V."PRICE",0) = 0 THEN 0
     ELSE
     (
       COALESCE(V."PRICE",0)
       - COALESCE(
           (SELECT "GET_COST"(T1."ItemCode", B."UomEntry",
            CAST(NULLIF(V."PRICELIST",'') AS INTEGER)) FROM DUMMY), 0)
     )
     / NULLIF(V."PRICE",0) * 100
   END AS "NMARG",

   CASE 
     WHEN IFNULL(T1."ItemCode",'') = '' 
     THEN 'Barcode Not Found/ Item Not found' 
     ELSE '' 
   END AS "REMARK",
NULL AS "StartDate"
   --V."UPDT" AS "StartDate"

FROM "TNX_PRICEUPDATE" V
LEFT JOIN OBCD B ON B."BcdCode" = V."BARCODE"
LEFT JOIN OITM T1 ON T1."ItemCode" = B."ItemCode"
LEFT JOIN OVTG T2 ON T2."Code" = T1."VatGourpSa"
LEFT JOIN OPLN T3 
   ON T3."ListNum" = CAST(NULLIF(V."PRICELIST",'') AS INTEGER)

-- ? ? ? ONLY LATEST ROW PER BARCODE + PRICELIST
WHERE V."CREATEDBY" = :USERNAME
  AND NULLIF(V."PRICELIST",'') IS NOT NULL
  AND V."LINENUM" =
  (
     SELECT MAX(X."LINENUM")
     FROM "TNX_PRICEUPDATE" X
     WHERE X."BARCODE"   = V."BARCODE"
       AND X."PRICELIST" = V."PRICELIST"
       AND X."CREATEDBY" = V."CREATEDBY"
  )

ORDER BY V."CREATEDBY", V."LINENUM", V."BARCODE", V."PRICELIST";

END;
 