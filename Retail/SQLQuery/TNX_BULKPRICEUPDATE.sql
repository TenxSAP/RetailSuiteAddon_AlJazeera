CREATE PROCEDURE "TNX_BULKPRICEUPDATE"(IN USERNAME VARCHAR(100))
AS
BEGIN

Select * From 
(
 SELECT T1."ItemCode"
 , B."UomEntry"
 , T1."ItemName"
 , T1."IUoMEntry"
 , v."BARCODE"
 , V."PRICELIST"
 , "PRICE" 
 , "LINENUM"
 , CASE WHEN (CASE WHEN T3."IsGrossPrc"='Y' then ("PRICE"/(1 + (T2."Rate"*0.01))) ELSE "PRICE" end) =0 then 0 else (((CASE WHEN T3."IsGrossPrc"='Y' then ("PRICE"/(1 + (T2."Rate"*0.01))) ELSE "PRICE" end)
 -(SELECT "GET_COST"(T1."ItemCode", B."UomEntry",V."PRICELIST") FROM DUMMY) )/(CASE WHEN T3."IsGrossPrc"='Y' then ("PRICE"/(1 + (T2."Rate"*0.01))) ELSE "PRICE" end)*100) end as "margin"
 , "CREATEDBY"
 , (SELECT "UomCode" FROM OUOM WHERE "UomEntry" = B."UomEntry") AS "UomName"
 , (SELECT "ListName" FROM OPLN WHERE "ListNum" = V."PRICELIST") AS "PriceListName"
 , (SELECT Y."BaseQty" / Y."AltQty" FROM OUGP Z INNER JOIN UGP1 Y ON Z."UgpEntry" = Y."UgpEntry" WHERE Z."UgpEntry" = T1."UgpEntry" AND Y."UomEntry" = B."UomEntry") AS "ConversionRate"
 , (SELECT "GET_COST"(T1."ItemCode", B."UomEntry",V."PRICELIST") FROM DUMMY) AS "COST"
 , (SELECT "GET_RSP"(T1."ItemCode", B."UomEntry",V."PRICELIST") FROM DUMMY) AS "RSP"
 , (CASE WHEN   IFNULL(T1."ItemCode",'')='' THEN 'Barcode Not Found/ Item Not found' ELSE '' END) AS "REMARK",v."UPDT" As "StartDate"
FROM   "TNX_PRICEUPDATE" v
left JOIN OBCD B ON B."BcdCode" = v."BARCODE"
LEFT JOIN OITM T1 on  T1."ItemCode" = B."ItemCode"
left join OVTG T2 on T2."Code" = T1."VatGourpSa"
left join OPLN T3 on T3."ListNum"= V."PRICELIST"
where "CREATEDBY"=:USERNAME and IFNULL("PRICELIST",'')<>''
--ORDER BY "CREATEDBY","LINENUM"

UNION ALL 

 SELECT Distinct T1."ItemCode"
 , B."UomEntry"
 , T1."ItemName"
 , T1."IUoMEntry"
 , v."BARCODE"
 , V."U_PricelistNo" As "PRICELIST"
 , "PRICE" 
 , "LINENUM"
 , CASE WHEN (CASE WHEN T3."IsGrossPrc"='Y' then ("PRICE"/(1 + (T2."Rate"*0.01))) ELSE "PRICE" end) =0 then 0 else (((CASE WHEN T3."IsGrossPrc"='Y' then ("PRICE"/(1 + (T2."Rate"*0.01))) ELSE "PRICE" end)
  -(SELECT "GET_COST"(T1."ItemCode", B."UomEntry",V."U_PricelistNo") FROM DUMMY) )/(CASE WHEN T3."IsGrossPrc"='Y' then ("PRICE"/(1 + (T2."Rate"*0.01))) ELSE "PRICE" end)*100) end as "margin"
 , "CREATEDBY"
 , (SELECT "UomCode" FROM OUOM WHERE "UomEntry" = B."UomEntry") AS "UomName"
 , (SELECT "ListName" FROM OPLN WHERE "ListNum" = V."U_PricelistNo") AS "PriceListName"
 , (SELECT Y."BaseQty" / Y."AltQty" FROM OUGP Z INNER JOIN UGP1 Y ON Z."UgpEntry" = Y."UgpEntry" WHERE Z."UgpEntry" = T1."UgpEntry" AND Y."UomEntry" = B."UomEntry") AS "ConversionRate"
 , (SELECT "GET_COST"(T1."ItemCode", B."UomEntry",V."U_PricelistNo") FROM DUMMY) AS "COST"
 , (SELECT "GET_RSP"(T1."ItemCode", B."UomEntry",V."U_PricelistNo") FROM DUMMY) AS "RSP"
 , (CASE WHEN   IFNULL(T1."ItemCode",'')='' THEN 'Barcode Not Found/ Item Not found' ELSE '' END) AS "REMARK",v."UPDT" As "StartDate"
FROM  (Select * FROM   "TNX_PRICEUPDATE", OBPL)v
left JOIN OBCD B ON B."BcdCode" = v."BARCODE"
LEFT JOIN OITM T1 on  T1."ItemCode" = B."ItemCode"
left join OVTG T2 on T2."Code" = T1."VatGourpSa"
left join OPLN T3 on T3."ListNum"= v."U_PricelistNo"
WHERE "CREATEDBY"=:USERNAME AND IFNULL(v."PRICELIST",'')=''
AND IFNULL(V."U_PricelistNo",'') <>''
AND V."BPLId" NOT IN ('12','15') 
)
ORDER BY "CREATEDBY","LINENUM","BARCODE", "PRICELIST";


END;