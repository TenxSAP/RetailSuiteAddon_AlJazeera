CREATE PROCEDURE "TNX_BarcodeCheck"(IN Barcode VARCHAR(30))
As
Begin
SELECT AC."ItemCode",AV."BcdCode",AC."ItemName",AB."UomCode", (Select "BaseQty"/"AltQty" from UGP1 where "UgpEntry"=AC."UgpEntry" and AV."UomEntry"="UomEntry" ) as "CF",
AZ."LPP", AZ."DocDate" as "Last Purchase Date", AZ."CardCode" as "Last Purchase Vendor", AZ."CardName" as "Last Purchase Vendor Name"



FROM OITM AC LEFT JOIN OBCD AV ON AV."ItemCode"=AC."ItemCode" LEFT JOIN OUOM AB on AB."UomEntry"=AV."UomEntry" 
LEFT JOIN (Select (case when B."UomEntry"=X."UomEntry"  and X."UomEntry"=Y."BaseUom" then B."Price" when B."UomEntry"=X."UomEntry"  and X."UomEntry"<>Y."BaseUom" then B."Price"*1/X."AltQty"*X."BaseQty" when B."UomEntry"<>X."UomEntry" and X."UomEntry"<>Y."BaseUom" then B."Price"*1/X."AltQty"*X."BaseQty" when B."UomEntry"<>X."UomEntry" and X."UomEntry"=Y."BaseUom" then B."Price"else 0 end) as "LPP",
B."DocDate", B."CardCode",B."CardName", B."ItemCode",X."UomEntry"
from OITM C inner join (SELECT A3."DocDate", A3."CardCode", A3."CardName",H."ItemCode",IFNULL((IFNULL(H."INMPrice", 0)/H."NumPerMsr"),0)-IFNULL((IFNULL(A1."INMPrice", 0)/A1."NumPerMsr"),0) as "Price", H."UomCode",H."UomEntry",H."NumPerMsr",Row_Number() over (partition by H."ItemCode" order by H."DocDate" Desc) as "RN" FROM "PDN1" H left join OPDN A3 on A3."DocEntry"=H."DocEntry" left join ORPC A2 on  A2."U_GRPO"=A3."DocNum"left join RPC1 A1 on A2."DocEntry"=A1."DocEntry" and A1."ItemCode"=H."ItemCode" and A1."NoInvtryMv"='Y' where IFNULL(H."U_normalqty",0)<>0)B on C."ItemCode"=B."ItemCode" left join OUGP Y on Y."UgpEntry"=C."UgpEntry"  left JOIN UGP1 X ON Y."UgpEntry" = X."UgpEntry" where "RN"=1 )AZ on AZ."ItemCode" = AV."ItemCode" and  AZ."UomEntry"= AV."UomEntry"

WHERE AC."ItemCode" in (select T."ItemCode" from OBCD T where T."BcdCode"=:BarCode);
END;