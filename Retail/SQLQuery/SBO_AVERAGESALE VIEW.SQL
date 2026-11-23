
CREATE FUNCTION  SBO_AverageSale()
      RETURNS TABLE 
   ( 
  
	ItemCode VARCHAR(215),
	ItemName  VARCHAR(215),
	Barcode VARCHAR(215),
	ItemArabicDescription  VARCHAR(215),
	QuartelySalesQuantity  Varchar(215),
	Halfyearlysalesquantity  Varchar(215),
	AnnualSalesQuantity Varchar(215),
	MonthlyAverageSalesQuantity Varchar(215)
	)
	LANGUAGE SQLSCRIPT
	SQL SECURITY INVOKER 
	AS
BEGIN
Return

Select T1."ItemCode" as "ITEMCODE",
T1."ItemName" as "ITEMNAME",
"CodeBars" as"BARCODE",
"FrgnName"as"ITEMARABICDESCRIPTION",

(SELECT SUM(T3."InvQty") as "Quartely  QTY" FROM OINV T2  INNER JOIN INV1 T3 ON T2."DocEntry" = T3."DocEntry"
  WHERE Month(T2."DocDate" ) Between '1' and '4' and Year(T2."DocDate")= YEAR(CURRENT_DATE)and T3."ItemCode"=T1."ItemCode" Group By T3."ItemCode")
 AS "QUARTELYSALESQUANTITY",
 
 (SELECT SUM(T5."InvQty") as "half QTY" FROM OINV T4  INNER JOIN INV1 T5 ON T4."DocEntry" = T5."DocEntry"
  WHERE Month(T4."DocDate" ) Between '1' and '6' and Year(T4."DocDate")= YEAR(CURRENT_DATE) and T5."ItemCode"=T1."ItemCode"Group By T5."ItemCode")
   AS "HALFYEARLYSALESQUANTITY",
   
   (SELECT SUM(T7."InvQty") as "annual  QTY" FROM OINV T6  INNER JOIN INV1 T7 ON T6."DocEntry" = T7."DocEntry"
  WHERE Month(T6."DocDate" ) Between '1' and '12' and Year(T6."DocDate")= YEAR(CURRENT_DATE)and T7."ItemCode"=T1."ItemCode" Group By T7."ItemCode")
   AS "ANNUALSALESQUANTITY",

   (SELECT SUM(T9."InvQty") as "Monthly Average QTY" FROM OINV T8 INNER JOIN INV1 T9 ON T8."DocEntry" = T9."DocEntry"
  WHERE T9."ItemCode"=T1."ItemCode" Group By T9."ItemCode")/MONTHS_BETWEEN((select MIN("DocDate") from OINM where "ItemCode"=T1."ItemCode" and "InQty">0 ),CURRENT_DATE)
   AS "MONTHLYAVERAGESALESQUANTITY"
  
  from OITM T1;
 
END;