/*Select
	 T2."DocDate" 
from OIGE T2 
where T2."DocDate" >= [%0] 
and T2."DocDate" <= [%1] */ 
/*SELECT
	 T2."DocDate",
	 T2."DocNum" as "Dispatch No",
	 T3."WhsCode" as "FromWhsCode",
	 (Select
	 "WhsName" 
	from OWHS 
	where "WhsCode"=T3."WhsCode") as "FromWhsName",
	 T3."U_METWHSCD",
	 T3."U_METWHSNM",
	 T2."Comments" as "Dispatch Remarks",
	 T2."U_MESTS" as "Dispatch Status",
	 T3."CodeBars",
	 T3."ItemCode",
	 T3."Dscription",
	 T3."Quantity" as "Dispatch Qty" ,
	 T3."OpenQty" as "Open Dispatch Qty",
	 TO_CHAR(T4."DocNum",
	 '') as "Receipt No",
	 T5."Quantity" as "Transfer Receipt Qty",
	 T4."Comments" as "Receipt Remarks" 
FROM OIGE T2 
LEFT JOIN IGE1 T3 ON T3."DocEntry" = T2."DocEntry" 
LEFT JOIN OIGN T4 on T4."DocEntry"=T2."U_MIDENT" 
LEFT JOIN IGN1 T5 on T5."DocEntry"=T4."DocEntry" 
and T5."ItemCode"=T3."ItemCode" 
WHERE IFNULL(T5."ItemCode",
	 '')='' 
and T2."DocDate">=[%0] 
and T2."DocDate"<=[%1] 
and T2."U_TYPE" in ('STR','Transfer') 
--IFNULL(T2."U_MIDNUM",'')='' and TO_CHAR(T2."UpdateDate",'YYYYMMDD')||LPAD(T2."UpdateTS",6,0)<='20240105124000'  

UNION ALL */
CREATE PROCEDURE "TNX_OpenDispatchReport"(IN FromDate Date,IN ToDate DATE)
As
Begin
SELECT
	 T2."DocDate",
	T2."DocNum" as "Dispatch No",
	T3."WhsCode" as "FromWhsCode",
	 (Select
	 "WhsName" 
	from OWHS 
	where "WhsCode"=T3."WhsCode") as "FromWhsName",
	T3."U_METWHSCD",
	 T3."U_METWHSNM",
	 T2."Comments" as "Dispatch Remarks",
	T2."U_MESTS" as "Dispatch Status",
	 T3."CodeBars",
	 T3."ItemCode",
	 T3."Dscription",
	 T3."Quantity" as "Dispatch Qty" ,
	 T3."OpenQty" as "Open Dispatch Qty",
	 STRING_AGG(T4."DocNum",
	',') as "Receipt No",
	 (T3."Quantity"-T3."OpenQty") as "Transfer Receipt Qty",
	 STRING_AGG(T4."Comments",
	',') as "Receipt Remarks" , T3."Quantity"-IFNULL(T5."Quantity",0) as "Variance"
FROM OIGE T2 
LEFT JOIN IGE1 T3 ON T3."DocEntry" = T2."DocEntry" 
LEFT JOIN IGN1 T5 on T5."BaseEntry"=T3."DocEntry" 
and T3."LineNum"=T5."BaseLine" 
and T3."ObjType"=T5."BaseType" 
LEFT JOIN OIGN T4 on T4."DocEntry"=T5."DocEntry" 
WHERE IFNULL(T3."OpenQty",
	0) <>0 
and T2."DocDate">=:FromDate
and T2."DocDate"<=:ToDate
and T2."U_TYPE" in ('STR','Transfer') 
and IFNULL(T3."U_METWHSCD",'')<>''

GROUP BY T2."DocDate",
	T2."DocNum" ,
	T3."WhsCode",
	 T3."U_METWHSCD",
	 T3."U_METWHSNM",
	 T2."Comments" ,
	T2."U_MESTS" ,
	 T3."CodeBars",
	 T3."ItemCode",
	 T3."Dscription",
	 T3."Quantity" ,T5."Quantity",
	 T3."OpenQty" 
order by T2."DocDate" ASC;
END;