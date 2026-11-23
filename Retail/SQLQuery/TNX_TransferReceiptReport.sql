CREATE PROCEDURE "TNX_TransferReceiptReport"(IN FromDate Date)
As
Begin

SELECT
	 A."DocNum",
	 A."U_MEPSDT",
	 A."U_MEREQBY",
	 A."U_MEFWHSNM",
	 A."U_METWHSNM",
	 A."U_MEBRCD",
	 A."U_MEITMCD",
	 A."U_MEITNM",
	 SUM("STR Qty") "STR Qty" ,
	 STRING_AGG("Dispatch No.s",
	 ',') "Dispatch No.s",
	 SUM("Dispatched Qty") "Dispatched Qty" ,
	 STRING_AGG("Receipt No",
	 ',') "Receipt No",
	 SUM("Transfer Receipt Qty") "Transfer Receipt Qty" ,
SUM("Dispatched Qty") -  SUM("Transfer Receipt Qty")  as "Variance"

FROM (SELECT
	 T0."DocNum",
	 T0."U_MEPSDT",
	 T0."U_MEREQBY",
	 T0."U_MEFWHSCD",
	 T0."U_MEFWHSNM",
	 T0."U_MEFBPID",
	 T0."U_METWHSCD",
	 T0."U_METWHSNM",
	 T0."U_METBPID",
	 T1."U_MEBRCD",
	 T1."U_MEITMCD",
	 T1."U_MEITNM",
	 (case when row_number() over (partition by T1."DocEntry",
	 T1."LineId" 
			order by T1."DocEntry",
	 T1."LineId") <>'1' 
		then 0 
		else T1."U_MEREQQTY" 
		end) as "STR Qty" ,
	 (case when IFNULL(A3."DocNum",
	 0)=0 
		then (case when row_number() over (partition by A4."DocEntry",
	 A4."LineNum" 
				order by A4."DocEntry",
	 A4."LineNum") <>'1' 
			then NULL 
			else A4."DocNum" 
			end) 
		else A3."DocNum" 
		end) as "Dispatch No.s",
	 (case when IFNULL(A3."Quantity",
	 0)=0 
		then (case when row_number() over (partition by A4."DocEntry",
	 A4."LineNum" 
				order by A4."DocEntry",
	 A4."LineNum") <>'1' 
			then 0 
			else IFNULL(A4."Quantity",0)
			end) 
		else IFNULL(A3."Quantity",0)
		end) as "Dispatched Qty" ,
	 (case when IFNULL(T4."DocNum",
	 0)=0 
		then (case when row_number() over (partition by T6."DocEntry",
	 T6."LineNum" 
				order by T6."DocEntry",
	 T6."LineNum") <>'1' 
			then NULL 
			else T7."DocNum" 
			end) 
		else T4."DocNum" 
		end) as "Receipt No",
	 (case when IFNULL(T5."Quantity",
	 0)=0 
		then (case when row_number() over (partition by T6."DocEntry",
	 T6."LineNum" 
				order by T6."DocEntry",
	 T6."LineNum") <>'1' 
			then NULL 
			else IFNULL(T6."Quantity",0)
			end) 
		else IFNULL(T5."Quantity" ,0)
		end) as "Transfer Receipt Qty" 
	FROM "@ME_STR" T0 
	INNER JOIN "@ME_STR_C0" T1 on T0."DocEntry"=T1."DocEntry" 
	LEFT JOIN (Select
	 T2."DocEntry",
	 T2."DocNum",
	 T2."U_INVTRREN",
	 T2."U_MIDENT",
	 T3."CodeBars",
	 T3."Quantity" 
		from OIGE T2 
		INNER JOIN IGE1 T3 on T3."DocEntry"=T2."DocEntry" 
		where TO_CHAR(T2."UpdateDate",
	 'YYYYMMDD')||LPAD(T2."UpdateTS",
	 6,
	 0) <='20240105124000')A3 ON T0."DocEntry" = A3."U_INVTRREN" 
	and T1."U_MEBRCD"=A3."CodeBars" 
	LEFT JOIN OIGN T4 on T4."DocEntry"=A3."U_MIDENT" 
	LEFT JOIN IGN1 T5 on T5."DocEntry"=T4."DocEntry" 
	and T5."CodeBars"=A3."CodeBars" 
	LEFT JOIN (Select
	 T2."DocEntry",
	 T2."DocNum",
	 T2."U_INVTRREN",
	 T2."U_MIDENT",
	 T3."CodeBars",
	 T3."Quantity" ,
	 T3."OpenQty",
	 T3."LineNum",
	 T2."ObjType" 
		from OIGE T2 
		INNER JOIN IGE1 T3 on T3."DocEntry"=T2."DocEntry" 
		where TO_CHAR(T2."UpdateDate",
	 'YYYYMMDD')||LPAD(T2."UpdateTS",
	 6,
	 0) >'20240105124000')A4 ON T0."DocEntry" = A4."U_INVTRREN" 
	and T1."U_MEBRCD"=A4."CodeBars" 
	LEFT JOIN IGN1 T6 on T6."BaseEntry"=A4."DocEntry" 
	and T6."BaseLine"=A4."LineNum" 
	and T6."BaseType"=A4."ObjType" 
	LEFT JOIN OIGN T7 on T6."DocEntry"=T7."DocEntry" 
	WHERE T0."U_MEPSDT" = :FromDate
	and T1."U_MEBRCD"<>'' -- and T0."DocNum"='7817' 
)A 
GROUP BY A."DocNum",
	 A."U_MEPSDT",
	 A."U_MEREQBY",
	 A."U_MEFWHSNM",
	 A."U_METWHSNM",
	 A."U_MEBRCD",
	 A."U_MEITMCD",
	 A."U_MEITNM";
	 END;