CREATE Procedure "TNX_APInvoiceAllocation"(IN DocNum VARCHAR(20),
	IN Series VARCHAR(20)) As 
Begin SELECT
	 T1."AcctCode",
	 T1."LineTotal" As "Debit",
	0 As "Credit",
	T1."U_Branch" as "BPLId" 
FROM OPCH T0 
INNER JOIN PCH1 T1 ON T0."DocEntry" = T1."DocEntry" 
WHERE T0."DocType" ='S' 
and T0."DocNum"=:DocNum 
UNION ALL SELECT
	 T1."AcctCode",
	 0 As "Debit",
	SUM(T1."LineTotal") As "Credit",
	T0."BPLId" 
FROM OPCH T0 
INNER JOIN PCH1 T1 ON T0."DocEntry" = T1."DocEntry" 
WHERE T0."DocType" ='S' 
and T0."DocNum"=:DocNum 
Group By T1."AcctCode",
	T0."BPLId" 
UNION ALL SELECT
	 (Select
	 "PmtClrAct" 
	From OBPL 
	Where "BPLId"=T0."BPLId"),
	 0 As "Debit",
	T1."LineTotal" As "Credit",
	T1."U_Branch" 
FROM OPCH T0 
INNER JOIN PCH1 T1 ON T0."DocEntry" = T1."DocEntry" 
WHERE T0."DocType" ='S' 
and T0."DocNum"=:DocNum 
UNION ALL SELECT
	 (Select
	 "PmtClrAct" 
	From OBPL 
	Where "BPLId"=T0."BPLId"),
	SUM(T1."LineTotal") As "Debit",
	0 As "Credit",
	T0."BPLId" 
FROM OPCH T0 
INNER JOIN PCH1 T1 ON T0."DocEntry" = T1."DocEntry" 
WHERE T0."DocType" ='S' 
and T0."DocNum"=:DocNum 
Group By T1."AcctCode",
	T0."BPLId" 
;    End;