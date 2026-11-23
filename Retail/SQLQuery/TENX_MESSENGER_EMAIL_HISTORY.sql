CREATE Procedure "TENX_MESSENGER_EMAIL_HISTORY"(
IN METHOD NVARCHAR(100),
IN DOCTYPE NVARCHAR(100),
IN DOCENTRY NVARCHAR(100) DEFAULT NULL,
IN FROMDATE DATE DEFAULT '19900101',
IN TODATE DATE DEFAULT '21301231'
)
AS
BEGIN


If METHOD='FormLoad' Then	
	SELECT * FROM (SELECT DISTINCT T0."U_DocType" AS "Document Type", T0."U_Date" AS "Document Date", T0."U_DocEntry" As "Document Entry", 
	T0."U_DocNum" As "Document No",T0."U_CardCode" As "Customer", T0."U_EmailID" As "Email ID", T0."U_Status" As "Status" 
	FROM "@SBO_MSGLOG" T0 Where T0."U_MsgType"='Email' AND T0."U_Date" BETWEEN :FROMDATE and :TODATE)X ORDER BY X."Document Date";
End If; 
If METHOD='ItemEvent_Success' Then	
	SELECT * FROM (SELECT DISTINCT T0."U_DocType" AS "Document Type", T0."U_Date" AS "Document Date", T0."U_DocEntry" As "Document Entry", 
	T0."U_DocNum" As "Document No",T0."U_CardCode" As "Customer", T0."U_EmailID" As "Email ID", T0."U_Status" As "Status" 
	FROM "@SBO_MSGLOG" T0 Where T0."U_MsgType"='Email' AND T0."U_Status"='Success' AND T0."U_Date" BETWEEN :FROMDATE and :TODATE)X ORDER BY X."Document Date";
End If; 
If METHOD='ItemEvent_Failed' Then	
	SELECT * FROM (SELECT DISTINCT '' as "Check",T0."U_DocType" AS "Document Type", T0."U_Date" AS "Document Date", T0."U_DocEntry" As "Document Entry", 
	T0."U_DocNum" As "Document No",T0."U_CardCode" As "Customer", T0."U_EmailID" As "Email ID", T0."U_Status" As "Status" 
	FROM "@SBO_MSGLOG" T0 Where T0."U_MsgType"='Email' AND T0."U_Status"='Failed' AND T0."U_Date" BETWEEN :FROMDATE and :TODATE)X ORDER BY X."Document Date";
End If; 



If METHOD='MarketingEmail' AND DOCTYPE='SO' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From RDR1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM ORDR  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocEntry"=:DOCENTRY and "U_SNEMAIL"='Yes';	
End If;

If METHOD='MarketingEmail' AND DOCTYPE='PO' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From POR1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM OPOR  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocEntry"=:DOCENTRY and "U_SNEMAIL"='Yes';	
End If;

If METHOD='MarketingEmail' AND DOCTYPE='SQ' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From QUT1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM OQUT  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocEntry"=:DOCENTRY and "U_SNEMAIL"='Yes';	
End If;

If METHOD='MarketingEmail' AND DOCTYPE='DL' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From DLN1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM ODLN  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocEntry"=:DOCENTRY and "U_SNEMAIL"='Yes';	
End If;

If METHOD='MarketingEmail' AND DOCTYPE='ARINV' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From INV1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM OINV  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocEntry"=:DOCENTRY and "U_SNEMAIL"='Yes';	
End If;

If METHOD='MarketingEmail' AND DOCTYPE='ARC' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From RIN1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM ORIN  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocEntry"=:DOCENTRY and "U_SNEMAIL"='Yes';	
End If;

If METHOD='MarketingEmail' AND DOCTYPE='ARRET' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From RDN1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM ORDN  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocEntry"=:DOCENTRY and "U_SNEMAIL"='Yes';	
End If;

If METHOD='MarketingEmail' AND DOCTYPE='PQ' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From PQT1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM OPQT  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocEntry"=:DOCENTRY and "U_SNEMAIL"='Yes';	
End If;

If METHOD='MarketingEmail' AND DOCTYPE='APINV' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From PCH1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM OPCH  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocEntry"=:DOCENTRY and "U_SNEMAIL"='Yes';	
End If;

If METHOD='MarketingEmail' AND DOCTYPE='GRPO' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From PDN1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM OPDN  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocEntry"=:DOCENTRY and "U_SNEMAIL"='Yes';	
End If;

If METHOD='MarketingEmail' AND DOCTYPE='GR' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From RPD1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM ORPD  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocEntry"=:DOCENTRY and "U_SNEMAIL"='Yes';	
End If; 	

If METHOD='MarketingEmail' AND DOCTYPE='APC' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From RPC1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM ORPC  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocEntry"=:DOCENTRY and "U_SNEMAIL"='Yes';	
End If; 

If METHOD='MarketingEmail' AND DOCTYPE='IP' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."E_Mail" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Email",0 "DiscSum",0 "ItemCount" 
	FROM "ORCT"  T0 WHERE T0."DocEntry"=:DOCENTRY;
End If; 

If METHOD='MarketingEmail' AND DOCTYPE='OP' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."E_Mail" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Email",0 "DiscSum",0 "ItemCount" 
	FROM "OVPM"  T0 WHERE T0."DocEntry"=:DOCENTRY;
End If; 

If METHOD='MarketingEmail' AND DOCTYPE='GRP' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."E_Mail" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Email",0 "DiscSum",0 "ItemCount" 
	FROM "OIGN"  T0 WHERE T0."DocEntry"=:DOCENTRY;
End If; 

If METHOD='MarketingEmail' AND DOCTYPE='GI' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."E_Mail" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Email",0 "DiscSum",0 "ItemCount" 
	FROM "OIGE"  T0 WHERE T0."DocEntry"=:DOCENTRY;
End If; 

If METHOD='MarketingEmail' AND DOCTYPE='ITR' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."E_Mail" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Email",0 "DiscSum",0 "ItemCount" 
	FROM "OWTQ"  T0 WHERE T0."DocEntry"=:DOCENTRY;
End If; 

If METHOD='MarketingEmail' AND DOCTYPE='IT' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."E_Mail" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Email",0 "DiscSum",0 "ItemCount" 
	FROM "OWTR"  T0 WHERE T0."DocEntry"=:DOCENTRY;
End If; 

If METHOD='MarketingEmail' AND DOCTYPE='JE' Then	
	SELECT T0."Number" "DocNum",T0."TransId" "DocEntry", T0."RefDate" "TaxDate", '' "CardCode",'' "CardName", 
	(SELECT SUM(X."Debit") FROM JDT1 X WHERE X."TransId"=T0."TransId") "DocTotal",0 "VatSum" ,'' As "Email",0 "DiscSum",0 "ItemCount" 
	FROM "OJDT"  T0 WHERE T0."TransId"=:DOCENTRY ;
End If; 


END;