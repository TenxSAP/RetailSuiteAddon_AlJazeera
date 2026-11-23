CREATE Procedure "TENX_MESSENGER_MAIN"(
IN METHOD NVARCHAR(100),
IN DOCTYPE NVARCHAR(100),
IN DOCNUM NVARCHAR(100) DEFAULT NULL,
IN SERIES NVARCHAR(100) DEFAULT NULL
)
AS
BEGIN

If  METHOD='WhatsAppUserEmp' Then
	SELECT (SELECT "PortNum" FROM OUSR WHERE TO_CHAR("USERID")=T1."U_UserID") As "Mobile" FROM "@MGRCONFIG"  T0  INNER JOIN "@MGRCONFIG_C0"  T1 ON T0."DocEntry"=T1."DocEntry"  
	Where T0."U_Doc"=:DOCTYPE and T0."U_Status"='Y' AND IFNULL(T0."U_User",'N')<>'N' and IFNULL(T1."U_UserID",'')<>''  
	UNION ALL  
	SELECT (SELECT "email" FROM OHEM WHERE TO_CHAR("empID")=T1."U_EmpID") As "Mobile" FROM "@MGRCONFIG"  T0  INNER JOIN "@MGRCONFIG_C0"  T1 ON T0."DocEntry"=T1."DocEntry"  
	Where T0."U_Doc"=:DOCTYPE and T0."U_Status"='Y' AND IFNULL(T0."U_Emp",'N')<>'N' and IFNULL(T1."U_EmpID",'')<>'' ;
End If;

If  METHOD='SMSUserEmp' Then
	SELECT (SELECT "PortNum" FROM OUSR WHERE TO_CHAR("USERID")=T1."U_UserID") As "Mobile" FROM "@MGRCONFIG"  T0  INNER JOIN "@MGRCONFIG_C0"  T1 ON T0."DocEntry"=T1."DocEntry"  
	Where T0."U_Doc"=:DOCTYPE and T0."U_Status"='Y' AND IFNULL(T0."U_UserSMS",'N')<>'N' and IFNULL(T1."U_UserID",'')<>''  
	UNION ALL  
	SELECT (SELECT "email" FROM OHEM WHERE TO_CHAR("empID")=T1."U_EmpID") As "Mobile" FROM "@MGRCONFIG"  T0  INNER JOIN "@MGRCONFIG_C0"  T1 ON T0."DocEntry"=T1."DocEntry"  
	Where T0."U_Doc"=:DOCTYPE and T0."U_Status"='Y' AND IFNULL(T0."U_EmpSMS",'N')<>'N' and IFNULL(T1."U_EmpID",'')<>'' ;
End If;

If  METHOD='EmailUserEmp' Then
	SELECT (SELECT "E_Mail" FROM OUSR WHERE TO_CHAR("USERID")=T1."U_UserID") As "Mobile" FROM "@MGRCONFIG"  T0  INNER JOIN "@MGRCONFIG_C0"  T1 ON T0."DocEntry"=T1."DocEntry"  
	Where T0."U_Doc"=:DOCTYPE and T0."U_Status"='Y' AND IFNULL(T0."U_UserEML",'N')<>'N' and IFNULL(T1."U_UserID",'')<>''  
	UNION ALL  
	SELECT (SELECT "email" FROM OHEM WHERE TO_CHAR("empID")=T1."U_EmpID") As "Mobile" FROM "@MGRCONFIG"  T0  INNER JOIN "@MGRCONFIG_C0"  T1 ON T0."DocEntry"=T1."DocEntry"  
	Where T0."U_Doc"=:DOCTYPE and T0."U_Status"='Y' AND IFNULL(T0."U_EmpEML",'N')<>'N' and IFNULL(T1."U_EmpID",'')<>'' ;
End If;

If  METHOD='MarketingWhatsAppMessage_Draft' Then
	SELECT B."Cellular" FROM ODRF A INNER JOIN OCRD B ON B."CardCode"=A."CardCode" WHERE A."DocNum"=:DOCNUM AND A."Series"=:SERIES AND A."DocStatus"='O';
End If;


If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='SO' Then	
	SELECT T0."DocEntry" FROM ORDR T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;
End If;

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='PO' Then	
	SELECT T0."DocEntry" FROM OPOR T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;		
End If;

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='SQ' Then	
	SELECT T0."DocEntry" FROM OQUT T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;		
End If;

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='DL' Then	
	SELECT T0."DocEntry" FROM ODLN T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;		
End If;

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='ARINV' Then	
	SELECT T0."DocEntry" FROM OINV T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;		
End If;

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='ARC' Then	
	SELECT T0."DocEntry" FROM ORIN T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;		
End If;

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='ARRET' Then	
	SELECT T0."DocEntry" FROM ORDN T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;	
End If;

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='PQ' Then	
	SELECT T0."DocEntry" FROM OPQT T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;		
End If;

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='APINV' Then	
	SELECT T0."DocEntry" FROM OPCH T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;		
End If;

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='GRPO' Then	
	SELECT T0."DocEntry" FROM OPDN T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;		
End If;

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='GR' Then	
	SELECT T0."DocEntry" FROM ORPD T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;		
End If; 	

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='APC' Then
	SELECT T0."DocEntry" FROM ORPC T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;
End If; 

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='IP' Then	
	SELECT T0."DocEntry" FROM ORCT T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;
End If; 

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='OP' Then	
	SELECT T0."DocEntry" FROM OVPM T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;
End If; 

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='GRP' Then	
	SELECT T0."DocEntry" FROM OIGN T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;
End If; 

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='GI' Then	
	SELECT T0."DocEntry" FROM OIGE T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;
End If; 

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='ITR' Then	
	SELECT T0."DocEntry" FROM OWTQ T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;
End If; 

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='IT' Then	
	SELECT T0."DocEntry" FROM OWTR T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES;
End If; 

If  METHOD='MarketingWhatsAppMessage' AND DOCTYPE='JE' Then	
	SELECT T0."TransId" "DocEntry" FROM OJDT T0 WHERE T0."Number"=:DOCNUM and T0."Series"=:SERIES;
End If;



If  METHOD='MarketingSMSMessage' AND DOCTYPE='SO' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From RDR1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."Cellular") As "Mobile" 
	FROM ORDR  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SMS"='Yes';	
End If;

If  METHOD='MarketingSMSMessage' AND DOCTYPE='PO' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From POR1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."Cellular") As "Mobile" 
	FROM OPOR  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SMS"='Yes';	
End If;

If  METHOD='MarketingSMSMessage' AND DOCTYPE='SQ' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From QUT1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."Cellular") As "Mobile" 
	FROM OQUT  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SMS"='Yes';	
End If;

If  METHOD='MarketingSMSMessage' AND DOCTYPE='DL' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From DLN1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."Cellular") As "Mobile" 
	FROM ODLN  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SMS"='Yes';	
End If;

If  METHOD='MarketingSMSMessage' AND DOCTYPE='ARINV' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From INV1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."Cellular") As "Mobile" 
	FROM OINV  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SMS"='Yes';	
End If;

If  METHOD='MarketingSMSMessage' AND DOCTYPE='ARC' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From RIN1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."Cellular") As "Mobile" 
	FROM ORIN  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SMS"='Yes';	
End If;

If  METHOD='MarketingSMSMessage' AND DOCTYPE='ARRET' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From RDN1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."Cellular") As "Mobile" 
	FROM ORDN  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SMS"='Yes';	
End If;

If  METHOD='MarketingSMSMessage' AND DOCTYPE='PQ' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From PQT1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."Cellular") As "Mobile" 
	FROM OPQT  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SMS"='Yes';	
End If;

If  METHOD='MarketingSMSMessage' AND DOCTYPE='APINV' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From PCH1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."Cellular") As "Mobile" 
	FROM OPCH  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SMS"='Yes';	
End If;

If  METHOD='MarketingSMSMessage' AND DOCTYPE='GRPO' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From PDN1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."Cellular") As "Mobile" 
	FROM OPDN  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SMS"='Yes';	
End If;

If  METHOD='MarketingSMSMessage' AND DOCTYPE='GR' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From RPD1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."Cellular") As "Mobile" 
	FROM ORPD  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SMS"='Yes';	
End If; 	

If  METHOD='MarketingSMSMessage' AND DOCTYPE='APC' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From RPC1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."Cellular") As "Mobile" 
	FROM ORPC  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SMS"='Yes';	
End If; 

If  METHOD='MarketingSMSMessage' AND DOCTYPE='IP' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."Cellular" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Mobile",0 "DiscSum",0 "ItemCount" 
	FROM "ORCT"  T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES ;
End If; 

If  METHOD='MarketingSMSMessage' AND DOCTYPE='OP' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."Cellular" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Mobile",0 "DiscSum",0 "ItemCount" 
	FROM "OVPM"  T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES ;
End If; 

If  METHOD='MarketingSMSMessage' AND DOCTYPE='GRP' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."Cellular" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Mobile",0 "DiscSum",0 "ItemCount" 
	FROM "OIGN"  T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES ;
End If; 

If  METHOD='MarketingSMSMessage' AND DOCTYPE='GI' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."Cellular" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Mobile",0 "DiscSum",0 "ItemCount" 
	FROM "OIGE"  T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES ;
End If; 

If  METHOD='MarketingSMSMessage' AND DOCTYPE='ITR' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."Cellular" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Mobile",0 "DiscSum",0 "ItemCount" 
	FROM "OWTQ"  T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES ;
End If; 

If  METHOD='MarketingSMSMessage' AND DOCTYPE='IT' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."Cellolar" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."Cellular" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Mobile",0 "DiscSum",0 "ItemCount" 
	FROM "OWTR"  T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES ;
End If; 

If  METHOD='MarketingSMSMessage' AND DOCTYPE='JE' Then	
	SELECT T0."Number" "DocNum",T0."TransId" "DocEntry", T0."RefDate" "TaxDate", '' "CardCode",'' "CardName", 
	(SELECT SUM(X."Debit") FROM JDT1 X WHERE X."TransId"=T0."TransId") "DocTotal",0 "VatSum" ,'' As "Mobile",0 "DiscSum",0 "ItemCount" 
	FROM "OJDT"  T0 WHERE T0."Number"=:DOCNUM and T0."Series"=:SERIES ;
End If; 

If  METHOD='MarketingEmail' AND DOCTYPE='SO' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From RDR1 Where "DocEntry"=T0."DocEntry") As "ItemCount" , IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM ORDR  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SNEMAIL"='Yes';	
End If;

If  METHOD='MarketingEmail' AND DOCTYPE='PO' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From POR1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM OPOR  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SNEMAIL"='Yes';	
End If;

If  METHOD='MarketingEmail' AND DOCTYPE='SQ' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From QUT1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM OQUT  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SNEMAIL"='Yes';	
End If;

If  METHOD='MarketingEmail' AND DOCTYPE='DL' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From DLN1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM ODLN  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SNEMAIL"='Yes';	
End If;

If  METHOD='MarketingEmail' AND DOCTYPE='ARINV' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From INV1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM OINV  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SNEMAIL"='Yes';	
End If;

If  METHOD='MarketingEmail' AND DOCTYPE='ARC' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From RIN1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM ORIN  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SNEMAIL"='Yes';	
End If;

If  METHOD='MarketingEmail' AND DOCTYPE='ARRET' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From RDN1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM ORDN  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SNEMAIL"='Yes';	
End If;

If  METHOD='MarketingEmail' AND DOCTYPE='PQ' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From PQT1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM OPQT  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SNEMAIL"='Yes';	
End If;

If  METHOD='MarketingEmail' AND DOCTYPE='APINV' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From PCH1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM OPCH  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SNEMAIL"='Yes';	
End If;

If  METHOD='MarketingEmail' AND DOCTYPE='GRPO' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From PDN1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM OPDN  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SNEMAIL"='Yes';	
End If;

If  METHOD='MarketingEmail' AND DOCTYPE='GR' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From RPD1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM ORPD  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SNEMAIL"='Yes';	
End If; 	

If  METHOD='MarketingEmail' AND DOCTYPE='APC' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", T0."DiscSum" ,
	(Select Count(*) From RPC1 Where "DocEntry"=T0."DocEntry") As "ItemCount" ,IFNULL((SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),C0."E_Mail") As "Email" 
	FROM ORPC  T0 INNER JOIN OCRD C0 ON C0."CardCode"=T0."CardCode" 
	WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES and "U_SNEMAIL"='Yes';	
End If; 

If  METHOD='MarketingEmail' AND DOCTYPE='IP' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."E_Mail" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Email" ,0 "DiscSum",0 "ItemCount" 
	FROM "ORCT"  T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES ;
End If; 

If  METHOD='MarketingEmail' AND DOCTYPE='OP' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."E_Mail" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Email" ,0 "DiscSum",0 "ItemCount" 
	FROM "OVPM"  T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES ;
End If; 

If  METHOD='MarketingEmail' AND DOCTYPE='GRP' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."E_Mail" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Email" ,0 "DiscSum",0 "ItemCount" 
	FROM "OIGN"  T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES ;
End If; 

If  METHOD='MarketingEmail' AND DOCTYPE='GI' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."E_Mail" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Email" ,0 "DiscSum",0 "ItemCount" 
	FROM "OIGE"  T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES ;
End If; 

If  METHOD='MarketingEmail' AND DOCTYPE='ITR' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."E_Mail" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Email" ,0 "DiscSum",0 "ItemCount" 
	FROM "OWTQ"  T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES ;
End If; 

If  METHOD='MarketingEmail' AND DOCTYPE='IT' Then	
	SELECT T0."DocNum",T0."DocEntry", T0."TaxDate", T0."CardCode",T0."CardName", T0."DocTotal", T0."VatSum", 
	IFNULL( (SELECT X."E_MailL" FROM OCPR X WHERE X."CntctCode"=T0."CntctCode"),(SELECT X."E_Mail" FROM OCRD X WHERE X."CardCode"=T0."CardCode")) As "Email" ,0 "DiscSum",0 "ItemCount" 
	FROM "OWTR"  T0 WHERE T0."DocNum"=:DOCNUM and T0."Series"=:SERIES ;
End If; 

If  METHOD='MarketingEmail' AND DOCTYPE='JE' Then	
	SELECT T0."Number" "DocNum",T0."TransId" "DocEntry", T0."RefDate" "TaxDate", '' "CardCode",'' "CardName", 
	(SELECT SUM(X."Debit") FROM JDT1 X WHERE X."TransId"=T0."TransId") "DocTotal",0 "VatSum" ,'' As "Email",0 "DiscSum",0 "ItemCount" 
	FROM "OJDT"  T0 WHERE T0."Number"=:DOCNUM and T0."Series"=:SERIES ;
End If; 

END;