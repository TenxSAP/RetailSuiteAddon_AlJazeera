CREATE PROCEDURE "SBO_ProcessBarCode"(IN ItemCode varchar(100))                          
AS 
ItemGrpName varchar(150);
Type1 varchar(5);
Itmgrpwise varchar(1);
temp_var_0 integer;
temp_var_1 integer;
temp_var_2 integer;
temp_var_3 integer;
temp_var_4 integer;
BEGIN

SELECT  (SELECT "ItmsGrpNam" FROM OITB WHERE "ItmsGrpCod" = (SELECT "ItmsGrpCod" FROM OITM WHERE "ItemCode" = :ItemCode)) INTO ItemGrpName   FROM DUMMY;
--SELECT (SELECT "U_Nodigit" FROM OITB WHERE "ItmsGrpCod" = (SELECT "ItmsGrpCod" FROM OITM WHERE "ItemCode" = :ItemCode)) INTO Type1 FROM DUMMY;
SELECT (SELECT Case When "U_AcceptWMQty"='1' Then '6' When "U_AcceptWMQty"='2' Then '13' End FROM OITM WHERE "ItemCode" = :ItemCode) INTO Type1 FROM DUMMY;
SELECT (SELECT "U_Itemgrpwise" FROM OITB WHERE "ItmsGrpCod" = (SELECT "ItmsGrpCod" FROM OITM WHERE "ItemCode" = :ItemCode)) INTO Itmgrpwise FROM DUMMY;
--SELECT :ItemCode,:ItemGrpName,:Type1,:Itmgrpwise FROM DUMMY;

IF :Itmgrpwise = '1' THEN 
SELECT (SELECT COUNT(*) FROM "@NT_BRST" WHERE "U_Itmgrpnm" = :ItemGrpName) INTO temp_var_0 FROM DUMMY;
IF :temp_var_0 > 0 THEN

CREATE LOCAL TEMPORARY TABLE #TEMP AS 
(
SELECT * FROM (
SELECT CASE WHEN T0."Type" = '6' THEN CAST(T0."BarPrefix" AS varchar(11)) || CAST(T0."BarSufixIncre6" AS varchar(11)) 
WHEN T0."Type" = '13' THEN CAST(T0."BarPrefix" AS varchar(11)) || CAST(T0."BarSufixIncre" AS varchar(11)) || CAST(T0."BarSufixIncre1" AS varchar(11)) 
END AS "BarCode", * FROM (SELECT *, RIGHT('00' || CAST(A1."BarSufixIncreOLD" AS varchar(10)), 3) AS "BarSufixIncre6", A1."BarSufixIncreOLD" AS "BarSufixIncre" 
FROM (SELECT A."ItemCode", B."BcdCode", C."ItmsGrpCod", C."ItmsGrpNam", D."U_Nodigit" AS "Type", D."U_Itprefix", D."U_Prefix", D."U_Start", D."U_Start1", D."U_End", 
CASE WHEN D."U_Nodigit" = '6' THEN D."U_Prefix" WHEN D."U_Nodigit" = '13' THEN D."U_Prefix" + D."U_Itprefix" END AS "Prefix", SUBSTRING(B."BcdCode", 1, 3) AS "BarPrefix", 
CASE WHEN D."U_Nodigit" = '6' THEN CAST(SUBSTRING(B."BcdCode", 4, 3) AS varchar(10)) WHEN D."U_Nodigit" = '13' THEN SUBSTRING(B."BcdCode", 4, 4) END + 1 
AS "BarSufixIncreOLD", SUBSTRING(B."BcdCode", 8, 5) || 1 AS "BarSufixIncre1",D."U_Itmgrpnm"  FROM OITM A LEFT OUTER JOIN OBCD B ON A."ItemCode" = B."ItemCode" INNER JOIN OITB C 
ON C."ItmsGrpCod" = A."ItmsGrpCod" LEFT OUTER JOIN "@NT_BRST" D ON D."U_Itmgrpnm" = C."ItmsGrpNam" ) AS A1) AS T0 
WHERE CAST(T0."BarPrefix" AS varchar(11))= CAST(T0."Prefix" AS varchar(11))) AS T1);

SELECT (SELECT COUNT(*) FROM "#TEMP" WHERE "U_Itmgrpnm" = :ItemGrpName AND "Type"= :Type1) INTO temp_var_2 FROM DUMMY;

IF :temp_var_2 > 0 THEN 

SELECT CAST(MAX(CAST("BarCode" AS bigint)) AS varchar(20)) AS "NextBarCode", "Type" FROM "#TEMP" WHERE "U_Itmgrpnm" = :ItemGrpName AND "Type"= :Type1 GROUP BY "Type";
ELSE SELECT CASE WHEN "U_Nodigit" = '6' THEN CAST("U_Prefix" AS varchar(11)) || "U_Start" WHEN "U_Nodigit" = '13' THEN CAST("U_Prefix" AS varchar(11)) || "U_Itprefix" || "U_Start" || "U_Start1" END AS "NextBarCode", "U_Nodigit" AS "Type", * FROM "@NT_BRST" WHERE "U_Itmgrpnm" = :ItemGrpName;

END IF;

DROP TABLE "#TEMP";

END IF;

ELSE

CREATE LOCAL TEMPORARY TABLE #TEMP1 AS (
SELECT * FROM (
SELECT CASE WHEN T0."Type" = '6' THEN CAST(T0."BarPrefix" AS varchar(11)) || CAST(T0."BarSufixIncre6" AS varchar(11)) WHEN T0."Type" = '13' 
THEN CAST(T0."BarPrefix" AS varchar(11)) || CAST(T0."BarSufixIncre" AS varchar(11)) || CAST(T0."BarSufixIncre1" AS varchar(11)) END AS "BarCode", * 
FROM (
SELECT *, RIGHT('00' || CAST(A1."BarSufixIncreOLD" AS varchar(10)), 3) AS "BarSufixIncre6", A1."BarSufixIncreOLD" AS "BarSufixIncre" FROM (

Select "BarSufixIncreOLD1"+1   As "BarSufixIncreOLD",
* From 
(
Select * From 
(
Select (Select ISNUMERIC("BarSufixIncreOLD1") from Dummy) As "NUMERIC"
,* From 
(
SELECT B."UpdateDate", A."ItemCode", B."BcdCode", C."ItmsGrpCod", C."ItmsGrpNam", D."U_Nodigit" AS "Type", D."U_Itprefix", D."U_Prefix", D."U_Start"
, D."U_Start1", D."U_End"
,CASE WHEN D."U_Nodigit" = '6' THEN D."U_Prefix" WHEN D."U_Nodigit" = '13' THEN D."U_Prefix" + D."U_Itprefix" END AS "Prefix"
,SUBSTRING(B."BcdCode", 1, 3) AS "BarPrefix"
 , CASE WHEN D."U_Nodigit" = '6' THEN CAST(SUBSTRING(B."BcdCode", 4, 3) AS varchar(10)) WHEN D."U_Nodigit" = '13' THEN SUBSTRING(B."BcdCode", 4, 4) END  
 AS "BarSufixIncreOLD1"
 , SUBSTRING(B."BcdCode", 8, 5) || 1 AS "BarSufixIncre1"
 , C."U_Itemgrpwise" 
 FROM OITM A 
 LEFT OUTER JOIN OBCD B ON A."ItemCode" = B."ItemCode" 
 INNER JOIN OITB C ON C."ItmsGrpCod" = A."ItmsGrpCod" 
 LEFT OUTER JOIN "@NT_BRST" D ON D."U_Itemgrpwise" = C."U_Itemgrpwise" 
 Where B."BcdCode" Not like '%*%'
 ) Where "BarSufixIncreOLD1" Not Like '%*%'
 )
  Where "NUMERIC"='1' and "BarSufixIncreOLD1"<>''
 ) As A
 
 ) AS A1) AS T0 WHERE CAST(T0."BarPrefix" AS varchar(11))= CAST(T0."Prefix" AS varchar(11)) ) AS T1);

SELECT (SELECT COUNT(*) FROM "#TEMP1" WHERE  "Type" = :Type1 AND IFNULL("U_Itemgrpwise", '') = :Itmgrpwise) INTO temp_var_4 FROM DUMMY;

IF :temp_var_4 > 0 THEN 

SELECT CAST(MAX(CAST("BarCode" AS bigint)) AS varchar(20)) AS "NextBarCode", "Type" FROM "#TEMP1" 
WHERE  "Type" = :Type1 AND IFNULL("U_Itemgrpwise", '') = :Itmgrpwise GROUP BY "Type";

ELSE 

SELECT CAST(MAX(CAST("NextBarCode" AS bigint)) AS varchar(20)) AS "NextBarCode", "U_Nodigit" AS "Type" 
FROM (SELECT CASE WHEN "U_Nodigit" = '6' THEN CAST("U_Prefix" AS varchar(11)) || "U_Start" WHEN "U_Nodigit" = '13' THEN CAST("U_Prefix" AS varchar(11)) || "U_Itprefix" || "U_Start" || "U_Start1" END AS "NextBarCode", * FROM "@NT_BRST" WHERE "U_Nodigit" = :Type1 AND IFNULL("U_Itemgrpwise", '') = :Itmgrpwise) AS T2 GROUP BY "U_Nodigit";

END IF;


DROP TABLE "#TEMP1";

END IF;

END;