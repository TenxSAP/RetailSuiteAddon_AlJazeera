CREATE PROCEDURE "EAN13_Check_Digit_Calculation" (IN ACode varchar(12)) AS tmpCode varchar(12)
;
tmpMulSup varchar(8000)
;
tmp varchar(8000)
;
i integer
;
j integer
;
z integer
;
SumDEven integer
;
SumDOdd integer
;
List varchar(8000)
;
tmpList varchar(8000)
;
CheckSum smallint
;

BEGIN SumDEven := 0
;
 SumDOdd := 0
;
 List := ''
;
 tmpList := ''
;
 tmp := ''
;
 tmpCode := :ACode
;
 /* 0. List builder */ j := LENGTH(:tmpCode) + 1
;
 i := 1
;
 
WHILE (:i <= LENGTH(:tmpCode)) DO List := :List || '|' || LTRIM(RTRIM(CAST (:j AS VARCHAR))) || ';' || SUBSTRING(:tmpCode,
	 :i,
	 1)
;
 j := (:j - 1)
;
 i := (:i + 1)
;/* 1. Add up the digits in even position */
 
END 
WHILE
;
 i := 1
;
 tmpList := :List
;
 
WHILE (LOCATE(:tmpList,
	 '|') > 0) DO j := LOCATE(:tmpList,
	 '|')
;
 z := LOCATE(:tmpList,
	 ';')
;
 
IF (MOD(CAST(SUBSTRING(:tmpList,
	 (:j + 1),
	 (:z - (:j + 1))) AS integer),
	 2)) = 0 
THEN SumDEven := :SumDEven + CAST(SUBSTRING(:tmpList,
	 (:z + 1),
	 1) AS integer)
;
 
END 
IF
;
 tmpList := SUBSTRING(:tmpList,
	 (:z + 2),
	 LENGTH(:tmpList))
;
 /* 2. Multiply the result of the previous step (the first step) to 3 (three) */ 
END 
WHILE
;
 SumDEven := (:SumDEven * 3)
;
 /* 3. Add up the digits in the odd positions */ i := 1
;
 tmpList := :List
;
 
WHILE (LOCATE(:tmpList,
	 '|') > 0) DO j := LOCATE(:tmpList,
	 '|')
;
 z := LOCATE(:tmpList,
	 ';')
;
 
IF (MOD(CAST(SUBSTRING(:tmpList,
	 (:j + 1),
	 (:z - (:j + 1))) AS integer),
	 2)) <> 0 
THEN SumDOdd := :SumDOdd + CAST(SUBSTRING(:tmpList,
	 (:z + 1),
	 1) AS integer)
;
 
END 
IF
;
 tmpList := SUBSTRING(:tmpList,
	 (:z + 2),
	 LENGTH(:tmpList))
;
 /* 4. Add up the results obtained in steps two 
and three */ 
END 
WHILE
;
 CheckSum := (:SumDEven + :SumDOdd)
;
 /* 5. Subtract the upper multiple of 10 
from the result obtained in step four */ 
IF ((MOD(:CheckSum,
	 10)) = 0) 
THEN CheckSum := 0
;
 
ELSE tmpMulSup := LTRIM(RTRIM(CAST (:CheckSum AS VARCHAR)))
;
 i := 1
;
 
WHILE :i <= (LENGTH(:tmpMulSup) - 1) DO tmp := :tmp || SUBSTRING(:tmpMulSup,
	 :i,
	 1)
;
 
IF (:i = LENGTH(:tmpMulSup) - 1) 
THEN tmp := LTRIM(RTRIM(CAST (CAST(:tmp AS integer) + 1 AS VARCHAR)))
;
 tmp := :tmp || '0'
;
 
END 
IF
;
 i := (:i + 1)
;
 
END 
WHILE
;
 CheckSum := CAST(:tmp AS integer) - :CheckSum
;
 
END 
IF
;
 SELECT
	 :CheckSum as "CheckDigit" 
FROM DUMMY
;
 /* 
If the result of the four step is a multiple of Ten (10),
	 like Twenty,
	 Thirty,
	 Forty 
and so on,
	 the check-digit will be equal to zero,
	 otherwise the check-digit will be the result of the fifth step */ --[Note:Stringifier] SAP HANA does not support returning values using the RETURN statement

END
;