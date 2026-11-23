SELECT T0."ItemCode", T0."BcdCode" FROM OBCD T0 where LEFT(T0."BcdCode",3)='901' and RIGHT(T0."BcdCode",6) > '028997' and RIGHT(T0."BcdCode",6)<='029999' --GROUP BY T0."BcdCode"
