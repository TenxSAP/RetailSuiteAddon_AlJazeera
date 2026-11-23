SELECT T0."ItemCode", T1."ItemName", T0."WhsCode", T0."OnHand", T0."MinStock" FROM OITW T0 INNER JOIN OITM T1 ON T0."ItemCode" = T1."ItemCode"  WHERE T0."OnHand" < T0."MinStock"
