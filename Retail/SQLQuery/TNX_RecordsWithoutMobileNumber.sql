SELECT T0."CardCode", T0."Phone1", T0."Cellular" FROM OCRD T0 WHERE T0."CardType" ='C' and  IFNULL(T0."Phone1",'')<>'' AND  IFNULL(T0."Cellular",'') =''
