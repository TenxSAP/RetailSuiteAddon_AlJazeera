SELECT T0."empID", T0."lastName",T0."U_MedicalExpiry" FROM OHEM T0 Where ADD_DAYS( "U_MedicalExpiry",-20)=CURRENT_DATE
