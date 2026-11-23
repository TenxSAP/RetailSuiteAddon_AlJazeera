CREATE PROCEDURE "SBO_GetMIUser"(IN User Varchar(30)) As 
Begin Select
	 "User" 
From ( Select
	 'manager' As "User" 
	From DUMMY 
	Union All Select
	 'D1001' As "User" 
	From DUMMY 
	Union All Select
	 'D1002' As "User" 
	From DUMMY 
	Union All Select
	 'D1006' As "User" 
	From DUMMY 
	Union All Select
	 'D1007' As "User" 
	From DUMMY 
	Union All Select
	 'D1008' As "User" 
	From DUMMY 
	Union All Select
	 'D1011' As "User" 
	From DUMMY 
	Union All Select
	 'D1012' As "User" 
	From DUMMY 
	Union All Select
	 'D1014' As "User" 
	From DUMMY 
	Union All Select
	 'D1015' As "User" 
	From DUMMY 
	Union All Select
	 'D1016' As "User" 
	From DUMMY 
	Union All Select
	 'D1017' As "User" 
	From DUMMY 
	Union All Select
	 'D1018' As "User" 
	From DUMMY 
	Union All Select
	 'D1019' As "User" 
	From DUMMY 
	Union All Select
	 'D1020' As "User" 
	From DUMMY 
	Union All Select
	 'D1025' As "User" 
	From DUMMY 
	Union All Select
	 'D1026' As "User" 
	From DUMMY 
	Union All Select
	 'D1027' As "User" 
	From DUMMY 
	Union All Select
	 'D1032' As "User" 
	From DUMMY 
	Union All Select
	 'D1034' As "User" 
	From DUMMY 
	Union All Select
	 'hr01' As "User" 
	From DUMMY 
	Union All Select
	 'hr02' As "User" 
	From DUMMY 
	Union All Select
	 'it01' As "User" 
	From DUMMY 
	Union All Select
	 'it02' As "User" 
	From DUMMY 
	Union All Select
	 'it03' As "User" 
	From DUMMY 
	Union All Select
	 'it04' As "User" 
	From DUMMY 
	Union All Select
	 'it05' As "User" 
	From DUMMY 
	Union All Select
	 'P1006' As "User" 
	From DUMMY 
	Union All Select
	 'SALMAN' As "User" 
	From DUMMY 
	Union All Select
	 'D1035' As "User" 
	From DUMMY 
	Union All Select
	 'P1008' As "User" 
	From DUMMY 
	Union All Select
	 'FAISAL' As "User" 
	From DUMMY 
	Union All Select
	 'D1031' As "User" 
	From DUMMY ) 
Where "User" = :User 
;

End
;