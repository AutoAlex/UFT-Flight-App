
 @@ hightlight id_;_2009247976_;_script infofile_;_ZIP::ssf75.xml_;_
 @@ hightlight id_;_2137339448_;_script infofile_;_ZIP::ssf83.xml_;_


Select Case DataTable.Value("Action", dtGlobalSheet)
	
	Case "_Login"
		Call subLogin
	Case "_InsertOrder"
		Call subInsertOrder
	Case "_OpenOrder"
		Call subOpenOrder
	Case "_UpdateOrder"
		Call subUpdateOrder
	Case "_DeleteOrder"
		Call subDeleteOrder
	Case "_Logout"
		Call subLogout
	Case "_OpenApp"
		Call subOpenApp
	Case "_inValid_Login"
	 	Call subInvalidLogin
	 Case "_selectFight"
	 	Call subFlightSelection
	 	Call subUpdateDataTable
	 	Call subUpdateExcelData

End Select


