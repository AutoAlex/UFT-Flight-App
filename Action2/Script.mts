REM Module: Select Case Statement for selecting Subroutines

'Selects Subroutine based on Keywords
Select Case DataTable.Value("Action", dtGlobalSheet)
	
	Case "_StartApp"
		Call subStartApp
	Case "_Login"
		Call subLogin
	Case "_InsertOrder"
		Call subInsertOrder
	Case "_UpdateOrder"
		Call subUpdateOrder
	Case "_DeleteOrder"
		Call subDeleteOrder
	Case "_Logout"
		Call subLogout
	
End Select


 @@ hightlight id_;_2030099416_;_script infofile_;_ZIP::ssf77.xml_;_
