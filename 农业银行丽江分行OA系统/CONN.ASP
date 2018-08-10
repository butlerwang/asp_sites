<% 
	dim adOpenKeyset
	dim adLockOptimistic
	dim iPage
	adOpenKeyset = 1
	adLockOptimistic = 3

	'创建并打开 Connection 对象。
	set con = Server.CreateObject("ADODB.Connection") 
	ConnStr="DBQ=" & Server.Mappath("db/mails1.mdb") & ";DRIVER={Microsoft Access Driver (*.mdb)};"
	con.Open(ConnStr)
	'创建并打开 Recordset 对象。
	set Record = Server.CreateObject("ADODB.Recordset")
	Record.ActiveConnection = con
	Record.CursorType = adOpenKeyset
	Record.LockType = adLockOptimistic
	Record.PageSize=Session("iPageSize")
if Session("send")<>"common" then
	Record.open "select * from "+Session("send")+Session("id") +" ORDER BY ilevel DESC,iDateTime DESC"
else
	Record.open "select * from message ORDER BY iDateTime DESC,ilevel DESC"
end if
%>