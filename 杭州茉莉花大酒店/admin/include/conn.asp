<!--#include file="config.asp"-->
<%
for each element in request.QueryString
if instr(request.QueryString(element),"'")>0 or instr(request.QueryString(element),";")>0 or instr(request.QueryString(element),"and")>0 or instr(request.QueryString(element),"%")>0 or instr(request.QueryString(element),"/add")>0 or instr(request.QueryString(element),"net")>0 then
response.Write("非法参数!")
response.End()
elseif instr(request.QueryString(element),"exec")>0 or instr(request.QueryString(element),"char")>0 or instr(request.QueryString(element),"&quot;")>0 or instr(request.QueryString(element),"truncate")>0 or instr(request.QueryString(element),"update")>0 or instr(request.QueryString(element),"Asc")>0 then
response.Write("非法参数！")
response.End() 
end if 
next

Function sqlhack(parameters)
dim regstr,regex
	set regex=New RegExp
	regex.pattern="^([;])+$"
	sqlhack=regex.test(parameters)
	set regex=Nothing
End Function
Dim Conn

Public Sub OpenDataBase()
on Error  Resume Next
set conn=Server.CreateObject(ado_conn)
conn.connectionstring="DBQ=" + Server.MapPath(DataBaseConnectStr) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
conn.open connstr

If Err Then
	Response.Write "数据库连接错误！"
	Response.End
End If
End Sub

Public Sub CloseDatabase()
'过程名称: CloseDataBase
'功能描述: 关闭数据库连接
'使用方法：Call CloseDataBase()
	Conn.Close:Set Conn=Nothing
End Sub
Private Sub OpenData()
'过程名称: OpenData
'功能描述: 打开数据库
'使用方法：Call OpenData()
	If IsEmpty(Conn) Then
		Call OpenDataBase()
		Exit Sub
	End If

	If Conn Is Nothing Then
		Call OpenDataBase()
	End if
End Sub
%>