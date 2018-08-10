<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="data.asp"-->
<!--#include file="check.asp"-->
<!--#include file="delmail.asp"-->
<%
   dim sql 
   dim rs

set rs=server.createobject("adodb.recordset")
sql="select * from user where id="&request("id")
rs.open sql,conn,1,3 
email=rs("信箱")


'--------------------------------------------------删除信箱--------------------------------------------------------
			
			set con2 = Server.CreateObject("ADODB.Connection") 
			'删除此用户的信箱数据
			ConnStr="DBQ=" & Server.Mappath("db/mails1.mdb") & ";DRIVER={Microsoft Access Driver (*.mdb)};"
			con2.open(ConnStr)
			set Record2 = Server.CreateObject("ADODB.Recordset")
			Record2.ActiveConnection = con2
			Record2.CursorType = adOpenKeyset
			Record2.LockType = adLockOptimistic
			
			DelAll "del" 
			DelAll "recived" 
			DelAll "sendout" 	
			con2.close
			set con2=nothing
'--------------------------------------------------删除信箱--------------------------------------------------------

rs.Delete
'   Dim objNewMail As CDONTS.NewMail
'   Set objNewMail = CreateObject("CDONTS.NewMail")
'Set objNewMail = CreateObject("CDONTS.NewMail")
'objNewMail.Send("webmaster@sxhighway.net", email, "您的帐号未通过审核或被删除", _
'"您的帐号未通过审核或被删除，具体情况请询问管理员", 0) '' low importance
'Set objNewMail = Nothing '' canNOT reuse it for another message



   conn.close
   set conn=nothing
   rs.close
   set rs=nothing  
response.redirect "userchk.asp"

%>
