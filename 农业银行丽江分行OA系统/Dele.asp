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
email=rs("����")


'--------------------------------------------------ɾ������--------------------------------------------------------
			
			set con2 = Server.CreateObject("ADODB.Connection") 
			'ɾ�����û�����������
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
'--------------------------------------------------ɾ������--------------------------------------------------------

rs.Delete
'   Dim objNewMail As CDONTS.NewMail
'   Set objNewMail = CreateObject("CDONTS.NewMail")
'Set objNewMail = CreateObject("CDONTS.NewMail")
'objNewMail.Send("webmaster@sxhighway.net", email, "�����ʺ�δͨ����˻�ɾ��", _
'"�����ʺ�δͨ����˻�ɾ�������������ѯ�ʹ���Ա", 0) '' low importance
'Set objNewMail = Nothing '' canNOT reuse it for another message



   conn.close
   set conn=nothing
   rs.close
   set rs=nothing  
response.redirect "userchk.asp"

%>
