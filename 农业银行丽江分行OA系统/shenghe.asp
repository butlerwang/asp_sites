<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="data.asp"-->
<!--#include file="check.asp"-->

<%response.buffer=false
dim sql
dim rs
dim id


id=request("id")
set rs=server.createobject("adodb.recordset")
sql="select * from user where id="&id
rs.open sql,conn,1,3 
rs("���")=true

		rs("iPageSize")=10
		rs("iAdd")=""
		
rs.update
'Dim objNewMail As CDONTS.NewMail
'Set objNewMail = CreateObject("CDONTS.NewMail")
'Set objNewMail = CreateObject("CDONTS.NewMail")
'objNewMail.Send("webmaster@sxhighway.net", rs("����"), "�����ʺ��Ѿ�ͨ�����", _
'"�����ʺ��Ѿ�ͨ����ˣ������ڿ���������½", 0) '' low importance
'Set objNewMail = Nothing '' canNOT reuse it for another message







'-----------------------------����Ϊ����������룬����ɾ��--------------------------------------------------------------------------------------------------------------------------------------------
		

		set con2 = Server.CreateObject("ADODB.Connection") 

		ConnStr="DBQ=" & Server.Mappath("db/mails1.mdb") & ";DRIVER={Microsoft Access Driver (*.mdb)};"
		con2.open(ConnStr)
		sql="create table recived"
		sql=sql+rs("�û���")+"(iDateTime varchar(50),iaddfile varchar(150), ifrom varchar(50),iinfo memo,ilevel char(1),cent varchar(50),iread char(1))"

		con2.Execute(sql)

		sql="create table sendout"
		sql=sql+rs("�û���")+"(iDateTime varchar(50),iaddfile varchar(150), ito varchar(50),iinfo memo,ilevel char(1),cent varchar(50),iread char(1))"

		con2.Execute(sql)

		sql="create table del"
		sql=sql+rs("�û���")+"(iDateTime varchar(50),iaddfile varchar(150), ifrom varchar(50),iinfo memo,ilevel char(1),cent varchar(50),iread char(1))"

		con2.Execute(sql)

		con2.close
		set con2=nothing






'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


rs.close
set rs=nothing


response.redirect "userchk.asp"
%>
