<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('���糬ʱ�������㻹û�е�½! ');this.location.href='../login.asp';</script>"
end if
 	temp_check_rights = Split(session("manconfig"),",")
	for temp_rights_count=0 to ubound(temp_check_rights)
	    if trim(temp_check_rights(temp_rights_count)) = "5" then
			rights_check_passkey = trim(temp_check_rights(temp_rights_count))
		end if
	next
	if rights_check_passkey <> "5" then
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('����Ȩ�޲��㣬�벻Ҫ�Ƿ�������������ģ�飬���������˺Ž���ϵͳ�Զ�ɾ��!');this.location.href='../login.asp';</Script>"
	Response.end
	end if%>
<%
Table_name=request.querystring("Table_name")
ItemID=request.querystring("ItemID")
intID=request.querystring("id")
Call OpenData()
Call Del(Table_name,ItemID,intID)
Call CloseDataBase()
Call page_back("ɾ���ɹ�")

%>