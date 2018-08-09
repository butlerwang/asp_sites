<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.SpaceCls.asp"-->
<%

on error resume next
Dim KS,UserName,SQL,I
Set KS=New PublicCls
response.expires=0
response.ContentType="text/xml"

UserName=KS.R(KS.S("UserName"))
Dim RS:Set RS=Server.CreateObject("adodb.recordset")
rs.open "select top 100 songname,url from ks_blogmusic where username='" & username & "' order by adddate desc",conn,1,1
If Not RS.Eof Then SQL=RS.GetRows(-1)
RS.Close:Set rs=nothing
closeconn()
set KS=Nothing
Response.CodePage=65001
Response.Addheader "Content-Type","text/html; charset=utf-8" 
%>

<?xml version="1.0" encoding="utf-8" ?>
<playlist version="1" xmlns="#">
    <title>music-box</title>
    <info>#</info>
    <trackList>
          <%
		  for i=0 to ubound(sql,2)
		  dim u:u=sql(1,i)
          response.write "<track>" & vbcrlf
          response.write "<annotation>" & sql(0,i) & "</annotation>" & vbcrlf
          response.write "<location>" & sql(1,i) & "</location>"& vbcrlf
          response.write "<info>#</info>" & vbcrlf
          response.write "</track>" & vbcrlf
		  next
        %>
        
    </trackList>
</playlist>
