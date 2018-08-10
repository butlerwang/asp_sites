<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->

<%
del_id=Request("delid")
del_bz=Request("delbz")
if del_bz<>"My_only" and del_bz<>"My_public" then
    Response.Redirect ("main.asp")
end if
if del_bz="My_public" and Session("Urule")<>"a" and Session("Urule")<>"b" then
    Response.Redirect ("main.asp")
end if
strSql="select * from jhtdata where id="&del_id&" and Á´½Ó<>'N/A'"
set del_rs=conn.Execute (strSql)
if not del_rs.eof then
    whichfile=server.mappath(del_rs("Á´½Ó"))
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set thisfile = fs.GetFile(whichfile)
  thisfile.Delete True
end if
del_rs.close
set del_rs=nothing
strSql="DELETE FROM jhtdata where id="&del_id
conn.Execute (strSql)
conn.Close
set conn = nothing
Response.Redirect "shouqu.asp"
%>
