<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->

<%
del_id=Request("delid")
del_bz=Request("delbz")
if del_bz<>"My_only" and del_bz<>"My_public" then
	Response.Redirect ("index.asp")
end if
if del_bz="My_public" and Session("Urule")<>"a" and Session("Urule")<>"b" then
	Response.Redirect ("index.asp")
end if
strSql="select * from jhtdata where id="&del_id&" and Á´½Ó<>'N/A'"
set del_rs=conn.Execute (strSql)
if not del_rs.eof then
	old_filename=del_rs("Á´½Ó")
	yb=instr(old_filename,"/")
	do while yb<>0
		old_filename=mid(old_filename,yb+1)
		yb=instr(old_filename,"/")
	loop
	new_fname=old_filename
	wulilj=application("updir")
	del_filename=wulilj&new_fname
	Set GoDelFile=Server.CreateObject("Scripting.FileSystemObject")
	if GoDelFile.fileExists(del_filename) then
		GoDelFile.DeleteFile del_filename
	end if
	set GoDelFile = nothing
end if
del_rs.close
set del_rs=nothing
strSql="DELETE FROM jhtdata where id="&del_id
conn.Execute (strSql)
conn.Close
set conn = nothing
Response.Redirect ("bbs.asp")
%>
