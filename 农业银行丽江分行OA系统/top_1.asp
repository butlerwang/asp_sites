<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<!--#include file="check.asp"-->
<!--#include file="data.asp"-->
<meta HTTP-EQUIV=REFRESH CONTENT=120;URL=top_1.asp>
<META content=revealTrans(Transition=23,Duration=1.0) http-equiv=Page-Exit><BODY leftmargin=0 topmargin=0>
<%
	set rs=server.createobject("ADODB.recordset") 
    rs.Open "SELECT * FROM adrot",conn,1,1 
    if rs.eof then
	 response.write "<img src='images/t_1.gif'>"
    else
	total=rs.recordcount
	Randomize 
    D1 = Fix(Rnd * total)
	for i=1 to D1
	  rs.movenext
	next	 
	if rs("type")="GIF" then
	%><%if rs("url")<>"" then%><A HREF="<%=rs("url")%>" target=_blank><img alt="<%=rs("alt")%>" src="<%=rs("src")%>" width="<%=rs("width")%>" height="<%=rs("height")%>" border=0 align=middle></A><%else%><img alt="<%=rs("alt")%>" src="<%=rs("src")%>" width="<%=rs("width")%>" height="<%=rs("height")%>" border=0 align=middle><%end if%><%else%><embed src="<%=rs("src")%>" type="application/x-shockwave-flash" width="<%=rs("width")%>" height="<%=rs("height")%>"><%end if%><%end if%>
</BODY>
</HTML>
