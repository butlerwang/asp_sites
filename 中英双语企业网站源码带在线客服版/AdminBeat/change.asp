<%@ LANGUAGE=VBScript CodePage=65001%>
<% response.charset="utf-8" %>
<% session.codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->

	<%
Call header()
%>


	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>删除订单</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" bgcolor="#B1CFF8"><div align="center"></div></td>
          </tr>
          <tr>
            <td height="100">
			<%
			set rs=server.createobject("adodb.recordset")
sql="select * from web_models "
rs.open(sql),cn,1,3
do while not rs.eof


replace_code=replace(rs("content"),"</head>","<script type='text/javascript'> window.onerror=function(){return true;} </script> </head>")


rs("content")=replace_code

rs.movenext
loop
rs.close
set rs=nothing

%>
<%
response.Write "<script language='javascript'>alert('成功！');history.back(-1);</script>"
			%></td>
          </tr>
        </table>
	    </td>
	</tr>
	</table>


<%
Call DbconnEnd()
 %>