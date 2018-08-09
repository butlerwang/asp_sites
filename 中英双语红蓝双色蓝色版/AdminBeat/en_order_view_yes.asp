<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/en_x_to_html/article_to_html.asp" -->

	<%
Call header()
%>


	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>处理订单</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" bgcolor="#B1CFF8"><div align="center"></div></td>
          </tr>
          <tr>
            <td height="100">
			<%page=request.querystring("page")
			act=request.querystring("act")
			keywords=request.querystring("keywords")
			article_id=cint(request.querystring("id"))
			set rs=server.createobject("adodb.recordset")
sql="select id,view_yes,article_id from en_web_order where id="&article_id&""
rs.open(sql),cn,1,3
a_id=rs("article_id")
if rs("view_yes")=0 then
rs("view_yes")=1
else
rs("view_yes")=0
end if
rs.update
rs.close
set rs=nothing
response.Write "<script language='javascript'>alert('处理成功！');location.href='en_order_list.asp?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
			%></td>
          </tr>
        </table>
	    </td>
	</tr>
	</table>


<%
Call DbconnEnd()
 %>