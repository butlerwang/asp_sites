<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Article_list_to_html.asp" -->
<!-- #include file="../inc/x_to_html/post_index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Blank_Content_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Recruit_list_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Case_List_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Search_index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/SiteMap_index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Article_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Product_to_html.asp" -->
<!-- #include file="../inc/x_to_html/order_index_to_html.asp" -->

	<%
Call header()
%>

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>生成所有页面</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class="TitleHighlight3"></td>
          </tr>
          <tr>
            <td height="100"><div align="center">
			<a href="html_all.asp"><font color="#FF0000"><b>点击这里生成所有页面</b></font></a><br /><br />
			<font color="#0000ff">（温馨提示：生成所有页面比较耗费时间，请耐心等待......）</font>
			</div></td>
          </tr>
        </table>
	    </td>
	</tr>
	</table>

<%
Call DbconnEnd()
 %>