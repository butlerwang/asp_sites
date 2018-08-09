<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="page_next.asp" -->




<script language="JavaScript">
<!--
function ask(msg) {
	if( msg=='' ) {
		msg='警告：删除后将不可恢复，是否确定删除？';
	}
	if (confirm(msg)) {
		return true;
	} else {
		return false;
	}
}
//-->
</script>
	<%
Call header()
%>

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>在线客服列表</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	
	
		 <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='TipTitle'>&nbsp;√ 操作提示</td>
          </tr>
          <tr>
            <td height="30" valign="top" class="TipWords">
		    <p>1、点击修改，把在线客服代码替换为自己的，即可。</p>
			<p>2、在线代码获取方法：</p>
			<p>（1）QQ在线代码生成网址：http://wp.qq.com/点击“商家沟通组件”</p>
			<p>（2）旺旺在线代码生成网址：http://www.taobao.com/wangwang/2011_seller/wangbiantianxia/</p>
			<p>（3）其他在线代码，如：MSN、Skype等，请到官方生成代码。</p>
			<p>3、<font color="#009900"><b>增加、修改、删除在线客服后，必须重新生成全站静态！方可生效！</b></font></p>
			<p>4、如果不想要在线客服，则删除所有，然后重新生成所有静态，就不会显示在线客服悬浮框了。</p>
			</td>
          </tr>
          <tr>
            <td height="10" ></td>
          </tr>
        </table><br />
		
	
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| <a href="ads_position_add.asp">添加在线客服</a></td>
          </tr>
      </table><br />
	   
	  
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="2">
          <tr>
            <td width="5%" height="30" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">编号</div></td>
            <td width="25%" height="30" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">在线客服名称</div></td>
            <td width="50%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">在线客服代码显示效果</div></td>
            <td width="20%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">操作</div></td>
          </tr>
<% '在线客服列表模块
strFileName="ads_position_list.asp" 
pageno=25
set rs = server.CreateObject("adodb.recordset")
s_sql="select * from web_ads_position order by id"
rs.Open (s_sql),cn,1,1
rscount=rs.recordcount
if not rs.eof and not rs.bof then
call showsql(pageno)
rs.move(rsno)
for p_i=1 to loopno
%>
<% if p_i mod 2 =0 then
class_style="forumRow"
else
class_style="forumRowHighLight"
end if%>
            <form name="form1" method="post" action="?action=edit&id=<%=rs("id")%>">
          <tr >
            <td   height="40" class='<%=class_style%>'><div align="center"><%=rs("id")%></div></td>
           <td class='<%=class_style%>' ><div align="center"><%=rs("name")%></div></td>
            <td class='<%=class_style%>' ><div align="center"><%=rs("memo")%></div></td>
            <td class='<%=class_style%>' >
            <div align="center"><a href="ads_position_edit.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">修改</a> | <a href="javascript:if(ask('警告：删除后将不可恢复，是否确定删除？')) location.href='ads_position_del.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>';">删除</a>            </div></td>
          </tr></form>
		  		  <%
		  rs.movenext
		  next
		  else
response.write "<div align='center'><span style='color: #FF0000'>暂无数据！</span></div>"
		  end if 
		  rs.close
		  set rs=nothing
		  %>
		    <tr  >
              <td height="35"  colspan="9" ><div align="center">
                <%call showpage(strFileName,rscount,pageno,false,true,"")%>
           </div></td>
		    </tr>
      </table>
	    <br></td>
	</tr>
	</table>

<%
Call DbconnEnd()
 %>