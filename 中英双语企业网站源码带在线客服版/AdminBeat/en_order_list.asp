<%@ LANGUAGE=VBScript CodePage=65001%>
<% response.charset="utf-8" %>
<% session.codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="page_next.asp" -->
<%
'文章文件夹获取
set rs_1=server.createobject("adodb.recordset")
sql="select FolderName from web_Models_type where [id]=40"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Article_FolderName=rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing%>

<% '搜索模块
act=request.querystring("act")
keywords=trim(request.form("keywords"))
if act="search" then
if keywords<>"" then
s_sql="select * from en_web_order where [content] like '%"&keywords&"%'   order by [time] desc "
else
s_sql="select * from en_web_order where [content] like '%"&keywords&"%'  order by [time] desc"
end if
else
s_sql="select * from en_web_order where [content] like '%"&keywords&"%'  order by [time] desc "

end if 
%>
<script language="javascript">

//全选JS
function unselectall(){
if(document.form2.chkAll.checked){
document.form2.chkAll.checked = document.form2.chkAll.checked&0;
}
}
function CheckAll(form){
for (var i=0;i<form.elements.length;i++){
var e = form.elements[i];
if (e.Name != 'chkAll'&&e.disabled==false)
e.checked = form.chkAll.checked;
}
}
</script>
<script language="JavaScript">
<!--
function ask(msg) {
	if( msg=='' ) {
		msg='警告：删除后将不可恢复，可能造成意想不到后果？';
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
	  <th width="100%" height=25 class='tableHeaderText'>订单列表</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
 <form name="form2" method="post" action="en_order_Del.asp?action=AllDel&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">
	    
		<table width="95%" border="0" align="center" cellpadding="0" cellspacing="2">
          <tr>
            <td width="2%" height="30" class="TitleHighlight">&nbsp;</td>
            <td width="4%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">编号</div></td>
            <td width="52%" height="30" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">订单详情</div></td>
            <td width="8%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">订购数量</div></td>
            <td width="7%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">处理</div></td>
            <td width="18%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">下单时间</div></td>
            <td width="9%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">订单操作</div></td>
          </tr>
<% '文章列表模块
strFileName="en_order_list.asp" 
pageno=10
set rs = server.CreateObject("adodb.recordset")
rs.Open (s_sql),cn,1,1
rscount=rs.recordcount
if not rs.eof and not rs.bof then
call showsql(pageno)
rs.move(rsno)
for p_i=1 to loopno
%>

          <tr >
            <td rowspan="2" class='forumRowHighLight'><div align="center"><input type="checkbox" name="Selectitem" id="Selectitem" value="<%=rs("id")%>"></div></td>
            <td rowspan="2" class='forumRowHighLight'><div align="center"><%=rs("id")%></div></td>
            <td height="30" class='forumRowHighLight' ><span style="color: #FF0000">订购产品：</span>
			<%
			set rst=server.createobject("adodb.recordset")
			sql="select [title],file_path from [en_article] where id="&rs("article_id")&""
			rst.open(sql),cn,1,1
			if not rst.eof and not rst.bof then
			response.write "<a href='/English/"&Article_FolderName&"/"&rst("file_path")&"' target='_blank'>"&rst("title")&"</a>"
			end if
			rst.close
			set rst=nothing
			%></td>
            <td class='forumRowHighLight' ><div align="center"><%=rs("ordercount")%>
            </div></td>
            <td class='forumRowHighLight'><div align="center"><a href="en_order_view_yes.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>"><%if rs("view_yes")=1 then%>已处理<%else%><span style="color: #FF0000">未处理</span><% end if%></a></div></td>
            <td class='forumRowHighLight' ><div align="center"><%=rs("time")%></div></td>
            <td class='forumRowHighLight' >
            <div align="center"><a href="javascript:if(ask('警告：删除后将不可恢复，可能造成意想不到后果？')) location.href='en_order_del.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>';">删除</a></div>				</td>
          </tr>
		            <tr >
            <td  colspan="5" valign="top" class='forumRow'  style="line-height:200%"><span style="color: #060">联系人：</span><%=rs("name")%>&nbsp;&nbsp;<br><span style="color: #060">联系地址：</span><%=rs("address")%>&nbsp;&nbsp;<br><span style="color: #060">联系电话：</span><%=rs("tel")%>&nbsp;&nbsp;<br><span style="color: #060">电子邮件：</span><%=rs("email")%>&nbsp;&nbsp;<br><span style="color: #060">QQ：</span><%=rs("qq")%>
            <br><span style="color: #060">备注：</span><%=rs("content")%>
            
            </td>
            
          </tr>


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
		            <td height="35"  colspan="9" >&nbsp;<input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>
                    全选/全不选&nbsp;<input type="submit" name="Submit" value="删除选中"></td>
          </tr>
		    <tr  >

              <td height="35"  colspan="7" ><div align="center">
                <%call showpage(strFileName,rscount,pageno,false,true,"")%>
           </div></td>
		    </tr>
      </table></form>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="20" class='forumRow'>&nbsp;</td>
          </tr>
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| 订单搜索</td>
          </tr>
          <tr>
            <td height="70"><form name="form1" method="post" action="?act=search">
                <div align="center">
               
                  <label>
                    <input name="keywords" type="text"  size="35" maxlength="40">
                  </label>
                  <label> &nbsp;
                    <input type="submit" name="Submit" value="搜 索">
                  </label>
                </div>
            </form></td>
          </tr>
        </table>
	    <br></td>
	</tr>
	</table>

<%
Call DbconnEnd()
 %>