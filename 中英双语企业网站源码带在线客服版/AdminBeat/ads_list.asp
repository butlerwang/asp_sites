<%@ LANGUAGE=VBScript CodePage=65001%>
<% response.charset="utf-8" %>
<% session.codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="page_next.asp" -->

<% '搜索模块
act=request.querystring("act")
keywords=trim(request.form("keywords"))
cid=request("cid")


if act="search" then


if cid=""  then
s_sql="select * from web_ads where [name]  like '%"&keywords&"%'  order by [position]"
else
search_sql="and [position]="&cid&""
s_sql="select * from web_ads where [name] like '%"&keywords&"%'"&search_sql&"  order by [order]"
end if

else
s_sql="select * from web_ads order by [position]"

end if

%>

<% '修改顺序模块
action1=request.querystring("action")
id1=request.querystring("id")
order1=trim(request.form("order"))
if action1="edit" then
if isnumeric(order1)=false then
response.Write "<script language='javascript'>alert('您输入的不是数字！');location.href='?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
else
set rs1=server.createobject("adodb.recordset")
sql="select id,order from web_ads where id="&id1&""
rs1.open(sql),cn,1,3
rs1("order")=cint(order1)
rs1.update
rs1.close
set rs1=nothing
call index_to_html()
end if
end if

%>
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
	  <th width="100%" height=25 class='tableHeaderText'>广告列表</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='TipTitle'>&nbsp;√ 操作提示</td>
          </tr>
          <tr>
            <td height="30" valign="top" class="TipWords"><p>1、几乎所有的广告形式都可以在后台进行更新。目前涵盖文字广告、图片广告、Flash广告和广告联盟代码广告。</p>
                <p>2、除广告联盟代码广告外，其它广告形式将以生成JS文件形式出现在网页中。</p>
                <p>3、除广告联盟代码广告外，其它广告的更新不需要生成相关网页即可看到修改的效果。</p>
                <p>4、广告文件均存放在根目录下的ADs文件夹中。</p></td>
          </tr>
          <tr>
            <td height="10" ></td>
          </tr>
        </table>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| <a href="ads_add.asp">添加新的广告</a></td>
          </tr>
          <tr>
            <td height="30"></td>
          </tr>
      </table>
	   
	  
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="2">
          <tr>
            <td width="4%" height="30" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">编号</div></td>
            <td width="24%" height="30" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">广告标题</div></td>
            <td width="14%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">广告类型</div></td>
            <td width="14%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">广告位置</div></td>
            <td width="10%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">显示排序</div></td>
            <td width="7%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">审核</div></td>
            <td width="18%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">添加时间</div></td>
            <td width="9%" bgcolor="#A5C6FC" class="TitleHighlight"><div align="center" style="font-weight: bold; color: #FFFFFF">操作</div></td>
          </tr>
<% '文章列表模块
strFileName="ads_list.asp" 
pageno=20
set rs = server.CreateObject("adodb.recordset")
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
           <td class='<%=class_style%>' ><%=rs("name")%></td>
            <td class='<%=class_style%>' ><div align="center">
              <%
			select case rs("ADtype")
			case 1
			response.write "文字广告"
			case 2
			response.write "图片广告"
			case 3
			response.write "Flash广告"
			case 4
			response.write "广告代码"
			end select%>
            </div></td>

            <td class='<%=class_style%>' ><div align="center"><%
			set rst=server.createobject("adodb.recordset")
			sql="select name from web_ads_position where [id]="&rs("position")&""
			rst.open(sql),cn,1,1
			if not rst.eof and not rst.bof then
			response.write rst("name")
			end if
			rst.close
			set rst=nothing
			%></div></td>
            <td class='<%=class_style%>' > <div align="center"><label>
            <input name="order" type="text" value="<%=rs("order")%>" size="5">
            <input type="submit" name="Submit" value="修改"  >
            </label>
           </div></td>
            <td class='<%=class_style%>' ><div align="center"><a href="ads_view_yes.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>"><%if rs("view_yes")=1 then%>已审核<%else%><span style="color: #FF0000">未审核</span><% end if%></a></div></td>
            <td class='<%=class_style%>' ><div align="center"><%=rs("time")%></div></td>
            <td class='<%=class_style%>' >
            <div align="center"><a href="ads_edit.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">修改</a> | <a href="javascript:if(ask('警告：删除后将不可恢复，可能造成意想不到后果？')) location.href='ads_del.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>';">删除</a>            </div></td>
          </tr></form>
		  		  <%
		  rs.movenext
		  next
		  else
response.write "<div align='center'><span style='color: #FF0000'>暂无链接！</span></div>"
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
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="20" class='forumRow'>&nbsp;</td>
          </tr>
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| 广告搜索</td>
          </tr>
          <tr>
            <td height="70"><form name="form1" method="post" action="?act=search"><div align="center">
<%
Dim count1,rsClass1,sqlClass1
set rsClass1=server.createobject("adodb.recordset")
sqlClass1="select id,name from web_ads_position  order by id" 
rsClass1.open sqlClass1,cn,1,1
%>
            <select name="cid" id="cid" onChange="changeselect1(this.value)">
              <option value="">选择分类</option>
              <%
count1 = 0
do while not rsClass1.eof
response.write"<option value="&rsClass1("ID")&">"&rsClass1("Name")&"</option>"
count1 = count1 + 1
rsClass1.movenext
loop
rsClass1.close
%>
            </select>
            <label>
<input name="keywords" type="text"  size="35" maxlength="40">
              </label>
                <label>
                       &nbsp;
                       <input type="submit" name="Submit" value="搜 索">
                </label>
              </div>
            </form>
            </td>
          </tr>
      </table>
	    <br></td>
	</tr>
	</table>

<%
Call DbconnEnd()
 %>