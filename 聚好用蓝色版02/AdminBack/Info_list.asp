<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="page_next.asp" -->

<%
'招聘职位文件夹获取
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=26"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Model_FileName=rs_1("FileName")
if rs_1("FolderName")<>"" then
Model_FolderName="/"&rs_1("FolderName")
end if
end if
rs_1.close
set rs_1=nothing%>

<% '搜索模块
act=request.querystring("act")
keywords=trim(request.form("keywords"))
if act="search" then
cid=request("cid")
pid=request("pid")
ppid=request("ppid")

if cid="" and pid="" and  ppid="" then
s_sql="select id,title,cid,pid,ppid,file_path,image,index_push,view_yes,hit,ip,time,AuthorID from web_info where [title] like '%"&keywords&"%'  order by time desc"
elseif pid="" and ppid="" then
search_sql="and cid='"&cid&"'"
s_sql="select id,title,cid,pid,ppid,file_path,image,index_push,view_yes,hit,ip,time,AuthorID from web_info where [title] like '%"&keywords&"%'"&search_sql&"   order by time desc"
elseif ppid="" then
search_sql="and pid='"&pid&"'"
s_sql="select id,title,cid,pid,ppid,file_path,image,index_push,view_yes,hit,ip,time,AuthorID from web_info where [title] like '%"&keywords&"%'"&search_sql&"   order by time desc"
else
search_sql="and ppid='"&ppid&"'"
s_sql="select id,title,cid,pid,ppid,file_path,image,index_push,view_yes,hit,ip,time,AuthorID from web_info where [title] like '%"&keywords&"%'"&search_sql&"   order by time desc"
end if
else
s_sql="select id,title,cid,pid,ppid,file_path,image,index_push,view_yes,hit,ip,time,AuthorID from web_info   order by time desc"

end if

%>
<script language="JavaScript">
<!--
function ask(msg) {
	if( msg=='' ) {
		msg='警告：删除后将不可恢复，确定要删除吗？';
	}
	if (confirm(msg)) {
		return true;
	} else {
		return false;
	}
}
//-->
</script>
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
</script>	<%
Call header()
%>
<%
if cid>1 then
juhaoyong_cid=cid
juhaoyong_pid=-1
juhaoyong_ppid=-1
end if

if pid>1 then
juhaoyong_cid=juhaoyongGetCategoryParentId(pid)
juhaoyong_pid=pid
juhaoyong_ppid=-1
end if

if ppid>1 then
juhaoyong_pid=juhaoyongGetCategoryParentId(ppid)
juhaoyong_cid=juhaoyongGetCategoryParentId(juhaoyong_pid)
juhaoyong_ppid=ppid
end if

Function juhaoyongGetCategoryParentId(id)
set juhaoyongRs=server.createobject("adodb.recordset")
juhaoyongSql="select id,pid,ppid,name from category where id="&id 
juhaoyongRs.open juhaoyongSql,cn,1,1

juhaoyongGetCategoryParentId=juhaoyongRs("pid")

juhaoyongRs.close
set juhaoyongRs=nothing
End Function
%>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>职位列表</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='TipTitle'>&nbsp;√ 操作提示</td>
          </tr>
          <tr>
            <td height="30" valign="top" class="TipWords"><p>1、职位列表显示您所添加的所有职位，标示“未审核”的职位将不会在网站中显示。</p>
                <p>2、删除职位将会同步删除数据库中的记录和职位的具体地址请慎重。</p>
            </td>
          </tr>
          <tr>
            <td height="10" ></td>
          </tr>
        </table>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| <a href="info_add.asp?juhaoyong_cid=<%=juhaoyong_cid%>&juhaoyong_pid=<%=juhaoyong_pid%>&juhaoyong_ppid=<%=juhaoyong_ppid%>">添加新的职位</a></td>
          </tr>
          
      </table>
 <form name="form2" method="post" action="info_Del.asp?action=AllDel&juhaoyong_cid=<%=juhaoyong_cid%>&juhaoyong_pid=<%=juhaoyong_pid%>&juhaoyong_ppid=<%=juhaoyong_ppid%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">
 	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="2">
          <tr>
            <td width="2%" height="30" class="TitleHighlight">&nbsp;</td>
            <td width="4%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">编号</div></td>
            <td width="31%" height="30" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">职位标题</div></td>
			<td width="31%" height="30" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">招聘分类</div></td>
            <td width="6%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">审核</div></td>
            <td width="17%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">添加时间</div></td>
            <td width="8%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">职位操作</div></td>
          </tr>
<% '职位列表模块
strFileName="info_list.asp" 
pageno=25
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
<%

%>
          <tr >
            <td   height="30" class='<%=class_style%>'><div align="center"><input type="checkbox" name="Selectitem" id="Selectitem" value="<%=rs("id")%>"></div></td>
            <td   height="30" class='<%=class_style%>'><div align="center"><%=rs("id")%></div></td>
            <td class='<%=class_style%>' >&nbsp;<a href="info_edit.asp?id=<%=rs("id")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>" ><%=left(rs("title"),16)%></a></td>
            
			<td class='<%=class_style%>' >&nbsp;
			<% '分类显示
			cid=cint(rs("cid"))

			set rs1=server.createobject("adodb.recordset")
			sql="select name from category where id="&cid&""
			rs1.open(sql),cn,1,1
			if not rs1.eof and not rs1.bof then
			response.write rs1("name")
			response.write "&nbsp;>&nbsp;"
			end if
			rs1.close
			set rs1=nothing
			
			if rs("pid")<>"" then
            pid=cint(rs("pid"))
						set rs1=server.createobject("adodb.recordset")
			sql="select name from category where id="&pid&""
			rs1.open(sql),cn,1,1
			if not rs1.eof and not rs1.bof then
			response.write rs1("name")
			response.write "&nbsp;>&nbsp;"
			end if
			rs1.close
			set rs1=nothing
			end if
			
			if rs("ppid")<>"" then
            ppid=cint(rs("ppid"))
						set rs1=server.createobject("adodb.recordset")
			sql="select name from category where id="&ppid&""
			rs1.open(sql),cn,1,1
			if not rs1.eof and not rs1.bof then
			response.write rs1("name")
			end if
			rs1.close
			set rs1=nothing
			end if
			%>            </td>
			
			<td class='<%=class_style%>' ><div align="center"><a href="info_view_yes.asp?id=<%=rs("id")%>&juhaoyong_cid=<%=juhaoyong_cid%>&juhaoyong_pid=<%=juhaoyong_pid%>&juhaoyong_ppid=<%=juhaoyong_ppid%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>"><%if rs("view_yes")=1 then%>已审核<%else%><span style="color: #FF0000">未审核</span><% end if%></a></div></td>
            <td class='<%=class_style%>' ><div align="center"><%=rs("time")%></div></td>
            <td class='<%=class_style%>' >
            <div align="center"><a href="info_edit.asp?id=<%=rs("id")%>&juhaoyong_cid=<%=juhaoyong_cid%>&juhaoyong_pid=<%=juhaoyong_pid%>&juhaoyong_ppid=<%=juhaoyong_ppid%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">修改</a> | <a href="javascript:if(ask('警告：删除后将不可恢复，确定要删除吗？')) location.href='info_del.asp?id=<%=rs("id")%>&juhaoyong_cid=<%=juhaoyong_cid%>&juhaoyong_pid=<%=juhaoyong_pid%>&juhaoyong_ppid=<%=juhaoyong_ppid%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>';">删除</a>            </div></td>
          </tr>
		  		  <%
		  rs.movenext
		  next
		  else
response.write "<div align='center'><span style='color: #FF0000'>暂无职位！</span></div>"
		  end if 
		  rs.close
		  set rs=nothing
		  %>
		          <tr  >
		            <td height="35"  colspan="9" >&nbsp;<input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>
                    全选/全不选&nbsp;<input type="submit" name="Submit" value="删除选中"></td>
          </tr>
            <tr  >
              <td height="35"  colspan="9" ><div align="center">
                <%call showpage(strFileName,rscount,pageno,false,true,"")%>
           </div></td>
		    </tr>
      </table> 
 </form>  
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="20" class='forumRow'>&nbsp;</td>
          </tr>
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| 职位搜索</td>
          </tr>
          <tr>
            <td height="70"><form name="form1" method="post" action="?act=search">
              <div align="center">
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