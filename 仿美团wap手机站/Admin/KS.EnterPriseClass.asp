<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="include/admin_Style.CSS" rel="stylesheet" type="text/css">
<title>行业类别管理</title>
<%
Dim KS:Set KS=New PublicCls
select case KS.G("Action")
  case "add","edit" Call AddClass()
  case "addsave" Call AddSave()
  case "del" Call DelClass()
  case "sub" Call SubMain()
  case else call main()
End Select


sub main()
dim rssort,sqlsort
set rssort=server.createobject("adodb.recordset")
sqlsort="select * from KS_EnterPriseClass where parentid=0 Order By orderid Asc,id desc"
rssort.open sqlsort,conn,1,1
%>
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
<script language="javascript">
function del(strurl,msg){
if (confirm(msg)){
window.location=strurl
}
}
</script>
</head>

<body>
<ul id='mt'><table width="100%" border="0" cellspacing="0" cellpadding="0">  <tr>    
<td height="23" align="left" valign="top">	
<td align="center"><strong>行业大类管理（点击相应的分类进行操作）-- <A href="?action=add">添加类别</A></strong></td>   
 </td>  
 </tr>
 </table>
</ul>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
	  <%dim i
	  i=1
	  do while not rssort.eof
	  %>
        <td height="25"><img src="../images/default/arrow_r.gif" width="6" height="10">&nbsp;<a href="?action=sub&id=<%=rssort("id")%>" title="进入小类管理"><%=rssort("ClassName")%></a><span class="style1">〖<a href="?action=add&parentid=<%=rssort("id")%>"><span class="style1">添加小类</a> | <a href="?action=edit&id=<%=rssort("id")%>"><span class="style1">修改</span></a>│<a href="?action=del&id=<%=rssort("id")%>" onClick="return(confirm('你确定要删除该类吗？'))"><span class="style1">删除</span></a>〗</span></td>
		<%
		if i mod 3<>0 then
		end if
		if i mod 3=0 then%>
      </tr>
	  <tr>
	  <%end if
	  rssort.movenext
	  i=i+1
	  loop
	  %>
	  </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
<%end sub


sub submain()
dim rssort,sqlsort
set rssort=server.createobject("adodb.recordset")
sqlsort="select * from KS_EnterPriseClass where parentid=" & ks.chkclng(ks.g("id")) & " Order By orderid Asc,id desc"
rssort.open sqlsort,conn,1,1
%>
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
<script language="javascript">
function del(strurl,msg){
if (confirm(msg)){
window.location=strurl
}
}
</script>
</head>

<body>
<ul id='mt'><table width="100%" border="0" cellspacing="0" cellpadding="0">  <tr>    
<td height="23" align="left" valign="top">	
<td align="center"><strong>行业大类管理（点击相应的分类进行操作）-- <A href="?action=add&parentid=<%=ks.g("id")%>">添加类别</A></strong></td>   
 </td>  
 </tr>
 </table>
</ul>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
	  <%dim i
	  i=1
	  do while not rssort.eof
	  %>
        <td height="25"><img src="../images/default/arrow_r.gif" width="6" height="10">&nbsp;<%=rssort("ClassName")%><span class="style1">〖<a href="?action=edit&id=<%=rssort("id")%>"><span class="style1">修改</span></a>│<a href="?pid=<%=ks.g("id")%>&action=del&id=<%=rssort("id")%>&flag=sub" onClick="return(confirm('你确定要删除该类吗？'))"><span class="style1">删除</span></a>〗</span></td>
		<%
		if i mod 3<>0 then
		end if
		if i mod 3=0 then%>
      </tr>
	  <tr>
	  <%end if
	  rssort.movenext
	  i=i+1
	  loop
	  %>
	  </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
<%end sub


sub AddClass()
dim orderid,rs,sql,parentid
set rs=server.CreateObject("adodb.recordset")
if ks.chkclng(ks.g("parentid"))=0 then
sql="select orderid from ks_enterpriseclass order by orderid"
else
sql="select orderid from ks_enterpriseclass where parentid=" & ks.chkclng(ks.g("parentid"))  & " order by orderid"
end if
rs.open sql,conn,1,1
if not rs.eof then
rs.movelast
orderid=rs(0)
end if
orderid=cint(orderid)+1
rs.close

rs.open "select * from ks_enterpriseclass where id=" & ks.chkclng(ks.g("id")),conn,1,1
if not rs.eof then
classname=rs("classname")
orderid=rs("orderid")
parentid=rs("parentid")
else
parentid=ks.chkclng(ks.g("parentid"))
end if
rs.close
%>
<body>
<ul id='mt'><table width="100%" border="0" cellspacing="0" cellpadding="0">  <tr>    
<td height="23" align="left" valign="top">	
<td align="center"><strong>添加大类</strong></td>   
 </td>  
 </tr>
 </table>
</ul>
<form name="form1" method="post" action="?action=addsave">
 <input type="hidden" value="<%=ks.g("id")%>" name="id">
      <table width="56%"  border="0" align="center" cellpadding="4" cellspacing="0">
        <tr>
          <td width="16%" height="23"><div align="right"><strong>所属大类:</strong></div></td>
          <td width="84%">
		  <select name="parentid">
		   <option value="0">-作为大类-</option>
		   <%
		    rs.open "select * from ks_enterpriseclass where parentid=0 order by orderid",conn,1,1
			do while not rs.eof
			 if trim(rs("id"))=trim(parentid) then
			 response.write "<option value='" & RS("id") & "' selected>" & rs("classname") & "</option>"
			 else
			 response.write "<option value='" & RS("id") & "'>" & rs("classname") & "</option>"
			 end if
			rs.movenext
			loop
			rs.close
		   %>
		  </select>
		  
		  </td>
        </tr>
        <tr>
          <td width="16%" height="23"><div align="right"><strong>类别名称:</strong></div></td>
          <td width="84%">
		  <%If request("action")="edit" then%>
		  	<input type="text" name="classname" value="<%=classname%>">
		  <%else%>
		  <textarea name="classname" style="width:300px;height:200px"><%=classname%></textarea>
		  <br/><font color=blue>批量添加时,一行表示一个分类</font>
		  <%end if%>
		  </td>
        </tr>
        <tr>
          <td height="23"><div align="right"><strong>排序:</strong></div></td>
          <td><input name="orderid" type="text" size="12" value="<%=orderid%>"></td>
        </tr>
        <tr>
          <td height="23">&nbsp;</td>
          <td>
		  <input type="submit" name="Submit" class="button" value=" 确 定 ">
            <input type="button" name="Submit"  class="button" value=" 返 回 " onClick="history.back(-1)"></td>
        </tr>
      </table>
    </form>
<%
set rs=nothing
end sub


sub AddSave()
	dim rs,sql,classarr,i
	set rs=server.createobject("adodb.recordset")
	sql="select * from KS_EnterPriseclass where id=" & ks.chkclng(ks.s("id"))
	rs.open sql,conn,1,3
	if rs.eof then
	  classarr=Split(replace(KS.G("ClassName")," ",""),vbcrlf)
	  for i=0 to ubound(classarr)
	    if classarr(i)<>"" then
	     rs.addnew
	    rs("classname")=classarr(i)
		rs("orderid")=ks.chkclng(ks.g("orderID"))+i
		rs("ParentID")=KS.ChkClng(ks.g("parentid"))
		rs.update
		end if
	  next
	else
	rs("ClassName")=ks.g("classname")
	rs("OrderID")=ks.chkclng(ks.g("orderID"))
	rs("ParentID")=KS.ChkClng(ks.g("parentid"))
	rs.update
	end if
	rs.close:set rs=nothing
	if ks.chkclng(ks.g("id"))=0 then
	response.write "<script>if (confirm('添加成功,继续添加吗？')){location.href='KS.EnterPriseClass.asp?parentid=" & ks.g("parentid") & "&action=add';}else{location.href='?';}</script>"
	else
		if ks.chkclng(ks.g("parentid"))=0 then
		response.write "<script>alert('修改成功!');location.href='KS.Enterpriseclass.asp';</script>"
		else
		response.write "<script>alert('修改成功!');location.href='KS.Enterpriseclass.asp?id=" & ks.g("parentid") & "&action=sub';</script>"
		end if
	end if
end Sub

sub DelClass()
dim id
id=cint(request.querystring("id"))
conn.execute "delete from ks_enterpriseclass where id="&id
conn.execute "delete from ks_enterpriseclass where parentid="&id
 KS.AlertHintScript "恭喜,删除成功!"
end sub
%>
