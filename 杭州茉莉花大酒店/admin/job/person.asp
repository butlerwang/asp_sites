<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%openData()
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('网络超时，或者你还没有登陆! ');this.location.href='../login.asp';</script>"
end if
  IF instr(webConfig,", 6")=0 or instr(session("manconfig"),", 6")=0 Then'网站功能配置
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('您的权限不足，请不要非法调用其它管理模块，否则您的账号将被系统自动删除!');this.location.href='../login.asp';</Script>"
Response.end
end if%>
<%Dim Act
  Act=Request("act")
  openData()
  Select case Act
    Case "del" : Call Del()
	Case "Ability" : Call Ability()
	Case Else : Call Main()	
  End Select
  Call CloseDataBase()
  
  Sub Del()
    ID=Request.Form("ID")
	If ID="" Then Call WriteErr("请选择要删除的信息！",1)
	sql="Delete From Sbe_Resume Where ID in("&ID&")"
	Conn.execute sql
	Response.Redirect(request.ServerVariables("HTTP_REFERER"))  
  End Sub
  
  
  Sub Ability()
     ID=Request.Form("ID")
	 If ID="" Then Call WriteErr("请选择要转移的信息！",1)
     Sql="Update Sbe_Resume set Ability=2 Where ID in("&ID&")"	 
	 Conn.execute sql
	 Response.Redirect(request.ServerVariables("HTTP_REFERER"))
	 Response.End()
  End Sub
  
  Sub Main()
	Keyword=replace(trim(Request.QueryString("Keyword")),"'","")
	Tid=replace(trim(Request.QueryString("Tid")),"'","") 
  	  Set rs = Server.CreateObject("ADODB.Recordset")
      Sql="select ID,Job,Sex,Education,AddDate,Ability from sbe_job "
Sql = "select a.ID,a.Job,a.Sex,a.Education,a.AddDate,a.Ability,a.RealName from Sbe_Resume as a left outer join sbe_job b on a.Job=b.ID  where a.Ability=1 and a.Job in(b.ID) "
if tid="Job" then Sql=Sql & " and b."&tid&" like '%"&Keyword&"%' "
 if tid="RealName" then Sql=Sql & " and a."&tid&" like '%"&Keyword&"%' "
 Sql=Sql&"order by a.AddDate desc"
'response.Write Sql
 'response.End
	   Rs.open Sql,conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理系统</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
   function SelectAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name == 'id')
       e.checked = form.ChkAll.checked;
    }
	}
	
	function check(){
	if(confirm("确定执行操作吗？")){	
	var chked;
	chked=false;
    for(var i=0;i<form1.elements.length;i++)
    {
       var e = form1.elements[i];
       if (e.name=='id'&&e.checked==true)
        { chked=true;
	       break;}
    }
	if(chked==false){
	alert("请选择要操作的信息！");
	return false;	
	}
	if(form1.act[0].checked==false&&form1.act[1].checked==false){
	alert("请选择要执行的操作！");
	return false;	
	}	
	return true;
	}
	else
	{return false;}
	
	}
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="21%" height="25"><font color="#6A859D">在线招聘 &gt;&gt; 应聘信息列表</font></td>
    <form name="formsearch" method="get" action="person.asp"> 
      <td width="79%"> <strong>【<font color="#FF0000"><img src="../images/i_search.gif" width="14" height="14">信息搜索</font>】 
        </strong> 
        <input type="text" name="keyword" value="<%=Keyword%>">
        <select name="tid">
          <option value="RealName" <%if tid="RealName" or tid="" then response.Write("selected") end if%>>姓名</option>
          <option value="Job" <%if tid="Job" then response.Write("selected") end if%>>求职岗位</option>
        </select><input type="submit" name="Submit" value="搜索" class="sbe_button">&nbsp;
        <input type="button" name="ref" value="不能自动刷新点击" onClick="location.href='person.asp'"  class="sbe_button" title="默认为查找所有信息">
      </td>
    </form>
  </tr>
  <tr> 
    <td height="1" colspan="2" background="../images/dot.gif"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="5"></td>
  </tr>
</table>
<table width="83%" border="0" align="center" cellpadding="0" cellspacing="0" id="loading">
	<tr> 
      
    <td height="63" colspan="8"><strong>正在载入数据，请稍候...</strong></td>
    </tr>
</table>
<%'response.Flush()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <tr class="sbe_table_title"> 
    <td class="sbe_table_title">&nbsp;</td>
    <td width="12%" height="25" class="sbe_table_title">姓名</td>

   <!-- <td height="25" class="sbe_table_title">部门</td>-->
    <td height="25" class="sbe_table_title">求职岗位</td>
    <td class="sbe_table_title" style="display:none">性别</td>
    <td class="sbe_table_title">学历</td>
    <td width="22%" height="25" class="sbe_table_title">提交时间</td>
    <td width="10%" height="25" class="sbe_table_title">状态</td>
    <td height="25" class="sbe_table_title">查看</td>
  </tr>
  <form name="form1" method="post" action="" onSubmit="return check()">
    <%
	  '   page=request("page") '获取当前页数
'		 if page="" or not IsNumeric(page) then page=1
'		 sp_where="Ability=0 "
'		 if Keyword<>"" then sp_where=sp_where & " and "&tid&" like '%"&Keyword&"%' "		
'		 '===================================
'		 '=========== 设置存储过程参数 =======
'		 '===================================
'		 Dim sp_table,sp_collist,sp_condition,sp_col,sp_orderby,sp_pagesize,sp_page,sp_records,Cmd
'		 '===================================
'		 sp_table     = "Sbe_Resume" '表名   : "News" －－ 字符串
'		 sp_collist   = "ID,RealName,Job,Sex,Education,AddDate,bumen,flag"           '要查询出的字段列表,*表示全部字段  ---字符串
'		 sp_condition = sp_where            'Where 语句 不用带Where : "show=1"  
'		 sp_col       = "id"          'order by 字段   : "id"   --字符串，必填
'		 sp_orderby   = 1             '排序,0-顺序 ,1-倒序 
'		 sp_pagesize  = 15            '每页记录数
'		 sp_page      = Cint(page)    '当前页数
'		 '===============End==================
'         Set Cmd=Server.CreateObject("adodb.Command")
'         Cmd.ActiveConnection=conn 
'         Cmd.CommandText="sp_page" 
'         Cmd.CommandType=4   
'         Cmd.Parameters.Append Cmd.CreateParameter("@tb",200,1,50,sp_table) 
'         Cmd.Parameters.Append Cmd.CreateParameter("@col",200,1,50,sp_col)
'         Cmd.Parameters.Append Cmd.CreateParameter("@coltype",3,1,4,0)
'         Cmd.Parameters.Append Cmd.CreateParameter("@orderby",3,1,4,sp_orderby)
'         Cmd.Parameters.Append Cmd.CreateParameter("@collist",200,1,800,sp_collist)
'         Cmd.Parameters.Append Cmd.CreateParameter("@pagesize",3,1,4,sp_pagesize)
'         Cmd.Parameters.Append Cmd.CreateParameter("@page",3,1,4,sp_page)
'         Cmd.Parameters.Append Cmd.CreateParameter("@condition",200,1,50,sp_condition)        
'		 Cmd.Parameters.Append Cmd.CreateParameter("@records",3,2)
'         set rs=Cmd.Execute 
'         Cmd.Execute
'		 sp_records=Cmd.Parameters("@records").value	
'		  if sp_records =0 then							  
		 %>
<%if rs.eof or rs.bof then%>
    <tr> 
      <td height="25" colspan="9">暂且没有找到信息...</td>
    </tr>
    <%	  else
	  rs.pagesize=15
      totalrecord=rs.recordcount
      totalpage=rs.pagecount
	  pagenum=rs.pagesize
      rs.movefirst
      nowpage=request("page")
      if nowpage="" then
         nowpage=1
      end if
      nowpage=cint(nowpage)  
      rs.absolutepage=nowpage
	  j=1
	  Do while not Rs.EOF and j<=pagenum
   %>
    <tr onMouseOver="this.style.backgroundColor='#E9EFF3'" onMouseOut="this.style.backgroundColor=''"> 
      <td width="6%" align="center"><input name="id" type="checkbox" id="id" value="<%=rs(0)%>"></td>
      <td height="25" align="center"><%=rs(6)%></td>
      <!--<td width="23%" height="21" align="center"><%'=rs("bumen")%></td>-->
      <td width="23%" height="21" align="center"><%
	  sql1="select Job from Sbe_Job  where ID="&rs(1)
	  set rs1=conn.execute(sql1)
	  if not rs.eof then
	    response.Write rs1(0)
	   end if
	  rs1.close
	  set rs1=nothing%></td>
      <td width="7%" align="center" style="display:none"><%=rs(2)%></td>
      <td width="11%" align="center"><%=rs(3)%></td>
      <td align="center"><%=rs(4)%></td>
      <td width="10%" align="center">人才库</td>
      <td width="9%" align="center"><a href="view.asp?id=<%=rs(0)%>"><img src="../images/4_1_1.gif" border="0"></a></td>
    </tr>
    <%j=j+1
	Rs.movenext
      loop
	%>
    <tr> 
      <td height="25" colspan="9"><input type="checkbox" name="ChkAll" onClick="SelectAll(this.form)">
        全选&nbsp;&nbsp; &nbsp; &nbsp;&nbsp; 
        <input type="radio" name="act" value="Ability">
        转出人才库 
        <input type="radio" name="act" value="del">
        删除 &nbsp; <input type="submit" name="Submit2" value="执行操作" class="sbe_button">
      </td>
    </tr>
  </form>
  <tr> 
    <td align="center" valign="middle" colspan="8">&nbsp;共<%=totalrecord%>条信息  分<%=totalpage%>页面 当前第 <%=nowpage%> 页 <%if nowpage>1 then%><a href="?Pid=<%=Pid%>&Gid=<%=Gid%>&page=<%=nowpage-1%>">上一页</a><%else%>上一页<%end if%>
   <%if nowpage<totalpage then%>
     <a href="?Pid=<%=Pid%>&Gid=<%=Gid%>&page=<%=nowpage+1%>">下一页</a> 
                    <%else%>
                    下一页 
                    <%end if%></td>
  </tr>
  <%end if
	Rs.close
	set Rs=nothing
	'Set Cmd=Nothing
  %>
</table>
<p>&nbsp;</p></body>
</html>
<script language="JavaScript">
loading.style.display="none";
</script>
<% End Sub%>