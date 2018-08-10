<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('网络超时，或者你还没有登陆! ');this.location.href='../login.asp';</script>"
end if
 	temp_check_rights = Split(session("manconfig"),",")
	for temp_rights_count=0 to ubound(temp_check_rights)
	    if trim(temp_check_rights(temp_rights_count)) = "4" then
			rights_check_passkey = trim(temp_check_rights(temp_rights_count))
		end if
	next
	if rights_check_passkey <> "4" then
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('您的权限不足，请不要非法调用其它管理模块，否则您的账号将被系统自动删除!');this.location.href='../login.asp';</Script>"
	Response.end
	end if%>
<%Dim Act
  Act=Request("act")
Call openData()
  Select case Act
    Case "up" : Call Up()
	Case "tuijian" : Call tuijian()	
	Case "down" : Call down()
	Case "show" : Call Show()
    Case "del" : Call Del()
	Case "move" : Call Moveto()
	Case "leibie" : Call leibie()
	Case Else : Call Main()	
  End Select
  Call CloseDataBase()
    Sub MoveTo()
	   id=Request.Form("id")
	   tid=request.Form("select")
	   if id = "" then Call WriteErr("请选择要操作的信息！",1)
	   if tid = "" then Call WriteErr("请选择要转移的分类！",1)
  sqlsize ="select * from Sbe_Down_Class where ID ="&Tid
  set rssize=conn.execute(sqlsize)
  if not (rssize.eof and rssize.bof) then
	if  rssize("Depth") = 0 then
	   bigclass=rssize("ID")
	   else
       bigclass = rssize("ParID")
	end if
  end if 
  rssize.close
 set rssize=nothing
	 Set rs=Server.CreateObject("adodb.recordset")
	   Sql="update Sbe_Down set tid="&tid&",bigclass="&bigclass&" where id in("&id&")"
	   conn.execute Sql   
	   Response.Redirect(request.ServerVariables("HTTP_REFERER"))
	End Sub
	Sub Del()
	   id=Request.Form("id")
	   if id = "" then Call WriteErr("请选择要操作的信息！",1)
	   Set Rs=Server.CreateObject("adodb.recordset")
	   Sql="Select * From Sbe_Down Where ID in ("&ID&")"
	   Rs.Open Sql,conn,1,3
	      Do while not rs.eof	   
			if rs("spic")<>"" and not isnull(rs("spic")) then Call DeleteFile(rs("spic"),"../../uploadfile")
			'if rs("bpic")<>"" and not isnull(rs("bpic")) then Call DeleteFile(rs("bpic"),"../../uploadfile")
			'if rs("uploadfile")<>"" and not isnull(rs("uploadfile")) then Call DeleteFile(rs("uploadfile"),"../../uploadfile")
		    rs.delete
		   rs.movenext
		   loop
		Rs.Close
		Set Rs=Nothing
		Response.Redirect(request.ServerVariables("HTTP_REFERER"))	
	End Sub
  
  Sub Tuijian()
     ID=Cint(Request.QueryString("ID"))
	 Set Rs=Server.CreateObject("adodb.recordset")
	 Sql="Select id,tuijian from Sbe_Down where id="&ID	
	 Rs.Open sql,Conn,1,3	  
	    If Rs("tuijian") Then
		   Rs("tuijian") = 0
		Else
		   RS("tuijian") =1
		end If
	  Rs.Update
	 Rs.Close
	 Set Rs=Nothing	 
    Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
  End Sub
    Sub leibie()
     ID=Cint(Request.QueryString("ID"))
	 Set Rs=Server.CreateObject("adodb.recordset")
	 Sql="Select id,leibie from Sbe_Down where id="&ID	
	 Rs.Open sql,Conn,1,3	  
	    If Rs(1) Then
		   Rs(1) = 0
		Else
		   RS(1) =1
		end If
	  Rs.Update
	 Rs.Close
	 Set Rs=Nothing	 
    Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
  End Sub 
    Sub Show()
     ID=Cint(Request.QueryString("ID"))
	 Set Rs=Server.CreateObject("adodb.recordset")
	 Sql="Select Show from Sbe_Down where id="&ID
	 Rs.Open sql,Conn,1,3
	    If Rs(0) Then
		   Rs(0) = 0
		Else
		   RS(0) =1
		end If
	  Rs.Update
	 Rs.Close
	 Set Rs=Nothing	 
    Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
  End Sub
  Sub Up()    
    ID=Cint(Request.QueryString("ID"))
	Keyword=replace(trim(Request.QueryString("Keyword")),"'","")
	Tid=replace(trim(Request.QueryString("Tid")),"'","")	
	leixing=replace(trim(Request.QueryString("leixing")),"'","")	
    Call UpSequence("Sbe_Down",ID,Keyword,Tid,leixing)
	response.Redirect(request.ServerVariables("HTTP_REFERER"))
  End Sub  
  Sub Down()    
    ID=Cint(Request.QueryString("ID"))
	Keyword=replace(trim(Request.QueryString("Keyword")),"'","")
	Tid=replace(trim(Request.QueryString("Tid")),"'","")
	leixing=replace(trim(Request.QueryString("leixing")),"'","")
    Call DownSequence("Sbe_Down",ID,Keyword,Tid,leixing)
	response.Redirect(request.ServerVariables("HTTP_REFERER"))
  End Sub
  Sub Main()
	Keyword=replace(trim(Request.QueryString("Keyword")),"'","")
	Tid=replace(trim(Request.QueryString("Tid")),"'","")
    leixing=trim(Request.QueryString("leixing"))
	 page=request("page") '获取当前页数
		 if page="" or not IsNumeric(page) then page=1
		 sp_where="1=1 "
		 if Keyword<>"" then sp_where=sp_where & " and pname like '%"&Keyword&"%' "	
		 if leixing<>"" then sp_where=sp_where & " and leibie="&leixing&" "	 
		 if tid<>"" then sp_where=sp_where&" and tid in ("&ChildrenID("Sbe_Down",Cint(tid))&")"
		 '===================================
		 '=========== 设置参数 =======
		 '===================================
		 Dim sp_table,sp_collist,sp_condition,sp_col,sp_orderby,sp_pagesize,sp_page,sp_records,Cmd
		 '===================================
		 sp_table     = "Sbe_Down" '表名   : "Product" －－ 字符串
		 sp_collist   = "ID,Pname,Tid,Tuijian,Show,Succeed,leibie"           '要查询出的字段列表,*表示全部字段  ---字符串
		 sp_condition = sp_where            'Where 语句 不用带Where : "show=1"  
		 sp_col       = "sequence"          'order by 字段   : "id"   --字符串，必填
		 sp_orderby   = "desc"             '排序,0-顺序 ,1-倒序 
		 sp_pagesize  = 15            '每页记录数
		 sp_page      = Cint(page)    '当前页数
		 
		 set rs=server.CreateObject("adodb.recordset")
		 sql="select "&sp_collist&" from "&sp_table&" where "&sp_where&" order by "&sp_col&" "&sp_orderby&" "
		 'response.Write(sql)
		 rs.open sql,conn,1,3
		 
		 sp_records=rs.recordcount
		 rs.pagesize=sp_pagesize
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
	if(form1.act[1].checked==true&&form1.tid.value==""){
	alert("请选择要转移的分类！");
	form1.tid.focus();
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
    <td width="24%" height="25"><font color="#6A859D">店铺形象管理中心 &gt;&gt; 信息列表</font></td>
    <form name="formsearch" method="get" action="list.asp"> 
      <td width="76%"> <strong>【<font color="#FF0000">信息搜索</font>】 </strong><%=Proname%> 
        <input type="text" name="keyword" value="<%=keyword%>">
              分类 
         <select name="tid">
            <option value="">所有分类...</option>
		    <%Call ShowClass(sp_table,tid)%>
         </select>          <select <%=banben_display%> name="leixing" class="sbe_button">
            <option value="" <%if leixing="" then response.Write("selected") end if%>>所有版本...</option>
		  <option value="1" <%if leixing="1" then response.Write("selected") end if%>>中文版</option>
		  <option value="2" <%if leixing="2" then response.Write("selected") end if%>>英文版</option>
        </select>
		 &nbsp;<input type="submit" name="Submit" value="搜索" class="sbe_button">
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
      
    <td height="63" colspan="11"><strong>正在载入数据，请稍候...</strong></td>
    </tr>
</table>
<%response.Flush()%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <tr class="sbe_table_title">
    <td class="sbe_table_title">操作</td>
    <td width="22%" height="25" class="sbe_table_title">
      名&nbsp;称</td>
    <td height="25" class="sbe_table_title">所属分类</td>
    <td height="25" class="sbe_table_title" style="display:none">起定量</td>
    <td height="25" class="sbe_table_title" style="display:none">交货期限</td>
    <td class="sbe_table_title">上移</td>
    <td height="25" class="sbe_table_title">下移</td>
    <td height="25" class="sbe_table_title" style="display:none;">首页推荐</td>
    <td height="25" class="sbe_table_title" <%=banben_display%>>类型</td>
    <td height="25" class="sbe_table_title">显示</td>
    <td height="25" class="sbe_table_title">编辑</td>
  </tr>
  <form name="form1" method="post" action="" onSubmit="return check()">
    <% 
		 if sp_records=0 then
		 %>
    <tr>
      <td height="25" colspan="11">暂且没有找到记录...</td>
    </tr>
    <%
   else
   rs.AbsolutePage=sp_page
   i=0
     while not rs.eof and i<sp_pagesize
	 i=i+1
   %>
    <tr onMouseOver="this.style.backgroundColor='#E9EFF3'" onMouseOut="this.style.backgroundColor=''">
      <td width="5%" align="center"><input name="id" type="checkbox" id="id" value="<%=rs(0)%>"></td>
      <td height="25"><font color="#0336699">
        <li type="circle"> <strong><a href="edit.asp?act=modify&id=<%=rs(0)%>"><%=RS(1)%></a> </strong></li>
      </font></td>
      <td width="17%" height="21" align="center"><%=ShowClassName(sp_table,rs(2))%></td>
      <td width="10%" align="center" style="display:none"><%=rs(3)%></td>
      <%'if ProShow1 Then%>
      <td width="9%" align="center" style="display:none"><%=rs(5)%></td>
      <td width="5%" align="center"><a href="list.asp?id=<%=rs(0)%>&act=up&keyword=<%=keyword%>&tid=<%=tid%>&leixing=<%=leixing%>"><img src="../images/up.gif" border="0" title="上移"></a></td>
      <td width="8%" align="center"><a href="list.asp?id=<%=rs(0)%>&act=down&keyword=<%=keyword%>&tid=<%=tid%>&leixing=<%=leixing%>"><img src="../images/downl.gif" border="0" title="下移"></a></td>
      <td width="8%" align="center" style="display:none;"><a href="?id=<%=rs(0)%>&act=tuijian">
        <%Call JudgeMent(rs(3))%>
      </a></td>
      <td width="8%" align="center" <%=banben_display%>><a href="?id=<%=rs(0)%>&act=leibie">
        <%Call JudgeMent1(rs(6))%>
      </a></td>
      <td width="8%" align="center"><a href="list.asp?id=<%=rs(0)%>&act=show">
        <%Call JudgeMent(Rs(4))%>
      </a></td>
      <td width="8%" align="center"><a href="edit.asp?act=modify&id=<%=rs(0)%>"><img src="../images/edit.gif" border="0"></a></td>
    </tr>
    <%Rs.movenext
     wend
	%>
    <tr>
      <td height="25" colspan="11"><input type="checkbox" name="ChkAll" onClick="SelectAll(this.form)">
        全选&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <input type="radio" name="act" value="del">
        删除
        <input type="radio" name="act" value="move">
        转移分类
        <select name="select" id="tid">
          <option selected>请选择类别...</option>
          <%Call ShowClass(sp_table,0)%>
        </select>
        <input type="submit" name="Submit2" value="执行操作" class="sbe_button"></td>
    </tr>
  </form>
  <tr>
    <td height="25" colspan="11"><% Call ShowPage("keyword="&keyword&"&tid="&tid&"&leixing="&leixing&"",sp_records,sp_pagesize,sp_page,true,2)%></td>
  </tr>
  <%end if
	Rs.close
	set Rs=nothing
  %>
</table>
</body>
</html>
<script language="JavaScript">
loading.style.display="none";
</script>
<%End Sub%>
<%
Sub UpSequence(ClassTitle,ID,Keyword,Tid,leixing)	
    set rsUp=server.CreateObject("adodb.recordset")
	sql="select Sequence from "&ClassTitle&" where ID="&ID
	rsUp.open sql,conn,1,3
	 set rs_up=server.CreateObject("adodb.recordset")
	 sql_up="select top 1 Sequence from "&ClassTitle&" where sequence>"&rsUp(0)
	 if tid<>"" then sql_up=sql_up&" and tid in("&ChildrenID(ClassTitle,tid)&")"
	 if keyword<>"" then sql_up=sql_up&" and pname like '%"&keyword&"%'"
	 if leixing<>"" then sql_up=sql_up&" and leibie ="&leixing&" "
	 sql_up=sql_up&" order by sequence"
	 rs_up.open sql_up,conn,1,3
	 if not rs_up.eof then
	    Temp_sequence=rs_up(0)
		rs_up(0)=rsUp(0)
		rs_up.update		
		rsUp(0)=Temp_sequence
		rsUp.update
     end if
	 rs_up.close
	 set rs_up=nothing
	rsUp.close
	set rsUp=nothing
  End Sub
  
  
Sub DownSequence(ClassTitle,ID,Keyword,Tid,leixing)
  set rs_DownSequence=server.CreateObject("adodb.recordset")
  sql="select Sequence from "&ClassTitle&" where id="&id
  rs_DownSequence.open sql,conn,1,3
     set rs_up=server.CreateObject("adodb.recordset")
	 sql_up="select top 1 Sequence from "&ClassTitle&" where sequence<"&rs_DownSequence(0)
	 if tid<>"" then sql_up=sql_up&" and tid in("&ChildrenID(ClassTitle,tid)&")"
	 if keyword<>"" then sql_up=sql_up&" and pname like '%"&keyword&"%'"
	 if leixing<>"" then sql_up=sql_up&" and leibie ="&leixing&" " 
	 sql_up=sql_up&" order by sequence desc"
	 rs_up.open sql_up,conn,1,3
	 if not rs_up.eof then
	    Temp_sequence=rs_up(0)
		rs_up(0)=rs_DownSequence(0)
		rs_up.update		
		rs_DownSequence(0)=Temp_sequence
		rs_DownSequence.update
     end if
	 rs_up.close
	 set rs_up=nothing
  rs_DownSequence.close  
  set rs_DownSequence=nothing
End Sub


%>