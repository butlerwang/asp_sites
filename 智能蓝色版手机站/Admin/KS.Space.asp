<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Space
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Space
        Private KS,Param
		Private Action,i,strClass,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		 With Response
					If Not KS.ReturnPowerResult(0, "KSMS10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
			  .Write "<html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""location.href='KS.Space.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>空间管理</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.SpaceSkin.asp?flag=2';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>模板管理</span></li>"
			  .Write "<li class='parent' onclick=""location.href='?action=class';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>空间分类</span></li>"
			  .Write "<li></li>"
			  if request("showtype")<>"1" then
			  .Write "<div><select name='classid' onchange=""location.href='?classid='+this.value;"">"
			   Dim RSC:Set RSC=Conn.Execute("Select ClassID,ClassName From KS_BlogClass order by orderid")
			   .Write "<option value=''>---按博客分类查看---</option>"
			   Do While Not RSC.Eof
			    .Write "<option value='" & RSC(0) & "'>" & rsc(1) & "</option>"
				rsc.movenext
			   Loop
			   RSC.Close
			   Set RSC=Nothing
			  .Write "</select></div>"
			  End If
			  .Write "</ul>"
		End With
		
		
		maxperpage = 30 '###每页显示数
		If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
			Response.Write ("错误的系统参数!请输入整数")
			Response.End
		End If
		If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
			CurrentPage = CInt(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CInt(CurrentPage) = 0 Then CurrentPage = 1
		
		
		Select Case KS.G("action")
		 Case "Del"
		  Call BlogDel()
		 Case "lock"
		  Call BlogLock()
		 Case "unlock"
		  Call BlogUnLock()
		 Case "verific"
		  Call Blogverific()
		 Case "recommend"
		  Call Blogrecommend()
		 Case "Cancelrecommend"
		  Call BlogCancelrecommend()
		 Case "add","modify" Call Modify()
		 case "modifysave" call ModifySave()
		 Case "class" ShowClass
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
    Dim classname
	if request("showtype")="1" then
		Param=" inner join ks_user u on a.username=u.username where u.usertype=0"
    ElseIf KS.S("ClassID")<>"" Then
	   classname="b.classname,"
	   Param=" left join ks_BLOGClass b on a.classid=b.classid where A.classid=" & KS.ChkClng(KS.G("ClassID"))
	Else
		Param=" where 1=1"
	End If
	
		if request("from")<>"" then
		 param=param & " and status=0"
		end if
		  
		If KS.G("KeyWord")<>"" Then
		  If KS.G("condition")=1 Then
		   Param= Param & " and blogname like '%" & KS.G("KeyWord") & "%'"
		  Else
		   Param= Param & " and username like '%" & KS.G("KeyWord") & "%'"
		  End If
		End If

		totalPut = Conn.Execute("Select Count(blogID) from KS_Blog a " & Param)(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td nowrap>站点名称</th>
	<td nowrap>创建者</th>
	<td nowrap>创建时间</th>
	<td nowrap>状态</th>
	<td nowrap>管理操作</th>
</tr>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select " & classname & " a.* from KS_Blog a  "& Param & " order by blogid desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>没有用户申请个人空间！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action="?">
<input type="hidden" name="action" id="action" value="Del">
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=rs("blogid")%>'></td>
	<td class="splittd">
	<%if request("showtype")="" and request("classid")<>"" then%>
	[<%=RS(0)%>]
	<%end if%>
	<a href="../space/?<%=rs("userid")%>" target="_blank"><%=Rs("blogname")%></a></td>
	<td class="splittd" align="center"><%=Rs("username")%></td>
	<td class="splittd" align="center"><%=Rs("adddate")%></td>
	<td class="splittd" align="center"><%
	select case rs("status")
	 case 0
	  response.write "未审"
	 case 1
	  response.write "<font color=red>已审</font>"
	 case 2
	  response.write "<font color=blue>锁定</font>"
	end select
	%></td>
	<td class="splittd" align="center"><a href="../space/?<%=rs("username")%>" target="_blank">浏览</a> <a href="?action=modify&id=<%=rs("blogid")%>" onclick="window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('空间门户管理 >> <font color=red>修改空间信息</font>')+'&ButtonSymbol=GOSave';">编辑</a> <a href="?Action=Del&ID=<%=rs("blogid")%>" onclick="return(confirm('删除站点将删除站点下的所有日志，确定删除吗？'));">删除</a> <%IF rs("recommend")="1" then %><a href="?Action=Cancelrecommend&id=<%=rs("blogid")%>"><font color=red>取消推荐</font></a><%else%><a href="?Action=recommend&id=<%=rs("blogid")%>">设为推荐</a><%end if%>&nbsp;<%if rs("status")=0 then%><a href="?Action=verific&id=<%=rs("blogid")%>">审核</a> <%elseif rs("status")=1 then%><a href="?Action=lock&id=<%=rs("blogid")%>">锁定</a><%elseif rs("status")=2 then%><a href="?Action=unlock&id=<%=rs("blogid")%>">解锁</a><%end if%>
	<a onclick="window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('空间门户管理 >> <font color=red>升级为企业空间</font>')+'&ButtonSymbol=GOSave';" href="KS.EnterPrise.asp?action=Edit&flag=update&username=<%=rs("username")%>" title="升级为企业空间">升级</a>
	</td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td  class="splittd" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class=Button type="submit" name="Submit2" value=" 删除选中的空间 " onclick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){$('#action').val('Del');this.document.selform.submit();return true;}return false;}">
	<input type="submit" value="批量审核/解锁" onclick="$('#action').val('verific');" class="button">
	<input type="submit" value="批量锁定" onclick="$('#action').val('lock');" class="button">
	<%if request("showtype")="1" then%>
	<input type="button" value="开通个人空间" onclick="window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('空间门户管理 >> <font color=red>开通个人空间</font>')+'&ButtonSymbol=GO';location.href='KS.Space.asp?action=add';" class="button">
	<%end if%>
	</td>
</tr>
</form>
<tr>
	<td colspan=10 align=right>
	<%
	 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>
<div>
<form action="KS.Space.asp" name="myform" method="get">
   <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
      &nbsp;<strong>快速搜索=></strong>
	 &nbsp;关键字:<input type="text" class='textbox' name="keyword">&nbsp;条件:
	 <select name="condition">
	  <option value=1>站点名称</option>
	  <option value=2>用 户 名</option>
	 </select>
	  &nbsp;<input type="submit" value="开始搜索" class="button" name="s1">
	  </div>
</form>
</div>
<%
End Sub

Sub Modify()
 Dim BlogName,username,domain,logo,ClassID,Descript,Announce,ContentLen,ListBlogNum,ListReplayNum,ListLogNum,ListGuestNum,status,templateid
 Dim ID:ID=KS.ChkClng(Request("id"))
 If Id<>0 Then
	 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	 RS.Open "select * from ks_blog where blogid=" & id,conn,1,1
	 If RS.Eof AND RS.Bof Then
	   RS.Close
	   Set RS=Nothing
	   KS.AlertHintScript "对不起，找不到记录！"
	 End If
	 BlogName = RS("BlogName")
	 username = RS("username")
	 domain   = RS("domain")
	 logo     = RS("logo")
	 ClassID  = RS("ClassID")
	 Descript = RS("Descript")
	 Announce = RS("Announce")
	 ContentLen=RS("ContentLen")
	 ListBlogNum=RS("ListBlogNum")
	 ListReplayNum=RS("ListReplayNum")
	 ListLogNum=RS("ListLogNum")
	 ListGuestNum= RS("ListGuestNum")
	 templateid  = RS("templateid")
	 status   = RS("status")
 Else 
     Announce = "暂无公告!"
	 ContentLen=500
	 ListBlogNum=10
	 ListReplayNum=10
	 ListLogNum=10
	 ListGuestNum= 10
     status=1
 End If
 %>
 <script type="text/javascript">
 function CheckForm()
 {
 <%if request("action")="add" then%>
   if ($("input[name=username]").val()=='')
   {
     alert('用户名称必须输入！');
	 $("input[name=username]").focus();
	 return false;
   }
 <%end if%>
   if ($("input[name=BlogName]").val()=='')
   {
     alert('空间名称必须输入！');
	 $("input[name=BlogName]").focus();
	 return false;
   }
   $("#myform").submit();
 }
 </script>
 <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
 <form name="myform" id="myform" action="?action=modifysave" method="post">
   <input type="hidden" value="<%=ID%>" name="id">
   <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl">
      <%if request("action")="add" then%>
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>用户名：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='username' id='username'/> <font color=red>*</font>请输入要开通个人空间的用户名</td>
          </tr>
	  <%else%>
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>创建人：</strong></td>
           <td height='28'>&nbsp;<%=username%></td>
          </tr>
	  <%end if%>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>空间名称：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='BlogName' value='<%=BlogName%>' size="40"> <font color=red>*</font></td>
          </tr> 
		  
		 <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>绑定模板：</strong></td>
           <td height='28'>&nbsp;<select name="templateid" id='templateid'>
		   <option value='0'>--选择模板--</option>
		   <%
		   dim flag
		   if request("action")="modify"  then
		     if conn.execute("select top 1 * from ks_enterprise where username='" & username & "'").eof then
		      flag=2
			 else
			  flag=4
			 end if
		   else
		    flag=2
		   end if
		   set rs=conn.execute("select * from KS_BlogTemplate where flag=" & flag &" order by id desc")
		   do while not rs.eof
		      if trim(templateid)=trim(rs("id")) then
		      response.write "<option value='" & rs("id") & "' selected>" & rs("templatename") &"</option>"
			  else
		      response.write "<option value='" & rs("id") & "'>" & rs("templatename") &"</option>"
			  end if
		   rs.movenext
		   loop
		   rs.close
		   set rs=nothing
		   %>
		   </select></td>
          </tr>
		   
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>空间域名：</strong></td>
           <td height='28'>&nbsp;
		   <label><input type="radio" <%if instr(domain,".")=0 then response.write " checked"%> name="domaintype" onclick="$('#domain0').show();$('#domain1').hide();" value="0" />二级域名</label>
			      <label><input type="radio" <%if instr(domain,".")<>0 then response.write " checked"%> name="domaintype" onclick="$('#domain0').hide();$('#domain1').show();" value="1" />顶级独立域名</label>
				  <div id='domain0'<%if instr(domain,".")<>0 then response.write " style='display:none'"%>>
                  &nbsp;<input class="textbox" name="domain" type="text" id="domain" style="width:50px; " value="<%=domain%>" maxlength="100" /><b>.<%response.write KS.SSetting(16)%></b> <span class="msgtips">如果不想绑定可以留空</span>
				  </div>
				  <div id='domain1' <%if instr(domain,".")=0 or KS.IsNul(domain) then response.write " style='display:none'"%>>
				    &nbsp;<input class="textbox" name="mydomain" type="text" id="mydomain" style="width:150px; " value="<%=domain%>" maxlength="100" /> <span class="msgtips">如# 需要将您的域名解释到本站服务器IP上。</span>
				  </div> 
		   
		   </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>空间Logo：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='logo' value='<%=logo%>' size="40"></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>空间分类：</strong></td>
           <td height='28'>&nbsp;<select class="textbox" size='1' name='ClassID' style="width:250">
                    <option value="0">-请选择类别-</option>
                    <% Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
							  RSC.Open "Select * From KS_BlogClass order by orderid",conn,1,1
							  If Not RSC.EOF Then
							   Do While Not RSC.Eof 
							   If Trim(ClassID)=trim(RSC("ClassID")) Then
								  Response.Write "<option value=""" & RSC("ClassID") & """ selected>" & RSC("ClassName") & "</option>"
							   Else
								  Response.Write "<option value=""" & RSC("ClassID") & """>" & RSC("ClassName") & "</option>"
							   End iF
								 RSC.MoveNext
							   Loop
							  End If
							  RSC.Close:Set RSC=Nothing
							  %>
                  </select>   </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>站点描述：</strong></td>
           <td height='28'>&nbsp;<textarea class="textbox" name="Descript" id="Descript" style="width:80%;height:60px" cols=50 rows=6><%=Descript%></textarea> </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>空间公告：</strong></td>
           <td height='28'>&nbsp;<textarea class="textbox" name="Announce" id="Announce" style="width:80%;height:80px" cols=50 rows=6><%=Announce%></textarea> </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>每页显示新鲜事条数：</strong></td>
           <td height='28'>&nbsp;<input class="textbox" name="ContentLen" type="text" id="ContentLen" style="width:250px; " value="<%=ContentLen%>" />            </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>每页显示日志篇数：</strong></td>
           <td height='28'>&nbsp;<input class="textbox" name="ListBlogNum" type="text" id="ListBlogNum" style="width:250px; " value="<%=ListBlogNum%>" />  </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>显示最新回复条数：</strong></td>
           <td height='28'>&nbsp;<input class="textbox" name="ListReplayNum" type="text" id="ListReplayNum" style="width:250px; " value="<%=ListReplayNum%>" />  </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>显示最新日志篇数：</strong></td>
           <td height='28'>&nbsp;<input class="textbox" name="ListLogNum" type="text" id="ListLogNum" style="width:250px; " value="<%=ListLogNum%>" />             </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>显示最新留言条数：</strong></td>
           <td height='28'>&nbsp;<input class="textbox" name="ListGuestNum" type="text" id="ListGuestNum" style="width:250px; " value="<%=ListGuestNum%>" />  </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>状态：</strong></td>
           <td height='28'>&nbsp;<input name="Status" type="radio" value="1"<%if status=1 then response.write " checked"%> /> 已审核 <input name="Status" type="radio" value="0" <%if status=0 then response.write " checked"%>/> 未审核<input name="Status" type="radio" value="2" <%if status=2 then response.write " checked"%>/> 锁定</td>
          </tr>  
         
   </form>
 </table>
 <%
End Sub

Sub ModifySave()
 Dim ID:ID=KS.ChkClng(Request("id"))
 Dim domaintype:domaintype=KS.ChkClng(KS.G("domaintype"))
 Dim templateid:templateid=KS.ChkClng(KS.G("templateid"))
 Dim Domain:Domain=KS.G("Domain")
 If DomainType=1 Then Domain=KS.G("mydomain")
 Dim UserID,UserName:UserName=KS.G("UserName")
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 If ID=0 Then  
   If KS.IsNul(UserName) Then
    KS.Die "<script>alert('请输入要开通空间的用户名!');history.back();</script>"
   End If
   RS.Open "select top 1 userid from KS_User Where UserName='" & UserName & "'",conn,1,1
   If RS.Eof And RS.Bof Then
     RS.Close
	 Set RS=Nothing
	 KS.Die "<script>alert('对不起，您输入的用户名不存在!');history.back();</script>"
   End If
   UserID=RS(0)
   RS.Close
 End If
 RS.Open "select top 1 * from ks_blog where blogid=" & id,conn,1,3
 If RS.Eof AND RS.Bof Then
   RS.addnew
   RS("UserName")=UserName
   RS("UserID")=UserId
   RS("AddDate")=Now
 End If
 RS("TemplateID")=TemplateId
 RS("BlogName")=KS.G("BlogName")
 RS("Domain")=Domain
 RS("Logo")=KS.G("Logo")
 RS("ClassID")=KS.ChkClng(KS.G("ClassID"))
 RS("Descript")=KS.G("Descript")
 RS("Announce")=KS.G("Announce")
 RS("ContentLen")=KS.ChkClng(KS.G("ContentLen"))
 RS("ListBlogNum")=KS.ChkClng(KS.G("ListBlogNum"))
 RS("ListReplayNum")=KS.ChkClng(KS.G("ListReplayNum"))
 RS("ListLogNum")=KS.ChkClng(KS.G("ListLogNum"))
 RS("ListGuestNum")=KS.ChkClng(KS.G("ListGuestNum"))
 RS("Status")=KS.ChkClng(KS.G("Status"))
 UserName=RS("UserName")
 RS.Update
 RS.Close
 Set RS=Nothing
  Conn.Execute("Update KS_EnterpriseNews Set [Domain]='" & Domain & "' Where UserName='" & UserName &"'")
  if id=0 then
 Response.Write "<script>alert('恭喜，空间添加成功！');$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr='+escape('空间门户管理 >> <font color=red>空间站点管理</font>');location.href='KS.Space.asp?showtype=1';</script>"
  else
 Response.Write "<script>alert('恭喜，空间修改成功！');$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr='+escape('空间门户管理 >> <font color=red>空间站点管理</font>');location.href='"& Request.Form("ComeUrl") & "';</script>"
 end if
End Sub

'删除日志
Sub BlogDel()
 Dim ID:ID=KS.G("ID")
 Dim UserName
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Dim RS:Set RS=Server.CreateOBject("ADODB.RECORDSET")
 RS.Open "Select * from KS_Blog Where BlogID in(" & id & ")",conn,1,1
 do while not rs.eof
  username=rs("username")
  Conn.execute("Delete From KS_BlogInfo Where username='" & username & "'")
  Conn.Execute("Delete From KS_BlogComment Where username='" & username &"'")
  Conn.execute("Delete From KS_BlogMessage Where Username='" & username & "'")
  rs.movenext
 loop
 rs.close:set rs=nothing
 Conn.Execute("Delete From KS_UploadFiles Where ChannelID=1025 and infoID in(" & ID & ")")
 Conn.execute("Delete From KS_Blog Where BlogID In("& id & ")")
 Response.Write "<script>alert('删除成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'设为精华
Sub Blogrecommend()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_Blog Set recommend=1 Where BlogID In("& id & ")")
 Conn.execute("Update KS_EnterPrise Set recommend=1 Where username In(select username from ks_blog where blogid in("& id & "))")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'取消精华
Sub BlogCancelrecommend()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_Blog Set recommend=0 Where BlogID In("& id & ")")
 Conn.execute("Update KS_EnterPrise Set recommend=0 Where username In(select username from ks_blog where blogid in("& id & "))")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'锁定
Sub BlogLock()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_Blog Set status=2 Where BlogID In("& id & ")")
 conn.execute("update ks_enterprise set status=2 where username in(select username from ks_blog where blogid in("&id&"))")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'解锁
Sub BlogUnLock()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_Blog Set status=1 Where BlogID In("& id & ")")
 conn.execute("update ks_enterprise set status=1 where username in(select username from ks_blog where blogid in("&id&"))")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'审核
Sub Blogverific
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_Blog Set status=1 Where BlogID In("& id & ")")
 conn.execute("update ks_enterprise set status=1 where username in(select username from ks_blog where blogid in("&id&"))")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

Sub ShowClass
%>		
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
		  <tr align="center"  class="sort"> 
			<td width="87"><strong>编号</strong></td>
			<td width="217"><strong>类型名称</strong></td>
			<td width="197"><strong>排序</strong></td>
			<td width="196"><strong>管理操作</strong></td>
		  </tr>
		  <%dim orderid
		  set rs = conn.execute("select * from KS_BlogClass order by orderid")
		    if rs.eof and rs.bof then
			  Response.Write "<tr><td colspan=""6"" height=""25"" align=""center"" class=""list"">还没有添加任何的博客类型!</td></tr>"
			else
			   do while not rs.eof%>
			  <form name="form1" method="post" action="?action=class&x=a">
				<tr  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
				  <td class="splittd" width="87" height="22" align="center"><%=rs("ClassID")%> <input name="ClassID" type="hidden" id="ClassID" value="<%=rs("ClassID")%>"></td>
				  <td class="splittd" width="217" align="center"><input name="ClassName" type="text" class="textbox" id="ClassName" value="<%=rs("ClassName")%>" size="25"></td>
				  <td class="splittd" width="197" align="center"><input style="text-align:center" name="OrderID" type="text" class="textbox" id="OrderID" value="<%=rs("OrderID")%>" size="8">				  </td>
				  <td class="splittd" align="center"><input name="Submit" class="button" type="submit"value=" 修改 ">&nbsp;<a  onclick='return(confirm("确定删除吗？"))' href="?action=class&x=c&ClassID=<%=rs("ClassID")%>">删除</a></td>
				</tr>
			  </form>
		  <%orderid=rs("orderid")
		   rs.movenext
		   loop
		 End IF
		rs.close%>
				<form action="?action=class&x=b" method="post" name="myform" id="form">
		    <tr>
		      <td class="splittd" height="25" colspan="4">&nbsp;&nbsp;<strong>&gt;&gt;新增空间分类<<</strong></td>
		    </tr>

			<tr valign="middle" class="list"> 
			  <td height="25"></td>
			  <td height="25" align="center"><input name="ClassName" type="text" class="textbox" id="ClassName" size="25"></td>
			  <td height="25" align="center"><input style="text-align:center" name="orderid" type="text" value="<%=orderid+1%>" class="textbox" id="orderid" size="8">
			  <td height="25" align="center"><input name="Submit3" class="button" type="submit" value="OK,提交"></td>
			</tr>
		</form>
</table>

		<% Select case request("x")
		   case "a"
				conn.execute("Update KS_BlogClass set ClassName='" & KS.G("ClassName") & "',orderid='" & KS.ChkClng(KS.G("OrderID")) &"' where ClassID="&KS.G("ClassID")&"")
				KS.AlertHintScript "恭喜,空间分类修改成功"
		   case "b"
		       If KS.G("ClassName")="" Then KS.Die "<script>alert('请输入类型名称!');history.back();</script>"
			   
				conn.execute("Insert into KS_BlogClass(ClassName,orderid)values('" & KS.G("ClassName") & "','" & KS.ChkClng(KS.G("OrderID")) &"')")
				KS.AlertHintScript "恭喜,空间分类添加成功"
		   case "c"
				conn.execute("Delete from KS_BlogClass where ClassID="&KS.G("ClassID")&"")
				KS.AlertHintScript "恭喜,空间分类删除成功"
		End Select
		%></body>
		</html>
<%End Sub
End Class
%> 
