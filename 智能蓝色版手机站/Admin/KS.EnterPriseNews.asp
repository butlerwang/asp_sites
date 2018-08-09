<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_EnterPriseNews
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_EnterPriseNews
        Private KS
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
					If Not KS.ReturnPowerResult(0, "KSMS10009") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
			  If KS.G("Action")="View" Then Call ShowNews():Exit Sub
			  .Write "<html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../ks_inc/Common.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  If KS.G("Action")<>"View" then
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""location.href='KS.Enterprise.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>企业管理</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.SpaceSkin.asp?flag=4';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>模板管理</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.EnterPrisePro.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>企业产品</span></li>"
			  .Write "</ul>"
			 End If
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
		totalPut = Conn.Execute("Select Count(id) From KS_EnterPriseNews")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Select Case KS.G("action")
		 Case "Del" Call BlogDel()
		 Case "verific"  Call Verify()
		 Case "unverific"  Call UnVerify()
		 Case "View" Call ShowNews()
		 Case "modify" call modify()
		 Case "DoSave" call DoSave()
		 Case Else  Call showmain
		End Select
End Sub

Sub Modify()
 Dim ID:id=KS.ChkClng(Request("id"))
 If ID=0 Then KS.Die "error!"
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "select top 1 * From KS_EnterpriseNews Where ID=" & ID,conn,1,1
 If RS.Eof And RS.Bof Then
   RS.Close:Set RS=Nothing
   KS.Die "<script>alert('出错啦!');history.back();</script>"
 End If
 
	%>
		<script language = "JavaScript">
				function CheckForm()
				{	
				if (document.myform.Title.value=="")
				  {
					alert("请输入新闻标题！");
					document.myform.Title.focus();
					return false;
				  }	
		
				    if (CKEDITOR.instances.Content.getData()=="")
					{
					  alert("新闻内容不能留空！");
					  CKEDITOR.instances.Content.focus();
					  return false;
					}
				 return true;  
				}
				</script>
				
				
				<table  width="99%" style="margin-top:5px" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form  action="?Action=DoSave&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
				    <tr  class="title">
					  <td colspan=2 align=center>修改企业新闻
					 </td>
					</tr>
                    <tr class="tdbg">
                       <td width="12%"  height="25" align="center"><span>新闻标题：</span></td>
                       <td width="88%"><input class="textbox" name="Title" type="text" id="Title" style="width:250px; " value="<%=rs("Title")%>" maxlength="100" />
                                        <span style="color: #FF0000">*</span> </td>
                    </tr>
					
						  
                     <tr class="tdbg">
                                <td align="center">发布时间：</td>
                                <td><input class="textbox" readonly name="AddDate" type="text" style="width:250px; " value="<%=rs("AddDate")%>" maxlength="100" /></td>
                              </tr>
                              <tr class="tdbg">
                                  <td align="center">新闻内容：</td>
								  <td>
							<%	
								Response.Write "<textarea name=""Content"" style=""display:none"">" & KS.HtmlCode(rs("Content")) & "</textarea>"
							%>	
							<script type="text/javascript" src="../editor/ckeditor.js" mce_src="../editor/ckeditor.js"></script>
							<script type="text/javascript">
						CKEDITOR.replace('Content', {width:"98%",height:"300px",toolbar:"Basic",filebrowserBrowseUrl :"../editor/ksplus/SelectUpFiles.asp",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
						</script> 
							</td>
                    </tr>
					
                    <tr class="tdbg">
                       <td width="12%"  height="25" align="center"><span>状态：</span></td>
                       <td width="88%"><input type="radio" name="status" value="1" <%if rs("status")="1" then response.write " checked"%>/>已审
					   <input type="radio" name="status" value="0" <%if rs("status")="0" then response.write " checked"%>/>未审
					   </td>
                    </tr>
					
                    <tr class="tdbg">
					  <td></td>
                      <td height="30">
					   <input type="submit" value="OK, 保 存" class="button"/>
					 </td>
                    </tr>
                  </form>
			    </table>
		  <%
	RS.Close:Set RS=Nothing
End Sub

Sub DoSave()
      Dim Title,Content,AddDATE
      Dim Id:Id=KS.ChkClng(Request("ID"))
	  Title=KS.S("Title")
	  Content=KS.HtmlEncode(Request.Form("Content"))
	  AddDate=KS.G("AddDate")
	  If NOt IsDate(AddDate) Then 
	  	Response.Write "<script>alert('日期格式不正确!');history.back();</script>"
		Exit Sub
	  End If
	  Dim RSObj
				  
				  If Title="" Then
				    Response.Write "<script>alert('你没有输入新闻标题!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Content="" Then
				    Response.Write "<script>alert('你没有输入新闻内容!');history.back();</script>"
				    Exit Sub
				  End IF
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select top 1 * From KS_EnterpriseNews Where ID=" & Id,Conn,1,3
				If not rsobj.eof then
				  
				  RSObj("Adddate")=KS.G("AddDate")
				  RSObj("Title")=Title
				  RSObj("Content")=Content
				  RSObj("Status")=KS.ChkClng(KS.S("Status"))
				 RSObj.Update
				End If
				 RSObj.Close:Set RSObj=Nothing
				 Response.Write "<script>alert('企业新闻修改成功!');location.href='KS.EnterpriseNews.asp';</script>"
  End Sub

Private Sub showmain()
%>
<script src="../ks_inc/kesion.box.js"></script>
<script>
function ShowIframe(id)
 {
    new KesionPopup().PopupCenterIframe("查看新闻","KS.EnterPriseNews.asp?action=View&newsid="+id,600,350,"auto")
 }
</script>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td nowrap>新闻标题</th>
	<td nowrap>添加</th>
	<td nowrap>添加时间</th>
	<td nowrap>状态</th>
	<td nowrap>管理操作</th>
</tr>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_EnterpriseNews order by id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>没有企业新闻！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action="?">
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td class="splittd"><a href="#" onclick="ShowIframe(<%=rs("id")%>)"><%=Rs("title")%></a></td>
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
	<td class="splittd" align="center">
	<a href="?action=modify&id=<%=rs("id")%>">修改</a>
	<a href="#" onclick="ShowIframe(<%=rs("id")%>)">浏览</a> <a href="?Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('确定删除吗？'));">删除</a> <a href="?Action=verific&id=<%=rs("id")%>">审核</a></td>
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
	<td class="splittd" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class=Button type="submit" name="Submit2" value=" 删除选中的新闻" onclick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){this.form.Action.value='Del';this.document.selform.submit();return true;}return false;}">
	<input type="button" value="批量审核" class="button" onclick="this.form.Action.value='verific';this.form.submit();">
	<input type="button" value="批量取消审核" class="button" onclick="this.form.Action.value='unverific';this.form.submit();">
	<input type="hidden" value="Del" name="Action">
	</td>
</tr>
</form>
<tr>
	<td colspan=7>
	<%
	 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>

<%
End Sub

'删除日志
Sub BlogDel()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Delete From KS_EnterPrisenews Where id In("& id & ")")
 Response.Write "<script>alert('删除成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub


Sub ShowNews()
	With Response	
		 .Write "<html>"
		 .Write"<head>"
		 .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		 .Write"<link href=""Include/Admin_style.CSS"" rel=""stylesheet"" type=""text/css"">"
		 .Write "<script src=""../ks_inc/Common.js"" language=""JavaScript""></script>"
		 .Write"</head>"
		 .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"

	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_EnterPriseNews where id=" &KS.ChkClng(KS.S("NewsID")),conn,1,1
		If Not RS.Eof Then
		   .WRITE "<div style=""margin-top:6px;font-weight:bold;text-align:center"">" & rs("title") & "</div>"
		   .Write "<div style=""text-align:center"">作者：" & RS("UserName") & "&nbsp;&nbsp;&nbsp;&nbsp;时间:" & RS("AddDate") & "</div>"
		   .Write "<hr size=1><div>" & KS.HTMLCode(rs("content")) & "</div>"
		End If
		RS.Close:Set RS=Nothing
   End With
End Sub
'审核
Sub Verify()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_EnterpriseNews Set status=1 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'取消审核
Sub UnVerify()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_EnterpriseNews Set status=0 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

End Class
%> 
