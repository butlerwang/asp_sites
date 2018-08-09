<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_EnterPrisePro
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_EnterPrisePro
        Private KS
		Private Action,i,strClass,sFileName,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
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
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  If KS.G("Action")<>"View" then
			   .Write "<div class='topdashed sort'>企业产品管理</div>"
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
		totalPut = Conn.Execute("Select Count(id) From KS_Product")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Select Case KS.G("action")
		 Case "Edit" Call ProEdit()
		 Case "EditSave" Call DoSave()
		 Case "Del" Call ProDel()
		 Case "verific"  Call Verify()
		 Case "unverific"  Call UnVerify()
		 Case "View" Call ShowNews()
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
%>

<script src="../ks_inc/kesion.box.js"></script>
<script>
function ShowIframe(id)
{
    new KesionPopup().PopupCenterIframe("查看产品","KS.EnterPrisePro.asp?action=View&ProID="+id,600,350,"auto")
}
</script>

<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td nowrap>产品名称</th>
	<td nowrap>添加</th>
	<td nowrap>产品型号</th>
	<td nowrap>产品价格</th>
	<td nowrap>属性</th>
	<td nowrap>状态</th>
	<td nowrap>管理操作</th>
</tr>
<%
	sFileName = "KS.EnterprisePro.asp?"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_Product order by id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>没有企业产品！</td></tr>"
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
	<td class="splittd"><a href="#" onclick="ShowIframe(<%=rs("id")%>)"><%=KS.Gottopic(Rs("Title"),45)%></a></td>
	<td class="splittd" align="center"><a href='../space/?<%=rs("inputer")%>' target='_blank'><%=Rs("inputer")%></a></td>
	<td class="splittd" align="center">&nbsp;<%=Rs("ProModel")%>&nbsp;</td>
	<td class="splittd" align="center"><%=Rs("Price")%> 元</td>
	<td class="splittd" align="center">
	 &nbsp;<% 
	 if rs("recommend")="1" then
	  response.write "<font color=blue>荐</font> "
	 end if
	 if rs("popular")="1" then
	  response.write "<font color=#ff6600>热</font> "
	 end if
	 if rs("istop")="1" then
	  response.write "顶"
	 end if
	 
	 %>
	</td>
	<td class="splittd" align="center"><%
	select case rs("verific")
	 case 0
	  response.write "<font color=red>未审</font>"
	 case 1
	  response.write "<font color=#999999>已审</font>"
	 case 2
	  response.write "<font color=blue>锁定</font>"
	end select
	%></td>
	<td class="splittd" align="center"><a href="#" onclick="ShowIframe(<%=rs("id")%>)">浏览</a> <a href="?Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('确定删除吗？'));">删除</a> <a href="?Action=verific&id=<%=rs("id")%>">审核</a></td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class=Button type="submit" name="Submit2" value=" 删除选中的记录" onclick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){this.form.Action.value='Del';this.form.submit();return true;}return false;}">
	<input type="button" value="批量审核" class="button" onclick="this.form.Action.value='verific';this.form.submit();">
	<input type="button" value="批量取消审核" class="button" onclick="this.form.Action.value='unverific';this.form.submit();">
	<input type="hidden" value="Del" name="Action">
	</td>
</tr>
</form>
<tr>
	<td colspan=10>
	<%
 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)	%></td>
</tr>
</table>

<%
End Sub

Sub ProEdit()
 Dim ID:ID=KS.ChkCLng(KS.G("id"))
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "Select * From KS_Product Where ID=" & ID,conn,1,1
 If RS.Eof And RS.Bof Then
  RS.Close:Set RS=Nothing
  Response.Write "<script>alert('参数传递出错！');history.back();</script>"
  Response.End
 End If
%>
<script>
function CheckForm()
{
if (document.myform.productname=='')
{
 alert('请输入产品名称');
 document.myform.productname.focus();
 return false;
}
document.myform.submit();
}
</script>
<br>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
 <form name="myform" action="?action=EditSave" method="post">
   <input type="hidden" value="<%=rs("id")%>" name="id">
   <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl">
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>产品名称：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='productname' value='<%=RS("productname")%>' size="40"> <font color=red>*</font></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>所属行业：</strong></td>
            <td height='28'>&nbsp;<%
		dim rss,sqls,count
		set rss=server.createobject("adodb.recordset")
		sqls = "select * from KS_enterpriseClass Where parentid<>0 order by orderid"
		rss.open sqls,conn,1,1
		%>
          <script language = "JavaScript">
		var onecount;
		subcat = new Array();
				<%
				count = 0
				do while not rss.eof 
				%>
		subcat[<%=count%>] = new Array("<%= trim(rss("id"))%>","<%=trim(rss("parentid"))%>","<%= trim(rss("classname"))%>");
				<%
				count = count + 1
				rss.movenext
				loop
				rss.close
				%>
		onecount=<%=count%>;
		function changelocation(locationid)
			{
			document.myform.SmallClassID.length = 0; 
			var locationid=locationid;
			var i;
			for (i=0;i < onecount; i++)
				{
					if (subcat[i][1] == locationid)
					{ 
						document.myform.SmallClassID.options[document.myform.SmallClassID.length] = new Option(subcat[i][2], subcat[i][0]);
					}        
				}
			}    
		
		</script>
		 <select class="face" name="BigClassID" onChange="changelocation(document.myform.BigClassID.options[document.myform.BigClassID.selectedIndex].value)" size="1">
		<% 
		dim rsb,sqlb
		set rsb=server.createobject("adodb.recordset")
        sqlb = "select * from ks_enterpriseClass where parentid=0 order by orderid"
        rsb.open sqlb,conn,1,1
		if rsb.eof and rsb.bof then
		else
		    do while not rsb.eof
					  If rs("BigClassID")=rsb("id") then
					  %>
                    <option value="<%=trim(rsb("id"))%>" selected><%=trim(rsb("ClassName"))%></option>
                    <%else%>
                    <option value="<%=trim(rsb("id"))%>"><%=trim(rsb("ClassName"))%></option>
                    <%end if
		        rsb.movenext
    	    loop
		end if
        rsb.close
			%>
                  </select>
                  <font color=#ff6600>&nbsp;*</font>
                  <select class="face" name="SmallClassID">
				   <option value='0'>--请选择-</option>
                    <%dim rsss,sqlss
						set rsss=server.createobject("adodb.recordset")
						sqlss="select * from ks_enterpriseclass where parentid="& rs("BigClassID")&" order by orderid"
						rsss.open sqlss,conn,1,1
						if not(rsss.eof and rsss.bof) then
						do while not rsss.eof
							  if rs("SmallClassID")=rsss("id") then%>
							<option value="<%=rsss("id")%>" selected><%=rsss("ClassName")%></option>
							<%else%>
							<option value="<%=rsss("id")%>"><%=rsss("ClassName")%></option>
							<%end if
							rsss.movenext
						loop
					end if
					rsss.close
					%>
                </select></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>产品价格：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='price' size='8' value='<%=RS("price")%>'> 元&nbsp;&nbsp;&nbsp;&nbsp;<strong>产品属性：</strong><input type='checkbox' name='recommend' value='1'<%if rs("recommend")="1" then response.write " checked"%>>推荐 <input type='checkbox' name='popular' value='1'<%if rs("popular")="1" then response.write " checked"%>>热门 <input type='checkbox' name='istop' value='1'<%if rs("istop")="1" then response.write " checked"%>>固顶</td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>产品产地：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='address' value='<%=RS("address")%>'></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>产品型号：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='promodel' value='<%=RS("promodel")%>'></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>详细介绍：</strong></td>
           <td height='28'> <%
		     Response.Write "<textarea id=""Intro"" name=""Intro"">"& KS.HTMLCode(rs("Intro")) &"</textarea>"
								%>		
		   </td>
          </tr> 
		  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>产品图片：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='photourl' size="45" value='<%=RS("photourl")%>'></td>
          </tr>  
		  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>发 布 人：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='username' size="45" value='<%=RS("username")%>'></td>
          </tr>  
 
		 </form>  
</table>
<%
RS.Close:Set RS=Nothing
End Sub

Sub DoSave
	Dim ID:ID=KS.ChkCLng(KS.G("id"))
	Dim RS:Set RS=Server.CreateOBject("ADODB.RECORDSET")
	RS.Open "Select * From KS_Product Where Id=" & ID,conn,1,3
	If RS.Eof And RS.Bof Then
	 RS.Close:Set RS=Nothing
	 Response.Write "<script>alert('参数传递出错！');history.back();</script>"
	 Response.End
	End If
	RS("ProductName")=KS.G("ProductName")
	RS("BigClassID")=KS.ChkCLng(KS.G("BigClassID"))
	RS("SmallClassID")=KS.ChkCLng(KS.G("SmallClassID"))
	RS("Price")=KS.ChkClng(KS.G("Price"))
	RS("Address")=KS.G("Address")
	RS("Intro")=KS.HtmlEncode(Request.Form("Intro"))
	RS("PhotoUrl")=KS.G("PhotoUrl")
	RS("UserName")=KS.G("UserName")
	RS("Recommend")=KS.ChkClng(KS.G("Recommend"))
	RS("Popular")=KS.ChkCLng(KS.G("Popular"))
	RS("Istop")=KS.ChkClng(KS.G("Istop"))
	RS.Update
	RS.Close:Set RS=Nothing
	Response.Write "<script>alert('恭喜，产品修改成功！');location.href='" & Request.Form("ComeUrl") & "';</script>"
End Sub

'删除
Sub ProDel()
 on error resume next
 Dim I,ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 ID=Split(ID,",")
 For I=0 To Ubound(ID)
  KS.DeleteFile(conn.execute("select photourl from KS_Product where id=" & ID(I))(0))
  Conn.execute("Delete From KS_Product Where id="& id(I))
 Next 
 Response.Write "<script>alert('删除成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

'审核
Sub ShowNews()
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_Product where id=" &KS.ChkClng(KS.S("ProID")),conn,1,1
		If Not RS.Eof Then
		   Response.WRITE "<div style='padding:30px'><div><strong>产品名称：</strong>" & rs("title") & "</div>"
		   Response.Write "<div style=""text-align:left""><strong>产品型号：</strong>" & RS("ProModel") & "</div>"
		   Response.Write "<div style=""text-align:left""><strong>商品规格：</strong>" & RS("ProSpecificat") & "</div>"
		   Response.Write "<div style=""text-align:left""><strong>产品价格：</strong>" & RS("price") & " 元</div>"
		   Response.Write "<div style=""text-aling:left""><strong>产品属性：</strong>"
			 if rs("recommend")="1" then
			  response.write "<font color=blue>荐</font> "
			 end if
			 if rs("popular")="1" then
			  response.write "<font color=#ff6600>热</font> "
			 end if
			 if rs("istop")="1" then
			  response.write "顶"
			 end if
	      response.write "</div>"
	  
		   Dim PhotoUrl:PhotoUrl=RS("PhotoUrl")
		   If PhotoUrl<>"" And Not IsNull(PhotoURL) Then
		   Response.Write "<div style=""text-align:left"">产品图片：<img src='" & RS("photourl") & "'></div>"
		   End If
		   Response.Write "<div>产品介绍：" & KS.HTMLCode(rs("prointro")) & "</div>"
		   Response.Write "</div>"
		End If
		RS.Close:Set RS=Nothing
End Sub
'审核
Sub Verify
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_Product Set verific=1 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'取消审核
Sub UnVerify
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_Product Set verific=0 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

End Class
%> 
