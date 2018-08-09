<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_EnterPriseZS
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_EnterPriseZS
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
					If Not KS.ReturnPowerResult(0, "KSMS10011") Then          '检查是权限
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
			  .Write "<div class='topdashed sort'>企业荣誉证书管理</div>"
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
		totalPut = Conn.Execute("Select Count(id) From KS_EnterpriseZS")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Select Case KS.G("action")
		 Case "Del" Call DelRecord()
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
onscrolls=false;
new KesionPopup().PopupCenterIframe("查看企业荣誉证书","KS.EnterpriseZS.asp?action=View&ProID="+id,550,350,'auto')
}
</script>

<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</td>
	<td nowrap>证书名称</td>
	<td nowrap>添加</td>
	<td nowrap>发证机构</td>
	<td nowrap>生效日期</td>
	<td nowrap>截止日期</td>
	<td nowrap>状态</td>
	<td nowrap>管理操作</td>
</tr>
<%
	sFileName = "KS.EnterpriseZS.asp?"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_EnterpriseZS order by id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>没有用户上传企业荣誉证书！</td></tr>"
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
	<td class="splittd"><a href="#" onclick="ShowIframe(<%=rs("id")%>)"><%=Rs("Title")%></a></td>
	<td class="splittd" align="center"><a href='../space/?<%=rs("username")%>' target='_blank'><%=Rs("username")%></a></td>
	<td class="splittd" align="center"><%=Rs("fzjg")%></td>
	<td class="splittd" align="center"><%=Rs("sxrq")%></td>
	<td class="splittd" align="center"><%=Rs("jzrq")%></td>
	<td class="splittd" align="center"><%
	select case rs("status")
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
<tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td  class="splittd" height='25' colspan=8>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class=Button type="submit" name="Submit2" value=" 删除选中的证书" onclick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){this.form.Action.value='Del';this.form.submit();return true;}return false;}">
	<input type="button" value="批量审核" class="button" onclick="this.form.Action.value='verific';this.form.submit();">
	<input type="button" value="批量取消审核" class="button" onclick="this.form.Action.value='unverific';this.form.submit();">
	<input type="hidden" value="Del" name="Action">
	</td>
</tr>
</form>
<tr>
	<td colspan=10>
	<%
	 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>

<%
End Sub

'删除日志
Sub DelRecord()
 Dim I,ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 ID=Split(ID,",")
 For I=0 To Ubound(ID)
  KS.DeleteFile(conn.execute("select photourl from ks_enterprisezs where id=" & ID(I))(0))
  Conn.execute("Delete From KS_EnterpriseZS Where id="& id(I))
 Next 
 Response.Write "<script>alert('删除成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

'审核
Sub ShowNews()
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_EnterpriseZS where id=" &KS.ChkClng(KS.S("ProID")),conn,1,1
		If Not RS.Eof Then
		   Response.WRITE "<div style='padding:30px'><div><strong>证书名称：</strong>" & rs("Title") & "</div>"
		   Response.Write "<div style=""text-align:left""><strong>发证机构：</strong>" & RS("fzjg") & "</div>"
		   Response.Write "<div style=""text-align:left""><strong>生效日期：</strong>" & RS("sxrq") & "</div>"
		   Response.Write "<div style=""text-align:left""><strong>截止日期：</strong>" & RS("jzrq") & "</div>"
		   Dim PhotoUrl:PhotoUrl=RS("PhotoUrl")
		   If PhotoUrl<>"" And Not IsNull(PhotoURL) Then
		   Response.Write "<div style=""text-align:left"">证书照片：<img src='" & RS("photourl") & "'></div>"
		   End If
		   Response.Write "</div>"
		End If
		RS.Close:Set RS=Nothing
End Sub
'审核
Sub Verify
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_EnterpriseZS Set status=1 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'取消审核
Sub UnVerify
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_EnterpriseZS Set status=0 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

End Class
%> 
