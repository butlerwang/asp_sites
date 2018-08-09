<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Main
KSCls.Kesion()
Set KSCls = Nothing

Class Main
        Private KS,Action,PKID
		Private I, totalPut, CurrentPage, SqlStr, RSObj
        Private MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub


		Public Sub Kesion()
			If Not KS.ReturnPowerResult(0, "KSMS20014") Then
			  Call KS.ReturnErr(1, "")
			  exit sub
			End If
			PKID=KS.ChkClng(Request("PKID"))
			Action=KS.G("Action")
			Select Case Action
			 Case "verify"
			      Call verify()
			 Case "del"
			      Call del()
			 Case Else
			   Call MainList()
			End Select
	    End Sub
		
		Sub MainList()
			If Request("page") <> "" Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
			With Response
			.Write "<html>"
			.Write "<head>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script language=""JavaScript"" src=""../ks_inc/Common.js""></script>"
			.Write "<script language=""JavaScript"" src=""../ks_inc/jquery.js""></script>"
			%>

			<%
			.Write "</head>"
			.Write "<body topmargin=""0"" leftmargin=""0"">"
			  .Write "<ul id='menu_top' style='font-weight:bold;text-align:center;padding-top:14px'>"
			  .Write "网友PK观点管理"
			  .Write "</ul>"
			

			.Write "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
			%>
			<form name='myform' method='Post' action='KS.PKGD.asp'>
		    <input type="hidden" value="del" name="action" id="action">
		    <input type="hidden" value="1" name="v">

			<%
			.Write "  <tr>"			
			.Write "          <td width=""40"" height=""25"" class=""sort"" align=""center"">选择</td>"
			.Write "          <td height=""25"" class=""sort"" align=""center"">观点内容</td>"
			.Write "          <td class=""sort"" align=""center"">PK主题</td>"
			.Write "          <td align=""center"" class=""sort"">用户</td>"
			.Write "          <td align=""center"" class=""sort"">时间</td>"
			.Write "          <td align=""center"" class=""sort"">观点</td>"
			.Write "          <td align=""center"" class=""sort"">状态</td>"
			.Write "          <td align=""center"" class=""sort"">管理操作</td>"
			.Write "  </tr>"
			 
			  dim param
			 if PKID<>0 then
			   param=" where a.pkid=" & PKID
			 end if
			 
			 Set RSObj = Server.CreateObject("ADODB.RecordSet")
					   SqlStr = "SELECT a.*,b.title FROM KS_PKGD a inner join KS_PKZT b on a.pkid=b.id" &param&" order by a.ID DESC"
					   RSObj.Open SqlStr, Conn, 1, 1
					 If RSObj.EOF And RSObj.BOF Then
					 Else
						totalPut = RSObj.RecordCount
			
								If CurrentPage < 1 Then CurrentPage = 1
								   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
									Else
										CurrentPage = 1
									End If
										Call showContent
				End If
				
			.Write "    </td>"
			.Write "  </tr>"
			.Write "</table>"
			.Write "</body>"
			.Write "</html>"
			End With
			End Sub
			Sub showContent()
			   on error resume next
			  With Response
					Do While Not RSObj.EOF
					  .Write "<tr align=""center"" onMouseOut=""this.className='list'"" onMouseOver=""this.className='listmouseover'"" id='u" & RSobj("ID") & "' onClick=""chk_iddiv('" & rsobj("ID") & "')""> "
					  .Write "<td align='center' width=""40"" height=""25"" class=""splittd""><input name=""id"" onClick=""chk_iddiv('" & rsobj("id") & "')"" type='checkbox' id='c" & rsobj("id") & "' value='" & rsobj("id") & "'></td>"
					  .Write "  <td align='left' class='splittd' height='20'>&nbsp;"
					  .Write "    <span style='cursor:default;' title='" & rsobj("content") & "'>" & KS.GotTopic(RSObj("content"), 45) & "</span> </td>"
					  .Write "  <td class='splittd' align='center'><a href='../plus/pk/pk.asp?id=" & rsobj("pkid") & "' target='_blank'>" & ks.gottopic(rsobj("title"),20) & "</a></td>"
					  .Write "  <td class='splittd' align='center'>" 
					  .write rsobj("username")
					  .Write " </td>"
					  .Write "  <td class='splittd' align='center'>" & rsobj("adddate") & "</td>"
					  .Write "  <td class='splittd' align='center'>"
					   if rsobj("role")="1" then
					    .write "<font color=blue>正方</font>"
					   elseif rsobj("role")="2" then
					    .write "<font color=green>反方</font>"
					   else
					    .write "<font color=red>第三方</font>"
					   end if
					  .Write "</td>"
					  .Write "  <td class='splittd' align='center'>"
					   if rsobj("status")=1 then
					    .write "<Font color=green>已审核</font>"
					   else
					    .write "<Font color=red>未审核</font>"
					   end if
					  .Write "</td>"
					  .Write "  <td class='splittd' align='center'>"
					  if rsobj("status")="1" then
					  .Write "<a href='?action=verify&v=0&id=" & rsobj("id") &"' title='取消审核'>取审</a>"
					  else
					  .Write "<a href='?action=verify&v=1&id=" & rsobj("id") &"' title='审核意见'>审核</a>"
					  end if
					  .Write" <a href='?action=del&id=" & rsobj("id") & "' onclick=""return(confirm('确定删除吗?'))"">删除</a></td>"
					  .Write "</tr>"
					 I = I + 1
					  If I >= MaxPerPage Then Exit Do
						   RSObj.MoveNext
					Loop
					  RSObj.Close
					  
					  %>
						  <tr>
						   <td colspan=6>
						   <div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a>
						   <input type="submit" class="button" value="删除选中" onClick="return(confirm('此操作不可逆,确定删除吗?'))">
						   <input type="submit" class="button" value="批量审核" onClick="$('#action').val('verify')">
							</div>
						   </td>
									</form>  
				 <td colspan=5>
					  
					  </td>
					  </tr>
				</table>
				<%

					  .Write "<tr><td height='26' colspan='8' align='right'>"
					 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
				End With
			End Sub
			
		  
		  '删除
		  Sub del()
		  		 Dim K, ZJID
				 ZJID = Trim(KS.G("ID"))
				 if zjid="" then
				   ks.alerthintscript "请选择要删除的意见!"
				 end if
				 ZJID = Split(ZJID, ",")
				 For k = LBound(ZJID) To UBound(ZJID)
					Conn.Execute ("Delete From KS_PKGD Where ID =" & ZJID(k))
				 Next
				 KS.AlertHintScript "恭喜,删除成功!"
		  End Sub
		  
		  sub verify()
		    dim id
			id=request("id")
			if id="" then
				   ks.alerthintscript "请选择要审核的意见!"
			end if
			conn.execute("update KS_PKGD set status=" & ks.chkclng(request("v")) & " where id in(" & ks.filterids(id) & ")")
			KS.AlertHintScript "恭喜,操作成功!"
		  end sub
	

End Class
%>
 
