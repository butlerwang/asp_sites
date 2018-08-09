<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdminiStratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Log_Main
KSCls.Kesion()
Set KSCls = Nothing

Class Log_Main
        Private KS,KSCls
		Private I
		Private totalPut
		Private CurrentPage
		Private SqlStr
        Private RSObj
        Private MaxPerPage
		Private Sub Class_Initialize()
		   MaxPerPage = 18
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
		If Request("page") <> "" Then
			  CurrentPage = CInt(Request("page"))
		Else
			  CurrentPage = 1
		End If
		With KS
		 .echo "<html>"
		 .echo "<head>"
		 .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		 .echo "<title>登录日志</title>"
		 .echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		 .echo "<script language=""JavaScript"">"
		 .echo "var Page='" & CurrentPage & "';"
		 .echo "</script>"
		 .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
		 .echo "<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>"
		
		If Not KS.ReturnPowerResult(0, "KMST10006") Then
		   .echo ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
		  Call KS.ReturnErr(1, "")
		End If
		 Select Case KS.G("Action")
		  Case "Del","DelAll"  Call LogDel()
		  Case Else  Call MainList()
		 End Select
		 End With
		End Sub
		
		Sub MainList()
		 With KS
		 %>
		<script language="javascript">
		
		function DelLog()
		{
		 var ids=get_Ids(document.myform);
		 if (ids!='')
		  {
		   if (confirm('真的要删除选中的日志吗,两天内的日志将不会被删除')){
		   $("#Action").val("Del");
		   $("#myform").submit();}
		  }
		 else
		  alert('请选择要删除的日志!');
		}
		function DelAllLog()
		{
		if (confirm('确定清空所有日志吗,两天内的日志将不会被清空')){
		   $("#Action").val("DelAll");
		   $("#myform").submit();}
		}
		function GetKeyDown()
		{ 
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 : location.reload(); break;
			 case 89 : DelAllLog();break;
			 case 68 : DelLog();break;
		   }	
		else	
		 if (event.keyCode==46)DelLog();
		}
		</script>
		<%
		 .echo "</head>"
		 .echo "<body topmargin=""0"" leftmargin=""0"" onclick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
		 .echo "<ul id='menu_top'>"
		 .echo "<li class='parent' onclick=""DelLog();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>删除选中日志</span></li>"
		 .echo "<li class='parent' onclick=""DelAllLog();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>全部清除</span></li>"
		 .echo "</ul>"

		 .echo "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		 .echo "<form name=""myform"" id=""myform"" action=""KS.Log.asp"" method=""post"">"
		 .echo "<input type='hidden' name='Action' value='Del' id='Action'>"
		 .echo "        <tr>"
		 .echo "          <td width=""35"" height=""25"" class=""sort""> <div align=""center"">选择</div></td>"
		 .echo "          <td height=""25"" class=""sort""> <div align=""center"">管理员</div></td>"
		 .echo "          <td width=""80"" class=""sort""><div align=""center"">操作结果</div></td>"
		 .echo "          <td align=""center"" class=""sort"">登录时间</td>"
		 .echo "          <td align=""center"" class=""sort"">登录IP</td>"
		 .echo "          <td align=""center"" class=""sort"">操作系统</td>"
		 .echo "          <td class=""sort""><div align=""center"">描 述</div></td>"
		 .echo "        </tr>"
		   
		   Set RSObj = Server.CreateObject("ADODB.RecordSet")
				 RSObj.Open "SELECT * FROM KS_Log order by LoginTime Desc", Conn, 1, 1
				 If Not RSObj.EOF Then
					totalPut = Conn.Execute("Select Count(ID) From KS_Log")(0)
		
							If CurrentPage < 1 Then	CurrentPage = 1
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
		
							If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RSObj.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
							Call showContent
			End If
		    RSObj.Close:Set RSObj=Nothing
			CloseConn
			.echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
	        .echo ("<tr><td width='180'><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> </div>")
	        .echo ("</td>")
	        .echo ("<td><input type='button' value='删 除' onclick=""DelLog();"" class='button'>&nbsp;<input type='button' value='清 空' onclick=""DelAllLog();"" class='button'></td>")
	        .echo ("</form><td align='right'>")
	         Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	        .echo ("</td></tr></table>")
		 .echo "</table>"
		 .echo "</body>"
		 .echo "</html>"
		 End With
		End Sub
		 Sub showContent()
		   Dim ID
		   With KS
				Do While Not RSObj.EOF
				   ID=RSObj("ID")
			       .echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & ID & "' onclick=""chk_iddiv('" &ID & "')"">"
			       .echo "<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" & ID & "')"" type='checkbox' id='c"& ID & "' value='" & ID & "'></td>"
				   .echo "  <td  height='20' class='splittd'>&nbsp; <span LogID='" & ID & "'><img src='Images/ico_friend.gif' width='16' height='16' align='absmiddle'>"
				  If RSObj("ResultTF") = 1 Then
				   .echo "  <span style='cursor:default;'>" & RSObj("UserName") & "</span></span></td>"
				   .echo "  <td class='splittd' align='center'>成功</td>"
				   .echo "  <td class='splittd' align='center'>" & RSObj("LoginTime") & "</td>"
				   .echo "  <td class='splittd' align='center'>" & RSObj("LoginIP") & "</td>"
				   .echo "  <td class='splittd' align='center'>" & RSObj("LoginOS") & "</td>"
				   .echo "  <td class='splittd' align='center'>" & RSObj("Description") & " </td>"
				  Else
				   .echo "    <span style='cursor:default;color:red'>" & RSObj("UserName") & "</span></span></td>"
				   .echo "  <td class='splittd' align='center'><font color=red>失败</font></td>"
				   .echo "  <td class='splittd' align='center'><FONT Color=red>" & RSObj("LoginTime") & "</font> </td>"
				   .echo "  <td class='splittd' align='center'><FONT Color=red>" & RSObj("LoginIP") & "</font> </td>"
				   .echo "  <td class='splittd' align='center'><FONT Color=red>" & RSObj("LoginOS") & "</font></td>"
				   .echo "  <td class='splittd' align='center'><font color=red>" & RSObj("Description") & "</font> </td>"
				  End If
				   .echo "</tr>"
				  I = I + 1
				  If I >= MaxPerPage Then Exit Do
					   RSObj.MoveNext
				Loop
			 End With
		End Sub
			 
		Sub LogDel()
			  Dim LogID,Action,Sql
			  
			 Action = KS.G("Action")   
			 If Action = "Del" Then
			     LogID = Trim(KS.G("ID"))
				 Sql = "Delete From KS_Log Where datediff(" & DataPart_D &",logintime," & SqlNowString & ")>2 And ID in(" & KS.FilterIDS(LogId) & ")"
				 Conn.Execute (Sql)
			ElseIf Action = "DelAll" Then
					Sql = "Delete From KS_Log Where datediff(" & DataPart_D &",logintime," & SqlNowString & ")>2"
					Conn.Execute (Sql)
			End If
			  KS.AlertHintScript ("恭喜,删除成功!")
		 End Sub
End Class
%> 
