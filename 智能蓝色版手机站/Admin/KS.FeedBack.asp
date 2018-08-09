<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New FeedBack
KSCls.Kesion()
Set KSCls = Nothing

Class FeedBack
        Private KS,Action,Page,KSCls
		Private I, totalPut, CurrentPage,MaxPerPage, SqlStr,RS,ID
		Private Sub Class_Initialize()
		  MaxPerPage =20
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
             With KS
		 	    .echo"<html>"
				.echo"<head>"
				.echo"<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
				.echo"<title>投诉建议管理</title>"
				.echo"<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
				.echo"<script src='../KS_Inc/common.js'></script>"
             Action=KS.G("Action")
			If Not KS.ReturnPowerResult(0, "KSMS20000") Then                  '栏目权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
			End iF

			 Page=KS.G("Page")
			If Not IsEmpty(Request("page")) Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
			 
			 Select Case Action
			  Case "Del" Call Del()
			  Case "ShowDetail" Call ShowDetail()
			  Case "DoSave" Call DoSave()
			  Case Else
			   Call MainList()
			 End Select
			.echo"</body>"
			.echo"</html>"
			End With
		End Sub
		
		Sub MainList()
		With KS
		 .echo"<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
		 .echo"<script language=""JavaScript"" src=""../KS_Inc/jQuery.js""></script>" & vbCrLf
       %>
	   		<SCRIPT language=javascript>
			
		function DelCompany()
		{
			var ids=get_Ids(document.myform);
			if (ids!='')
			 { 
				if (confirm('真的要删除选中的记录吗?'))
				{
				$("form[name=myform]").action="KS.JobTraining.asp?Action=Del&ID="+ids;
				$("form[name=myform]").submit();
				}
			}
			else 
			{
			 alert('请选择要删除的个人简历!');
			}
		}
	
		</SCRIPT>

	   <%
	
		.echo"</head>"
		
		.echo"<body scroll=no topmargin='0' leftmargin='0'>"
		.echo"<ul id='mt'> <div id='mtl'>投诉建议管理:</div><li>"
		.echo"</ul>"
		.echo"<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		.echo"    <tr class='sort'>"
		.echo"    <td width='30' align='center'>选中</td>"
		.echo"    <td  align='center'>编号</td>"
		.echo"    <td align='center'>投诉主题</td>"
		.echo"    <td width='90' align='center'>投诉时间</td>"
		.echo"    <td width='80' align='center'>投诉对象</td>"
		.echo"    <td width='60' align='center'>受理人</td>"
		.echo"    <td align='center'>受理时间</td>"
		.echo"    <td width='70' align='center'>状态</td>"
		.echo"    <td align='center'>管理操作</td>"
		.echo"  </tr>"
		 Set RS = Server.CreateObject("ADODB.RecordSet")
		 SqlStr = "SELECT * FROM [KS_FeedBack] Order By ID Desc"
		 RS.Open SqlStr, conn, 1, 1
		   If RS.EOF And RS.BOF Then
			 .echo"<tr><td  class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"" colspan=10 height='25' align='center'>没有人反馈!</td></tr>"
		   Else
					totalPut = RS.RecordCount
		
							If CurrentPage < 1 Then	CurrentPage = 1	
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
		
							If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
								RS.Move (CurrentPage - 1) * MaxPerPage
							Else 
							   CurrentPage=1
							End If
							Call showContent(RS)
			End If
		  .echo"  </td>"
		  .echo"</tr>"

		 .echo"</table>"
		  .echo("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
		 .echo("<tr><td width='180'><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> </div>")
		 .echo("</td>")
		 .echo("<td><input type='submit' class='button' value=' 删 除 ' onclick='return(confirm(""确定删除选中的意见吗?""))'></td>")
	     .echo("</form><td align='right'>")
		    Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	     .echo("</td></tr></form></table>")
		End With
		End Sub
		Sub showContent(RS)
		  With KS
          .echo (" <form name=""myform"" method=""Post"" action=""KS.FeedBack.asp"">")
		  .echo("<input type='hidden' name='action' id='action' value='Del'>")
			 Do While Not RS.EOF
		  .echo"<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & RS("ID") & "' onclick=""chk_iddiv('" & RS("ID") & "')"">"
		  .echo"<td class='splittd' align='center'><input name='id' onclick=""chk_iddiv('" &RS("ID") & "')"" type='checkbox' id='c"& RS("ID") & "' value='" &RS("ID") & "'></td>"
		  
		  dim bh:bh=rs("id")
		  IF LEN(BH)=1 THEN 
			 BH="00"& bh
		  ElseIf LEN(BH)=2 Then
			 Bh="0" & bh
		 End If
		  bh="YJ" & year(rs("adddate")) & month(rs("adddate")) & bh
						  
		  .echo" <td height='22' class='splittd' align='center'>"
		   .echo bh
		   
		   .echo"</td>"
		   .echo" <td class='splittd' align='center'><a href='?action=ShowDetail&id=" & rs("id") & "'>" & rs("title") & "</a></td>"
		   .echo" <td class='splittd' align='center'>" & formatdatetime(rs("adddate"),2) & "</td>"
		   .echo" <td class='splittd' align='center'>&nbsp;" & KS.Gottopic(RS("object"),24) & "</a></td>"
		   .echo" <td class='splittd' align='center'>"
		   
		   Dim AcceptTime,Delstr,strs
		  if rs("Accepted")="" or isnull(rs("accepted")) then
		   .echo "未处理"
		   AcceptTime="---"
		   Delstr="<a onclick=""return(confirm('确定删除吗?'))"" href='?action=del&id=" & rs("id") & "'>删除</a>"
		   strs="<font color=red>待受理</font>"
		  else
		   .echo rs("Accepted")
		   AcceptTime=RS("AcceptTime")
		   Delstr="<a href='#' disabled>删除</a>"
		   strs="<font color=green>已受理</font>"
		  end if
		   .echo"</td>"
		   .echo" <td class='splittd' align='center'>" & accepttime & "</td>"
		   .echo" <td align='center' class='splittd'>"
		   
		   .echo strs
		   
		   .echo" </td>"
		   

		   .echo" <td class='splittd' align='center'><a href='?action=ShowDetail&id=" & rs("id") & "'>查看受理</a> "
		   .echo"<a href='?action=Del&id=" & rs("id") & "' onclick='return(confirm(""确定删除该反馈吗?""))'>删除</a> "
		  
		   
		   
		   .echo"</td></tr>"
							  I = I + 1
								If I >= MaxPerPage Then Exit Do
							   RS.MoveNext
							   Loop
								RS.Close
								 
		  End With
		 End Sub
		 
		 Sub ShowDetail()
		  Dim ID:ID=KS.ChkClng(KS.G("ID"))
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select * from ks_feedback where id=" & ID,Conn,1,1
		  If RS.EoF Then 
		   rs.close:set rs=nothing
		   KS.Echo "<script>alert('出错!');history.back();</script>"
		   response.end
		  End If
		  dim accepted,accepttime,acceptresult
		  accepted=rs("accepted")
		  if accepted="" or isnull(accepted) then 
		  accepted=ks.c("adminname")
		  accepttime=now
		  else
		  accepttime=rs("accepttime")
		  end if
		  acceptresult=rs("acceptresult")
		  %>
		  
		  <div class='topdashed sort'>查看受理意见及投诉</div><br>
		  
		  <table cellspacing=0 cellpadding=0 width=100% align=center  border=0>
                    <tr>
                      <td class=title height=21 align="center"><font style="MARGIN-TOP: 2px; MARGIN-LEFT: 10px" color=#ffffff><strong>会员意见</strong></font></td>
                    </tr>
          </table>
            <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
                  <tr class="tdbg">
                    <td width="17%" height="25" align="right" class="clefttitle"></strong> 意见主题：</td>
                    <td><%=rs("title")%></td>
                    <td width="17%" height="25" align="right" class="clefttitle">意见对象：</td>
                    <td height="25" class="tdbg">
					<%=rs("object")%>					</td>
                  </tr>
                  <tr class="tdbg">
                    <td width="17%" height="25" align="right" class="clefttitle">意见内容：</td>
                    <td height="25" class="tdbg">
					<%=rs("content")%>					</td>
                    <td height="25" align="right" class="tdbg">投 诉 人：</td>
                    <td height="25" class="tdbg"><%=rs("username")%></td>
                  </tr>
				  <%if rs("hopesolution")<>"" and not isnull(rs("hopesolution")) then%>
                  <tr class="tdbg">
                    <td width="17%" height="25" align="right" class="clefttitle">希望处理方案：</td>
                    <td height="25" class="tdbg" colspan=3>
					<%=rs("hopesolution")%>					</td>
                  </tr>
				  <%end if%>
               </table>		  
		  
		  <table cellspacing=0 cellpadding=0 width=100% align=center  border=0>
                    <tr>
                      <td class=title height=21 align="center"><font style="MARGIN-TOP: 2px; MARGIN-LEFT: 10px" color=#ffffff><strong>处理结果</strong></font></td>
                    </tr>
          </table>
            <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
			  <form name="myform" action="?action=DoSave" method="post">
			  <input type="hidden" value="<%=id%>" name="id">
                  <tr class="tdbg">
                    <td width="17%" height="25" align="right" class="clefttitle"></strong> 受理人：</td>
                    <td><input type="text" value="<%=accepted%>" name="accepted" class="textbox"></td>
                    <td width="17%" height="25" align="right" class="clefttitle">受理时间：</td>
                    <td height="25" class="tdbg">
					<input type="text" name="accepttime" value="<%=accepttime%>" calss="textbox">
					</td>
                  </tr>
                  <tr class="tdbg">
                    <td width="17%" height="25" align="right" class="clefttitle">受理结果：</td>
                    <td height="25" class="tdbg" colspan=3>
					<textarea name="acceptresult" style="width:90%;height:150px"><%=acceptresult%></textarea>
					</td>
                  </tr>
				  <tr class="tdbg">
                    <td width="17%" height="25" colspan=4 style="text-align:center">
					<input type="submit" value="保存受理结果" class="button">
					<input type="button" value=" 返 回 " onclick="history.back()" class="button">
					</td>
                  </tr>
			  </form>
</table>
	  

		  <br>
		  
		  <%RS.Close:Set RS=Nothing
		  
		
		  
		  
		  
		 End Sub
		 
		 
		
		 
		 Sub DoSave()
             Dim ID:ID=KS.ChkClng(KS.G("ID"))
			 if not isdate(KS.G("accepttime")) then
			  KS.echo "<script>alert('受理时间格式不正确!');history.back();</script>"
			  response.end
			 end if
			 if KS.G("AcceptResult")="" Then
			  KS.echo "<script>alert('受理结果不能为空!');history.back();</script>"
			  response.end
			 END IF
			 
			  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select * From KS_FeedBack Where ID=" & ID,conn,1,3
			 If RS.Eof And RS.Bof Then
			  KS.echo "error!"
			  response.end
			 End If
			  RS("Accepted") = KS.G("Accepted")
			  RS("AcceptResult")  = KS.G("AcceptResult")
			  RS("AcceptTime")=KS.G("AcceptTime")
			  RS.Update
			  RS.Close:Set RS=Nothing
			  
			  KS.echo "<script>alert('受理结果已保存！');location.href='KS.FeedBack.asp';</script>"

		 End Sub


		 
		 Sub Del()
			Dim ID:ID = KS.FilterIDS(KS.G("ID"))
			if id<>"" then
	         conn.execute("delete from ks_feedback where id in(" & ID  & ")")
		    end if
		    response.redirect request.servervariables("http_referer") 
		 End Sub
		 
		
End Class
%> 
