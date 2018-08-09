<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_UserLog
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_UserLog
        Private KS,Action,Page,KSCls
		Private I, totalPut, CurrentPage,MaxPerPage, SqlStr,RS
		Private ID
		
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		  Set KSCls= New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
             With KS
		 	    .echo "<html>"
				.echo "<head>"
				.echo "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
				.echo "<title>用户使用记录管理</title>"
				.echo "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		        .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
		        .echo "<script language=""JavaScript"" src=""../KS_Inc/Jquery.js""></script>" & vbCrLf
		        .echo "<script language=""JavaScript"" src=""../KS_Inc/Kesion.Box.js""></script>" & vbCrLf
               Action=KS.G("Action")
				If Not KS.ReturnPowerResult(0, "KMUA10015") Then                 '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
				End iF

			 Page=KS.G("Page")
			 Select Case Action
			  Case "Del" ItemDelete
			  Case "DelAllRecord" DelAllRecord
			  Case Else MainList()
			 End Select
			.echo "</body>"
			.echo "</html>"
			End With
		End Sub
		
		Sub MainList()
			If Not IsEmpty(Request("page")) Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
		With KS
%>	   		
     <SCRIPT language=javascript>
		function DelDiggList()
		{
			var ids=get_Ids(document.myform);
			if (ids!='')
			 { 
				if (confirm('真的要删除选中的记录吗?'))
				{
				$("#myform").action="KS.UserUseLog.asp?Action=Del&show=<%=KS.G("show")%>&ID="+ids;
				$("#myform").submit();
				}
			}
			else 
			{
			 alert('请选择要删除的评论!');
			}
		}
		function DelDigg()
		{
			if (confirm('真的要删除选中的记录吗?'))
				{
				$("#myform").submit();
				}
		}
		function show(t,m,d)
		{
		new KesionPopup().PopupCenterIframe('查看详情[<font color=red>'+t+'</font>]记录','KS.UserUseLog.asp?action=list&infoid='+d,750,440,'auto')
		}
		function ShowCode(){
		new KesionPopup().PopupCenterIframe('查看Digg调用代码','KS.UserUseLog.asp?action=ShowCode',750,440,'no')
		}

		</SCRIPT>

	   <%
	
		.echo "</head>"
		
		.echo "<body topmargin='0' leftmargin='0'>"
		If KS.S("Action")="list" Then Call DiggDetail() : Exit Sub
		.echo "<div class='topdashed sort'>查看会员使用记录</div>"
		.echo "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		.echo(" <form name=""myform"" id=""myform"" method=""Post"" action=""KS.UserUseLog.asp?Action=Del"">")
		.echo "    <tr class='sort'>"
		.echo "    <td width='30' align='center'>选中</td>"
		.echo "    <td align='center'>标题</td>"
		.echo "    <td align='center'>时间</td>"
		.echo "    <td align='center'>用户名</td>"
		.echo "    <td align='center'>类型</td>"
		.echo "    <td align='center'>分类</td>"
		.echo "    <td align='center'>管理</td>"
		.echo "  </tr>"
		 Set RS = Server.CreateObject("ADODB.RecordSet")
		   Dim Param:Param=" where L.ChannelID=I.ChannelID"
		   
				  SqlStr="Select l.*,i.hits,i.tid,i.fname From KS_LogConsum l inner join ks_iteminfo i on l.infoid=i.infoid " & Param & " order by logid desc"
				  RS.Open SqlStr, conn, 1, 1
				 If RS.EOF And RS.BOF Then
				  .echo "<tr><td  class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"" colspan=16 height='25' align='center'>没有使用记录!</td></tr>"
				 Else
					        totalPut = Conn.Execute("Select count(logid) from KS_LogConsum l inner join ks_iteminfo i on l.infoid=i.infoid " & Param)(0)
							If CurrentPage < 1 Then CurrentPage = 1
							
		
							If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
							Call showContent
			End If
		  .echo "  </td>"
		  .echo "</tr>"

		 .echo "</table>"
		 .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
		 .echo ("<tr><td width='170'><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> </div>")
		 .echo ("</td>")
	     .echo ("<td><input type=""button"" value=""删除选中的记录"" onclick=""DelDiggList();"" class=""button""></td>")
	     .echo ("</form><td align='right'>")
		 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	     .echo ("</td></tr></form></table>")
		 .echo ("<form action='KS.UserUseLog.asp?action=DelAllRecord' method='post' target='_hiddenframe'>")
		 .echo ("<iframe src='about:blank' style='display:none' name='_hiddenframe' id='_hiddenframe'></iframe>")
		 .echo ("<div class='attention'><strong>特别提醒： </strong><br>当站点运行一段时间后,网站的会员 使用记录表可能存放着大量的记录,为使系统的运行性能更佳,建议一段时间后清理一次。")
		 .echo ("<br /> <strong>删除范围：</strong><input name=""deltype"" type=""radio"" value=1>10天前 <input name=""deltype"" type=""radio"" value=""2"" /> 1个月前 <input name=""deltype"" type=""radio"" value=""3"" />2个月前 <input name=""deltype"" type=""radio"" value=""4"" />3个月前 <input name=""deltype"" type=""radio"" value=""5"" /> 6个月前 <input name=""deltype"" type=""radio"" value=""6"" checked=""checked"" /> 1年前  <input onclick=""$(parent.frames['FrameTop'].document).find('#ajaxmsg').toggle();"" type=""submit""  class=""button"" value=""执行删除"">")
		 .echo ("</div>")
		 .echo ("</form>")
		End With
		End Sub
		Sub showContent()
		  Dim LogID
		  With KS
			 Do While Not RS.EOF
			 LogID=RS("LogID")
			.echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & LogID & "' onclick=""chk_iddiv('" & LogID & "')"">"
		   .echo "<td class='splittd'><input name='id' onclick=""chk_iddiv('" & LogID & "')"" type='checkbox' id='c"& LogID & "' value='" & LogID & "'></td>"
		  .echo " <td class='splittd' height='22'><span style='cursor:default;'>"
		   .echo  RS("title")  & "</td>"
		   .echo " <td class='splittd' align='center'>" & RS("adddate") & " </td>"
		   .echo " <td class='splittd' align='center'>" & RS("username") & " </td>"
		   .echo " <td class='splittd' align='center'>" 
		   Select Case rs("basictype")
	      Case 1:Response.WRite "文章"
		  Case 2:Response.Write "图片"
		  Case 3:Response.Write "下载"
		  Case 4:Response.Write "动漫"
		  Case 7:Response.Write "影片"
		  Case 9:Response.Write "试卷"
		 End Select
		   .echo "</td>"
		   .echo " <td class='splittd' align='center'>" & KS.C_C(RS("Tid"),1) & " </td>"
		   .echo " <td class='splittd' align='center'><a href='?action=Del&id=" & logid & "' onclick=""return(confirm('确定删除吗?'))"">删除</a> </td>"
		   .echo "</tr>"
			I = I + 1:	If I >= MaxPerPage Then Exit Do
			RS.MoveNext
			Loop
		  RS.Close
		  End With
		 End Sub
		 
		 
		 Sub ItemDelete()
			Dim ID:ID = KS.G("ID")
			If ID="" Then KS.AlertHintScript "您没有选择要删除的记录!"
			conn.Execute ("Delete From KS_LogConsum Where LogID IN(" & KS.FilterIds(ID) & ")")
		    response.redirect request.servervariables("http_referer") 
		 End Sub
		 
		
		 Sub DelAllRecord()
		  Dim Param
		  Select Case KS.ChkClng(KS.G("DelType"))
		   Case 1 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>11"
		   Case 2 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>31"
		   Case 3 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>61"
		   Case 4 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>91"
		   Case 5 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>181"
		   Case 6 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>366"
		  End Select
   		  If Param<>"" Then Conn.Execute("Delete From KS_LogConsum Where " & Param)
          KS.echo "<script>$(parent.document).find('#ajaxmsg').toggle();alert('恭喜,删除指定日期内的记录成功!');</script>"
		 End Sub
End Class
%> 
