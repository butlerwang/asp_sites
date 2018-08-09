<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.FsoVarCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_UserLog
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_UserLog
        Private KS,Action,KSCls
		Private I, totalPut, MaxPerPage, SqlStr,RS,ID
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
  %>
  <!--#include file="../ks_cls/ubbfunction.asp"-->
  <%

		Public Sub Kesion()
             With KS
		 	    .echo "<html>"
				.echo "<head>"
				.echo "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
				.echo "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		        .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
		        .echo "<script language=""JavaScript"" src=""../KS_Inc/Jquery.js""></script>" & vbCrLf
               Action=KS.G("Action")
				If Not KS.ReturnPowerResult(0, "KSMS20016") Then                 '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
				End iF
%>	   		
     <SCRIPT language=javascript>
		function DelDiggList(){
			var ids=get_Ids(document.myform);
			if (ids!=''){ 
				if (confirm('真的要删除选中的记录吗?')){
				$("#myform").action="KS.UserLog.asp?Action=Del&show=<%=KS.G("show")%>&ID="+ids;
				$("#myform").submit();}
			}else { alert('请选择要删除的记录!');}
		}
		function DelDigg(){if (confirm('真的要删除选中的记录吗?')){$("#myform").submit();}	}
		</SCRIPT>
		<style type="text/css">
			.imglist{margin-top:10px;height:70px;}
				 .imglist ul{}
				 .imglist ul li{float:left;width:70px;padding:margin:10px;}
				 .imglist ul li img{width:60px;height:60px;border:1px solid #ccc;padding:1px;}
				 .intropic{margin:10px;}
				 .intro{margin:10px;color:#999;}
				 .intropic img{width:120px;border:1px solid #ccc;padding:2px}
				 a.logtitle{font-size:14px;}
		</style>

	   <%
		.echo "</head>"
		
		.echo "<body topmargin='0' leftmargin='0'>"
		.echo "<div class='topdashed sort' ><a href='ks.userlog.asp'>微博数据管理</a> | <a href='?action=comment'>微博评论数据</a></div>"
			 Select Case Action
			  Case "comment" comment
			  Case "DelComment" DelComment
			  Case "DelAllComment" DelAllComment
			  Case "Del" ItemDelete
			  Case "DelAllRecord" DelAllRecord
			  Case Else MainList()
			 End Select
			.echo "</body>"
			.echo "</html>"
			End With
		End Sub
		
		Sub MainList()
         With KS 
		.echo "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		.echo(" <form name=""myform"" id=""myform"" method=""Post"" action=""KS.UserLog.asp?Action=Del"">")
		.echo "    <tr class='sort'>"
		.echo "    <td width='40' nowrap align='center'>选中</td>"
		.echo "    <td style='text-align:left'>微博内容</td>"
		.echo "  </tr>"
		   Dim Param:Param=""
		   If KS.G("Key")<>"" Then Param=" where a.UserName='" & KS.S("Key") & "'"
			SQLStr=" select b.id,a.userid,a.username,a.transtime,a.msg,b.adddate,b.copyfrom,b.note,b.cmtnum,b.username as busername,b.userid as buserid,b.transnum,a.type,a.id as rid from ks_userlogr a left join ks_userlog b on a.msgid=b.id " & param & " order by a.id desc"
			Set RS = Server.CreateObject("AdoDb.RecordSet")
			RS.Open SQLStr, conn, 1, 1
			If RS.EOF And RS.BOF Then
				  .echo "<tr><td  class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"" colspan=6 height='25' align='center'>没有微博数据!</td></tr>"
				 Else
					totalPut = Conn.Execute("select count(1) from ks_userlogr a left join ks_userlog b on a.msgid=b.id "  & Param)(0)
					If CurrentPage >1 and  (CurrentPage - 1) * MaxPerPage < totalPut Then RS.Move (CurrentPage - 1) * MaxPerPage
					Dim i:I=0
					Do While Not RS.EOF
			.echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & RS("rID") & "' onclick=""chk_iddiv('" & RS("rID") & "')"">"
		   .echo "<td class='splittd' style='text-align:center'><input name='id' onclick=""chk_iddiv('" &RS("rID") & "')"" type='checkbox' id='c"& RS("rID") & "' value='" &RS("rID") & "'></td>"
		  .echo " <td class='splittd' height='22' style='word-break:break-all;padding-top:10px;'><span style='cursor:default;'>"
		   .echo  "<a href='../user/weibo.asp?userid=" & rs("userid") & "' target='_blank'>" & RS("username")  & "</a>&nbsp;&nbsp;<span style='color:#999'>-" & RS("adddate") & " - " & rs("copyfrom") & "</span> <a href='?action=Del&id=" & rs("rid") & "' onclick=""return(confirm('确定删除吗?'))"">删除</a> <div style='color:#888;margin:10px'>" & ubbcode(Replace(RS("note")&"","{$GetSiteUrl}",KS.GetDomain),i) & "</div></td>"
		   .echo "</tr>"
			I = I + 1:	If I >= MaxPerPage Then Exit Do
			RS.MoveNext
			Loop
		  RS.Close
			End If
		  .echo "  </td>"
		  .echo "</tr>"

		 .echo "</table>"
		 .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
		 .echo ("<tr><td width='170'><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> </div>")
		 .echo ("</td>")
	     .echo ("<td><input type=""button"" value=""删除选中的记录"" onclick=""DelDiggList();"" class=""button""></td>")
	     .echo ("</form></td><td><form name='sform' action='?' method='post'><strong>按用户名搜索：</strong><input class='textbox' type='text' name='key'> <input class='button' type='submit' value='搜索'/></form></td></tr></table>")
	      Call KS.ShowPage(totalput, MaxPerPage, "",CurrentPage,true,true)
		 .echo ("<br /> <br /> <br /> <form action='KS.UserLog.asp?action=DelAllRecord' method='post' target='_hiddenframe'>")
		 .echo ("<iframe src='about:blank' style='display:none' name='_hiddenframe' id='_hiddenframe'></iframe>")
		 .echo ("<div class='attention'><strong>特别提醒： </strong><br>当站点运行一段时间后,网站的微博记录表可能存放着大量的记录,为使系统的运行性能更佳,建议一段时间后清理一次。")
		 .echo ("<br /> <strong>删除范围：</strong><input name=""deltype"" type=""radio"" value=1>10天前 <input name=""deltype"" type=""radio"" value=""2"" /> 1个月前 <input name=""deltype"" type=""radio"" value=""3"" />2个月前 <input name=""deltype"" type=""radio"" value=""4"" />3个月前 <input name=""deltype"" type=""radio"" value=""5"" /> 6个月前 <input name=""deltype"" type=""radio"" value=""6"" checked=""checked"" /> 1年前  <input  type=""submit""  class=""button"" value=""执行删除"">")
		 .echo ("</div>")
		 .echo ("</form>")
		End With
		End Sub

		 Sub Comment()
			 With KS  
			.echo "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
			.echo(" <form name=""myform"" id=""myform"" method=""Post"" action=""KS.UserLog.asp?Action=DelComment"">")
			.echo "    <tr class='sort'>"
			.echo "    <td width='40' nowrap align='center'>选中</td>"
			.echo "    <td style='text-align:left'>评论内容</td>"
			.echo "    <td>评论人</td>"
			.echo "    <td>评论时间</td>"
			.echo "    <td style='text-align:left'>操作</td>"
			.echo "  </tr>"
			   Dim Param:Param=""
			   If KS.G("Key")<>"" Then Param=" where UserName='" & KS.S("Key") & "'"
				SQLStr=" select * from KS_UserLogCMT " & param & " order by id desc"
				Set RS = Server.CreateObject("AdoDb.RecordSet")
				RS.Open SQLStr, conn, 1, 1
				If RS.EOF And RS.BOF Then
					  .echo "<tr><td  class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"" colspan=6 height='25' align='center'>没有评论数据!</td></tr>"
					 Else
					    TotalPut=Conn.Execute("select count(1) From  KS_UserLogCMT " & param)(0)
						If CurrentPage >1 and  (CurrentPage - 1) * MaxPerPage < totalPut Then RS.Move (CurrentPage - 1) * MaxPerPage
					    Dim i:I=0
						Do While Not RS.EOF
						.echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & RS("ID") & "' onclick=""chk_iddiv('" & RS("ID") & "')"">"
					   .echo "<td class='splittd' style='text-align:center'><input name='id' onclick=""chk_iddiv('" &RS("ID") & "')"" type='checkbox' id='c"& RS("ID") & "' value='" &RS("ID") & "'></td>"
					  .echo " <td class='splittd' height='22' style='word-break:break-all;padding-top:10px;'>" & ubbcode(Replace(RS("content"),"{$GetSiteUrl}",KS.GetDomain),i) & "</td><td class='splittd'><span style='cursor:default;'>"
					   .echo  "<a href='../user/weibo.asp?userid=" & rs("userid") & "' target='_blank'>" & RS("username")  & "</a><td class='splittd'><span style='color:#999'>" & RS("adddate") & "</span></td><td class='splittd'> <a href='?action=DelComment&id=" & rs("id") & "' onclick=""return(confirm('确定删除吗?'))"">删除</a> </td>"
					   .echo "</tr>"
						I = I + 1:	If I >= MaxPerPage Then Exit Do
						RS.MoveNext
						Loop
					  RS.Close
				End If
			  .echo "  </td>"
			  .echo "</tr>"
	
			 .echo "</table>"
			 .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
			 .echo ("<tr><td width='170'><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> </div>")
			 .echo ("</td>")
			 .echo ("<td><input type=""button"" value=""删除选中的记录"" onclick=""DelDiggList();"" class=""button""></td>")
			 .echo ("</form></td><td><form name='sform' action='?Action=comment' method='post'><strong>按用户名搜索：</strong><input class='textbox' type='text' name='key'> <input class='button' type='submit' value='搜索'/></form></td></tr></table>")
			  Call KS.ShowPage(totalput, MaxPerPage, "",CurrentPage,true,true)
			 .echo ("<br /> <br /> <br /> <form action='KS.UserLog.asp?action=DelAllComment' method='post' target='_hiddenframe'>")
			 .echo ("<iframe src='about:blank' style='display:none' name='_hiddenframe' id='_hiddenframe'></iframe>")
			 .echo ("<div class='attention'><strong>特别提醒： </strong><br>当站点运行一段时间后,网站的微博评论记录表可能存放着大量的记录,为使系统的运行性能更佳,建议一段时间后清理一次。")
			 .echo ("<br /> <strong>删除范围：</strong><input name=""deltype"" type=""radio"" value=1>10天前 <input name=""deltype"" type=""radio"" value=""2"" /> 1个月前 <input name=""deltype"" type=""radio"" value=""3"" />2个月前 <input name=""deltype"" type=""radio"" value=""4"" />3个月前 <input name=""deltype"" type=""radio"" value=""5"" /> 6个月前 <input name=""deltype"" type=""radio"" value=""6"" checked=""checked"" /> 1年前  <input  type=""submit""  class=""button"" value=""执行删除"">")
			 .echo ("</div>")
			 .echo ("</form>")
			End With
		 End Sub
		 
		 Sub DelComment()
			Dim I,ID:ID =KS.FilterIds(KS.S("ID"))
			If ID="" Then KS.AlertHintScript "您没有选择要删除的记录!"
			ID=Split(ID,",")
			Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			For I=0 To Ubound(ID)
			  rs.open "select top 1 * From KS_UserLogCMT Where id=" & ID(I),CONN,1,1
			  If Not RS.Eof Then
			    Conn.Execute("Update KS_Userlog Set CmtNum=CmtNum-1  Where CmtNum>1 and id=" & rs("msgid"))
			  End If
			  rs.close
			Next
			Set RS=Nothing
			Conn.Execute("delete From KS_UserLogCMT Where ID in (" & KS.FilterIds(KS.S("ID")) & ")")
			response.redirect request.servervariables("http_referer") 
		 End Sub
		 Sub DelAllComment()
		  Dim Param
		  Select Case KS.ChkClng(KS.G("DelType"))
		   Case 1 Param="datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>11"
		   Case 2 Param="datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>31"
		   Case 3 Param="datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>61"
		   Case 4 Param="datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>91"
		   Case 5 Param="datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>181"
		   Case 6 Param="datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>366"
		  End Select
		   Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			  rs.open "select top 1 * From KS_UserLogCMT Where " & Param,CONN,1,1
			   do while not rs.eof
			    Conn.Execute("Update KS_Userlog Set CmtNum=CmtNum-1  Where CmtNum>1 and id=" & rs("msgid"))
				rs.movenext
			   loop
			  rs.close
			  Set RS=Nothing
			Conn.Execute("delete From KS_UserLogCMT Where " & param)
			KS.echo "<script>alert('恭喜,删除指定日期内的记录成功!');</script>"
		 End Sub
		 
		 Sub ItemDelete()
			Dim I,ID:ID =KS.FilterIds(KS.S("ID"))
			If ID="" Then KS.AlertHintScript "您没有选择要删除的记录!"
			ID=Split(ID,",")
			For I=0 To Ubound(ID)
			 Call DelTalk(ID(I))
			Next
		    response.redirect request.servervariables("http_referer") 
		 End Sub
		 
	Sub DelTalk(id)
	  Dim RS:Set RS=Server.CreateObject("adodb.recordset")
	  RS.Open "select top 1 * From KS_UserLogR Where ID=" & id,conn,1,1
	  If RS.Eof AND RS.Bof Then
	    RS.Close:Set RS=Nothing
	    Exit Sub
	  End If
	  Dim bType:bType=RS("Type")
	  Dim UserName:UserName=RS("UserName")
	  Dim MsgId:MsgId=RS("MsgId")
	  RS.Close:Set RS=Nothing

	  If BType=0 Then
	    Conn.Execute("Delete From KS_UserLog Where ID=" & MsgId)
		Conn.Execute("Delete From KS_UserLogCMT Where MsgID=" & MsgId)
	  Else
	    Conn.Execute("Update KS_UserLog set TransNum=TransNum-1  Where id=" & MsgId &" and TransNum>=1")
	  End If
	    Conn.Execute("Delete From KS_UserLogR Where ID=" & id)
	    Conn.Execute("Update KS_User set MsgNum=MsgNum-1  Where UserName='" & UserName &"' and MsgNum>=1")
	End Sub
		 
		 
		
		 Sub DelAllRecord()
		  Dim Param
		  Select Case KS.ChkClng(KS.G("DelType"))
		   Case 1 Param="datediff(" & DataPart_D & ",transtime," & SqlNowString & ")>11"
		   Case 2 Param="datediff(" & DataPart_D & ",transtime," & SqlNowString & ")>31"
		   Case 3 Param="datediff(" & DataPart_D & ",transtime," & SqlNowString & ")>61"
		   Case 4 Param="datediff(" & DataPart_D & ",transtime," & SqlNowString & ")>91"
		   Case 5 Param="datediff(" & DataPart_D & ",transtime," & SqlNowString & ")>181"
		   Case 6 Param="datediff(" & DataPart_D & ",transtime," & SqlNowString & ")>366"
		  End Select
		  
		  Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		  RS.Open "select * From KS_UserLogR Where " & param,conn,1,1
		  do while not rs.eof
			  Dim bType:bType=RS("Type")
			  Dim UserName:UserName=RS("UserName")
			  Dim MsgId:MsgId=RS("MsgId")
			  RS.Close:Set RS=Nothing
		
			  If BType=0 Then
				Conn.Execute("Delete From KS_UserLog Where ID=" & MsgId)
				Conn.Execute("Delete From KS_UserLogCMT Where MsgID=" & MsgId)
			  Else
				Conn.Execute("Update KS_UserLog set TransNum=TransNum-1  Where id=" & MsgId &" and TransNum>=1")
			  End If
				Conn.Execute("Delete From KS_UserLogR Where ID=" & id)
				Conn.Execute("Update KS_User set MsgNum=MsgNum-1  Where UserName='" & UserName &"' and MsgNum>=1")
		   rs.movenext
		  loop
		  rs.close
		  set rs=nothing
          KS.echo "<script>alert('恭喜,删除指定日期内的记录成功!');</script>"
		 End Sub
End Class
%> 
