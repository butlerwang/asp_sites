<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_PhotoVote
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_PhotoVote
        Private KS,Action,Page,KSCls
		Private I, totalPut, CurrentPage,MaxPerPage, SqlStr,ChannelID,ItemName,ItemName1,RS
		Private OriginName, ID, Sex, Birthday, Telphone, UnitName, UnitAddress, Zip, Email, QQ, HomePage, Note, OriginType
		
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
		 	    .echo "<html>"
				.echo "<head>"
				.echo "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
				.echo "<title>投票记录管理</title>"
				.echo "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
             Action=KS.G("Action")
			 ChannelID=KS.ChkClng(KS.G("ChannelID"))
			 If ChannelID=0 Then ChannelID=2
				If Not KS.ReturnPowerResult(0, "KMST10014") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
				End iF

			 Page=KS.G("Page")
			 ItemName=KS.C_S(ChannelID,3)
			 
			 Select Case Action
			  Case "Del"
			    Call DiggDel()
		      Case "Clear"
			    Call ClearVote()
			  Case "SaveClear"
			    Call SaveClear()
			  Case Else
			   Call MainList()
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
		.echo "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		.echo "<script language='JavaScript' src='../KS_Inc/jquery.js'></script>"
		.echo "</head>"
		
		.echo "<body scroll=no topmargin='0' leftmargin='0'>"
		.echo "<ul id='mt'> <div id='mtl'>"
		
		.echo "<strong>查看详细的投票记录:</strong><select id='channelid' name='channelid' onchange=""if (this.value!=0){location.href='?channelid='+this.value;}"">"
			.echo " <option value='0'>---请选择模型---</option>"
	
			If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,Node
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1 and @ks6=2]")
			  If trim(ChannelID)=trim(Node.SelectSinglenode("@ks0").text) Then
			    .echo "<option value='" &Node.SelectSingleNode("@ks0").text &"' selected>" & Node.SelectSingleNode("@ks1").text & "</option>"

			  Else
			   .echo "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
			  End If
			next
			.echo "</select>"
		.echo "</div>"
		.echo "<li><a href='KS.PhotoVote.asp?Action=Clear&channelid=" & channelid &"'>按栏目清零票数</a></li>"
		.echo "</ul>"

    	.echo (" <div style=""position:absolute;top:4;right:8; overflow:hidden;"" >")
		.echo "  <select OnChange=""location.href='KS.PhotoVote.asp?ChannelID=" & ChannelID & "&id='+this.options[this.options.selectedIndex].value+'';"" style='width:120px' name='id'>"
		.echo "<option value=''>按栏目查看投票记录...</option>"
		.echo Replace(KS.LoadClassOption(ChannelID,false),"value='" & KS.S("ID") & "'","value='" & KS.S("ID") &"' selected") & "</select>&nbsp;&nbsp;"
		.echo ("</div>")
			
		.echo "<table width='100%'border='0' cellpadding='0' cellspacing='0'>"
        .echo(" <form name=""myform"" method=""Post"" action=""KS.PhotoVote.asp?Action=Del&ChannelID=" & ChannelID & """>")
		.echo "    <tr class='sort'>"
		.echo "    <td width='30' align='center'>选中</td>"
		.echo "    <td width='78' align='center'>参与用户</td>"
		.echo "    <td align='center'>投票时间</td>"
		.echo "    <td width='80' align='center'>参与用户IP</td>"
		.echo "    <td align='center'>栏目</td>"
		.echo "    <td align='center'>所投" & KS.C_S(ChannelID,3) & "</td>"
		.echo "  </tr>"
		 Set RS = Server.CreateObject("ADODB.RecordSet")
		   Dim Param:Param=" where ChannelID="& ChannelID
             If KS.S("ID")<>"" Then Param=Param & " and ClassID='" & KS.S("ID") & "'"
			  Param=Param & " order by VoteTime desc"
		   
				   SqlStr = "SELECT * FROM [KS_PhotoVote] " & Param
				   RS.Open SqlStr, conn, 1, 1
				 If RS.EOF And RS.BOF Then
				  .echo "<tr><td class='splittd' class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"" colspan=6 height='25' align='center'>没有投票记录!</td></tr>"
				 Else
					totalPut = RS.RecordCount
		
							If CurrentPage < 1 Then
								CurrentPage = 1
							End If
		
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
		
							If CurrentPage = 1 Then
								Call showContent
							Else
								If (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
									
									Call showContent
								Else
									CurrentPage = 1
									Call showContent
								End If
							End If
			End If
		 
		 .echo "</table>"
		 .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
		 .echo ("<tr><td width='180'><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> </div>")
		 .echo ("</td>")
	     .echo ("<td><input type=""submit"" value=""删除选中的记录"" onclick=""return confirm('确定要执行删除操作吗,删除后相应的票数将减少？');"" class=""button""></td>")
	     .echo ("<td align='right'>")
	Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	     .echo ("</td></tr></form></table>")

		End With
		End Sub
		Sub showContent()
		  With KS
			 Do While Not RS.EOF
			.echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & RS("ID") & "' onclick=""chk_iddiv('" & RS("ID") & "')"">"
		   .echo "<td class='splittd' align='center'><input name='id' onclick=""chk_iddiv('" &RS("ID") & "')"" type='checkbox' id='c"& RS("ID") & "' value='" &RS("ID") & "'></td>"
		  .echo " <td class='splittd' height='22'>" & RS("UserName") & "</td>"
		   
		   .echo " <td class='splittd' align='center'>" & RS("voteTime") & "</td>"
		   .echo " <td class='splittd' align='center'>" & RS("UserIP") & "</td>"
		   .echo " <td class='splittd' align='center'>" & KS.C_C(RS("ClassID"),1) & " </td>"
		   .echo " <td class='splittd' width='445' align='center'>" 
		    
			Dim RI:Set RI=Conn.Execute("select title from " & ks.c_s(channelid,2) & " where id in(" & rs("infoid") & ")")
			do while not RI.eof
			  .echo ri(0) & "  " 
			RI.MoveNext
			Loop
			RI.Close:Set RI=Nothing
		   
		   .echo "</td>"
		   .echo "</tr>"
							  I = I + 1
								If I >= MaxPerPage Then Exit Do
							   RS.MoveNext
							   Loop
								RS.Close
								 
		  End With
		 End Sub
		 
         Sub ClearVote()
		  With KS
		 	    .echo "<html>"
				.echo "<head>"
				.echo "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
				.echo "<title>投票记录管理</title>"
				.echo "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
			 ChannelID=KS.ChkClng(KS.G("ChannelID"))
		   .echo "</head>"
		
		.echo "<body scroll=no topmargin='0' leftmargin='0'>"
		.echo "<ul id='mt'> <div id='mtl'>查看详细的投票记录</div>"
		.echo "<li><a href='KS.PhotoVote.asp?Action=Clear&channelid=" & channelid &"'>按栏目清零票数</a></li>"
		.echo "</ul>"

    	.echo (" <div style=""position:absolute;top:4;right:8; overflow:hidden;"" >")
		.echo "  <select OnChange=""location.href='KS.PhotoVote.asp?ChannelID=" & ChannelID & "&id='+this.options[this.options.selectedIndex].value+'';"" style='width:120px' name='id'>"
		.echo "<option value=''>按栏目查看投票记录...</option>"
		.echo Replace(KS.LoadClassOption(ChannelID,false),"value='" & KS.S("ID") & "'","value='" & KS.S("ID") &"' selected") & "</select>&nbsp;&nbsp;"
		.echo ("</div>")
%>
             <table border='0' width='100%' cellspacing='1' cellpadding='1' class='ctable'>
			  <form name="myform" action="KS.PhotoVote.asp?Action=SaveClear&ChannelID=<%=ChannelID%>" method="post">
			  <tr class="tdbg">
			   <td width="200" height="30" class="clefttitle" align="right">选择要清零的栏目</td>
			   <td><select style='width:120px' name='id'>
			  <option value=''>按栏目清零...</option>
			   <%=Replace(KS.LoadClassOption(ChannelID,false),"value='" & KS.S("ID") & "'","value='" & KS.S("ID") &"' selected")%>
			  </select>
			  </td>
			  </tr>
			  <tr class="tdbg">
			    <td colspan=2 align="center"><input class="button" type="submit" value="确定清零" onclick="return(confirm('此操作不可逆，确定清零吗？'));"/></td>
			  </tr>
			  </form>
			 </table>
			 <br/>
			 <strong>说明：</strong><br>
			 &nbsp;清零操作将所选栏目的得票数设为“0”，一般在开始投票前操作，否则请不要进行此操作
<%
		  End With
		 End Sub
		 
		 Sub SaveClear()
		   If KS.S("ID")="" Then Call KS.AlertHistory("对不起，你没有选择栏目!",-1):Exit Sub
		   Conn.Execute("Update " & KS.C_S(ChannelID,2) & " set score=0 where tid='" & KS.S("ID") & "'")
		   Conn.Execute("Delete from KS_PhotoVote where classid='" & KS.S("ID") & "'")
		   KS.AlertHintScript "恭喜,清零成功!"
		 End Sub
		 
		 Sub DiggDel()
			Dim ID:ID = KS.G("ID")
			Dim IDArr:IDArr=Split(id,",")
			Dim I
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			For I=0 To Ubound(IDArr)
			 RS.Open "Select infoID From KS_PhotoVote Where ID=" & IDArr(i),conn,1,3
			 If Not RS.Eof Then
			  Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set score=score-1 Where ID in(" & RS(0) & ")")
			  RS.Delete
			 End iF
			 RS.Close
			Next
			Set RS=Nothing
		    response.redirect request.servervariables("http_referer") 
		 End Sub
End Class
%> 
