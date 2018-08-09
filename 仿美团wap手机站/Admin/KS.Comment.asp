<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Comment
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Comment
        Private KS,ChannelID,Page,ChannelHomeUrl,KSCls,Action
		Private I, totalPut, CurrentPage, SqlStr,InfoID, ClassID,ProjectID
        Private RSObj,MaxPerPage
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
		   InfoID = KS.G("InfoID")
		   ClassID = KS.G("ClassID")
		   If Trim(ClassID) = "" Then ClassID = "0"
		   If ClassID <> "0" Then ClassID = "'" & Replace(ClassID, ",", "','") & "'"
		   If InfoID = "" Then InfoID = "0"
		   If InfoID <> "0" Then  InfoID = "'" & Replace(InfoID, ",", "','") & "'"
           Page = KS.G("Page")
		   ChannelID=KS.ChkClng(KS.G("ChannelID"))
		   ProjectID=KS.ChkClng(Request("ProjectID"))
		   
			If Not KS.ReturnPowerResult(ChannelID, "M010002") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
			End iF

		   Select Case KS.C_S(ChannelID,6)
			  Case 1:ChannelHomeUrl="KS.Article.asp"
			  Case 2:ChannelHomeUrl="KS.Picture.asp"
			  Case 3:ChannelHomeUrl="KS.Down.asp"
			  Case 4:ChannelHomeUrl="KS.Flash.asp"
			  Case 5:ChannelHomeUrl="KS.Shop.asp"
			  Case 7:ChannelHomeUrl="KS.Movie.asp"
			  Case 8:ChannelHomeUrl="KS.Supply.asp"
			 End Select
			 Action=KS.G("Action")
			 Select Case Action
			  Case "View"  Call CommentView()
			  Case "Verific" Call CommentVerific()
			  Case "Del" Call CommentDel()
			  Case "DelAllRecord" Call DelAllRecord()
			  Case "DelUnVerify" Call DelUnVerify()
			  Case Else	 Call CommentList()
			 End Select
		
		End Sub
		
		Sub DelUnVerify()
		  Conn.Execute("Delete From KS_Comment Where verific=0")
		  KS.Die "<script>alert('恭喜，一键清除了所有未审核的评论!');location.href='KS.Comment.asp?ProjectID=" & ProjectID &"';</script>"
		End Sub
		
		Sub DelAllRecord()
		  Dim Param
		  Select Case KS.ChkClng(KS.G("DelType"))
		   Case 1 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>11"
		   Case 2 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>31"
		   Case 3 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>61"
		   Case 4 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>91"
		   Case 5 Param="datediff(" & DataPart_D & ",AddDate," & SqlNowString & ")>181"
		   Case 6 Param="datediff(" & DataPart_Y & ",AddDate," & SqlNowString & ")>=1"
		   Case 7 Param="datediff(" & DataPart_Y & ",AddDate," & SqlNowString & ")>=2"
		  End Select
		  If ProjectID<>0 Then Param=Param &" and ProjectID=" & ProjectID
   		  If Param<>"" Then 
		   	 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select UserName,Anonymous,ID,Content,AddDate,channelid,infoid From KS_Comment Where " & Param,conn,1,1
			 Do While Not RS.Eof
			  Call ProcessUserScore(RS)
			  RS.MoveNext
			 Loop
			 RS.Close:Set RS=Nothing
		     Conn.Execute("Delete From KS_Comment Where " & Param)
		  End If
		  KS.echo "<script src=""../ks_inc/jquery.js""></script>"
          KS.echo "<script>$(parent.document).find('#ajaxmsg').toggle();alert('恭喜,删除指定日期评论成功!');location.href='KS.Comment.asp?ProjectID=" & ProjectID &"';</script>"
		 End Sub
		
        Sub CommentList
		If Request("page") <> "" Then
			  CurrentPage = KS.chkclng(Request("page"))
		Else
			  CurrentPage = 1
		End If
	With KS
	  .echo "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
	  .echo "<head>"
	  .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
	  .echo "<title>评论管理</title>"
	  .echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
	  .echo "<script language=""JavaScript"">" & vbCrLf
	  .echo "var ChannelID=""" & ChannelID & """;" & vbCrLf
	  .echo "var Page='" & CurrentPage & "';" & vbCrLf
	  .echo "var InfoID=""" & InfoID & """;" & vbCrLf
	  .echo "var ClassID=""" & ClassID & """;" & vbCrLf
	  .echo "</script>" & vbCrLf
	
	  .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
	  .echo "<script language=""JavaScript"" src=""../KS_Inc/Kesion.Box.js""></script>" & vbCrLf
	  .echo "<script language=""JavaScript"" src=""../KS_Inc/JQuery.js""></script>" & vbCrLf
%>
	<script language="javascript">
	function set(v)
	{
	{
	 if (v==1)
	 Verific(1,0);
	 else if (v==2)
	 Verific(0,0);
	 else if(v==3)
	  DelComment();
	 }
	}
	function Verific(OpType,CommentID)
	{
	if (CommentID==0) 
	 {
	 var ids=get_Ids(document.myform);
	if (ids!='')
	 {
	   location.href="KS.Comment.asp?ProjectID=<%=ProjectID%>&Action=Verific&ChannelID="+ChannelID+"&VerificType="+OpType+"&InfoID="+InfoID+"&ClassID="+ClassID+"&Page="+Page+"&CommentID="+ids;
	 }	
	else
	 alert('请选择评论!');
	 }
	 else
	   location.href="KS.Comment.asp?ProjectID=<%=ProjectID%>&Action=Verific&ChannelID="+ChannelID+"&VerificType="+OpType+"&InfoID="+InfoID+"&ClassID="+ClassID+"&Page="+Page+"&CommentID="+CommentID;
	}
	function DelComment()
	{
		var ids=get_Ids(document.myform);
		if (ids!=''){ 
	     if (confirm('真的要删除选中的评论吗?'))
	       location="KS.Comment.asp?ProjectID=<%=ProjectID%>&ChannelID="+ChannelID+"&Action=Del&InfoID="+InfoID+"&ClassID="+ClassID+"&Page="+Page+"&CommentID="+ids;
		}
		else{ alert('请选择要删除的评论!');}
	}
	function GetKeyDown()
	{ 
	if (event.ctrlKey)
	  switch  (event.keyCode)
	  {  case 90 : location.reload(); break;
		 case 65 : Select(0);break;
		 case 86 : event.keyCode=0;event.returnValue=false;ViewComment(0); break;
		 case 83 : event.keyCode=0;event.returnValue=false;Verific(1,0);break;
		 case 67 : event.keyCode=0;event.returnValue=false;Verific(0,0);break;
		 case 68 : DelComment();break;
	   }	
	else	
	 if (event.keyCode==46)DelComment();
	}
</script>
<%
	  .echo "</head>"
	'  .echo "<body scroll=no topmargin=""0"" leftmargin=""0"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
	  .echo "<ul id='menu_top'>"
	  .echo "<li onclick='javascript:Verific(1,0);' class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/verify.gif' border='0' align='absmiddle'>审核评论</span></li>"
	  .echo "<li onclick='Verific(0,0);' class='parent' onclick='Delete()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/unverify.gif' border='0' align='absmiddle'>取消审核</span></li>"
	  .echo "<li onclick='DelComment()' class='parent' onclick='Delete()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>删除评论</span></li>"
	  .echo "<li></li><form action='?' method='post'><div class='quicktz'>条件:<select name='searchtype'><option value='1'>文档标题</option><option value='2'>评论者</option><option value='3'>评论内容</option></select>关键字:<input type='text' class='textbox' name='keyword'> <input class='button' type='submit' value=' 搜 索 '></div></form></ul>"
	  if ProjectID<>0 then
	   dim rst:set rst=conn.execute("select top 1 * from KS_MoodProject where id=" & ProjectID)
	   if not rst.eof then
	    .echo "<div style='font-size:14px;margin:10px;font-weight:bold'>查看点评项目[<font color=green>" & RST("ProjectName") &"</font>]的所有用户点评&nbsp;<input type='button' class='button' value='返回点评管理中心' onclick=""location.href='KS.Mood.asp?typeflag=1'""/></div>"
	   end if
	   rst.close:set rst=nothing
	  end if
	  .echo "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
	  .echo ("<form name='myform' id='myform' method='Post' action='?channelid="& channelid & "'>")
	  .echo ("<input type='hidden' name='action' id='action' value='" & Action & "'>")
	  .echo ("<input type='hidden' name='ProjectID' value='" & ProjectID &"'>")
	  .echo "        <tr>"
	  .echo "          <td class=""sort"" width='30' align='center'>选择</td>"
	  .echo "          <td class=""sort"" align=""center"">评论内容</td>"
	  .echo "          <td width=""10%"" class=""sort"" align=""center"">发表人</td>"
	  .echo "          <td width=""10%"" class=""sort"" align=""center"">评论IP</td>"
	  .echo "          <td width=""15%"" align=""center"" class=""sort"">发表时间</td>"
	  .echo "          <td width=""10%"" class=""sort"" align=""center"">状态</td>"
	  .echo "          <td width=""12%"" class=""sort"" align=""center"">管理操作</td>"
	  .echo "        </tr>"

		      Set RSObj = Server.CreateObject("ADODB.RecordSet")
		 
			   Dim Param
			   If KS.G("ComeFrom")="Verify" Then
			   Param=" Where verific=0"
			   Else
			   Param=" Where 1=1"
			   End If
			   If ProjectID<>0 Then  Param=Param & " and projectid=" & KS.ChkClng(Request("ProjectID"))
			   If ChannelID<>0 Then Param=Param & " and ChannelID="& ChannelID&" "

			   If InfoID <> "0" Then
				 Param = Param & " And InfoID IN  (" & InfoID & ")"
			   End If
			   If KS.G("KeyWord")<>"" Then
			    Select Case KS.ChkClng(KS.S("SearchType"))
				 Case 1 Param=Param & " and InfoID In (select InfoID From [KS_ItemInfo] Where Title Like '%" & KS.G("KeyWord") & "%')"
				 Case 2 Param=Param & " and username='" & KS.G("KeyWord") & "'"
				 Case 3 Param=Param & " and Content Like '%" & KS.G("KeyWord") & "%'"
				End Select
			   End If
			   
			  SqlStr ="SELECT * FROM KS_Comment " & Param & " order by AddDate desc"
			   RSObj.Open SqlStr, conn, 1, 1
			 If RSObj.EOF And RSObj.BOF Then
			 Else
				        totalPut = conn.execute("select count(id) from ks_comment " & param)(0)
						If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
								RSObj.Move (CurrentPage - 1) * MaxPerPage
						End If
						Dim CommentXml:Set CommentXml=KS.ArrayToxml(RSObj.GetRows(MaxPerPage),RSObj,"row","xmlroot")
						Call showContent1(CommentXml)
						Set CommentXml=Nothing

		End If

      RSObj.Close:Set RSOBj=Nothing
	  CloseConn
	  .echo "</table>"
	  .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
	  .echo ("<tr><td width='180'><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> </div>")
	  .echo ("</td>")
	  .echo ("<td><select onchange='set(this.value)' name='setattribute'><option value=0>快速设置...</option><option value='1'>设为已审</option><option value='2'>设为未审</option><option value='3'>执行删除</option></select></td>")
	  .echo ("</form><td align='right'>")
	  .echo ("</td></tr></table>")
	  
	  	  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		 
		 .echo ("<div style=""clear:both""></div>")
	     .echo ("<form action='KS.Comment.asp?action=DelAllRecord&ProjectID=" & ProjectID&"' method='post' target='_hiddenframe'>")
		 .echo ("<iframe src='about:blank' style='display:none' name='_hiddenframe' id='_hiddenframe'></iframe>")
		 .echo ("<div class='attention'><strong>特别提醒： </strong><br>当站点运行一段时间后,网站的评论表可能存放着大量的记录,为使系统的运行性能更佳,您可以一段时间后清理一次。")
		 .echo ("<br /> <strong>删除范围：</strong><input name=""deltype"" type=""radio"" value=1>10天前 <input name=""deltype"" type=""radio"" value=""2"" /> 1个月前 <input name=""deltype"" type=""radio"" value=""3"" />2个月前 <input name=""deltype"" type=""radio"" value=""4"" />3个月前 <input name=""deltype"" type=""radio"" value=""5"" /> 6个月前 <input name=""deltype"" type=""radio"" value=""6""/> 1年前  <input name=""deltype"" type=""radio"" value=""7"" checked=""checked"" /> 2年前<input onclick=""$(parent.frames['FrameTop'].document).find('#ajaxmsg').toggle();"" type=""submit""  class=""button"" value=""执行删除"">")
		 .echo ("</form>")
		 
		 
		 .echo ("<br /><input type=""button"" onclick=""if (confirm('此操作不可逆，确定删除所有未审核的评论吗？')){location.href='?ProjectID=" & ProjectID &"&action=DelUnVerify';}""  class=""button"" value=""一键删除所有未审核的记录"">")
		 
		 
		 .echo ("</div>")
	  .echo "</body>"
	  .echo "</html>"
	 End With
	End Sub
	Sub ShowContent1(CommentXml)
	  With KS
	  Dim Node,ID
	  If IsObject(CommentXml) Then
		  For Each Node In CommentXml.DocumentElement.SelectNodes("row")
			  ID=Node.SelectSingleNode("@id").text
			    .echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & ID & "' onclick=""chk_iddiv('" & ID & "')"">"
			    .echo "<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" &ID & "')"" type='checkbox' id='c"& ID & "' value='" &ID & "'></td>"
			    .echo "  <td height='20' class='splittd' title='双击查看详细内容'><span CommentID='" & ID & "' ondblclick=""this.submit()"" title=""" & Node.SelectSingleNode("@content").text & """><img src='Images/ico_friend.gif' align='absmiddle'>"
			    .echo "  <span style='cursor:default;'>" & KS.GotTopic(Node.SelectSingleNode("@content").text, 42) & " "
			  If Node.SelectSingleNode("@replycontent").text<>"" Then   .echo "<font color=red>已回复</font>"
			    .echo " </span></span> </td>"
			    .echo "  <td align='center' class='splittd'>" & Node.SelectSingleNode("@username").text & " </td>"
			    .echo "  <td align='center' class='splittd'>" &Node.SelectSingleNode("@userip").text & " </td>"
			    .echo "  <td align='center' class='splittd'><FONT Color=red>" & Node.SelectSingleNode("@adddate").text & "</font> </td>"
			  If Node.SelectSingleNode("@verific").text = 0 Then
			     .echo "  <td align='center' class='splittd'><font color=red><span style='cursor:pointer' onclick='Verific(1," & ID & ")'>未审</span></font></td>"
			  Else
			     .echo "  <td align='center' class='splittd'><span style='cursor:pointer' onclick='Verific(0," & ID & ")'>已审</span></td>"
			  End If
			    .echo "  <td align='center' class='splittd'><a href='KS.Comment.asp?ProjectID=" & ProjectID &"&Action=View&ChannelID=" & ChannelID & "&CommentID=" & ID & "'>查看/回复</a>  <a href='KS.Comment.asp?ChannelID=" & ChannelID & "&Action=Del&CommentID=" & ID & "&ProjectID=" & ProjectID &"' onclick=""return(confirm('确定删除吗?'))"">删除</a></td>"
			    .echo "</tr>"	  
		  Next
	  End If
	 End With
	End Sub
	

          '删除评论
    Sub CommentDel()
		 	 Dim K, CommentID
			 CommentID = KS.FilterIds(KS.G("CommentID"))
			 If CommentID="" Then Call KS.AlertHintScript("没有选择记录!",-1)
			 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select UserName,Anonymous,ID,Content,AddDate,channelid,infoid From KS_Comment Where ID In(" & CommentID & ")",conn,1,1
			 Do While Not RS.Eof
			  Call ProcessUserScore(RS)
			  RS.MoveNext
			 Loop
			 RS.Close:Set RS=Nothing
			 Conn.Execute("Delete From KS_Comment Where id in(" & CommentID & ")")
			
			KS.Echo ("<script>location.href='KS.Comment.asp?ProjectID=" & ProjectID&"&ChannelID=" & ChannelID&"&ClassID=" & KS.G("ClassID") & "&InfoID=" & KS.G("InfoID") & "&Page=" & Page & "';</script>")
		 End Sub
		 '扣除一个月内被删除的用户积分
		 Sub ProcessUserScore(RS)
		      If Cint(RS(1))=0 And DateDiff("m",RS(4),Now)<1 Then
			     Dim RSU:Set RSU=Server.CreateObject("ADODB.RECORDSET")
				 RSU.Open "Select top 1 groupid From KS_User Where UserName='" & RS(0) & "'",conn,1,1
				 If Not RSU.Eof Then
				    If KS.ChkClng(KS.U_S(RSU(0),7))>0 and not Conn.Execute("Select top 1 id From KS_LogScore Where UserName='" & rs(0) & "' and ChannelID=1002 and InfoID=" & rs("channelid") & "" & rs("InfoID") & " And InOrOutFlag=1").Eof then
					Conn.Execute("Update KS_User Set Score=Score-" & KS.ChkClng(KS.U_S(RSU("GroupID"),7))  & " Where UserName='" & RS(0) & "'")
					
				    Dim CurrScore:CurrScore=Conn.Execute("Select top 1 Score From KS_User Where UserName='" & RS(0) & "'")(0)
					
			        Conn.Execute("Insert into KS_LogScore(UserName,InOrOutFlag,Score,CurrScore,[User],Descript,Adddate,IP,Channelid,InfoID) values('" & RS(0) & "',2," & KS.ChkClng(KS.U_S(RSU("GroupID"),7)) & ","&CurrScore & ",'系统','评论[" & KS.GotTopic(KS.HTMLEncode(RS(3)),36) & "]被删除!'," & SqlNowString & ",'" & replace(ks.getip,"'","""") & "',1002," & RS("ChannelID") & RS("InfoID") & ")")
					
					End If
				 End If
				 RSU.Close
			   End If
		 End Sub
		 
		 '审核评论
		 Sub CommentVerific()
		 	Dim K , CommentID,VerificType
			 VerificType = KS.ChkClng(KS.G("VerificType"))
			 CommentID = KS.FilterIds(KS.G("CommentID"))
			 If CommentID="" Then Call KS.AlertHintScript("没有选择记录!",-1)
			 If VerificType=1 Then 
			    Dim IDArr:IDArr=Split(CommentID,",")
				For K=0 To Ubound(IDArr)
				  Call VerifyAddScore(IDArr(k))
				Next
			 End If
			 Conn.Execute ("Update KS_Comment set Verific=" & VerificType & " Where ID in(" & CommentID & ")")
			 
			KS.Echo ("<script>location.href='KS.Comment.asp?ProjectID=" & ProjectID &"&ChannelID=" & ChannelID&"&ClassID=" & KS.G("ClassID") & "&InfoID=" & KS.G("InfoID") & "&Page=" & Page & "';</script>")
		 End Sub
		 
		sub VerifyAddScore(ID)
		          Dim RS:Set RS=Server.CreateObject("adodb.recordset")
				  rs.open "select top 1 u.userName,u.groupid,c.channelid,c.infoid from ks_comment c inner join ks_user u on c.userName=u.username where c.anonymous=0 and c.id=" & id,conn,1,1
				  If Not RS.Eof Then
				    If rs("channelid")<>1000 and KS.ChkClng(KS.U_S(rs(1),6))>0 Then
					 Dim RSA:Set RSA=Server.CreateObject("adodb.recordset")
					 RSA.Open "Select top 1 Title,Tid,Fname From " & KS.C_S(rs("ChannelID"),2) & " Where ID=" & rs("InfoID"),conn,1,1
					 If Not RSA.Eof Then
					 
						 Call  KS.ScoreInOrOut(rs("UserName"),1,KS.ChkClng(KS.U_S(rs("GroupID"),6)),"系统","参与文档[<a href=""" & KS.GetItemUrl(rs("channelid"),rsa(1),rs("infoid"),rsa(2)) & """ target=""_blank"">" & RSa(0) & "</a>]的评论!",1002,""&rs("ChannelID")&""&rs("InfoID"))
					 
					 End If
					 RSA.Close:Set RSA=Nothing
					End If
				  End If
				  rs.close:set rs=nothing
		End Sub
		
		'查看评论 
		Sub CommentView()
    	Dim CommentID:CommentID = KS.G("CommentID")
		Dim RSObj:Set RSObj=Server.CreateObject("ADODB.Recordset")
		RSObj.Open "Select top 1 * From KS_Comment Where ID=" & CommentID, conn, 1, 3
		If RSObj.EOF And RSObj.BOF Then KS.Echo ("参数传递出错!"):Exit Sub
		If KS.G("Flag")="Save" Then
		 RSObj("verific")=KS.ChkClng(Request.Form("verific"))
		 RSObj("Content")=Request.Form("Content")
		 RSObj("ReplyContent")=Request.Form("ReplyContent")
		 RSObj("ReplyTime")=Request.Form("ReplyTime")
		 RSObj("ReplyUser")=Request.Form("ReplyUser")
		 ChannelID=RSObj("ChannelID")
		 Dim InfoID:InfoID=RSOBj("InfoID")
		 RSObj.Update
		 If KS.ChkClng(Request.Form("verific"))=1 Then
		  Call VerifyAddScore(CommentID)
		    '自动生成内容页
			 If KS.C_S(Channelid,7)=1 or KS.C_S(ChannelID,7)=2 Then
				 Dim KSRObj:Set KSRObj=New Refresh
				Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
				RS.Open "select top 1 * From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID,Conn,1,1
						 Dim DocXML:Set DocXML=KS.RsToXml(RS,"row","root")
						 Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
						  KSRObj.ModelID=ChannelID
						  KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 
						  Call KSRObj.RefreshContent()
						  Set KSRobj=Nothing
				RS.Close
				Set RS=Nothing
			 End If
		 End If
		 KS.Echo "<script>alert('恭喜,评论修改成功!');location.href='" & Request.Form("ComeUrl") & "';</script>"
		End If
        With KS
			Dim ARS, Url,SqlStr,ChannelID,ReplyTime,ReplyUser
			ChannelID=KS.ChkClng(RSObj("ChannelID"))
			if channelid=1000 then
			 sqlstr="select top 1 ID,subject as Title,classid as Tid,0,0,0,0 from KS_GroupBuy Where ID=" & RSObj("InfoID")
			Else
			Select Case KS.C_S(ChannelID,6)
			 Case 1 SQLStr="select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes from " & KS.C_S(ChannelID,2) &" Where ID=" & RSObj("InfoID")
			 Case 2 SQLStr="select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,0 from " & KS.C_S(ChannelID,2) &" Where ID=" & RSObj("InfoID")
			 Case 3 SQLStr="select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,0 from " & KS.C_S(ChannelID,2) &" Where ID=" & RSObj("InfoID")
			 Case 4 SQLStr="select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,0 from " & KS.C_S(ChannelID,2) &" Where ID=" & RSObj("InfoID")
			 Case 5 SQLStr="select top 1 ID,Title,Tid,0,0,Fname,0 from " & KS.C_S(ChannelID,2) &" Where ID=" & RSObj("InfoID")
			 Case 7 SQLStr="select top 1 ID,Title,Tid,ReadPoint,InfoPurview,Fname,0 from " & KS.C_S(ChannelID,2) &" Where ID=" & RSObj("InfoID")
			 Case 8 SqlStr="select top 1 ID,Title,Tid,0,0,Fname,0 from " & KS.C_S(ChannelID,2) &" Where ID=" & RSObj("InfoID")
			End Select
		 End If
			
			ReplyTime=RSObj("ReplyTime")
			If ReplyTime="" Or IsNull(ReplyTime) Then
			 ReplyTime=Now
			End If
			ReplyUser=RSObj("ReplyUser")
			If ReplyUser=""  Or IsNull(ReplyUser) Then
			ReplyUser=KS.C("AdminName")
			End If
			dim itemname
			if channelid=1000 then
			itemname="团购"
			else
			itemname=KS.C_S(ChannelID,3)
			end if
			
				  .echo "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd"">"
	              .echo "<html xmlns=""http://www.w3.org/1999/xhtml"">"

				  .echo "<head>"
				  .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
				  .echo "<link href=""include/Admin_Style.css"" rel=""stylesheet"">"
				  .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
				  .echo "<title>查看评论</title>"
				  .echo "</head>"
				  .echo "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
				  .echo "<div class='topdashed sort'>评论查看/回复</div>"
				  .echo "  <br>"
				  .echo "   <table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""1"" Class=""Ctable"">"
				  .echo "    <form name='myform' action='KS.Comment.asp' method='post'>"
				  .echo "    <input type='hidden' value='" & Request.ServerVariables("HTTP_REFERER") & "' name='ComeUrl'>"
				  .echo "    <input type='hidden' value='" & ChannelID & "' name='ChannelID'>"
				  .echo "    <input type='hidden' value='" & CommentID & "' name='CommentID'>"
				  .echo "    <input type='hidden' value='" & ProjectID & "' name='ProjectID'>"
				  .echo "    <input type='hidden' value='View' name='Action'>"
				  .echo "    <input type='hidden' value='Save' name='Flag'>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td width=""200"" class='clefttitle' height=""25"">" &itemname &"标题</td>"
				  .echo "            <td> "
				  
				   Set Ars= Conn.Execute(SqlStr)
				   If Not ARS.EOF Then
					 Url = KS.GetItemUrl(ChannelID,aRS(2),ars(0),ars(5))
					 If ChannelID=1000 Then
					    Url="../shop/groupbuyshow.asp?id=" & ars(0)
					 Else
						 If ChannelID=1 Then
						  If ARS("Changes")=1 Then Url=ARS("Fname")
						 End IF
					 End If
					   .echo "<a href=""" & Url & """ target=""_blank"">" & ARS("title") & "</a>"
				   End If
				   ARS.Close:Set ARS = Nothing
				  .echo "          </td></tr>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td class='clefttitle' height=""25"">发表人</td>"
				  .echo "            <td> " & RSObj("UserName") & " 发表于 " & RSObj("AddDate") & " 用户IP:" & RSObj("UserIP") &"</td>"
				  .echo "          </tr>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td height=""25"" class='clefttitle' align=""center"">票数</td>"
				  .echo "            <td>支持:" & RSObj("score") & "票  反对" & RSObj("oscore") & "票</td>"
				  .echo "          </tr>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td height=""25"" class='clefttitle' align=""center"">评论状态</td><td>"
				  .echo " <input type='radio' value='1' name='verific'"
				 If RSObj("verific")=1 Then   .echo " checked"
				  .echo ">已审核"
				  .echo " <input type='radio' value='0' name='verific'"
				 If RSObj("verific")=0 Then   .echo " checked"
				  .echo ">未审核"
				  .echo "          </td></tr>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td height=""25"" class='clefttitle' align=""center"">评论内容"
				If RSObj("QuoteContent")<>"" And Not IsNull(RSObj("QuoteContent")) Then
				   .echo "<div style='color:red;font-weight:bold'><br />该评论内容有引用</div>"
				End If
				  .echo "</td>"
				  .echo "            <td><textarea name='Content' style=""overflow:auto;height:100px; width:380px;"">" & ReplaceFace(RSObj("Content")) & "</textarea></td>"
				  .echo "          </tr>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td height=""25"" class='clefttitle' align=""center"">回复内容</td>"
				  .echo "            <td><textarea name='ReplyContent' style=""overflow:auto;height:60px; width:380px;"">" & RSObj("ReplyContent") & "</textarea></td>"
				  .echo "          </tr>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td height=""25"" class='clefttitle' align=""center"">回复时间</td>"
				  .echo "            <td><input type='text' name='ReplyTime' class='textbox' value='" & ReplyTime & "'></td>"
				  .echo "          </tr>"
				  .echo "          <tr class='tdbg'>"
				  .echo "            <td height=""25"" class='clefttitle' align=""center"">回复人</td>"
				  .echo "            <td><input type='text' name='ReplyUser' class='textbox' value='" & ReplyUser & "'></td>"
				  .echo "          </tr>"
				
				  .echo "        </table>"

				  .echo "  <table width=""100%"" height=""30"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
				  .echo "    <tr>"
				  .echo "      <td height=""40"" align=""center"">"
				  .echo "        <input type='submit' class='button' value='确定修改'>"
				  .echo "        <input type=""button"" name=""Submit1"" onclick=""javascript:window.open('" & Url & "','new','');"" value=""查看" & itemname &""" class='button'>"
				  .echo "      </td>"
				  .echo "    </tr>"
				  .echo "</form>"
				  .echo "  </table>"
				  .echo "  <br>"
				  .echo "</body>"
				  .echo "</html>"
			End With
		End Sub
		
		Function ReplaceFace(c)
		 Dim str:str="惊讶|撇嘴|色|发呆|得意|流泪|害羞|闭嘴|睡|大哭|尴尬|发怒|调皮|呲牙|微笑|难过|酷|非典|抓狂|吐|"
		 Dim strArr:strArr=Split(str,"|")
		 Dim K
		 For K=0 To 19
		  c=replace(c,"[e"&K &"]","<img title=""" & strarr(k) & """ src=""" & KS.Setting(3) & "images/emot/" & K & ".gif"">")
		 Next
		 ReplaceFace=C
		End Function

End Class
%> 
