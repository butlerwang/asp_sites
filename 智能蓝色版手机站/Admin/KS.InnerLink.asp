<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Inner_Main
KSCls.Kesion()
Set KSCls = Nothing

Class Inner_Main
        Private KS,KSCls
		Private I
		Private totalPut
		Private CurrentPage
		Private SqlStr
		Private RSObj,XML
        Private MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		  Call KS.DelCahe(KS.SiteSN & "_InnerLink")
		 CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
		If Not KS.ReturnPowerResult(0, "KMST10004") Then          '检查是否有基本信息设置的权限
			Call KS.ReturnErr(1, "")
			Response.End
		End If
		If Request("page") <> "" Then
			  CurrentPage = CInt(Request("page"))
		Else
			  CurrentPage = 1
		End If
		With KS
		
		 .echo "<html>"
		 .echo "<head>"
		 .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		 .echo "<title>站点公告</title>"
		 .echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		 .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
		 .echo "<script language=""JavaScript"" src=""../KS_Inc/Jquery.js""></script>"
		
		 Select Case KS.G("Action")
		   Case "Add","Edit"   Call InnerAdd()
		   Case "AddSave"  Call InnerAddSave()
		   Case "Del"  Call InnerDel()
		   Case Else  Call InnerMainList()
		 End Select
		 End With
		End Sub
		
		Sub InnerMainList()
		
		With KS
	%>
	   <script language="javascript">
	  
		function InnerAdd()
		{
			new parent.KesionPopup().PopupCenterIframe('添加新站内链接','KS.InnerLink.asp?Action=Add',630,250,'no')
		}
		function EditInner(id)
		{
			new parent.KesionPopup().PopupCenterIframe('添加新站内链接',"KS.InnerLink.asp?Action=Edit&Flag=Edit&InnerID="+id,630,250,'no')
		}
		function DelInner(id)
		{
		 if (confirm('真的要删除选中的站内链接吗?'))
		 $("#myform").submit();
		}
		function InnerControl(op)
		{  var alertmsg='';
		   var ids=get_Ids(document.myform);
			if (ids!='')
			 {  if (op==1)
				{
				if (ids.indexOf(',')==-1) 
					EditInner(ids)
				  else alert('一次只能编辑一条站内链接!')	
				}
			  else if (op==2)    
				 DelInner(ids);
			 }
			else 
			 {
			 if (op==1)
			  alertmsg="编辑";
			 else if(op==2)
			  alertmsg="删除"; 
			 else
			  {
			  alertmsg="操作" 
			  }
			 alert('请选择要'+alertmsg+'的站内链接');
			  }
		}
		function GetKeyDown()
		{ 
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 : location.reload(); break;
			 case 65 : SelectAllElement();break;
			 case 78 : event.keyCode=0;event.returnValue=false; InnerAdd();break;
			 case 69 : event.keyCode=0;event.returnValue=false;InnerControl(1);break;
			 case 68 : InnerControl(2);break;
		   }	
		else	
		 if (event.keyCode==46)InnerControl(2);
		}
	   </script>
	<%
		 .echo "</head>"
		 .echo "<body topmargin=""0"" leftmargin=""0"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
		 .echo "<ul id='menu_top'>"
		 .echo "<li class='parent' onclick=""InnerAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加关键字</span></li>"
		 .echo "<li class='parent' onclick=""InnerControl(1);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>编辑关键字</span></li>"
		 .echo "<li class='parent' onclick=""InnerControl(2);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>删除关键字</span></li>"
		 .echo "</ul>"

		 .echo "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
		 .echo ("<form name='myform' id='myform' method='Post' action='?'>")
		 .echo ("<input type='hidden' value='Del' name='Action'>")
         .echo "  <tr>"
		 .echo "          <td height=""28"" class=""sort"" align=""center"">选择</div></td>"
		 .echo "          <td height=""28"" class=""sort"" align=""center"">待替换文字</div></td>"
		 .echo "          <td class=""sort""><div align=""center"">链接地址</div></td>"
		 .echo "          <td align=""center"" class=""sort"">新增时间</td>"
		 .echo "          <td align=""center"" class=""sort"">开始搜索位置</td>"
		 .echo "          <td align=""center"" class=""sort"">替换次数</td>"
		 .echo "          <td class=""sort""><div align=""center"">是否启用</div></td>"
		 .echo "  </tr>"
		  Set RSObj = Server.CreateObject("ADODB.RecordSet")
		  SqlStr = "SELECT * FROM KS_InnerLink order by AddDate desc"
		   RSObj.Open SqlStr, Conn, 1, 1
				 If RSObj.EOF And RSObj.BOF Then
				 Else
					       totalPut = Conn.Execute("Select Count(*) From KS_InnerLink")(0)
		
							If CurrentPage < 1 Then	CurrentPage = 1
		
							If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RSObj.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
							Set XML=KS.ArrayToXml(RSObj.GetRows(MaxPerPage),RSObj,"row","root")
							Call showContent
			End If
		   RSObj.Close
		   CloseConn

		    .echo "</table>"
			.echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
	        .echo ("<tr><td width='180'><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> </div>")
	        .echo ("</td>")
	        .echo ("<td><input type='button' value='删 除' onclick=""InnerControl(2)"" class='button'></td>")
	        .echo ("</form><td align='right'>")
			Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	        .echo ("</td></tr></table>")
			.echo "</body>"
		 .echo "</html>"
		 End With
		End Sub
		
		Sub showContent()
		 Dim Node,ID
		 With KS
		 If IsObject(XML) Then
		   For Each Node In XML.DocumentElement.SelectNodes("row")
		        ID=Node.SelectSingleNode("@id").text
			    .echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & ID & "' onclick=""chk_iddiv('" &ID & "')"">"
			    .echo "<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" & ID & "')"" type='checkbox' id='c"& ID & "' value='" & ID & "'></td>"
				.echo "  <td class='splittd' height='20'> &nbsp;&nbsp; <span InnerID='" & ID & "' ondblclick=""EditInner(this.InnerID)"">"
				.echo "    <span style='cursor:default;'>" & KS.GotTopic(Node.SelectSingleNode("@title").text, 45) & "</span></span> "
				.echo "  </td>"
				.echo "  <td class='splittd' align='center'><a href='" & Node.SelectSingleNode("@url").text & "' target='_blank'>" & Node.SelectSingleNode("@url").text & " </a></td>"
				.echo "  <td class='splittd' align='center'><FONT Color=red>" & Node.SelectSingleNode("@adddate").text & "</font> </td>"
				.echo "  <td class='splittd' align='center'>" & Node.SelectSingleNode("@casetf").text & " </td>"
				.echo "  <td class='splittd' align='center'>" & Node.SelectSingleNode("@times").text & "</td>"
				  If Node.SelectSingleNode("@opentf").text = "1" Then
				   .echo "  <td class='splittd' align='center'><font color=red>是</font></td>"
				  Else
				   .echo "  <td class='splittd' align='center'>否</td>"
				  End If
				.echo "</tr>"
			 Next
		  End If
		 End With
		 Set XML=Nothing
		End Sub
			 
		Sub InnerAdd()
			   With KS
			 		Dim InnerID, RSObj, SqlStr, Content, Title, Url, OpenTF, OpenType,CaseTF,Times,Start
					Dim Flag, Page
					OpenTF = 1
					Flag = KS.G("Flag")
					Page = KS.G("Page")
					If Page = "" Then Page = 1
					If Flag = "Edit" Then
						InnerID = KS.G("InnerID")
						Set RSObj = Server.CreateObject("Adodb.Recordset")
						SqlStr = "SELECT * FROM KS_InnerLink Where ID=" & InnerID
						RSObj.Open SqlStr, Conn, 1, 1
						  Title = RSObj("Title")
						  Url = RSObj("Url")
						  OpenType = RSObj("OpenType")
						  OpenTF = RSObj("OpenTF")
						  CaseTF = RSObj("CaseTF")
						  Times = RSObj("Times")
						  Start = RSObj("Start")
						RSObj.Close
					Else
					  Flag = "Add":OpenTF = 1:CaseTF=1:Times=-1 :Start=1
					End If
					
				

					 .echo "<table style=""margin-top:4px"" width=""99%"" align=""center"" class=""Ctable"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
					 .echo "  <form name=InnerForm method=post action=""KS.InnerLink.asp?Action=AddSave"">"
					 .echo "   <input type=""hidden"" name=""Flag"" value=""" & Flag & """>"
					 .echo "   <input type=""hidden"" name=""InnerID"" value=""" & InnerID & """>"
					 .echo "   <input type=""hidden"" name=""Page"" value=""" & Page & """>"
					 .echo "   <input type=""hidden"" name=""Content"" ID=""Content"" value=""" & Server.HTMLEncode(Content) & """>"
					 .echo "    <tr class=""tdbg"">"
					 .echo "        <td height=""28"" width=""150"" height=""28"" align=""right"" class=""clefttitle""><strong>待替换文字:</strong></td>"
					 .echo "        <td><input name=""Title"" type=""text"" id=""Title"" value=""" & server.HTMLEncode(Title) & """ class=""textbox"" style=""width:50%""> 如AspCMS,KesionCMS等</td>"
					 .echo "    </tr>"
					 .echo "    <tr class=""tdbg"">"
					 .echo "            <td height=""28"" class=""clefttitle"" align=""right"" clefttitle""><strong>链 接(URL):</strong></td>"
					 .echo "            <td>  <input name=""Url"" type=""text"" id=""Url""  value="""
					If Flag = "Edit" Then
					    .echo (Url)
					Else
					   .echo "http://"
					End If
					 .echo """ class=""textbox"" style=""width:50%""> </td>"
					 .echo "    </tr>"
					 .echo "    <tr class=""tdbg"">"
					 .echo "            <td height=""28"" align=""right"" class=""clefttitle""><strong>是否新窗口打开:</strong></td><td>"
								If OpenType = "_blank" Then
								  .echo ("<input type=""radio"" name=""OpenType"" value=""_blank"" checked>是")
								  .echo ("<input type=""radio"" name=""OpenType"" value=""""> 否")
								 Else
								  .echo ("<input type=""radio"" name=""OpenType"" value=""_blank"">是")
								  .echo ("<input type=""radio"" name=""OpenType"" value="""" checked> 否")
								 End If
					 .echo "             </td>"					
					 .echo "        </tr>"

					 .echo "          <tr class=""tdbg"">"
					 .echo "            <td height=""28"" align=""right"" class=""clefttitle""><strong>是否开启替换:</strong></td><td>"
								If OpenTF = 1 Then
								  .echo ("<input type=""radio"" name=""OpenTF"" value=""1"" checked>开启")
								  .echo ("<input type=""radio"" name=""OpenTF"" value=""0""> 关闭")
								 Else
								  .echo ("<input type=""radio"" name=""OpenTF"" value=""1"">开启")
								  .echo ("<input type=""radio"" name=""OpenTF"" value=""0"" checked> 关闭")
								 End If
					 .echo "             </td>"
					 .echo "    </tr>"
					 .echo "   <tr class=""tdbg"">"
					 .echo "            <td height=""28"" align=""right"" class=""clefttitle""><strong>区分大小写:</strong></td><td>"
					            If CaseTF = 1 Then
								  .echo ("<input type=""radio"" name=""CaseTF"" value=""1"" checked>不区分")
								  .echo ("<input type=""radio"" name=""CaseTF"" value=""0""> 区分")
								 Else
								  .echo ("<input type=""radio"" name=""CaseTF"" value=""1"">不区分")
								  .echo ("<input type=""radio"" name=""CaseTF"" value=""0"" checked> 区分")
								 End If
					 .echo "   </td>"
					 .echo "    </tr>"
					 .echo "   <tr class=""tdbg"">"
					 .echo "            <td height=""28"" align=""right"" class=""clefttitle""><strong>查找位置:</strong></td><td><input type=""text"" name=""Start""  value=""" & Start & """>缺省值是 1,表示从第一个位置开始查找.   </td>"
					 .echo "    </tr>"
					 .echo "   <tr class=""tdbg"">"
					 .echo "            <td height=""28"" align=""right"" class=""clefttitle""><strong>替换次数:</strong></td><td><input type=""text"" name=""times""  value=""" & times & """>替换的次数。如果忽略，缺省值是 -1,表示全部替换.   </td>"
					 .echo "    </tr>"
					
					
					 .echo "  </form>"
					 .echo "</table>"
						
					.echo "<div id='save'>"
					.echo "<li class='parent' onclick=""return(CheckForm())""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/save.gif' border='0' align='absmiddle'>确定保存</span></li>"
					.echo "<li class='parent' onclick=""parent.closeWindow();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>关闭取消</span></li>"
					.echo "</div>"

					 .echo "<script language=""JavaScript"">" & vbCrLf
					 .echo "<!--" & vbCrLf
					 .echo "function CheckForm()" & vbCrLf
					 .echo "{ var form=document.InnerForm;" & vbCrLf
					 .echo "  if (form.Title.value=='')" & vbCrLf
					 .echo "   {" & vbCrLf
					 .echo "    alert('请输入待替换文字!');" & vbCrLf
					 .echo "    form.Title.focus();" & vbCrLf
					 .echo "    return false;" & vbCrLf
					 .echo "   }" & vbCrLf
					 .echo "   if (form.Url.value=='')" & vbCrLf
					 .echo "   {" & vbCrLf
					 .echo "    alert('请输入链接网址!');" & vbCrLf
					 .echo "    form.Url.focus();" & vbCrLf
					 .echo "    return false;" & vbCrLf
					 .echo "   }" & vbCrLf
						 
					 .echo "   form.submit();" & vbCrLf
					 .echo "   return true;" & vbCrLf
					 .echo "}" & vbCrLf
					 .echo "//-->" & vbCrLf
					 .echo "</script>" & vbCrLf
				End With
			 End Sub
			 
			 Sub InnerAddSave()
				Dim InnerID, RSObj, SqlStr, Title, Url, AddDate, OpenType, OpenTF,Times,CaseTF,Start
				Dim Flag, Page, RSCheck
				Set RSObj = Server.CreateObject("Adodb.RecordSet")
				Flag = KS.G("Flag")
				InnerID = KS.G("InnerID")
				Title = KS.G("Title")
				Url = KS.G("Url")
				OpenType = KS.G("OpenType")
				OpenTF = KS.G("OpenTF")
				CaseTF  = KS.ChkClng(KS.G("CaseTF"))
				Times  = KS.ChkClng(KS.G("Times"))
				Start  = KS.ChkClng(KS.G("Start"))
				If OpenTF = "" Then OpenTF = 0
				If Title = "" Then Call KS.AlertHistory("待替换文字不能为空!", -1)
				If Url = "" Then Call KS.AlertHistory("链接地址不能为空!", -1)
				
				Set RSObj = Server.CreateObject("Adodb.Recordset")
				If Flag = "Add" Then
				   RSObj.Open "Select ID From KS_InnerLink Where Title='" & Title & "'", Conn, 1, 1
				   If Not RSObj.EOF Then
					  RSObj.Close
					  Set RSObj = Nothing
					  KS.AlertHintScript ("对不起,待替换文字已存在!")
					  Exit Sub
				   Else
					RSObj.Close
					RSObj.Open "SELECT top 1 * FROM KS_InnerLink Where (ID is Null)", Conn, 1, 3
					RSObj.AddNew
					  RSObj("Title") = Title
					  RSObj("Url") = Url
					  RSObj("AddDate") = Now
					  RSObj("OpenType") = OpenType
					  RSObj("OpenTF") = OpenTF
					  RSObj("CaseTF")=CaseTF
					  RSObj("Times")=Times
					  RSObj("Start")=Start
					RSObj.Update
					 RSObj.Close
				  End If
				   Set RSObj = Nothing
				    KS.Echo ("<script> if (confirm('站内链接添加成功!继续添加吗?')) {location.href='KS.InnerLink.asp?Action=Add';}else{top.frames[""MainFrame""].location.reload();top.closeWindow();}</script>")
				ElseIf Flag = "Edit" Then
				  Page = KS.G("Page")
				  RSObj.Open "Select ID FROM KS_InnerLink Where Title='" & Title & "' And ID<>" & InnerID, Conn, 1, 1
				  If Not RSObj.EOF Then
					 RSObj.Close
					 Set RSObj = Nothing
					 KS.AlertHintScript ("对不起,标题已存在!")
					 Exit Sub
				  Else
				   RSObj.Close
				   SqlStr = "SELECT * FROM KS_InnerLink Where ID=" & InnerID
				   RSObj.Open SqlStr, Conn, 1, 3
					 RSObj("Title") = Title
					 RSObj("Url") = Url
					 RSObj("OpenType") = OpenType
					 RSObj("OpenTF") = OpenTF
					 RSObj("CaseTF")=CaseTF
					 RSObj("Times")=Times
					 RSObj("Start")=Start
				   RSObj.Update
				   RSObj.Close
				   Set RSObj = Nothing
				  End If
				  Response.Write ("<script>alert('站内链接修改成功!');top.frames[""MainFrame""].location.reload();top.closeWindow();</script>")
				End If
			 End Sub
			
			'删除 
			 Sub InnerDel()
			     Dim InnerID
				 InnerID = Trim(KS.G("ID"))
				 Conn.Execute ("Delete From KS_InnerLink Where ID in(" & InnerID &")")
				 KS.AlertHintScript "恭喜,删除成功!"
			 End Sub
End Class
%>