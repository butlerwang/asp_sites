<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.CollectCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Collect_ItemModify2
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_ItemModify2
        Private KS
		Private KMCObj
		Private ConnItem
		Private Rs, Sql, FoundErr, ErrMsg, Action
		Private SqlItem, RsItem
		Private ItemID, ItemName, WebName, WebUrl, ChannelID, strChannelDir, ClassID, SpecialID, ItemDemo, LoginType, LoginUrl, LoginPostUrl, LoginUser, LoginPass, LoginFalse
		Private ListUrl, LsString, LoString, ListPageType, LPsString, LPoString, ListStr, ListPageStr1, ListPageStr2, ListPageID1, ListPageID2, ListPageStr3,CharsetCode
		Private tClass, tSpecial
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KMCObj=New CollectPublicCls
		  Set ConnItem = KS.ConnItem()
		End Sub
        Private Sub Class_Terminate()
		 Call KS.CloseConnItem()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KMCObj=Nothing
		End Sub
		Sub Kesion()
		FoundErr = False
		
		Action = Trim(Request("Action"))
		ItemID = Trim(Request("ItemID"))
		
		If ItemID = "" Then
				ItemName = Trim(KS.G("ItemName"))
				WebName = Trim(KS.G("WebName"))
				WebUrl = Trim(KS.G("WebUrl"))
				ChannelID = Trim(KS.G("ChannelID"))
				ClassID = Trim(KS.G("ClassID"))
				SpecialID = Trim(KS.G("SpecialID"))
				ItemDemo = Trim(KS.G("ItemDemo"))
				LoginType = KS.G("LoginType")
				LoginUrl = Trim(KS.G("LoginUrl"))
				LoginPostUrl = Trim(KS.G("LoginPostUrl"))
				LoginUser = Trim(KS.G("LoginUser"))
				LoginPass = Trim(KS.G("LoginPass"))
				LoginFalse = Trim(KS.G("LoginFalse"))
				CharsetCode=KS.G("CharsetCode")
				
				If ItemName = "" Then
				   FoundErr = True
				   ErrMsg = ErrMsg & "<br><li>项目名称不能为空</li>"
				End If
				If WebName = "" Then
				   FoundErr = True
				   ErrMsg = ErrMsg & "<br><li>网站名称不能为空</li>"
				End If
		
				If ChannelID = "" Or ChannelID = 0 Then
				   FoundErr = True
				   ErrMsg = ErrMsg & "<br><li>未指定频道</li>"
				Else
				   ChannelID = CLng(ChannelID)
				End If
				If ClassID = "" Then
				   FoundErr = True
				   ErrMsg = ErrMsg & "<br><li>未指定栏目</li>"
				Else
				   Set Rs = conn.Execute("select * From KS_Class Where ID='" & ClassID & "'")
				   If Rs.BOF And Rs.EOF Then
						 FoundErr = True
						 ErrMsg = ErrMsg & "<br><li>找不到指定的栏目</li>"
					End If
					strChannelDir = Rs("Folder")
					Set Rs = Nothing
				End If
				
				If SpecialID = "" Then   SpecialID = 0
		
				
				If LoginType = "" Then
				   FoundErr = True
				   ErrMsg = ErrMsg & "<br><li>请选择登录类型</li>"
				Else
				   LoginType = CLng(LoginType)
				   If LoginType = 1 Then
						 If LoginUrl = "" Or LoginPostUrl = "" Or LoginUser = "" Or LoginPass = "" Or LoginFalse = "" Then
						 FoundErr = True
						 ErrMsg = ErrMsg & "<br><li>请将登录参数填写完整</li>"
					  End If
				   End If
				End If
				
				If FoundErr <> True Then
				   SqlItem = "Select top 1 ItemID,ItemName,WebName,WebUrl,ChannelID,ChannelDir,ClassID,SpecialID,ItemDemo,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,CharsetCode From KS_CollectItem Where ItemName='" & ItemName & "'"
				   Set RsItem = Server.CreateObject("adodb.recordset")
				   RsItem.Open SqlItem, ConnItem, 1, 3
				   If RsItem.EOF And RsItem.BOF Then
				   RsItem.AddNew
				   RsItem("ItemName") = ItemName
				   RsItem("WebName") = WebName
				   RsItem("WebUrl") = WebUrl
				   RsItem("ChannelID") = ChannelID
				   RsItem("ChannelDir") = strChannelDir
				   RsItem("ClassID") = ClassID
				   RsItem("SpecialID") = SpecialID
				   RsItem("CharsetCode") = CharsetCode
				   If ItemDemo <> "" Then
					  RsItem("ItemDemo") = ItemDemo
				   End If
				   RsItem("LoginType") = LoginType
				   If LoginType = 1 Then
					  RsItem("LoginUrl") = LoginUrl
					  RsItem("LoginPostUrl") = LoginPostUrl
					  RsItem("LoginUser") = LoginUser
					  RsItem("LoginPass") = LoginPass
					  RsItem("LoginFalse") = LoginFalse
				   End If
				   ItemID = RsItem("ItemID")
				   RsItem.Update
				   Else
					 FoundErr = True
					 ErrMsg = "<br><li>存在相同的项目名称</li>"
				   End If
				   RsItem.Close: Set RsItem = Nothing
				End If

		Else
		   ItemID = CLng(ItemID)
		   SqlItem = "Select * From KS_CollectItem Where ItemID=" & ItemID
		   Set RsItem = Server.CreateObject("adodb.recordset")
		   RsItem.Open SqlItem, ConnItem, 1, 1
		   If RsItem.EOF And RsItem.BOF Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "<br><li>没有找到该项目!</li>"
		   Else
			  LoginType = RsItem("LoginType")
			  LoginUrl = RsItem("LoginUrl")
			  LoginPostUrl = RsItem("LoginPostUrl")
			  LoginUser = RsItem("LoginUser")
			  LoginPass = RsItem("LoginPass")
			  LoginFalse = RsItem("LoginFalse")
			  ListStr = RsItem("ListStr")
			  LsString = RsItem("LsString")
			  LoString = RsItem("LoString")
			  ListPageType = RsItem("ListPageType")
			  LPsString = RsItem("LPsString")
			  LPoString = RsItem("LPoString")
			  ListPageStr1 = RsItem("ListPageStr1")
			  ListPageStr2 = RsItem("ListPageStr2")
			  ListPageID1 = RsItem("ListPageID1")
			  ListPageID2 = RsItem("ListPageID2")
			  ListPageStr3 = RsItem("ListPageStr3")
			  If ListPageStr3 <> "" Then
				 ListPageStr3 = Replace(ListPageStr3, "|", Chr(13))
			  End If
		   End If
		   RsItem.Close
		   Set RsItem = Nothing
		End If
		
		If Action = "SaveEdit" And FoundErr <> True Then
		   Call SaveEdit
		End If
		
		If FoundErr = True Then
		   Call KS.AlertHistory(ErrMsg,-1)
		Else
		   Call Main
		End If
		End Sub
		
		Sub Main()
		
		   If FoundErr = True Then
			  Call KS.AlertHistory(ErrMsg,-1)
		
		   Else
		
		Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd"">"
		Response.Write "<html xmlns=""http://www.w3.org/1999/xhtml"">"
		Response.Write "<head>"
		Response.Write "<title>采集系统</title>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Admin_Style.css"">"
		Response.Write "<style type=""text/css"">"
		Response.Write "<!--" & vbCrLf
		Response.Write ".STYLE1 {color: #0000CC}" & vbCrLf
		Response.Write "-->" & vbCrLf
		Response.Write "</style>" & vbCrLf
		Response.Write "</head>"
		Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		Response.Write "<div class='topdashed'>"& KMCObj.GetItemLocation(2,ItemID) &"</div>"
		Response.Write "<br>"
		Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""ctable"" >"
		Response.Write "<form method=""post"" action=""Collect_ItemModify3.asp"" name=""form1"">"
		Response.Write "    <tr class='tdbg'>"
		Response.Write "      <td width=""20%"" height=""30"" align=""center"" class='clefttitle'>列表索引页面：</td>"
		Response.Write "      <td width=""75%"">"
		Response.Write "        <input class=""textbox"" name=""ListStr"" type=""text"" size=""50"" maxlength=""200"" value=""" & ListStr & """>"
		Response.Write "&nbsp;&nbsp;列表的第一页 </td>"
		Response.Write "    </tr>"
		Response.Write "    <tr class='tdbg'>"
		 Response.Write "     <td width=""20%"" align=""center"" class='clefttitle'>列表开始标记：</td>"
		 Response.Write "     <td width=""75%"">"
		 Response.Write "     <textarea name=""LsString"" cols=""49"" rows=""7"">" & LsString & "</textarea><br>      </td>"
		 Response.Write "   </tr>"
		 Response.Write "   <tr class='tdbg'>"
		 Response.Write "     <td width=""20%"" align=""center"" class='clefttitle'>列表结束标记：</td>"
		 Response.Write "     <td width=""75%"">"
		 Response.Write "     <textarea name=""LoString"" cols=""49"" rows=""7"">" & LoString & "</textarea><br>      </td>"
		 Response.Write "   </tr>"
		
		Response.Write "    <tr class='tdbg'>"
		Response.Write "      <td width=""20%"" height=""30"" align=""center"" class='clefttitle'> 列表索引分页：</td>"
		Response.Write "      <td width=""75%"">"
		 Response.Write "       <input type=""radio"" value=""0"" name=""ListPageType"" "
		 If ListPageType = 0 Then Response.Write "checked"
		 Response.Write " onClick=""ListPage1.style.display='none';ListPage12.style.display='none';ListPage2.style.display='none';ListPage3.style.display='none'"">不作设置&nbsp;"
		 Response.Write "       <input type=""radio"" value=""1"" name=""ListPageType"""
				 If ListPageType = 1 Then Response.Write "checked"
				 Response.Write " onClick=""ListPage1.style.display='';ListPage12.style.display='';ListPage2.style.display='none';ListPage3.style.display='none'"">设置标签&nbsp;"
		 Response.Write "       <input type=""radio"" value=""2"" name=""ListPageType"" "
		 If ListPageType = 2 Then Response.Write "checked"
		 Response.Write " onClick=""ListPage1.style.display='none';ListPage12.style.display='none';ListPage2.style.display='';ListPage3.style.display='none'"">批量生成&nbsp;"
		 Response.Write "       <input type=""radio"" value=""3"" name=""ListPageType"" "
		 If ListPageType = 3 Then Response.Write "checked"
		 Response.Write " onClick=""ListPage1.style.display='none';ListPage12.style.display='none';ListPage2.style.display='none';ListPage3.style.display=''"">手动添加      </td>"
		 Response.Write "   </tr>"
		 Response.Write "   <tr class='tdbg' id=""ListPage1"" style=""display:"
		 If ListPageType <> 1 Then Response.Write "none"
		 Response.Write """>"
		Response.Write "      <td width=""20%"" align=""center"" class='clefttitle'>下页开始标记："
		 Response.Write "       <p>　</p><p>　</p>"
		 Response.Write "       下页结束标记： </td>"
		 Response.Write "     <td width=""75%"">"
		 Response.Write "       <textarea name=""LPsString"" cols=""49"" rows=""7"">" & LPsString & "</textarea><br>"
		 Response.Write "       <textarea name=""LPoString"" cols=""49"" rows=""7"">" & LPoString & "</textarea>      </td>"
		 Response.Write "   </tr>"
		 Response.Write "   <tr class='tdbg' id=""ListPage12"" style=""display:"
		 If ListPageType <> 1 Then Response.Write "none"
		 Response.Write """>"
		 Response.Write "     <td width=""20%"" align=""center"" class='clefttitle'><span class=""STYLE1"">索引分页重定向： </span></td>"
		 Response.Write "     <td width=""75%"">"
		 Response.Write "       <input name=""ListPageStr1"" class=""textbox"" type=""text"" size=""58"" maxlength=""200"" value=""" & ListPageStr1 & """>      </td>"
		 Response.Write "   </tr>"
		 Response.Write "   <tr class='tdbg' id=""ListPage2"" style=""display:"
		 If ListPageType <> 2 Then Response.Write "none"
		 Response.Write """>"
		 Response.Write "     <td width=""20%"" align=""center"" class='clefttitle'><span class=""STYLE1"">批量生成：</span></td>"
		 Response.Write "     <td width=""75%"">"
		 Response.Write "       原字符串：<br>"
		  Response.Write "      <input class=""textbox"" name=""ListPageStr2"" type=""text"" size=""58"" maxlength=""200"" value=""" & ListPageStr2 & """><br>"
		  Response.Write "              格式：http://www.xxxxxx.com/list.asp?page={$ID}<br><br>"
		  Response.Write "      生成范围：<br>"
		 Response.Write "       <input class=""textbox"" name=""ListPageID1"" type=""text"" size=""8"" maxlength=""200"" value=""" & ListPageID1 & """><span lang=""en-us""> To </span><input class=""textbox"" name=""ListPageID2"" type=""text"" size=""8"" maxlength=""200"" value=""" & ListPageID2 & """><br>"
		 Response.Write "              格式：只能是数字，可升序或者降序。      </td>"
		 Response.Write "   </tr>"
		 Response.Write "   <tr class='tdbg' id=""ListPage3"" style=""display:"
		 If ListPageType <> 3 Then Response.Write "none"
		 Response.Write """>"
		 Response.Write "     <td width=""20%"" align=""center"" class='clefttitle'><span class=""STYLE1"">手动添加： </span></td>"
		 Response.Write "     <td width=""75%"">"
		 Response.Write "     <textarea name=""ListPageStr3"" cols=""49"" rows=""7"">" & ListPageStr3 & "</textarea><br>"
		 Response.Write "     格式：输入一个网址后按回车，再输入下一个。      </td>"
		 Response.Write "   </tr>"
		
		 Response.Write "   <tr class='tdbg'>"
		 Response.Write "     <td height=""30"" colspan=""2"" style=""text-align:center"">"
		 Response.Write "       <input name=""ItemID"" type=""hidden"" id=""ItemID"" value=""" & ItemID & """>"
		 Response.Write "       <input name=""Action"" type=""hidden"" id=""Action"" value=""SaveEdit"">"
		 Response.Write "     <input  type=""submit"" class='button' name=""Submit"" value=""下&nbsp;一&nbsp;步""></td>"
		 Response.Write "   </tr>"
		Response.Write "</form>"
		Response.Write "</table>"
		Response.Write "</body>"
		Response.Write "</html>"
		End If
		End Sub
		Sub SaveEdit()
		   ItemName = Trim(Request.Form("ItemName"))
		   WebName = Trim(Request.Form("WebName"))
		   WebUrl = Trim(Request.Form("WebUrl"))
		   ChannelID = Trim(Request.Form("ChannelID"))
		   ClassID = Trim(Request.Form("ClassID"))
		   SpecialID = Trim(Request.Form("SpecialID"))
		   LoginType = Trim(Request.Form("LoginType"))
		   LoginUrl = Trim(Request.Form("LoginUrl"))
		   LoginPostUrl = Trim(Request.Form("LoginPostUrl"))
		   LoginUser = Trim(Request.Form("LoginUser"))
		   LoginPass = Trim(Request.Form("LoginPass"))
		   LoginFalse = Trim(Request.Form("LoginFalse"))
		   ItemDemo = Request.Form("ItemDemo")
		   CharsetCode =KS.G("CharsetCode")
			  If ItemName = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "项目名称不能为空"
			  End If
			  If WebName = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "网站名称不能为空"
			  End If

			  If ChannelID = "" Or ChannelID = 0 Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "未指定频道"
			  Else
				 ChannelID = CLng(ChannelID)
			  End If
		
				If ClassID = "" Then
				   FoundErr = True
				   ErrMsg = ErrMsg & "未指定栏目"
				Else
				   ClassID = ClassID
				   Set Rs = conn.Execute("select * From KS_Class Where ID='" & ClassID & "'")
				   If Rs.BOF And Rs.EOF Then
						 FoundErr = True
						 ErrMsg = ErrMsg & "目标栏目不存在，请先建立"
					Else
					strChannelDir = Rs("Folder")
				   End If
					Set Rs = Nothing
				End If
			  
				 SpecialID = 0
		
			  If LoginType = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "请选择网站登录类型"
			  Else
				 LoginType = CLng(LoginType)
				 If LoginType = 1 Then
					If LoginUrl = "" Or LoginPostUrl = "" Or LoginUser = "" Or LoginPass = "" Or LoginFalse = "" Then
					   FoundErr = True
					   ErrMsg = ErrMsg & "网站登录信息不完整"
					End If
				 End If
			  End If
		   If FoundErr <> True Then
			  SqlItem = "Select top 1 *  From KS_CollectItem Where ItemID=" & ItemID
			  Set RsItem = Server.CreateObject("adodb.recordset")
			  RsItem.Open SqlItem, ConnItem, 2, 3
			  RsItem("ItemName") = ItemName
			  RsItem("WebName") = WebName
			  RsItem("CharsetCode") = CharsetCode
			  RsItem("WebUrl") = WebUrl
			  RsItem("ChannelID") = ChannelID
			  RsItem("ChannelDir") = strChannelDir
			  RsItem("ClassID") = ClassID
			  RsItem("SpecialID") = SpecialID
			  RsItem("LoginType") = LoginType
			  If LoginType = 1 Then
				 RsItem("LoginUrl") = LoginUrl
				 RsItem("LoginPostUrl") = LoginPostUrl
				 RsItem("LoginUser") = LoginUser
				 RsItem("LoginPass") = LoginPass
				 RsItem("LoginFalse") = LoginFalse
			  End If
			  RsItem("ItemDemo") = ItemDemo
			  RsItem.Update
			  RsItem.Close
			  Set RsItem = Nothing
		   End If
		End Sub
End Class
%> 
