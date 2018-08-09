<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<!--#include file="Label/LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New LabelFolder
KSCls.Kesion()
Set KSCls = Nothing

Class LabelFolder
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim LabelFolderRS, FolderName, Descript, ParentID, FolderID, Action
		Dim  TS, ID, RSCheck, LabelType
		FolderID = Trim(Request("FolderID"))
		LabelType = Request("LabelType")
		Action = Request("Action")
		If FolderID = "" Then FolderID = "0"
		ParentID = FolderID
		Set LabelFolderRS = Server.CreateObject("ADODB.RecordSet")
		If Action = "EditFolder" Then
			LabelFolderRS.Open "select * from KS_LabelFolder where ID='" & FolderID & "'", Conn, 3, 3
		   If Not LabelFolderRS.EOF Then
			FolderName = LabelFolderRS("FolderName")
			Descript = LabelFolderRS("Description")
			ParentID = LabelFolderRS("ParentID")
			LabelType = LabelFolderRS("FolderType")
		 Else
			FolderName = ""
			Descript = ""
			ParentID = FolderID
		 End If
			LabelFolderRS.Close
		End If
		Select Case (Request.Form("Action"))
		 Case "Submit"
			If Request.Form("FolderName") = "" Then
				Call KS.alert("错误提示:\n\n请填写目录名称", "LabelFolder.asp?LabelType=" & LabelType & "&FolderID=" & FolderID)
				Response.End
			End If
			If KS.strLength(Trim(Request("FolderName"))) > 50 Then
			   Call KS.alert("目录名称不能超过25个汉字(50个英文字符)!", "LabelFolder.asp?LabelType=" & LabelType & "&FolderID=" & FolderID)
			   Response.End
			 End If
			 If KS.strLength(Trim(Request.Form("Descript"))) > 255 Then
			   Call KS.alert("目录描述不能超过个125汉字(255个英文字符)!", "LabelFolder.asp?LabelType=" & LabelType & "&FolderID=" & FolderID)
			  
			   Response.End
			 End If
			LabelFolderRS.Open "SELECT * FROM KS_LabelFolder WHERE FolderType=" & LabelType & " and FolderName='" & Request.Form("FolderName") & "'", Conn, 1, 1
			 If Not LabelFolderRS.EOF Then
			  LabelFolderRS.Close
			  Call KS.alert("错误提示:\n\n该目录已存在!", "LabelFolder.asp?LabelType=" & LabelType & "&FolderID=" & FolderID)
			
			   Response.End
			 Else
			   LabelFolderRS.Close
				If Request.Form("ParentID") <> "" Then
				   ParentID = Request.Form("ParentID")
				Else
				 ParentID = "0"
				End If
			   LabelFolderRS.Open "SELECT * FROM [KS_LabelFolder] WHERE ID='" & ParentID & "'", Conn, 1, 1
			   If Not LabelFolderRS.EOF Then
			   TS = LabelFolderRS("TS")
			   Else
			   TS = ""
			   End If
			   LabelFolderRS.Close
				'生成目录ID 年+12位随机
					Do While True
					ID = Year(Now()) & KS.MakeRandom(12)
					Set RSCheck = Conn.Execute("Select ID from [KS_LabelFolder] Where ID='" & ID & "'")
					 If RSCheck.EOF And RSCheck.BOF Then
					  RSCheck.Close
					  Set RSCheck = Nothing
					  Exit Do
					  End If
					Loop
				TS = TS & ID & ","
				LabelFolderRS.Open "SELECT * FROM [KS_LabelFolder]", Conn, 3, 3
				LabelFolderRS.AddNew
				LabelFolderRS("ID") = ID
				LabelFolderRS("FolderName") = Request.Form("FolderName")
				LabelFolderRS("Description") = Request.Form("Descript")
				LabelFolderRS("ParentID") = ParentID
				LabelFolderRS("TS") = TS
				LabelFolderRS("FolderType") = LabelType
				LabelFolderRS("AddDate") = Now
				LabelFolderRS("OrderID") = 0
				LabelFolderRS.Update
				Response.Write ("<script>if (confirm('成功提示:\n\n添加目录分类成功,继续添加目录分类吗?')){location.href='LabelFolder.asp?LabelType=" & LabelType & "&Folderid=" & FolderID & "';}else{top.frames['MainFrame'].location.reload();top.frames['MainFrame'].box.close();}</script>")
			End If
		Case "EditOk"
			  If Request.Form("FolderName") = "" Then
				Call KS.alert("错误提示:\n\n请填写目录名称", "LabelFolder.asp?Action=EditFolder&LabelType=" & LabelType & "&FolderID=" & FolderID)
				
				Response.End
			  End If
			  If KS.strLength(Trim(Request("FolderName"))) > 50 Then
			   Call KS.alert("目录名称不能超过25个汉字(50个英文字符)!", "LabelFolder.asp?Action=EditFolder&LabelType=" & LabelType & "&FolderID=" & FolderID)
			   
			   Response.End
			 End If
			 If KS.strLength(Trim(Request.Form("Descript"))) > 255 Then
			   Call KS.alert("目录描述不能超过个125汉字(255个英文字符)!", "LabelFolder.asp?Action=EditFolder&LabelType=" & LabelType & "&FolderID=" & FolderID)
			   
			   Response.End
			 End If
			   LabelFolderRS.Open "Select * From [KS_LabelFolder] Where FolderType=" & LabelType & " and ID<>'" & FolderID & "' AND FolderName='" & Request.Form("FolderName") & "'", Conn, 1, 1
			   If Not LabelFolderRS.EOF Then
				Call KS.alert("错误提示:\n\n目录名称已存在", "LabelFolder.asp?Action=EditFolder&LabelType=" & LabelType & "&FolderID=" & FolderID)
				
				Response.End
			   End If
			   LabelFolderRS.Close
			   ParentID=Request.Form("ParentID")
			   LabelFolderRS.Open "SELECT * FROM [KS_LabelFolder] Where ID='" & FolderID & "'", Conn, 3, 3
			    If  ParentID<> "" and Trim(ParentID)<>Trim(LabelFolderRS("ParentID")) Then
				  If  ParentID=FolderID Then 
				    LabelFolderRS.Close:Set LabelFolderRS=Nothing
					Call KS.AlertHintScript("错误提示:\n\n所属父栏目不能是自己!")
					Exit Sub
				  End If
				   
				   Dim PID
				   Dim RST:Set RST=Conn.Execute("Select top 1 * From [KS_LabelFolder] Where ID='" & ParentID & "'")
				   If Not RST.Eof Then
				      TS=RST("TS")
					  PID=RST("ID")
					  If Not Conn.Execute("Select top 1 * From [KS_LabelFolder] Where Ts like '" &PID & "%' And ParentID='" & FolderID & "'").EOF Then
				      RST.Close:Set RST=Nothing
					  Call KS.AlertHintScript("错误提示:\n\n所属父栏目不能是自己子栏目!")
					  Exit Sub
					  End If
				   End If
				   RST.Close:Set RST=Nothing
				   
				   LabelFolderRS("ParentID") = ParentID
				   
				   LabelFolderRS("TS")=TS & FolderID & ","
				End If

				LabelFolderRS("FolderName") = Request.Form("FolderName")
				LabelFolderRS("Description") = Request.Form("Descript")
				LabelFolderRS("AddDate") = Now
				LabelFolderRS.Update
				LabelFolderRS.Close
				Set LabelFolderRS=Nothing
				
				UpdateTS(FolderID)
				Response.Write ("<script>alert('成功提示:\n\n标签目录修改成功!');top.frames['MainFrame'].location.reload();top.frames['MainFrame'].box.close();</script>")
		End Select
		Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd"">"
        Response.Write "<html xmlns=""http://www.w3.org/1999/xhtml"">"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<title>标签类型</title>"
		Response.Write "<link href=""admin_style.css"" rel=""stylesheet"">"
		Response.Write "<script language=""JavaScript"" src=""Common.js""></script>"
		Response.Write "</head>"
		Response.Write "<body style=""background: #EAF0F5;"" scroll=no topmargin=""0"" leftmargin=""0"">"
		Response.Write "<br>"
		Response.Write "  <form name=""LabelFolderAddForm"" method=""post"" action=""LabelFolder.asp"">"
		Response.Write "  <input type=""hidden"" value=""" & LabelType & """ name=""LabelType"">"
		Response.Write "  <table width=""96%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class='border'>"
		Response.Write "    <tr class='title'>"
		Response.Write "      <td colspan=2 align=center height=25>"
		If Action = "EditFolder" Then
		   Response.Write "修改目录"
		   Else
		   Response.Write "创建新目录"
		  End If
		Response.Write "      </td></tr>"
		Response.Write "      <tr class='tdbg'>"
		Response.Write "      <td height=""30""> <div align=""center"">父目录</div></td>"
		Response.Write "      <td height=""30"">"
		   
				Response.Write ReturnLabelFolderTree(ParentID, LabelType)
			  
		Response.Write "      </td>"
		Response.Write "    </tr>"
		Response.Write "<tr class='tdbg'>"
		Response.Write "      <td width=""100"" height=""30""> <div align=""center"">目录名称</div></td>"
		Response.Write "      <td height=""30""> <input  name=""FolderName"" class='textbox' value=""" & FolderName & """ size=""35"">"
		Response.Write "              <font color=""#FF0000"">* 请输入有效的名称,不能包含特殊字符</font>"
		If Action = "" Then
		Response.Write "              <input type=""hidden"" name=""Action"" value=""Submit"">"
		Else
		Response.Write "              <input type=""hidden"" name=""Action"" value=""EditOk"">"
		End If
		Response.Write "</td>"
		Response.Write "        <input type='hidden' name='FolderID' value='" & FolderID & "'>"
		Response.Write "    </tr>"
		Response.Write "    <tr class='tdbg'>"
		Response.Write "      <td><div align=""center"">目录描述</div></td>"
		Response.Write "      <td><textarea name=""Descript"" cols=""60""  rows=""8"" class='textbox'>" & Descript & "</textarea></td>"
		Response.Write "       </td>"
		Response.Write "    </tr>"
		Response.Write "  </table>"
		Response.Write "  <table width=""100%"" height=""30"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "    <tr>"
		Response.Write "      <td height=""40"" align=""center"">"
		Response.Write "        <input type=""button"" name=""Submit"" class='button' Onclick=""CheckForm()"" value="" 确 定 "">"
		Response.Write "        <input type=""button"" name=""Submit2"" class='button' onclick=""top.frames['MainFrame'].box.close()"" value="" 取 消 "">"
		Response.Write "      </td>"
		Response.Write "    </tr>"
		Response.Write "  </table>"
		Response.Write "  </form>"
		Response.Write "</body>"
		Response.Write "</html>"
		Response.Write "<Script Language=""javascript"">" & vbCrLf
		Response.Write "function CheckForm()" & vbCrLf
		Response.Write "{ var form=document.LabelFolderAddForm;" & vbCrLf
		Response.Write "   if (form.FolderName.value=="""")" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""请输入目录名称!"");" & vbCrLf
		Response.Write "     form.FolderName.focus();"
		Response.Write "    return false;" & vbCrLf
		Response.Write "    }"
		Response.Write "    if (form.FolderName.value.length>50)" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""目录名称不能超过25个汉字(50个英文字符)!"");" & vbCrLf
		Response.Write "     form.FolderName.focus();"
		Response.Write "    return false;"
		Response.Write "    }"
		Response.Write "    if (form.Descript.value.length>255)" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""目录名称不能超过125个汉字(255个英文字符)!"");" & vbCrLf
		Response.Write "     form.Descript.focus();" & vbCrLf
		Response.Write "    return false;" & vbCrLf
		Response.Write "    }"
		Response.Write "    form.submit();" & vbCrLf
		Response.Write "    return true;" & vbCrLf
		Response.Write "}" & vbCrLf
		Response.Write "</Script>"
		End Sub
		
		Sub UpdateTS(ParentID)
		    Dim RS,TS
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From [KS_LabelFolder] Where  ID='" & ParentID & "'",Conn,1,1
			If Not RS.Eof Then
			  TS=RS("TS")  
			Else
			  Exit Sub
			End iF
			RS.Close
		    Set RS=Nothing

				Dim RST:Set RST=Server.CreateObject("ADODB.RECORDSET")
				RST.Open "Select * From [KS_LabelFolder] Where ParentID='" & ParentID & "'",Conn,1,3
				Do While Not RST.Eof
					 RST("TS")=TS & RST("ID") & ","
					 RST.Update
					 UpdateTS(RST("ID"))
					 RST.MoveNext
				Loop
				RST.Close
				Set RST=Nothing
		End Sub

End Class
%> 
