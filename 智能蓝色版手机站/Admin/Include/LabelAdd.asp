<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<!--#include file="Label/LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New LabelAdd
KSCls.Kesion()
Set KSCls = Nothing

Class LabelAdd
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim LabelID, LabelRS, SQLStr, LabelName, Descript, LabelContent, LabelFlag, ParentID
		Dim Action, Page, RSCheck, FolderID
		Dim KeyWord, SearchType, StartDate, EndDate
		  
		'收集搜索参数
		KeyWord = Request("KeyWord")
		SearchType = Request("SearchType")
		StartDate = Request("StartDate")
		EndDate = Request("EndDate")
		
		With Response
		 if request.Form("Action")="AddNewSubmit" Then
		 .Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
		 End If
		 .Write "<head>"
		 .Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		 .Write "<script src='../../ks_inc/jquery.js'></script>"
		 .Write "<script src='../../ks_inc/lhgdialog.js'></script>"
		 .Write ("<body>")

		Set LabelRS = Server.CreateObject("Adodb.RecordSet")
		Action = Request.QueryString("Action")
		Page = Request("Page")
		If Action = "EditLabel" Then
			LabelID = Request("LabelId")
			Set LabelRS = Server.CreateObject("Adodb.Recordset")
			SQLStr = "SELECT top 1 * FROM [KS_Label] Where ID='" & LabelID & "'"
			LabelRS.Open SQLStr, Conn, 1, 1
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			Descript = LabelRS("Description")
			FolderID =LabelRS("FolderID")
			LabelContent = Server.HTMLEncode(LabelRS("LabelContent"))
			LabelRS.Close
		Else
		  LabelName=Request.QueryString("LabelName")
		  Descript=Request.QueryString("Description")
		  FolderID = Request.QueryString("FolderID")
		  LabelContent=Request.QueryString("LabelContent")
		  If LabelContent="" Then LabelContent="请输入您自定义的html代码"
		End If
		Select Case Request.Form("Action")
		 Case "AddNewSubmit"
			ParentID = Request.Form("ParentID")
			LabelName = Replace(Replace(Trim(Request.Form("LabelName")), """", ""), "'", "")
			Descript = Replace(Trim(Request.Form("Description")), "'", "")
			LabelContent = Trim(Request.Form("LabelContent"))
			LabelFlag = Request.Form("LabelFlag")
			FolderID = Request.Form("FolderID")
			If LabelName = "" Then
			   Call KS.AlertHistory("标签名称不能为空!", -1)
			   Set KS = Nothing
			   .End
			End If
			If LabelContent = "" Then
			  Call KS.AlertHistory("标签内容不能为空!", -1)
			  Set KS = Nothing
			  .End
			End If
			LabelName = "{LB_" & LabelName & "}"
			LabelRS.Open "Select top 1 LabelName From [KS_Label] Where LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
			  Call KS.AlertHistory("标签名称已经存在!", -1)
			  LabelRS.Close
			  Conn.Close
			  Set LabelRS = Nothing
			  Set Conn = Nothing
			  Set KS = Nothing
			  .End
			Else
				LabelRS.Close
				LabelRS.Open "Select * From [KS_Label] Where (ID is Null)", Conn, 1, 3
				LabelRS.AddNew
				  Do While True
					'生成ID  年+12位随机
					LabelID = Year(Now()) & KS.MakeRandom(10)
					Set RSCheck = Conn.Execute("Select ID from [KS_Label] Where ID='" & LabelID & "'")
					 If RSCheck.EOF And RSCheck.BOF Then
					  RSCheck.Close
					  Set RSCheck = Nothing
					  Exit Do
					 End If
				  Loop
				 LabelRS("ID") = LabelID
				 LabelRS("LabelName") = LabelName
				 LabelRS("LabelContent") = LabelContent
				 LabelRS("LabelFlag") = LabelFlag
				 LabelRS("Description") = Descript
				 LabelRS("FolderID") = FolderID
				 LabelRS("AddDate") = Now
				 LabelRS("LabelType") = 1
				 LabelRS("OrderID") = 1
				 LabelRS.Update
				 Call KS.FileAssociation(1021,1,LabelContent,0)
				 KS.echo "<script>$.dialog.confirm('恭喜，添加标签成功,继续添加标签吗?',function(){location.href='LabelAdd.asp?Action=AddNew&mode=text&LabelType=1&FolderID=" & FolderID & "';},function(){$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & FolderID & "&OpStr=标签管理 >> 自定义静态标签&ButtonSymbol=FreeLabel';parent.frames['MainFrame'].location.href='Label_Main.asp?LabelType=1&FolderID=" & FolderID & "';});</script>"
			End If
		Case "EditSubmit"
			LabelID = Trim(Request.Form("LabelID"))
			ParentID = Request.Form("ParentID")
			LabelName = Replace(Replace(Trim(Request.Form("LabelName")), """", ""), "'", "")
			Descript = Replace(Trim(Request.Form("Description")), "'", "")
			LabelContent = Trim(Request.Form("LabelContent"))
			LabelFlag = Request.Form("LabelFlag")
			If LabelName = "" Then
			   Call KS.AlertHistory("标签名称不能为空!", -1)
			   Set KS = Nothing
			   .End
			End If
			If LabelContent = "" Then
			  Call KS.AlertHistory("标签内容不能为空!", -1)
			  Set KS = Nothing
			  .End
			End If
			LabelName = "{LB_" & LabelName & "}"
			LabelRS.Open "Select top 1 LabelName From [KS_Label] Where ID <>'" & LabelID & "' AND LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
			  Call KS.AlertHistory("标签名称已经存在!", -1)
			  LabelRS.Close:Conn.Close:Set LabelRS = Nothing:Set Conn = Nothing
			  Set KS = Nothing
			  .End
			Else
				LabelRS.Close
				LabelRS.Open "Select top 1 * From [KS_Label] Where ID='" & LabelID & "'", Conn, 1, 3
				 LabelRS("LabelName") = LabelName
				 LabelRS("LabelContent") = LabelContent
				 LabelRS("LabelFlag") = LabelFlag
				 LabelRS("Description") = Descript
				 LabelRS("AddDate") = Now
				 LabelRS("FolderID") = Request.Form("ParentID")
				 LabelRS.Update
				 '遍历所有标签内容，找出所有标签的图片
				 Dim Node,UpFiles,RCls
				 UpFiles=LabelContent
				 if Not IsObject(Application(KS.SiteSN&"_labellist")) Then
				     Set RCls=New Refresh
				     Call Rcls.LoadLabelToCache()
					 Set Rcls=Nothing
				 End If
					 For Each Node in Application(KS.SiteSN&"_labellist").DocumentElement.SelectNodes("labellist")
					   UpFiles=UpFiles & Node.Text
					 Next
				 Call KS.FileAssociation(1021,1,UpFiles,1)
				 '遍历及入库结束
				 
				 If KeyWord = "" Then
				   	KS.Echo "<script>$.dialog.tips('<br/>恭喜，标签修改成功!',1,'success.gif',function(){$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & ParentID & "&OpStr=标签管理  >> 自定义静态标签&ButtonSymbol=FreeLabel';location.href='Label_main.asp?Page=" & Page & "&LabelType=1&FolderID=" & ParentID & "';});</script>"
				 Else
				   	KS.Echo "<script>$.dialog.tips('<br/>恭喜，标签修改成功!',1,'success.gif',function(){$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?OpStr=标签管理 >> <font color=red>搜索自定义静态标签结果</font>&ButtonSymbol=FreeLabelSearch';location.href='Label_main.asp?Page=" & Page & "&LabelType=1&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';});</script>"

				 End If
			End If
		End Select
		
		.Write "<html>"
		.Write "<head>"
		.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		.Write "<title>新建标签</title>"
		.Write "</head>"
		.Write "<link href=""Admin_Style.CSS"" rel=""stylesheet"">"
		.Write "<script language=""JavaScript"" src=""../../ks_inc/Common.js""></script>"
		%>
				<script language = 'JavaScript'>

		function show_ln(txt_ln,txt_main){
			var txt_ln  = document.getElementById(txt_ln);
			var txt_main  = document.getElementById(txt_main);
			txt_ln.scrollTop = txt_main.scrollTop;
			while(txt_ln.scrollTop != txt_main.scrollTop)
			{
				txt_ln.value += (i++) + '\n';
				txt_ln.scrollTop = txt_main.scrollTop;
			}
			return;
		}
		function editTab(){
			var code, sel, tmp, r
			var tabs=''
			event.returnValue = false
			sel =event.srcElement.document.selection.createRange()
			r = event.srcElement.createTextRange()
			switch (event.keyCode){
				case (8) :
				if (!(sel.getClientRects().length > 1)){
					event.returnValue = true
					return
				}
				code = sel.text
				tmp = sel.duplicate()
				tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)
				sel.setEndPoint('startToStart', tmp)
				sel.text = sel.text.replace(/\t/gm, '')
				code = code.replace(/\t/gm, '').replace(/\r\n/g, '\r')
				r.findText(code)
				r.select()
				break
			case (9) :
				if (sel.getClientRects().length > 1){
					code = sel.text
					tmp = sel.duplicate()
					tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)
					sel.setEndPoint('startToStart', tmp)
					sel.text = '\t'+sel.text.replace(/\r\n/g, '\r\t')
					code = code.replace(/\r\n/g, '\r\t')
					r.findText(code)
					r.select()
				}else{
					sel.text = '\t'
					sel.select()
				}
				break
			case (13) :
				tmp = sel.duplicate()
				for (var i=0; tmp.text.match(/[\t]+/g) && i<tmp.text.match(/[\t]+/g)[0].length; i++) tabs += '\t'
				sel.text = '\r\n'+tabs
				sel.select()
				break
			default  :
				event.returnValue = true
				break
				}
			}
			
		//-->
		</script>
		<%
		.Write "<script>"
		Response.Write "function LabelInsertCode(Val)" & vbcrlf
		Response.Write "{"
		Response.Write " if (Val==null)" & vbcrlf
		Response.Write "  Val=OpenWindow('LabelFrame.asp?sChannelID=0&TemplateType=0&url=InsertLabel.asp&pagetitle='+escape('插入标签'),260,350,window);"&vbcrlf
		Response.Write "if (Val!='')"
		Response.Write "{"
		If KS.G("Mode")="text" then
		Response.Write "document.LabelForm.LabelContent.focus();" & vbcrlf
		Response.Write "  var str = document.selection.createRange();" & vbcrlf
		Response.Write "  str.text = Val;"
		Else
		Response.Write "CKEDITOR.instances.LabelContent.insertHtml(Val);"
		End If
		Response.Write " }" & vbcrlf
		Response.Write "}" & vbcrlf
		Response.Write "function InsertFunctionLabel(Url,Width,Height)" & vbcrlf
        Response.Write "{" & vbcrlf
        Response.Write "var Val = OpenWindow(Url,Width,Height,window);"
		Response.Write "if (Val!=''&&Val!=null)"
		Response.Write "{ document.LabelForm.LabelContent.focus();" & vbcrlf
		Response.Write "  var str = document.selection.createRange();" & vbcrlf
		Response.Write "  str.text = Val;"
		Response.Write " }" & vbcrlf
        Response.Write "}" & vbcrlf
		.Write "</script>"
		.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		.Write "<table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		.Write "  <form name=""LabelForm"" id=""LabelForm"" method=post action="""" onSubmit=""return(CheckForm())"">"
		.Write "    <input type=""hidden"" name=""LabelFlag"" value=""3"">"
		.Write "    <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.Write "    <input type=""hidden"" name=""FolderID"" value=""" & FolderID & """>"
		.Write "    <input type=""hidden"" name=""Page"" value=""" & Page & """>"
			
			If Action = "AddNew" Or Action = "" Then .Write "<input type='hidden' name='Action' value='AddNewSubmit'>"
			If Action = "EditLabel" Then .Write "<input type='hidden' name='Action' value='EditSubmit'>"
			
		   .Write " <tr>"
		   .Write "   <td height=""25"" colspan=""2""> "
		 .Write "<table width='100%' height='25' border='0' cellpadding='0' cellspacing='1' bgcolor='#efefef' class='sort'>"
		 .Write "<tr><td><div align='center'><font color='#990000'>"
		  If Action = "EditLabel" Then
		   .Write "修改自定义静态标签"
		   Else
		   .Write "新建自定义静态标签"
		  End If
		 .Write "</font></div></td></tr>"
		.Write "</table>"
		.Write " </td>"
		.Write "    </tr>"
		.Write "    <tr class=""tableBorder1"">"
		.Write "      <td height=""19"">标签名称</td>"
		.Write "      <td><input value=""" & LabelName & """ name=""LabelName"" style=""width:200;"">"
		.Write "        <font color=""#FF0000""> 例如标签名称：&quot;推荐文章列表&quot;，则在模板中调用：&quot;{LB_推荐文章列表}&quot;（注意英文大小写及全半角）。</font></td>"
		.Write "    </tr>"
		.Write "    <tr class=""tableBorder1"">"
		.Write "      <td width=""60"" height=""19""> <div align=""left"">标签目录</div></td>"
		.Write "      <td>" & ReturnLabelFolderTree(FolderID, 1) & "<font color=""#FF0000"">请选择标签归属目录，以便日后管理标签</font></td>"
		.Write "    </tr>"
		.Write "    <tr class=""tableBorder1"">"
		.Write "      <td width=""60"" height=""16""><div align=""left"">标签简介</div></td>"
		.Write "      <td><textarea name=""Description"" rows=""3"" style=""width:100%;"">" & Descript & "</textarea></td>"
		.Write "    </tr>"
		.Write "    <tr><td colspan=""2"" align=""center"" height=""25"" class=""tableBorder1""><strong>自 定 义 静 态 标 签 内 容</strong></td></tr>"

		 Response.Write "   <tr class=""tableBorder1"" height=25>"
		 Response.Write "	<td  colspan=""2"">"
		 Response.Write "    &nbsp;&nbsp;&nbsp;&nbsp;"
		 Response.Write " <select name=""mylabel"" id=""mylabel"" style=""width:160px"">"
		 Response.Write " <option value="""">==选择系统函数标签==</option>"
		   Dim RS:Set RS=Server.Createobject("adodb.recordset")
		   rs.open "select LabelName from KS_Label Where LabelType<>5 order by adddate desc",conn,1,1
		   If not Rs.eof then
		    Do While Not Rs.Eof
			 Response.Write "<option value=""" & RS(0) & """>" & RS(0) & "</option>"
			 RS.MoveNext
			Loop 
		   End If
		  Response.Write "</select>&nbsp;<input class='button' type='button' onclick='LabelInsertCode(document.getElementById(""mylabel"").value);' value='插入标签'>"
		  RS.Close:Set RS=Nothing
		 Response.Write "&nbsp;<input type=""button"" class='button' onclick=""javascript:LabelInsertCode();"" value=""选择更多标签"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</Td>"
		 Response.Write "      </Tr>"
			If KS.G("Mode")="text" then

		.Write "    <tr class=""tableBorder1""><td  height=""250"" align='right'><textarea id='txt_ln' name='rollContent' cols='6' style='overflow:hidden;height:100%;background-color:highlight;border-right:0px;text-align:right;font-family: tahoma;font-size:12px;font-weight:bold;color:highlighttext;cursor:default;' readonly>"
		 Dim N
		 For N=1 To 3000
			.Write N & "&#13;&#10;"
		 Next
		 .Write"</textarea></td><td>"
		 .Write "<textarea name='LabelContent' style='width:100%;height:100%' ROWS='15' id='txt_main' onkeydown='editTab()' onscroll=""show_ln('txt_ln','txt_main')"" wrap='on'>" & LabelContent & "</textarea>" & vbNewLine
		 .Write "	<script>for(var i=3000; i<=3000; i++) document.getElementById('txt_ln').value += i + '\n';</script></td></tr>"
		Else
		.Write "    <tr><td colspan=""2"" height=""50""><textarea style=""width:100%"" type=""hidden"" ROWS='6' COLS='108' name=""LabelContent"">" & LabelContent & "</textarea></td></tr>"
		.Write "    <tr valign=""top"">"
		.Write "      <td colspan='2'> "
		%>
			<script type="text/javascript" src="../../editor/ckeditor.js" mce_src="../../editor/ckeditor.js"></script>
			<script type="text/javascript">
             CKEDITOR.replace('LabelContent', {width:"100%",height:"300",toolbar:"NewsTool",filebrowserBrowseUrl :"SelectPic.asp?from=ckeditor&Currpath=<%=KS.GetUpFilesDir()%>",filebrowserWindowWidth:650,filebrowserWindowHeight:290,fullPage :false});
			 </script> 

		<%
		.Write "      </td>"
		.Write "    </tr>"
		End iF
		
		.Write "  </form>"
		.Write "</table>"
		.Write "</body>"
		.Write "</html>"
		.Write "<script language=""JavaScript"">" & vbCrLf
		.Write "<!--" & vbCrLf
		.Write "function CheckForm()" & vbCrLf
		.Write "{ var form=document.LabelForm;"
		.Write "  if (form.LabelName.value=='')"
		.Write "   {"
		.Write "    alert('请输入标签名称!');"
		.Write "    form.LabelName.focus();"
		.Write "    return false;"
		.Write "   }"
		 If KS.G("Mode")="text" then
		 .Write " if (form.LabelContent.value==''||form.LabelContent.value=='请输入您自定义的html代码')"
		 .Write " {"
		 .Write "   alert('请输入标签内容!');"
		 .Write "   form.LabelContent.focus();"
		 .Write "   return false;"
		 .Write "  }"
		 End If
		 .Write "  if (form.Description.value.length>255)"
		 .Write "   {"
		 .Write "    alert(""目录名称不能超过125个汉字(255个英文字符)!"");"
		 .Write "    form.Description.focus();"
		 .Write "   return false;"
		 .Write "   }"
		 .Write "  form.submit();"
		 .Write "  return true;"
		.Write "}" & vbCrLf
		.Write "//-->" & vbCrLf
		.Write "</script>"
		
		Set Conn = Nothing
		
		End With
End Sub
End Class
%> 
