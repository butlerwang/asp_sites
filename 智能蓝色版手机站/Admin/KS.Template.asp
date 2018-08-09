<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.FileIcon.asp"-->
<!--#include file="Include/Session.asp"-->
<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1

Dim KSCls
Set KSCls = New Template
KSCls.Kesion()
Set KSCls = Nothing

Class Template
        Private KS
		'===========================================================================
		Private I, totalPut, TemplateSql, KS_T_RS
		Private TemplateType, ChannelID,DomainStr
		Private FolderObj, FileObj, FileItem, CurrPath, ParentPath,InstallDir,FsoObj,Path,PhysicalPath
		'=============================================================================
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		   If KS.G("Action")<>"SelectTemplate" Then
			If Not KS.ReturnPowerResult(0, "KMTL10007") Then                '模板管理的权限检查
			  Call KS.ReturnErr(1, "")
			  Exit Sub
			End If
		   End If
			Set FsoObj = KS.InitialObject(KS.Setting(99))
		    CurrPath = KS.G("CurrPath")
		Response.Write "<html><head><title>模板管理-添加模板</title>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"

			Select Case KS.G("Action")
			 Case "Del"
			   Call TemplateDel()
			 Case "AddTemplateFile"
			   Call AddTemplateFile()
			 Case "AddTemplate","Modify"
			   Call AddTemplate()
			 Case "TemplateSave"
			   Call TemplateSave()
			 Case "SelectTemplate"
			   Call SelectTemplate()
			 Case Else
			   If KS.G("Action") = "FileReName" Then
				Dim NewFileName, OldFileName
				Path = Request("Path")
				If Path <> "" Then
					NewFileName = Request("NewFileName")
					OldFileName = Request("OldFileName")
					
					 If Instr(NewFileName ,".")=0 Then
						Call KS.AlertHistory("模板文件名格式不正确!", -1)
						Set KS = Nothing:Response.End
					 Else
					   If right(lcase(NewFileName),4) <>".css" and right(lcase(NewFileName),4) <>".htm" and right(lcase(NewFileName),5)<>".html" and right(lcase(NewFileName),6)<>".shtml" and right(lcase(NewFileName),5)<>".shtm" Then
						Call KS.AlertHistory("模板文件名式不正确,只能以html,htm,shtml或shtm为扩展名!", -1)
						Set KS = Nothing:Response.End
					   End If
					 End If
					 
					 If InStr(lcase(NewFileName),".asp")>0 or InStr(lcase(NewFileName),".asa")>0 or InStr(lcase(NewFileName),".php")>0 or InStr(lcase(NewFileName),".cer")>0  or InStr(lcase(NewFileName),".cdx")>0 Then
						Call KS.AlertHistory("模板文件名式不正确,请不要录入.asp|.php的扩展名!", -1)
						Set KS = Nothing:Response.End
					 End If
								
					
					
					If (NewFileName <> "") And (OldFileName <> "") Then
						PhysicalPath = Server.MapPath(Path) & "\" & OldFileName
						If FsoObj.FileExists(PhysicalPath) = True Then
							PhysicalPath = Server.MapPath(Path) & "\" & NewFileName
							If FsoObj.FileExists(PhysicalPath) = False Then
								Set FileObj = FsoObj.GetFile(Server.MapPath(Path) & "\" & OldFileName)
								FileObj.name = NewFileName
								Set FileObj = Nothing
							End If
						End If
					End If
					
					If right(lcase(NewFileName),4) =".css" Then
					  Response.Redirect "KS.Template.asp?flag=skin"
					End If
				End If
			ElseIf KS.G("Action") = "FolderReName" Then
				Dim NewPathName, OldPathName
				Path = Request("Path")
				If Path <> "" Then
					NewPathName = Request("NewPathName")
					OldPathName = Request("OldPathName")
					
					 If Instr(NewPathName ,".")<>0 or instr(NewPathName,";")<>0 Then
						Call KS.AlertHistory("目录名格式不正确!", -1)
						Set KS = Nothing:Response.End
					 End If
					
					
					If (NewPathName <> "") And (OldPathName <> "") Then
						PhysicalPath = Server.MapPath(Path) & "\" & OldPathName
						If FsoObj.FolderExists(PhysicalPath) = True Then
							PhysicalPath = Server.MapPath(Path) & "\" & NewPathName
							If FsoObj.FolderExists(PhysicalPath) = False Then
								Set FileObj = FsoObj.GetFolder(Server.MapPath(Path) & "\" & OldPathName)
								FileObj.name = NewPathName
								Set FileObj = Nothing
							End If
						End If
					End If
				End If
			End If
			   Call TemplateList()
			End Select
		End Sub
		
		Sub TemplateList()
		With Response
		InstallDir=KS.Setting(3)
        If CurrPath = "" Then
			ParentPath = ""
			CurrPath= InstallDir & KS.Setting(90)
		Else
			ParentPath = Mid(CurrPath, 1, InStrRev(CurrPath, "/") - 1)
			If ParentPath = "" Then
				ParentPath = Left(InstallDir, Len(InstallDir) - 1)
			End If
		End If
		
		If Request("flag")="skin" Then  '只显示css文件
		  If Ks.IsNul(KS.Setting(178)) Then
		    CurrPath=InstallDir & "imgaes/"  
		  Else
		    CurrPath=InstallDir & KS.Setting(178)  
		  End If
		End If
		
		If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath,Len(CurrPath)-1)		
		
		.Write "<script language=""JavaScript"">"
		.Write "var ParentPath=escape('" & ParentPath & "');" & vbCrLf
		.Write "var CurrPath='" & CurrPath & "';" & vbCrLf
		.Write "var ChannelID='" & ChannelID & "';" & vbCrLf
		.Write "var TemplateType='" & TemplateType & " ';" & vbCrLf
		.Write "</script>"
		.Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
		.Write "<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>"
		.Write "<script language=""JavaScript"" src=""Include/ContextMenu1.js""></script>"
		.Write "<script language=""JavaScript"" src=""Include/SelectElement.js""></script>"
		%>
		<script language="javascript">
		var DocElementArrInitialFlag=false;
		var DocElementArr = new Array();
		var DocMenuArr=new Array();
		var SelectedFile='',SelectedFolder='';
		$(document).ready(function()
			{  
			if (DocElementArrInitialFlag) return;
			InitialDocElementArr('FolderID','SelectObjID');
			InitialContextMenu();
			DocElementArrInitialFlag=true;
		});
		
		function InitialContextMenu()
		{   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.AddTemplateFile('');",'<%if request("flag")="skin" then response.write "新建CSS" Else Response.Write"新建模板"%>(N)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.TextEdit('text');",'文本编辑(W)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.TextEdit('');",'可视编辑(E)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.TemplateControl(2);",'删 除(D)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.location.reload();",'刷 新(Z)','disabled');
		}
		function DocDisabledContextMenu()
		{
			DisabledContextMenu('FolderID','SelectObjID','可视编辑(E),文本编辑(W),删 除(D)','文本编辑(W),可视编辑(E)','','','','')
		}
		function GoBack()
		{
		 location.href='?CurrPath='+ParentPath;
		}
		function OpenFolder(Obj)
		{
			var SubmitPath='';
			if (CurrPath=='/') SubmitPath=CurrPath+Obj.SelectObjID;
			else SubmitPath=CurrPath+'/'+Obj.SelectObjID;
			location.href='?CurrPath='+escape(SubmitPath);
		}
		function EditFolder(filename)   
		{
			var ReturnValue='';
			ReturnValue=prompt('修改的名称：',filename);
			if ((ReturnValue!='') && (ReturnValue!=null)) window.location.href='?Action=FolderReName&Path='+CurrPath+'&CurrPath='+CurrPath+'&OldPathName='+escape(filename)+'&NewPathName='+escape(ReturnValue);
				else if(ReturnValue!=null){alert('请填写要更名的名称');}
		}
		function EditFile(filename)
		{
			var ReturnValue='';
			ReturnValue=prompt('修改的名称：',filename);
			if ((ReturnValue!='') && (ReturnValue!=null)) window.location.href='?Action=FileReName&Path='+CurrPath+'&CurrPath='+CurrPath+'&OldFileName='+escape(filename)+'&NewFileName='+escape(ReturnValue);
				else if(ReturnValue!=null){alert('请填写要更名的名称');}
		}
		function AddTemplateFile()
		{
		  OpenWindow('KS.Frame.asp?t=<%=request("flag")%>&Action=AddTemplateFile&PageTitle='+escape('添加新的模板文件')+'&URL=KS.Template.asp&currpath='+escape(CurrPath),350,100,window);
		 location.reload();
		}
		
		
		function EditTemplate(id)
		{
		window.parent.parent.frames['MainFrame'].location.href='KS.Template.asp?Action=Modify&TemplateID='+id;
		$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape("模板管理中心 >> 编辑模板")+'&ButtonSymbol=TemplateAdd';
		}
		function TextEdit(Flag)
		{
			GetSelectStatus('FolderID','SelectObjID');
		 if (SelectedFile!='')
			if (SelectedFile.indexOf(',')==-1) 
			{
			 location.href='KS.Template.asp?Action=Modify&Flag='+Flag+'&FileName='+escape(SelectedFile)+'&CurrPath=<%=Server.UrlEncode(CurrPath)%>';
			 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape("模板管理中心 >> 编辑模板")+'&ButtonSymbol=Gosave';
			}
			else alert('一次只能编辑一个模板文件!')	 
	     else
		 alert('请选择要一个模板!');
		}
		function DelTemplate(id)
		{
		if (confirm('删除后将导致已绑定的信息找不到模板，确认操作吗?'))
		 location="KS.Template.asp?Action=Del&FileName="+id+'&CurrPath='+escape(CurrPath);
		}
		function TemplateControl(op)
		{
			var alertmsg='';
			GetSelectStatus('FolderID','SelectObjID');
			if (SelectedFile!='')
			 {  switch (op)
				{                    
				 case 2 :   
				   DelTemplate(SelectedFile); 
				   break;
				}   
			 }   
			else 
			 {
			   switch (op)
				{case 1 :alertmsg="编辑"; break;
				 case 2: alertmsg="删除"; break;
				 case 3: alertmsg="设置默认"; break;
				 default:alertmsg="操作" ;break;
				 } 
			 alert('请选择要'+alertmsg+'的模板');
			  }
		}
		function GetKeyDown()
		{ event.returnValue=false;
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 : location.reload(); break;
			 case 65 : AddTemplate('');break;
			 case 77 : AddTemplateFile();break;
			 case 69 : TextEdit('');break;
			 case 87 : TextEdit('text');break;
			 case 68 : TemplateControl(2);break;
			 case 83 : TemplateControl(3);break;
		   }	
		else	
		 if (event.keyCode==46)TemplateControl(2);
		}
		</script>
		<%
		.Write "</head>"
		.Write "<body scroll=no topmargin=""0"" leftmargin=""0"" onclick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"

        .Write "<ul id='menu_top'>"
		.Write "<li class='parent' onclick=""AddTemplateFile();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>"
		If Request("flag")="skin" then .write "新建CSS" Else .Write"新建模板"
		.Write"</span></li>"
		.Write "<li class='parent' onclick=""location.href='KS.Template.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>模板管理</span></li>"
		.Write "<li class='parent' onclick=""location.href='KS.Template.asp?flag=skin';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>CSS管理</span></li>"
		.Write "<li class='parent' "
		If Lcase(CurrPath & "/")=lcase(InstallDir & KS.Setting(90)) Then .Write " Disabled"
	    .Write " onclick=""GoBack();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>回上一级</span></li>"
		.Write "</ul>"	
		
		.Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")	
		.Write "<table width=""100%"" height=""25"" border=""0"" cellpadding=""0"" cellspacing=""1"">"
		.Write "  <tr align=""center"">"
		 If request("flag")<>"skin" Then
		.Write "    <td height=""25"" class=""sort""> <div align=""center""><font color=""#990000"">模板名称</font></div></td>"
		Else
		.Write "    <td height=""25"" class=""sort""> <div align=""center""><font color=""#990000"">CSS名称</font></div></td>"
		End If
		.Write "    <td width=""121"" class=""sort"">类型</td>"
		.Write "    <td width=""71"" class=""sort"">大小</td>"
		.Write "    <td width=""143"" class=""sort"">修改时间</td>"
		.Write "    <td width=""267"" class=""sort"">操作管理</td>"
		.Write "  </tr>"
		
		call ShowContent
		  
		.Write "</table>"
		.Write "</div>"
		.Write "</body>"
		.Write "</html>"
		End With
		End Sub
		Sub showContent()
		With Response
		Dim FsoItem,FileExtName
		Dim FolderObj:Set FolderObj = FsoObj.GetFolder(Server.MapPath(CurrPath))
		Dim SubFolderObj:Set SubFolderObj = FolderObj.SubFolders
		DiM FileObj:Set FileObj = FolderObj.Files

    If request("flag")<>"skin" Then
		For Each FsoItem In SubFolderObj
		  .Write "<tr class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
		  .Write "  <td class='splittd'><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		  .Write "      <tr>"
		  .Write "        <td height=""20"">&nbsp;"
		  .Write "         <img src=""Images/Folder/folderclosed.gif"" align=""absmiddle"">"
		  .Write "          <span style=""cursor:default""><a SelectObjID=" & FsoItem.name & " href=""?CurrPath=" & CurrPath &"/" & FsoItem.Name & """>" & FsoItem.name & "</a></span></td>"
		  .Write "      </tr>"
		  .Write "    </table></td>"
		  .Write "    <td align='center'  class='splittd'>文件夹</td>"
		  .Write "  <td align=""center""  class='splittd'>"
		  if FsoItem.Size<100 then
			 .Write FsoItem.Size &"Byte"
		  Else
			 .Write FormatNumber(FsoItem.Size/1024,1,-1) &"KB"
		  End if
		  .Write "  </td>"
		
		  .Write ("<td align='center'  class='splittd'>" & FsoItem.DateLastModified & " </td>")
		  .Write ("<td align='center'  class='splittd'><a href=""javascript:EditFolder('" & FsoItem.name & "')"">重命名</a> | <a href='KS.Template.asp?Action=Del&FileName=" & Server.URLEncode(FsoItem.name)&"&CurrPath=" & Server.URLEncode(CurrPath) & "' onclick=""return(confirm('此操作不可逆，确定删除吗？'))"">删除</a> </td>")
		  .Write "</tr>"
		  Next
	 End If	  
		 For Each FsoItem In FileObj
			FileExtName = LCase(Mid(FsoItem.name, InStrRev(FsoItem.name, ".") + 1))
		  If (request("flag")="skin" and FileExtName="css") or FileExtName="html" or FileExtName="htm" or FileExtName="shtml" Then
		  .Write "<tr class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
		  .Write "  <td  class='splittd'><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		  .Write "      <tr>"
		  .Write "        <td height=""20"">&nbsp;"
		  .Write "         <span SelectObjID=" & FsoItem.name & " onDblClick=""TextEdit('text');"">"
		  .Write "         <img src=""../Editor/KSPlus/FileIcon/"&GetFileIcon(FsoItem.name)&""" align=""absmiddle"">"
		  .Write "          <span style=""cursor:default"">" & FsoItem.name & "</span></span></td>"
		  .Write "      </tr>"
		  .Write "    </table></td>"
		  .Write "    <td align='center'  class='splittd'>" & FsoItem.Type & "</td>"
		  .Write "  <td align=""center""  class='splittd'>"
		  if FsoItem.Size<100 then
			 .Write FsoItem.Size &"Byte"
		  Else
			 .Write FormatNumber(FsoItem.Size/1024,1,-1) &"KB"
		  End if
		  .Write "  </td>"
		
		  .Write ("<td align='center'  class='splittd'>" & FsoItem.DateLastModified & " </td>")
		  .Write ("<td align='center'  class='splittd'><a href='KS.Template.asp?t=" & request("flag") & "&Action=Modify&FileName=" &Server.URLEncode(FsoItem.name)&"&Flag=text&CurrPath=" & Server.URLEncode(CurrPath) & "' onclick=""$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=模板管理中心 >> 编辑模板&ButtonSymbol=GoSave'"">文本编辑</a>")
		  If Request("flag")<>"skin" then
		  .Write (" | <a href='KS.Template.asp?Action=Modify&FileName=" & Server.URLEncode(FsoItem.name) &"&CurrPath=" & Server.URLEncode(CurrPath) & "'onclick=""$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=模板管理中心 >> 编辑模板&ButtonSymbol=GoSave'"">可视化编辑</a>")
		  End If
		  .Write (" | <a href=""javascript:EditFile('" & FsoItem.name & "')"">重命名</a> | <a href='KS.Template.asp?Action=Del&FileName=" & Server.URLEncode(FsoItem.name)&"&CurrPath=" & Server.URLEncode(CurrPath) & "' onclick=""return(confirm('此操作不可逆，确定删除吗？'))"">删除</a></td>")
		  .Write "</tr>"
		  End If
		  Next	
		  .write "</table>"
		  if request("flag")="skin" then
		  .Write "<br/><br/><div class=""attention"" style=""text-align:left""><font color=red>说明：只有放在" & KS.Setting(178) & "目录下的css样式文件才能在此管理。</font></div>" 
		  end if
		 End With 	   
	     End Sub
			 
			 '删除模板
		Sub TemplateDel
			Dim FileName,CurrPath
			FileName = Request("FileName")
			CurrPath = Request("CurrPath")
			Call KS.DeleteFile(CurrPath & "/" & FileName)
			Call KS.DeleteFolder(CurrPath & "/" & FileName)
			If right(lcase(FileName),4) <>".css" Then
		    Response.Write "<script>window.location.href='KS.Template.asp?CurrPath=" & CurrPath & "'</script>"
			Else
		    Response.Write "<script>window.location.href='KS.Template.asp?flag=skin&CurrPath=" & CurrPath & "'</script>"
			End If
       End Sub
	   
	  
	   
	   '添加新的模板文件
	   Sub AddTemplateFile()
	     Dim ChannelID,TemplateType,TemplateName,TemplateFile,FsoObj,PhysicalPath,NewTemplateStr,KS_RS_Obj,FsoFileName,IsDefault,SQLStr
	   %>
	   <html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; chaRSet=utf-8">
			<title>新建模板文件</title>
			<link href="Include/Admin_Style.css" rel="stylesheet">
		    <META HTTP-EQUIV="pragma" CONTENT="no-cache">
            <META HTTP-EQUIV="Cache-Control" CONTENT="no-cache, must-revalidate">
            <META HTTP-EQUIV="expires" CONTENT="Wed, 26 Feb 1997 08:21:57 GMT">
			<script language="JavaScript" src="../KS_Inc/common.js"></script>
			</head>
			<body style="background: #EAF0F5;" topmargin="0" leftmargin="0" scroll=no>
			<table width="90%" align="center" border="0" cellspacing="0" cellpadding="0">
			  <form name="myform" action="?Action=AddTemplateFile" method="post" onSubmit="return(CheckForm())">
			  <input type="hidden" value="Add" Name="Flag">
			  <tr> 
				<td height="18">&nbsp;</td>
			  </tr>
			  <tr> 
				<td  width="80" height="30" align="center">
				文 件 名：
				</td>
				<td>
				<%if request("t")="skin" then%>
				<%=KS.Setting(3)%><%=KS.Setting(178)%><input type="text" value="new.css" class='textbox' name="TemplateFile" size="25">
				<%else%>
				<%=KS.Setting(3) & KS.Setting(90)%><input type="text" value="<%=Replace(Request("CurrPath") & "/",KS.Setting(3) &  KS.Setting(90),"")%>Untitled.html" class='textbox' name="TemplateFile" size="25">
				<br><font color=#ff6666>请以html、htm为扩展名,如Article/Aritcle.html</font>
				<%end if%>
				</td>
			  </tr>
			  <tr align="center"> 
				<td height="30" colspan=2>
				 <input type="hidden" name="t" value="<%=request("t")%>"/>
				 <input type="hidden" name="CurrPath" value="<%=Request("CurrPath")%>">
				 <input type="submit" class="button" name="button1" value="确定新建"> 
				  &nbsp; <input type="button" class="button" onClick="window.close();" name="button2" value=" 取消 "> 
				</td>
			  </tr>
			  </form>
			</table>
			</body>
			</html>
			<script>
			function CheckForm()
			 {
				//if (CheckEnglishStr(document.myform.TemplateFile,"模板文件")==false) 
				 //  return false;
				if (!IsExt(document.myform.TemplateFile.value,'css')&&!IsExt(document.myform.TemplateFile.value,'htm')&&!IsExt(document.myform.TemplateFile.value,'html'))
				   { alert('模板文件的扩展名必须是.html或.htm');
					  document.myform.TemplateFile.focus(); 
					  return false;
				   }
			 return true;
			}
			</script>
	   <%
	   If Request.Form("Flag") = "Add" Then
		  TemplateFile=Request.Form("TemplateFile")
		  
		 If InStr(lcase(TemplateFile),".asp")>0 or InStr(lcase(TemplateFile),".asa")>0 or InStr(lcase(TemplateFile),".php")>0  or InStr(lcase(TemplateFile),".cer")>0 Then
			Call KS.AlertHistory("模板文件名格式不正确,请不要录入.asp|.php之类的扩展名!", -1)
			Set KS = Nothing:Response.End
		 End If
          
		  
		  If lcase(Right(TemplateFile,4))<>".css" and lcase(Right(TemplateFile,4))<>"html" And lcase(Right(TemplateFile,3))<>"htm" Then Call KS.AlertHistory("文件名的扩展名必须是html或htm",-1)
		  if request("t")="skin" then
		   TemplateFile=KS.Setting(3) & KS.Setting(178) & TemplateFile
		  Else
		   TemplateFile=Replace(KS.Setting(3) & KS.Setting(90) & TemplateFile,"//","/")
		  End If
		  KS.CreateListFolder(Replace(TemplateFile,Split(TemplateFile,"/")(Ubound(Split(TemplateFile,"/"))),""))
		 
		 if request("t")<>"skin" then
	      NewTemplateStr = "<html>" & vbnewline &"<head>" & vbnewline & "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">" & vbnewline
	      NewTemplateStr = NewTemplateStr & "<title>无标题文档</title>" & vbnewline & "</head>" & vbnewline & "<body>" & vbnewline & vbnewline & "</body>" & vbnewline & "</html>"
		 end if
		  Set FsoObj = KS.InitialObject(KS.Setting(99))
		  PhysicalPath=Server.Mappath(TemplateFile)

			if FsoObj.FileExists(PhysicalPath) = False then
				Set FileObj = FsoObj.CreateTextFile(PhysicalPath)
				FileObj.WriteLine(NewTemplateStr)
				Set FileObj = Nothing
			else
				Call KS.AlertHistory("文件名已经存在,请另取一个名称!",-1)
			end if
		if request("t")<>"skin" then
			Response.Write "<script>if (confirm('新建模板成功，继续添加吗?')){location.href='KS.Template.asp?Action=AddTemplateFile&ChannelID=" & ChannelID & "&TemplateType=" & TemplateType & "&CurrPath=" & Request("CurrPath") & "';}else{ window.close();}</script>"
		Else
			Response.Write "<script>if (confirm('增加CSS文件成功，继续添加吗?')){location.href='KS.Template.asp?t=skin&Action=AddTemplateFile&ChannelID=" & ChannelID & "&TemplateType=" & TemplateType & "&CurrPath=" & Request("CurrPath") & "';}else{ window.close();}</script>"
		End If
	   End If
	   End Sub
	   
	   '导入模板
	  Sub AddTemplate()
		Dim Action, TemplateID, ChannelID, TemplateType, TemplateName, FsoFileName, FnameType, TemplateContent, TemplateFileName, TemplateFromFileContent,Action1
		Dim  CurrPath, InstallDir, TemplateDIr,FileName
		InstallDir  = KS.Setting(3)
		CurrPath=Request("CurrPath")
		FileName=Request("FileName")
		TemplateFileName=CurrPath & "/" & FileName
        TemplateFromFileContent=KS.ReadFromFile(TemplateFileName)
		Response.Write "<html><head><title>模板管理-添加模板</title>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
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
		Response.Write "<script language=""JavaScript"" type=""text/JavaScript"" src=""../KS_Inc/common.js""></script>"
		Response.Write "<script language=""JavaScript"" type=""text/JavaScript"" src=""../KS_Inc/jQuery.js""></script>"
		Response.Write "</head>"
		if KS.G("Flag")<>"text" then
		Response.Write "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
		else
		Response.Write "<body scroll=no leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
		end if
		Response.Write " <table width='100%' height='25' border='0' cellpadding='0' cellspacing='1' bgcolor='#efefef'>"
        Response.Write " <tr>"
        Response.Write " <td class='sort'><div align='center'><font color='#990000'>"
		If Request("t")="skin" Then
		 Response.Write "在线修改CSS"
		Else
		 Response.Write "修改模板"
		End If
		Response.Write "</font></div></td>"
        Response.Write "  </tr>"
        Response.Write "</table>"
		 Response.Write "<table width='100%' height=""350"" style=""background-color:#EEEEEE;padding-right: 2px;padding-left: 2px;padding-bottom: 0px;"" border='0' align='center' cellpadding='0' cellspacing='0' class='ctable'>"
		 Response.Write " <form name=""TemplateForm"" method=""post"" action=""KS.Template.asp?Action=TemplateSave"" onSubmit=""return(CheckForm())"">"	
		
		 Response.Write "   <tr class=""clefttitle"">"
		 Response.Write "     <td height=""30"" style='text-align:left'><b>"
		 If Request("t")="skin" Then
		  Response.Write "CSS地址"
		 Else
		  Response.Write "模板地址"
		 End If
		 Response.Write "：</b><input name=""TemplateFileName"" type=""text"" id=""TemplateFileName"" size=""50"" Value=""" & TemplateFileName & """ readonly>"
		 Response.Write "  </td></tr>"
		 Response.Write "   <tr id=""toplabelarea"" class=""clefttitle"""
		 If Request("t")="skin" Then Response.Write " style='display:none'"
		 Response.Write ">"
		 Response.Write "	<td valign=""top"" style='text-align:left'><strong>插入标签：</strong>"
		 Response.Write "<select id=""mylabel"" style=""width:160px"">"
		 Response.Write " <option value="""">==选择系统函数标签==</option>"
		   Dim RS:Set RS=KS.InitialObject("adodb.recordset")
		   rs.open "select top 500 LabelName from KS_Label Where LabelType<>5 order by adddate desc",conn,1,1
		   If not Rs.eof then
		    Do While Not Rs.Eof
			 Response.Write "<option value=""" & RS(0) & """>" & RS(0) & "</option>"
			 RS.MoveNext
			Loop 
		   End If
		  Response.Write "</select>&nbsp;<input class='button' type='button' onclick='LabelInsertCode($(""#mylabel"").val());' value='插入标签'>"
		  RS.Close:Set RS=Nothing
		 Response.Write "&nbsp;<input type=""button"" class='button' onclick=""javascript:LabelInsertCode();"" value=""选择更多标签"">&nbsp;&nbsp;"
		IF Request("flag")="text" then Response.Write "<input type=""button"" value=""复制代码""  class=""button"" onclick=""copy()"">&nbsp;&nbsp;&nbsp;&nbsp;"
		 Response.Write " </td>"
		 Response.Write "</tr>"
		if KS.G("flag")="text" then 
		 Response.Write "   <tr id=""codearea"">"
		 Response.Write "   <td>"
		 Response.Write "     <table border='0' cellspacing='0' cellspadding='0'>"
		 Response.Write "	  <tr>"
		 Response.Write "       <td valign=""top"" width='20'>"
		 %>
		  <textarea name="txt_ln" id="txt_ln" cols="6" style="overflow:hidden;height:383px;background-color:highlight;border-right:0px;text-align:right;font-family: tahoma;font-size:12px;font-weight:bold;color:highlighttext;cursor:default;" readonly="">
<% Dim N
For N=1 To 3000
 Response.Write N & vbcrlf
Next
%>
</textarea>
		             </td>
		             <td valign="top"><textarea name="Content" rows="2" cols="30" id="Content" onscroll="show_ln('txt_ln','Content')" onKeyDown="editTab()" onChange="TemplateContent.SetContentIni();" style="height:382px;width:770;"><%=Server.HTMLEncode(TemplateFromFileContent)%></textarea>
</td>
		             </tr>
					 </table>
		</td>
		</tr>
					 <%
   else
		 Response.Write "   <tr id=""editorarea"">"
		 Response.Write "    <td colspan=""2"" height=""510"">"
		 %>
		 
		 	<textarea name="Content" rows="2" cols="30" id="Content" onscroll="show_ln('txt_ln','Content')" onKeyDown="editTab()" style="height:422px;width:770[x;"><%=Server.HTMLEncode(TemplateFromFileContent)%></textarea>  
		    <script type="text/javascript" src="../editor/ckeditor.js" mce_src="../editor/ckeditor.js"></script>
			<script type="text/javascript">
             CKEDITOR.replace('Content', {width:"99%",height:"400",toolbar:"Full",filebrowserBrowseUrl :"Include/SelectPic.asp?from=ckeditor&Currpath=<%=KS.GetUpFilesDir()%>",filebrowserWindowWidth:650,filebrowserWindowHeight:220,fullPage :true});
			 </script> 
		 <%

		 Response.Write "     </td>"
		 Response.Write "   </tr>"
end if
		 Response.Write " </form>"
		 Response.Write "</table>"
		 Response.Write "</body>"
		 Response.Write "</html>"
			 Conn.Close:Set Conn = Nothing
		Response.Write "<script language=""JavaScript"">" & vbCrLf
		Response.Write " SetFsoFileNameArea(" & TemplateType & ");" & vbCrLf
		Response.Write "function SetFsoFileNameArea(num)" & vbCrLf
		Response.Write "{if (num=='9993')" & vbCrLf
		Response.Write "document.all.FsoFileNameArea.style.display='';" & vbCrLf
		Response.Write "else" & vbCrLf
		Response.Write "document.all.FsoFileNameArea.style.display='none';" & vbCrLf
		Response.Write "}"

		Response.Write "    function copy()" & vbcrlf
		Response.Write "{" & vbcrlf
        Response.Write "    document.TemplateForm.Content.select();" & vbcrlf
        Response.Write "    textRange = document.TemplateForm.Content.createTextRange();" & vbcrlf
        Response.Write "    textRange.execCommand(""Copy"");" & vbcrlf
		Response.Write "    alert('恭喜，当前代码已复制到剪贴板!');" & vbcrlf
        Response.Write "}" & vbcrlf
		Response.Write "function LabelInsertCode(Val)" & vbcrlf
		Response.Write "{"
		Response.Write " if (Val==null)" & vbcrlf
		Response.Write "  Val=OpenWindow('Include/LabelFrame.asp?sChannelID=" & ChannelID &"&TemplateType=" & TemplateType &"&url=InsertLabel.asp&pagetitle='+escape('插入标签'),260,350,window);"&vbcrlf
		Response.Write "if (Val!='')"
		Response.Write "{ "
		If Request("flag")="text" Then
		Response.Write "document.TemplateForm.Content.focus();" & vbcrlf
		Response.Write "  var str = document.selection.createRange();" & vbcrlf
		Response.Write "  str.text = Val;"
		Else
		Response.Write "CKEDITOR.instances.Content.insertHtml(Val);"
		End If
		Response.Write " }" & vbcrlf
		Response.Write "}" & vbcrlf
		Response.Write "function WapLabelInsertCode(Val)" & vbcrlf
		Response.Write "{"
		Response.Write " if (Val==null)" & vbcrlf
		Response.Write "  Val=OpenWindow('Include/LabelFrame.asp?sChannelID=" & ChannelID &"&TemplateType=" & TemplateType &"&url=../Wap/InsertLabel.asp&pagetitle='+escape('插入WAP标签'),250,300,window);"&vbcrlf
		Response.Write "if (Val!='')"
		Response.Write "{ document.TemplateForm.Content.focus();" & vbcrlf
		Response.Write "  var str = document.selection.createRange();" & vbcrlf
		Response.Write "  str.text = Val;"
		Response.Write " }" & vbcrlf
		Response.Write "}" & vbcrlf
		Response.Write "function InsertFunctionLabel(Url,Width,Height)" & vbcrlf
        Response.Write "{" & vbcrlf
        Response.Write "var Val = OpenWindow(Url,Width,Height,window);"
		Response.Write "if (Val!=''&&Val!=null)"
		Response.Write "{ document.TemplateForm.Content.focus();" & vbcrlf
		Response.Write "  var str = document.selection.createRange();" & vbcrlf
		Response.Write "  str.text = Val;"
		Response.Write " }" & vbcrlf
        Response.Write "}" & vbcrlf
		Response.Write "function CheckForm()" & vbCrLf
		Response.Write "{  if (document.TemplateForm.TemplateFileName.value=="""")"
		Response.Write "     {"
		Response.Write "     alert(""请导入模板!"");"
		Response.Write "     return false;"
		Response.Write "     }" & vbCrLf

		Response.Write "    document.TemplateForm.submit();"
		Response.Write "    return true;"
		Response.Write "}" & vbCrLf
		Response.Write "</script>" & vbCrLf
	  End Sub
	  
	  Sub TemplateSave()
	  	Dim Action, ChannelID, TemplateType, TemplateName, TemplatConTent, TemplateFileName, TemplateID, FsoFileName, TemplateContent
		Dim ObjRS, SQLStr, IsDefault, TemplateFilePath, OpStr
		 TemplateName = Trim(Request("TemplateName"))
		 TemplateContent = Trim(Request("Content"))
		 TemplateFileName = Request("TemplateFileName")      '导入模板文件名及相对路径
		 TemplateFilePath = Replace(TemplateFileName, Mid(TemplateFileName, InStrRev(TemplateFileName, "/")), "")
		 If Instr(TemplateFileName ,".")=0 Then
			Call KS.AlertHistory("模板文件名格式不正确!", -1)
			Set KS = Nothing:Response.End
		 Else
		   If right(lcase(TemplateFileName),4) <>".css" and right(lcase(TemplateFileName),4) <>".htm" and right(lcase(TemplateFileName),5)<>".html" and right(lcase(TemplateFileName),6)<>".shtml" and right(lcase(TemplateFileName),5)<>".shtm" Then
			Call KS.AlertHistory("模板文件名格式不正确,只能以html,htm,shtml或shtm为扩展名!", -1)
			Set KS = Nothing:Response.End
		   End If
		 End If
		 
		 If InStr(lcase(TemplateFileName),".asp")>0 or InStr(lcase(TemplateFileName),".asa")>0 or InStr(lcase(TemplateFileName),".php")>0 or InStr(lcase(TemplateFileName),".cer")>0 or InStr(lcase(TemplateFileName),".cdx")>0 Then
			Call KS.AlertHistory("模板文件名格式不正确!", -1)
			Set KS = Nothing:Response.End
		 End If
		 
		 
		 
		 

			'检查数据正确性
			If TemplateFileName = "" Then
				  Call KS.AlertHistory("您还没有导入模板!", -1)
				  Set KS = Nothing
				  Response.End
			End If
			 TemplateContent = ReplaceBadStr(Replace(Replace(Replace(TemplateContent, "contentEditable=true", ""), KS.GetDomain, "/"), KS.Setting(2), ""))
			 
			Dim regEx:Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True
			regEx.Pattern = "@\s*LANGUAGE\s*=\s*[""]?\s*(vbscript|jscript|javascript).encode\b"
			If regEx.Test(TemplateContent) Then
				  Call KS.AlertHistory("您的模板格式不正确,请不要包含可执行注入脚本!", -1)
				  Set KS = Nothing
				  Response.End
			end if	
			
			regEx.Pattern = "execute\s*request"
			If regEx.Test(TemplateContent) Then
				  Call KS.AlertHistory("您的模板格式不正确,发现类似一句话木马!", -1)
				  Set KS = Nothing
				  Response.End
			end if
			
			regEx.Pattern = "executeglobal\s*request"
			If regEx.Test(TemplateContent) Then
				  Call KS.AlertHistory("您的模板格式不正确,发现类似一句话木马!", -1)
				  Set KS = Nothing
				  Response.End
			end if
			regEx.Pattern = "<script.*runat.*server(\n|.)*execute(\n|.)*<\/script>"
			If regEx.Test(TemplateContent) Then
				  Call KS.AlertHistory("您的模板格式不正确,发现类似一句话木马!", -1)
				  Set KS = Nothing
				  Response.End
			end if
			 
			 
			 If (Instr(TemplateContent,"<%")<>0) or (instr(TemplateContent,"<?")<>0 and instr(TemplateContent,"?>")<>0)  or instr(lcase(TemplateContent),"createobject(""adodb.stream"")")>0 Then
				  Call KS.AlertHistory("您的模板格式不正确,请不要包含可执行代码!", -1)
				  Set KS = Nothing
				  Response.End
			 End If
			 
			IF Instr(TemplateFileName,KS.Setting(3))=0 Then
		     TemplateFileName=Replace(KS.Setting(3) & TemplateFileName,"//","/")
		    End IF

			  If KS.CheckFile(TemplateFileName) = False Then KS.CreateListFolder (TemplateFilePath)
			  If KS.WriteTOFile(TemplateFileName, TemplateContent) = True Then
			   If right(lcase(TemplateFileName),4) =".css" Then
			    Response.Write ("<script> alert('恭喜，CSS在线修改成功!');top.frames['BottomFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?OpStr=模板管理&ButtonSymbol=Disabled'; location.href='KS.Template.asp?flag=skin';</script>")
			   Else
			    Response.Write ("<script> alert('成功提示:模板修改成功!');top.frames['BottomFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?OpStr=模板管理&ButtonSymbol=Disabled'; location.href='KS.Template.asp';</script>")
			   End If
			  Else
				Call KS.AlertHistory("错误提示:1.保存失败，模板文件不存在;\n2.请拷贝后，重新打开文件再保存", -1)
				Set KS = Nothing
			  End If
		End Sub
		Function ReplaceBadStr(Content)
			Dim regEx, Matches, Match
			Set regEx = New RegExp
			regEx.Pattern = "/" & KS.Setting(89) & "([A-Z]|[a-z]|\.|\?|\=|&|;|[0-9])*#"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			ReplaceBadStr = Content
			For Each Match In Matches
				On Error Resume Next
				ReplaceBadStr = Replace(ReplaceBadStr, Match.Value, "#")
			Next
		End Function
       '选择模板
	   Sub SelectTemplate()
	   	Dim CurrPath:CurrPath = Request("CurrPath")
		Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
		Response.Write "<title>选择模板文件</title><link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'></head>"
		Response.Write "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
		Response.Write "<table style=""width:100%;height:100%"" border='0' cellspacing='0' cellpadding='0' align='right'>"
		Response.Write "  <tr><td style=""width:100%""><select onChange='ChangeFolder(this.value);' id='FolderSelectList' style='width:100%;' name='select'>"
		Response.Write "        <option selected value='" & CurrPath & "'>" & CurrPath & "</option>"
		Response.Write "      </select></td>"
		Response.Write "  </tr>"
		Response.Write "  <tr>"
		Response.Write "    <td style=""width:100%;height:100%""><iframe id=""FolderList"" width='100%' height='100%' frameborder='1' src='Include/FolderList.asp?CurrPath=" & CurrPath & "'></iframe></td>"
		Response.Write "  </tr>"
		Response.Write "  <tr>"
		Response.Write "     <td height='25' align='center' background='images/titlebg.png'>"
		Response.Write "      <input type='button' class='button' onClick='SelectFile();' name='Submit' value=' 确 定 '>"
		Response.Write "    &nbsp;&nbsp;<input class='button' onClick='top.close();' type='button' name='Submit3' value=' 取 消 '>"
		Response.Write "      </td>"
		Response.Write "  </tr>"
		Response.Write "</table>"
		Response.Write "</body>"
		Response.Write "</html>"
		Response.Write "<script>"
		Response.Write "function ChangeFolder(CurrPath)"
		Response.Write "{"
		Response.Write "    frames[""FolderList""].location='Include/FolderList.asp?CurrPath='+escape(CurrPath);"
		Response.Write "}"
		Response.Write "function SelectFile(file)"&vbcrlf
		Response.Write "{  if (file==null) file=frames[""FolderList""].FileName;"&vbcrlf
		Response.Write "    if (file!='')"&vbcrlf
		Response.Write "    {"&vbcrlf
		Response.Write "       var templatedir=document.getElementById('FolderSelectList').value+'/'+file;"&vbcrlf
		Response.Write "       templatedir=templatedir.replace('" & KS.Setting(3) & KS.Setting(90) & "','{@TemplateDir}/');"&vbcrlf
		Response.Write "    if (document.all){window.returnValue=templatedir;}else{ top.opener.setVal(templatedir);}"&vbcrlf
		Response.Write "        top.close();"&vbcrlf
		Response.Write "    }"&vbcrlf
		Response.Write "    else{"&vbcrlf
		Response.Write "        alert('操作提示:\n\n对不起，您没有选择模板文件!');}"&vbcrlf
		Response.Write "}"&vbcrlf
		Response.Write "window.onunload=SetReturnValue;"
		Response.Write "function SetReturnValue()"
		Response.Write "{ "
		Response.Write "    if (typeof(window.returnValue)!='string') window.returnValue='';"
		Response.Write "}"
		Response.Write "</script>"

	   End Sub
	   
	  
 End Class
%>