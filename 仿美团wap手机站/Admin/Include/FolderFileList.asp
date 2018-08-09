<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.FileIcon.asp"-->
<!--#include file="Session.asp"-->
<%

Dim KSCls
Set KSCls = New FolderFileList
KSCls.Kesion()
Set KSCls = Nothing

Class FolderFileList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		'主体部分
		Sub Kesion()
		Dim CurrPath, FsoObj, SubFolderObj, FolderObj, FileObj, I, FsoItem, OType
		Dim ParentPath, FileExtName, AllowShowExtNameStr
		Dim ShowVirtualPath
		Dim CanBackFlag
		Set FsoObj = KS.InitialObject(KS.Setting(99))
		Dim InstallDir, UpFilesDir, Hypothesized, ChannelID
		
		InstallDir = KS.Setting(3)
		UpFilesDir = KS.Setting(91)
		Hypothesized = KS.Setting(3)
		
		On Error Resume Next
		OType = Request("Type")
		If OType <> "" Then
			Dim Path, PhysicalPath
			If OType = "DelFolder" Then
				Path = Request("Path")
				If Path <> "" Then
					Path = Server.MapPath(Path)
					If FsoObj.FolderExists(Path) = True Then FsoObj.DeleteFolder Path
				End If
			ElseIf OType = "DelFile" Then
				Dim DelFileName
				Path = Request("Path")
				DelFileName = Request("FileName")
				If (DelFileName <> "") And (Path <> "") Then
					Path = Server.MapPath(Path)
					If FsoObj.FileExists(Path & "\" & DelFileName) = True Then FsoObj.DeleteFile Path & "\" & DelFileName
				End If
			ElseIf OType = "AddFolder" Then
				Path = Replace(Request("Path"),".","")
				If Path <> "" Then
					Path = Server.MapPath(Path)
					if instr(path,";")>0  or instr(lcase(path),".asp")>0 or instr(lcase(path),".php")>0 or instr(lcase(path),".asa")>0 then
					 Response.Write ("<script>alert('对不起，目录不合法！');location.href='?CurrPath=" & KS.GetUpFilesDir & "';</script>")
					 Response.end
					end if
					If FsoObj.FolderExists(Path) = True Then
						Response.Write ("<script>alert('对不起，目录已经存在！');</script>")
					Else
						FsoObj.CreateFolder Path
					End If
				End If
			ElseIf OType = "FileReName" Then
				Dim NewFileName, OldFileName
				Path = Request("Path")
				If Path <> "" Then
					NewFileName = Request("NewFileName")
					if instr(lcase(NewFileName),".asp")<>0 or instr(lcase(NewFileName),".php")<>0 or instr(lcase(NewFileName),".asa")<>0  or instr(lcase(NewFileName),".cer")<>0 or instr(NewFileName,";")<>0 then 
					 Response.Write ("<script>alert('对不起，扩展名不合法！');location.href='?CurrPath=" & KS.GetUpFilesDir & "';</script>")
					 Response.end
					else
						OldFileName = Request("OldFileName")
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
					end if
				End If
			ElseIf OType = "FolderReName" Then
				Dim NewPathName, OldPathName
				Path = Replace(Request("Path"),".","")
				If Path <> "" Then
					NewPathName = Replace(Request("NewPathName"),".","")
					OldPathName = Replace(Request("OldPathName"),".","")
					if instr(NewPathName,";")>0  or instr(lcase(NewPathName),".asp")>0 or instr(lcase(NewPathName),".php")>0 or instr(lcase(NewPathName),".asa")>0 then
					 Response.Write ("<script>alert('对不起，目录不合法！');location.href='?CurrPath=" & KS.GetUpFilesDir & "';</script>")
					 Response.end
					end if
					
					
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
		End If
		
		ShowVirtualPath = KS.G("ShowVirtualPath")
		AllowShowExtNameStr = "jpg,txt,gif,bmp"
		CurrPath = KS.G("CurrPath")
		ChannelID = KS.G("ChannelID")
		If ChannelID = "" Or Not IsNumeric(ChannelID) Then ChannelID = 0
		If CurrPath = "" Then
			ParentPath = ""
		Else
			ParentPath = Mid(CurrPath, 1, InStrRev(CurrPath, "/") - 1)
			If ParentPath = "" Then
				ParentPath = Left(InstallDir, Len(InstallDir) - 1)
			End If
		End If
		If ChannelID <> 0 Then
		 Session("CurrPath") = CurrPath
		 KS.CreateListFolder(CurrPath)
		End If
		Set FolderObj = FsoObj.GetFolder(Server.MapPath(CurrPath))
		Set SubFolderObj = FolderObj.SubFolders
		Set FileObj = FolderObj.Files
		
		
		Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<title>文件和目录列表</title>"
		Response.Write "<script src=""../../ks_inc/jquery.js""></script>"
		Response.Write "</head>"
		%>
		<script>
		function OpenFolder(Obj)
		{  	 
		    var SubmitPath='';
			if (CurrPath=='/') SubmitPath=CurrPath+Obj;
			else SubmitPath=CurrPath+'/'+Obj;
			location.href='FolderFileList.asp?CurrPath='+SubmitPath;
			AddFolderList(parent.document.getElementById('FolderSelectList'),SubmitPath,SubmitPath);
		}
		
		function SetFile(File)
		{  
		    parent.SetFileUrl();
			return;
		}
		
		function AddFolderList(SelectObj,Label,LabelContent)
		{
			var i=0,AddOption;
			if (!SearchOptionExists(SelectObj,Label))
			{
				AddOption = document.createElement("OPTION");
				AddOption.text=Label;
				AddOption.value=LabelContent;
				SelectObj.add(AddOption);
				SelectObj.options(SelectObj.length-1).selected=true;
			}
		}
        function SearchOptionExists(Obj,SearchText)
		{ 
			var i,flag=false;
			var AddOption;
			for(i=0;i<Obj.length;i++)
			{
				if (Obj.options(i).text==SearchText)
				{
					Obj.options(i).selected=true;
					flag=true;
					return true;
				}
			}
			
			//if (!flag)
			//{
			//	AddOption = document.createElement("OPTION");
			//	AddOption.text=SearchText;
			//	AddOption.value=SearchText;
			//	Obj.add(AddOption);
			//	Obj.options(Obj.length-1).selected=true;
			//	}
			return false;
		}		
		function SelectFile(Obj,file)
		{	
		    PreviewFile(file);
			try{
		    var PathArticle='',TempPath='';
			for (var i=0;i<document.all.length;i++)
			{
				if (document.all(i).className=='FolderSelectItem') document.all(i).className='FolderItem';
			}
			Obj.className='FolderSelectItem';
			}catch(e){}
			 if (Hypothesized!='')
			 {
					TempPath=CurrPath;
					PathArticle=TempPath.substr(TempPath.indexOf(Hypothesized)+Hypothesized.length);
			}else{
					PathArticle=CurrPath;
				}
		    var vstr='';
			if (CurrPath=='/')	vstr=Hypothesized+PathArticle+file;
			else vstr=Hypothesized+PathArticle+'/'+file;
			<%If KS.Setting(97)="1" Then
			  response.write "var dstr='" & KS.GetDomain & "';" & vbcrlf
			%>
			 if (vstr.substr(0,1)=='/'){
			  vstr=vstr.slice(1);
			 }
			 vstr=dstr+vstr;
			<%End If%>
			parent.document.getElementById('FileUrl').value=vstr;
			
		}
		function PreviewFile(File)
		{  
			var Url='';
			var Path='';
			if (CurrPath=='/') Path=escape(CurrPath+File);
			else Path=escape(CurrPath+'/'+File);
			Url="Preview.asp?FilePath="+escape(Path);
			parent.frames["PreviewArea"].location=Url;
		}
		function OpenParentFolder()
		{
			if (CanBackFlag==0) return;
			   location.href='FolderFileList.asp?CurrPath='+ParentPath;
			  SearchOptionExists(parent.document.all.FolderSelectList,ParentPath);
		}
		</script>
		<%
		Response.Write "<link href=""Admin_Style.CSS"" rel=""stylesheet"">"
		Response.Write "<body topmargin=""0"" leftmargin=""0"" onClick=""SelectFolder();"">"
		Response.Write "<table width=""99%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""0"">"
		Response.Write "  <tr>"
		Response.Write "    <td width=""70%"" class=""sort""> <div align=""center"">文件/文件夹名</div></td>"
		Response.Write "    <td width=""30%"" class=""sort""> <div align=""center"">大小</div></td>"
		Response.Write "  </tr>"
		 
		   If (CurrPath <> InstallDir & Left(UpFilesDir, Len(UpFilesDir) - 1)) And (ParentPath <> "") And Session("CurrPath") <> CurrPath Then
			 CanBackFlag = 1 '设置状态为可返回
		  
		Response.Write "  <tr title=""上级目录" & ParentPath & """>"
		Response.Write "    <td><table width=""117"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "        <tr>"
		Response.Write "          <td width=""24""><font color=""#FFFFFF""><img src=""../Images/arrow.gif""></font></td>"
		Response.Write "          <td><span style=""cursor:default"" onClick=""SelectUpFolder(this);"" onDblClick=""OpenParentFolder();"">返回上级目录</span></td>"
		Response.Write "        </tr>"
		Response.Write "      </table></td>"
		Response.Write "    <td></td>"
		Response.Write "  </tr>"
		
		 Else
		 CanBackFlag = 0  '设置状态为不可返回
		End If
		For Each FsoItem In SubFolderObj
		
		Response.Write "  <tr>"
		Response.Write "    <td width=""30%""><table border=""0"" cellspacing=""0"" cellpadding=""0"">"
		Response.Write "        <tr title=""双击鼠标进入此目录"">"
		Response.Write "          <td  width=""24""><img src=""../images/Folder/folderclosed.gif""></td>"
		Response.Write "         <td> <span class=""FolderItem"" Path=""" & FsoItem.name & """ onDblClick=""OpenFolder('" & FsoItem.name & "');"">"
		Response.Write FsoItem.name
		Response.Write "            </span> </td>"
		Response.Write "        </tr>"
		Response.Write "      </table></td>"
		Response.Write "    <td><div align=""Right"">"
		Response.Write FsoItem.size
		Response.Write "      字节 </div></td>"
		Response.Write "  </tr>"
		
		Next
		For Each FsoItem In FileObj
			FileExtName = LCase(Mid(FsoItem.name, InStrRev(FsoItem.name, ".") + 1))

		
		Response.Write "  <tr>"
		Response.Write "    <td width=""30%""><table border=""0"" cellspacing=""0"" cellpadding=""0"">"
		Response.Write "        <tr title=""双击鼠标进入此目录"">"
		Response.Write "          <td  width=""24""><img src='"&InstallDir&"Editor/KSPlus/FileIcon/"&GetFileIcon(FsoItem.name)&"' border=0 width=""16"" height=""16"" align=""absmiddle"" alt='"& FsoItem.Type&"'</td>"
		Response.Write "         <td> <span class=""FolderItem"" File=""" & FsoItem.name & """ onDblClick=""SetFile('" & replace(FsoItem.name,"'","\'") &"');"" onClick=""SelectFile(this,'" & replace(FsoItem.name,"'","\'") &"');"">"
		Response.Write FsoItem.name
		Response.Write "            </span> </td>"
		Response.Write "        </tr>"
		Response.Write "      </table></td>"
		Response.Write "    <td><div align=""Right"">"
		Response.Write FsoItem.size
		Response.Write "        &nbsp;字节 </div></td>"
		Response.Write "  </tr>"
		
		Next
		
		Response.Write "</table>"
		Response.Write "</body>"
		Response.Write "</html>"
		
		If Err Then
		 Response.Write ("<Script> alert('系统检测到网站信息配置有误,请检查!');window.close()</Script>")
		 Err.Clear
		End If
		Set FsoObj = Nothing
		Set SubFolderObj = Nothing
		Set FileObj = Nothing
		
		Response.Write "<script language=""JavaScript"">" & vbCrLf
		Response.Write "var CurrPath='" & CurrPath & "';" & vbCrLf
		Response.Write "var Hypothesized='" & Hypothesized & "';" & vbCrLf
		Response.Write "var ShowVirtualPath='" & ShowVirtualPath & "';" & vbCrLf
		Response.Write "var ParentPath='" & ParentPath & "';" & vbCrLf
		Response.Write "var CanBackFlag=" & CanBackFlag & ";" & vbCrLf
		Response.Write "</script>"
		%>
		<script>
		
		var SelectedObj=null;
		var DocMenuArr=new Array();
		DocElementArrInitialFlag=false;
		var DocPopupContextMenu=window.createPopup();
		document.oncontextmenu=new Function("return ShowMouseRightMenu(window.event);");
		function ShowMouseRightMenu(event)
		{
			DocDisabledContextMenu();
			var width=100;
			var height=0;
			var lefter=event.clientX;
			var topper=event.clientY;
			var ObjPopDocument=DocPopupContextMenu.document;
			var ObjPopBody=DocPopupContextMenu.document.body;
			var MenuStr='';
			for (var i=0;i<DocMenuArr.length;i++)
			{
				if (DocMenuArr[i].ExeFunction=='seperator')
				{
					MenuStr+=FormatSeperator();
					height+=16;
				}
				else
				{
					MenuStr+=FormatMenuRow(DocMenuArr[i].ExeFunction,DocMenuArr[i].Description,DocMenuArr[i].EnabledStr);
					height+=20;
				}
			}
			MenuStr="<TABLE border=0 cellpadding=0 cellspacing=0 class=Menu width=100>"+MenuStr
			MenuStr=MenuStr+"<\/TABLE>";
			ObjPopDocument.open();
			ObjPopDocument.write("<head><link href=\"ContextMenu.css\" type=\"text/css\" rel=\"stylesheet\"></head><body scroll=\"no\" onConTextMenu=\"event.returnValue=false;\" onselectstart=\"event.returnValue=false;\">"+MenuStr);
			ObjPopDocument.close();
			height+=4;
			if(lefter+width > document.body.clientWidth) lefter=lefter-width;
			DocPopupContextMenu.show(lefter, topper, width, height, document.body);
			return false;
		}
		function FormatSeperator()
		{
			var MenuRowStr="<tr><td height=16 valign=middle><hr><\/td><\/tr>";
			return MenuRowStr;
		}
		function FormatMenuRow(MenuOperation,MenuDescription,EnabledStr)
		{
			var MenuRowStr="<tr "+EnabledStr+"><td align=left height=20 class=MouseOut onMouseOver=this.className='MouseOver'; onMouseOut=this.className='MouseOut'; valign=middle"
			if (EnabledStr=='') MenuRowStr+=" onclick=\""+MenuOperation+"parent.DocPopupContextMenu.hide();\">&nbsp;&nbsp;&nbsp;&nbsp;";
			else MenuRowStr+=">&nbsp;&nbsp;&nbsp;&nbsp;";
			MenuRowStr=MenuRowStr+MenuDescription+"<\/td><\/tr>";
			return MenuRowStr;
		}
		$(document).ready(function(){
			if (DocElementArrInitialFlag) return;
			InitialDocMenuArr();
			DocElementArrInitialFlag=true;
		});
		function ContextMenuItem(ExeFunction,Description,EnabledStr)
		{
			this.ExeFunction=ExeFunction;
			this.Description=Description;
			this.EnabledStr=EnabledStr;
		}
		function DocDisabledContextMenu()
		{
			SelectFolder();
		}
		function InitialDocMenuArr()
		{
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.AddFolderOperation();",'新建目录','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("if (confirm('确定要删除吗？')==true) parent.DelFolderFile();",'删除','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.EditFolder();",'重命名','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.location.reload();",'刷新页面','disabled');
		}
		function SelectFolder()
		{
			Obj=event.srcElement,DisabledContextMenuStr='';
			if (SelectedObj!=null) SelectedObj.className='FolderItem';
			if ((Obj.Path!=null)||(Obj.File!=null))
			{
				Obj.className='FolderSelectItem';
				SelectedObj=Obj;
			}
			else SelectedObj=null;
			if (SelectedObj!=null)	DisabledContextMenuStr='';
			else DisabledContextMenuStr=',删除,重命名,';
			for (var i=0;i<DocMenuArr.length;i++)
			{
				if (DisabledContextMenuStr.indexOf(DocMenuArr[i].Description)!=-1) DocMenuArr[i].EnabledStr='disabled';
				else  DocMenuArr[i].EnabledStr='';
			}
		}
		
		function SelectUpFolder(Obj)
		{
			for (var i=0;i<document.all.length;i++)
			{
				if (document.all(i).className=='FolderSelectItem') document.all(i).className='FolderItem';
			}
			Obj.className='FolderSelectItem';
		}
		
		window.onunload=CheckReturnValue;
		function CheckReturnValue()
		{
			if (typeof(window.returnValue)!='string') window.returnValue='';
		}
		function AddFolderOperation()
		{  
			var ReturnValue=prompt('请输入新建目录名称：','');
			if ((ReturnValue!='') && (ReturnValue!=null))
				window.location.href='FolderFileList.asp?Type=AddFolder&Path='+CurrPath+'/'+ReturnValue+'&CurrPath='+CurrPath;
		}
		function DelFolderFile()
		{
			if (SelectedObj!=null)
			{
				if (SelectedObj.Path!=null) window.location.href='?Type=DelFolder&Path='+CurrPath+'/'+SelectedObj.Path+'&CurrPath='+CurrPath;
				if (SelectedObj.File!=null) window.location.href='?Type=DelFile&Path='+CurrPath+'&FileName='+SelectedObj.File+'&CurrPath='+CurrPath;
			}
			else alert('请选择要删除的目录');
		}
		function EditFolder()
		{
			var ReturnValue='';
			if (SelectedObj!=null)
			{
				if (SelectedObj.Path!=null)
				{
					ReturnValue=prompt('修改的名称：',SelectedObj.Path);
					if ((ReturnValue!='') && (ReturnValue!=null)) window.location.href='?Type=FolderReName&Path='+CurrPath+'&CurrPath='+CurrPath+'&OldPathName='+SelectedObj.Path+'&NewPathName='+ReturnValue;
				}
				if (SelectedObj.File!=null)
				{
					ReturnValue=prompt('修改的名称：',SelectedObj.File);
					if ((ReturnValue!='') && (ReturnValue!=null)) window.location.href='?Type=FileReName&Path='+CurrPath+'&CurrPath='+CurrPath+'&OldFileName='+SelectedObj.File+'&NewFileName='+ReturnValue;
				}
			}
			else alert('请填写要更名的目录名称');
		}
		</script>
		<%
		End Sub
		Function CheckFileShowTF(AllowShowExtNameStr, ExtName)
			If ExtName = "" Then
				CheckFileShowTF = False
			Else
				If InStr(1, AllowShowExtNameStr, ExtName) = 0 Then
					CheckFileShowTF = False
				Else
					CheckFileShowTF = True
				End If
			End If
		End Function
End Class
%> 
