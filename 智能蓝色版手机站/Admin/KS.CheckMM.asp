<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.FileIcon.asp"-->
<!--#include file="Include/Session.asp"-->
<%


Const FilterFiles="Kesion.CommonCls.asp,Kesion.Label.SQLCls.asp,Kesion.IfCls.asp,Collect_ItemModify4.asp,Comment.asp,Mood.asp,User_PayReceive.asp,KS_Char.asp,Wap_FilesCls.asp,MyFunction.asp,Upload.asp,User_Photo.asp,Alipay_NotifyUrl.asp,KS.Template.asp,Upfilesave.asp,Kesion.UpFileCls.asp,ex.asp,user_files.asp,user_blog.asp,qqBind.asp,Kesion.Label.FunctionCls.asp,Kesion.Thumbs.asp,rnd.asp,cls_api.asp,KS.Shop.asp,cls_api.asp,Kesion.VersionCls.asp,function.asp"  '定义过滤不检测的文件,多个文件用逗号隔开
dim Report,delnum:delnum=0
Dim KS:Set KS=New PublicCls
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>在线找木马</title>
<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../KS_Inc/common.js"></script></head>
<script language="JavaScript" src="../KS_Inc/jquery.js"></script></head>
<body scroll="no" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div class='topdashed sort'>在线检测木马</div>
<div style="height:95%; overflow: auto; width:100%" align="center">
<%
	if KS.G("act")<>"scan" then
%>
				<form action="?act=scan" method="post">

 <table style="margin-top:4px" width="99%" align="center" class="Ctable" border="0" cellpadding="0" cellspacing="0">
				
				<tr class="tdbg">       
				 <td height="28" width="150"align="right" class="clefttitle"><strong>检查的路径：</strong></td>              <td><input name="path" type="text" style="border:1px solid #999" value="\" size="30" />
				 * 网站根目录的相对路径，填“\”即检查整个网站
				 </td>    
				</tr>
				<tr class="tdbg">
				 <td height="28" width="150" align="right" class="clefttitle"><strong>检查的扩展名：</strong></td>              <td><input name="FileExt" type="text" style="border:1px solid #999" value="asp,asa,gif,jpg" size="30" />
				 *多个扩展名,请用逗号隔开
				 </td>    
				</tr>
				<tr class="tdbg">
				  <td height="28" align="right" class="clefttitle"><strong>以下扩展名的文件将被直接删除：</strong></td>
				  <td><input type="text" name="delfilelist" id="delfilelist" value="cdx,asa,cer" size='30'> <font color=red>说明KesionCMS系统默认是不包含以上类型的文件,一般情况下如果您的网站多出以上类型的文件,考虑是被上传了木马,如果不删除请留空。
				  </td>
                </tr>
				<tr class="tdbg">
				  <td height="28" align="right" class="clefttitle"><strong>高危险文件处理：</strong></td>
				  <td>&nbsp;
				  <input type="radio" name="delfile" value="1">直接删除
				  <input type="radio" name="delfile" value="0" checked>提示我删除
				  </td>
                </tr>
				</table>
				
		     <div style="text-align:center;margin-top:20px">
				<input type="submit" value=" 开始扫描 " onClick="if($('#delfilelist').val()!=''){return(confirm($('#delfilelist').val()+'的文件将被删除,确定开始扫描吗?'))}" class="button" />
			 </div>
				</form>
			<div style="line-height:24px;text-align:left;padding:10px;background:#ffffee;margin:5px 2px;border:1px #f9c943 solid">
			 使用说明:<br />
			  ①、执行本操作需要耗费几分钟时间，请在访问量少的时候执行本操作。
			  <br />
			  ②、您可以用此工具检查您的站点是否存在木马文件。
			  <br />③、由于新的木马变种经常变化出现,本工具不保证所有木马都可以查出来。
			</div>
<%
	else
	%>
	<br>
	<table border="0" width="98%" align="center">
	<tr>
	 <td id='message' style="line-height:24px;padding:10px;background:#ffffee;border:1px #f9c943 solid">   
		<table border="0" width="100%">
		<tr>
		 <td width="150" height="50"><img src="images/wait.gif"></td>
		 <td id="msg"></td>
		</tr>
		</table>
	</td>
   </tr>
  </table>
<%
		server.ScriptTimeout = 90000
		DimFileExt = Request("FileExt")
		If DimFileExt="" Then DimFileExt="asp"
		delfilelist= Request("delfilelist")
		If delfilelist="" Then delfilelist="0"
		If delfilelist="0" Then DimFileExt=DimFileExt & ",cdr,asa,cer"   '如果没有直接删除，加入检测列表
		
		Sun = 0
		SumFiles = 0
		SumFolders = 1
		if request.Form("path")="" then
			response.Write("<script>alert('请输入要检测的路径!');history.back();</script>")
			response.End()
		end if
		timer1 = timer
		if request.Form("path")="\" then
			TmpPath = Server.MapPath("\")
		elseif request.Form("path")="." then
			TmpPath = Server.MapPath(".")
		else
			TmpPath = Server.MapPath("\")&"\"&request.Form("path")
		end if
		Call ShowAllFile(TmpPath)
		
		Dim Msg
		If Sun=0 Then
		 Msg="<img src=images/succeed.gif align=absmiddle> 扫描完毕！恭喜您,您的系统很安全,没有发现木马,希望您能保持！"
		Else
		 Msg="<img src=images/succeed.gif align=absmiddle> 扫描完毕！一共检查文件夹<font color=""#FF0000"">" & SumFolders & "</font>个，文件<font color=""#FF0000"">" & SumFiles & "</font>个，发现可疑点<font color=""#FF0000"">" & Sun & "</font>个,删除危险文件<dont color=red>" & delnum & "</font>个"
		End If
		Response.Write "<script>message.innerHTML='" & MSG & "';</script>"
		Response.Flush()


%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">

  <tr>
    <td class="CPanel" style="padding:5px;line-height:170%;clear:both;font-size:12px">
       
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
	 <tr>
		 <td valign="top">
		  <% If Sun<>0 Then%>
			 <table width="99%" align="center" border="1" cellpadding="0" cellspacing="0" style="padding:5px;line-height:170%;clear:both;font-size:12px">
			 <tr>
			   <td>文件相对路径</td>
			   <td>特征码</td>
			   <td width="230">描述</td>
			   <td>创建/修改时间</td>
			   <td>处理情况</td>
			   </tr>
			   <tbody id='tablemsg'>
			   </tbody>
			 <%=Report%>
			 </table>
		 <%end if%>	 
			 </td>
	 </tr>
	</table>
</td></tr></table>
<%
	timer2 = timer
	thetime=cstr(int(((timer2-timer1)*10000 )+0.5)/10)
	response.write "<br><font size=""2"">本次检测共用了"&thetime&"毫秒</font>"
end if

%>
<br />
</body>
</html>
<%

'遍历处理path及其子目录所有文件
Sub ShowAllFile(Path)
	Set FSO = CreateObject("Scripting.FileSystemObject")
	if not fso.FolderExists(path) then exit sub
	Set f = FSO.GetFolder(Path)
	Set fc2 = f.files
	For Each myfile in fc2
	     ext=FSO.GetExtensionName(path&"\"&myfile.name)
		If CheckExt(ext,DimFileExt) or CheckExt(FSO.GetExtensionName(path&"\"&myfile.name),delfilelist) Then
			Response.Write "<script>msg.innerHTML='正在检测文件:" & myfile.name & "';</script>"
		    If KS.FoundInArr(lcase(FilterFiles),lcase(myfile.name),",")=false Then
			 Call ScanFile(Path&Temp&"\"&myfile.name, "",ext)
			End If
			SumFiles = SumFiles + 1
			Response.Flush()
		End If
	Next
	Set fc = f.SubFolders
	For Each f1 in fc
	     if instr(lcase(f1.name),".")<>0 Then
			Report = Report&"<tr><td>"&replace(path&"\"&f1.name,server.MapPath("\")&"\","",1,1,1)&"</td><td>不合法的文件夹名称</td><td colspan=2>危险文件夹，一般利用IIS的文件名执行漏洞,形式为 x.asp下的图片木马可能被执行</td><td><font color=red>请手工删除该文件夹</font></td></tr>"
		 End If
		  ShowAllFile path&"\"&f1.name
		  SumFolders = SumFolders + 1
    Next
	Set FSO = Nothing
End Sub

'检测文件
Sub ScanFile(FilePath, InFile,ext)
	If InFile <> "" Then
		Infiles = "该文件被<a href=""http://"&Request.Servervariables("server_name")&"\"&InFile&""" target=_blank>"& InFile & "</a>文件包含执行"
	End If
	 temp = "<a href=""http://"&Request.Servervariables("server_name")&"\"&replace(FilePath,server.MapPath("\")&"\","",1,1,1)&""" target=_blank>"&replace(FilePath,server.MapPath("\")&"\","",1,1,1)&"</a>"

	if instr(FilePath,";")<>0 then
		Report = Report&"<tr><td>"&temp&"</td><td>不合法的文件名</td><td>危险文件，一般利用IIS的文件名执行漏洞,形式为 *.asp;*.gif等</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td>" & DeleteFile(filepath) & "</td></tr>"
		Sun = Sun + 1
		exit sub
	end if
	if CheckExt(ext,delfilelist) Then
		Report = Report&"<tr><td>"&temp&"</td><td>非法扩展名</td><td>危险文件，非KesionCMS系统文件！</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td>" & DeleteFile(filepath) & "</td></tr>"
		Sun = Sun + 1
	  exit sub
	End If
	
    	
	Set FSOs = CreateObject("Scripting.FileSystemObject")
	on error resume next
	set ofile = fsos.OpenTextFile(FilePath)
	filetxt = Lcase(ofile.readall())
	If err Then Exit Sub end if
	if len(filetxt)>0 then
		    '特征码检查
			'Check "WScr"&DoMyBest&"ipt.Shell"
			If instr( filetxt, Lcase("WScr"&DoMyBest&"ipt.Shell") ) or Instr( filetxt, Lcase("clsid:72C24DD5-D70A"&DoMyBest&"-438B-8A42-98424B88AFB8") ) then
				Report = Report&"<tr><td>"&temp&"</td><td>WScr"&DoMyBest&"ipt.Shell 或者 clsid:72C24DD5-D70A"&DoMyBest&"-438B-8A42-98424B88AFB8</td><td>危险组件，一般被ASP木马利用。"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>请手工确认</font></td></tr>"
				Sun = Sun + 1
			End if
			'Check "She"&DoMyBest&"ll.Application"
			If instr( filetxt, Lcase("She"&DoMyBest&"ll.Application") ) or Instr( filetxt, Lcase("clsid:13709620-C27"&DoMyBest&"9-11CE-A49E-444553540000") ) then
				Report = Report&"<tr><td>"&temp&"</td><td>She"&DoMyBest&"ll.Application 或者 clsid:13709620-C27"&DoMyBest&"9-11CE-A49E-444553540000</td><td>危险组件，一般被ASP木马利用。"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>请手工确认</font></td></tr>"
				Sun = Sun + 1
			End If
			'Check .Encode
			Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True
			regEx.Pattern = "@\s*LANGUAGE\s*=\s*[""]?\s*(vbscript|jscript|javascript).encode\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>(vbscript|jscript|javascript).Encode</td><td>似乎脚本被加密了，一般ASP文件是不会加密的。"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>请手工确认</font></td></tr>"
				Sun = Sun + 1
			End If
			'Check my ASP backdoor :(
			regEx.Pattern = "\bEv"&"al\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>Ev"&"al</td><td>e"&"val()函数可以执行任意ASP代码，被一些后门利用。其形式一般是：ev"&"al(X)<br>但是javascript代码中也可以使用，有可能是误报。"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>请手工确认</font></td></tr>"
				Sun = Sun + 1
			End If
			'Check exe&cute backdoor
			regEx.Pattern = "[^.]\bExe"&"cute\b"
			If regEx.Test(filetxt) and instr(filetxt,"conn.execute")=0 Then
				Report = Report&"<tr><td>"&temp&"</td><td>Exec"&"ute</td><td>e"&"xecute()函数可以执行任意ASP代码，被一些后门利用。其形式一般是：ex"&"ecute(X)。<br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>请手工确认</font></td></tr>"
				Sun = Sun + 1
			End If
			
			'===================10-31号增加=========================
			dim findcontent:findcontent=lcase(filetxt)
			if (instr(findcontent,"exec"&"utestatement")<>0 or instr(findcontent,"msscript"&"control.scriptcontr")<>0 or instr(findcontent,"clsid:72c24dd5-d70"&"a-438b-8a42-98424b88afb8")<>0 or instr(findcontent,"clsid:f935dc22-1cf0-11d0-adb9"&"-00c04fd58a0b")<>0 or instr(findcontent,"clsid:093ff999-1ea0-4079-9525-961"&"4c3504b74")<>0 or instr(findcontent,"clsid:f935dc26-1cf0-11d0-adb9-"&"00c04fd58a0b")<>0 or instr(findcontent,"clsid:0d43fe01"&"-f093-11cf-8940-00a0c9054228")<>0) then
				Report = Report&"<tr><td>"&temp&"</td><td>Execute"&"Global/execute/clsid</td><td>Execute"&"Global()函数可以执行任意ASP代码，被一些后门利用。其形式一般是：Execute"&"Global(X)。<br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>请手工确认</font></td></tr>"
				Sun = Sun + 1
			End If
			
			
			regEx.Pattern = "execute\s*request"
			If regEx.Test(findcontent) and instr(lcase(filename),"scan.asp")=0 Then
				Report = Report&"<tr><td>"&temp&"</td><td>Execute"&"Global/execute</td><td>Execute"&"Global()函数可以执行任意ASP代码，被一些后门利用。其形式一般是：Execute"&"Global(X)。<br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>请手工确认</font></td></tr>"
				Sun = Sun + 1
			end if
			
			regEx.Pattern = "executeglobal\s*request"
			If regEx.Test(findcontent) and instr(lcase(filename),"scan.asp")=0 Then
				Report = Report&"<tr><td>"&temp&"</td><td>Execute"&"Global/execute</td><td>Execute"&"Global()函数可以执行任意ASP代码，被一些后门利用。其形式一般是：Execute"&"Global(X)。<br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>请手工确认</font></td></tr>"
				Sun = Sun + 1
			end if
			regEx.Pattern = "<script.*runat.*server(\n|.)*execute(\n|.)*<\/script>"
			If regEx.Test(findcontent) and instr(lcase(filename),"scan.asp")=0 Then
				Report = Report&"<tr><td>"&temp&"</td><td>Execute"&"Global/execute</td><td>Execute"&"Global()函数可以执行任意ASP代码，被一些后门利用。其形式一般是：Execute"&"Global(X)。<br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>请手工确认</font></td></tr>"
				Sun = Sun + 1
			end if

			Set regEx = Nothing
			
			
			'===================增强检测结束============================================
			
	 
			
		'Check include file
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "<!--\s*#include\s*file\s*=\s*"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			ext=FSOs.GetExtensionName(tFile)
			If Not CheckExt(ext,DimFileExt) Then
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1),ext )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check include virtual
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "<!--\s*#include\s*virtual\s*=\s*"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")     
			ext=FSOs.GetExtensionName(tFile)
			If Not CheckExt(ext,DimFileExt) Then
				Call ScanFile( Server.MapPath("\")&"\"&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1),ext )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check Server&.Execute|Transfer
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "Server.(Exec"&"ute|Transfer)([ \t]*|\()"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			ext=FSOs.GetExtensionName(tFile)
			If Not CheckExt(ext,DimFileExt) Then
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1),ext )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check Server&.Execute|Transfer
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "Server.(Exec"&"ute|Transfer)([ \t]*|\()[^""]\)"
		If regEx.Test(filetxt) Then
			Report = Report&"<tr><td>"&temp&"</td><td>Server.Exec"&"ute</td><td>不能跟踪检查Server.e"&"xecute()函数执行的文件。请管理员自行检查。<br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>请手工确认</font></td></tr>"
			Sun = Sun + 1
		End If
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check Crea"&"teObject
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "CreateO"&"bject[ |\t]*\(.*\)"
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			If Instr(Match.Value, "&") or Instr(Match.Value, "+") or Instr(Match.Value, """") = 0 or Instr(Match.Value, "(") <> InStrRev(Match.Value, "(") Then
				Report = Report&"<tr><td>"&temp&"</td><td>Creat"&"eObject</td><td>Crea"&"teObject函数使用了变形技术，仔细复查。"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>请手工确认</font></td></tr>"
				Sun = Sun + 1
				exit sub
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing

	end if
	set ofile = nothing
	set fsos = nothing
End Sub

'检查文件后缀，如果与预定的匹配即返回TRUE
Function CheckExt(FileExt,CheckFileExt)
	If DimFileExt = "*" Then CheckExt = True
	Ext = Split(CheckFileExt,",")
	For i = 0 To Ubound(Ext)
		If Lcase(FileExt) = Ext(i) Then 
			CheckExt = True
			Exit Function
		End If
	Next
End Function

Function GetDateModify(filepath)
	Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(filepath) 
	s = f.DateLastModified 
	set f = nothing
	set fso = nothing
	GetDateModify = s
End Function

Function GetDateCreate(filepath)
	Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(filepath) 
	s = f.DateCreated 
	set f = nothing
	set fso = nothing
	GetDateCreate = s
End Function

'删除文件
Public Function DeleteFile(FileStr)
       if request("delfile")="1" Then
		   Dim FSO
		   On Error Resume Next
		   Set FSO = CreateObject("Scripting.FileSystemObject")
			FSO.DeleteFile FileStr, True
		   Set FSO = Nothing
		   If Err.Number <> 0 Then
			Err.Clear
			DeleteFile="<font color=green>删除失败,请手工删除</font>"
		   Else
		   delnum=delnum+1
			DeleteFile="<font color=red>已删除</font>"
		   End If
	   else
	     DeleteFile="<font color=blue>请确认并手工删除</font>"
	   end if
End Function

Set KS=Nothing
CloseConn
%>