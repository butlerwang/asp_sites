<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- #include file="../inc/access.asp" -->
<!-- #include file="inc/functions.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="css/common.css" type="text/css" />
<title>管理区域</title>
</head>
<body>
<div id="man_zone">
<%
'====================系统空间占用=======================
sub SpaceSize()
On error resume next
GetSysInfo()
Dim t
't = GetAllSpace
Dim FoundFso
FoundFso = False
FoundFso = IsObjInstalled("Scripting.FileSystemObject")
%>

  <table width="95%" border="0" align="center"  cellpadding="3" cellspacing="1" class="table_style">
     <tr>
      <td colspan="2"  >&nbsp;服务器相关信息</td>
    </tr> 
    <tr>
      <td width="18%" class="left_title_1"><span class="left-title">网站域名</span></td>
      <td width="82%">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
    </tr>
    <tr>
      <td class="left_title_2">网站IP地址</td>
      <td>&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
    </tr>
    <tr>
      <td class="left_title_1">运行端口</td>
      <td>&nbsp;<%=Request.ServerVariables("server_port")%></td>
    </tr>
    <tr>
      <td class="left_title_2">ASP脚本解释引擎</td>
      <td>&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
    </tr>
    <tr>
      <td class="left_title_1">IIS 版本</td>
      <td>&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%> </td>
    </tr>
    <tr>
      <td class="left_title_2">服务器操作系统</td>
      <td>&nbsp;<%=Request.ServerVariables("OS")%></td>
    </tr>
    <tr>
      <td class="left_title_1">服务器CPU数量</td>
      <td>&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%>个</td>
    </tr>
    <tr>
      <td colspan="2"  >&nbsp;主要组件信息</td>
    </tr>
    <tr>
      <td class="left_title_1">FSO文件读写</td>
      <td>&nbsp;<%
If FoundFso Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td class="left_title_2">Jmail发送邮件支持</td>
      <td>&nbsp;<%
If IsObjInstalled("JMail.SmtpMail") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%></td>
    </tr>
    <tr>
      <td class="left_title_1">CDONTS发送邮件支持</td>
      <td>&nbsp;<%
If IsObjInstalled("CDONTS.NewMail") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td class="left_title_2">AspEmail发送邮件支持</td>
      <td>&nbsp;<%
If IsObjInstalled("Persits.MailSender") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%></td>
    </tr>
    <tr>
      <td class="left_title_1">无组件上传支持</td>
      <td>&nbsp;<%
If IsObjInstalled("Adodb.Stream") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td class="left_title_2">AspUpload上传支持</td>
      <td>&nbsp;<%
If IsObjInstalled("Persits.Upload") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%></td>
    </tr>    
    <tr>
      <td class="left_title_1">SA-FileUp上传支持</td>
      <td>&nbsp;<%
If IsObjInstalled("SoftArtisans.FileUp") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td class="left_title_2">DvFile-Up上传支持</td>
      <td>&nbsp;<%
If IsObjInstalled("DvFile.Upload") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%></td>
    </tr>
    <tr>
      <td class="left_title_1">CreatePreviewImage生成图片</td>
      <td>&nbsp;<%
If IsObjInstalled("CreatePreviewImage.cGvbox") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td class="left_title_2">AspJpeg生成预览图片</td>
      <td>&nbsp;<%
If IsObjInstalled("Persits.Jpeg") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%></td>
    </tr>  
    <tr>
      <td class="left_title_1">SA-ImgWriter生成预览图片</td>
      <td>&nbsp;<%
If IsObjInstalled("SoftArtisans.ImageGen") Then
	Response.Write "<font color=green><b>√</b></font>"
Else
	Response.Write "<font color=red><b>×</b></font>"
End If
%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td class="left_title_2">ADO(数据库访问)版本</td>
      <td>&nbsp;<%=cn.Version%></td>
    </tr>
    <tr>
      <td class="left_title_1">磁盘文件操作速度测试</td>
      <td>&nbsp;<%
	Response.Write "正在重复创建、写入和删除文本文件50次..."

	Dim thetime3,tempfile,iserr,t1,FsoObj,tempfileOBJ,t2
	Set FsoObj=Server.CreateObject("Scripting.FileSystemObject")

	iserr=False
	t1=timer
	tempfile=server.MapPath("./") & "\aspchecktest.txt"
	For i=1 To 50
		Err.Clear

		Set tempfileOBJ = FsoObj.CreateTextFile(tempfile,true)
		If Err <> 0 Then
			Response.Write "创建文件错误！"
			iserr=True
			Err.Clear
			Exit For
		End If
		tempfileOBJ.WriteLine "Only for test. Ajiang ASPcheck"
		If Err <> 0 Then
			Response.Write "写入文件错误！"
			iserr=True
			Err.Clear
			Exit For
		End If
		tempfileOBJ.close
		Set tempfileOBJ = FsoObj.GetFile(tempfile)
		tempfileOBJ.Delete 
		If Err <> 0 Then
			Response.Write "删除文件错误！"
			iserr=True
			Err.Clear
			Exit For
		end if
		Set tempfileOBJ=Nothing
	Next
	t2=timer
	If Not iserr Then
		thetime3=cstr(int(( (t2-t1)*10000 )+0.5)/10)
		Response.Write "...已完成！本服务器执行此操作共耗时 <font color=red>" & thetime3 & " 毫秒</font>"
	End If
%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td class="left_title_2">ASP脚本解释和运算速度测试</td>
      <td>&nbsp;<%

	Response.Write "整数运算测试，正在进行50万次加法运算..."
	dim lsabc,thetime,thetime2
	t1=timer
	for i=1 to 500000
		lsabc= 1 + 1
	next
	t2=timer
	thetime=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...已完成！共耗时 <font color=red>" & thetime & " 毫秒</font><br>"


	Response.Write "浮点运算测试，正在进行20万次开方运算..."
	t1=timer
	for i=1 to 200000
		lsabc= 2^0.5
	next
	t2=timer
	thetime2=cstr(int(( (t2-t1)*10000 )+0.5)/10)
	Response.Write "...已完成！共耗时 <font color=red>" & thetime2 & " 毫秒</font><br>"
%></td>
    </tr>                  
  </table>
  <%
end sub


'=====================系统空间参数=========================
Sub ShowSpaceInfo(drvpath)
	dim fso,d,size,showsize
	set fso=server.createobject("scripting.filesystemobject") 		
	drvpath=server.mappath(drvpath) 		 		
	set d=fso.getfolder(drvpath) 		
	size=d.size
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write "<font face=verdana>" & showsize & "</font>"
End Sub	
 	
Sub Showspecialspaceinfo(method)
	dim fso,d,fc,f1,size,showsize,drvpath 		
	set fso=server.createobject("scripting.filesystemobject")
	drvpath=server.mappath("../index.asp")
	drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
	set d=fso.getfolder(drvpath)
	if method="All" then 		
		size=d.size
	elseif method="Program" then
		set fc=d.Files
		for each f1 in fc
			size=size+f1.size
		next	
	end if
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write "<font face=verdana>" & showsize & "</font>"
end sub 	 	 	
	
Function Drawbar(drvpath)
	dim fso,drvpathroot,d,size,totalsize,barsize
	set fso=server.createobject("scripting.filesystemobject")
	drvpathroot=server.mappath("../index.asp")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	drvpath=server.mappath(drvpath)
	if fso.FolderExists(drvpath) then		
		set d=fso.getfolder(drvpath)
		size=d.size
	End If
	barsize=cint((size/totalsize)*400)
	Drawbar=barsize
End Function 	
 	
Function Drawspecialbar()
	dim fso,drvpathroot,d,fc,f1,size,totalsize,barsize
	set fso=server.createobject("scripting.filesystemobject")
	drvpathroot=server.mappath("../index.asp")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	set fc=d.files
	for each f1 in fc
		size=size+f1.size
	next
	barsize=cint((size/totalsize)*400)
	Drawspecialbar=barsize
End Function
	
Function GetAllSpace()
	Dim fso,drvpath,d,size
	set fso=server.createobject("scripting.filesystemobject")
	drvpath=server.mappath("../index.asp")
	drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
	set d=fso.getfolder(drvpath)	
	size=d.size
	set fso=nothing
	GetAllSpace = size
End Function

Function GetFileSize(FileName)
	Dim fso,drvpath,d,size,showsize
	set fso=server.createobject("scripting.filesystemobject")
	drvpath=server.mappath(FileName)
	set d=fso.getfile(drvpath)	
	size=d.size
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(Size/1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	set fso=nothing
	GetFileSize = showsize
End Function

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

Sub GetSysInfo()
	On Error Resume Next
	Dim WshShell,WshSysEnv
	Set WshShell = Server.CreateObject("WScript.Shell")
	Set WshSysEnv = WshShell.Environment("SYSTEM")
	okOS = Cstr(WshSysEnv("OS"))
	okCPUS = Cstr(WshSysEnv("NUMBER_OF_PROCESSORS"))
	okCPU = Cstr(WshSysEnv("PROCESSOR_IDENTIFIER"))
	If IsNull(okCPUS) Then
		okCPUS = Request.ServerVariables("NUMBER_OF_PROCESSORS")
	ElseIf okCPUS="" Then
		okCPUS = Request.ServerVariables("NUMBER_OF_PROCESSORS")
	End If
	If Request.ServerVariables("OS")="" Then okOS=okOS & "(可能是 Windows Server 2003)"
End Sub



Call SpaceSize()
%>
</div>
</body>
</html>
