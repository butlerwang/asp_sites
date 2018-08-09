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
      <td class="left_title_2">ADO(数据库访问)版本</td>
      <td>&nbsp;<%=cn.Version%></td>
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
