<%@ Language="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.Commoncls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../KS_Cls/Kesion.CollectCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
Response.Buffer=true
Server.ScriptTimeout=9999999
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"

'是否允许自动检测最新版本 true 允许 false 不允许
Dim EnabledAutoUpdate:EnabledAutoUpdate=true 
'网站程序使用的编码,一定要设置正确,否则可能导致网站出现乱码
const Encoding="utf-8"
'官方远程文件版本地址
const Kesion_Version_XmlUrl="#" 

'官方远程文件更新列表地址,必须/结束   
const Kesion_Update_FileUrl="#"   

Dim SuccNum,ErrNum,LocalVersion,RemoteVersion
Dim KS:Set KS=New PublicCls

Dim TempDownFileDir : TempDownFileDir = KS.Setting(3) & KS.Setting(91) & "update/"      '下载升级文件的临时目录,不需要修改
Dim xmlObj : set xmlObj = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)

'普通管理员屏蔽自动升级
If KS.C("SuperTF")<>"1" And EnabledAutoUpdate=true Then EnabledAutoUpdate=false

select case Request("action")
  case "check"  checkIsNewestVersion
  case "downfile" begindown
  case "update" beginupdate
  case "showupdateinfo" showupdateinfo
  case "showintro" showintro
end select

Set KS=Nothing
CloseConn

sub LoadRemoteXML()
	  on error resume next
	  xmlObj.load(server.mappath("include/version.xml"))
	  if isObject(xmlObj) then
		LocalVersion=xmlObj.getElementsByTagName("kesioncms/version")(0).Text
	  end if
	  if err.number<>0 then
		err.clear
		KS.Die "localversionerr"
	  end if
	 
    set xmlObj = KS.InitialObject("Microsoft.XMLDOM")
	xmlObj.async = "false"
	xmlObj.resolveExternals = "false"
	xmlObj.setProperty "ServerHTTPRequest", true
	xmlObj.load(Kesion_Version_XmlUrl&"?v=" &LocalVersion &"&b=" & IsBusiness & "&d=" &GetTrueDomain(Request.ServerVariables("SERVER_NAME")))
end sub

Function GetTrueDomain(domain)
				Dim x:x = split(domain,".")
				Dim sdomain:sdomain= ""
				Dim start:start = 2
				Dim k :k= 1
				if ubound(x)<=1 then GetTrueDomain=domain:exit function
				if (ubound(x) >= 3) then start = 3
				dim i:i=start
				do while i > 0
					if (i=start) then
						sdomain = sdomain & x(ubound(x)-start+k)
					else
						sdomain = sdomain & "." & x(ubound(x)-start+k)
					end if
					k=k+1
					i=i-1
				loop
				GetTrueDomain=sdomain
End function

function checkIsNewestVersion()
    If EnabledAutoUpdate=false Then KS.Die "enabled"
    Call LoadRemoteXML()
	if xmlObj.readystate=4 and xmlObj.parseError.errorCode=0 Then 
	  Dim Node,XMLNode:Set XMLNode=XmlObj.getElementsByTagName("root/item")
	  if XMLNode.length>0 then
	   RemoteVersion=XMLNode.item(XMLNode.length-1).SelectSingleNode("version").text
	  end if
	else
	   KS.Die "remoteversionerr"
	end if
  
  if RemoteVersion>LocalVersion Then
    'If xmlObj.getElementsByTagName("root/item/allowupdateonline")(0).Text="false"Then   
	' KS.Echo "unallow"
	'ElseIf KS.ChkClng(split(RemoteVersion,".")(0))>KS.ChkCLng(split(LocalVersion,".")(0)) Then '增加判断是不是同一版本号的
	' KS.Echo "unallowversion"
	'End If
  else
    KS.Echo "false"
  end if
end function

sub showupdateinfo()
 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
<link href='Include/admin_style.css' rel='stylesheet'>
<script language='JavaScript' src='../KS_Inc/Common.js'></script>
<script language='JavaScript' src='../KS_Inc/Jquery.js'></script>
<style type="text/css">
.listinfo li{height:23px;line-height:23px;background:url(images/37.gif) no-repeat 0px 7px;padding-left:10px}
.bborder{border:2px solid #cccccc }
</style>
<script type="text/javascript">
 function check(){
  if ($('#isupdate').attr("checked")){
   return true;
  }else{
   alert('您没有选择需要在线升级的补丁!');
   return false;
  }
 }
</script>
<body class="tdbg">
  <table border="0" style="margin-top:20px" width="95%" align="center" cellspacing="0" cellpadding="0">
   <form name="myform" id="myform" action="KS.Update.asp" method="post">
  <tr>
   <td class="listinfo">
   <input type="hidden" name="action" value="downfile"/>
 <%
 If EnabledAutoUpdate=false Then KS.Die "对不起，您没有开启在线升级功能"

  
    Call LoadRemoteXML()
	Dim Num,Node,XMLNode,showupdatebutton:showupdatebutton=true
	num=1
	if xmlObj.readystate=4 and xmlObj.parseError.errorCode=0 Then 
	  Set XMLNode=XmlObj.getElementsByTagName("root/item")
	  response.write "<table width='99%' align='center'>"
	  For Each Node In XMLNode
	    If lcase(Node.SelectSingleNode("authorize").text)="false" Then
		    showupdatebutton=false
			response.write "<tr class='tdbg'>" & vbcrlf
			response.write " <td class='splittd' style='color:blue;font-size:14px;height:60px;line-height:32px' colspan=6><strong>温馨提示:</strong><br/><font color=red>1、系统检测到有新的可升级补丁，但发现您没有安装在授权域名下，商业授权版本请安装在授权域名下方可使用在线升级功能;"
			response.write "<br/>2、如需要手工升级，请登录官方<a style='color:red' href=><u>商业用户自动中心</u></a>下载补丁包;"
			response.write "<br/>3、对授权有任何疑问，请登录官方网站<a style='color:red' href='#/sysq/sqcx.html' target='_blank'><u>授权查询中心</u></a>查询;"
			response.write "</font></td></tr>"
			exit for
	    ElseIf lcase(Node.SelectSingleNode("authorize").text)="expire" Then
		    showupdatebutton=false
			response.write "<tr class='tdbg'>" & vbcrlf
			response.write " <td class='splittd' style='color:blue;font-size:14px;height:60px;line-height:32px' colspan=6>"
			response.write "</td></tr>"
			exit for
		Elseif KS.ChkClng(split(Node.SelectSingleNode("version").text,".")(0))>KS.ChkCLng(split(LocalVersion,".")(0)) then
		    set node=XMLNode.item(XMLNode.length-1)
			response.write "<tr class='tdbg'>" & vbcrlf
			response.write " <td class='splittd' style='color:blue;font-size:14px;height:60px;line-height:22px' colspan=6>"
			response.write "</td></tr>"
			exit for
		else
			response.write "<tr class='tdbg'>" & vbcrlf
			response.write " <td class='splittd' height='23' style='width:20px'><strong>" & num & "、</strong>"
			response.write "</td>"& vbcrlf
			response.write " <td class='splittd' style='text-align:left'>名称：<font color=green>" & Node.SelectSingleNode("title").text & "</font> <a href='javascript:;' onclick=""$('#v" & num & "').toggle();"">说明</a></td>"& vbcrlf
			response.write " <td class='splittd'>补丁号：<font color=red>v" & Node.SelectSingleNode("version").text &"</font> 适合版本：<font color=red>" & Node.SelectSingleNode("forversion").text & "</font></td>"& vbcrlf
			response.write " <td class='splittd'>时间：<font color=red>" & formatdatetime(Node.SelectSingleNode("adddate").text,2) &"</font></td>"& vbcrlf
			response.write " <td class='splittd'>" 
			 if num=1 then
			  if cbool(Node.SelectSingleNode("allowupdateonline").text)=true then
				response.write "<label><input type='checkbox' name='isupdate' id='isupdate' value='" &Node.SelectSingleNode("@id").text & "' checked>在线升级</label>"
			  else
			  response.write "<label style='color:#ff6600'>该补丁不支持在线升级，<br/>请到 <a style='color:#ff6600' href='#' target='_blank'>#</a> 下载补丁包手工升级。</label>"
			  end if
			 else
			  response.write "<label style='color:#999'><input type='checkbox' value=1 disabled>在线升级</label>"
			 end if
	
			response.write " </td>"
			response.write "</tr>"& vbcrlf
			response.write "<tbody style='display:none' id='v" & num & "'><tr class='tdbg'><td colspan=5 class='bborder'><iframe src='KS.Update.asp?action=showintro&id=" & Node.SelectSingleNode("@id").text & "' frameborder='0' width='100%' height='180' scrolling='auto'></iframe></td></tr></tbody>" & vbcrlf

			num=num+1
		end if
	  Next
	  response.write "</table>"& vbcrlf
	else
	 KS.Die "读取远程版本信息出错"
	end if
  
  if RemoteVersion>LocalVersion Then
    KS.echo (xmlObj.getElementsByTagName("kesioncms/message")(0).Text)
  else
    'KS.Echo "没有可升级的文件"
  end if
  if num<=2 then
  %>
  <script>$("#v1").show();</script>
  <%end if%>
  </td>
 </tr>
  <%if showupdatebutton=true then%>
 <tr><td style='height:40px;text-align:center'>
   <input type='submit' onclick="return(check())" value=" 开始下载升级文件 " class="button"/> <input type="button" value=" 暂不升级 " class="button" onclick="$(parent.frames['MainFrame'].document).find('#updateInfo').html('<font color=green>您选择了暂时不升级操作!</font>');parent.closeWindow();"/></td></tr>
 <%if num-1>1 then%>
 <tr><td style='height:40px;text-align:left;line-height:20px'><br/><strong>在级升级说明：</strong><br/>
  您当前的版本号为 <font color=red>V<%=LocalVersion%></font>,系统检测到有 <font color=red><%=num-1%></font> 个可更新的补丁，在线更新补丁必须是按顺序从早期版本的补丁一个个升级过来,并且只有同一个大版本号的补丁才可以在线升级。</td></tr>
 <%end if%>
<%else%>
 <tr><td style='height:40px;text-align:center'>
   <input type='submit' onclick="parent.closeWindow();" value=" 关闭提示窗口 " class="button"/> </td></tr>
<%end if%>
</table>
</body>
</html>
<%
end sub

sub showintro()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>开始升级操作</title>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
<link href='Include/admin_style.css' rel='stylesheet'>
<style type="text/css">
h1{margin:3px 0px 0px 10px;padding:0px;border-bottom:1px solid #cccccc;height:25px;line-height:25px;font-weight:bold;font-size:14px;}
.listinfo{margin-left:10px}
.listinfo ul{margin:0px;padding:0px}
.listinfo li{line-height:23px;background:url(images/37.gif) no-repeat 0px 7px;padding-left:10px}
</style>
<body class="tdbg">
<%
 Dim ID:id=KS.ChkClng(Request("id"))
 Call LoadRemoteXML()
 if xmlObj.readystate=4 and xmlObj.parseError.errorCode=0 Then
%>
<h1>更新列表<span style='color:#FF6600'>(更新时间:<%=formatdatetime(xmlObj.DocumentElement.selectsinglenode("item[@id='" & id & "']/adddate").text,2)%>)</span></h1>
<div class="listinfo">
<ul>
<%
  response.write xmlObj.DocumentElement.selectsinglenode("item[@id='" & id & "']/message").text
%>
</ul>
</div>
<%
 end if
%>
</body>
</html>
<%
end sub

sub begindown()
dim id:id=KS.ChkClng(Request("id"))
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>开始升级操作</title>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
<link href='Include/admin_style.css' rel='stylesheet'>
<script language='JavaScript' src='../KS_Inc/Common.js'></script>
<script language='JavaScript' src='../KS_Inc/Jquery.js'></script>
<style type="text/css">
#uptips{margin:10px}
.line{margin:0 auto;border-bottom:1px solid #efefef}
.c{background:#FBFDFF;border-top:2px solid #E1EEFF;
	height:28px;
	line-height:28px;
	letter-spacing:2px;
	font-weight:bold;
	border-bottom:1px solid #E1EEFF;
}
.line li{height:23px;line-height:23px}
.line .l{width:40%;float:left;}
.line .r{width:50%;float:left}
.clear{clear:both;height:1px;}
</style>
<body class="tdbg">
  <form name="myform" action="KS.Update.asp" method="post">
  <input type="hidden" name="action" value="update"/>
  <input type="hidden" name="id" value="<%=id%>"/>
   <div style="margin:0 auto;height:270px; overflow: auto; width:95%" id="uptips">
<%
  Dim FileList,FileLen,FileArr,I,Node,XMLNode,RemoteVersion
  
    Call LoadRemoteXML()
	if xmlObj.readystate=4 and xmlObj.parseError.errorCode=0 Then 
	 Set XMLNode=xmlObj.DocumentElement.selectsinglenode("item[@id='" & id & "']")
     RemoteVersion=xmlnode.selectsinglenode("version").text
     FileList=xmlnode.selectsinglenode("filelist").text
	else
     Call InnerHtml("<b>无法连接官方服务器，版本信息获取失败,请稍候再试...</b>")
	 ks.die ""
	end if
  FileList=Replace(FileList,vbcrlf,"")
  FileArr=Split(FileList,",")
  FileLen=Ubound(FileArr)
   response.write "<strong><font color=blue>需要升级的文件列表:</font></strong>"
   response.write "<div class='line c'><li class='l'>文件名</li><li class='r' style='padding-left:6px'>状态</li></div>"
   
  For I=0 To FileLen
    response.write "<div class='line' class='splittd'><li class='l'>" & replace(lcase(FileArr(i)),".txt",".asp") &"</li><li class='r' id='d" & i & "' style='color:#999999'>等待下载</li></div>"
  Next
  response.write "</div><div class='clear'></div>"

  SuccNum=0 : ErrNum=0
  Dim LocalFileName,RemoteFileUrl,DownSuccess,DownTimes,ErrListFile
  KS.DeleteFolder(TempDownFileDir)
  'Call InnerHtml("<b>正在下载文件在临时目录" & TempDownFileDir & "</b>")
  For I=0 To FileLen
     LocalFileName=replace(replace(trim(lcase(FileArr(i))),chr(10),""),".txt",".asp")
	 If Left(LocalFileName,1)="/" Then LocalFileName=Right(LocalFileName,Len(LocalFileName)-1)
	 LocalFileName=TempDownFileDir & replace(LocalFileName,"admin/",KS.Setting(89))
	 If IsBusiness=true then
      RemoteFileUrl=Kesion_Update_FileUrl & "vip/" & Encoding & "/" & RemoteVersion & replace(trim(FileArr(i)),chr(10),"")
	 else
      RemoteFileUrl=Kesion_Update_FileUrl & "free/" & Encoding & "/" & RemoteVersion & replace(trim(FileArr(i)),chr(10),"")
	 end if
	 DownSuccess=false
	 DownTimes=0
	  Call InnerHtml1(i,"<font color=blue>正在下载文件...</font>")
	 do while (DownSuccess=false)
	    DownSuccess=SaveRemoteFile(LocalFileName,RemoteFileUrl)
		DownTimes=DownTimes+1
		If DownTimes>10 Then Exit Do
	 loop
	 If DownSuccess Then
	    SuccNum=SuccNum+1
        Call InnerHtml1(i,"<font color=green>下载成功!</font>")
	 Else
	    ErrNum=ErrNum+1
	    Call InnerHtml1(i,"<font color=red>下载失败!</font>")
		if ErrListFile="" then
		ErrListFile=Replace(lcase(FileArr(i)),".txt",".asp")
		else
		ErrListFile=ErrListFile &","& Replace(lcase(FileArr(i)),".txt",".asp")
		end if
	 End If
  Next
  If ErrNum>0 Then
   response.write ("<div style='margin:2px'><font color=blue>发现有文件没有成功下载，下载失败的文件:" &ErrListFile &",建议直接到<a href='#' target='_blank'><font color=blue><u>官方网站</u></font></a>下载升级包覆盖!</font></div>")
  Else
    response.write ("<div style='margin:2px'>共成功下载了 <font color=red>" & SuccNum & "</font>个更新文件到临时目录，请按“立即更新”按钮覆盖。</div>")
  End If
  %> 
   
   <div style='height:40px;text-align:center'><input type='submit'  value=" 立即更新 " class="button"/> <input type="button" value=" 暂不升级 " class="button" onclick="$(parent.frames['MainFrame'].document).find('#updateInfo').html('<font color=green>您选择了暂时不升级操作!</font>');parent.closeWindow();"/></div>
</body>
</html>
 <%
end sub

sub beginupdate()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>开始升级操作</title>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
<link href='Include/admin_style.css' rel='stylesheet'>
<script language='JavaScript' src='../KS_Inc/Common.js'></script>
<script language='JavaScript' src='../KS_Inc/Jquery.js'></script>
<style type="text/css">
.listinfo li{height:23px;line-height:23px;background:url(images/37.gif) no-repeat 0px 7px;padding-left:10px}
</style>
<body class="tdbg">
  <table border="0" style="margin-top:20px" width="95%" align="center" cellspacing="0" cellpadding="0">
  <tr>
   <td class="listinfo">
   <div style="height:200px; overflow: auto; width:100%" id="uptips">
   </div>
<%
    Call InnerHtml("<b>正在开始更新升级...</b>")
    Dim FsoObj:Set FsoObj = KS.InitialObject(KS.Setting(99))
   	Dim FsoItem,FName
	Dim FolderObj:Set FolderObj = FsoObj.GetFolder(Server.MapPath(TempDownFileDir))
	Dim SubFolderObj:Set SubFolderObj = FolderObj.SubFolders
	Dim FileObj:Set FileObj = FolderObj.Files
    For Each FsoItem In SubFolderObj
	  If CopyMyFolder(TempDownFileDir & FsoItem.name,KS.Setting(3) & FsoItem.name)=true then
	  Call InnerHtml("<font color=green>成功的将目录“" & TempDownFileDir & FsoItem.name & "” 更新到 “" & KS.Setting(3) & FsoItem.name & "”</font>")
	  Else
	   Call InnerHtml("<font color=red>目录“" & TempDownFileDir & FsoItem.name & "/”更新失败，请手工将目录“" & TempDownFileDir & FsoItem.name & "/”复制到“" & KS.Setting(3) & FsoItem.name & "/”覆盖</font>")
	  End If
	Next
	For Each FsoItem In FileObj
	  FName=FsoItem.name
	  If cbool(KS.CopyFile(TempDownFileDir & FName,KS.Setting(3)& FName))=true Then
	   Call InnerHtml("<font color=green>成功的将文件“" & TempDownFileDir & FName & "” 更新到 “" & KS.Setting(3) & FName & "”</font>")
	  Else
	   Call InnerHtml("<font color=red>文件“" & TempDownFileDir & FName & "/”更新失败，请手工将文件“" & TempDownFileDir & FName & "/”复制到“" & KS.Setting(3) & FName & "/”覆盖</font>")
	  End If
	Next
    
	 '更新当前版本号
	 Call LoadRemoteXML()
	 id=ks.chkclng(request("id"))
	 if xmlObj.readystate=4 and xmlObj.parseError.errorCode=0 Then 
		 Set XMLNode=xmlObj.DocumentElement.selectsinglenode("item[@id='" & id & "']")
		 RemoteVersion=xmlnode.selectsinglenode("version").text
	   If KS.ChkClng(RemoteVersion)<>0 Then
		 Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		 Doc.async = false
		 Doc.setProperty "ServerHTTPRequest", true 
		 Doc.load(Server.MapPath("include/version.xml"))
		 doc.documentElement.selectSingleNode("//kesioncms/version").text=RemoteVersion
		 doc.save(Server.MapPath("include/version.xml"))
	   End If
	end if
	 Call InnerHtml("<b><font color=red>恭喜，您已成功的升级到版本 V" &RemoteVersion & "！</font></b>")
	
	 
  %> 
    </div>
    </td>
 </tr>
 <tr><td style='height:40px;text-align:center'><input type="button" value="关闭窗口" class="button" onclick="top.location.reload()"/></td></tr>
</table>
</body>
</html>
 <%
end sub

Function CopyMyFolder(FolderName,FolderPath) 
	on error resume next
	Dim sFolder:sFolder=server.mappath(FolderName) 
	Dim oFolder:oFolder=server.mappath(FolderPath) 
	dim fso:set fso=KS.InitialObject(KS.Setting(99))
	if fso.folderexists(sFolder) Then     '检查原文件夹是否存在 
		if fso.folderexists(oFolder) Then '检查目标文件夹是否存在 
		 fso.copyfolder sFolder,oFolder 
		Else '目标文件夹如果不存在就创建 
			KS.CreateListFolder(FolderPath)
			fso.copyfolder sFolder,oFolder 
		End if
		if err then
		 CopyMyFolder=false
		else 
		 CopyMyFolder=true
		end if
	Else 
		CopyMyFolder=false
	End If 
	set fso=nothing 
End Function


'JS InnerHtml （对象名，追加信息）
Sub InnerHtml(msg)
	Response.Write "<SCRIPT LANGUAGE=""JavaScript"">uptips.innerHTML += ""<li>"&msg&"</li>"";</SCRIPT>" &vbcrlf
	Response.Flush
End Sub
Sub InnerHtml1(i,msg)
	Response.Write "<SCRIPT LANGUAGE=""JavaScript"">d" & i &".innerHTML = """&msg&""";document.getElementById('uptips').scrollTop = "&(i+1)*23&";</SCRIPT>" &vbcrlf
	Response.Flush
End Sub


Function SaveRemoteFile(LocalFileName,RemoteFileUrl) 
    Dim DownDir:DownDir=left(LocalFileName,InStrRev(LocalFileName, "/"))
    KS.CreateListFolder(DownDir)
	SaveRemoteFile=True 
	dim Ads,Retrieval,GetRemoteData 
	Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP") 
	With Retrieval 
	.Open "Get", RemoteFileUrl, False, "", "" 
	.Send 
	If .Readystate<>4 or .status<>200 then 
		SaveRemoteFile=False 
		Exit Function 
	End If 
	GetRemoteData = .ResponseBody 
	End With 
	Set Retrieval = Nothing 
	Set Ads = Server.CreateObject("Adodb.Stream") 
	With Ads 
	.Type = 1 
	.Open 
	.Write GetRemoteData 
	.SaveToFile server.MapPath(LocalFileName),2 
	.Cancel() 
	.Close() 
	End With 
	Set Ads=nothing 
End Function 

%>