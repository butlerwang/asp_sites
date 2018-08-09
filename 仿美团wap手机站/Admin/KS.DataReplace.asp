<%Option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.Asp"-->
<%
Response.Buffer=true
Response.CharSet="utf-8"
Server.ScriptTimeout=9999999
Dim KSCls
Set KSCls = New Admin_Replace
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Replace
        Private KS,Action,I,BeginTime,EndTime
		Private Sub Class_Initialize()
		   Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		Sub Kesion()
		 If Not KS.ReturnPowerResult(0, "KMST10008") Then                '检查在线执行SQL语句
				  Call KS.ReturnErr(1, "")
			  Response.End
         End If
%>
<html>
<head>
<title>数据库内容替换程序</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="include/admin_style.CSS" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
if trim(request("type1"))="GetChird" then
	Call ShowChird()
else
		Response.Write "<ul id='menu_top'>"
		Response.Write "<li class='parent' onclick=""location.href='KS.DataReplace.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>普通模式</span></li>"
		Response.Write "<li class='parent' onclick=""location.href='?Action=Main2';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/s.gif' border='0' align='absmiddle'>高级模式</span></li>"
		Response.Write "<li class='parent' onclick=""location.href='?Action=Main3';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>赋值模式</span></li>"
		Response.Write "<li class='parent' onclick=""location.href='?Action=Main4';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>所有图片地址替换</span></li>"
		Response.Write "</ul>"

		i=0
		Action=trim(request("action"))
		Select Case Action
		Case "Replace1","Replace2","Replace3" call step1()
		Case "Main2" call Main2()
		Case "Main3" call Main3()
		Case "Main4" call Main4()
		Case "Replace4"  call Replace4()
		Case Else
			call Main1()
		End Select
end if
%>
</body>
</html>
<%End Sub

Sub Main1()
%>
<script language = "JavaScript">
function changedb(){
    tablechird.location.href="KS.DataReplace.asp?type1=GetChird&type="+document.myform.TableName.value;
}
function CheckForm(){
  if (document.myform.TableName.value==''){
    alert('数据表名不能为空！');
    document.myform.TableName.focus();
    return false;
  }
  if (document.myform.ColumnName.value==''){
    alert('字段名不能为空！');
    //document.myform.ColumnName.focus();
    return false;
  }
  if (document.myform.strOld.value==''){
    alert('替换字符不能为空！');
    document.myform.strOld.focus();
    return false;
  }
  return true;  
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="ctable">
<form method="post" name="myform" onSubmit="return CheckForm();" action="KS.DataReplace.asp" target="_self">
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" width=80 class="clefttitle"><strong>数据表名：</strong></td>
	<td height="30"><%Call ShowMain()%><font color=red> *</font></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="clefttitle"><strong>字 段 名：</strong></td>
	<td height="30"><input name="ColumnName" type="hidden" value="" size="40"><iframe style="top:2px;" ID="tablechird" src="KS.DataReplace.asp?type1=GetChird" frameborder=0 scrolling=no width="100%" height="22"></iframe></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="clefttitle"><strong>将 字 符：</strong></td>
	<td height="30"><textarea name="strOld" cols="60" rows="4"></textarea><font color=red> *</font></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="clefttitle"><strong>替 换 成：</strong></td>
	<td height="30"><textarea name="strNew" cols="60" rows="4"></textarea></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="clefttitle"><strong>注意事项：</strong></td><td height="30">	1、执行操作前，请备份数据库文件。<br>2、本操作的更新时间视您数据的多少以及服务器（或本地机器）的配置决定，如果数据很多，更新可能很慢，在这过程千万不能刷新页面或关闭浏览器，如果出现超时或者错误提示，请使用备份数据重新进行操作。</td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'"><td height="30" colspan="2" style="text-align:center"><input type="hidden" name="Action" value="Replace1"><input class="button" type="submit" name="Submit" value="开始替换"></td></tr>
</table>
</form>
<%
End Sub

Sub Main2()
%>
<script language = "JavaScript">
function changedb(){
    tablechird.location.href="KS.DataReplace.asp?type1=GetChird&type="+document.myform.TableName.value;
}
function CheckForm1(){
  if (document.myform.TableName.value==''){
    alert('数据表名不能为空！');
    document.myform.TableName.focus();
    return false;
  }
  if (document.myform.ColumnName.value==''){
    alert('字段名不能为空！');
    //document.myform.ColumnName.focus();
    return false;
  }
  if (document.myform.strOld.value==''){
    alert('替换开始代码不能为空！');
    document.myform.strOld.focus();
    return false;
  }
   if (document.myform.strOld1.value==''){
    alert('替换结束代码不能为空！');
    document.myform.strOld1.focus();
    return false;
  }
  return true;  
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="CTable">
<form method="post" name="myform" onSubmit="return CheckForm1();" action="KS.DataReplace.asp" target="_self">
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" width=80 class="clefttitle"><strong>数据表名：</strong></td>
	<td height="30"><%Call ShowMain()%><font color=red> *</font></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="clefttitle"><strong>字 段 名：</strong></td>
	<td height="30"><input name="ColumnName" type="hidden" value="" size="40"><iframe style="top:2px;" ID="tablechird" src="KS.DataReplace.asp?type1=GetChird" frameborder=0 scrolling=no width="100%" height="22"></iframe></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="clefttitle"><strong>字符开始：</strong></td>
	<td height="30"><textarea name="strOld" cols="60" rows="3"></textarea><font color=red> *</font></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="clefttitle"><strong>字符结束：</strong></td>
	<td height="30"><textarea name="strOld1" cols="60" rows="3"></textarea><font color=red> *</font></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="clefttitle"><strong>替 换 成：</strong></td>
	<td height="30"><textarea name="strNew" cols="60" rows="3"></textarea></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="clefttitle"><b>特别事项</b>：</td><td height="30">1、<font color=red>高级模式较易出错，且替换速度较慢、较占服务器CPU和内存资源,请谨慎使用!如无特殊需要,建议使用普通模式!</font><br>	2、执行操作前，请备份数据库文件。<br>3、本操作的更新时间视您数据的多少以及服务器（或本地机器）的配置决定，如果数据很多，更新可能很慢，在这过程千万不能刷新页面或关闭浏览器，如果出现超时或者错误提示，请使用备份数据重新进行操作。</td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'"><td height="30" colspan="2" style="text-align:center"><input type="hidden" name="Action" value="Replace2"><input type="submit" name="Submit" value="开始替换" class="button"></td></tr>
</form>
</table>
<%
End Sub

Sub Main3()
%>
<script language = "JavaScript">
function changedb(){
    tablechird.location.href="KS.DataReplace.asp?type1=GetChird&type="+document.myform.TableName.value;
}
function CheckForm1(){
  if (document.myform.TableName.value==''){
    alert('数据表名不能为空！');
    document.myform.TableName.focus();
    return false;
  }
  if (document.myform.ColumnName.value==''){
    alert('字段名不能为空！');
    return false;
  }
  return true;  
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="CTable">
<form method="post" name="myform" onSubmit="return CheckForm1();" action="KS.DataReplace.asp" target="_self">
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" width=80 class="clefttitle"><strong>数据表名：</strong></td><td height="30"><%Call ShowMain()%><font color=red> *</font></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="clefttitle"><strong>字 段 名：</strong></td>
	<td height="30"><input name="ColumnName" type="hidden" value="" size="40"><iframe style="top:2px;" ID="tablechird" src="KS.DataReplace.asp?type1=GetChird" frameborder=0 scrolling=no width="100%" height="22"></iframe></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="clefttitle"><strong>修 改 成：</strong></td>
	<td height="30"><textarea name="strNew" cols="60" rows="3"></textarea></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="clefttitle"><b>特别事项</b>：</td><td height="30">1、<font color=red>赋值模式不理会原有数据,而是把选定字段的值直接修改为新的值,这会导致原数据丢失并变为新数据!请谨慎使用!</font>	<br>2、执行操作前，请备份数据库文件。<br>3、本操作的更新时间视您数据的多少以及服务器（或本地机器）的配置决定，如果数据很多，更新可能很慢，在这过程千万不能刷新页面或关闭浏览器，如果出现超时或者错误提示，请使用备份数据重新进行操作。</td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'"><td height="30" colspan="2" style="text-align:center"><input type="hidden" name="Action" value="Replace3"><input class="button" type="submit" name="Submit" value="开始替换"></td></tr>
</form>
</table>
<%
End Sub

'图片字段内容替换
Sub Main4()
%>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="ctable">
<form method="post" name="myform" action="?" target="_self">
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td align="right" class="clefttitle"><strong>注意事项：</strong></td>
	<td ><br/>1、此功能一般用于网站域名更换时才使用，将所有涉及图片的字段的原域名替换为新的域名，使用前请先备份好您的数据库，如果您对此不太了解请慎用！！！<br/>
	2、本操作的更新时间视您数据的多少以及服务器（或本地机器）的配置决定，如果数据很多，更新可能很慢，在这过程千万不能刷新页面或关闭浏览器，如果出现超时或者错误提示，请使用备份数据重新进行操作。<br/>3、要替换的图片字段可以打开config/photofield.txt文件增加或删除。<br/><br/></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="clefttitle" nowrap="nowrap"><strong>将 字 符：</strong></td>
	<td height="30"><textarea name="strOld" cols="60" rows="4"></textarea><font color=red> *</font>如 http://www.kesion.cn</td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="clefttitle"><strong>替 换 成：</strong></td>
	<td height="30"><textarea name="strNew" cols="60" rows="4"></textarea>&nbsp;</td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'"><td height="30" colspan="2" style="text-align:center"><input type="hidden" name="Action" value="Replace4"><input class="button" type="submit" name="Submit" value="开始替换"></td></tr>
</table>
</form>
<%
End Sub

'**************************************************
'过程名：Step1
'作  用：数据替换主调用程序
'参  数：无
'**************************************************
Sub Step1()
	dim rs,sql
	BeginTime=timer
	dim TableName,ColumnName,strOld,strOld1,strNew
	TableName	= KS.R(trim(request("TableName")))
	ColumnName	= KS.R(trim(request("ColumnName")))
	strOld		= trim(request("strOld"))
	strOld1		= trim(request("strOld1"))
	strNew		= trim(request("strNew"))
	if TableName="" then
		response.write "请输入要替换的数据表名！"
		exit sub
	end if
	if ColumnName="" then
		response.write "请输入要替换的字段名！"
		exit sub
	end if
	if action="Replace1" then
		if strOld="" then
			response.write "请输入要替换的字符！"
			exit sub
		end if
	else 
		if action="Replace2" then
			if strOld="" then
				response.write "请输入要替换的字符开始代码！"
				exit sub
			end if
			if strOld1="" then
				response.write "请输入要替换的字符结束代码！"
				exit sub
			end if
		End if		
	End if
	on error resume next
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "select " & ColumnName & "  from " & TableName
	OpenConn : rs.open sql,conn,1,1
	if err.number<>0 then
		response.write "<font color=red>数据库操作失败，请检查数据表名和字段名添写是否正确</font>"
		exit sub
	end if
	set rs=nothing
	on error GoTo 0
	response.write "<br>正在替换有关数据……<font color=red>在此过程中请勿刷新页面或关闭窗口！</font><br>"
	call ReplaceData(TableName,ColumnName,strOld,strOld1,strNew)	
	EndTime=timer	

    Response.Write "<table border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""  class=""ctable"" width=""90%""><tr align='center' class='tdbg'><td height='22'><strong>恭喜你！</strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='100' valign='top'><br>" & "共替换：<font color='#0000ff'>"&i&"</font> 项数据。<br>共耗时：<font color='#0000ff'>"&FormatNumber((EndTime-BeginTime)*1000,2)&"</font> 毫秒。" & "</td></tr>" & vbCrLf
    Response.Write "  <tr align='center' class='tdbg' height='30'><td><input type='button' class='button' onclick=history.go(-1)  value='返回上一页'></td></tr></table>"& vbCrLf

End Sub

'**************************************************
'过程名：ReplaceData
'作  用：显示数据表列表
'参  数：TableName	----数据表名
'        ColumnName	----字段名
'        strOld		----查找字符(或查找开始代码)
'        strOld1	----查找结束代码
'        strNew		----替换字符
'**************************************************
Sub ReplaceData(TableName,ColumnName,strOld,strOld1,strNew)
	dim rs,sql,tt
	Response.Write "<li>正在替换数据</li>&nbsp;&nbsp;"
	Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "select " & ColumnName & "  from " & TableName 
	rs.open sql,conn,1,3
	Select Case Action
	Case "Replace1"
		do while not rs.eof
			if instr(rs(ColumnName),strOld)<>0 then
				rs(ColumnName)=replace(rs(ColumnName),strOld,strNew)
				i=i+1
				Response.Write "."
			end if
			rs.update
			rs.movenext
		loop
	Case "Replace2"
		Set tt=new RegExp
  		tt.IgnoreCase =true
  		tt.Global=True
  		tt.Pattern=strOld&"[^"&strOld1&"]*"&strOld1
		response.write strOld&"[^"&strOld1&"]*"&strOld1
		do while not rs.eof
 			rs(ColumnName) = tt.Replace(rs(ColumnName),strNew) 
  			i=i+1
			Response.Write "."
			rs.update
			rs.movenext
		loop
		Set tt=Nothing
	Case "Replace3"
		do while not rs.eof
			Response.Write strNew
			rs(ColumnName)=strNew
			i=i+1
			Response.Write "."
			rs.update
			rs.movenext
		loop
	End Select
	rs.close:set rs=nothing
	response.write "&nbsp;&nbsp;<font color='#009900'>替换数据成功！</font>"
End Sub


Sub Replace4()
	dim rs,sql,tt,i,strOld,strNew,k,table,farr,ii
	strOld=request("strOld")
	strNew=request("strNew")
	BeginTime=timer
	i=0
	dim rfield:rfield=KS.ReadFromFile("../config/photofield.txt")&""
	rfield=split(rfield,vbcrlf)
	Response.Write "<li>正在替换数据</li>&nbsp;&nbsp;"
	for k=0 to ubound(rfield)
	 if not ks.isnul(rfield(k)) then
	   table=split(rfield(k),"|")(0)
	   farr=split(split(rfield(k),"|")(1),",")
	   for ii=0 to ubound(farr)
	     if not ks.isnul(farr(ii)) then
		  if lcase(table)="ks_article" then
		    call replacemodelphoto(1,farr(ii),strOld,strNew)
		  elseif lcase(table)="ks_photo" then
		    call replacemodelphoto(2,farr(ii),strOld,strNew)
		  elseif lcase(table)="ks_download" then
		    call replacemodelphoto(3,farr(ii),strOld,strNew)
		  elseif lcase(table)="ks_flash" then
		    call replacemodelphoto(4,farr(ii),strOld,strNew)
		  else
			 Response.Write "<li>正在替换数据表" & table & "的字段" & farr(ii) & "</li>&nbsp;&nbsp;"
			 response.Flush()
			 Call DoReplaceData(table,farr(ii),strOld,strNew)	
		  end if
		 end if
	   next
	 end if
	next
	EndTime=timer	
    Response.Write "<table border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""  class=""ctable"" width=""98%""><tr align='center' class='tdbg'><td height='22'><br/><strong><font color='#009900'>所有数据替换成功！</font></strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='30' valign='top'>共耗时：<font color='#0000ff'>"&FormatNumber((EndTime-BeginTime)*1000,2)&"</font> 毫秒。" & "</td></tr>" & vbCrLf
    Response.Write "  <tr align='center' class='tdbg' height='30'><td><input type='button' class='button' onclick=history.go(-1)  value='返回上一页'></td></tr></table>"& vbCrLf
	
end sub	

sub replacemodelphoto(basictype,fieldname,strOld,strNew)
   dim rs:set rs=conn.execute("select channeltable from ks_channel where basictype=" & basictype)
			do while not rs.eof
			 Response.Write "<li>正在替换数据表" & rs(0) & "的字段" & fieldname & "</li>&nbsp;&nbsp;"
			 response.Flush()
			 Call DoReplaceData(rs(0),fieldname,strOld,strNew)	
			rs.movenext
			loop
			rs.close
	set rs=nothing
end sub
	
'**************************************************
'过程名：ReplaceData
'作  用：显示数据表列表
'参  数：TableName	----数据表名
'        ColumnName	----字段名
'        strOld		----查找字符
'        strNew		----替换字符
'**************************************************
Sub DoReplaceData(TableName,ColumnName,strOld,strNew)	
	    dim rs:Set rs = Server.CreateObject("ADODB.Recordset")
		dim sql:sql = "select " & ColumnName & "  from " & TableName 
	    rs.open sql,conn,1,3
		i=0
		do while not rs.eof
			if instr(rs(ColumnName),strOld)<>0 and instr(lcase(rs(ColumnName)),"#")=0 then
				rs(ColumnName)=replace(rs(ColumnName),strOld,strNew,1,-1,1)
				i=i+1
				'Response.Write "."
			end if
			rs.update
			rs.movenext
		loop
		response.write "&nbsp;<font color=#999999>共替换了<font color=red>" & i & "</font>项</font>"
	rs.close:set rs=nothing
End Sub

'**************************************************
'过程名：ShowMain
'作  用：显示数据表列表
'参  数：无
'**************************************************
Sub ShowMain()
	dim rs,tablename,temptable
	OpenConn : Set rs = Conn.OpenSchema(4)
	tablename=""
	response.write "<select name='TableName' onChange='changedb()'><option value=''>请选择一个数据表</option>"
	Do Until rs.EOF
		temptable=rs("Table_name")
		if temptable <> tablename and temptable <> "KS_Admin" and temptable <> "KS_NotDown" and temptable <> "MSysAccessXML" and temptable <> "MSysAccessObjects" then
			Response.write "<option value='" & temptable & "'>" & temptable & "</option>"
			Tablename = temptable
		end if
	rs.MoveNext
	Loop
	response.write "</select>"
	rs.close:set rs=nothing
End Sub

'**************************************************
'过程名：ShowChird
'作  用：显示指定数据表的字段列表
'参  数：无
'**************************************************
Sub ShowChird()
	dim rs
	response.write "<body class='tdbg'><form method='post' name='myform11' action='KS.DataReplace.asp'><select name='dbname2' onChange=parent.document.myform.ColumnName.value=document.myform11.dbname2.value><option value=''>请选择一个字段　</option>"
	if trim(request("type"))<>"" then	
		OpenConn : Set rs = Conn.OpenSchema(4)	
		Do Until rs.EOF or rs("Table_name") = trim(request("type"))
			rs.MoveNext
		Loop
		Do Until rs.EOF or rs("Table_name") <> trim(request("type"))
			response.write "<option value='"&rs("column_Name")&"'>"&rs("column_Name")&"</option>"
			rs.MoveNext
		loop
		rs.close:set rs=nothing
	End if
	response.write "</select><font color=red> *</font></form><script language = 'JavaScript'>parent.document.myform.ColumnName.value=document.myform11.dbname2.value;</script></body>"
End Sub
End Class
%> 
