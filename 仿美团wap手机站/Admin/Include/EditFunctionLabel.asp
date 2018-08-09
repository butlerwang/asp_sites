<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
<link href='admin_style.css' rel='stylesheet'>
<script language='JavaScript' src='../../KS_Inc/Common.js'></script>
<script language='JavaScript' src='../../KS_Inc/Jquery.js'></script>
<script language="javascript">
function CheckForm(){
 frames["LabelShow"].CheckForm();
}
</script>
<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0' scroll="no">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="25">
<div class="topdashed sort">编 辑 系 统 函 数 标 签 属 性</div>
<%
Dim LabelID,LabelRS,Str,LabelContent,FileName
LabelID=Trim(Request.QueryString("LabelID"))
Set LabelRS=Server.CreateObject("Adodb.Recordset")
 Str="SELECT LabelContent FROM KS_Label Where ID='" & LabelID &"'"
 LabelRS.Open Str,Conn,1,1
IF LabelRS.Eof and LabelRS.Bof THEN
 LabelRS.Close
 Set LabelRS=Nothing
 Response.Write("<Script>alert('参数传递出错!');window.close();</Script>")
 Response.End
End if 
 LabelContent=LabelRS(0)
 LabelRS.Close
 Set LabelRS=Nothing
 'on error resume next
 'Str=mid(LabelContent, InStrrev(LabelContent, "("))
 'FileName=Replace(Replace(LabelContent,Str,""),"{$","")
 'If Err Then
	 Str=Split(LabelContent," ")(0)
	 FileName=Replace(Str,"{Tag:","")
 'End If
 FileName=FileName & ".asp?Action=Edit&LabelID=" & LabelID
 Response.WRITE "<script>"
 'Response.Write "location.href='Label/" & FILENAME & "';"
 Response.Write "</script>"
%>
</td>
</tr>
<tr>
 <td>
<iframe name="LabelShow" id="LabelShow" src="Label/<%=FileName%>" style="width:100%;height:100%" frameborder="0"  scrolling="auto"></iframe>
 </td>
</tr>
</body>
</html>
 
