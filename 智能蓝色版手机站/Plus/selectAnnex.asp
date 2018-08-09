<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Dim KS:Set KS=New PublicCls
Dim KSUser:Set KSUser=New UserCls
if KSUser.UserLoginChecked=false then
 set ks=nothing : set ksuser=nothing
 ks.die "对不起,您没有登录!"
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>选择已上传的附件</title>
<META HTTP-EQUIV="pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache, must-revalidate">
<META HTTP-EQUIV="expires" CONTENT="Wed, 26 Feb 1997 08:21:57 GMT">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<style type="text/css">
a{text-decoration: none;} /* 链接无下划线,有为underline */ 
a:link {color: #000000;} /* 未访问的链接 */
a:visited {color: #000000;} /* 已访问的链接 */
a:hover{color: #FF0000;text-decoration: underline;} /* 鼠标在链接上 */ 
a:active {color: #FF0000;} /* 点击激活链接 */
td	{font-family:  "Verdana, Arial, Helvetica, sans-serif"; font-size: 12px;  text-decoration:none ; text-decoration:none ; }
body  {  margin:0px; font:9pt 宋体; FONT-SIZE: 9pt;text-decoration: none;}
#fenye{clear:both;}
#fenye a{text-decoration:non;}
#fenye .prev,#fenye .next{width:52px; text-align:center;}
#fenye a.curr{width:22px;background:#1f3a87; border:1px solid #dcdddd; color:#fff; font-weight:bold; text-align:center;}
#fenye a.curr:visited {color:#fff;}
#fenye a{margin:5px 4px 0 0; color:#1E50A2;background:#fff; display:inline-table; border:1px solid #dcdddd; float:left; text-align:center;height:22px;line-height:22px}
#fenye a.num{width:22px;}
#fenye a:visited{color:#1f3a87;} 
#fenye a:hover{color:#fff; background:#1E50A2; border:1px solid #1E50A2;float:left;}
#fenye span{display:block;margin:10px}
form{margin:0px;padding:0px}
.list{margin-left:5px}
.list li{float:left;width:160px;border:1px solid #cccccc;margin:2px;margin-bottom:9px;height:22px;line-height:22px;text-align:center}
</style>
</head>
<body>
<div>
 <form name="myform" action="selectAnnex.asp" method="post" >
 <strong>搜索附件=></strong>  附件名称<input type="text" name="key"> <input style="padding:2px" type="submit" value=" 快速查找 " >
 </form>
</div>
<hr size='1' color='#cccccc'/>
<div class="list">
<%
Dim TotalPut,MaxPerPage,Xml,Node,upfrom
upfrom=KS.S("upfrom")
If upfrom="Admin" then  upfrom="top.MainFrame" Else upfrom="parent"
MaxPerPage=20
Dim Param: Param=" Where IsAnnex=1 And InfoID<>0"
If Not KS.IsNul(Request("Key")) Then
Param=Param & " and title like '%" & KS.S("Key") & "%'"
End If

If KS.C("SuperTF")<>"1" Then
  If Not KS.IsNul(KS.C("AdminName")) Then
   Param=Param & " and username='" & KS.C("AdminName") & "'"
  Else
   Param=Param & " and username='" & KSUser.UserName & "'"
  End If
End If
Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
RS.Open "Select * From [KS_UploadFiles] " & Param & " Order By Id Desc",Conn,1,1
IF RS.Eof And RS.Bof Then
  totalput=0
  RS.Close:Set RS=Nothing
   If Request("key")<>"" then
  Response.write  "<div style='text-align:center'>对不起,找不到含有<font color=""red"">" & KS.CheckXSS(KS.S("Key")) & "</font>的附件!</div>"
   else
  Response.write  "<div style='text-align:center'>对不起,您没有上传过附件!</div>"
  end if
 Else
	TotalPut=Conn.Execute("Select Count(1) From KS_UploadFiles" & Param)(0)
	If CurrentPage < 1 Then CurrentPage = 1
	If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
			RS.Move (CurrentPage - 1) * MaxPerPage
	Else
			CurrentPage = 1
	End If
	on error resume next
	Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),rs,"row","")
	RS.Close:Set RS=Nothing
	If IsObject(Xml) Then
	  For Each Node In XML.DocumentElement.SelectNodes("row")
	    Dim FileName:FileName=Node.SelectSingleNode("@filename").text
		If Instr(FileName,"http")<>0 Then FileName=KS.Setting(3) & "UploadFiles/" & Split(lcase(FileName),"uploadfiles/")(1)
	    response.write "<li><a href=""javascript:;"" onclick=""" & upfrom &".InsertFileFromUp('" & FileName & "'," & KS.GetFieSize(Server.MapPath(FileName)) &"," & Node.SelectSingleNode("@id").text & ",'" & Node.SelectSingleNode("@title").text & "');" & upfrom &".box.close();"" title='附件id:" & Node.SelectSingleNode("@id").text & " &#13;附件名称:" & Node.SelectSingleNode("@title").text & "&#13;下载次数:" & Node.SelectSingleNode("@hits").text & "次&#13;文件名:" & Node.selectsinglenode("@filename").text & "'>" & KS.Gottopic(Node.SelectSingleNode("@title").text,30) &"</a></li>"
	  Next
	End If
End IF
%>
</div>
<%=KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,false)%>
</body>
</html>
<%
Set KS=Nothing
CloseConn
%>