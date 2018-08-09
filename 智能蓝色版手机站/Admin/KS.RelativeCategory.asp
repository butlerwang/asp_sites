<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim ChannelID,RS
Dim KS:Set KS= New PublicCls
ChannelID = KS.ChkClng(KS.G("ChannelID"))
If ChannelID = 0 Then ChannelID = 1
Dim Doc,Node,I,OpStr
set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
Doc.async = false
Doc.setProperty "ServerHTTPRequest", true 
Doc.load(Server.MapPath(KS.Setting(3)&"Config/relativeType.xml"))

If KS.S("Action")="dosave" Then
 Dim CategoryItemArr,CategoryItem:CategoryItem=Replace(KS.S("Item")," ","")
 If KS.IsNul(CategoryItem) Then KS.AlertHintScript "请输入分类名称!"
 CategoryItemArr=Split(CategoryItem,",")
 Set Node=Doc.DocumentElement.SelectSingleNode("model[@channelid=" &channelid&"]")
 if not node is nothing then  Doc.DocumentElement.RemoveChild(Node)
 
 Set Node=Doc.documentElement.appendChild(Doc.createNode(1,"model",""))
 Node.attributes.setNamedItem(Doc.createNode(2,"channelid","")).text=channelid
 For I=0 To Ubound(CategoryItemArr)
	 Dim Nn:Set NN=Node.appendChild(Doc.createNode(1,"item",""))
	 NN.text=CategoryItemArr(i)
	 OpStr =OpStr &"<option>" & CategoryItemArr(i)&"</option>"
 Next
Doc.save(Server.MapPath(KS.Setting(3)&"Config/relativeType.xml"))
Application(KS.SiteSN&"_Configrelativetype")=empty
response.write "<script>parent.setrelcategory('" & opstr & "');</script>"
set ks=nothing
closeconn
response.end
End If
%> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
<script type="text/JavaScript">
	var rowtypedata = [
		[
			[1,'<input type="text" name="item"  size="20" class="textbox"/>', 'tdbg']
		],
	];

var addrowdirect = 0;
function addrow(obj, type) {
	var table = obj.parentNode.parentNode.parentNode.parentNode;
	if(!addrowdirect) {
		var row = table.insertRow(obj.parentNode.parentNode.parentNode.rowIndex);
	} else {
		var row = table.insertRow(obj.parentNode.parentNode.parentNode.rowIndex + 1);
	}
	var typedata = rowtypedata[type];
	for(var i = 0; i <= typedata.length - 1; i++) {
		var cell = row.insertCell(i);
		cell.colSpan = typedata[i][0];
		var tmp = typedata[i][1];
		if(typedata[i][2]) {
			cell.className = typedata[i][2];
		}
		tmp = tmp.replace(/\{(\d+)\}/g, function($1, $2) {return addrow.arguments[parseInt($2) + 1];});
		cell.innerHTML = tmp;
	}
	addrowdirect = 0;
}
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="myform" action="KS.RelativeCategory.asp" method="post">
 <input type="hidden" name="channelid" value="<%=channelid%>">
 <input type="hidden" name="action" value="dosave">
 <table cellspacing="1" width="96%" align="center" cellpadding="1" class="ctable" border="0">
    <tr><td class='clefttitle' style="text-align:center"><strong>分类名称</strong></td></tr>
<%
  If IsObject(Doc) Then
	For Each Node In Doc.DocumentElement.SelectNodes("model[@channelid=" & channelid &"]/item")
	%>
	<tr><td class='tdbg'><input type="text" name="item" size="20" value="<%=Node.text%>" class="textbox" /></td></tr>
	<%
	Next
 Else%>
  	<tr><td><input type="text" name="item" size="20" class="textbox" /></td></tr>
 <%
  End If

%>
	<tr><td class='tdbg'><div><img src="images/accept.gif" align="absmiddle"/> <a href="#" onClick="addrow(this, 0)" class="addtr">增加一项</a></div></td>
	</tr>
	</table>
	<div style='text-align:center'>
   <input type='submit' value='确定保存' class='button'/>
    </div>
 </form>
</body>
</html>
<%
set ks=nothing
closeconn
%>