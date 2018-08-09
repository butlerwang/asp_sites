<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<%Dim Action,JSID,EditUrl,FolderID
Action=Request.QueryString("Action")
JSID=Request.QueryString("JSID")
FolderID=Request.QueryString("FolderID")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="JavaScript" src="../../KS_Inc/jQuery.js"></script>
<script language="javascript">
function CheckForm(){
 frames["JSFrame"].CheckForm();
}
</script>

<link href="../Admin_Style.CSS" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" scroll=no>
<div class="topdashed sort">
			 <%IF Action="Edit" Then
			 Response.Write("<Strong>编辑 JS</Strong>")
			 Else
			 %>
               新建系统JS
			<%end if%>
</div>
     <table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top">
	<%IF Action="Edit" Then
	 EditUrl=Request.Querystring("EditUrl")
	 Response.Write("<iframe src=""" & EditUrl &""" name=""JSFrame"" id=""JSFrame"" width=""100%"" height=""93%"" frameborder=""0"" scrolling=""auto""></iframe>")
	else
	 Response.Write("<iframe src=""getgenericlist.asp?Channelid=1&JSID=" & JSID & "&FolderID=" & FolderID &"&Action=" & Action &""" name=""JSFrame"" id=""JSFrame"" width=""100%"" height=""100%"" frameborder=""0"" scrolling=""auto""></iframe>")
	End IF%>
</td>
  </tr>
</table>
</body>
</html>
<script>
 function SelectJSType(ObjValue)
  {
   frames['JSFrame'].location.href=ObjValue;
  }
</script> 
