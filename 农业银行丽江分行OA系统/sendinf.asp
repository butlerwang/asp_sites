<%if Session("Urule")<>"a" then
	Response.write "你没有足够权限"
	response.end
end if%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<LINK href="oa.css" rel=stylesheet>
<title>发布通知</title>
<script language="JavaScript">
<!--
function  validate()
    {
       
        if  (document.myform.biaoti.value=="")
        {
            alert("标题不能为空");
            document.myform.biaoti.focus();
            return false ;
        }
		if  (document.myform.neirong.value=="")
        {
            alert("内容不能为空");
            document.myform.neirong.focus();
            return false ;
        }
		}
//-->
</script>
</head>

<BODY
style="BACKGROUND-ATTACHMENT: scroll; BACKGROUND-IMAGE: url(images/main_bg.gif); BACKGROUND-POSITION: left bottom; BACKGROUND-REPEAT: no-repeat">
<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
  <TBODY>
	  <TR> 
            <TD>
			<form method="POST" name="myform" action="inf_to_server.asp"  onsubmit="return  validate();">
 
  <table border="1" width="100%" cellspacing="1" height="121" bordercolorlight=#000000 bordercolordark=#ffffff>
    <tr>
      <td width="21%" align="center" height="21">
        标　　题</td> 
      <td width="79%" height="21"><input type="text" name="biaoti" size="21" style="width: 288; height: 18;background-color: #ffffff;filter:chroma(color=#ffffff);"></td>
    </tr>
    <tr>
      <td width="21%" align="center" height="122">
        内　　容</td>
      <td width="79%"  valign="middle"><textarea rows="20" name="neirong" cols="38" style="font-size: 10pt;background-color: #ffffff;filter:chroma(color=#ffffff);overflow: auto"></textarea></td>
    </tr>
  </table>
  <p align="center">
    <input type="submit" value="发布通知" name="B1" style="border:1pt solid #636563;height:18px;background-color: #ffffff;filter:chroma(color=#ffffff);">&nbsp;&nbsp;<INPUT TYPE="reset" value="重新填写" style="border:1pt solid #636563;height:18px;background-color: #ffffff;filter:chroma(color=#ffffff);">
    </p>
</form></TD>
            
      </TR>
	  
      </table>



</body>

</html>

