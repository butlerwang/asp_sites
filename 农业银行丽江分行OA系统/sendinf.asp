<%if Session("Urule")<>"a" then
	Response.write "��û���㹻Ȩ��"
	response.end
end if%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<LINK href="oa.css" rel=stylesheet>
<title>����֪ͨ</title>
<script language="JavaScript">
<!--
function  validate()
    {
       
        if  (document.myform.biaoti.value=="")
        {
            alert("���ⲻ��Ϊ��");
            document.myform.biaoti.focus();
            return false ;
        }
		if  (document.myform.neirong.value=="")
        {
            alert("���ݲ���Ϊ��");
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
        �ꡡ����</td> 
      <td width="79%" height="21"><input type="text" name="biaoti" size="21" style="width: 288; height: 18;background-color: #ffffff;filter:chroma(color=#ffffff);"></td>
    </tr>
    <tr>
      <td width="21%" align="center" height="122">
        �ڡ�����</td>
      <td width="79%"  valign="middle"><textarea rows="20" name="neirong" cols="38" style="font-size: 10pt;background-color: #ffffff;filter:chroma(color=#ffffff);overflow: auto"></textarea></td>
    </tr>
  </table>
  <p align="center">
    <input type="submit" value="����֪ͨ" name="B1" style="border:1pt solid #636563;height:18px;background-color: #ffffff;filter:chroma(color=#ffffff);">&nbsp;&nbsp;<INPUT TYPE="reset" value="������д" style="border:1pt solid #636563;height:18px;background-color: #ffffff;filter:chroma(color=#ffffff);">
    </p>
</form></TD>
            
      </TR>
	  
      </table>



</body>

</html>

