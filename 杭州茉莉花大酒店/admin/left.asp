<!--#include file="check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨����ϵͳ</title>
<link href="include/style.css" rel="stylesheet" type="text/css">
<style  type="text/css">  
BODY {
	SCROLLBAR-FACE-COLOR: #EFF3F7;  FONT-SIZE: 9pt; BACKGROUND: #ffffff; SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; SCROLLBAR-SHADOW-COLOR: #C6D3DE; SCROLLBAR-3DLIGHT-COLOR: #C6D3DE; SCROLLBAR-ARROW-COLOR: #C6D3DE; SCROLLBAR-TRACK-COLOR: #ffffff; SCROLLBAR-DARKSHADOW-COLOR: #ffffff; TEXT-DECORATION: none
}
</style>  

</head>
<base target="right">
<body>
<br>
<%IF session("name")="" Then%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="6%" align="center">&nbsp;</td>
    <td width="94%" height="20"><img src="images/ducument.gif" width="25" height="13"> <strong>��վ��̨����ϵͳ</strong></td>
  </tr>
  <tr>
    <td height="23">&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="images/next(1).gif" width="9" height="10"> 
            &nbsp;<a href="login.asp" target="_top">��̨����ϵͳ</a></td>
        </tr>
      </table>      
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="6%" align="center">&nbsp;</td>
    <td width="94%" height="20"><img src="images/ducument.gif" width="25" height="13"> <span class="font_bold">��վ��ҳ</span></td>
  </tr>
<tr>
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="images/next(1).gif" width="9" height="10"> &nbsp;<a href="../index.asp" target="_blank">��վ��ҳ</a></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<!--<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="6%" align="center">&nbsp;</td>
    <td width="94%" height="20"><img src="images/ducument.gif" width="25" height="13"> <span class="font_bold">����ͳ��ϵͳ</span></td>
  </tr>
<tr>
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="images/next(1).gif" width="9" height="10"> &nbsp;<a href="../count/index.asp" target="_blank">����ͳ��ϵͳ</a></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>-->
<!--<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="6%" align="center">&nbsp;</td>
    <td width="94%" height="20"><img src="images/ducument.gif" width="25" height="13"> <span class="font_bold">���ݿ����</span></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="images/next(1).gif" width="9" height="10"> &nbsp;<a href="weblink/">���ݿⱸ��</a></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td height="20"><img src="images/next(1).gif" width="9" height="10"> ���ݿ⵼��EXCEL</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
-->
<%
else
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="6%" align="center">&nbsp;</td>
    <td width="94%" height="20"><img src="images/ducument.gif" width="25" height="13"> <strong>��վ����ϵͳ</strong></td>
  </tr>
  <tr>
    <td height="23">&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="8%">&nbsp;</td>
        <td width="92%" height="20"><img src="images/next(1).gif" width="9" height="10"> &nbsp;��ӭ��/Welcome��</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="images/next(1).gif" width="9" height="10"> &nbsp;<a href="member/edit_password.asp">�ʺŹ���</a></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="images/next(1).gif" width="9" height="10"> &nbsp;<a href="../" target="_blank">��վ��ҳ</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<!--<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="6%" align="center">&nbsp;</td>
    <td width="94%" height="20"><img src="images/ducument.gif" width="25" height="13"> <span class="font_bold">����ͳ��ϵͳ</span></td>
  </tr>
<tr>
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="8%">&nbsp;</td>
          <td width="92%" height="20"><img src="images/next(1).gif" width="9" height="10"> &nbsp;<a href="../count/index.asp" target="_blank">����ͳ��ϵͳ</a></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>-->
<%
end if%>
</body>
</html>