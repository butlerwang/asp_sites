<%
if session("Urule")<>"a" then
response.redirect "error.asp?id=admin"
end if
%>
<!--#include file="data.asp"-->
<script language="javascript" src="ShowProcessBar.js"></script>

<html><head><title>����ļ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="oa.css" rel=stylesheet>
</head>
<BODY>
<Script Language="javaScript">
    function  validate()
    {
       
        if  (document.myform.src.value=="")
        {
            alert("������Ʋ���Ϊ��");
            document.myform.src.focus();
            return false ;
        }
        if  (document.myform.content.value=="")
        {
            alert("˵������Ϊ��");
            document.myform.content.focus();
            return false ;
        }
        if  (document.myform.height.value=="")
        {
            alert("�߶Ȳ���Ϊ��");
            document.myform.height.focus();
            return false ;
        }
        if  (document.myform.width.value=="")
        {
            alert("��Ȳ���Ϊ��");
            document.myform.width.focus();
            return false ;
        }
     
}
</Script>

<form method="post" action="adrot_save.asp" enctype="multipart/form-data" name=myform  onSubmit='return validate()'>
  <div align="center">
        <TABLE border=1 bordercolorlight='000000' bordercolordark=#ffffff cellspacing=0 cellpadding=0 align=center>
      <TR> 
            <TD height=20 width=60>&nbsp;���ӵ�ַ</TD>
            <TD height=20> 
              <INPUT name="url" 
            size=30>
            </TD>
          </TR>
      <TR> 
            <TD height=20 width=60> 
              <div align="center">��  ��<BR><BR>˵  ��</div>
            </TD>
            <TD height=20> 
              <TEXTAREA cols=40 name="content" rows=6 class="txt" style="overflow:auto"></TEXTAREA>
            </TD>
          </TR>
      <TR> 
            <TD height=20 colspan=2> 
              &nbsp;�������:&nbsp; <SELECT NAME="type" style="height:18px;font-size:9pt"><option value="GIF" selected>GIF</option><option value="SWF">SWF</option></SELECT> ��泤��<INPUT TYPE="text" NAME="width" size=2 value="485">��<INPUT TYPE="text" NAME="height" size=1 value="75">
            </TD>
          </TR>
      <tr>
        <TD height=20 width=60>&nbsp;���</TD>
        <td height="16"><input type="file" name="src" size="20" value="���">&nbsp;&nbsp;<input type="submit" value="�ϴ�" name="B1" class="txt " IsShowProcessBar="True"></td>
      </tr></table>
    
  </div></form></body></html>
