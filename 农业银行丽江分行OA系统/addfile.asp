<%
if Session("Urule")="c" then
response.redirect "error.asp?id=admin"
end if
%>
<!--#INCLUDE FILE="check.asp" -->

<script language="javascript" src="ShowProcessBar.js"></script>

<html><head><title>upload_file_form</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="oa.css" rel=stylesheet>
</head>
<script language="JavaScript">
<!--

if (window.Event) 
document.captureEvents(Event.MOUSEUP); 

function nocontextmenu() 
{
event.cancelBubble = true
event.returnValue = false;

return false;
}

function norightclick(e) 
{
if (window.Event) 
{
if (e.which == 2 || e.which == 3)
return false;
}
else
if (event.button == 2 || event.button == 3)
{
event.cancelBubble = true
event.returnValue = false;
return false;
}

}

document.oncontextmenu = nocontextmenu; // for IE5+
document.onmousedown = norightclick; // for all others
//-->
</script>


<BODY bgColor=#ffffff leftMargin=0 
style="BACKGROUND-ATTACHMENT: scroll; BACKGROUND-IMAGE: url(images/main_bg.gif); BACKGROUND-POSITION: left bottom; BACKGROUND-REPEAT: no-repeat" 
topMargin=0>
<Script Language="javaScript">
    function  validate()
    {
       
        if  (document.myform.biaoti.value=="")
        {
            alert("说明不能为空");
            document.myform.biaoti.focus();
            return false ;
        }
        if  (document.myform.biaoti.value.length>100)
        {
            alert("字数超过规定范围");
            document.myform.biaoti.focus();
            return false ;
        }
        if  (document.myform.lianjie.value=="")
        {
            alert("文件连接不能为空");
            document.myform.lianjie.focus();
            return false ;
        }
     
}
</Script>

<form method="POST" action="file_add_db.asp" enctype="multipart/form-data" name=myform onsubmit="return validate()">
  <div align="center">
<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
  <TBODY>
  <TR>
    <TD bgColor=#4e5960 class=heading colSpan=2 height=3></TD></TR>
  <TR>
    <TD bgColor=#4e5960 class=heading>　<FONT 
color=#ffffff><B>上报文件</B></FONT></TD>
    <TD align=right bgColor=#4e5960 class=heading height=20></TD></TR>
  <TR>
    <TD align=middle vAlign=top width=109></TD>
    <TD align=middle>
        <TABLE bgColor=#666666 border=0 cellPadding=1 cellSpacing=1 
        width="100%">
      <TR> 
            <TD bgColor=#efefef height=20 width=60>&nbsp;单位名称</TD>
            <TD bgColor=#ffffff height=20> 
              <INPUT maxLength=100 name="Upart" 
            size=30 value="<%=Session("Upart")%>">
            </TD>
          </TR>
      <TR> 
            <TD bgColor=#efefef height=20 width=60 >&nbsp;报送人</TD>
            <TD bgColor=#ffffff height=20> 
              <INPUT maxLength=60 name="Rname" 
            size=30 value="<%=Session("Rname")%>">
            </TD>
          </TR>
      
      <TR> 
            <TD bgColor=#efefef height=20 width=60> 
              <div align="center">附  加<BR><BR>说  明<BR><BR><FONT COLOR="red">(字数不得超过100)</FONT></div>
            </TD>
            <TD bgColor=#ffffff height=20> 
              <TEXTAREA cols=40 name="biaoti" rows=6></TEXTAREA>
            </TD>
          </TR>
      <tr>
        <TD bgColor=#efefef height=20 width=60>&nbsp; 附件</TD>
        <td height="16" bgcolor="#ffffFF"><input type="file" name="lianjie" size="20" class="smallInput">&nbsp;&nbsp;<input type="submit" value="开始上传" name="B1" class="buttonface "  IsShowProcessBar="True"></td>
      </tr></table>
   
<%
 Response.Cookies("Type") = ""
 Response.Cookies("Type").Expires = "December 31, 2001"
 Response.Cookies("Type").Domain = ""
 Response.Cookies("Type").Path = "/www/home"
 Response.Cookies("Type").Secure = FALSE
%>

 </td></tr></table>
  </div></form></body></html>

