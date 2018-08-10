<%
if Session("Urule")<>"a" then
response.redirect "error.asp?id=admin"
end if
%>
<!--#include file="data.asp"-->
<html><head><title>添加文件</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="oa.css" rel=stylesheet>
</head>
<BODY>
<Script Language="javaScript">
    function  validate()
    {
       
        if  (document.myform.softname.value=="")
        {
            alert("软件名称不能为空");
            document.myform.softname.focus();
            return false ;
        }
        if  (document.myform.content.value=="")
        {
            alert("说明不能为空");
            document.myform.content.focus();
            return false ;
        }
        if  (document.myform.size.value=="")
        {
            alert("大小不能为空");
            document.myform.size.focus();
            return false ;
        }
        if  (document.myform.url.value=="")
        {
            alert("文件连接不能为空");
            document.myform.url.focus();
            return false ;
        }
     
}
</Script>

<form method="POST" action="addsoft_save.asp" name=myform  onSubmit='return validate()'>
  <div align="center">
        <TABLE border=1 bordercolorlight='000000' bordercolordark=#ffffff cellspacing=0 cellpadding=0 align=center>
	  <TR> 
            <TD height=20 width=60>&nbsp;软件名称</TD>
            <TD height=20> 
              <INPUT name="softname" 
            size=20 class="txt">
            </TD>
          </TR>
	  <TR> 
            <TD height=20 width=60>&nbsp;软件名称</TD>
            <TD height=20> 
              <INPUT name="size" 
            size=20 value="K">
            </TD>
          </TR>
      <tr>
        <TD height=20 width=60>&nbsp;下载路径</TD>
        <td height="16"><input name="url" size="42" class="txt"></td>
      </tr>
	  <TR> 
            <TD height=20 width=60> 
              <div align="center">软  件<BR><BR>简  介</div>
            </TD>
            <TD height=20> 
              <TEXTAREA cols=41 name="content" rows=6 class="txt" style="overflow:auto"></TEXTAREA>
            </TD>
          </TR>
     	  <TR> 
            <TD height=20 width=60> 
              <div align="center">&nbsp;是否推荐</div>
            </TD>
            <TD height=20>&nbsp; 是<INPUT TYPE="radio" NAME="best"  value="true"> 否<INPUT TYPE="radio" name="best" value="0" checked>
            </TD>
          </TR>
 <tr>
        <TD height=20 width=60>&nbsp; 提交表单</TD>
        <td height="16"><input type="submit" value="   提   交   " class="txt">  <INPUT TYPE="reset" value="   重   置   " class="txt"></td>
      </tr></table>
    
  </div></form></body></html>