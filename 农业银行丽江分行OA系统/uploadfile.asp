<%
if session("Urule")<>"a" then
response.redirect "error.asp?id=admin"
end if
%>
<!--#include file="data.asp"-->
<script language="javascript" src="ShowProcessBar.js"></script>

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
        if  (document.myform.url.value=="")
        {
            alert("文件连接不能为空");
            document.myform.url.focus();
            return false ;
        }
     
}
</Script>

<form method="post" action="softadd.asp" enctype="multipart/form-data" name=myform  onSubmit='return validate()'>
  <div align="center">
        <TABLE border=1 bordercolorlight='000000' bordercolordark=#ffffff cellspacing=0 cellpadding=0 align=center>
	  <TR> 
            <TD height=20 width=60>&nbsp;软件名称</TD>
            <TD height=20> 
              <INPUT name="softname" 
            size=30 class="txt">
            </TD>
          </TR>
	  <TR> 
            <TD height=20 width=60> 
              <div align="center">软  件<BR><BR>简  介</div>
            </TD>
            <TD height=20> 
              <TEXTAREA cols=40 name="content" rows=6 class="txt" style="overflow:auto"></TEXTAREA>
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
        <TD height=20 width=60>&nbsp; 软件</TD>
        <td height="16"><input type="file" name="url" size="20" class="txt">&nbsp;&nbsp;<input type="submit" value="开始上传" name="B1" class="txt " IsShowProcessBar="True"></td>
      </tr></table>
    
  </div></form></body></html>