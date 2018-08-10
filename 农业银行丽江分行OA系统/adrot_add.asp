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
       
        if  (document.myform.src.value=="")
        {
            alert("软件名称不能为空");
            document.myform.src.focus();
            return false ;
        }
        if  (document.myform.content.value=="")
        {
            alert("说明不能为空");
            document.myform.content.focus();
            return false ;
        }
        if  (document.myform.height.value=="")
        {
            alert("高度不能为空");
            document.myform.height.focus();
            return false ;
        }
        if  (document.myform.width.value=="")
        {
            alert("宽度不能为空");
            document.myform.width.focus();
            return false ;
        }
     
}
</Script>

<form method="post" action="adrot_save.asp" enctype="multipart/form-data" name=myform  onSubmit='return validate()'>
  <div align="center">
        <TABLE border=1 bordercolorlight='000000' bordercolordark=#ffffff cellspacing=0 cellpadding=0 align=center>
      <TR> 
            <TD height=20 width=60>&nbsp;链接地址</TD>
            <TD height=20> 
              <INPUT name="url" 
            size=30>
            </TD>
          </TR>
      <TR> 
            <TD height=20 width=60> 
              <div align="center">广  告<BR><BR>说  明</div>
            </TD>
            <TD height=20> 
              <TEXTAREA cols=40 name="content" rows=6 class="txt" style="overflow:auto"></TEXTAREA>
            </TD>
          </TR>
      <TR> 
            <TD height=20 colspan=2> 
              &nbsp;广告类型:&nbsp; <SELECT NAME="type" style="height:18px;font-size:9pt"><option value="GIF" selected>GIF</option><option value="SWF">SWF</option></SELECT> 广告长宽：<INPUT TYPE="text" NAME="width" size=2 value="485">×<INPUT TYPE="text" NAME="height" size=1 value="75">
            </TD>
          </TR>
      <tr>
        <TD height=20 width=60>&nbsp;广告</TD>
        <td height="16"><input type="file" name="src" size="20" value="浏览">&nbsp;&nbsp;<input type="submit" value="上传" name="B1" class="txt " IsShowProcessBar="True"></td>
      </tr></table>
    
  </div></form></body></html>
