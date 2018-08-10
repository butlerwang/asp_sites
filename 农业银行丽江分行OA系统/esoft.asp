<%
if session("Urule")<>"a" then
response.redirect "error.asp?id=admin"
end if
%>
<!--#include file="html.asp"-->
<!--#include file="data.asp"-->
<%
	set rs=server.createobject("adodb.recordset")
	sql="select * from soft where id="&request("id")
	rs.open sql,conn,3,3
if request("edit")="" then
%>
<html><head><title>修改软件</title>
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
        if  (document.myform.size.value=="")
        {
            alert("软件大小不能为空");
            document.myform.size.focus();
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

<form method="POST" action="esoft.asp" name=myform  onSubmit='return validate()'>
  <div align="center"><INPUT TYPE="hidden" name="id" value="<%=rs("id")%>">
        <TABLE border=1 bordercolorlight='000000' bordercolordark=#ffffff cellspacing=0 cellpadding=0 
        width="400" align=center>
	  <TR> 
            <TD height=20 width=60>&nbsp;软件名称</TD>
            <TD height=20> 
              <INPUT name="softname" 
            size=20 class="txt" value="<%=rs("name")%>">
            </TD>
          </TR>
	  <TR> 
            <TD height=20 width=60>&nbsp;软件大小</TD>
            <TD height=20> 
              <INPUT name="size" 
            size=20 class="txt" value="<%=rs("size")%>">
            </TD>
          </TR>
      <tr>
        <TD height=20 width=60>&nbsp;下载路径</TD>
        <td height="16"><input name="url" size="42" class="txt" value="<%=rs("url")%>"></td>
      </tr>
	  <TR> 
            <TD height=20 width=60> 
              <div align="center">软  件<BR><BR>简  介</div>
            </TD>
            <TD height=20> 
              <TEXTAREA cols=41 name="content" rows=6 class="txt" style="overflow:auto"><%=replace(replace(rs("content"),"<br>",chr(13)),"&nbsp;"," ")%></TEXTAREA>
            </TD>
          </TR>
	  <TR> 
            <TD height=20 width=60> 
              <div align="center">&nbsp;是否推荐</div>
            </TD>
            <TD height=20>&nbsp; 是<INPUT TYPE="radio" NAME="best"  value="true"<%if rs("best")=true then response.write " checked" end if%>> 否<INPUT TYPE="radio" name="best" value="0" <%if rs("best")=0 then response.write " checked" end if%>>
            </TD>
          </TR>
      <tr>
        <TD height=20 width=60>&nbsp; 提交表单</TD>
        <td height="16"><input type="submit" value="   提   交   " class="txt" name="edit">  <INPUT TYPE="reset" value="   重   置   " class="txt"></td>
      </tr></table>
    
  </div></form></body></html>
<%else 
		 
	rs("name") =htmlencode2(request("softname"))
	rs("content") = htmlencode2(request("content"))
	rs("url") =request("url")
	rs("size")=request("size")
	rs("best")=request("best")
	rs.update
%>

<LINK href="oa.css" rel=stylesheet>
<script language=javascript>
opener.location=opener.location
</script>
<BODY>
        <TABLE border=1 bordercolorlight='000000' bordercolordark=#ffffff cellspacing=0 cellpadding=0 align=center>
	  <TR> 
            <TD>文件已经成功修改，是否继续修改……<BR>
<P><P><A HREF="esoft.asp">继续修改</A>&nbsp;&nbsp;<A HREF="javascript:window.close()">关闭窗口</A></TD>
            
      </TR>
	  
      </table>
<%
end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
