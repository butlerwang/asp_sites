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
<html><head><title>�޸����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="oa.css" rel=stylesheet>
</head>
<BODY>
<Script Language="javaScript">
    function  validate()
    {
       
        if  (document.myform.softname.value=="")
        {
            alert("������Ʋ���Ϊ��");
            document.myform.softname.focus();
            return false ;
        }
        if  (document.myform.size.value=="")
        {
            alert("�����С����Ϊ��");
            document.myform.size.focus();
            return false ;
        }
        if  (document.myform.content.value=="")
        {
            alert("˵������Ϊ��");
            document.myform.content.focus();
            return false ;
        }
        if  (document.myform.url.value=="")
        {
            alert("�ļ����Ӳ���Ϊ��");
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
            <TD height=20 width=60>&nbsp;�������</TD>
            <TD height=20> 
              <INPUT name="softname" 
            size=20 class="txt" value="<%=rs("name")%>">
            </TD>
          </TR>
	  <TR> 
            <TD height=20 width=60>&nbsp;�����С</TD>
            <TD height=20> 
              <INPUT name="size" 
            size=20 class="txt" value="<%=rs("size")%>">
            </TD>
          </TR>
      <tr>
        <TD height=20 width=60>&nbsp;����·��</TD>
        <td height="16"><input name="url" size="42" class="txt" value="<%=rs("url")%>"></td>
      </tr>
	  <TR> 
            <TD height=20 width=60> 
              <div align="center">��  ��<BR><BR>��  ��</div>
            </TD>
            <TD height=20> 
              <TEXTAREA cols=41 name="content" rows=6 class="txt" style="overflow:auto"><%=replace(replace(rs("content"),"<br>",chr(13)),"&nbsp;"," ")%></TEXTAREA>
            </TD>
          </TR>
	  <TR> 
            <TD height=20 width=60> 
              <div align="center">&nbsp;�Ƿ��Ƽ�</div>
            </TD>
            <TD height=20>&nbsp; ��<INPUT TYPE="radio" NAME="best"  value="true"<%if rs("best")=true then response.write " checked" end if%>> ��<INPUT TYPE="radio" name="best" value="0" <%if rs("best")=0 then response.write " checked" end if%>>
            </TD>
          </TR>
      <tr>
        <TD height=20 width=60>&nbsp; �ύ��</TD>
        <td height="16"><input type="submit" value="   ��   ��   " class="txt" name="edit">  <INPUT TYPE="reset" value="   ��   ��   " class="txt"></td>
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
            <TD>�ļ��Ѿ��ɹ��޸ģ��Ƿ�����޸ġ���<BR>
<P><P><A HREF="esoft.asp">�����޸�</A>&nbsp;&nbsp;<A HREF="javascript:window.close()">�رմ���</A></TD>
            
      </TR>
	  
      </table>
<%
end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
