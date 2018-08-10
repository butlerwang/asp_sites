<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<!--#INCLUDE FILE="html.asp" -->
<%
name=request("name")
url=request("url")


if name="" or url="" then
%>
<HTML><HEAD><TITLE> 添加网址 </TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
</HEAD>

<BODY>
<FORM action="addurl.asp" method=post>
        <TABLE bgColor=#666666 border=0 cellPadding=1 cellSpacing=1 
        width="100%">
          <TBODY> 
          <TR> 
            <TD bgColor=#000000 class=heading colSpan=2><FONT 
            color=#ffffff><b>添加常用网址</b></FONT></TD>
          </TR>
          <TR> 
            <TD bgColor=#efefef height=20 width=60>&nbsp;网站名称 </TD>
            <TD bgColor=#ffffff height=20> 
              <INPUT maxLength=100 name="name" 
            size=30>
            </TD>
          </TR>
          <TR> 
            <TD bgColor=#efefef height=20 width=60 >&nbsp;网站地址</TD>
            <TD bgColor=#ffffff height=20> 
              http://<INPUT maxLength=60 name="url" 
            size=23>
            </TD>
          </TR>
		  <TR> 
            <TD bgColor=#efefef height=20 width=60 >&nbsp;网站说明 </TD>
            <TD bgColor=#ffffff height=20> 
                <TEXTAREA NAME="shuoming" ROWS="8" COLS="29"></TEXTAREA>
			</TD>
          </TR>
          
          <TR> 
            <TD bgColor=#efefef height=20 width=60> 
              &nbsp;
            </TD>
            <TD bgColor=#ffffff height=20> 
               <INPUT TYPE="image" SRC="images/add_off.gif">
            </TD>
          </TR>
        
          </TBODY> 
        </TABLE>
      </FORM>
</BODY>
</HTML>

<%
else
set rs=server.createobject("ADODB.recordset") 
rs.Open "SELECT * FROM url Where ID is null",conn,1,3 
rs.addnew

rs("网站名称")=name
rs("网址")=url
rs("网站说明")=htmlencode2(request("shuoming"))
rs.update 
rs.close
set rs=nothing
%>
<script language=javascript>
opener.location=opener.location;
</script>
<HTML><HEAD><TITLE>已经成功添加 </TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
</HEAD>

<BODY>
<TABLE bgColor=#666666 border=0 cellPadding=1 cellSpacing=1 
        width="100%">
          <TBODY> 
          <TR> 
            <TD bgColor=#000000 class=heading colSpan=2><FONT 
            color=#ffffff><b>已经成功添加</b></FONT></TD>
          </TR>
          <TR> 
            <TD bgColor=#efefef height=20 width=60>&nbsp;网站名称</TD>
            <TD bgColor=#ffffff height=20> 
              <%=name%>
            </TD>
          </TR>
          <TR> 
            <TD bgColor=#efefef height=20 width=60 >&nbsp;网站地址</TD>
            <TD bgColor=#ffffff height=20> 
             <%=url%>
            </TD>
          </TR>
		  <TR> 
            <TD bgColor=#efefef height=20 width=60 >&nbsp;网站说明</TD>
            <TD bgColor=#ffffff height=20> 
             <%=request("shuoming")%>
            </TD>
          </TR>
          
          <TR> 
            <TD bgColor=#efefef height=20 width=60> 
              &nbsp;
            </TD>
            <TD bgColor=#ffffff height=20> 
  <a  href="javaScript:window.close()"><img   border="0" src="images/close_1.gif"></a>     
            </TD>
          </TR>
        
          </TBODY> 
        </TABLE>
		</BODY>
</HTML>
<%end if%>