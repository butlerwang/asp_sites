<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<!--#INCLUDE FILE="html.asp" -->
<%
name=request("name")
url=request("url")


if name="" or url="" then
%>
<HTML><HEAD><TITLE> �����ַ </TITLE>
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
            color=#ffffff><b>��ӳ�����ַ</b></FONT></TD>
          </TR>
          <TR> 
            <TD bgColor=#efefef height=20 width=60>&nbsp;��վ���� </TD>
            <TD bgColor=#ffffff height=20> 
              <INPUT maxLength=100 name="name" 
            size=30>
            </TD>
          </TR>
          <TR> 
            <TD bgColor=#efefef height=20 width=60 >&nbsp;��վ��ַ</TD>
            <TD bgColor=#ffffff height=20> 
              http://<INPUT maxLength=60 name="url" 
            size=23>
            </TD>
          </TR>
		  <TR> 
            <TD bgColor=#efefef height=20 width=60 >&nbsp;��վ˵�� </TD>
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

rs("��վ����")=name
rs("��ַ")=url
rs("��վ˵��")=htmlencode2(request("shuoming"))
rs.update 
rs.close
set rs=nothing
%>
<script language=javascript>
opener.location=opener.location;
</script>
<HTML><HEAD><TITLE>�Ѿ��ɹ���� </TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
</HEAD>

<BODY>
<TABLE bgColor=#666666 border=0 cellPadding=1 cellSpacing=1 
        width="100%">
          <TBODY> 
          <TR> 
            <TD bgColor=#000000 class=heading colSpan=2><FONT 
            color=#ffffff><b>�Ѿ��ɹ����</b></FONT></TD>
          </TR>
          <TR> 
            <TD bgColor=#efefef height=20 width=60>&nbsp;��վ����</TD>
            <TD bgColor=#ffffff height=20> 
              <%=name%>
            </TD>
          </TR>
          <TR> 
            <TD bgColor=#efefef height=20 width=60 >&nbsp;��վ��ַ</TD>
            <TD bgColor=#ffffff height=20> 
             <%=url%>
            </TD>
          </TR>
		  <TR> 
            <TD bgColor=#efefef height=20 width=60 >&nbsp;��վ˵��</TD>
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