<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<%if request("name")="" or request("phone")="" then%>
<HTML><HEAD><TITLE> ��ӵ绰 </TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
</HEAD>
<BODY>
<FORM action="addphone.asp" method=post>
        <TABLE bgColor=#666666 border=0 cellPadding=1 cellSpacing=1 
        width="100%">
          <TBODY> 
          <TR> 
            <TD bgColor=#000000 class=heading colSpan=2><FONT 
            color=#ffffff><b>��ӵ绰����</b></FONT></TD>
          </TR>
          <TR> 
            <TD bgColor=#efefef height=20 width=60>&nbsp;�绰����</TD>
            <TD bgColor=#ffffff height=20> 
              <INPUT maxLength=100 name="name" 
            size=30>
            </TD>
          </TR>
          <TR> 
            <TD bgColor=#efefef height=20 width=60 >&nbsp;�绰����</TD>
            <TD bgColor=#ffffff height=20> 
              <INPUT maxLength=60 name="phone" 
            size=30>
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
name=request("name")
phone=request("phone")

set rs=server.createobject("ADODB.recordset") 
rs.Open "SELECT * FROM tel Where ID is null",conn,1,3 
rs.addnew

rs("�绰����")=name
rs("�绰����")=phone
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
            <TD bgColor=#efefef height=20 width=60>&nbsp;�绰����</TD>
            <TD bgColor=#ffffff height=20> 
              <%=name%>
            </TD>
          </TR>
          <TR> 
            <TD bgColor=#efefef height=20 width=60 >&nbsp;�绰����</TD>
            <TD bgColor=#ffffff height=20> 
             <%=phone%>
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