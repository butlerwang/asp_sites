<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<%
Set my_rs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from jhtdata where id="&request("view_id")
my_rs.open strSql,Conn,1,1 
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="oa.css" rel=stylesheet>
<title>�ڲ�֪ͨ��<%=my_rs("����")%></title>
</head>

<body style="background-image: url('images/main_bg.gif'); background-attachment: scroll; background-repeat: no-repeat; background-position: left bottom">

<div align="center">
  <center>
  <table border="1" width="100%" cellspacing="0" cellpadding=0 height="209" bordercolorlight=#000000 bordercolordark=#ffffff>
    <tr valign=top>
      <td width="21%" align="center" height="14">
        ��������</td>
      <td width="79%" height="14"><%=my_rs("����")%></td>
    </tr>
    <tr>
      <td width="21%" align="center" height="18">
        �� �� ��</td> 
      <td width="79%" height="18"><%=my_rs("����")%>��<%=my_rs("��ʵ����")%></td>
    </tr>
    <tr>
      <td width="21%" align="center" height="17">
        ����ʱ��</td>
      <td width="79%" height="17"><%=my_rs("ʱ��")%></td>
    </tr>
    <tr valign=top>
      <td width="21%" align="center" height="122">
        �ڡ�����</td>
      <td width="79%" height="122"><%=my_rs("����")%></td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>
<% my_rs.Close %>
<% my_Conn.Close %>
<% set my_Conn = nothing %>

