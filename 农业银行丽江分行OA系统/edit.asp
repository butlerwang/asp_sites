<!--#include file="data.asp"-->
<!--#include file="check.asp"-->
<%if session("Urule")<>"a" then
respons.write "��û���㹻��Ȩ��:P"
respons.end
end if
%>
<%
Set myrs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from bumen"
myrs.open strSql,Conn,1,1 
%>

<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
<title>�޸��û�����</title>
</head>

<body>
<script Language="javaScript">
    function  validate()
    {
        if  (document.myform.name.value=="")
        {
            alert("��������Ϊ��");
            document.myform.name.focus();
            return false ;
        }
        if  (document.myform.Userid.value=="")
        {
            alert("��¼�ʺŲ���Ϊ��");
            document.myform.Userid.focus();
            return false ;
        }
        if  (document.myform.company.value=="")
        {
            alert("�������Ʋ���Ϊ��");
            document.myform.company.focus();
            return false ;
        }
        if  (document.myform.tel.value=="")
        {
            alert("�绰���벻��Ϊ��");
            document.myform.tel.focus();
            return false ;
        }
        if  (document.myform.email.value=="")
        {
            alert("�����ʼ�����Ϊ��");
            document.myform.email.focus();
            return false ;
        }
        if  (document.myform.password.value=="")
        {
            alert("���벻��Ϊ��");
            document.myform.password.focus();
            return false ;
         }
        if  (document.myform.ilevel.value=="")
        {
            alert("���伶����Ϊ��");
            document.myform.ilevel.focus();
            return false ;
	}
        return  true;
    }
</script> <link rel="stylesheet" href="eintrdemo.css"> </head> <%
dim sql
dim rs
 sql="select * from user where id="&request("id")
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
                %> <body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"> 
<form method="post" action="saveedit.asp?id=<%=request("id")%>" name="myform" onsubmit="return  validate()"> 
<table width="100%" border="0" cellspacing="1" cellpadding="2"> <tr > <td class="heading"> 
<div align="center"> <center> <table width="81%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000" bordercolorlight="#000000" style="font-size:9pt"> 
<tr> <td width="2%" align="right"></td><td align="left" height="25"> <p align="center"><font color="#FFFFFF"><b>�� 
�� �� ��</b></font></p></td><td width="3%"></td></tr> </table></center></div></td></tr> 
</table><div align="center"> <table width="80%" border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF" bordercolor="#FFFFFF" bordercolorlight="#000000" style="font-size:9pt"> 
<tr> <td width="17%" valign="top"> <p align="right">�������:</p></td><td width="83%"> 
<input type="text" name="name" class="form" value="<%=rs("����")%>" size="24"> </td></tr> 
<tr> <td width="17%" valign="top" height="6"> <p align="right">��¼�ʺ�:</p></td><td width="83%" height="6"> 
	    <input type="hidden" name="Userid"  value="<%=rs("�û���")%>"  >
        <input type="text" name="Userid2" class="form" value="<%=rs("�û���")%>" size="24" disabled>


</td></tr> <tr> <td width="17%"  valign="top" height="16"> <p align="right">��¼����:</p></td><td width="83%" height="16"> 
<input type="password" name="password" class="form" size="24" value="<%=rs("����")%>"> 
</td></tr> <tr> <td width="17%"  valign="top" height="16"> <p align="right">��������:</p></td><td width="83%" height="16"> 
<input type="text" name="question" class="form" size="24" value="<%=rs("����")%>"> 
</td></tr> <tr> <td width="17%"  valign="top" height="16"> <p align="right">�����:</p></td><td width="83%" height="16"> 
<input type="text" name="answer" class="form" size="24" value="<%=rs("��")%>"> 
</td></tr> 
<tr> 
<td width="17%"  valign="top"> 
<p align="right">��������:</td>
<td width="83%"> 
<select NAME="company"> <%if myrs.eof and myrs.bof then
response.write "<font color='red'>��û���κ�����</font>"
else

do while not (myrs.eof or myrs.bof)
if myrs("type")=rs("����") then
sel="selected"
else 
sel=""
end if
%> <option value="<%=myrs("type")%>" <%=sel%>><%=myrs("type")%></option> <%myrs.movenext 
loop 
end if%> </select> </td></tr> 
<tr> 
<td width="17%" valign="top"> 
<p align="right">���伶��:</p>
</td>
<td width="83%"> 
<input type="text" name="ilevel" class="from" value="<%=rs("ilevel")%>" size="1"> 
</td>
</tr> 
<tr> 
<td width="17%"  valign="top"> 
<p align="right">�绰����:</p>
</td>
<td width="83%"> 
<input type="text" name="tel" class="form" value="<%=rs("�绰")%>" size="24"> </td></tr> 
<tr> <td width="17%"  valign="top"> <p align="right">�����ʼ�:</p></td><td width="83%"> 
<input type="text" name="email" class="form" value="<%=rs("����")%>" size="24"> 
</td></tr> <tr> <td width="17%"  valign="top"> <p align="right">����Ȩ��:</p></td><td width="83%"> 
<select NAME="admin"> <option value="a" <%if rs("Ȩ��")="a" then%>selected<%end if%>>�����û�</option> 
<option value="b" <%if rs("Ȩ��")="b" then%>selected<%end if%>>����Ա</option> <option value="c" <%if rs("Ȩ��")="c" then%>selected<%end if%>>��ͨ�û�</option> 
</select> </td></tr> 
</table>
</div><div align="center"><input type=image  src="images/modify_off.gif">&nbsp;&nbsp; 
<a  href="javaScript:window.close()"><img   border="0" src="images/close_1.gif"></a> 
</div></form>     


</body>
</html>
