<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="mouse.js" -->
<%
Set rs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from bumen"
rs.open strSql,Conn,1,1 
%>
<html><head><title>������������칫ϵͳ----�����ʺ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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
            alert("��λ���Ʋ���Ϊ��");
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
		if  (document.myform.question.value=="")
        {
            alert("�������ⲻ��Ϊ��");
            document.myform.question.focus();
            return false ;
        }
		if  (document.myform.answer.value=="")
        {
            alert("����𰸲���Ϊ��");
            document.myform.answer.focus();
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
            document.myform.password.focus();
            return false ;
        }
        return  true;
    }
</script>

<link rel="stylesheet" href="oa.css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form method="post" action="zhuce.asp" name="myform" onsubmit="return  validate()"> 
<table width="100%" border="0" cellspacing="1" cellpadding="2"> <tr > <td class="heading"> 
<div align="center"> <center> <table width="81%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000" bordercolorlight="#000000"> 
<tr> <td width="2%" align="right"></td><td align="left" height="25"> <p align="center"><font color="#FFFFFF"><b>�� 
�� �� ��</b></font></p></td><td width="3%"></td></tr> </table></center></div></td></tr> 
</table><div align="center"> <table WIDTH="80%" BORDER="1" CELLSPACING="0" CELLPADDING="0" BORDERCOLORDARK="#FFFFFF" BORDERCOLOR="#FFFFFF" BORDERCOLORLIGHT="#000000"> 
<tr> <td WIDTH="17%" VALIGN="top"> <p ALIGN="right">�������:</p></td><td WIDTH="83%"> 
<input TYPE="text" NAME="name" CLASS="form" SIZE="24">[�����������] </td></tr> <tr> <td WIDTH="17%" VALIGN="top" HEIGHT="6"> 
<p ALIGN="right">��¼����:</p></td><td WIDTH="83%" HEIGHT="6"> <input TYPE="text" NAME="Userid" CLASS="form" SIZE="24">[�����������]<br></td></tr> 
<tr> <td WIDTH="17%"  VALIGN="top" HEIGHT="16"> <p ALIGN="right">��¼����:</p></td><td WIDTH="83%" HEIGHT="16"> 
<input TYPE="password" NAME="password" CLASS="form" SIZE="24"> [���μ��������]</td></tr> <tr> 
<td WIDTH="17%"  VALIGN="top" HEIGHT="16"> <p ALIGN="right">��������:</p></td><td WIDTH="83%" HEIGHT="16"> 
<input TYPE="text" NAME="question" CLASS="form" SIZE="24" value=���ù�>  [���Բ��ù�]</td></tr> <tr> <td WIDTH="17%"  VALIGN="top" HEIGHT="16"> 
<p ALIGN="right">�����:</p></td><td WIDTH="83%" HEIGHT="16"> <input TYPE="text" NAME="answer" CLASS="form" SIZE="24" value=���ù�> [���Բ��ù�]
</td></tr> <tr> <td WIDTH="17%"  VALIGN="top"> <p ALIGN="right">��λ����: </td><td WIDTH="83%"> 
<select NAME="company"> <option selected> --==��λ����==--</option> <%if rs.eof and rs.bof then
response.write "<font color='red'>��û���κζ���</font>"
else

do while not (rs.eof or rs.bof)
%> <option VALUE="<%=rs("type")%>"><%=rs("type")%></option> <%rs.movenext 
loop 
end if%> </select> </td></tr> <tr> <td WIDTH="17%"  VALIGN="top"> <p ALIGN="right">�ֻ�����:</p></td><td WIDTH="83%"> 
<input TYPE="text" NAME="mobile" CLASS="form" SIZE="24">Ϊ����ϵ��������д </td></tr> <tr> <td WIDTH="17%"  VALIGN="top" HEIGHT="27"> 
<p ALIGN="right">�绰����:</p></td><td WIDTH="83%" HEIGHT="27"> <input TYPE="text" NAME="tel" CLASS="form" SIZE="24"> 
Ϊ����ϵ��������д</td></tr><tr> <td WIDTH="17%"  VALIGN="top"> <p ALIGN="right">�����ʼ�:</p></td><td WIDTH="83%"> 
<p><input TYPE="text" NAME="email" CLASS="form" SIZE="24" value="XX@LJ.YN.ABC"></p></td></tr> <%if session("id")<>"" then %> 
<tr> <td WIDTH="17%"  VALIGN="top"> <p ALIGN="right">���伶��:</p></td><td WIDTH="83%"> 
<p><input TYPE="text" NAME="ilevel" CLASS="form" SIZE="1"></p></td></tr> <tr> 
<td width="17%"  valign="top"> <p align="right">����Ȩ��:</p></td><td width="83%"> 
<select NAME="admin"> <option value="a" <%if rs("Ȩ��")="a" then%>selected<%end if%>>�����û�</option> 
<option value="b" <%if rs("Ȩ��")="b" then%>selected<%end if%>>����Ա</option> <option value="c" <%if rs("Ȩ��")="c" then%>selected<%end if%>>��ͨ�û�</option> 
</select> </td></tr> <%end if%> </table></div><div align="center"><input type=image  src="images/add_off.gif">&nbsp;&nbsp; 
<a  href="javaScript:window.close()"><img   border="0" src="images/close_1.gif"></a> 
</div></form><div align="center"> <center> <table border="1" width="80%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF" bordercolor="#C0C0C0" HEIGHT="102"> 
<tr> <td width="100%" HEIGHT="101"> ϵͳ������:<FONT COLOR="#FF0000">����������ʵ��������ע�ᣬ����������ˣ���δ�����ͨ���ģ������ܽ���ϵͳ��ʹ���κι��ܣ�</FONT> <ul><li>���������<font color="#FF0000">����</font>����&nbsp;</li><li>��¼������<font color="#FF0000">����</font>����&nbsp;</li><li>��¼���룺��<font color="#FF0000">����</font>��</li><li>��λ���ƣ�<font color="#FF0000">����</font>��</li><li>�ֻ�����</li><li>�����ʼ�</li></ul></td></tr> 
</table></center></div>
       
</body>       
</html>

