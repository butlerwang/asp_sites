<!--#include file="data.asp"-->
<!--#include file="html.asp"-->
<!--#INCLUDE FILE="mouse.js" -->
<%
Set myrs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from bumen"
myrs.open strSql,Conn,1,1 
 name=htmlencode2(request("name"))
 password=request("password")
 userid=htmlencode2(request("userid"))
 question=htmlencode2(request("question"))
 answer=htmlencode2(request("answer"))
 email=request("email")
 mobile=request("mobile")
 tel=request("tel")
 ilevel=request("ilevel")
 department=htmlencode2(request("company"))
 ip= Request.ServerVariables("REMOTE_ADDR")
 nowtime=now()
sj=cstr(year(nowtime))+"-"+cstr(month(nowtime))+"-"+cstr(day(nowtime))+" "+cstr(hour(nowtime))+":"+right("0"+cstr(minute(nowtime)),2)+":"+right("0"+cstr(second(nowtime)),2)

set rs=server.createobject("ADODB.recordset")
rs.open "select * from user where �û���='"& userid &"'order by id",conn,3,3
if rs.eof or rs.bof then
 else if userid=rs("�û���") then
  userid=""
  password=""
  %>
<link rel="stylesheet" href="oa.css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2"> <tr > <td class="heading"> 
<div align="center"> <center> <table width="81%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000" bordercolorlight="#000000"> 
<tr> <td width="2%" align="right"></td><td align="left" height="25"> <p align="center"><font color="#FFFFFF"><b>���Ѿ��ɹ������ʺ�</b></font></p></td><td width="3%"></td></tr> 
</table></center></div></td></tr> </table><div align="center"> <form method="post" action="saveedit1.asp?id=<%=id%>" name="myform" onsubmit="return  validate()"> 
<table width="80%" border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF" bordercolor="#FFFFFF" bordercolorlight="#000000"> 
<tr> <td> <font COLOR="red">���ʺ��Ѿ�����</font> </td></tr> </table></div><div align="center"><a  href="javascript:history.back(1)"><img border="0" src="images/previous.gif"></a>&nbsp;&nbsp; 
<a  href="javaScript:window.close()"><img   border="0" src="images/close_1.gif"></a> 
</div></form> <div align="center"> <center> <table border="1" width="80%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF" bordercolor="#C0C0C0"> 
<tr> <td width="100%"> ϵͳ������: <ul> <li>���������<font color="#FF0000">����</font>��</li><li>��¼�ʺţ�<font color="#FF0000">����</font>��</li><li>��¼���루<font COLOR="#FF0000">����</font>��</li><li>���伶��<font COLOR="#FF0000">����</font>��</li><li>��˾���ƣ�<font color="#FF0000">����</font>��</li><li>�����ʼ�����ʾ�������д�����Զ���д�������ʼ���ַ�������ʼ���</li></ul></td></tr> 
</table></center></div><%
  response.end
end if
end if
rs.close


set rs=server.createobject("ADODB.recordset") 
rs.Open "SELECT * FROM user Where ID is null",conn,1,3 
rs.addnew

rs("�û���")=userid
rs("����")=password
rs("����")=email
rs("����")=department
rs("����")=question
rs("��")=answer
rs("Ȩ��")=request("admin")
rs("���")=false
rs("ʱ��")=sj
rs("IP")=ip
rs("�绰")=tel
rs("����")=name
rs("mobile")=mobile
rs("ilevel")=ilevel
if rs("ilevel")="" then rs("ilevel")="1"
if rs("Ȩ��")="" then rs("Ȩ��")="c"
rs.update 
id=rs("id")
%> <title>���Ѿ��ɹ������ʺ�</title> <meta http-equiv="Content-Type" content="text/html; charset=gb2312"> 
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
<table width="100%" border="0" cellspacing="1" cellpadding="2"> 
<tr > <td class="heading"> <div align="center"> <center> <table width="81%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000" bordercolorlight="#000000"> 
<tr> <td width="2%" align="right"></td><td align="left" height="25"> <p align="center"><font color="#FFFFFF"><b>���Ѿ��ɹ������ʺ�</b></font></p></td><td width="3%"></td></tr> 
</table></center></div></td></tr> </table><div align="center">
 <form method="post" action="saveedit1.asp?id=<%=id%>" name="myform" onsubmit="return  validate()" > 
<table width="80%" border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF" bordercolor="#FFFFFF" bordercolorlight="#000000"> 
<tr> <td width="17%" valign="top"> <p align="right">�������:</p></td><td width="83%"> 
<input type="text" name="name" class="form" value="<%=rs("����")%>" size="24"> </td></tr> 
<tr> <td width="17%" valign="top" height="6"> <p align="right"><font size="2">��¼�ʺ�:</font></p></td><td width="83%" height="6"> 
	    <input type="hidden" name="Userid"  value="<%=rs("�û���")%>"  >
        <input type="text" name="Userid2" class="form" value="<%=rs("�û���")%>" size="24" disabled>


</td></tr> <tr> <td width="17%"  valign="top" height="16"> <p align="right"><font size="2">��¼����:</font></p></td><td width="83%" height="16"> 
<input type="password" name="password" class="form" size="24" value="<%=rs("����")%>"> 
</td></tr> <tr> <td width="17%"  valign="top" height="16"> <p align="right"><font size="2">��������:</font></p></td><td width="83%" height="16"> 
<input type="text" name="question" class="form" size="24" value="<%=rs("����")%>"> 
</td></tr> <tr> <td width="17%"  valign="top" height="16"> <p align="right"><font size="2">�����:</font></p></td><td width="83%" height="16"> 
<input type="text" name="answer" class="form" size="24" value="<%=rs("��")%>"> 
</td></tr> <tr> <td width="17%"  valign="top"> <p align="right"><font size="2">��λ����:</font> 
</td><td width="83%"> <select NAME="company">

 <%if myrs.eof and myrs.bof then
response.write "<font color='red'>��û���κζ���</font>"
else

do while not (myrs.eof or myrs.bof)
if myrs("type")=rs("����") then
sel="selected"
else 
sel=""
end if
%> <option value="<%=myrs("type")%>" <%=sel%>><%=myrs("type")%></option> <%myrs.movenext 
loop 
end if%> </select> </td></tr> <tr> <td width="17%"  valign="top"> <p align="right"><font size="2">�ֻ�����:</font></p></td><td width="83%"> 
<input type="text" name="tel" class="form" value="<%=rs("mobile")%>" size="24"> 
</td></tr> <tr> <td width="17%"  valign="top"> <p align="right"><font size="2">�绰����:</font></p></td><td width="83%"> 
<input type="text" name="tel" class="form" value="<%=rs("�绰")%>" size="24"> </td></tr> 
<tr> <td width="17%"  valign="top"> <p align="right"><font size="2">�����ʼ�:</font></p></td><td width="83%"> 
<input type="text" name="email" class="form" value="<%=rs("����")%>" size="24"> 
</td></tr>

<%if session("id")<>"" then %>
 <tr> 
<td WIDTH="17%"  VALIGN="top"> <p ALIGN="right">���伶��:</p></td><td WIDTH="83%"> 
<p><input TYPE="text" NAME="ilevel" CLASS="form" SIZE="1" value="<%=rs("ilevel")%>"></p></td></tr>
<tr> <td width="17%"  valign="top"> <p align="right">����Ȩ��:</p></td><td width="83%"> 
<select NAME="admin"> <option value="a" <%if rs("Ȩ��")="a" then%>selected<%end if%>>�����û�</option> 
<option value="b" <%if rs("Ȩ��")="b" then%>selected<%end if%>>����Ա</option> <option value="c" <%if rs("Ȩ��")="c" then%>selected<%end if%>>��ͨ�û�</option> 
</select> </td></tr> 
<%end if%> 

</table></div><div align="center"><input type=image  src="images/modify_off.gif">&nbsp;&nbsp; 
<a  href="javaScript:window.close()"><img   border="0" src="images/close_1.gif"></a> 
</div></form> <div align="center"> <center> <table border="1" width="80%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF" bordercolor="#C0C0C0"> 
<tr> <td width="100%"> ϵͳ������: <ul> <li>���������<font color="#FF0000">����</font>��</li><li>��¼�ʺţ�<font color="#FF0000">����</font>��</li><li>��¼���루<font color="#FF0000">����</font>��</li><li>���伶��<font COLOR="#FF0000">����</font>��</li><li>��˾���ƣ�<font color="#FF0000">����</font>��</li><li>�����ʼ�����ʾ�������д�����Զ���д�������ʼ���ַ�������ʼ���</li></ul></td></tr> 
</table></center></div>       


</body>
</html>
<%rs.close
set rs=nothing
%>