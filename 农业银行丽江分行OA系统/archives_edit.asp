<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<%

Set rs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from userinfo where userid="&session("Uid")
rs.open strSql,Conn,1,1 
if rs.eof then
response.write "no record"
end if
check=split(rs("check"), ",", -1, 1)

%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
</head>
<title>���˵���</title>  
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" style="background-image: url('images/main_bg.gif'); background-attachment: scroll; background-repeat: no-repeat; background-position: left bottom" onLoad="MM_preloadImages('images/more_on.gif')" >
<form  method="post" action="archives_save.asp">
<table border="1" cellpadding="0" cellspacing="0" width="95%" bordercolorlight=#000000 bordercolordark=#ffffff align=right>
  <tr> 
    <td align="center" width="15%"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td width="30%">&nbsp;<%=session("Rname")%></td>
    <td align="center" width="15%"><b>��&nbsp;��&nbsp;��</b></td>
    <td width="25%"><INPUT TYPE="text" name="Uname" value="<%=rs("Uname")%>" size=10><input type="radio" name="c1" value="yes" <% if check("0")="yes" then response.write "checked"%>>Y<input type="radio" name="c1" value="no" <% if check("0")="no" then response.write "checked"%>>N</td>
    </tr>
  <tr> 
    <td align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td><SELECT NAME="sex"><option value="��" <%if rs("sex")="��" then response.write " selected" end if%>>��</option><option value="Ů" <%if rs("sex")="Ů" then response.write " selected" end if%>>Ů</option></SELECT>
        ����:<input type="radio" name="c2" value="yes" <% if check("1")="yes" then response.write "checked"%>>Y<input type="radio" name="c2" value="no" <% if check("1")="no" then response.write "checked"%>>N</td>
    <td align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("nation")%>" name=nation size="10"><input type="radio" name="c3" value="yes" <% if check("2")="yes" then response.write "checked"%>>Y<input type="radio" name="c3" value="no" <% if check("2")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td align="center"><b>��������</b></td>
    <td>&nbsp;<%=Session("Upart")%></td>
    <td align="center"><b>ְ&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("duty")%>" name="duty" size="10"><input type="radio" name="c4" value="yes" <% if check("3")="yes" then response.write "checked"%>>Y<input type="radio" name="c4" value="no" <% if check("3")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td align="center"><b>ְ&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("grade")%>" name="grade" size="10"><input type="radio" name="c5" value="yes" <% if check("4")="yes" then response.write "checked"%>>Y<input type="radio" name="c5" value="no" <% if check("4")="no" then response.write "checked"%>>N
      </td>
    <td align="center"><b>��������</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("birthday")%>" name="birthday" size="10"><input type="radio" name="c6" value="yes" <% if check("5")="yes" then response.write "checked"%>>Y<input type="radio" name="c6" value="no" <% if check("5")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td align="center"><b>������ò</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("polity")%>" name="polity" size="10"><input type="radio" name="c7" value="yes" <% if check("6")="yes" then response.write "checked"%>>Y<input type="radio" name="c7" value="no" <% if check("6")="no" then response.write "checked"%>>N
      </td>
    <td align="center"><b>����״��</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("health")%>" name="health" size="10"><input type="radio" name="c8" value="yes" <% if check("7")="yes" then response.write "checked"%>>Y<input type="radio" name="c8" value="no" <% if check("7")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td align="center" height="20"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("Nplace")%>" name="Nplace" size="10"><input type="radio" name="c9" value="yes" <% if check("8")="yes" then response.write "checked"%>>Y<input type="radio" name="c9" value="no" <% if check("8")="no" then response.write "checked"%>>N
      </td>
    <td align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("weight")%>" name="weight" size="10"><input type="radio" name="c10" value="yes" <% if check("9")="yes" then response.write "checked"%>>Y<input type="radio" name="c10" value="no" <% if check("9")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>���֤��</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("idcard")%>" name="idcard" size="10"><input type="radio" name="c11" value="yes" <% if check("10")="yes" then response.write "checked"%>>Y<input type="radio" name="c11" value="no" <% if check("10")="no" then response.write "checked"%>>N
      </td>
    <td align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("height")%>" name="height" size="10"><input type="radio" name="c12" value="yes" <% if check("11")="yes" then response.write "checked"%>>Y<input type="radio" name="c12" value="no" <% if check("11")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td align="center" height="20"><b>����״��</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("marriage")%>" name="marriage" size="10"><input type="radio" name="c13" value="yes" <% if check("12")="yes" then response.write "checked"%>>Y<input type="radio" name="c13" value="no" <% if check("12")="no" then response.write "checked"%>>N
      </td>
    <td align="center"><b>��ҵԺУ</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("Fschool")%>" name="Fschool" size="10"><input type="radio" name="c14" value="yes" <% if check("13")="yes" then response.write "checked"%>>Y<input type="radio" name="c14" value="no" <% if check("13")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td align="center" height="20"><b>���˳ɷ�</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("member")%>" name="member" size="10"><input type="radio" name="c15" value="yes" <% if check("14")="yes" then response.write "checked"%>>Y<input type="radio" name="c15" value="no" <% if check("14")="no" then response.write "checked"%>>N
      </td>
    <td align="center"><b>ר&nbsp;&nbsp;&nbsp;&nbsp;ҵ</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("speciality")%>" name="speciality" size="10"><input type="radio" name="c16" value="yes" <% if check("15")="yes" then response.write "checked"%>>Y<input type="radio" name="c16" value="no" <% if check("15")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("length")%>" name="length" size="10"><input type="radio" name="c17" value="yes" <% if check("16")="yes" then response.write "checked"%>>Y<input type="radio" name="c17" value="no" <% if check("16")="no" then response.write "checked"%>>N
      </td>
    <td align="center"><b>ѧ&nbsp;&nbsp;&nbsp;&nbsp;��</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("study")%>" name="study" size="10"><input type="radio" name="c18" value="yes" <% if check("17")="yes" then response.write "checked"%>>Y<input type="radio" name="c18" value="no" <% if check("17")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��������</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("foreign")%>" name="foreign" size="10"><input type="radio" name="c19" value="yes" <% if check("18")="yes" then response.write "checked"%>>Y<input type="radio" name="c19" value="no" <% if check("18")="no" then response.write "checked"%>>N
      </td>
    <td align="center"><b>����ˮƽ</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("Elevel")%>" name="Elevel" size="10"><input type="radio" name="c20" value="yes" <% if check("19")="yes" then response.write "checked"%>>Y<input type="radio" name="c20" value="no" <% if check("19")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>���������</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("Clevel")%>" name="Clevel" size="10"><input type="radio" name="c21" value="yes" <% if check("20")="yes" then response.write "checked"%>>Y<input type="radio" name="c21" value="no" <% if check("20")="no" then response.write "checked"%>>N
      </td>
    <td align="center"><b>�������ڵ�</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("Hplace")%>" name="Hplace" size="10"><input type="radio" name="c22" value="yes" <% if check("21")="yes" then response.write "checked"%>>Y<input type="radio" name="c22" value="no" <% if check("21")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr>
    <td height="20" align="center"><b>QQ����</b></td>
    <td>
        <INPUT TYPE="text" value="<%=rs("QQ")%>" name="QQ" size="10"><input type="radio" name="c23" value="yes" <% if check("22")="yes" then response.write "checked"%>>Y<input type="radio" name="c23" value="no" <% if check("22")="no" then response.write "checked"%>>N
      </td>
    <td align="center"><b>EMAIL</b></td>
    <td>&nbsp;<%=Session("email")%></td>
  </tr>
  <tr>
    <td height="20" align="center"><b>���õ绰</b></td>
    <td>&nbsp;<%=Session("tel")%></td>
    <td align="center"><b>�ֻ�����</b></td>
    <td>&nbsp;<%=session("mobile")%></td>
  </tr>
  <tr>
    <td height="20" align="center"> <b>��������</b> </td>
    <td>
        <INPUT TYPE="text" value="<%=rs("call")%>" name="call1" size="10"><input type="radio" name="c24" value="yes" <% if check("23")="yes" then response.write "checked"%>>Y<input type="radio" name="c24" value="no" <% if check("23")="no" then response.write "checked"%>>N
      </td>
    <td align="center">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��&nbsp;ס&nbsp;ַ</b></td>
    <td colspan="3">
        <INPUT size=45 TYPE="text" value="<%=rs("place")%>" name="place"><input type="radio" name="c25" value="yes" <% if check("24")="yes" then response.write "checked"%>>Y<input type="radio" name="c25" value="no" <% if check("24")="no" then response.write "checked"%>>N</td></tr>
  <tr> 
    <td height="20" align="center"><b>����ר��<br>
      �Լ�����</b></td>
    <td colspan="3">
        <TEXTAREA NAME="love" ROWS="2" COLS="44"><%=rs("love")%></TEXTAREA><input type="radio" name="c26" value="yes" <% if check("25")="yes" then response.write "checked"%>>Y<input type="radio" name="c26" value="no" <% if check("25")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��������<br>
      �����ֽ�<br>
      ���ʹ���</b></td>
    <td colspan="3">
        <TEXTAREA NAME="award" ROWS="3" COLS="44"><%=rs("award")%></TEXTAREA><input type="radio" name="c27" value="yes" <% if check("26")="yes" then response.write "checked"%>>Y<input type="radio" name="c27" value="no" <% if check("26")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��������</b></td>
    <td colspan="3">
        <TEXTAREA NAME="experience" ROWS="2" COLS="44"><%=rs("experience")%></TEXTAREA><input type="radio" name="c28" value="yes" <% if check("27")="yes" then response.write "checked"%>>Y<input type="radio" name="c28" value="no" <% if check("27")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��ͥ���</b></td>
    <td colspan="3">
        <TEXTAREA NAME="family" ROWS="2" COLS="44"><%=rs("family")%></TEXTAREA><input type="radio" name="c29" value="yes" <% if check("28")="yes" then response.write "checked"%>>Y<input type="radio" name="c29" value="no" <% if check("28")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;��<br>
      ��ϵ��ʽ</b></td>
    <td colspan="3">
        <TEXTAREA NAME="contact" ROWS="2" COLS="44"><%=rs("contact")%></TEXTAREA><input type="radio" name="c30" value="yes" <% if check("29")="yes" then response.write "checked"%>>Y<input type="radio" name="c30" value="no" <% if check("29")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td height="20" align="center"><b>��&nbsp;&nbsp;&nbsp;&nbsp;ע</b></td>
    <td colspan="3">
        <TEXTAREA NAME="remark" ROWS="2" COLS="44"><%=rs("remark")%></TEXTAREA><input type="radio" name="c31" value="yes" <% if check("30")="yes" then response.write "checked"%>>Y<input type="radio" name="c31" value="no" <% if check("30")="no" then response.write "checked"%>>N
      </td>
  </tr>
  <tr> 
    <td height="20" align="center" colspan="4"><b><%=session("Rname")%>�ĸ��˻�������</b>&nbsp; 
      <input type="submit" value=" �� �� " name="submit">
      &nbsp; </td>
  </tr>

</table>
</form>
