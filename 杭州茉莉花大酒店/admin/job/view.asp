<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%openData()
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('���糬ʱ�������㻹û�е�½! ');this.location.href='../login.asp';</script>"
end if
  IF instr(webConfig,", 6")=0 or instr(session("manconfig"),", 6")=0 Then'��վ��������
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('����Ȩ�޲��㣬�벻Ҫ�Ƿ�������������ģ�飬���������˺Ž���ϵͳ�Զ�ɾ��!');this.location.href='../login.asp';</Script>"
Response.end
end if%>
<%Dim Act
  Act=Request.Form("act")
  OpenData()
  Call Main()  
  Call CloseDataBase()
  Sub Main()
  ID=Cint(Request.QueryString("iD"))
  Set Rs=Server.CreateObject("adodb.recordset")
  Sql="Select Job,RealName,Sex,Birthday,Address,School,Education,Profession,working,Tel,Email,Content,fuqin,muqin,jk,aihao,hj,bysj From Sbe_Resume Where id="&ID
  Rs.Open SQL,Conn,1,1
 ' set rs1=server.CreateObject("adodb.recordset")
'rs1.open "select * from [Sbe_Resume] where id="&ID ,conn,1,3
'rs1("Ability")=2
'rs1.update
'rs1.close
'set rs1=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨����ϵͳ</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function check(){
  if(form1.Job.value==""){
     alert("��ѡ��д��λ���ƣ�");
	 form1.Job.focus();
	 return false;
  }
  return true;  
}
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="25"><font color="#6A859D"> ������Ƹ &gt;&gt; ��ְ��Ϣ����</font></td>
  </tr>
  <tr> 
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>
  
<br>
<table width="71%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <form name="form1" method="post" action="" onSubmit="return check()">
    <tr > 
      <td height="25" colspan="2"  class="sbe_table_title">��ְ��Ϣ����</td>
    </tr>
    <tr > 
      <td width="106" height="25" align="right" class="szise3">������</td>
      <td>&nbsp;<%=rs(1)%></td>
    </tr>
            <tr>
              <td height="25" align="right" class="szise3">ӦƸ��λ��</td>
              <td width="588" style="padding-left:4px">&nbsp;<% sql1="select Job from Sbe_Job  where ID="&rs(0)
	  set rs1=conn.execute(sql1)
	  if not rs.eof then
	    response.Write rs1(0)
	   end if
	  rs1.close
	  set rs1=nothing%></td>
            </tr>
            <tr>
              <td height="25" align="right" class="szise3">�������£�</td>
              <td>&nbsp;<%=rs(3)%></td>
            </tr>
            <tr>
              <td height="25" align="right" class="szise3">��ѧרҵ��</td>
              <td>&nbsp;<%=rs(7)%></td>
            </tr>
            <tr>
              <td height="25" align="right" class="szise3">��ҵʱ�䣺</td>
              <td>&nbsp;<%=rs(17)%></td>
    </tr>
            <tr>
              <td height="25" align="right" class="szise3">ѧλ��</td>
              <td>&nbsp;<%=rs(6)%></td>
            </tr>
            <tr>
              <td height="25" align="right" class="szise3">�������</td>
              <td>&nbsp;<%=rs(16)%></td>
            </tr>
            
            <tr>
              <td height="25" align="right" class="szise3">��Ȥ���ã�</td>
              <td>&nbsp;<%=rs(15)%></td>
            </tr>
            
            <tr>
              <td height="25" align="right" class="szise3">����״����</td>
              <td>&nbsp;<%=rs(14)%></td>
    </tr>
            <tr>
              <td height="25" align="right" class="szise3">����ְҵ��</td>
              <td>&nbsp;<%=rs(12)%></td>
            </tr>
            <tr>
              <td height="25" align="right" class="szise3">ĸ��ְҵ��</td>
              <td>&nbsp;<%=rs(13)%></td>
            </tr>
            
            
            <tr>
              <td height="76" align="right" class="szise3">����������</td>
              <td valign="middle">&nbsp;<%=HTMLcode(rs(11))%></td>
    </tr>
  </form>
</table>
<div align="center"><br>
  <input type="button" name="Submit" onClick="history.go(-1);return false;" value="����" />
  <br>
</div>
</body>
</html>
<%Rs.Close
Set rS=Nothing
  End Sub%>
