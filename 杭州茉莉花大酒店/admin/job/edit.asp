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
<%Dim Act,ID
  ID=Cint(Request("ID"))  
  Act=Request.Form("act")
  Select Case Act
    Case "save" : Call SaveData()
    Case else : Call Main()
  End Select
  Call CloseDataBase()  
  Sub SaveData()
    Department = Request.Form("Department")
	Job = Request.Form("Job")
	Sex = Request.Form("Sex")
	Age = Request.Form("Age")
	Education = Request.Form("Education")
	Years = Request.Form("Years")
	Money = Request.Form("Money")
	Num = Request.Form("Num")
	EffectTime = Request.Form("EffectTime")
	Contact = Request.Form("Contact")
	Tel = Request.Form("Tel")
	Content = Request.Form("Content")
	AddDate= Request.Form("AddDate")
	leibie= Request.Form("leibie")
	address= Request.Form("address")
	workingway= Request.Form("workingway")
	yingjie= Request.Form("yingjie")
	Show=trim(Request.Form("Show"))
	Other= Request.Form("Other")
	If Job="" Then Call WriteErr("����д��Ƹְλ��",1)
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="Select * From Sbe_Job Where ID="&id
	Rs.Open Sql,Conn,1,3	  
	   Rs("Department")=Department
	   Rs("workingway")=workingway
	   Rs("yingjie")=yingjie   
	   Rs("Job")=Job
	   Rs("Sex")=Sex
	   Rs("Age")=Age
	   Rs("Education")=Education
	   Rs("Years")=Years
	   Rs("Money")=Money
	   Rs("Num")=Num
	   Rs("EffectTime")=EffectTime
	   Rs("Contact")=Contact
	   Rs("Tel")=Tel
	   Rs("Content")=Content
	   Rs("AddDate")=AddDate
	   Rs("leibie")=leibie	
	   Rs("address")=address
	  ' response.Write(show)
	   'response.End()
	   Rs("Show")=Show
	   Rs("Other")=Other
	   Rs.Update
	Rs.Close
	Set Rs=Nothing
	Response.Write("<script language=javascript>alert('��Ƹ��Ϣ�޸ĳɹ���');window.location.href='"&Request.Form("url")&"';</script>")
	Response.End()	
  End Sub  
  
  Sub Main()
  Set Rs=Server.CreateObject("adodb.recordset")
  Sql="Select Department,Job,Sex,Age,Education,Years,Money,Num,EffectTime,Contact,Tel,Content,leibie,AddDate,address,workingway,yingjie,Other,Show From Sbe_Job Where ID="&ID
  Rs.Open Sql,Conn,1,1  
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨����ϵͳ</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../include/meizzDate.js"></script>
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
    <td height="25"><font color="#6A859D"> ������Ƹ &gt;&gt; �޸���Ƹ��Ϣ</font></td>
  </tr>
  <tr> 
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>
  
<br>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <form name="form1" method="post" action="" onSubmit="return check()">
    <tr > 
      <td height="25" colspan="2"  class="sbe_table_title"><strong>�޸���Ƹ��Ϣ</strong></td>
    </tr>
    <tr> 
      <td height="25" align="center"><strong>��������</strong></td>
      <td height="21"> <input name="Department" type="text" class="input" id="Department" value="<%=rs(0)%>"></td>
    </tr>
    <tr > 
      <td width="13%" height="25" align="center"><strong>��λ����</strong></td>
      <td width="87%" height="21"><input name="Job" type="text" class="input" id="Job" value="<%=rs(1)%>" size="50"></td>
    </tr>
    <tr> 
      <td height="25" align="center"><strong>��Ƹ����</strong></td>
      <td height="21"><input name="Num" type="text" class="input" id="Num" value="<%=rs(7)%>"></td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>�Ա�Ҫ��</strong></td>
      <td height="21">
	    <select name="Sex" class="sbe_button" id="Sex">
          <option value="��Ů����" <%Call ReturnSel(rs(2),"��Ů����",1)%>>��Ů����</option>
          <option value="����" <%Call ReturnSel(rs(2),"����",1)%>>����</option>
          <option value="Ů��" <%Call ReturnSel(rs(2),"Ů��",1)%>>Ů��</option>
        </select></td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>����Ҫ��</strong></td>
      <td height="21"><input name="Age" type="text" class="input" id="Age" value="<%=rs(3)%>"> </td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>ѧ��Ҫ��</strong></td>
      <td height="21">
	   <select name="Education" class="sbe_button" id="Education">
          <option value="ѧ������"  <%Call ReturnSel(rs(4),"ѧ������",1)%>>ѧ������</option>
          <option value="��ʿ" <%Call ReturnSel(rs(4),"��ʿ",1)%>>��ʿ</option>
          <option value="˶ʿ" <%Call ReturnSel(rs(4),"˶ʿ",1)%>>˶ʿ</option>
          <option value="��ѧ����" <%Call ReturnSel(rs(4),"��ѧ����",1)%>>��ѧ����</option>
          <option value="��ר" <%Call ReturnSel(rs(4),"��ר",1)%>>��ר</option>
          <option value="��ר" <%Call ReturnSel(rs(4),"��ר",1)%>>��ר</option>
          <option value="ְ��/��У" <%Call ReturnSel(rs(4),"ְ��/��У",1)%>>ְ��/��У</option>
          <option value="����" <%Call ReturnSel(rs(4),"����",1)%>>����</option>
          <option value="����" <%Call ReturnSel(rs(4),"����",1)%>>����</option>
        </select></td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>�Ƿ�Ӧ��</strong></td>
      <td height="21"><input name="yingjie" type="radio" value="Ӧ��" <%if rs(16)="Ӧ��" then response.Write("checked") end if%>>
      Ӧ�� &nbsp;<input name="yingjie" type="radio" value="�ѹ���" <%if rs(16)="�ѹ���" then response.Write("checked") end if%>>�ѹ��� &nbsp;<input name="yingjie" type="radio" value="Ӧ�졢�ѹ�������" <%if rs(16)="Ӧ�졢�ѹ�������" then response.Write("checked") end if%>>
      Ӧ�졢�ѹ�������</td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>��������</strong></td>
      <td height="21"><input name="Years" type="text" class="input" id="Years" value="<%=rs(5)%>"></td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>нˮ��Χ</strong></td>
      <td height="21"><input name="Money" type="text" class="input" id="Money" value="<%=rs(6)%>"></td>
    </tr>
    <tr class="display">
      <td height="25" align="center"><strong>�� ϵ ��</strong></td>
      <td height="21"><input name="Contact" type="text" class="input" id="Contact" value="<%=rs(9)%>"></td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>��ϵ�绰</strong></td>
      <td height="21"><input name="Tel" type="text" class="input" id="Tel" value="<%=rs(10)%>"></td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>������ʽ</strong></td>
      <td height="21"><input name="workingway" type="text" class="input" id="workingway" value="<%=rs(15)%>"></td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>�����ص�</strong></td>
      <td height="21"><input name="address" type="text" class="input" id="address" value="<%=rs(14)%>"></td>
    </tr>
    <tr> 
      <td height="25" align="center"><strong>ְλҪ��</strong></td>
      <td height="21"><textarea name="Content" cols="80" rows="8" class="input" id="Content"><%=rs(11)%></textarea></td>
    </tr>
    <tr> 
      <td height="25" align="center"><strong>����</strong></td>
      <td height="21"><textarea name="Other" cols="80" rows="5" class="input" id="Content"><%=rs(17)%></textarea></td>
    </tr>
    <tr> 
      <td height="25" align="center"><strong>����ʱ��</strong></td>
      <td height="21"><input name="AddDate" type="text" class="input" id="AddDate" onFocus="setday(this)" 
	  value="<%if rs(13)<>"" then 
	              response.Write rs(13) 
				else
	             response.Write date() 
			end if%>">
      Ĭ��һ�㲻�޸�,ע��ʱ���ʽ!</td>
    </tr>
 <tr> 
      <td height="25" align="center"><strong>��ֹ����</strong></td>
      <td height="21"><input name="EffectTime" type="text" class="input" id="EffectTime" onFocus="setday(this)"  value="<%if rs(8)<>"" then 
	                                                                                   response.Write rs(8) 
																					 else
	                                                                      response.Write date()+3 end if%>"></td>
    </tr>
   <tr <%=banben_display%>> 
    <td align="center"><strong>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</strong></td>
    <td colspan="2">
 <input type="radio" name="leibie" value="1" <%if Rs(12)=1 then response.Write("checked") end if%>>
        �� &nbsp;&nbsp; <input name="leibie" type="radio" value="0" <%if Rs(12)=2 then response.Write("checked") end if%>>
        Ӣ</td>
  </tr>
 <tr> 
      <td height="25" align="center"><strong>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ʾ</strong></td>
      <td height="21"><input name="Show" type="radio" id="Show" value="1" <%if rs(18)=-1 then%> checked="checked"<%end if%>>�� <input name="Show" type="radio" id="Show" value="0"  <%if rs(18)<>-1 then%> checked="checked"<%end if%>>
      ��</td>
    </tr>
    <tr> 
      <td height="25" align="center">&nbsp;</td>
      <td height="21">
	    <input type="submit" name="Submit" value="�޸���Ϣ" class="sbe_button"> 
        <input name="act" type="hidden" id="act2" value="save">
        <input type="reset" name="Submit2" value="����" class="sbe_button">
        <input name="id" type="hidden" id="id" value="<%=id%>">
        <input name="url" type="hidden" id="url" value="<%=request.ServerVariables("HTTP_REFERER")%>"></td>
    </tr>
  </form>
</table>
<br>
</body>
</html>
<%Rs.Close
  Set Rs=Nothing
  End Sub%>
