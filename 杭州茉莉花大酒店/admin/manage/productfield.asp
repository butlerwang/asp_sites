<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%If Session("flag")<>99 then
Session.Abandon()
response.Write "<script LANGUAGE=javascript>alert('����Ȩ�޲��㣬�벻Ҫ�Ƿ�������������ģ�飬���������˺Ž���ϵͳ�Զ�ɾ��! ');this.location.href='../login.asp';</script>"
end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨����ϵͳ</title>

<link href="../include/style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="25"><font color="#6A859D">��վ���� &gt;&gt; ��Ʒ�ֶ�����</font></td>
  </tr>
  <tr>
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>
<br>

<%
openData()
  Dim Act
  Act=Request("act")
  Select Case act
    Case "save" : Call SaveData()
	Case "del" : Call Del()
	Case "show" : Call Show()
	Case "modify" : Call Modify()
	Case "savemodify" : Call SaveModify()
	Case else : Call Main()
  End Select
  Call CloseDataBase()
  
  Sub Show()
    id=Cint(Request.QueryString("id"))
	Set Rs=Server.CreateObject("adodb.recordset")
	sql="select * from SBE_Product_Field where id="&id
	rs.open sql,conn,1,3
	  if rs("show")=true then
	     rs("show")=0
	  else 
	     rs("show")=1
	  end if
	  rs.update
	rs.close
	set rs=nothing	
	response.Redirect(request.ServerVariables("HTTP_REFERER"))
  End Sub
  
  Sub Del()
    id=Cint(Request.QueryString("id"))
	Set Rs=Server.CreateObject("adodb.recordset")
	sql="select * from SBE_Product_Field where id="&id
	rs.open sql,conn,1,3
	   if not Rs.Eof then
	     sql="ALTER TABLE Sbe_Product DROP COLUMN "&rs("FieldTitle")
		 Conn.execute sql
	     rs.delete
	   end if
	rs.close
	set rs=nothing
	response.Redirect(request.ServerVariables("HTTP_REFERER"))  
  End Sub
  
    Sub SaveModify()
	id=Cint(Request.Form("id"))
    Dim FieldName,FieldShow,FieldShowLength,showa  
	FieldName=Trim(Request.Form("FieldName"))
	FieldShow=Request.Form("FieldShow")
	FieldShowLength=Cint(Request.Form("FieldShowLength"))
	showa=Request.Form("show")
	FieldLength=Request.Form("FieldLength")
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="Select * From SBE_Product_Field Where id="&id
	Rs.Open Sql,Conn,1,3	
		 Rs("FieldName")=FieldName
		 Rs("FieldShow")=FieldShow
		 Rs("FieldShowLength")=FieldShowLength	 
	  	 Rs("Show")=showa
		 Rs("FieldLength")=FieldLength 
		 Rs.Update	
	Rs.Close
	Set Rs=Nothing
	Response.Write("<script language=javascript>alert('�޸ĳɹ���');window.location.href='"&request.Form("url")&"';</script>")
	response.End()
  End Sub
  
  
  
  Sub SaveData()
    Dim FieldTitle,FieldLength,FieldName,FieldShow,FieldShowLength,Build
    FieldTitle=Trim(Request.Form("FieldTitle"))
	FieldLength=Request.Form("FieldLength")
	FieldName=Trim(Request.Form("FieldName"))
	FieldShow=Request.Form("FieldShow")
	FieldShowLength=Cint(Request.Form("FieldShowLength"))
	Build=Request.Form("Build")
	
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="Select * From SBE_Product_Field Where FieldTitle='"&FieldTitle&"'"
	Rs.Open Sql,Conn,1,3
	  If Not Rs.Eof Then
	     Call WriteErr("���ֶ��Ѿ����ڣ�",1)
	  Else
	     Rs.AddNew
		 Rs("FieldTitle")=FieldTitle
		 Rs("FieldLength")=FieldLength
		 Rs("FieldName")=FieldName
		 Rs("FieldShow")=FieldShow
		 Rs("FieldShowLength")=FieldShowLength
		 Rs("Show")=1
		 Rs("Lock")=0
		 Rs.Update
	  End If
	Rs.Close
	Set Rs=Nothing
	
	If Build = 1 Then
	  sql="ALTER TABLE Sbe_Product ADD "&FieldTitle&" NVARCHAR("&FieldLength&")"
	  Conn.execute sql	
	End If
	Response.Redirect(request.ServerVariables("HTTP_REFERER"))
  
  End Sub

  Sub Main()
%>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <tr class="sbe_table_title"> 
    <td width="17%" height="26" class="sbe_table_title">�ֶ���</td>
    <td height="26" class="sbe_table_title">�ֶγ���</td>
    <td class="sbe_table_title">�ֶ�����</td>
    <td class="sbe_table_title">�ı���</td>
    <td class="sbe_table_title">�ı��򳤶�</td>
    <td height="26" class="sbe_table_title">�Ƿ�ʹ��</td>
    <td height="26" class="sbe_table_title">�Ƿ�ϵͳ�ֶ�</td>
    <td height="26" class="sbe_table_title">�༭</td>
    <td height="26" class="sbe_table_title">ɾ��</td>
  </tr>
  <%
   
   Set rs=server.CreateObject("Adodb.recordset")
   Sql="select * from Sbe_Product_Field order by Sequence"
   Rs.open Sql,conn,1,1
     do while not rs.eof

	%>
  <tr> 
    <td height="25" align="center"><%=rs("FieldTitle")%></td>
    <td width="11%" height="21" align="center" bgcolor="#E9EFF3"><%=rs("FieldLength")%></td>
    <td width="14%" align="center"><%=rs("fieldname")%></td>
    <td width="9%" align="center" bgcolor="#E9EFF3">
      <%
	  if rs("FieldShow")=1 then
	     Response.Write("����")
	   elseif rs("FieldShow")=2 then
	      Response.Write("����")
	   elseif rs("FieldShow")=3 then
	      Response.Write("����")
	   elseif rs("FieldShow")=4 then
	      Response.Write("�༭")		  
	   elseif rs("FieldShow")=5 then
	      Response.Write("��ѡ")		  
	   elseif rs("FieldShow")=6 then
	      Response.Write("��ѡ")		  
	   elseif rs("FieldShow")=7 then
	      Response.Write("����") 
	  end if
	  %>
    </td>
    <td width="12%" align="center"><%=rs("fieldShowLength")%></td>
    <td width="9%" align="center" bgcolor="#E9EFF3"><a href="productfield.asp?act=show&id=<%=rs("id")%>"><%=JudgeMent(rs("show"))%></a></td>
    <td width="12%" align="center"><%=JudgeMent(rs("Lock"))%></td>
    <td width="8%" align="center" bgcolor="#E9EFF3"><a href="productfield.asp?act=modify&id=<%=rs("id")%>"><img src="../images/edit.gif" border="0" ></a> 
    </td>
    <td width="8%" align="center"> 
	<%if rs("lock")=false then%>
	<a href="productfield.asp?act=del&id=<%=rs("id")%>" onClick="javascript:return confirm('��ɾ������ֶ���������Ϣ����ʧ!\nȷ��ɾ����')"> 
      <img src="../images/delete.gif" border="0"></a>
	 <%End If%> 
	  </td>
  </tr>
  <%
    Rs.movenext
	loop
	RS.close
	set rs=nothing
	%>
</table>
<br>
<script language="JavaScript">
function check(){
 if(form1.FieldTitle.value==""){
    alert("����д�ֶ�����");
	form1.FieldTitle.focus();
	return false;
 }
 re=/^[0-9]+$/;
 if(!re.test(form1.FieldLength.value)){
   alert("����д�ֶγ��ȣ�");
   form1.FieldLength.focus();
   return false;
 }
 if(!re.test(form1.FieldShowLength.value)){
    alert("����д�ı��򳤶ȣ�");
	form1.FieldShowLength.focus();
	return false;
 }
return true;
}
</script>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <form name="form1" method="post" action="productfield.asp?act=save" onSubmit="return check()">
    <tr> 
      <td align="center">�ֶ���</td>
      <td><input name="FieldTitle" type="text" class="input" id="FieldTitle"></td>
      <td align="center">�ֶγ���</td>
      <td height="25"><input name="FieldLength" type="text" class="input" id="FieldLength"></td>
    </tr>
    <tr> 
      <td align="center">�ֶ�����</td>
      <td><input name="FieldName" type="text" class="input" id="FieldName"></td>
      <td align="center">�ı���</td>
      <td height="25"><input name="FieldShow" type="radio" class="input" value="1" checked>
        ����&nbsp;&nbsp; 
        <input name="FieldShow" type="radio" class="input" value="2">
        ����&nbsp;&nbsp; 
        <input name="FieldShow" type="radio" class="input" value="3">
        ����&nbsp;&nbsp; 
        <input name="FieldShow" type="radio" class="input" value="4">
        �༭<br>
        <input name="FieldShow" type="radio" class="input" value="5">
        ��ѡ&nbsp;&nbsp; 
        <input name="FieldShow" type="radio" class="input" value="6">
		��ѡ&nbsp;&nbsp; 
        <input name="FieldShow" type="radio" class="input" value="7">
		����
      </td>
    </tr>
    <tr> 
      <td align="center">�ı��򳤶�</td>
      <td>
       <input name="FieldShowLength" type="text" class="input" id="FieldShowLength"></td>
      <td align="center">ֱ�ӽ��ֶ�</td>
      <td height="25"><input name="Build" type="checkbox" id="Build" value="1" checked></td>
    </tr>
    <tr> 
      <td width="17%" align="center">&nbsp;</td>
      <td width="37%"><input name="Submit" type="submit" value=" ���� " class="sbe_button">
        <input name="act" type="hidden" id="act" value="save"></td>
      <td width="14%" align="center">&nbsp;</td>
      <td width="32%" height="25">&nbsp;&nbsp; </td>
    </tr>
  </form>
</table>
<%End Sub


  Sub Modify()
  id=Cint(request.QueryString("id"))
  Set Rs=Server.CreateObject("adodb.recordset")
  sql="select * from Sbe_Product_Field where id="&id
  rs.open sql,conn,1,1  
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <form name="form1" method="post" action="productfield.asp" onSubmit="return check()">
    <tr> 
      <td align="center">�ֶ���</td>
      <td><%=rs("FieldTitle")%></td>
      <td align="center">�ֶγ���</td>
      <td height="25"><input name="FieldLength" type="text" class="input" id="FieldLength" value="<%=rs("FieldLength")%>"></td>
    </tr>
    <tr> 
      <td align="center">�ֶ�����</td>	  
      <td>	  
      <input name="FieldName" type="text" class="input" id="FieldName" value="<%=rs("FieldName")%>">	  	  
	  </td>
      <td align="center">�ı���</td>
      <td height="25">
	  
	  <input name="FieldShow" type="radio" class="input" value="1" <%Call ReturnSel(rs("FieldShow"),1,2)%>>
        ����&nbsp;&nbsp; 
        <input type="radio" name="FieldShow" class="input" value="2" <%Call ReturnSel(rs("FieldShow"),2,2)%>>
        ����&nbsp;&nbsp; 
        <input name="FieldShow" type="radio" class="input" value="3" <%Call ReturnSel(rs("FieldShow"),3,2)%>>
        ����&nbsp;&nbsp; 
        <input name="FieldShow" type="radio" class="input" value="4" <%Call ReturnSel(rs("FieldShow"),4,2)%>>
        �༭<br>
        <input name="FieldShow" type="radio" class="input" value="5" <%Call ReturnSel(rs("FieldShow"),5,2)%>>
        ��ѡ&nbsp;&nbsp; 
        <input name="FieldShow" type="radio" class="input" value="6" <%Call ReturnSel(rs("FieldShow"),6,2)%>>
		��ѡ&nbsp;&nbsp; 
        <input name="FieldShow" type="radio" class="input" value="7" <%Call ReturnSel(rs("FieldShow"),7,2)%>>
		����
	  </td>
    </tr>
    <tr> 
      <td align="center">�ı��򳤶�</td>
      <td>	 
       <input name="FieldShowLength" type="text" class="input" id="FieldShowLength" value="<%=rs("FieldShowLength")%>">
	   
      </td>
      <td align="center">�Ƿ�ʹ��</td>
      <td height="25"><input name="show" type="radio" value="1"  <%Call ReturnSel(rs("show"),true,2)%>>
        ʹ��&nbsp;&nbsp; 
        <input type="radio" name="show" value="0"  <%Call ReturnSel(rs("show"),false,2)%>>
        ���� </td>
    </tr>
    <tr> 
      <td width="17%" align="center">&nbsp;</td>
      <td width="37%"><input name="Submit" type="submit" value=" �޸� " class="sbe_button">
	    <input name="act" type="hidden" id="act" value="savemodify">
        <input name="id" type="hidden" id="id" value="<%=id%>">
		<input name="url" type="hidden" id="url" value="<%=request.ServerVariables("HTTP_REFERER")%>">
		</td>
      <td width="14%" align="center">&nbsp;</td>
      <td width="32%" height="25">&nbsp;&nbsp; </td>
    </tr>
  </form>
</table>
<%rs.close
  set rs=nothing
  End Sub%>

</body>
</html>
