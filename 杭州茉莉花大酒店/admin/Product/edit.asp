<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<script language="JavaScript" src="../include/meizzDate.js"></script>
<%
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('���糬ʱ�������㻹û�е�½! ');this.location.href='../login.asp';</script>"
end if
 	temp_check_rights = Split(session("manconfig"),",")
	for temp_rights_count=0 to ubound(temp_check_rights)
	    if trim(temp_check_rights(temp_rights_count)) = "2" then
			rights_check_passkey = trim(temp_check_rights(temp_rights_count))
		end if
	next
	if rights_check_passkey <> "2" then
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('����Ȩ�޲��㣬�벻Ҫ�Ƿ�������������ģ�飬���������˺Ž���ϵͳ�Զ�ɾ��!');this.location.href='../login.asp';</Script>"
	Response.end
	end if%>
<%Dim Act,ID
  Act=Request.Form("act")
  ID=Cint(Request("id"))
  openData()
  If Act="save" Then
     Call SaveData()
  Else
     Call Main()
  End If
  Call CloseDataBase() 
  Sub SaveData()
     Pname=Trim(Request.Form("Pname"))
	 Tid=Cint(Request.Form("Tid"))
	 Hot=Cint(Request.Form("Hot"))
	 Ptype=Trim(Request.Form("Ptype"))
	 Bpic=Trim(Request.Form("Bpic"))
	 spic=Trim(Request.Form("Spic"))
	 Price=Trim(Request.Form("Price"))
	 if price="" then price=0
	 leibie=Trim(Request.Form("leibie"))
	  num=Trim(Request.Form("num"))
	 Tuijian=Request.Form("Tuijian")
	 Show=Request.Form("Show")
	 datet=Request.Form("datet")
	 detail =Request.Form("detail")
	 password=trim(Request.Form("password"))
	 shifou =trim(Request.Form("shifou"))
  sqlsize ="select * from Sbe_Product_Class where ID ="&Tid
  set rssize=conn.execute(sqlsize)
  if not (rssize.eof and rssize.bof) then
    if  rssize("Depth") = 0 then
	   bigclass=rssize("ID")
	   else
       bigclass = rssize("ParID")
	end if
  end if 
  rssize.close
 set rssize=nothing
     Content = ""
     For i = 1 To Request.Form("content").Count
       Content = Content & Request.Form("content")(i)
     Next
	 Uploadfile=request.Form("Uploadfile") 
	 Set Rs=Server.CreateObject("adodb.recordset")
	 sql="select * From Sbe_Product Where ID="&ID
	 Rs.Open Sql,Conn,1,3	  
		Rs("Pname")=Pname
		Rs("Tid")=Tid
		Rs("Ptype")=Ptype
		Rs("Bpic")=Bpic
		Rs("spic")=spic
		Rs("Price")=Price
		Rs("Tuijian")=Tuijian
		Rs("Content")=Content
		Rs("Uploadfile")=Uploadfile
		Rs("Hot")=Hot
		Rs("leibie")=leibie
		Rs("datet")=datet
		Rs("detail")=detail
		Rs("Show")=Show
		Rs("bigclass")=bigclass
		Rs("password")=password	
		Rs("shifou")=shifou	
		rs("num")=num
		'Set Rs1=Server.CreateObject("adodb.recordset")
'		Sql="Select FieldTitle From Sbe_Product_Field Where Lock=0"
'		Rs1.Open Sql,Conn,1,1
'		   do While Not Rs1.Eof			
'		      Rs(Cstr(rs1(0)))=request.Form(Cstr(rs1(0)))
'		   Rs1.MoveNext
'		   Loop
'		Rs1.Close
'		Set Rs1=Nothing
		rs.update
	rs.Close
	Set Rs=Nothing 
    response.Write("<script language=javascript>alert('�ͷ���Ϣ�޸ĳɹ���');window.location.href='"&request.Form("returnurl")&"';</script>")
	response.End()
  End Sub
  
  Sub Main()
  Tid=request("tid")
  if tid="" then tid=0
  tid=cint(tid) 
  Set Rs2=Server.CreateObject("adodb.recordset")
  Sql="Select * From Sbe_Product Where ID="&ID
  Rs2.Open Sql,Conn,1,1  
  Set Rs=Server.CreateObject("adodb.recordset")
  Sql="select picAuto from Sbe_WebConfig"
  Rs.Open sql,Conn,1,1
     PicAuto=rs(0)
  Rs.Close
  Set Rs=Nothing  
  Set Rs=Server.CreateObject("adodb.recordset")
  sql="Select FieldName,FieldShow,FieldShowLength,Show,FieldLength from Sbe_Product_Field Where Lock=1 order by Sequence "
  Rs.Open Sql,Conn,1,1
    if rs.recordcount<>10 then
	   response.Write("ϵͳ�ֶζ�ʧ������SBE_PRODUCT_FIELD��")
	   Response.End()
	end if  
%>
<html>                                                                               
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨����ϵͳ</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function check(){
  if(form1.Tid.value==""){
     alert("��ѡ����࣡");
	 form1.Tid.focus();
	 return false;
  }
  if(form1.Pname.value==""){
     alert("����д�ͷ����ƣ�");
	 form1.Pname.focus();
	 return false;
  }
  if(document.form1.shifou[1].checked==true){
  if (form1.password.value==""){
     alert("����д�鿴�û�����");
	 form1.password.focus();
	 return false;
  } 
 }
 document.form1.addbtn.disabled=true;
 document.form1.addbtn.value="���Ժ�..."
  return true;
}  
 function show_user_rights_menu(menu_id)
{
if (menu_id==0)
{
eval("show_user_rights.style.display=\"none\";");
}
else
{
eval("show_user_rights.style.display=\"\";");
}
}
  function PasswordShow(flag){
   if (flag==1){
       Showpassword.style.display="";
	   }
   if (flag==0){
       Showpassword.style.display="none";
	   }
  }
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="25"><font color="#6A859D">�ͷ����� &gt;&gt; �޸Ŀͷ���Ϣ</font></td>
  </tr>
  <tr>
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>
<br>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <form name="form1" method="post" action="edit.asp" onSubmit="return check()">
    <tr <%Call OpenClose(rs(3))%>> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td colspan="2"><select name="Tid" class="sbe_button">
          <option>��ѡ��...&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
          <%
		    Call ShowClass("sbe_product",rs2("tid"))%>
        </select> </td>
    </tr>
    <%
	rs.movenext '�ƶ���Ptype�ֶ�
	%>
    <tr <%Call OpenClose(rs(3))%>> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td colspan="2"> <%Call ShowField("Pname",rs(1),rs(2),rs2("pname"),rs(4))%> </td>
    </tr>
    <%
	rs.movenext '�ƶ���Ptype�ֶ�
	%>
    <tr <%Call OpenClose(rs(3))%>  style="display:none"> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td colspan="2"><%Call ShowField("Ptype",rs(1),rs(2),rs2("Ptype"),rs(4))%></td>
    </tr>
    <%
	rs.movenext '�ƶ���Bpic�ֶ�
	%>
    <tr <%Call OpenClose(rs(3))%>  style="display:none"> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td width="23%"> <%Call ShowField("Bpic",rs(1),rs(2),rs2("Bpic"),rs(4))%></td>
      <td width="61%"><iframe src="../upload/upload.asp?Form_Name=form1&UploadFile=Bpic" width="64%" height="25" frameborder="0" scrolling="no"></iframe>
      (ͼƬ��ѳߴ�:225*300)</td>
    </tr>
    <%
	rs.movenext '�ƶ���Spic�ֶ�
	%>
    <tr <%Call OpenClose(rs(3))%>> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td><%Call ShowField("Spic",rs(1),rs(2),rs2("Spic"),rs(4))%></td>
      <td> <iframe src="../upload/upload.asp?Form_Name=form1&UploadFile=Spic" width="64%" height="25" frameborder="0" scrolling="no"></iframe> 
        <%if PicAuto=false then%><%else%>
        <%end if%>
        (ͼƬ��ѳߴ�:393*278)</td>
    </tr>
    <%
	rs.movenext '�ƶ���Price�ֶ�
	%>
<!--    <tr <%'Call OpenClose(rs(3))%>  style="display:none" > 
      <td height="25" align="center"><%'=rs(0)%></td>
      <td colspan="2"><%'Call ShowField("Price",rs(1),rs(2),rs2("Price"),rs(4))%></td>
    </tr>-->
    <%
	rs.movenext '�ƶ���Tuijian�ֶ�
	%>
    <tr <%Call OpenClose(rs(3))%>  style="display:none"> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td colspan="2"> <input type="radio" name="Tuijian" value="1"  <%Call ReturnSel(rs2("Tuijian"),true,2)%>>
        �� &nbsp;&nbsp; <input name="Tuijian" type="radio" value="0"  <%Call ReturnSel(rs2("Tuijian"),false,2)%>>��</td>
    </tr>
<!-- onclick=show_user_rights_menu(1)
  <tr id="show_user_rights" <%'if Tuijian=false then response.write("style='display:none;'") end if%>>
    <td align="center">�ϴ�ͼƬ</td> 
    <td width="23%"><input name="pic" type="text" class="input" value="<%=pic%>" size="25"></td>
    <td width="61%"><iframe src="../upload/upload.asp?Form_Name=form1&UploadFile=pic" width="304" height="25" frameborder="0" scrolling="no"></iframe> ͼƬ�ߴ磺112*148</td>
  </tr>-->
    <%
	rs.movenext '�ƶ�����Ʒ�����ֶ�
	%>
    <tr <%if rs(3)=true then%><%=banben_display%><%else%><%Call OpenClose(rs(3))%><%end if%>> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td colspan="2"> <input type="radio" name="leibie" value="1" <%Call ReturnSel(rs2("leibie"),1,2)%>>
        �� &nbsp;&nbsp; <input name="leibie" type="radio" value="2"  <%Call ReturnSel(rs2("leibie"),2,2)%>>
        Ӣ</td>
    </tr>
<tr  style="display:none"> 
      <td height="25" align="center">�鿴Ȩ��</td>
      <td colspan="2"> <input type="radio" name="shifou" value="0" checked="checked" onClick="PasswordShow(0)" <%Call ReturnSel(rs2("shifou"),false,2)%>>
        �����û��� &nbsp;&nbsp; <input name="shifou" type="radio" value="1" onClick="PasswordShow(1)" <%Call ReturnSel(rs2("shifou"),true,2)%>>
        ��Ҫ�û���<span id="Showpassword" <%if rs2("shifou")=false then response.Write("style='display:none'") end if%>>&nbsp;&nbsp;
        <input name="password" type="text" class="input" style="ime-mode:Disabled;" value="<%=rs2("password")%>" size="25" maxlength="20">
        &nbsp;(<font color="#FF0000">�������û���</font>)</span></td>
    </tr>
   <%Set Rs1=Server.CreateObject("adodb.recordset")
     Sql="Select FieldName,FieldShow,FieldShowLength,FieldTitle,Show,FieldLength from Sbe_Product_Field Where Lock=0 order by Sequence"
	 rs1.open sql,conn,1,1 
	   do while not rs1.eof  
   %>
    <tr <%if Rs1(4)=0 then response.Write("class=""display""") end if%>>
      <td height="25" align="center"><%=rs1(0)%></td>
      <td colspan="2"><%Call ShowField(rs1(3),rs1(1),rs1(2),rs2(CStr(rs1(3))),rs1(5))%></td>
    </tr>
	<% rs1.movenext
	   Loop
	  Rs1.Close
	  Set Rs1=Nothing
	%>
    <%
	rs.movenext '�ƶ���Content�ֶ�
	%>	
<!--    <tr <%Call OpenClose(rs(3))%> style="display:none"> 
      <td height="12" align="center"><%=rs(0)%> <input name="content" type="hidden" id="content" value="<%=server.HTMLEncode(rs2("content"))%>"> <input name="uploadfile" type="hidden" id="uploadfile" value="<%=rs2("uploadfile")%>"></td>
      <td colspan="2"><iframe ID="eWebEditor1" src="../eWebEditor/ewebeditor.asp?id=content&style=sbe&savefilename=uploadfile" frameborder="0" scrolling="no" width="100%" HEIGHT="350"></iframe></td>
    </tr>-->
    <tr <%Call OpenClose(rs(3))%>>
      <td height="13" align="center">����</td>
      <td colspan="2"><input name="num" type="text" id="num"  value="<%=rs2("num")%>"></td>
    </tr>
    <tr <%Call OpenClose(rs(3))%>>
      <td height="25" align="center">����</td>
      <td colspan="2"><textarea name="content" cols="50" rows="8" id="content"><%=rs2("content")%></textarea></td>
    </tr>
    <tr <%Call OpenClose(rs(3))%>>
      <td height="25" align="center">�۸�</td>
      <td colspan="2"><textarea name="price" cols="50" rows="8" id="price"><%=rs2("price")%></textarea></td>
    </tr>
    <%
	rs.movenext '�ƶ���Show�ֶ�
	%>
    <tr <%Call OpenClose(rs(3))%>> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td colspan="2"> <%Call ShowField("Show",rs(1),rs(2),rs2("Show"),rs(4))%></td>
    </tr>
    <%rs.close
	  set rs=Nothing
	  '�ر�
	  %>
    <tr> 
      <td width="16%" height="40" align="center">&nbsp;</td>
      <td colspan="2"> <input name="addbtn" type="submit" value=" �޸� " class="sbe_button"> 
        &nbsp; <input type="reset" name="Submit2" value=" ��ԭ " class="sbe_button">
        <input name="act" type="hidden" id="act" value="save">
        <input name="returnurl" type="hidden" id="returnurl" value="<%=request.ServerVariables("HTTP_REFERER")%>">
        <input name="id" type="hidden" id="id" value="<%=id%>"></td>
    </tr>
  </form>
</table>
<br>
</body>
</html>
<%
 rs2.close
 Set Rs2 = Nothing
 End Sub
 %>
<%Sub ShowField(FieldName,FieldType,FieldLength,FieldValue,chandu)
   If FieldType=5 Then%>
<input type="radio" name="<%=FieldName%>" value="1" <%Call ReturnSel(FieldValue,true,2)%>>
        �� &nbsp;&nbsp; <input name="<%=FieldName%>" type="radio" value="0"  <%Call ReturnSel(FieldValue,false,2)%>>
        ��	  
<%
   Elseif FieldType=2 Then 
      Response.Write("<textarea name="""&FieldName&""" cols="""&FieldLength&""" rows=""3"" class=""input"">"&FieldValue&"</textarea>")
   elseIf FieldType=3 Then
      Response.Write("<input type=""password"" name="""&FieldName&""" size="""&FieldLength&""" value="""&FieldValue&""" class=""input"" maxlength="""&chandu&""">")
   elseIf FieldType=4 Then
      Response.Write("<input type=""hidden"" name="""&FieldName&""" value="""&FieldValue&""" class=""input""><iframe ID=""eWebEditor1"" src=""../eWebEditor/ewebeditor.asp?id=content&style=sbe&savefilename=uploadfile"" frameborder=""0"" scrolling=""no"" width=""100%"" HEIGHT=""350""></iframe>")
   else
 '  If FieldType=1 Then
   if FieldName="datet" then
      Response.Write("<input type=""text"" name="""&FieldName&""" onFocus=""setday(this)"" size="""&FieldLength&""" class=""input"" value="""&FieldValue&""" maxlength="""&chandu&""" readonly>")
	  else
      Response.Write("<input type=""text"" name="""&FieldName&""" size="""&FieldLength&""" class=""input"" value="""&FieldValue&""" maxlength="""&chandu&""">")
	 end if
   'elseIf FieldType=6 Then
'      Response.Write("<input type=""password"" name="""&FieldName&""" size="""&FieldLength&""" value="""&FieldValue&""" class=""input"">")
'   elseIf FieldType=7 Then
'      Response.Write("<input type=""password"" name="""&FieldName&""" size="""&FieldLength&""" value="""&FieldValue&""" class=""input"">")
   End If 
 End Sub
 
 Sub OpenClose(Flag)
   If Flag=false Then Response.Write("style=""display:none""")
 End Sub

%>