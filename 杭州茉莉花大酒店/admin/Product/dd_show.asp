<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%
Call OpenData()%>
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
<%
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('���糬ʱ�������㻹û�е�½! ');this.location.href='../login.asp';<'/script>"
end if
 	temp_check_rights = Split(session("manconfig"),",")
	for temp_rights_count=0 to ubound(temp_check_rights)
	    if trim(temp_check_rights(temp_rights_count)) = "8" then
			rights_check_passkey = trim(temp_check_rights(temp_rights_count))
		end if
	next
	if rights_check_passkey <> "8" then
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('����Ȩ�޲��㣬�벻Ҫ�Ƿ�������������ģ�飬���������˺Ž���ϵͳ�Զ�ɾ��!');this.location.href='../login.asp';<'/Script>"
	Response.end
	end if%>
<%if request("act")="add" then
'     response.Write request("flag")
'     response.End
     Set Rs= Server.CreateObject("ADODB.RecordSet") 
     Rs.open "Select * from Sbe_order where ID=" & clng(request("id")),conn,1,3
     Rs("showtime")= date()
     Rs("status")= request("flag")
     rs.update  
     rs.close
     Set rs=nothing
     response.Redirect("dingdan.asp")
	 response.End	
	'Response.Write("<script>alert(""�ظ��ɹ�"");location.href=""dingdan.asp"";</'script>") 	
else
  if request("id") ="" then
     response.Write "<script LANGUAGE=javascript>alert('��������! ');history.go(-1);</script>"
     response.End
  else
   id=request("id")
   Sql = "Select * from Sbe_order where ID = "&id
   set rs=conn.execute(Sql)
   if rs.eof and rs.bof then
      response.Write "<script LANGUAGE=javascript>alert('��������! ');history.go(-1);</script>"
      response.End
    else
	  huiyuan=rs("huiyuan")  
	  username=rs("username")         '
	  usertel=rs("usertel")                 '           '
	  useremail=rs("useremail")             '
	  useraddress=rs("useraddress")
	  remarks=rs("remarks")
	  productname=rs("productname")
	  category=rs("category")
	  xinghao=rs("xinghao")           '
	  jiage=rs("jiage")           '
	  content=rs("content")          '
	  status1=rs("status")          '	
	  timing=rs("timing")
	  showtime=rs("showtime")
	  detail=rs("detail")
	  if rs("status")=1 then
	     a="disabled"
	   else
         a=""
       end if
     end if
   rs.close
   set rs=nothing
  end if
end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��������</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript>
// ��������
NS4 = document.layers && true;
IE4 = document.all && parseInt(navigator.appVersion) >= 4;
</script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="19%" height="25"><font color="#6A859D">�������� &gt;&gt;��������</font></td>   
      <td width="81%">&nbsp;       </td>   
  </tr>
  <tr> 
    <td height="1" colspan="2" background="../images/dot.gif"></td>
  </tr>
</table>
<table width="60%" border="0" align="center" cellpadding="0" cellspacing="0"  id="sbe_table">
                <form name=form method=post onSubmit="return checked();" action="dd_show.asp?act=add">
				 <tr align="center"> 
                    <td height="30" colspan="2" bgcolor="#EFEFEF" class="sbe_table_title">�������� >> ��������</td>
                  </tr>
				  	<tr> 
                    <td class=M align="right" bgcolor="#EFEFEF"><strong>��Ʒ��Ϣ</strong>��</td>
                    <td>               </td>
                  </tr>
				  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">��Ʒ���</td>
                    <td>&nbsp;<%=category%>                 </td>
                  </tr>
                 <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">��Ʒ���ƣ�</td>
                    <td>&nbsp;<%=productname%></td>
                  </tr>

                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">��Ʒ��ţ�</td>
                    <td>&nbsp;<%=xinghao%></td>
                  </tr>                 
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF" width="120">��Ʒ���</td>
                    <td>&nbsp;<%=jiage%>
					</td>
                  </tr>
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF" width="120">��Ʒ��װ��</td>
                    <td>&nbsp;<%=detail%>
					</td>
                  </tr>
<!--                  <tr style="display:none"> 
                    <td class=M bgcolor="#EFEFEF" align="right">QQ/MSN��</td>
                    <td><input name="URL2" type="text" id="URL2" value="<%=lyqq%>" size="40" readonly=""></td>
                  </tr>-->
<tr > 
                    <td class=M bgcolor="#EFEFEF" align="right">��Ʒ���ݣ�</td>
                    <td class=M>&nbsp;<%=content%></td>
                  </tr>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right"><strong>�ͻ�����</strong>��</td>
                    <td></td>
                  </tr>
				  <%if trim(huiyuan)<>"" then%>
				  <tr style="display:none;"> 
                    <td class=M bgcolor="#EFEFEF" align="right">��Ա�ʺţ�</td>
                    <td> &nbsp;<%=huiyuan%>
                    </td>
                  </tr>
				  <%end if%>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">�ͻ�������</td>
                    <td> &nbsp;<%=username%>
                    </td>
                  </tr>
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">�ͻ��绰��</td>
                    <td>&nbsp;<%=usertel%>
                    </td>
                  </tr>
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">�ͻ�Email��</td>
                    <td>&nbsp;<%=useremail%>
                    </td>
                  </tr>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">�ͻ���ַ��</td>
                    <td>&nbsp;<%=useraddress%>
                    </td>
                  </tr>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">�������ݣ�</td>
                    <td> 
                      &nbsp;<%=remarks%>
                    </td>
                  </tr>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">�Ƿ���</td>
                    <td>&nbsp;<input name="flag"  type="radio" class="input" value="1" <%if trim(status1)=1 then response.Write("checked") end if%>> 
                    ����&nbsp;&nbsp;
                    <input name="flag" type="radio" class="input" value="0" <%if trim(status1)=0 then response.write("checked") end if%> <%=a%>>
                    �ݲ�����
                    </td>
                  </tr>
                  <tr> 
                    <td align="right" bgcolor="#EFEFEF" class=M style="height:30px;">��</td>
                    <td> 
                      <input name="submit" type="submit" class="sbe_button" value=" ȷ �� ">
                      &nbsp;
                      <input name="submit2" type="hidden" class="sbe_button" value=" �� �� ">
                      &nbsp;
                      <input type="hidden" name="id" value=<%=id%>>
                    </td>
                  </tr>
                </form>
</table>
<table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="20">&nbsp;</td>
  </tr>
</table>

<Script Language="JavaScript">
	<!--
 function checked(){
//  if(document.form.hftheme.value == ""){
//   alert("�ظ����ⲻ��Ϊ��!");
//   document.form.hftheme.focus();
//   return false;
//  }
 // if(document.form.hfremark.value == ""){
//   alert("�ظ����ݲ���Ϊ��!");
//   document.form.hfremark.focus();
//   return false;
//  }
     //if(confirm('Do you add this order,If you add,You will unEdit?')){
  // return true;}
   //{
   //return false;
   //}
//return true;
}
   // -->
	</Script>
<%Call CloseDataBase()%>
</body>
</html>
