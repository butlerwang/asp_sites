<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%Call OpenData()
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('���糬ʱ�������㻹û�е�½! ');this.location.href='../login.asp';</script>"
end if
  IF instr(webConfig,", 7")=0 or instr(session("manconfig"),", 7")=0 Then'��վ��������
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('����Ȩ�޲��㣬�벻Ҫ�Ƿ�������������ģ�飬���������˺Ž���ϵͳ�Զ�ɾ��!');this.location.href='../login.asp';</Script>"
Response.end
end if
if request("id") ="" then
   response.Redirect("list.asp")
 else
   id=request("id")
 Set Rs= Server.CreateObject("ADODB.RecordSet") 	
 Sql = "Select * from Guest_book where ID = "&id
 rs.open Sql,conn,1,3
 if rs.eof and rs.bof then
    response.Redirect("list.asp")
  else
    if trim(Rs("status"))=0 then Rs("status")= 1
	Rs.update 
	lytheme=rs("lytheme")
	lyname=rs("lyname")
	lytel=rs("lytel")
	lyemail=rs("lyemail")
	lysex=rs("lysex")
	lyremark=rs("lyremark")
	lyaddress=rs("lyaddress")
	lytime=rs("lytime")
	status1=rs("status")
	leibie=rs("leibie")
	hfname=rs("hfname")
	hftheme=rs("hftheme")
	hfremark=rs("hfremark")
	hftime=rs("hftime")
	lycheck=rs("lycheck")
 end if
 rs.close
 set rs=nothing
end if
if request("act")="add" then
'if hy_message=true then
  Set Rs= Server.CreateObject("ADODB.RecordSet") 
  sql= "Select * from Guest_book where ID=" & clng(request("id"))
  rs.open sql,conn,1,3
'  response.Write(sql)
'  response.End()
  Rs("hfremark")= Request("hfremark")
  Rs("hftime")= date()
  Rs("status")= 1
  hfname="����Ա"
  Rs("hfname")= hfname
  if request("lycheck")<>"" then
  rs("lycheck")=1
  else
  rs("lycheck")=0
  end if
  rs.update  
  rs.close
  Set rs=nothing
' end if
   response.Redirect("list.asp")	
	'Response.Write("<script>alert(""�ظ��ɹ�"");location.href=""list.asp"";</script>") 
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
    <td width="19%" height="25"><font color="#6A859D">�������� &gt;&gt; �������</font></td>   
      <td width="81%">&nbsp;       </td>   
  </tr>
  <tr> 
    <td height="1" colspan="2" background="../images/dot.gif"></td>
  </tr>
</table>
<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0"  id="sbe_table">
                <form name=form method=post onSubmit="return checked();" action="ly_show.asp?act=add">
				 <tr align="center"> 
                    <td height="30" colspan="3" bgcolor="#EFEFEF" class="sbe_table_title">�������� >> �������</td>
                  </tr>
                  <tr> 
                    <td width="23%" align="right" bgcolor="#EFEFEF" class=M>�� ϵ �ˣ�</td>
                    <td width="48%"><%=lyname%></td>
                    <td width="29%">�ύ���ڣ�<%=lytime%></td>
                  </tr>
                  <tr class="display" > 
                    <td class=M bgcolor="#EFEFEF" align="right">��</td>
                    <td colspan="2" class=M><%=lytel%></td>
                  </tr>
                  <tr class="display"> 
                    <td align="right" bgcolor="#EFEFEF" class=M>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</td>
                    <td colspan="2" class=M><%=lysex%></td>
                  </tr>
                  <tr class="display"> 
                    <td align="right" bgcolor="#EFEFEF" class=M>E - mail��</td>
                    <td colspan="2" class=M><%=lyemail%></td>
                  </tr>
                  <tr class="display"> 
                    <td align="right" bgcolor="#EFEFEF" class=M>�绰��</td>
                    <td colspan="2" class=M><%=lytel%></td>
                  </tr>
                  <tr class="display"> 
                    <td align="right" bgcolor="#EFEFEF" class=M>��ϵ��ַ��</td>
                    <td colspan="2" class=M><%=lyaddress%></td>
                  </tr>
                  <tr > 
                    <td class=M bgcolor="#EFEFEF" align="right"> �������⣺</td>
                    <td colspan="2"><%=lytheme%></td>
                  </tr>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">�������ݣ�</td>
                    <td colspan="2"><%=HTMLcode(lyremark)%>                    </td>
                  </tr>
				   <!--<tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">�ظ����⣺</td>
                    <td> 
                      <input name="hftheme" type="text" id="hftheme" value="<%=hftheme%>" size="45" maxlength="100" <%=a%>>
                     <font color="#FF6600">*</font>                    </td>
                  </tr>-->
				  <%if hy_message=true then%>
                  <tr > 
                    <td class=M bgcolor="#EFEFEF" align="right">�ظ����ݣ�</td>
                    <td> 
                      <textarea name="hfremark" cols="46" rows="6" class="input" id="hfremark" <%'=a%>><%=hfremark%></textarea>
                    <font color="#FF6600">*</font>                    </td>
                  </tr>
                  <tr >
                    <td class=M bgcolor="#EFEFEF" align="right">��ˣ�</td>
                    <td><input name="lycheck" type="checkbox" id="lycheck" value="1" <%if lycheck=1 then%> checked="checked"<%end if%>></td>
                  </tr>
				  <%end if%>
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">��</td>
                    <td colspan="2">&nbsp;
                      <input type="button" name="submit2"class="sbe_button" value=" �� �� " onClick="javascript:history.go(-1);">
                      &nbsp;
                    <input type="hidden" name="id" value=<%=id%>>                    </td>
                  </tr>
                </form>
</table>
  <%if hy_message=true then%>
<Script Language="JavaScript">
	<!--
 function checked(){
//  if(document.form.hfremark.value == ""){
//   alert("�ظ����ݲ���Ϊ��!");
//   document.form.hfremark.focus();
//   return false;
//  }
     //if(confirm('Do you add this order,If you add,You will unEdit?')){
  // return true;}
   //{
   //return false;
   //}
return true;
}
   // -->
	</Script>
	<%end if%>
<%Call CloseDataBase()%>
</body>
</html>
