<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('���糬ʱ�������㻹û�е�½! ');this.location.href='../login.asp';</script>"
end if
 	temp_check_rights = Split(session("manconfig"),",")
	for temp_rights_count=0 to ubound(temp_check_rights)
	    if trim(temp_check_rights(temp_rights_count)) = "4" then
			rights_check_passkey = trim(temp_check_rights(temp_rights_count))
		end if
	next
	if rights_check_passkey <> "4" then
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('����Ȩ�޲��㣬�벻Ҫ�Ƿ�������������ģ�飬���������˺Ž���ϵͳ�Զ�ɾ��!');this.location.href='../login.asp';</Script>"
	Response.end
	end if%>
<%Dim Act
  Act=Request.Form("act")
  openData()
  If Act="save" Then
     Call SaveData()
  Else
     Call Main()
  End If
  Call CloseDataBase()
  
  Sub SaveData()
     Pname=Trim(Request.Form("Pname"))
	 Ptype=Trim(Request.Form("Ptype"))
'  sqlsize ="select Pname from Sbe_Down where Ptype ='"&Ptype&"'"
'  set rssize=conn.execute(sqlsize)
'  if not (rssize.eof and rssize.bof) then
'	Response.Write "<Script Language=JavaScript>alert('���ݿ����Ѵ���ͬ���ĵ���������!');this.location.href='javascript:history.back();';</'Script>"
'	response.End
' end if
'  rssize.close
' set rssize=nothing
	 Tid=Cint(Request.Form("Tid"))
	 Bpic=Trim(Request.Form("Bpic"))
	 spic=Trim(Request.Form("Spic"))
	 Price=Trim(Request.Form("Price"))
	 if price="" then
	 price=0
	 end if
	 leibie=Request.Form("leibie")
	 Tuijian=Request.Form("Tuijian")
	 succeed=Request.Form("succeed")
     Content = ""
     For i = 1 To Request.Form("content").Count
       Content = Content & Request.Form("content")(i)
     Next
     Uploadfile=request.Form("Uploadfile")
    ' Content2 = ""
    ' For i = 1 To Request.Form("content2").Count
       Content2 = Request.Form("content2")
   '  Next
     Uploadfile2=request.Form("Uploadfile2")
       content3 = Request.Form("content3")
     Uploadfile3=request.Form("Uploadfile3") 
	 
	 set rs_max=server.CreateObject("adodb.recordset")
     sql="select max(sequence) as maxid from Sbe_Down"
     rs_max.open sql,conn,1,1
     if isnull(rs_max("maxid")) then
        sequence=1
     else
        sequence=rs_max("maxid")+1
     end if
     rs_max.close
     set rs_max=nothing	 
  sqlsize ="select * from Sbe_Down_Class where ID ="&Tid
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
	 Set Rs=Server.CreateObject("adodb.recordset")
	 sql="select * From Sbe_Down Where ID=0"
	 Rs.Open Sql,Conn,1,3
	    Rs.AddNew
		Rs("Pname")=Pname
		Rs("Tid")=Tid
		Rs("bigclass")=bigclass
		Rs("Ptype")=Ptype
		Rs("Bpic")=Bpic
		'if spic<>"" then
		Rs("spic")=spic
		'end if
		Rs("Price")=Price
		Rs("Tuijian")=Tuijian
		Rs("Content")=Content
		Rs("Uploadfile")=Uploadfile
		Rs("Content2")=Content2
		Rs("Uploadfile2")=Uploadfile2
		Rs("Content3")=Content3
		Rs("Uploadfile3")=Uploadfile3
		Rs("Show")=request("Show")
		Rs("leibie")=leibie
		Rs("Sequence")=Sequence
		Rs("Succeed")=succeed
		Rs("gg")=trim(request("gg"))
		Rs("fileSize")=trim(request("fileSize"))		
		if request("datet")<>"" then Rs("datet")=request("datet")
		rs.update
	rs.Close
	Set Rs=Nothing 
    response.Write("<script language=javascript>alert('����������ӳɹ���');window.location.href='add.asp?tid="&tid&"';</script>")
	response.End()
  End Sub
  
  Sub Main()
  Tid=request("tid")
  if tid="" then tid=0
  tid=cint(tid)
  Set Rs=Server.CreateObject("adodb.recordset")
  Sql="select picAuto from Sbe_WebConfig"
  Rs.Open sql,Conn,1,1
     PicAuto=rs(0)
  Rs.Close
  Set Rs=Nothing
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
 document.form1.addbtn.disabled=true;
 document.form1.addbtn.value="���Ժ�..."
  return true;
}
</script>
<script language="JavaScript" src="../include/meizzDate.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="25"><font color="#6A859D">������������ &gt;&gt; ��ӵ���������Ϣ</font></td>
  </tr>
  <tr>
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>
<br>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <form name="form1" method="post" action="add.asp" onSubmit="return check()">
    <tr> 
      <td width="15%" height="25" align="center">��������</td>
      <td colspan="2"><select name="Tid" class="input_length">
<!--          <option>��ѡ��...</option>-->
         <%Call ShowClass("Sbe_Down",tid)%>
        </select> </td>
    </tr>
    <tr> 
      <td height="25" align="center">�� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</td>
      <td colspan="2"> <input name="Pname" type="text" id="Pname" size="30" maxlength="200"  class="input"/></td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center">�� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</td>
      <td colspan="2"> <input name="Ptype" type="text" id="Ptype" size="30" maxlength="200"  class="input"/></td>
    </tr>	
    <tr > 
      <td height="25" align="center">��ͼ</td>
      <td width="23%"> <input name="Bpic" type="text" id="Bpic" size="30" maxlength="200"  class="input"/></td>
      <td width="61%"> <iframe src="../upload/upload.asp?Form_Name=form1&UploadFile=Bpic" width="80%" height="25" frameborder="0" scrolling="no"></iframe>360*240</td>
    </tr>
    <tr> 
      <td height="25" align="center">Сͼ</td>
      <td width="29%"> <input name="Spic" type="text" id="Spic" size="30" maxlength="200"  class="input"/></td>
      <td width="56%" valign="middle"> <iframe style="top:2px" ID="UploadFiles" src="../upload/Download_Photo.asp?PhotoUrlID=1" frameborder=0 scrolling=no width="320" height="25"></iframe>
      75*50</td>
    </tr>
	     <tr  style="display:none">
            <td height="25" align="center">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</td>
            <td colspan="2"><input name="gg" type="text" id="gg" size="30" maxlength="200"  class="input"/></td>
    </tr>
          <tr  style="display:none"> 
            <td height="25" align="center">�ļ���С</td>
            <td colspan="2">
<input name="FileSize" type="text" class="input" id="fileSize" size="30">
            K </td>
          </tr>
    <tr  style="display:none"> 
      <td height="25" align="center">��������</td>
      <td colspan="2"><input type="text" name="succeed" id="succeed" value="" class="input"></td>
    </tr>
   <tr  style="display:none"> 
      <td height="25" align="center">��ϸ˵��:
      <input name="content" type="hidden" id="content"> <input name="uploadfile" type="hidden" id="uploadfile"></td>
      <td colspan="2"><iframe ID="eWebEditor1" src="../eWebEditor/ewebeditor.asp?id=content&style=sbe&savefilename=uploadfile" frameborder="0" scrolling="no" width="100%" HEIGHT="350"></iframe></td>
    </tr>
<tr class="display">
    <td align="center">��ҳ�Ƽ�:</td>
    <td colspan="2">
 <input type="radio" name="Tuijian" value="1">
        �� &nbsp;&nbsp; <input name="Tuijian" type="radio" value="0" checked="checked">
        ��</td>
  </tr>
<tr <%=banben_display%>>
    <td align="center">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��:</td>
    <td colspan="2">
 <input type="radio" name="leibie" value="1" checked="checked">
        �� &nbsp;&nbsp; <input name="leibie" type="radio" value="2">
        Ӣ</td>
  </tr>
  <tr>
    <td align="center"><span class="lv">����</span>ʱ��:</td>
    <td colspan="2"><input name="datet" type="text" id="datet" onFocus="setday(this)" value="<%=date()%>" class="input"></td>
  </tr>
<tr>
    <td align="center">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ʾ:</td>
    <td colspan="2">
 <input type="radio" name="Show" value="1" checked="checked">
 �� &nbsp;&nbsp; <input name="Show" type="radio" value="0">
 ��</td>
  </tr>
    <tr> 
      <td height="40" align="center">&nbsp;</td>
      <td colspan="2"> <input name="addbtn" type="submit" value=" ���� " class="sbe_button"> 
        &nbsp; <input type="reset" name="Submit2" value=" ��� " class="sbe_button">
        <input name="act" type="hidden" id="act" value="save"> </td>
    </tr>
  </form>
</table>
<br>
</body>
</html>
<% End Sub%>