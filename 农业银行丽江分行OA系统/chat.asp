<%@ language="javascript" %>
<html>
<head>
<title>�����������</title>
<link rel="stylesheet" href="images/vblife.css" type="text/css">
</head>
<script language="vbscript" runat=server>
dim kk
kk=cstr(now())
</script>
<% 
var date=new Date()
var sendd;
switch (parseInt(Request.Form("op")))
{
case 1:
	Application("MeetingBegin")=1
	Application("title")="" + Request("MeetingName")
	Application("msg")=""
	Application("num")=0
	Application("build")=Session("id")
	break;
case 2:
	if (Application("MeetingBegin")){
    mmessage="����������ġ�"
    Session("isin")=1
	Application("num")=Application("num")+1
	
	Application.lock
		Application("msg")="<font color=blue>"+Session("id")+":"+"</font>"+mmessage+"<br>"+Application("msg");
		if (Application("msg").length>2000)		Application("msg")= Application("msg").substr(0,1000);
	Application.UnLock
	Session("sendd")=1}else{ Session("isin")=0}
    break;
case 4:
	if (Application("MeetingBegin")){
	mmessage="" + Request.Form("message");
	if (mmessage=="") mmessage="�����ұ��ֳ�Ĭ����"
	Application.lock
		Application("msg")="<font color=blue>"+Session("id")+":"+"</font>"+mmessage+"<br>"+Application("msg");
		if (Application("msg").length>2000)		Application("msg")= Application("msg").substr(0,1000);
	Application.UnLock
	Session("sendd")=1}else{ Session("isin")=0}
    break;
case 5:
    Application("msg")="";
	Session("sendd")=1
	break;
case 6:
	Session("isin")=0;
	Application("num")=Application("num")-1
	Session("sendd")=0
	break;
case 7:
	Application("MeetingBegin")=0;
	Session("isin")=0;
	Application("num")=0
	Session("sendd")=0
	
	break;
case 8:
   	var connstr="DBQ="+Server.MapPath("db/system1.asa")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
    var conn=Server.CreateObject("ADODB.CONNECTION")
    conn.Open(connstr)
    var sql="insert into learn(title,content,type) values('" +Application("title")+"','"+Application("msg")+"','43')"
	Response.Write(sql)
	conn.Execute(sql);
	Application("msg")="";
	Application("title")=""
};


%>
<body leftmargin="0" topmargin="0" <% if (Session("sendd")==1){%>OnLoad="vbscript:document.all.form1.message.focus"<%}%>>

<%if (Application("MeetingBegin")!=1)  {  if  ((Session("level")>="2")||(Session("level")=="0")) {

    if (Application("build")==Session("id")) {%>
	<form method="post" action="chat.asp">
	<input type="submit" name="Submit3" value="��������¼" class=css0 >
		<input type="hidden" name="op" value=8 >
	</form>
	
	<% Application("build")="";Response.End}

%>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>

<table width="100%" border="0"   class=css4 align="center">
  <form method="post" action="chat.asp">
	<tr> 
	  <td align="right"> <font color="#FFFFFF" > </font></td>
	  <td width="88" bgcolor="#000000" align="right"><font color="#FFFFFF">�������ƣ�</font></td>
	  <td width="268" bgcolor="#00FFFF"> 
		<input type="text" name="MeetingName"  maxlength="40" size="40" class="css0">
	  </td>
	  <td>&nbsp;</td>
	</tr>
	<tr> 
	  <td colspan="2">&nbsp;</td>
	  <td colspan="2"> 
		<input type="submit" name="Submit2" value="��ʼ����"  class="css0">
		<input type="hidden" name="op" value=1 >
	  </td>
	</tr>
  </form>
</table>
<% 
Response.End();
} else
{
Response.Write("����û���κλ������ٿ�����û�з�������Ȩ�ޣ�");
Response.End();}
}%>
<% if (Session("isin")!=1) { %>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table width="100%" border="0"   class=css4 align="center">
  <form method="post" action="chat.asp">
    <tr> 
      <td align="right" height="24">&nbsp;</td>
      <td height="24" bgcolor="#000000" align="right"><font color="#FFFFFF">�������ƣ�</font></td>
      <td height="24" bgcolor="#00FFFF"><%=(Application("title"))%> </td>
      <td height="24">&nbsp;</td>
    </tr>
    <tr> 
      <td align="right" height="24">&nbsp;</td>
      <td height="24" bgcolor="#000000" align="right"><font color="#FFFFFF">���鷢���ˣ�</font></td>
      <td height="24" bgcolor="#00FFFF"><%=(Application("build"))%> </td>
      <td height="24">&nbsp;</td>
    </tr>
    <tr> 
      <td align="right" height="24">&nbsp;</td>
      <td height="24" bgcolor="#000000" align="right"><font color="#FFFFFF">�᳡������</font></td>
      <td height="24" bgcolor="#00FFFF"><%=(Application("num"))%> </td>
      <td height="24">&nbsp;</td>
    </tr>
    <tr> 
      <td align="right"> <font color="#FFFFFF" > </font></td>
      <td width="88" bgcolor="#000000" align="right"><font color="#FFFFFF" >���������</font></td>
      <td width="268" bgcolor="#00FFFF"> 
		<input type="text" name="UserName" value="<%=Session("id")%>"  maxlength="40" size="40" class="css0" >
      </td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="2">&nbsp;</td>
      <td colspan="2"> 
        <input type="submit" name="Submit2" value="����᳡"  class="css0">
		<input type="hidden" name="op" value=2 >
      </td>
    </tr>
  </form>
</table>
<% } %> <% if (Session("isin")==1 ) { %> 
<table border="0" cellspacing="1" cellpadding="1" class=css2 width="100%" height="100%" align="center">
  <tr> 
	<td  align="left" height="20" width="451" bgcolor="#00FFFF"    ><b>��������</b>��<font color="#0000FF"><%=Application("title")%> 
	  </font></td>
    <td align="left" width="161" height="20" bgcolor="#00FFFF"  ><b>�᳡����</b>��<font color="#0000FF"><%=(Application("num"))%> 
	  </font></td>
    <td  align="left" width="157" height="20" bgcolor="#00FFFF" ><b>������</b>��<font color="#0000FF"><%=Session("id")%> 
	  </font></td>
  </tr>
  <tr> 
	<td  align="right" valign="top" colspan=3 class="title"> <iframe name=chat frameborder="no" class=css0 width=100% height=100% src="chattext.asp"></iframe> 
	</td>
  </tr>

  <form method="post" name= "form1" action="chat.asp"> 
	<tr valign="top" > 
	  <td bgcolor="#00FFFF"   height="24" align="left" valign="middle" colspan=3> 
		<b>��Ҫ����</b>�� 
		<input type="text" name="message" maxlength="400" size="44" class="css0">
		<input type="submit" name="Submit" value="����" class=css0 OnClick="vbscript:Document.all.form1.op.value=4">
       <%if (Application("build")==Session("id")) {%><input type="submit" name="Submit3" value="���" class=css0  OnClick="vbscript:Document.all.form1.op.value=5"><%}%>
        <input type="submit" name="Submit3" value="�˳�" class=css0 OnClick="vbscript:Document.all.form1.op.value=6">
        <%if (Application("build")==Session("id")) {%><input type="submit" name="Submit3" value="��������" class=css0 OnClick="vbscript:Document.all.form1.op.value=7"><%}%>
        
		<input type="hidden" name="op" value=2 >
	  </td>
  </tr>  </form>
</table>
<% } %> 
</body>
</html> 















