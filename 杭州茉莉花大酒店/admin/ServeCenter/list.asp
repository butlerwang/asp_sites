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
if Request("method")="ok" then
	if Request("chkid") = "" then
	Response.Write("<script>alert(""�γ�����,��Ҫɾ��������ID����Ϊ�ա�"");location.href="""& Request.ServerVariables("HTTP_REFERER") &""";</script>")
	   Response.end     
	end if
	chkid = Request("chkid")
	Conn.Execute("delete from Guest_book where ID in("& Request("chkid") &")")
	Response.Write("<script>alert(""ɾ���ɹ�"");location.href="""& Request.ServerVariables("HTTP_REFERER") &""";</script>") 
	Response.End()
end if
if request("keyword")="ok" then
      keyword=request("keyword")
      ly_time1=Cdate(request("ly_time1"))
      ly_time2=Cdate(request("ly_time2"))
	  leibie=request("leibie")
      flag=request("flag")
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��������</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript>
// ��������
NS4 = document.layers && true;
IE4 = document.all && parseInt(navigator.appVersion) >= 4;
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall')
       e.checked = form.chkall.checked;
    }
  }
</script>
<script language="JavaScript" src="../include/meizzDate.js"></script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="19%" height="25"><font color="#6A859D">�������� &gt;&gt; �������� </font></td>   
      <td width="81%">&nbsp;       </td>   
  </tr>
  <tr> 
    <td height="1" colspan="2" background="../images/dot.gif"></td>
  </tr>
</table>
<table width="98%"  border="0" cellpadding="0" cellspacing="1" align="center" bgcolor="#99CCFF">
  <tr bgcolor="#E8F1FF">
    <td height="22" colspan="10" class="table_1_2"><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="table_1_2">
      <tr align="left" class="table_1_3_1">
        <form name="newsearch" action="list.asp" method="post">
          <td width="88" align="right"  valign="middle" height="25" class="sbe_table_title">���Բ��ң�</td>
          <td width="806"  colspan="9" align="left"  valign="middle" class="sbe_table_title">ʱ�䣺
            <input name="keyword" type="hidden" value="ok">
                <input name="ly_time1" type="text" class="sbe_button" id="ly_time1"   style="ime-mode:Disabled" onFocus="setday(this)" onKeyPress="return   event.keyCode>=48&&event.keyCode<=57||event.keyCode==45"  <%if ly_time1<>"" then response.Write("value='"&ly_time1&"'") else response.Write("value='"&date()-6&"'") end if%> size="15" onpaste="return!clipboardData.getData('text').match(/\D/)" ondragenter="return false">
            ��
            <input name="ly_time2" type="text" class="sbe_button" id="ly_time2"   style="ime-mode:Disabled" onFocus="setday(this)" onKeyPress="return   event.keyCode>=48&&event.keyCode<=57||event.keyCode==45" <%if ly_time2<>"" then response.Write("value='"&ly_time2&"'") else response.Write("value='"&date()&"'") end if%> size="15" onpaste="return!clipboardData.getData('text').match(/\D/)" ondragenter="return false">
            &nbsp;	״̬��
			<%if hy_message=true then%>
            <select name="flag" size="1" class="sbe_button" style="width:100">
              <option value="" <%if flag ="" then response.Write("selected") end if%>>ȫ������</option>
              <option value="0" <%if flag ="0" then response.Write("selected") end if%>>δ�鿴������</option>
              <option value="1" <%if flag ="1" then response.Write("selected") end if%>>δ�ظ�������</option>
              <option value="2" <%if flag ="2" then response.Write("selected") end if%>>�ѻظ�������</option>
            </select>
			<%else%>
			<select name="flag" size="1" class="sbe_button" style="width:100">
              <option value="" <%if flag ="" then response.Write("selected") end if%>>ȫ������</option>
              <option value="0" <%if flag ="0" then response.Write("selected") end if%>>δ�鿴������</option>
              <option value="1" <%if flag ="1" then response.Write("selected") end if%>>�Ѳ鿴������</option>
            </select>
			<%end if%>
			&nbsp;
			<select <%=banben_display%> name="leibie" size="1" class="sbe_button">
              <option value="" <%if leibie ="" then response.Write("selected") end if%>>ȫ�����</option>
              <option value="1" <%if leibie ="1" then response.Write("selected") end if%>>��������</option>
              <option value="2" <%if leibie ="2" then response.Write("selected") end if%>>Ӣ������</option>
            </select>
            <input type="submit" name="Submit2" value="��ʼ����" class="sbe_button" title="��Ϣ����">
            &nbsp;
            <input type="button" name="ref" value="ˢ��ҳ��" onClick="location.href='list.asp'"  class="sbe_button" title="�����Զ�ˢ�µ��"></td>
        </form>
      </tr>
    </table></td>
  </tr>
  <tr align="left" bgcolor="#E8F1FF">
    <td  width="5%" align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button" ><strong>&nbsp;����</strong>&nbsp;</td>
    <td width="16%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"  style="display:none"><strong>&nbsp;��������</strong></td>
    <td width="17%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"  ><strong>��������</strong></td>
    <td width="9%"  height="22" align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>&nbsp;��ϵ��&nbsp;</strong></td>
    <td width="13%"  height="22" align="center" valign="middle" nowrap bgcolor="#FFFFFF"  style="display:none"><strong>&nbsp;��ϵ�绰&nbsp;</strong></td>
    <td width="12%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF"   style="display:none"><strong>&nbsp;E-mail</strong>&nbsp;</td>
    <td width="10%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>��������</strong></td>
    <td width="6%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button" <%=banben_display%>><strong>&nbsp;���&nbsp;</strong></td>
    <td width="6%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>&nbsp;״̬&nbsp;</strong></td>
    <td width="6%" align="center"  valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>�鿴</strong></td>
  </tr>
  <% 
  set rs1=server.createobject("adodb.recordset")
  Sql = "Select * from Guest_book where flag=1 "
  if flag <> "" then Sql = Sql&"and status ="&flag&" "
  if leibie<>"" then Sql=Sql&" and leibie="&leibie&""
  if ly_time1 <> "" and ly_time2<>"" then Sql = Sql&"and lytime between #"&ly_time1&"# and #"&ly_time2&"# "
  Sql = Sql&" order by ID desc"
  rs1.open Sql,conn,1,1
  if rs1.eof and rs1.bof then
  if hy_message=true then
    if flag="2" then
	    Response.Write "<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time&"��û��<font color='#FF0000'>�ѻظ�</font>��������Ϣ��</td></tr>"
    elseif flag="1" then
        response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time1&"��"&ly_time2&"��û��<font color='#FF0000'>δ�ظ�</font>��������Ϣ</td></tr>")
      elseif flag="0" then
        response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time1&"��"&ly_time2&"��û��<font color='#FF0000'>δ�鿴</font>��������Ϣ</td></tr>")
	  else
	    response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>��û��������Ϣ</td></tr>")
     end if
	else
	 if flag="1" then
        response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time1&"��"&ly_time2&"��û��<font color='#FF0000'>�Ѳ鿴</font>��������Ϣ</td></tr>")
      elseif flag="0" then
        response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time1&"��"&ly_time2&"��û��<font color='#FF0000'>δ�鿴</font>��������Ϣ</td></tr>")
	  else
	    response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>��û��������Ϣ</td></tr>")
     end if
	end if
	 else
	  %>
  <tr align="left" bgcolor="#E8F1FF" style="display:none">
    <td height="23" colspan="10" valign="middle" bgcolor="#FFFFFF" class="sbe_button"><%
	  rs1.pagesize=15
      totalrecord=rs1.recordcount
      totalpage=rs1.pagecount
	  pagenum=rs1.pagesize
      rs1.movefirst
      if request("page")="" then
         nowpage=1
		 elseif cint(request("page"))>totalpage then
         nowpage=totalpage
        else
      nowpage=request("page")
      end if
      nowpage=cint(nowpage)
      rs1.absolutepage=nowpage%></td>
  </tr>
  <form action="list.asp?method=ok" name="form1" method="post">
	<%p=1
	  Do while not rs1.EOF and p<=pagenum%>
    <tr align="left" bgcolor="#E8F1FF">
      <td align="center" valign="middle" bgcolor="#FFFFFF"><input type="checkbox" name="chkid" value="<%=rs1("ID")%>"></td>
      <td align="left" valign="middle" nowrap bgcolor="#FFFFFF" style="display:none">&nbsp;<%=gotTopic(rs1("lyremark"),50)%>&nbsp;</td> 
      <td align="left" valign="middle" nowrap bgcolor="#FFFFFF" ><%=gotTopic(rs1("lytheme"),20)%></td>
      <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><%=gotTopic(rs1("lyname"),20)%></td>  
      <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"   style="display:none"><%=rs1("lytel")%></td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF"    style="display:none"><%=(rs1("lyemail"))%></td>
      <td align="center" valign="middle" bgcolor="#FFFFFF"><%=rs1("lytime")%></td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF" <%=banben_display%>><%if rs1("leibie")=1 then response.Write("��") else response.Write("Ӣ") end if%></td> 
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF">&nbsp;
      <%
	  if hy_message=true then
	  if rs1("status")=2 then
	       response.Write("�ѻظ�")
	    elseif rs1("status") = 1 then 
	       response.Write("<font color='#000099'>δ�ظ�</font>")
         else
	       response.Write("<font color='#FF0000'>δ�鿴</font>")
	    end if
	  else
	  if rs1("status")=1 then
	       response.Write("�Ѳ鿴")
         else
	       response.Write("<font color='#FF0000'>δ�鿴</font>")
	    end if
	end if
	 %>
        &nbsp;</td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF"><a href="ly_show.asp?id=<%=rs1("ID")%>"><img src="../images/edit.gif" border="0"></a></td>
    </tr>
	<%p=p+1
      rs1.moveNext
      loop%>
    <tr align="left" bgcolor="#E8F1FF">
      <td  height="25" colspan="10" valign="middle" bgcolor="#FFFFFF">&nbsp;���������ȫ��ѡ��
        <input type="checkbox" name="chkall" onClick="javascript:CheckAll(this.form)">
          <input type="submit" name="Submit" value="ɾ��" class="sbe_button" >
        &nbsp;&nbsp; </td>
    </tr>
  </form>
  <tr>
              <td height="25" colspan="10" align="right" bgcolor="#FFFFFF" class="zi11"><a href="?page=1&keyword=<%=keyword%>&ly_time1=<%=ly_time1%>&ly_time2=<%=ly_time2%>&leibie=<%=leibie%>&flag=<%=flag%>" title="��ҳ" class="zi11">��ҳ</a>&nbsp;&nbsp;
    <%if nowpage>1 then%><a href="?page=<%=nowpage-1%>&keyword=<%=keyword%>&ly_time1=<%=ly_time1%>&ly_time2=<%=ly_time2%>&leibie=<%=leibie%>&flag=<%=flag%>" title="��һҳ" class="zi11">��һҳ</a><%else%>��һҳ<%end if%>&nbsp;&nbsp;<%if nowpage<totalpage then%><a href="?page=<%=nowpage+1%>&keyword=<%=keyword%>&ly_time1=<%=ly_time1%>&ly_time2=<%=ly_time2%>&leibie=<%=leibie%>&flag=<%=flag%>" title="��һҳ" class="zi11">��һҳ</a><%else%>��һҳ<%end if%>&nbsp;&nbsp;<a href="?page=<%=totalpage%>&keyword=<%=keyword%>&ly_time1=<%=ly_time1%>&ly_time2=<%=ly_time2%>&leibie=<%=leibie%>&flag=<%=flag%>" title="���ҳ" class="zi11">���ҳ</a>&nbsp;&nbsp;ҳ�Σ�<%=nowpage%>/<%=totalpage%>&nbsp;&nbsp;<%=pagenum%>����¼/ҳ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  </tr>
  		  <%end if
		  rs1.close%>
</table>
<%Call CloseDataBase()%>
</body>
</html>
