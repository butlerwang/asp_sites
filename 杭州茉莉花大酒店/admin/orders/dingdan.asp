<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%
Call OpenData()
%>
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
	Response.Write "<Script Language=JavaScript>alert('����Ȩ�޲��㣬�벻Ҫ�Ƿ�������������ģ�飬���������˺Ž���ϵͳ�Զ�ɾ��!');this.location.href='../login.asp';</Script>"
	Response.end
	end if%>
<%
if Request("method")="ok" then
	if Request("chkid") = "" then
	Response.Write("<script>alert(""�γ�����,��Ҫɾ���Ķ���ID����Ϊ�ա�"");location.href="""& Request.ServerVariables("HTTP_REFERER") &""";</script>")
	   Response.end     
	end if
	chkid = Request("chkid")
	Conn.Execute("delete from Sbe_order where ID in("& Request("chkid") &")")
	Response.Write("<script>alert(""ɾ���ɹ�"");location.href="""& Request.ServerVariables("HTTP_REFERER") &""";</script>") 
	Response.End()
end if
if request("keyword")="ok" then
   keyword=request("keyword")
   ly_time1=Cdate(request("ly_time1"))
   ly_time2=Cdate(request("ly_time2"))
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
    <td width="19%" height="25"><font color="#6A859D">�������� &gt;&gt;�������� </font></td>   
      <td width="81%">&nbsp;       </td>   
  </tr>
  <tr> 
    <td height="1" colspan="2" background="../images/dot.gif"></td>
  </tr>
</table>
<table width="98%"  border="0" cellpadding="0" cellspacing="1" align="center" bgcolor="#99CCFF">
  <tr bgcolor="#E8F1FF">
    <td height="22" colspan="9" class="table_1_2"><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="table_1_2">
      <tr align="left" class="table_1_3_1">
        <form name="newsearch" action="dingdan.asp" method="post">
          <td width="88" align="right"  valign="middle" height="25" class="sbe_table_title">�������ң�</td>
          <td width="806"  colspan="9" align="left"  valign="middle" class="sbe_table_title">ʱ�䣺
            <input name="keyword" type="hidden" value="ok">
                <input name="ly_time1" type="text" class="sbe_button" id="ly_time1"   style="ime-mode:Disabled" onFocus="setday(this)" onKeyPress="return   event.keyCode>=48&&event.keyCode<=57||event.keyCode==45"  <%if ly_time1<>"" then response.Write("value='"&ly_time1&"'") else response.Write("value='"&date()-6&"'") end if%> size="15" onpaste="return!clipboardData.getData('text').match(/\D/)" ondragenter="return false">
            ��
            <input name="ly_time2" type="text" class="sbe_button" id="ly_time2"   style="ime-mode:Disabled" onFocus="setday(this)" onKeyPress="return   event.keyCode>=48&&event.keyCode<=57||event.keyCode==45" <%if ly_time2<>"" then response.Write("value='"&ly_time2&"'") else response.Write("value='"&date()&"'") end if%> size="15" onpaste="return!clipboardData.getData('text').match(/\D/)" ondragenter="return false">
            &nbsp;	״̬��
            <select name="flag" size="1" class="sbe_button" style="width:150">
              <option value="" <%if flag ="" then response.Write("selected") end if%>>ȫ������</option>
              <option value="0" <%if flag ="0" then response.Write("selected") end if%>>δ����Ķ���</option>
              <!--                      <option value="1" <%'if flag ="1" then response.Write("selected") end if%>>δ�ظ��Ķ���</option>-->
              <option value="1" <%if flag ="1" then response.Write("selected") end if%>>�Ѵ���Ķ���</option>
            </select>
            <input type="submit" name="Submit2" value="��ʼ����" class="anniugaodu" title="��������">
            &nbsp;
            <input type="button" name="ref" value="ˢ��ҳ��" onClick="location.href='dingdan.asp'"  class="anniugaodu" title="�����Զ�ˢ�µ��"></td>
        </form>
      </tr>
    </table></td>
  </tr>
  <tr align="left" bgcolor="#E8F1FF">
    <td  width="4%" align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button" ><strong>&nbsp;����</strong>&nbsp;</td>
    <td width="16%"  height="22" align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>&nbsp;��������&nbsp;</strong></td>
    <td width="20%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>&nbsp;����&nbsp;</strong></td>
    <td width="9%"  height="22" align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>&nbsp;��סʱ��&nbsp;</strong></td>
    <td width="18%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>�뿪ʱ��</strong></td>
    <td width="12%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>�绰</strong>&nbsp;</td>
    <td width="11%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>�ύ����</strong></td>
    <td width="5%"  align="center" valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button"><strong>&nbsp;״̬&nbsp;</strong></td>
    <td width="5%" align="center"  valign="middle" nowrap bgcolor="#FFFFFF" class="sbe_button">&nbsp;<strong>�鿴</strong>&nbsp;</td>
  </tr>
  <%
		'-----�޸�ÿҳ��ʾ���� Start--------
	const MaxPerPage=15
	'-----�޸�ÿҳ��ʾ���� End  --------
   	dim totalPut
   	dim CurrentPage
	if not isempty(request("page")) then
      		currentPage=cint(request("page"))
   	else
      		currentPage=1
   	end if 
  set rs=server.createobject("adodb.recordset")
  Sql = "Select * from Sbe_order where 5=5 "
  if flag <> "" then Sql = Sql&"and status ="&flag&" "  
  if ly_time1 <> "" and ly_time2<>"" then Sql = Sql&" and timing between #"&ly_time1&"# and #"&ly_time2&"# "
  Sql = Sql&" order by ID desc"
  rs.open Sql,conn,1,1
  if rs.eof and rs.bof then
'     if flag="2" then
'	    Response.Write "<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time&"��û��<font color='#FF0000'>�ѻظ�</font>�Ķ�����Ϣ��</td></tr>"
    if flag="1" then
        response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time1&"��"&ly_time2&"��û��<font color='#FF0000'>�Ѵ���</font>�Ķ�����Ϣ</td></tr>")
      elseif flag="0" then
        response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>"&ly_time1&"��"&ly_time2&"��û��<font color='#FF0000'>δ����</font>�Ķ�����Ϣ</td></tr>")
	  else
	    response.Write("<tr align=left bgcolor=#E8F1FF><td height=20 colspan=12 valign=middle>��û�ж�����Ϣ</td></tr>")
     end if
	 else
	  %>
  <tr align="left" bgcolor="#E8F1FF" style="display:none">
    <td height="20" colspan="9" valign="middle" bgcolor="#FFFFFF" class="sbe_button"><%
 
       		totalPut=rs.recordcount
      		if currentpage<1 then
          		currentpage=1
      		end if
      		if (currentpage-1)*MaxPerPage>totalput then
	   		if (totalPut mod MaxPerPage)=0 then
	     			currentpage= totalPut \ MaxPerPage
	  		else
	      			currentpage= totalPut \ MaxPerPage + 1
	   		end if
      		end if
       		if currentPage=1 then
           		showpage totalput,MaxPerPage,"dingdan.asp"
            		showContent
            		showpage totalput,MaxPerPage,"dingdan.asp"
       		else
          		if (currentPage-1)*MaxPerPage<totalPut then
            			rs.move  (currentPage-1)*MaxPerPage
            			dim bookmark
            			bookmark=rs.bookmark
           			showpage totalput,MaxPerPage,"dingdan.asp"
            			showContent
             			showpage totalput,MaxPerPage,"dingdan.asp"
        		else
	        		currentPage=1
           			showpage totalput,MaxPerPage,"dingdan.asp"
           			showContent
           			showpage totalput,MaxPerPage,"dingdan.asp"
	      		end if
	   		end if
		rs.close
		set rs = nothing	
   	end if 
	sub showContent
	dim i 
	   	i=0
  %></td>
  </tr>
  <form action="dingdan.asp?method=ok" name="form1" method="post">
    <%do while not rs.eof%>
    <tr align="left" bgcolor="#E8F1FF">
      <td align="center" valign="middle" bgcolor="#FFFFFF"><input type="checkbox" name="chkid" value="<%=rs("ID")%>"></td>
      <td height="22" align="center" valign="middle" bgcolor="#FFFFFF">	 
	   <%set rs2=server.CreateObject("adodb.recordset")
	  sql="select * from sbe_product_class where id="&rs("roomtype")&""
	  rs2.open sql,conn,1,1
	  fang=rs2("classname")
	  rs2.close%>
	  <%=fang%>
</td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF">&nbsp;<%=rs("username")%>&nbsp;</td> 
      <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><a href="dd_show.asp?id=<%=rs("ID")%>"><%=gotTopic(rs("kssj"),20)%></a></td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF"><%=gotTopic(rs("lksj"),30)%></td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF"><%=(rs("tel"))%></td>
      <td align="center" valign="middle" bgcolor="#FFFFFF"><%=rs("time")%></td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF">&nbsp;
          <%if rs("status") = 1 then 
	       response.Write("<font color='#FF0000'>�Ѵ���</font>")
        else
	       response.Write("<font color='#000099'>δ����</font>")
	   end if
	 %>
        &nbsp;</td>
      <td align="center" valign="middle" nowrap bgcolor="#FFFFFF"><a href="dd_show.asp?id=<%=rs("ID")%>"><img src="../images/edit.gif" border="0"></a></td>
    </tr>
    <%
  i=i+1
	      if i>=MaxPerPage then exit do
  rs.movenext
  loop
  %>
    <tr align="left" bgcolor="#E8F1FF">
      <td  height="25" colspan="9" valign="middle" bgcolor="#FFFFFF">&nbsp;���������ȫ��ѡ��
        <input type="checkbox" name="chkall" onClick="javascript:CheckAll(this.form)">
          <input type="submit" name="Submit" value="ɾ��" class="sbe_button" >
        &nbsp;&nbsp; </td>
    </tr>
  </form>
  <tr align="left" bgcolor="#E8F1FF">
    <td  height="25" colspan="9" valign="middle" bgcolor="#FFFFFF"><%
    end sub
	function showpage(totalnumber,maxperpage,filename)
  	dim n

  	if totalnumber mod maxperpage=0 then
     		n= totalnumber \ maxperpage
  	else
     		n= totalnumber \ maxperpage+1          
  	end if
  	response.write "<table cellspacing=1 width='100%' border=0 colspan='4' ><form method=Post action="""&filename&"?keyword="& keyword&"&flag="& flag &"&year_time="& year_time &"&day_time="& day_time &"&month_time="& month_time &"""><tr><td align=right> "
  	if CurrentPage<2 then
    		response.write "��<b><font color=red>"&totalnumber&"</font></b>����¼&nbsp;��ҳ ��һҳ&nbsp;"
  	else
    		response.write "��<b><font color=red>"&totalnumber&"</font></b>����¼&nbsp;<a href="&filename&"?page=1&keyword="& keyword&"&flag="& flag &"&year_time="& year_time &"&day_time="& day_time &"&month_time="& month_time &">��ҳ</a>&nbsp;"
    		response.write "<a href="&filename&"?page="&CurrentPage-1&"&keyword="& keyword&"&flag="& flag &"&year_time="& year_time &"&day_time="& day_time &"&month_time="& month_time &">��һҳ</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		response.write "��һҳ βҳ"
  	else
    		response.write "<a href="&filename&"?page="&(CurrentPage+1)&"&keyword="& keyword &"&flag="& flag &"&year_time="& year_time &"&day_time="& day_time &"&month_time="& month_time &">"
    		response.write "��һҳ</a> <a href="&filename&"?page="&n&"&keyword="& keyword&"&flag="& flag &"&year_time="& year_time &"&day_time="& day_time &"&month_time="& month_time &">βҳ</a>"
  	end if
   	response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&CurrentPage&"</font>/"&n&"</strong>ҳ "
    	response.write "&nbsp;<b>"&maxperpage&"</b>����¼/ҳ "
%>
      ת����
      <select name='page' size='1' class="sbe_button" style="font-size: 9pt" onChange='javascript:submit()'>
          <%for i = 1 to n%>
          <option value='<%=i%>' <%if CurrentPage=cint(i) then%> selected <%end if%>>��<%=i%>ҳ</option>
          <%next%>
        </select>
        <%   
	response.write "</td></tr></FORM></table>"
end function
%></td>
  </tr>
</table>
<%Call CloseDataBase()%>
</body>
</html>
