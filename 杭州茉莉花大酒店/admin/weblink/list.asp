<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%Call OpenData()
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('���糬ʱ�������㻹û�е�½! ');this.location.href='../login.asp';</script>"
end if
  IF instr(webConfig,", 9")=0 or instr(session("manconfig"),", 9")=0 Then'��վ��������
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('����Ȩ�޲��㣬�벻Ҫ�Ƿ�������������ģ�飬���������˺Ž���ϵͳ�Զ�ɾ��!');this.location.href='../login.asp';</Script>"
Response.end
end if
act=Request("act")
linkid=Request("id")
IF act="list" Then	
	IF Request("chkid")<>"" Then
	  msg="ɾ���ɹ�"
	  Conn.Execute("delete from Sbe_Weblink where id in("& Request("chkid") &")")
	  Call MessageBoxOK(msg)
     End IF
ElseIF act="up"	Then
  Call Up()
ElseIf act="down" Then
  Call down()
ElseIF act="pass" Then
  Call pass()
ElseIF act="leibie" Then
  Call leibie()
end if

Private Sub MessageBoxOK(strValue)

	With Response
		.Write "<script>" & vbcrlf
		.Write "alert('"+strValue+"');" & vbcrlf
		.Write "this.location.href='list.asp'" & vbcrlf
		.Write "</script>" & vbcrlf
	End With
End Sub

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

// ѡ��ָ����tab.
function selectTab(tab) {
    var form   = document.tabform;
    var TabLayer1 = getLayerStyle("TabLayer1");
    var TabLayer2 = getLayerStyle("TabLayer2");

    if (tab == "TabLayer2") {
        _showLayer(TabLayer1, false);
        _showLayer(TabLayer2, true);


    } else {
        _showLayer(TabLayer2, false);
        _showLayer(TabLayer1, true);

    }
    return true;
}

function _showLayer(layer, display) {
    if (layer) {
        if (display) {
            layer.display = "block";
        } else {
            layer.display = "none";
        }
    }
}

// ȡ��ָ��id��layer
function getLayerStyle(id) {
    if (IE4 && document.all(id)) {
        return document.all(id).style;
    } else if (NS4 && document.layers[id]) {
        return document.layers[id];
    } else {
        return null;
    }
}

function SelectAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name == 'chkid')
       e.checked = form.ChkAll.checked;
    }
	}

</SCRIPT> 
</head>

<body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="19%" height="25"><font color="#6A859D">¥�̱�־</font><font color="#6A859D"> &gt;&gt;¥�̱�־</font></td>   
      <td width="81%">&nbsp;       </td>   
  </tr>
  <tr> 
    <td height="1" colspan="2" background="../images/dot.gif"></td>
  </tr>
</table>
<br>
<table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0" id="sbe_table">
  <tr>
    <td width="6%">&nbsp;</td>
    <td width="16%" align="center">����</td>
    <td width="19%" align="center">���ӵ�ַ</td>
    <td width="14%" align="center">��������</td>
    <td width="14%" align="center">ͼƬ</td>
    <td width="7%" align="center">����</td>
    <td width="6%" align="center">����</td>
    <td width="6%" align="center" <%=banben_display%>>���</td>
    <td width="6%" align="center">���</td>
    <td width="6%" align="center">�޸�</td>
  </tr>
  <form action="list.asp?act=list" method="post">
  <%
  Set oRs=Server.CreateObject("adodb.recordset")
  sql="select * from Sbe_Weblink order by orderid desc"
  oRs.Open sql,conn,1,1
  IF not(oRs.eof and oRs.bof) Then
     oRs.pagesize=10
     totalrecord=oRs.recordcount
     totalpage=oRs.pagecount
     pagenum=oRs.pagesize
     oRs.movefirst
   if request("page")="" then
     nowpage=1
   elseif cint(request("page"))>totalpage then
     nowpage=totalpage
   else
     nowpage=request("page")
   end if
   nowpage=cint(nowpage)
   oRs.absolutepage=nowpage
   j_5=1
   Do while not oRs.EOF and j_5<=pagenum%>
  <tr>
    <td align="center"><input name="chkid" type="checkbox" id="chkid" value="<%=oRs("id")%>"></td>
    <td align="center"><%=oRs("companyname")%></td>
    <td align="center"><%=oRs("URL")%></td>
    <td align="center"><%IF	oRs("linktype")=true Then Response.Write "ͼƬ����" Else Response.Write"��������" End IF%></td>
    <td align="center"><%If oRs("picurl")<>"" Then%><img src="../../uploadfile/<%=oRs("picurl")%>" width="88" height="31"><%else%>������������<%End IF%></td>
    <td align="center"><a href="?id=<%=oRs("id")%>&act=up"><img src="../images/up.gif" border="0" title="����"></a></td>
    <td align="center"><a href="?id=<%=oRs("id")%>&act=down"><img src="../images/downl.gif" border="0" title="����"></a></td>
    <td align="center" <%=banben_display%>><a href="?id=<%=oRs("id")%>&act=leibie"><%Call JudgeMent1(oRs("leibie"))%></a></td>
    <td align="center"><a href="?id=<%=oRs("id")%>&act=pass"><%Call JudgeMent(oRs("status"))%></a></td>
    <td align="center"><a href="weblink.asp?act=modify&id=<%=oRs("id")%>"><img src="../images/edit.gif" border="0"></a></td>
  </tr>
<%j_5=j_5+1
  oRs.Movenext
  Loop%>
  <tr>
          <td height="18" valign="middle" class="ziti3" colspan="10">
		  
<a href="?page=1" title="��ǰҳ" class="ziti3">��ǰҳ</a>  <%if nowpage>1 then%><a href="?page=<%=nowpage-1%>" title="��һҳ" class="ziti3">��һҳ</a><%else%>��һҳ<%end if%> &nbsp;<%
	if totalpage<=6 then
	   totalpage_1=totalpage
	 else
	   totalpage_1=6
	 end if
	   for i=1 to totalpage_1
	      response.Write("<span class='ziti3'><a href='?page="&i&"' title='"&i&"' class='ziti3'>"&i&"</a></span>")
		  if i<totalpage_1 then response.Write("&nbsp;")
	next
	%>&nbsp;&nbsp;<%if nowpage<totalpage then%><a href="?page=<%=nowpage+1%>" title="��һҳ" class="ziti3">��һҳ</a><%else%>��һҳ<%end if%>  <a href="?page=<%=totalpage%>" title="���ҳ" class="ziti3">���ҳ</a>    ҳ�Σ�<span class="ziti3"><%=nowpage%></span>/<span class="ziti3"><%=totalpage%></span>ҳ    ��<%=totalrecord%>����¼ <span class="ziti13"><%=pagenum%></span>����¼/ҳ
		  </td>
        </tr>    
<%End IF
 oRs.Close:Set oRs=Nothing
 
 %> 
  <tr>
    <td colspan="10">&nbsp;&nbsp;<input type="checkbox" name="ChkAll" onClick="SelectAll(this.form)">
        ȫѡ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      <input type="submit" name="Submit" value="ȫ��ɾ��" class="sbe_button" ></td>
  </tr>
  </form>  
</table>
<br>
<%Call CloseDataBase()%>
</body>
</html>
<%
Private Sub pass()
'��˹���
id=request.Querystring("id")
'IF id="" Then Exit Sub
Set objRs=Server.Createobject("adodb.recordset")
sql="select status from Sbe_Weblink where id=" &id
objRs.Open sql,conn,1,3
  IF objRs.Fields(0).Value=True Then
  objRs.Fields(0).Value=0 
Else
  objRs.Fields(0).Value=1
End IF
  objRs.Update
 msg="������óɹ�" 
 objRs.Close:set objRs=Nothing
' Call MessageBoxOK(msg)
  Response.Redirect(request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub leibie()
'��˹���
id=request.Querystring("id")
'IF id="" Then Exit Sub
Set objRs=Server.Createobject("adodb.recordset")
sql="select leibie from Sbe_Weblink where id=" &id
objRs.Open sql,conn,1,3
  IF objRs.Fields(0).Value=1 Then
  objRs.Fields(0).Value=2 
Else
  objRs.Fields(0).Value=1
End IF
  objRs.Update
 objRs.Close:set objRs=Nothing
 Response.Redirect(request.ServerVariables("HTTP_REFERER"))
End Sub

Private Sub Up()  
 '����  
    ID=Cint(Request.QueryString("ID"))	
    set rsUp=server.CreateObject("adodb.recordset")
	sql="select orderid from Sbe_Weblink where ID="&ID
	rsUp.open sql,conn,1,3	
	 set rs_up=server.CreateObject("adodb.recordset")
	 sql_up="select top 1 orderid from Sbe_Weblink where orderid >"&rsUp(0)	
	 sql_up=sql_up&" order by orderid"	
	 rs_up.open sql_up,conn,1,3
	 if not rs_up.eof then
	    Temp_sequence=rs_up(0)
		rs_up(0)=rsUp(0)
		rs_up.update		
		rsUp(0)=Temp_sequence
		rsUp.update
     end if
	 rs_up.close
	 set rs_up=nothing
	rsUp.close
	set rsUp=nothing
	'msg="���Ƴɹ�"
	'Call MessageBoxOK(msg)
End Sub
  
Private Sub Down()    
    ID=Cint(Request.QueryString("ID"))
	 set rs_DownSequence=server.CreateObject("adodb.recordset")
  sql="select orderid from Sbe_Weblink where id="&id
  rs_DownSequence.open sql,conn,1,3
     set rs_up=server.CreateObject("adodb.recordset")
	 sql_up="select top 1 orderid from Sbe_Weblink where orderid <"&rs_DownSequence(0)	 
	 sql_up=sql_up&" order by orderid desc"
	 rs_up.open sql_up,conn,1,3
	 if not rs_up.eof then
	    Temp_sequence=rs_up(0)
		rs_up(0)=rs_DownSequence(0)
		rs_up.update		
		rs_DownSequence(0)=Temp_sequence
		rs_DownSequence.update
     end if
	 rs_up.close
	 set rs_up=nothing
  rs_DownSequence.close  
  set rs_DownSequence=nothing
End Sub 
%>