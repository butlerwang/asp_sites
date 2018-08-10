<!-- #include file="inc/conn.asp"-->
<!-- #include file="Check_Sql.asp"-->
<!-- #include file="inc/lib.asp"-->
<%OpenData()%>
<%set rs=server.CreateObject("adodb.recordset")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>杭州茉莉花大酒店</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<link href="css/css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="810" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9"><img  src="images/new1_01.jpg" width="9" height="152"></td>
        <td width="631" background="images/new1_02.jpg">&nbsp;</td>
        <td width="170"><img  src="images/new1_04.jpg" width="170" height="152" border="0" usemap="#Map" ></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td bgcolor="#FFF9D7"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr valign="top">
        <td width="640">
		<%sql="select title,content,newsdate,clicks,writer from sbe_news where id="&request("id")&""
		rs.open sql,conn,1,3
		if not rs.eof then
		rs(3)=rs(3)+1
		rs.update%>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" >
          <tr>
            <td height="100" align="center" valign="bottom"><strong class="notice"><%=rs(0)%></strong></td>
          </tr>
          <tr>
            <td height="50" align="right" valign="bottom"><span class="ziti1">来源:<%=rs(4)%> 浏览量:<%=rs(3)%> 发布时间:<%=rs(2)%></span></td>
          </tr>
          <tr>
            <td valign="top" style="padding-bottom:10px; padding-top:10px">
			<div style="padding-left:12px;  padding-right:12px;l width:100%;overflow:auto;height:270;">
			<table height="200" width="100%"><tr><td align="left" valign="top">
			<%=rs(1)%>
			</td></tr></table>
			</div>
			            </td>
          </tr>
          
        </table>
		<%end if
		rs.close%>
		</td>
        <td width="69"><img id="new1_06" src="images/new1_06.jpg" width="69" height="115" alt="" /></td>
        <td width="101">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><img id="new1_09" src="images/new1_09.jpg" width="810" height="7" alt="" /></td>
  </tr>
  <tr>
    <td height="90" valign="top" bgcolor="#F5E9C3" class="ziti7">
	<%
	if request("type")<>"" then
	sql="select id,title,sequence from sbe_news where tid="&request("type")&" and sequence>"&request("sequence")&"  and show=-1"
	else
	sql="select id,title,sequence from sbe_news where tid=1 and sequence>"&request("sequence")&"  and show=-1"
	end if
	rs.open sql,conn,1,1
	if not rs.eof then
	%>
	上一篇：<a href="?id=<%=rs(0)%>&sequence=<%=rs(2)%>&type=<%=request("type")%>"><%=rs(1)%></a>
	<%end if
	rs.close%>
	
	<%
	if request("type")<>"" then
	sql="select top 1 id,title,sequence from sbe_news where tid="&request("type")&" and sequence<"&request("sequence")&" and show=-1"
	else
	sql="select top 1 id,title,sequence from sbe_news where tid=1 and sequence<"&request("sequence")&"  and show=-1"
	end if
	rs.open sql,conn,1,1
	if not rs.eof then
	%>
	下一篇：<a href="?id=<%=rs(0)%>&sequence=<%=rs(2)%>&type=<%=request("type")%>"><%=rs(1)%></a>
	<%end if
	rs.close%>
	</td>
  </tr>
</table>

<map name="Map"><area shape="rect" coords="144,8,169,25" href="#" onClick="javascript:window.close();"></map></body>
</html>
