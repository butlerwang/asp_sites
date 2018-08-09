<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_BlogMusic
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_BlogMusic
        Private KS
		Private Action,i,strClass,sFileName,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		With Response
					If Not KS.ReturnPowerResult(0, "KSMS10005") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
	    if ks.s("action")="play" then
		 call MusicPlay()
		 response.end
		end if
		.Write "<script src='../KS_Inc/common.js'></script>"
		.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		.Write "<ul id='mt'>"
		.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		.Write "  <tr>"
		.Write "    <td height=""23"" align=""left"" valign=""top"">"
		.Write "	<td align=""center""><strong>用 户 歌 曲 管 理</strong></td>"
		.Write "    </td>"
		.Write "  </tr>"
		.Write "</table>"
		.Write "</ul>"
		End With	
		
		
		maxperpage = 30 '###每页显示数
		If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
			Response.Write ("错误的系统参数!请输入整数")
			Response.End
		End If
		If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
			CurrentPage = CInt(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CInt(CurrentPage) = 0 Then CurrentPage = 1
		totalPut = Conn.Execute("Select Count(ID) from KS_BlogMusic")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
        
		if ks.s("action")="Del" then
		 call SongDel
		else
		  Call showmain
		end if
End Sub

Private Sub showmain()
%>		<script src="../ks_inc/kesion.box.js" language="JavaScript"></script>
<script>
	   function play(s,t)
	   {
			onscrolls=false;
            new KesionPopup().PopupCenterIframe("歌曲试听","?action=play&songname="+t+"&url="+s,550,150,'no')

	   }
		</script>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</td>
	<td width="29%" nowrap>歌曲名称</td>
	<td width="11%" nowrap>上传用户</td>
	<td width="11%" nowrap>歌 手</td>
	<td width="16%" nowrap>上传时间</t>
  <td width="18%" nowrap>管理操作</td>
</tr>
<%
	sFileName = "KS.SpaceMessage.asp?"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_BlogMusic order by id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr class='list'><td height=""25"" align=center colspan=7>没有人发表歌曲！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action=?action=Del>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="22" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=Rs("id")%>'></td>
	<td class="splittd"><a href="#"><%=Rs("songname")%></a></td>
	<td class="splittd" align="center"><%=rs("username")%></td>
	<td class="splittd" align="center"><%=Rs("singer")%></td>
	<td class="splittd" align="center"><%=Rs("adddate")%></td>
	<td class="splittd" align="center"><a href="#" onClick="play('<%=rs("url")%>','<%=rs("songname")%>')"><img src="../user/images/radio.gif" align="absmiddle" border="0">试听</a> <a href="?Action=Del&ID=<%=RS("ID")%>" onClick="return(confirm('确定删除该歌曲吗？'));">删除</a> </td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" >
	<td class="splittd" height='25' colspan=8>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class=Button type="submit" name="Submit2" value=" 删除选中的歌曲 " onClick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){this.document.selform.submit();return true;}return false;}"></td>
</tr>

</form>
<tr>
	<td colspan=8>
	<%
	  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>

<%
End Sub

		Sub MusicPlay()
		 Response.Expires = -1 
		Response.ExpiresAbsolute = Now() - 1 
		Response.cachecontrol = "no-cache" 
		dim url:url=request("url")
		 %>
			<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
			<title>用户管理中心</title>
			<link href="../user/images/css.css" type="text/css" rel="stylesheet" />
			<META HTTP-EQUIV="pragma" CONTENT="no-cache"> 
			<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache, must-revalidate"> 
			<META HTTP-EQUIV="expires" CONTENT="Wed, 26 Feb 1997 08:21:57 GMT">
			<style>
			 .tt{font-size:14px;color:#191970}
			 .tt span{font-size:12px;color:#999999}
			</style>
			</head>
			<body leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">
			<br>
			<table  width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
			  <tr class="tdbg">
                 
                 <td height="25" align="center" class="tt"> 
				 
				  <object id="MediaPlayer1" width="350" height="64" classid="CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6" 
codebase="http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=6,4,7,1112"
align="baseline" border="0" standby="Loading Microsoft Windows Media Player components..." 
type="application/x-oleobject">
    <param name="URL" value="<%=url%>">
    <param name="autoStart" value="true">
    <param name="invokeURLs" value="false">
    <param name="playCount" value="100">
    <param name="defaultFrame" value="datawindow">
       
		<embed src="<%=url%>" align="baseline" border="0" width="350" height="68"
			type="application/x-mplayer2"
			pluginspage=""
			name="MediaPlayer1" showcontrols="1" showpositioncontrols="0"
			showaudiocontrols="1" showtracker="1" showdisplay="0"
			showstatusbar="1"
			autosize="0"
			showgotobar="0" showcaptioning="0" autostart="1" autorewind="0"
			animationatstart="0" transparentatstart="0" allowscan="1"
			enablecontextmenu="1" clicktoplay="0" 
			defaultframe="datawindow" invokeurls="0">
		</embed>
</object>
				
				<!--<EMBED style="WIDTH: 272px; HEIGHT: 29px" src=<%=url%> width=299 height=10 type=audio/x-wav autostart="true" loop="true"></DIV></EMBED>
				-->
                   <!--
				     <object type='application/x-shockwave-flash' height='20' width='200' data='/ks_inc/dewplayer.swf?son=<%=url%>&autoplay=1&autoreplay=1'>
    <param value='/ks_inc/dewplayer.swf?son=<%=url%>&autoplay=1&autoreplay=1'name='movie' />
    <param name="wmode" value="transparent" />
    <param name="bgcolor" value="" />
  </object>-->
				   
				<br><span><%=Request("songname")%></span></td>
              </tr>

			 </table>
	
			 <div style="text-align:center">&nbsp;<input type="button" value="关闭窗口" onClick="parent.closeWindow();" class="button"></div>
			 </form>
		 	</body>
			</html>
		<%
		End Sub

'删除歌曲
Sub SongDel()
		  on error resume next
		  Dim i,id:id=KS.FilterIDs(KS.S("id"))
		  if (id="") then Call KS.AlertHistory("对不起，参数传递出错!",-1):exit sub
		  dim ids:ids=split(id,",")
		  for i=0 to ubound(ids)
		    ks.deletefile(conn.execute("select url from ks_blogmusic where id=" & ids(i))(0))
		  next
		  Conn.Execute("delete from ks_blogmusic where id in(" & id & ")")
		  Conn.Execute("delete from KS_UploadFiles Where ChannelID=1027 and infoid in(" & id & ")")
		  Call KS.AlertHintScript("恭喜，删除成功!")
End Sub


End Class
%> 
