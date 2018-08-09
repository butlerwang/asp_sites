<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%

Dim KSCls
Set KSCls = New RefreshHtml
KSCls.Kesion()
Set KSCls = Nothing

Class RefreshHtml
        Private KS,ChannelID, ChannelStr
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		ChannelID = KS.G("ChannelID")		
		ChannelStr =KS.C_S(ChannelID,3)
        Select Case KS.S("Action")
		 Case "ref"
		   Call  refreshlist
		 Case Else
		   Call Main
		End Select
	 End Sub
	 
	 Sub Main
		With Response
		.Write "<html>"
		.Write "<head>"
		.Write "<link href=""Admin_Style.css"" rel=""stylesheet"">"
		.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		.Write "<title>生成内容页管理</title>"
		%>
		 <style type="text/css">
		 #mt{}
		 #mt li{border:1px #a7a7a7 dashed;padding-top:3px;height:20px;margin:2px;}
		 </style>
		<%
		.Write "</head>"
		.Write "<body topmargin=""0"" leftmargin=""0"" scroll='no'>"
		.Write "<table border='0' height='100%' width='100%' cellspacing='0' cellpadding='0'>"
		.Write "<tr>"
		.Write "<td>"
        .Write "<ul id='mt'>"
		.Write " <div id='mtl'>发布选项：</div>"
		.Write " <a href='refreshindex.asp' target='main'>发布首页</a>&nbsp;|&nbsp;<a href='refreshspecial.asp' target='main'>发布专题</a>&nbsp;|&nbsp;<a href='refreshjs.asp' target='main'>发布JS</a>&nbsp;|&nbsp;<a href='refreshcommonpage.asp' target='main'>自定义页面</a>&nbsp;|&nbsp;<a href='Refresh_Sitemap.asp' target='main' title='生成Google地图'>Google/Baidu</a>"

		.Write "</ul>"
		.Write "</td>"
		.Write "</tr>"
		.Write "<tr>"
		.Write " <td height='100%'>"
		.Write " <iframe name=""main"" id='main' scrolling=""auto"" frameborder=""0"" src=""RefreshHtml.asp?Action=ref&channelid=" & ChannelID & """ width=""100%"" height=""100%""></iframe>"
		.Write "</td>"
		.Write "</tr>"
		.Write "</table>"
	  End With
	End Sub
	
	Sub refreshlist()
		With Response
		.Write "<html>"
		.Write "<head>"
		.Write "<link href=""Admin_Style.css"" rel=""stylesheet"">"
		.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		.Write "<title>生成内容页管理</title>"
		.Write "</head>"
		.Write "<script language=""JavaScript"" src=""../../ks_inc/Common.js""></script>"
		.Write "<script language=""JavaScript"" src=""../../ks_inc/popcalendar.js""></script>"
		.Write "<body topmargin=""0"" leftmargin=""0"">"
		.Write "<table width='100%'>"
		.Write "<tr>"
		.Write " <td width='180' valign='top' style='border:1px solid #cccccc'  class='tdbg' align='center'><div style='margin:6px'><strong>请选择要发布的模型</strong></div>"
		.Write "<select name='schannelid' style='width:180px;height:550px' size='2' onchange=""if (this.value!=''){location.href='?action=ref&channelid='+this.value;}"">"
		 Dim RS:Set RS=KS.InitialObject("ADODB.RECORDSET")
		 RS.Open "Select ChannelID,ChannelName From KS_Channel Where ChannelStatus=1 and channelid<>9 and channelid<>6 and channelid<>10 order by channelid",conn,1,1
		 do while not RS.Eof
				If trim(ChannelID)=trim(rs(0)) Then
				.Write "<option value='" & RS(0) & "' selected>" & RS(1) & "</option>"
				Else
				.Write "<option value='" & RS(0) & "'>" & RS(1) & "</option>"
				End If
		  RS.MoveNext
		 Loop
		 RS.Close:Set RS=Nothing
		 .Write "</select>"
		 
		.Write "</td>"
		.Write " <td style='border:1px solid #cccccc'>"
		.Write " <table width=""100%"" style='margin-top:2px'  border=""0"" cellpadding=""0"" align=""center"" cellspacing=""1"">"
		.Write "       <tr class='sort'>"
		.Write "          <td colspan=2>发布" & ChannelStr & "内容页操作</td>"
		.Write "      <tr>"
		.Write "  <form action=""RefreshHtmlSave.asp?Types=Content&RefreshFlag=New&ChannelID=" & ChannelID & """ method=""post"" name=""ArticleNewForm"" onSubmit=""return(CheckTotalNumber())"">"
		.Write "    <tr>"
		.Write "      <td height=""35"" align=""center""  class='tdbg'> 发布最新添加的</td>"
		.Write "      <td width=""78%"" height=""35""> <input name=""TotalNum"" onBlur=""CheckNumber(this,'" & ChannelStr & "');"" type=""text"" id=""TotalNum"" style=""width:20%"" value=""50"">"
		.Write "        " & KS.C_S(ChannelID,4) & ChannelStr
		.Write "        <input name=""Submit2"" type=""submit"" id=""Submit2"" class=""button"" value="" 发 布 &gt;&gt;"" border=""0"">"
		.Write "      </td>"
		.Write "    </tr>"
		.Write "  </form>"
		.Write "  <form action=""RefreshHtmlSave.asp?Types=Content&RefreshFlag=InfoID&ChannelID=" & ChannelID & """ method=""post"" name=""IDForm"">"
	  .Write "    <tr>"
	  .Write "      <td height=""35"" align=""center""  class='tdbg'>按" & ChannelStr & "ID发布</td>"
	  .Write "      <td height=""35""> 从"
	  .Write "        <input name=""StartID"" type=""text"" value=""1"" id=""StartID"">"
	  .Write "        到"
	  .Write "        <input name=""EndID"" type=""text"" value=""100"" id=""EndID"">"
	  .Write "        <input name=""SubmitID"" class=""button"" type=""submit"" id=""SubmitID"" value="" 发 布 &gt;&gt;"" border=""0"">"
	  .Write "      </td>"
	  .Write "    </tr>"
	  .Write "  </form>"
		.Write "  <form action=""RefreshHtmlSave.asp?Types=Content&RefreshFlag=Date&ChannelID=" & ChannelID & """ method=""post"" name=""DateForm"">"
		.Write "    <tr>"
		.Write "      <td height=""35"" align=""center""  class='tdbg'>按日期发布</td>"
		.Write "      <td height=""35""> 从"
		.Write "        <input name=""StartDate"" onclick=""popUpCalendar(this, this, dateFormat,-1,-1)"" type=""text"" id=""StartDate"" readonly style=""width:20%"" value=""" & Date & """>"
		.Write "        <b><a href=""javascript:;"" onclick=""popUpCalendar(document.DateForm.StartDate, document.DateForm.StartDate, dateFormat,-1,-1)""><img src=""../Images/date.gif"" border=""0"" align=""absmiddle"" title=""选择日期""></a></b>"
		.Write "        到"
		.Write "        <input name=""EndDate"" onclick=""popUpCalendar(this, this, dateFormat,-1,-1)"" type=""text"" id=""EndDate"" readonly style=""width:20%"" value=""" & Date & """>"
		.Write "        <b><a href=""javascript:;"" onclick=""popUpCalendar(document.DateForm.EndDate, document.DateForm.EndDate, dateFormat,-1,-1)""><img src=""../Images/date.gif"" border=""0"" align=""absmiddle"" title=""选择日期""></a></b>的" & ChannelStr
		.Write "        <input name=""Submit23"" type=""submit"" class=""button"" id=""Submit23"" value="" 发 布 &gt;&gt;"" border=""0"">"
		.Write "      </td>"
		.Write "    </tr>"
		.Write "  </form>"
		.Write "  <form action=""RefreshHtmlSave.asp?Types=Content&RefreshFlag=Folder&ChannelID=" & ChannelID & """ onSubmit=""return(CheckForm(this))"" method=""post"" name=""ClassForm"">"
		.Write "    <tr>"
		.Write "      <td height=""50"" align=""center""  class='tdbg'> 按" & ChannelStr & "栏目发布</td>"
		.Write "      <td height=""50""> <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		.Write "          <tr>"
		.Write "            <td width=""39%"">"
		.Write "            <input type=""hidden"" name=""FolderID"">"
		.Write "            <select name=""TempFolderID"" size=10 multiple id=""TempFolderID"" style=""width:260"">"
		.Write KS.LoadClassOption(ChannelID,false)
		.Write "              </select></td>"
		.Write "            <td width=""61%""><input type='radio' value='1' name='refreshtf' checked>仅发布未生成过Html的" & ChannelStr & "<br> <input type='radio' value='0' name='refreshtf'>发布所有页面<br>&nbsp;限制最新添加的<input type='text' name='TotalNum' value='50' size='4' style='text-align:center'>篇文档<br><input  class=""button"" name=""Submit22"" type=""submit"" id=""Submit222"" value="" 发布选中栏目的" & ChannelStr & " &gt;&gt;"" border=""0"">"
		.Write "              <br> <font color=""#FF0000""> 　<br>"
		.Write "              　提示：<br>"
		.Write "              　按住""CTRL""或""Shift""键可以进行多选</font></td>"
		.Write "          </tr>"
		.Write "        </table></td>"
		.Write "    </tr>"
		.Write "  </form>"
		.Write "  <form action=""RefreshHtmlSave.asp?Types=Content&RefreshFlag=All&ChannelID=" & ChannelID & """ method=""post"" name=""AllForm"">"
		.Write "    <tr>"
		.Write "      <td height=""30"" align=""center""  class='tdbg'> 发布所有" & ChannelStr & "页面</td>"
		.Write "      <td height=""30"">"
		.Write "        <input type='radio' value='1' name='refreshtf' checked>仅发布未生成过Html的" & ChannelStr & " <input type='radio' value='0' name='refreshtf'>发布所有页面"
		.Write "        <input name=""SubmitAll"" type=""submit"" class=""button"" value=""发布 &gt;&gt;"" border=""0"">"
		.Write "      </td>"
		.Write "    </tr>"
		.Write "  </form>"
		.Write "</table>"
		
		.Write "<table width=""100%"" style='margin-top:2px'  border=""0"" cellpadding=""0"" cellspacing=""1"" align='center'>"
		.Write "  <tr class='sort'>"
		.Write "     <td colspan=2>发布" & ChannelStr & "栏目(频道)操作</td>"
		.Write "   </tr>"		
		.Write "   <tr>"	
		.Write "  <Form action=""RefreshHtmlSave.asp?Types=Folder&RefreshFlag=All&ChannelID=" & ChannelID & """ method=""post"" name=""FolderAllForm"">"
		.Write "    <tr>"
		.Write "      <td height=""30"" align=""center""  class='tdbg'>发布全部栏目</td>"
		.Write "      <td>"
		.Write "<table><tr><td><input type='radio' value='1' name='fsotype'>更新所有列表分页(<font color=blue>较占用资源</font>)<br>"
		.Write "<input type='radio' value='2' name='fsotype' checked>仅发布每个列表页的前<input type='text' name='FsoListNum' value='" & KS.C_S(ChannelID,35) & "' size='6' style='text-align:center'>页"

		.Write " </td><td><input class=""button"" name=""Submit2222"" type=""submit"" id=""Submit2222"" value="" 发布全部栏目(频道) &gt;&gt;"" border=""0""></td></tr></table></td>"
		.Write "    </tr>"
		.Write "  </Form>"
		.Write "  <form action=""RefreshHtmlSave.asp?Types=Folder&RefreshFlag=Folder&ChannelID=" & ChannelID & """ method=""post"" onSubmit=""return(CheckForm(this))"" name=""FolderForm"">"
		.Write "    <tr>"
		.Write "      <td align=""center"" class='tdbg'> 栏目(频道）发布</td>"
		.Write "      <td width=""78%"" height=""50""> <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		.Write "          <tr>"
		.Write "            <td width=""39%"">"
		.Write "             <input type=""hidden"" name=""FolderID"">"
		.Write "             <select name=""TempFolderID"" size=12 multiple id=""TempFolderID"" style=""width:260px"">"
		Dim Node,K,SQL,NodeText,Pstr,TJ,SpaceStr,TreeStr
		KS.LoadClassConfig()
		If KS.ChkClng(ChannelID)<>0 Then Pstr="[@ks12=" & channelid & "]"
		For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class" & Pstr&"")
		  SpaceStr=""
		  TJ=Node.SelectSingleNode("@ks10").text
		  If TJ>1 Then
			 For k = 1 To TJ - 1
				 SpaceStr = SpaceStr & "──"
			 Next
			.Write "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & SpaceStr & Node.SelectSingleNode("@ks1").text & " </option>"
		  Else
		    .Write "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & Node.SelectSingleNode("@ks1").text & " </option>"
		  End If
		  
		Next
		
		.Write "              </select></td>"
		.Write "            <td width=""61%"">"
		.Write "<input type='radio' value='1' name='fsotype'>更新所有列表分页(<font color=blue>较占用资源</font>)<br>"
		.Write "<input type='radio' value='2' name='fsotype' checked>仅发布每个列表页的前<input type='text' name='FsoListNum' value='" & KS.C_S(ChannelID,35) & "' size='6' style='text-align:center'>页"
		.Write "              <input class=""button"" name=""Submit222"" type=""submit"" id=""Submit223"" value="" 发布选中的栏目 &gt;&gt;"" border=""0"">"
		.Write "              <br> <font color=""#FF0000""> 　<br>"
		.Write "              　提示：<br>"
		.Write "              　按住""CTRL""或""Shift""键可以进行多选</font></td>"
		.Write "          </tr>"
		.Write "        </table></td>"
		.Write "    </tr>"
		.Write "  </Form>"
		.Write "</table>"
		.Write "</td>"
		.Write "</tr>"
		.Write "</table>"
		.Write "<br><div align='center'><font color=#ff6600>友情提示：发布操作会比较占用系统资源及时间，每次发布时请尽量仅发布最新添加的信息</font></div>"
		.Write ""
		.Write "</body>"
		.Write "</html>"
		.Write "<script>" & vbCrLf
		.Write " function CheckForm(FormObj)" & vbCrLf
		.Write " {var tempstr='';" & vbCrLf
		.Write " for (var i=0;i<FormObj.TempFolderID.length;i++){" & vbCrLf
		.Write "     var KM = FormObj.TempFolderID[i];" & vbCrLf
		.Write "    if (KM.selected==true)" & vbCrLf
		.Write "       tempstr = tempstr + "" '"" + KM.value + ""',""" & vbCrLf
		.Write "    }" & vbCrLf
		.Write "    if (tempstr=='')" & vbCrLf
		.Write "    {" & vbCrLf
		.Write "    alert('请选择您要发布的(栏目)频道!');" & vbCrLf
		.Write "    return false;" & vbCrLf
		.Write "    }" & vbCrLf
		.Write "    FormObj.FolderID.value=tempstr.substr(0,(tempstr.length-1));" & vbCrLf
		.Write "  return true;" & vbCrLf
		.Write " }" & vbCrLf
		.Write "function CheckTotalNumber()"
		.Write "{"
		.Write "    if (document.ArticleNewForm.TotalNum.value=='') {alert('请填写新闻数量');document.ArticleNewForm.TotalNum.focus();return false;}"
		.Write "    else return true;"
		.Write "}"
		.Write "</script>"
		End With
		End Sub
		
	
End Class
%> 
