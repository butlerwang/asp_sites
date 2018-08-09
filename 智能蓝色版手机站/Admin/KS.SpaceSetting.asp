<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_System
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_System
        Private KS,KSMCls
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSMCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call KS.DelCahe(KS.SiteSn & "_Config")
		 Call KS.DelCahe(KS.SiteSn & "_Date")
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		       Call SetSystem()
		End Sub
	
		'系统基本信息设置
		Sub SetSystem()
		Dim SqlStr, RS, InstallDir, FsoIndexFile, FsoIndexExt
		Dim SetType
		SetType = KS.G("SetType")
		With Response
			If Not KS.ReturnPowerResult(0, "KSMS10000") Then          '检查是否有基本信息设置的权限
			  Call KS.ReturnErr(1, "")
			 .End
			End If
	
			SqlStr = "select SpaceSetting from KS_Config"
			Set RS = Server.CreateObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1, 3
			Dim Setting:Setting=Split(RS(0)&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","^%^")
			If KS.G("Flag") = "Edit" Then
			    Dim N					
			    Dim WebSetting
				For n=0 To 60
				  if n=56 then
				   WebSetting=WebSetting & KS.G("SynchOption") & "^%^"
				  else
				   WebSetting=WebSetting & Replace(KS.G("Setting(" & n &")"),"^%^","") & "^%^"
				  end if
				Next
				RS("SpaceSetting")=WebSetting
				RS.Update				
				.Write ("<script>alert('空间参数修改成功！');location.href='KS.SpaceSetting.asp';</script>")
			End If
			
		  	.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
			.Write "<title>空间参数设置</title>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			.Write "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			.Write "<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>"
			.Write "<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "</head>" & vbCrLf

			.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.Write "<div class='topdashed sort'>空间参数配置</div>"
			.Write "<br>"
			.Write "<div class=tab-page id=spaceconfig>"
			.Write "  <form name='myform' id='myform' method=post action="""" onSubmit=""return(CheckForm())"">"
            .Write " <SCRIPT type=text/javascript>"
            .Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""spaceconfig"" ), 1 )"
            .Write " </SCRIPT>"
             
			.Write " <div class=tab-page id=site-page>"
			.Write "  <H2 class=tab>空间配置</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "<input type=""hidden"" value=""Edit"" name=""Flag"">"
			.Write "<input type=""hidden"" name=""SynchOption"" id=""SynchOption"" value=""" & Setting(56) & "0000000000000000000000000""/>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""20%"" height=""30"" class=""clefttitle"" align=""right""><div><strong>空间状态：</strong></div></td>"
			.Write "      <td width=""50%"" height=""30"">" 
			
				.Write " <input type=""radio"" name=""Setting(0)"" value=""1"" "
				If Setting(0) = "1" Then .Write (" checked")
				.Write "> 打开"
				.Write "    <input type=""radio"" name=""Setting(0)"" value=""0"" "
				If Setting(0) = "0" Then .Write (" checked")
				.Write "> 关闭"

			
			.Write "     </td><td class=""tips"">如果选择“关闭”那么前台注册会员将无法使用空间站点功能。</td>"
			.Write "    </tr>"
			
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>运行模式：</strong></div></td>"
			.Write "      <td>" 
			
				.Write " <input type=""radio"" name=""Setting(21)"" onclick=""$('#ext').hide();"" value=""0"" "
				If Setting(21) = "0" Then .Write (" checked")
				.Write "> 动态模式"
				.Write "    <input type=""radio"" name=""Setting(21)"" onclick=""$('#ext').show();"" value=""1"" "
				If Setting(21) = "1" Then .Write (" checked")
				.Write "> 伪静态"

             If Setting(21)="1" Then
			  .Write "<div id='ext'>"
			 Else
			  .Write "<div id='ext' style='display:none'>"
			 End If
			.Write "伪静态扩展名:<input class='textbox' type='text' size='8' name='Setting(22)' value='" & Setting(22) & "'>,更改此配置,需要修改ISAPI_Rewrite的配置文件httpd.ini</div>"
			.Write "    </td><td class=""tips"">选择伪静态功能需要服务器安装ISAPI_Rewrite组件。</td>"
			.Write "    </tr>"

			
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td class=""clefttitle"" disabled align=""right""><div><strong>是否启用二级域名：</strong></div></td>"
			.Write "      <td height=""30"">" 
			
				.Write "    <label><input disabled type=""radio"" name=""Setting(14)"" value=""0"" "
				If Setting(14) = "0" Then .Write (" checked")
				.Write "> 否(不支持)</label><br/>"

				.Write " <label><input disabled type=""radio"" name=""Setting(14)"" value=""1"" "
				If Setting(14) = "1" Then .Write (" checked")
				.Write "> 仅允许绑定本站的二级域名</label><br/>"
				.Write " <label><input disabled type=""radio"" name=""Setting(14)"" value=""2"" "
				If Setting(14) = "2" Then .Write (" checked")
				.Write "> 允许绑定本站二级域名和独立域名(<span style='color:green'>独立域名需解释到我的服务器</span>)</label>"
			
			.Write "     </td><td disabled class=""tips"">此功能必须自己有独立服务器或是您的空间支持泛域名解释,仅对商业版本开放，若不支持请选择否。</td>"
			.Write "    </tr>"
			
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>空间首页域名：</strong></div></td>"
			 .Write "  <td height=""30"">"			 .Write " <input type=""text"" class='textbox' name=""Setting(15)"" size=15 value=""" & Setting(15) & """>"
			 .Write "    </td><td class=""tips"">此项功能需要开启二级域名才生效</td>"
			 .Write "</tr>"
			 
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>空间站点二级域名：</strong></div></td>"
			 .Write "  <td height=""30"">"
			 .Write " <input  type=""text"" class='textbox' name=""Setting(16)"" size=15 value=""" & Setting(16) & """> 			 .Write "    </td><td class=""tips"">关闭二级域名功能请留空,若设置为三级域名则用户站点访问,若设置二级域名则用户站点访问</td>"
			 .Write "</tr>"
			
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle""align=""right""><div><strong>会员注册是否自动注册个人空间：</strong></div></td>"
			 .Write "     <td height=""30"">"
			 	.Write " <input type=""radio"" name=""Setting(1)"" value=""1"" "
				If Setting(1) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(1)"" value=""0"" "
				If Setting(1) = "0" Then .Write (" checked")
				.Write "> 否"

			 .Write "</td><td class=""tips"">如果选择“是”那么注册会员的同时将同时拥有一个个人空间站点。</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			  .Write "    <td height=""30"" class=""clefttitle""> <div align='right'><strong>申请空间是否需要审核：</strong></div></td>"
			 .Write "    <td height=""30""> <input type='radio' name='Setting(2)' value='0'"
			 If Setting(2) = "0" Then .Write (" checked")
			 .Write "/>不需要审核<br/><input type='radio' name='Setting(2)' value='1'"
			 If Setting(2) = "1" Then .Write (" checked")
			 .Write "/>申请个人空间不需要审核，申请企业空间要审核<br/> <input type='radio' name='Setting(2)' value='2'"
			 If Setting(2) = "2" Then .Write (" checked")
			 .Write "/>申请个人和企业空间都需要审核<br/>"
			 
			 .Write "      </td><td class=""tips"">建议设置为“申请个人空间不需要审核，申请企业空间要审核”</td>"
			 .Write "   </tr>"
 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>副模板更多每页显示设置：</strong></div></td>"
			 .Write "  <td colspan=""2"">"
			 .Write " 空间每页显示<input type=""text"" name=""Setting(9)"" class='textbox' style=""text-align:center"" size=5 value=""" & Setting(9) & """> 个 日志每页显示<input type=""text"" name=""Setting(10)"" class='textbox' style=""text-align:center"" size=5 value=""" & Setting(10) & """> 篇 圈子每页显示<input type=""text"" name=""Setting(11)"" class='textbox' style=""text-align:center"" size=5 value=""" & Setting(11) & """> 个"
			 .Write "    </td>"
			 .Write "</tr>"


			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>副模板更多相册每页显示：</strong></div></td>"
			 .Write "  <td>"
			 .Write " <input type=""text"" name=""Setting(12)"" class='textbox' style=""text-align:center"" size=5 value=""" & Setting(12) & """> 本相册 每行显示<input type=""text"" name=""Setting(13)"" class='textbox' style=""text-align:center"" size=5 value=""" & Setting(13) & """> 本"
			 .Write "    </td><td class=""tips""></td>"
			 .Write "</tr>"			 
			 
			 .Write "</table>"
			 .Write "</div>"
			 
			 
			.Write " <div class=tab-page id=post-page>"
			 .Write "  <H2 class=tab>发表设置</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""post-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"

            .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align='right'> <div><strong>积分限制：</strong></div></td>"
			 .Write "     <td>发表日志要求达到<input type='text' style='width:40px;text-align:center' name='Setting(36)' value='" & Setting(36) &"' class='textbox'/>分积分 上传照片要求达到<input type='text' style='width:40px;text-align:center' name='Setting(37)' value='" & Setting(37) &"' class='textbox'/>分积分 创建圈子要求达到积分<input type='text' style='width:40px;text-align:center' name='Setting(38)' value='" & Setting(38) &"' class='textbox'/>分积分  发布企业新闻要求达到积分<input type='text' style='width:40px;text-align:center' name='Setting(39)' value='" & Setting(39) &"' class='textbox'/>分积分  上传企业荣誉证书要求达到积分<input type='text' style='width:40px;text-align:center' name='Setting(40)' value='" & Setting(40) &"' class='textbox'/>分积分 添加音乐要求达到积分<input type='text' style='width:40px;text-align:center' name='Setting(41)' value='" & Setting(41) &"' class='textbox'/>分积分"

			 .Write "    </td><td class=""tips"">可以有效启用防止发帖机发布作用，不限制，请输入“0”。</td>"
			 .Write "   </tr>"
			 

			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align='right'> <div><strong>发表日志是否需要审核：</strong></div></td>"
			 .Write "     <td>"
			 	
				.Write " <input type=""radio"" name=""Setting(3)"" value=""1"" "
				If Setting(3) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(3)"" value=""0"" "
				If Setting(3) = "0" Then .Write (" checked")
				.Write "> 否"

			 .Write "       </td><td class=""tips""></td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align='right'> <div><strong>发表日志是否允许上传附件：</strong></div></td>"
			 .Write "     <td>"
			 	
				.Write " <input type=""radio"" onclick=""$('#fj').show();"" name=""Setting(26)"" value=""1"" "
				If Setting(26) = "1" Then .Write (" checked")
				.Write "> 允许"
				.Write "    <input type=""radio"" onclick=""$('#fj').hide();"" name=""Setting(26)"" value=""0"" "
				If Setting(26) = "0" Then .Write (" checked")
				.Write "> 不允许"
				If Setting(26) = "1" Then
                .Write "<div id='fj' style='color:blue'>"
				Else
                .Write "<div id='fj' style='display:none;color:blue'>"
				End If
				.Write "允许上传的附件扩展名:<input class='textbox' type='text' value='" & Setting(27) & "' name='Setting(27)' /> 多个扩展名用 |隔开,如gif|jpg|rar等<Br/>允许上传的文件大小：<input class='textbox' name=""Setting(28)"" type=""text"" value=""" & Setting(28) &""" style=""text-align:center"" size='8'>KB<br/>每天上传文件个数：<input class='textbox' name=""Setting(29)"" type=""text"" value=""" & Setting(29) &""" style=""text-align:center"" size='8'>个,不限制请填0</font><br/>"
			 .Write "<br/><strong>允许上传附件的用户组:</strong>"
			 .Write KS.GetUserGroup_CheckBox("Setting(30)",Setting(30),5)
			 .Write "</div>"
			 .Write "       </td><td class=""tips"">如果允许所有会员组上传附件，请不要勾选用户组</td>"
			 .Write "   </tr>"
			 
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>发表日志远程图片是否自动保存到本地：</strong></div></td>"
			  .Write "    <td>"
			  
			  	.Write " <input type=""radio"" name=""Setting(35)"" value=""1"" "
				If Setting(35) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(35)"" value=""0"" "
				If Setting(35) = "0" Then .Write (" checked")
				.Write "> 否"
			  
			  .Write "    </td><td class=""tips"">选择“是”，则用户转载的日志里含有远程图片将自动保存到您的服务器上</td>"
			 .Write "   </tr>"			 
			 
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>创建相册是否需要审核：</strong></div></td>"
			  .Write "    <td>"
			  
			  	.Write " <input type=""radio"" name=""Setting(4)"" value=""1"" "
				If Setting(4) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(4)"" value=""0"" "
				If Setting(4) = "0" Then .Write (" checked")
				.Write "> 否"
			  
			  .Write "    </td><td class=""tips""></td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>最大允许上传的单张照片：</strong></div></td>"
			  .Write "    <td>"
			  
			  	.Write " <input type=""text"" class='textbox' name=""Setting(32)""  size='5' style='text-align:center' value=""" & Setting(32) &"""> K"
			  
			  .Write "    </td><td class=""tips"">建议不要超过1024K</td>"
			 .Write "   </tr>"
			 
			
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>创建圈子是否需要审核：</strong></div></td>"
				.Write "  <td>"
				
				.Write " <input type=""radio"" name=""Setting(5)"" value=""1"" "
				If Setting(5) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(5)"" value=""0"" "
				If Setting(5) = "0" Then .Write (" checked")
				.Write "> 否"
				
				.Write "    </td><td class=""tips""></td>"
				.Write "</tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>用户留言是否需要审核：</strong></div></td>"
				.Write "  <td height=""30"">"
				
				.Write " <input type=""radio"" name=""Setting(24)"" value=""1"" "
				If Setting(24) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(24)"" value=""0"" "
				If Setting(24) = "0" Then .Write (" checked")
				.Write "> 否"
				
				.Write "    </td><td class=""tips"">启用后,用户的留言只有经过后台管理员审核后,前台的空间才可以看到</td>"
				.Write "</tr>"
				
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>允许游客在空间里评论/留言：</strong></div></td>"
			  .Write "    <td>"
			  
			  	.Write " <input type=""radio"" name=""Setting(25)"" value=""1"" "
				If Setting(25) = "1" Then .Write (" checked")
				.Write "> 允许"
				.Write "    <input type=""radio"" name=""Setting(25)"" value=""0"" "
				If Setting(25) = "0" Then .Write (" checked")
				.Write "> 不允许"
			  
			  .Write "    </td><td class=""tips"">建议设置不允许,可以有效阻止一些注册机留言</td>"
			 .Write "   </tr>"				
				
				

			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>每个会员允许创建圈子数：</strong></div></td>"
				.Write "  <td>"
				.Write " <input type=""text"" name=""Setting(6)"" class='textbox' style=""text-align:center"" size=5 value=""" & Setting(6) & """>个"
				.Write "    </td><td class=""tips"">如果不想限制请输入“0”</td>"
				.Write "</tr>"

				
			

			 .Write " </table>"
			 .Write "</div>"
			 
			 .Write " <div class=tab-page id=weibo-page>"
			.Write "  <H2 class=tab>微博参数设置</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""weibo-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			  .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>频道状态：</strong></div></td><td width='500'>"
			  	.Write " <input type=""radio"" name=""Setting(55)"" value=""1"" "
				If Setting(55) = "1" Then .Write (" checked")
				.Write "> 开通"
				.Write "    <input type=""radio"" name=""Setting(55)"" value=""0"" "
				If Setting(55) = "0" Then .Write (" checked")
				.Write "> 关闭"
              .Write "</td><td class=""tips"">只有开通状态，前台会员才可以发布微博广播。</td></tr>"
			 
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 dim maxlen
			 IF KS.ChkClng(Setting(34))>255 Then MaxLen=255 Else MaxLen=Setting(34)
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right"" width='130'><div><strong>广播字数限制：</strong></div></td><td>最少<input type=""text"" class='textbox' name=""Setting(33)""  size='5' style='text-align:center' value=""" & Setting(33) &"""> 个字符&nbsp;&nbsp;最多<input type=""text"" class='textbox' name=""Setting(34)""  size='5' style='text-align:center' value=""" & maxlen &"""> 个字符</td><td class=""tips"">不能超过255个字符。</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>允许上传图片：</strong></div></td><td width='500'>"
			  	.Write " <input type=""radio"" name=""Setting(50)"" value=""1"" "
				If Setting(50) = "1" Then .Write (" checked")
				.Write "> 允许"
				.Write "    <input type=""radio"" name=""Setting(50)"" value=""0"" "
				If Setting(50) = "0" Then .Write (" checked")
				.Write "> 不允许"
			 .Write "<font color=blue><br/>允许上传的图片文件大小：<input style='text-align:center' type='text' name='Setting(51)' value='" & Setting(51) & "' class='textbox' size='5'/>KB"
			 .Write "<br/>每天上传文件个数限制：<input style='text-align:center' type='text' name='Setting(52)' value='" & Setting(52) & "' class='textbox' size='5'/>个，不限制请输入“0”"
			 .Write "<br/>允许上传图片的用户组："
			 .Write KS.GetUserGroup_CheckBox("Setting(53)",Setting(53),5)
			 .write "</font></td><td class=""tips"">选择允许，在广播里才能上传图片,如果允许所有会员组上传图片，请不要勾选用户组</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>上传总目录：</strong></div></td><td>"
			  	.Write " <input type=""text"" name=""Setting(54)"" value=""" & Setting(54) & """ class=""textbox""></td><td class=""tips"">如WeiboFiles则表示微博里所有上传文件都上传到UploadFiles/WeiboFiles/目录下,后面不要带""/""。</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>广播同步设置：</strong></div></td><td>"
			   Dim Wbtb:Wbtb=Setting(56)&"00000000000000000000000000000000000000"
			  	.Write " <label><input type=""checkbox"" name=""Synch"""
				If mid(wbtb,1,1)="1" then .write " checked"
				.Write " value=""1"">论坛新帖</label>"
			  	.Write " <label><input type=""checkbox"" name=""Synch"""
				If mid(wbtb,2,1)="1" then .write " checked"
				.Write " value=""2"">空间博文</label>"
			  	.Write " <label><input type=""checkbox"" name=""Synch"""
				If mid(wbtb,3,1)="1" then .write " checked"
				.Write " value=""3"">空间照片</label>"
			  	.Write " <label><input type=""checkbox"" name=""Synch"""
				If mid(wbtb,4,1)="1" then .write " checked"
				.Write " value=""4"">空间圈子</label>"
			  	.Write " <label><input type=""checkbox"" name=""Synch"""
				If mid(wbtb,5,1)="1" then .write " checked"
				.Write " value=""5"">模型投稿</label>"
			  	.Write " <br/><label><input type=""checkbox"" name=""Synch"""
				If mid(wbtb,6,1)="1" then .write " checked"
				.Write " value=""6"">更换头像</label>"
			  	.Write " <label><input type=""checkbox"" name=""Synch"""
				If mid(wbtb,7,1)="1" then .write " checked"
				.Write " value=""7"">企业新闻</label>"
			  	.Write " <label><input type=""checkbox"" name=""Synch"""
				If mid(wbtb,8,1)="1" then .write " checked"
				.Write " value=""8"">企业证书</label>"
			  	.Write " <label><input type=""checkbox"" name=""Synch"""
				If mid(wbtb,9,1)="1" then .write " checked"
				.Write " value=""9"">招聘频道</label>"
			 .Write "   </td><td class=""tips"">设置同步的频道，当有会员发布时，自动截取一定介绍到微博广播大厅发布。</td></tr>"

			 .Write " </table>"
			 .Write "</div>"
			 
			 
			 
			.Write " <div class=tab-page id=template-page>"
			.Write "  <H2 class=tab>模板绑定</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""template-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
            .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""150"" height=""30"" class=""clefttitle"" align=""right""><div><strong>空间首页模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(7)"" id='Setting7' type=""text"" value=""" & Setting(7) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting7')[0]") & "</td><td class=""tips"">对应页面<a href='../user/space.asp' target='_blank'>/user/space.asp</a></td>"
			.Write "    </tr>"            
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>空间副模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(8)"" id='Setting8' type=""text"" value=""" & Setting(8) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting8')[0]") & "</td><td class=""tips"">空间的副模板，用于显示更多日志、相册、圈子等，必须包含标签“{$ShowMain}”。</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>交友首页模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(23)"" id='Setting23' type=""text"" value=""" & Setting(23) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting23')[0]") & "</td><td class=""tips"">对应<a href='..//space/friend/index.asp' target='_blank'>/space/friend/index.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>微博首页模板：</strong></div></td>"
			.Write "      <td nowrap> <input class='textbox' name=""Setting(31)"" id='Setting31' type=""text"" value=""" & Setting(31) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting31')[0]") & "</td><td class=""tips"">对应<a href='../user/weibo.asp' target='_blank'>/user/weibo.asp</a> </td>"
			.Write "    </tr>"
			 .Write " </table>"
			.Write " </div>"

			.Write " <div class=tab-page id=user-page>"
			.Write "  <H2 class=tab>企业空间设置</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""user-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
            .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""clefttitle"" align=""right""><div><strong>允许升级为企业空间的用户组：</strong></div><font color=red>不限制,请不要选</font></td>"
			.Write "      <td width=""63%"" height=""30""> &nbsp;" & KS.GetUserGroup_CheckBox("Setting(17)",Setting(17),5) & "</td>"
			.Write "    </tr>"            
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""clefttitle"" align=""right""><div><strong>发布企业新闻是否需要审核：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30"">"
				.Write " <input type=""radio"" name=""Setting(18)"" value=""1"" "
				If Setting(18) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(18)"" value=""0"" "
				If Setting(18) = "0" Then .Write (" checked")
				.Write "> 否"
			.Write "</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""clefttitle"" align=""right""><div><strong>发布企业产品是否需要审核：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30"">"
				.Write " <input type=""radio"" name=""Setting(19)"" value=""1"" "
				If Setting(19) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(19)"" value=""0"" "
				If Setting(19) = "0" Then .Write (" checked")
				.Write "> 否"
			.Write "</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""clefttitle"" align=""right""><div><strong>发布荣誉证书是否需要审核：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30"">"
				.Write " <input type=""radio"" name=""Setting(20)"" value=""1"" "
				If Setting(20) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(20)"" value=""0"" "
				If Setting(20) = "0" Then .Write (" checked")
				.Write "> 否"
			.Write "</td>"
			.Write "    </tr>"
			.Write " </table>"
			.Write " </div>"
			

			.Write ""

			.Write " </body>"
			.Write " </html>"
			.Write " <Script Language=""javascript"">"
			.Write " <!--" & vbCrLf
			
			.Write " function CheckForm()" & vbCrLf
			.Write " {" & vbCrLf
			.Write "if ($('#Setting7').val()=='')" & vbCrLf
			.Write "{ alert('请选择空间首页模板!');" & vbCrLf
			.Write "  $('#Setting7').focus();" & vbCrLf
			.Write "  return false;" & vbCrLf
			.Write "}" & vbCrLf
			.Write "if ($('#Setting8').val()=='')" & vbCrLf
			.Write "{ alert('请选择空间副模板!');" & vbCrLf
			.Write "  $('#Setting8').focus();" & vbCrLf
			.Write "  return false;" & vbCrLf
			.Write "}" & vbCrLf
			.Write "var Synch='';" &vbcrlf
			.Write " $(""input[name=Synch]"").each(function(){ " &vbcrlf
			.Write "     if ($(this).attr(""checked"")==true){" &vbcrlf
			.Write "	  Synch=Synch+'1'}else{Synch=Synch+'0'}" &vbcrlf
			.Write " })" &vbcrlf
			.Write " $(""#SynchOption"").val(Synch);" &vbcrlf
			.Write " $('#myform').submit();" & vbCrLf
			.Write " }" & vbCrLf
			.Write " //-->" & vbCrLf
			.Write " </Script>" & vbCrLf
			RS.Close:Set RS = Nothing:Set Conn = Nothing
		End With
		End Sub
	
		

End Class
%> 
