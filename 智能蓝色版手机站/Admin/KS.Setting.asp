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
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		 With Response
		  	.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
			.Write "<title>网站基本参数设置</title>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			.Write "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			.Write "<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>"
			.Write "<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<style type=""text/css"">"
			.Write "<!--" & vbCrLf
			.Write ".STYLE1 {color: #FF0000}" & vbCrLf
			.Write ".STYLE2 {color: #FF6600}" & vbCrLf
			.Write ".tips {color: #999999;padding:2px}" & vbCrLf
			.Write ".txt {color: #666;border:1px solid #ccc;height:22px;line-height:22px}" & vbCrLf
			.Write "textarea {color: #666;border:1px solid #ccc;}" & vbCrLf
			.Write "-->" & vbCrLf
			.Write "</style>" & vbCrLf
			.Write "</head>" & vbCrLf

		  Select Case KS.G("Action")
		  
		   Case "Space"
		     	If Not KS.ReturnPowerResult(0, "KMST10010") Then          
				   Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
				   Call KS.ReturnErr(1, ""): Exit Sub
				Else
		           Call GetSpaceInfo()
				End If
		   Case "CopyRight"
		     	If Not KS.ReturnPowerResult(0, "KMST10011") Then         
				   Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
				   Call KS.ReturnErr(1, ""): Exit Sub
				Else
		           Call GetCopyRightInfo()
				End If
		   Case Else
		       Call SetSystem()
		  End Select
		 End With
		End Sub
	
		'系统基本信息设置
		Sub SetSystem()
		Dim SqlStr, RS, InstallDir, FsoIndexFile, FsoIndexExt
		With Response
			
					If Not KS.ReturnPowerResult(0, "KMST10001") Then          '检查是否有基本信息设置的权限
					 .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back()';</script>")
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
			
			dim strDir,strAdminDir
			strDir=Trim(request.ServerVariables("SCRIPT_NAME"))
			strAdminDir=split(strDir,"/")(Ubound(split(strDir,"/"))-1) & "/"
			InstallDir=left(strDir,instr(lcase(strDir),"/"&Lcase(strAdminDir)))
			
			If Instr(UCASE(InstallDir),"/W3SVC")<>0 Then
			   InstallDir=Left(InstallDir,Instr(InstallDir,"/W3SVC"))
			End If
	
	
			SqlStr = "select * from KS_Config"
			Set RS = KS.InitialObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1, 3
			
			 Dim Setting:Setting=Split(RS("Setting")&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","^%^")
			 Dim TBSetting:TBSetting=Split(RS("TBSetting"),"^%^")
			 FsoIndexFile = Split(Setting(5), ".")(0)
			 FsoIndexExt = Split(Setting(5), ".")(1)
			If KS.G("Flag") = "Edit" Then
			            'IP设置
			            Dim LockIP,I,PartIPArr
						Dim LockIPWhite:LockIPWhite=KS.G("LockIPWhite")
						Dim LockIPBlack:LockIPBlack=KS.G("LockIPBlack")
						If  LockIPWhite<>"" Then
							Dim LockIPWhiteArr:LockIPWhiteArr=Split(LockIPWhite,vbcrlf)
							For I=0 To Ubound(LockIPWhiteArr)
							 If LockIPWhiteArr(i)<>"" and instr(LockIPWhiteArr(i),"----")>0 Then
								 PartIPArr=Split(LockIPWhiteArr(i),"----")
								 If I=0 Then
								   LockIP=LockIP & KS.EncodeIP(PartIPArr(0))&"----" & KS.EncodeIP(PartIPArr(1))
								 Else
								   LockIP=LockIP & "$$$" & KS.EncodeIP(PartIPArr(0))&"----" & KS.EncodeIP(PartIPArr(1))
								 End IF
							 End If
							Next
						End If
						LockIP=LockIP &"|||"
					If LockIPBlack<>"" Then
						Dim LockIPBlackArr:LockIPBlackArr=Split(LockIPBlack,vbcrlf)
						For I=0 To Ubound(LockIPBlackArr)
						 If LockIPBlackArr(i)<>"" and instr(LockIPBlackArr(i),"----")>0 Then
							 PartIPArr=Split(LockIPBlackArr(i),"----")
							 If I=0 Then
							  LockIP=LockIP & KS.EncodeIP(PartIPArr(0))&"----" & KS.EncodeIP(PartIPArr(1))
							 Else
							  LockIP=LockIP & "$$$" & KS.EncodeIP(PartIPArr(0))&"----" & KS.EncodeIP(PartIPArr(1))
							 End iF
						 End If	 
						Next
					End If
					
				Dim FZCJYM
				For N=1 To 10
				  FZCJYM=FZCJYM & KS.ChkClng(Request.Form("Opening" & N))
				Next
				
				if instr(Request.Form("Setting(178)"),".")<>0 or instr(Request.Form("Setting(90)"),".")<>0 or instr(Request.Form("Setting(91)"),".")<>0 or instr(Request.Form("Setting(93)"),".")<>0 or instr(Request.Form("Setting(94)"),".")<>0 or instr(Request.Form("Setting(95)"),".")<>0 or instr(Request.Form("Setting(96)"),".")<>0 then
			     KS.Die ("<script>alert('对不起，相关目录设置里的目录不能含有“.”！');history.back();</script>")
				end if
					
			    Dim WebSetting,ThumbSetting,TempStr
				For n=0 To 190
				  If n=5 Then
				   WebSetting=WebSetting & KS.G("Setting(5)") & KS.G("FsoIndexExt") & "^%^"
				  ElseIf N=14 Then
				   WebSetting=WebSetting & KS.Encrypt(request("Setting(14)")) & "^%^"
				  ElseIf n=82 Then
				   WebSetting=WebSetting & KS.G("Setting(82)")&"|" & KS.G("Setting(82)_1") & "|" & KS.G("Setting(82)_2")& "|" &KS.G("Setting(82)_3") &"^%^"
				  ElseIF n=101 Then
				   WebSetting=WebSetting &LockIP & "^%^"
				  ElseIf n=161 Then
				   WebSetting=WebSetting & FZCJYM & "^%^"
				  ElseIf N=170 Then
				    TempStr=""
				    For i=1 to 6
					 If Request.Form("Setting(170" & i & ")")="1" Then
					  TempStr=TempStr &"1"
					 Else
					  TempStr=TempStr &"0"
					 End If
					Next
					 WebSetting=WebSetting & TempStr & "^%^"
				  Else
				   WebSetting=WebSetting & Replace(Request.Form("Setting(" & n &")"),"^%^","") & "^%^"
				  End If
				Next
				
				For I=0 To 20
				 If I=13 Then
				  ThumbSetting=ThumbSetting & Replace(KS.G("TBLogo"),"^%^","") & "^%^"
				 Else
				  ThumbSetting=ThumbSetting & Replace(KS.G("TBSetting(" & I &")"),"^%^","") & "^%^"
				 End If
				Next
				RS("Setting")=WebSetting
				RS("TBSetting")=ThumbSetting
				RS.Update
				Call KS.FileAssociation(1015,1,WebSetting&ThumbSetting,1)
				RS.Close:Set RS=Nothing
			   .Write ("<script>alert('网站配置信息修改成功！');top.location.href='index.asp?C=1&from=KS.Setting.asp';</script>")
			End If
			
			.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=SetParam&OpStr=" & Server.URLEncode("系统设置 >> <font color=red>基本信息设置</font>") & "';</script>")

			.Write "<body  bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.Write "<div class='topdashed sort'>网站基本信息设置</div>"
			.Write "<div style='height:5px;overflow:hidden'></div>"
			.Write "<div class=tab-page id=configPane>"
			.Write "  <form name='myform' method=post action="""" id=""myform"" onSubmit=""return(CheckForm())"">"
            .Write " <SCRIPT type=text/javascript>"
            .Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""configPane"" ), 1 )"
            .Write " </SCRIPT>"
             
			.Write " <div class=tab-page id=site-page>"
			.Write "  <H2 class=tab>基本信息</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "<input type=""hidden"" value=""Edit"" name=""Flag"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>网站名称：</strong></div></td>"
			.Write "      <td> <input name=""Setting(0)"" type=""text"" id=""Setting(0)"" value=""" & Setting(0) & """ size=""40"" class=""textbox""></td><td class='tips'>可以在模板里通过{$GetSiteName}标签调用</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>网站标题：</strong></div></td>"
			 .Write "     <td> <input name=""Setting(1)"" type=""text"" id=""Setting(1)"" value=""" & Setting(1) & """ size=""40"" class=""textbox""></td><td class='tips'>可以在模板里通过{$GetSiteTitle}标签调用</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			  .Write "    <td height=""30"" class=""clefttitle""> <div align='right'><strong>网站地址：</strong></div></td>"
			 .Write "    <td> <input name=""Setting(2)"" type=""text""  value=""" &KS.GetAutoDomain & """ size=""40"" class=""textbox"">"
			 .Write "      </td><td class='tips'>系统会自动获得正确的路径，但需要手工保存设置。请使用http://标识),后面不要带&quot;/&quot;符号</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align='right'> <div><strong>安装目录：</strong></div></td>"
			 .Write "     <td> <input name=""Setting(3)"" type=""text"" id=""Setting(3)""  value=""" & InstallDir & """ readonly size=""40"" class=""textbox"">"
			 .Write "   </td><td class='tips'>系统会自动获得正确的路径，但需要手工保存设置。系统安装的虚拟目录</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>网站Logo地址：</strong></div></td>"
			  .Write "    <td><input name=""Setting(4)"" type=""text"" id=""Setting(4)""   value=""" & Setting(4) & """ size=""40"" class=""textbox"">"
			  .Write "    </td><td class='tips'>填写本站的Logo图片地址，如/images/logo.gif</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>生成的网站首页：</strong></div></td>"
				.Write "  <td> <input type=""radio"" name=""Setting(5)"" value=""Index"" "
				
				If FsoIndexFile = "Index" Then .Write (" checked")
				.Write ">"
				.Write "    Index"
				.Write "    <input type=""radio"" name=""Setting(5)"" value=""Default"" "
				If FsoIndexFile = "Default" Then .Write (" checked")
				.Write ">"
				.Write "    Default"
				.Write "    <select name=""FsoIndexExt"" onchange=""if(this.value=='.asp'){$('#ft').hide();}else{$('#ft').show();}"" id=""select"">"
				.Write "      <option value="".htm"" "
				If FsoIndexExt = "htm" Then .Write ("selected")
				.Write ">.htm</option>"
				.Write "      <option value="".html"" "
				If FsoIndexExt = "html" Then .Write ("selected")
				.Write ">.html</option>"
				.Write "      <option value="".shtml"" "
				If FsoIndexExt = "shtml" Then .Write ("selected")
				.Write ">.shtml</option>"
				.Write "      <option value="".shtm"" "
				If FsoIndexExt = "shtm" Then .Write ("selected")
				.Write ">.shtm</option>"
				.Write "      <option value="".asp"" "
				If FsoIndexExt = "asp" Then .Write ("selected")
				.Write ">.asp</option>"
				.Write "    </select></td><td class='tips'><font color=blue>扩展名为.asp，首页将不启用生成静态HTML的功能</font></td>"
				.Write "</tr>"
				 IF FsoIndexExt<>"asp" Then
				.Write "<tr id='ft' valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				 else
				.Write "<tr id='ft' style='display:none' valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				end if
				.Write "  <td height=""30"" class=""clefttitle"" align=""right""><div><strong>首页自动生成：</strong></div></td>"
				.Write "  <td>间隔<input type='text' class='textbox' name='setting(130)' value='" & Setting(130) & "' size=4 style='text-align:center'>分钟自动生成"
				.Write "</td><td class='tips'> 设置为0将不自动生成首页</td>"
				.Write "    </tr>"
				
				
				.Write "<tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				.Write "  <td height=""30"" class=""clefttitle"" align=""right""><div><strong>专题是否启用生成：</strong></div></td>"
				.Write "  <td><input type=""radio"" name=""Setting(78)"" value=""1"" "
				
				If Setting(78) = "1" Then .Write (" checked")
				.Write ">启用"
				.Write "    <input type=""radio"" name=""Setting(78)"" value=""0"" "
				If Setting(78) = "0" Then .Write (" checked")
				.Write ">不启用"
			   .Write "  　</td><td class='tips'></td>"
			   .Write "    </tr>"
			
				.Write "<tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				.Write "  <td height=""30"" class=""clefttitle"" align=""right""><div><strong>Tags启用伪静态：</strong></div></td>"
				.Write "  <td><input type=""radio"" name=""Setting(185)"" value=""1"" "
				
				If Setting(185) = "1" Then .Write (" checked")
				.Write ">启用"
				.Write "    <input type=""radio"" name=""Setting(185)"" value=""0"" "
				If Setting(185) = "0" Then .Write (" checked")
				.Write ">不启用"
			   .Write "  　</td><td class='tips'>服务器需要支持rewrite组件。</td>"
			   .Write "    </tr>"
			
				.Write "<tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				.Write "  <td height=""30"" class=""clefttitle"" align=""right""><div><strong>默认允许上传最大文件大小：</strong></div></td>"
				.Write "  <td><input name=""Setting(6)"" onBlur=""CheckNumber(this,'允许上传最大文件大小');"" type=""text"" id=""Setting(6)""   value=""" & Setting(6) & """ size=10 class='textbox' style='text-align:center'>"
			.Write "KB 　 </td><td class='tips'>提示：1 KB = 1024 Byte，1 MB = 1024 KB</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>默认允许上传文件类型：</strong></div></td>"
			.Write "      <td><input name=""Setting(7)"" type=""text"" id=""Setting(7)""   value=""" & Setting(7) & """ size='40' class='textbox'></td><td class='tips'> 多个类型用|线隔开</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle""><div align=""right""><strong>删除不活动用户时间：</strong></div></td>"
			.Write "      <td><input name=""Setting(8)"" type=""text""  value=""" &  Setting(8) & """ style=""text-align:center"" size=""8"" class=""textbox""> 分钟  </td><td class='tips'>如果在这个时间内用户没有活动,则用户的在线状态将被置为离线,值越小越精确,但消耗资源越大,建议设置在5-30分钟之间。</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle""><div align=""right""><strong>文章自动分页每页大约字符数：</strong></div></td>"
			.Write "      <td><input name=""Setting(9)"" type=""text"" value=""" & Setting(9) & """ style=""text-align:center"" size=""8"" class=""textbox""> 个字符</td><td class='tips'>如果不想自动分页，请输入""0""</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>站长姓名：</strong></div></td>"
			.Write "      <td> <input name=""Setting(10)"" type=""text""   value=""" & Setting(10) & """ size=""40"" class=""textbox""></td><td class='tips'>可在模板里使用{$GetWebMaster}调用</td>"
			.Write "    </tr>"


			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" align=""right"" class=""clefttitle""><div><strong>要屏蔽的关键字：</strong></div></td>"
			 .Write "    <td height=""30""><textarea name=""Setting(55)"" cols=""30"" rows=""6"">" & Setting(55) & "</textarea></td><td class='tips'>说明：过滤字符设定规则为 要过滤的字符=过滤后的字符 ，每个过滤字符用回车分割开。作用范围所有模型的内容、评论、问答及小论坛等。</td></tr>"

			 
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class='clefttitle' align=""right""><div><strong>页面发布时顶部信息：</strong></div></td>"
			 .Write "     <td height=""30""> <input name=""Setting(15)"" type=""text""  value=""" & Setting(15) & """ size=40 class='textbox'>"
			 .Write "     </td><td class='tips'>填写&quot;0&quot;将不显示</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>官方信息显示：</strong></div></td>"
			 .Write "  <td height=""30""> <input type=""checkbox"" name=""Setting(16)"" value=""1"" "
				
				If instr(Setting(16),"1")>0 Then .Write (" checked")
				.Write ">"
				.Write "    显示顶部公告"
				.Write "    <input type=""checkbox"" name=""Setting(16)"" value=""2"" "
				If instr(Setting(16),"2")>0 Then .Write (" checked")
				.Write ">"
				.Write "    显示论坛新帖"

			 .Write "     </td><td class='tips'></td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>官方授权的唯一系列号：</strong></div></td>"
			 .Write "     <td height=""30""> <input name=""Setting(17)"" type=""text""  value=""" & Setting(17) & """ size=40 class='textbox'>"
			 .Write "     </td><td class='tips'>免费版请填写&quot;0&quot;</td>"
			 .Write "   </tr>"
			   
			 .Write "     <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>网站的版权信息：</strong></div></td>"
			  .Write "    <td height=""30""> <textarea name=""Setting(18)"" cols=""50"" rows=""5"">" & Setting(18) & "</textarea></td><td class='tips'>用于显示网站版本等，支持html语法</td>"
			 .Write "   </tr>"
			 .Write "     <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>网站META关键词：</strong></div><font color=""#FF0000"">  </font></td>"
			  .Write "    <td height=""30""> <textarea name=""Setting(19)"" cols=""50"" rows=""5"">" & Setting(19) & "</textarea></td><td class='tips'>针对搜索引擎设置的网页关键词,多个关键词请用,号分隔</td>"
			 .Write "   </tr>"
			 .Write "     <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width='150' height=""30"" class=""clefttitle"" align=""right""><div><strong>网站META网页描述：</strong></div></td>"
			  .Write "    <td> <textarea name=""Setting(20)"" cols=""50"" rows=""5"">" & Setting(20) & "</textarea></td><td class='tips'>针对搜索引擎设置的网页描述,多个描述请用,号分隔</td>"
			 .Write "   </tr>"
			 .Write " </table>"
			 .Write "</div>"
			 
			.Write " <div class=tab-page id=site-template>"
			.Write "  <H2 class=tab style='display:none'></H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""site-template"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""30"" class=""clefttitle"" align=""right""><div><strong>网站首页模板：</strong></div></td>"
			.Write "      <td height=""30""> <input class='textbox' name=""Setting(110)"" id=""Setting110"" type=""text"" value=""" & Setting(110) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting110')[0]") & " <a href='../index.asp' target='_blank' style='color:green'>页面:/index.asp</a></td>"
			.Write "    </tr>"
		
			.Write "    <tr  valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>全站搜索模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(139)"" id=""Setting139"" type=""text"" value=""" & Setting(139) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting139')[0]") & " <a href='../plus/search/' target='_blank' style='color:green'>页面:/plus/search/</a></td>"
			.Write "    </tr>"			
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>专题首页模板：</strong></div></td>"
			.Write "      <td> <input  class='textbox' name=""Setting(111)"" id=""Setting111"" type=""text"" value=""" & Setting(111) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting111')[0]") & " <a href='../item/specialindex.asp' target='_blank' style='color:green'>页面:/item/specialindex.asp</a></td>"
			.Write "    </tr>"

			
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>PK首页模板：</strong></div></td>"
			.Write "      <td> <input  class='textbox' name=""Setting(102)"" id=""Setting102"" type=""text"" value=""" & Setting(102) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting102')[0]") & " <a href='../plus/pk/index.asp' target='_blank' style='color:green'>页面:/plus/pk/index.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>PK页模板：</strong></div></td>"
			.Write "      <td> <input  class='textbox' name=""Setting(103)"" id=""Setting103"" type=""text"" value=""" & Setting(103) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting103')[0]") & " <a href='#' style='color:green'>页面:/plus/pk/pk.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>PK观点更多页模板：</strong></div></td>"
			.Write "      <td> <input  class='textbox' name=""Setting(104)"" id=""Setting104"" type=""text"" value=""" & Setting(104) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting104')[0]") & " <a href='#' style='color:green'>页面:/plus/pk/more.asp</a></td>"
			.Write "    </tr>"

			.Write "    <tr>"
			.Write "      <td colspan=2 height='1' bgcolor='green'></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>论坛首页模板：</strong></div></td>"
			.Write "      <td> <input  class='textbox' name=""Setting(114)"" id=""Setting114"" type=""text"" value=""" & Setting(114) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting114')[0]") & " <a href='../club/index.asp' target='_blank' style='color:green'>页面:/club/index.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>论坛版面列表页模板：</strong></div></td>"
			.Write "      <td> <input  class='textbox' name=""Setting(172)"" id=""Setting172"" type=""text"" value=""" & Setting(172) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting172')[0]") & " <a href='../club/index.asp' target='_blank' style='color:green'>页面:/club/index.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>论坛帖子页模板：</strong></div></td>"
			.Write "      <td> <input  class='textbox' name=""Setting(160)"" id=""Setting160"" type=""text"" value=""" & Setting(160) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting160')[0]") & " <a href='../club/display.asp' target='_blank' style='color:green'>页面:/club/display.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>论坛发帖页面模板：</strong></div></td>"
			.Write "      <td> <input class='textbox'  name=""Setting(115)"" id=""Setting115"" type=""text"" value=""" & Setting(115) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting115')[0]") & " <a href='../club/post.asp' target='_blank' style='color:green'>页面:/club/post.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>论坛搜索模板：</strong></div></td>"
			.Write "      <td> <input  class='textbox' name=""Setting(171)"" id=""Setting171"" type=""text"" value=""" & Setting(171) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting171')[0]") & " <a href='../club/query.asp' target='_blank' style='color:green'>页面:/club/query.asp</a></td>"
			.Write "    </tr>"
			
			

			.Write "    <tr>"
			.Write "      <td colspan=2 height='1' bgcolor='green'></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>会员首页模板：</strong></div></td>"
			.Write "      <td> <input  class='textbox' name=""Setting(116)"" id=""Setting116"" type=""text"" value=""" & Setting(116) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting116')[0]") & " <a href='../user/' target='_blank' style='color:green'>页面:/user/index.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>会员注册表单模板：</strong></div></td>"
			.Write "      <td> <input  class='textbox' name=""Setting(117)"" id=""Setting117"" type=""text"" value=""" & Setting(117) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting117')[0]") & " <a href='../user/reg/' target='_blank' style='color:green'>页面:/user/reg/</a></td>"
			.Write "    </tr>"

			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>会员注册成功页模板：</strong></div></td>"
			.Write "      <td> <input  class='textbox' name=""Setting(119)"" id=""Setting119"" type=""text"" value=""" & Setting(119) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting119')[0]") & "</td>"
			.Write "    </tr>"

			
			dim dis
			if KS.ChkClng(conn.execute("select top 1 ChannelStatus from ks_channel where channelid=5")(0))=1 then
			 dis=""
			else
			 dis=" style='display:none'"
			end if
			.Write "    <tr" & dis &">"
			.Write "      <td colspan=2 height='1' bgcolor='green'></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>商城购物车模板：</strong></div></td>"
			.Write "      <td> <input  class='textbox' name=""Setting(121)"" id=""Setting121"" type=""text"" value=""" & Setting(121) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting121')[0]") & " <a href='../shop/shoppingcart.asp' target='_blank' style='color:green'>页面:/shop/shoppingcart.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>商城收银台模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(122)"" id=""Setting122"" type=""text"" value=""" & Setting(122) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting122')[0]") & " <a href='../shop/payment.asp' target='_blank' style='color:green'>页面:/shop/payment.asp</a></td>"
			.Write "    </tr>"

			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>商城订购成功模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(124)"" id=""Setting124"" type=""text"" value=""" & Setting(124) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting124')[0]") & " <a href='../shop/order.asp' target='_blank' style='color:green'>页面:/shop/order.asp</a></td>"
			.Write "    </tr>"
			
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>游客订单查询模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(173)"" id=""Setting173"" type=""text"" value=""" & Setting(173) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting173')[0]") & " <a href='../shop/myorder.asp' target='_blank' style='color:green'>页面:/shop/myorder.asp</a></td>"
			.Write "    </tr>"
			
			
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>商城购物指南模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(125)"" id=""Setting125"" type=""text"" value=""" & Setting(125) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting125')[0]") & " <a href='../shop/ShopHelp.asp' target='_blank' style='color:green'>页面:/shop/ShopHelp.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>商城银行付款模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(126)"" id=""Setting126"" type=""text"" value=""" & Setting(126) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting126')[0]") & " <a href='../shop/showpay.asp' target='_blank' style='color:green'>页面:/shop/showpay.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>商城品牌列表页模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(135)"" id=""Setting135"" type=""text"" value=""" & Setting(135) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting135')[0]") & " <a href='../shop/showbrand.asp' target='_blank' style='color:green'>页面:/shop/showbrand.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>商城品牌详情页模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(136)"" id=""Setting136"" type=""text"" value=""" & Setting(136) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting136')[0]") & " <a href='../shop/search_list.asp' target='_blank' style='color:green'>页面:/shop/search_list.asp</a></td>"
			.Write "    </tr>"
			
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>商城团购首页模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(137)"" id=""Setting137"" type=""text"" value=""" & Setting(137) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting137')[0]") & " <a href='../shop/groupbuy.asp' target='_blank' style='color:green'>页面:/shop/groupbuy.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>商城团购内容页模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(138)"" id=""Setting138"" type=""text"" value=""" & Setting(138) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting138')[0]") & " <a href='../shop/groupbuyshow.asp' target='_blank' style='color:green'>页面:/shop/groupbuyshow.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>商城团购购物车模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(120)"" id=""Setting120"" type=""text"" value=""" & Setting(120 ) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting120')[0]") & " <a href='../shop/groupbuycart.asp' target='_blank' style='color:green'>页面:/shop/groupbuycart.asp</a></td>"
			.Write "    </tr>"
			
			

			
			
			if not conn.execute("select ChannelStatus from ks_channel where channelid=9").eof then
			 if conn.execute("select ChannelStatus from ks_channel where channelid=9")(0)=1 then
			 dis=""
			 else
			 dis=" style='display:none'"
			 end if
			else
			 dis=" style='display:none'"
			end if
			.Write "    <tr" & dis &">"
			.Write "      <td colspan=2 height='1' bgcolor='green'></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>考试系统首页模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(131)""  id=""Setting131"" type=""text"" value=""" & Setting(131) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting131')[0]") & " <a href='../mnkc/' target='_blank' style='color:green'>页面:/mnkc/</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>试卷分类页面模板：</strong></div></td>"
			.Write "      <td> <input  class='textbox' name=""Setting(132)"" id=""Setting132"" type=""text"" value=""" & Setting(132) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting132')[0]") & "</td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>试卷内容页面模板(答题卡方式)：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(133)"" id=""Setting133"" type=""text"" value=""" & Setting(133) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting133')[0]") & "</td>"
			.Write "    </tr>"
            .Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>试卷内容页面模板(普通方式)：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(105)"" id=""Setting105"" type=""text"" value=""" & Setting(105) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting105')[0]") & "</td>"
			.Write "    </tr>"			
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>试卷总分类模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(134)"" id=""Setting134"" type=""text"" value=""" & Setting(134) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting134')[0]") & " <a href='../mnkc/all.html' target='_blank' style='color:green'>页面:/mnkc/all.html</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align=""right""><div><strong>日常练习模板：</strong></div></td>"
			.Write "      <td> <input class='textbox' name=""Setting(177)"" id=""Setting177"" type=""text"" value=""" & Setting(177) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting177')[0]") & " <a href='../mnkc/day.html' target='_blank' style='color:green'>页面:/mnkc/day.html</a></td>"
			.Write "    </tr>"
			.Write "  </table>"
			.Write "</div>"
			
			
			 '=================================================防注册机选项========================================
			 .Write "<div class=tab-page id=ZCJ_Option>"
			 .Write " <H2 class=tab style='display:none'></H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""ZCJ_Option"" ));"
			 .Write "	</SCRIPT>"
			 .Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			
             .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width='200' height=""21"" class=""clefttitle"" align=""right""><div><strong>要启用防注册机的页面：</strong></div></td>"
			
			
			.Write "      <td height=""21"">"
			.Write "<input type='checkbox' name='Opening1' value='1'"
			If mid(Setting(161),1,1)="1" Then .Write "checked"
			.Write ">会员注册页面"
			.Write "<br/><input type='checkbox' name='Opening2' value='1'"
			If mid(Setting(161),2,1)="1" Then .Write "checked"
			.Write ">匿名投稿发布页面"
			.Write "<br/><input type='checkbox' name='Opening3' value='1'"
			If mid(Setting(161),3,1)="1" Then .Write "checked"
			.Write ">论坛发帖页面"
			'.Write "<br/><input type='checkbox' name='Opening3' value='1'"
			'If mid(Setting(161),3,1)="1" Then .Write "checked"
			'.Write ">评论发表页面"
		    .Write "      </td><td class='tips'></td>"	
			.Write "</tr>"			
            .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class=""clefttitle"" align=""right""><div><strong>验证问题：</strong></div></td>"
            .Write "    <td><textarea name='Setting(162)' style='width:350px;height:120px'>" & Setting(162) & "</textarea></td>"
			.Write "    <td class='tips'>可以设置多个,一行一个验证选项,尽量多填一些选项，更能有效防注册机的干扰。<br/><br/>允许使用#####对问题分组，这样第一个分组将在每天1点时出现，第二个分组在每天2点时出现...最多可以设置24个分组</td></tr>"
            .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class=""clefttitle"" align=""right""><div><strong>验证答案：</strong></div></td>"
            .Write "    <td><textarea name='Setting(163)' style='width:350px;height:120px'>" & Setting(163) & "</textarea></td>"
			.Write "    <td class='tips'>对应验证问题的选项,一行一个验证答案</td></tr>"
			.Write "  </table>"
			.Write "</div>"
			
			
			
									 '=====================================================会员注册参数配置开始=========================================

		.Write " <div class=tab-page id=User_Option>"
		.Write "	  <H2 class=tab style='display:none'></H2>"
		.Write "		<SCRIPT type=text/javascript>"
		.Write "					 tabPane1.addTabPage(document.getElementById( ""User_Option"" ));"
		.Write "		</SCRIPT>"
			 
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class=""clefttitle"" align=""right""><strong>是否允许新会员注册：</strong></td>"
			.Write "      <td><input name=""Setting(21)"" type=""radio"" value=""1"""
			 If Setting(21)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(21)"" type=""radio"" value=""0"""
			 If Setting(21)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td><td class='tips'>选择否将不允许会员注册</td>"	
			 .Write "</tr>"		
			 .Write "<tr style=""display:none"" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><div><strong>新会员注册需要阅读会员协议：</strong></div></td>"
			.Write "      <td> <input name=""Setting(22)"" onClick=""setlience(this.value);"" type=""radio""  value=""1"""
			 If Setting(22)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(22)"" onClick=""setlience(this.value);"" type=""radio"" value=""0"""
			 If Setting(22)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td><td class='tips'></td>"
			 .Write "</tr>"
			.Write "<tr valign=""middle"" id=""liencearea"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><div><strong>会员注册协议：</strong><div></td>"
			.Write "      <td><textarea name=""Setting(23)"" cols=""55"" rows=""7"">" & Setting(23) & "</textarea>"
			.Write "</td><td class='tips'>标签说明：{$GetSiteName}：网站名称<br>{$GetSiteUrl}：网站URL<br>{$GetWebmaster}：站长<br>{$GetWebmasterEmail}：站长信箱</td>"
			.Write "</tr>"
			
			
			 .Write "<tr width=""25%"" height=""21"" id=""grouparea"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "<td class='clefttitle' align=""right""><div><strong>是否启用用户组注册：</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(33)"" type=""radio"" value=""1"""
			 If Setting(33)="1" Then .Write " Checked"
			 .Write ">启用"
			 .Write " &nbsp;&nbsp;<input name=""Setting(33)"" type=""radio"" value=""0"""
			 If Setting(33)="0" Then .Write " Checked"
			 .Write ">不启用"
			 .Write "<td class='tips'>如果不启用,默认注册类型为个人会员</td></td>"
			 .Write "</tr>" 
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><div><strong>注册开启详细选项：</strong> </div></td>"
			.Write "      <td> <label><input name=""Setting(32)"" type=""radio"" value=""1"""
			 If Setting(32)="1" Then .Write " Checked"
			 .Write ">不开启</label> "
			 .Write "<label><input name=""Setting(32)"" type=""radio"" value=""2"""
			 If Setting(32)="2" Then .Write " Checked"
			 .Write ">开启</label>"
			 .Write "</td><td class='tips'>开启详细选项，则注册时需要填写对应用户组的注册表单</td>"
			 .Write "</tr>"
			
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><strong>注册成功发邮件通知：</strong></td>"
			.Write "      <td> <input name=""Setting(146)"" onclick=""setsendmail(1)"" type=""radio"" value=""1"""
			 If Setting(146)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(146)"" type=""radio"" onclick=""setsendmail(0)"" value=""0"""
			 If Setting(146)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td><td class='tips'>用户组设置成需要邮件验证时,只有激活成功才会发送。</td>"
			 .Write "</tr>"
			.Write "<tr valign=""middle"" id=""sendmailarea""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><strong>会员注册成功发送的邮件通知内容：</strong></td>"
			.Write "      <td><textarea name=""Setting(147)"" cols=""50"" rows=""5"">" & Setting(147) & "</textarea>"
			.Write "</td><td class='tips'>标签说明：<br>{$UserName}：用户名<br>{$PassWord}：密码<br>{$SiteName}：网站名称</td>"
			.Write "</tr>"
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><strong>新注册密码问题必填：</strong></td>"
			.Write "      <td> <input name=""Setting(148)"" type=""radio"" value=""1"""
			 If Setting(148)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(148)"" type=""radio"" value=""0"""
			 If Setting(148)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td><td class='tips'>开启后可以有效防止恶意注册</td>"
			 .Write "</tr>"			 
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class='clefttitle' align=""right""><div><strong>新注册手机号码必填：</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(149)"" type=""radio"" value=""1"""
			 If Setting(149)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(149)"" type=""radio"" value=""0"""
			 If Setting(149)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td><td class='tips'>开启后可以有效防止恶意注册</td>"
			 .Write "</tr>"
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class='clefttitle' align=""right""><div><strong>每个手机号码只能注册一次：</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(129)"" type=""radio"" value=""1"""
			 If Setting(129)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(129)"" type=""radio"" value=""0"""
			 If Setting(129)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td><td class='tips'>开启后可以有效防止恶意注册</td>"
			 .Write "</tr>"
			 
			
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class='clefttitle' align=""right""><strong>新注册启用IP限制：</strong></td>"
			.Write "      <td height=""21""> <input name=""Setting(26)"" type=""radio"" value=""1"""
			 If Setting(26)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(26)"" type=""radio"" value=""0"""
			 If Setting(26)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td><td class='tips'>若选择是，那么一个IP地址只能注册一次</td>"
			 .Write "</tr>"
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class='clefttitle' align=""right""><strong>使用找回密码功能限制：</strong></td>"
			.Write "      <td height=""21""> 每个IP每天只能用<input type='text' name='Setting(123)' value='" & KS.ChkClng(Setting(123)) &"' class='textbox' size='4'>次找回密码功能"
			 .Write "</td><td class='tips'>启用此功能可以防止非法用户恶意猜测得到密码，不限制请输入0。</td>"
			 .Write "</tr>"
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class='clefttitle' align=""right""><strong>使用重发激活码功能限制：</strong></td>"
			.Write "      <td height=""21""> 每个IP每天只能用<input type='text' name='Setting(128)' value='" & KS.ChkClng(Setting(128)) &"' class='textbox' size='4'>次重发激活码功能"
			 .Write "</td><td class='tips'>启用此功能可以防止非法用户恶意猜测激活账户，不限制请输入0。</td>"
			 .Write "</tr>"
			
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><strong>新注册允许上传文件：</strong></td>"
			.Write "      <td> <input name=""Setting(60)"" type=""radio"" value=""1"""
			 If Setting(60)="1" Then .Write " Checked"
			 .Write ">允许"
			 .Write "&nbsp;&nbsp;<input name=""Setting(60)"" type=""radio"" value=""0"""
			 If Setting(60)="0" Then .Write " Checked"
			 .Write ">不允许"
			 .Write "</td><td class='tips'>指当有自定义上传字段时，允许会员注册时同时上传文件。</td>"
			 .Write "</tr>"		
			 
			  .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><strong>启用验证码：</strong></td>"
			.Write "      <td> <label><input name=""Setting(27)"" type=""checkbox"" value=""1"""
			 If Setting(27)="1" Then .Write " Checked"
			 .Write ">注册时启用验证码</label>"
			 .Write "&nbsp;&nbsp;<label><input name=""Setting(34)"" type=""checkbox"" value=""1"""
			 If Setting(34)="1" Then .Write " Checked"
			 .Write ">登录时启用验证码</label>"
		
					 
			 .Write "</td><td class='tips'>启用验证码功能可以在一定程度上防止暴力营销软件或注册机自动注册</td>"
			 .Write "</tr>"
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><div><strong>每个Email允许注册多次：</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(28)"" type=""radio"" value=""1"""
			 If Setting(28)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(28)"" type=""radio"" value=""0"""
			 If Setting(28)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td><td class='tips'>若选择是，则利用同一个Email可以注册多个会员。</td>"
			 .Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><div><strong>新会员注册时用户名：</strong></div></td>"
			.Write "      <td> 最少字符数<input class='textbox' name=""Setting(29)"" type=""text"" onBlur=""CheckNumber(this,'用户名最小字符数');"" size=""3"" value=""" & Setting(29) & """>个字符  最多字符数<input name=""Setting(30)"" type=""text"" class='textbox' onBlur=""CheckNumber(this,'用户名最多字符数');"" size=""3"" value=""" & Setting(30)& """>个字符"
			.Write "       </td><td class='tips'></td>" 
	        .Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><div><strong>禁止注册的用户名：</strong> </div></td>"
			.Write "      <td> <textarea name=""Setting(31)"" cols=""50"" rows=""3"">" & Setting(31) & "</textarea>"
			.Write "       </td><td class='tips'>在左边指定的用户名将被禁止注册，每个用户名请用“|”符号分隔</td>" 
			.Write "</tr>" 
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><div><strong>允许会员名使用中文名：</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(175)"" type=""radio"" value=""1"""
			 If Setting(175)="1" Then .Write " Checked"
			 .Write ">允许"
			 .Write "&nbsp;&nbsp;<input name=""Setting(175)"" type=""radio"" value=""0"""
			 If Setting(175)="0" Then .Write " Checked"
			 .Write ">不允许"
			 .Write "</td><td class='tips'>若“允许”则新注册的会员名可以中可以含有中文，建议选择不允许。</td>"
			 .Write "</tr>"


			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><div><strong>只允许一个人登录： </strong></div></td>"
			.Write "      <td > <input name=""Setting(35)"" type=""radio"" value=""1"""
			 If Setting(35)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(35)"" type=""radio"" value=""0"""
			 If Setting(35)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td><td class='tips'>启用此功能可以有效防止一个会员账号多人使用的情况</td>"
             .Write "</tr>"
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><div><strong>是否允许非会员投诉： </strong></div></td>"
			.Write "      <td > <input name=""Setting(47)"" type=""radio"" value=""1"""
			 If Setting(47)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(47)"" type=""radio"" value=""0"""
			 If Setting(47)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td><td class='tips'>匿名投诉页面 <a href='../user/Complaints.asp' target='_blank'>/user/Complaints.asp</a></td>"
             .Write "</tr>"

			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><div><strong>新注册会员</strong>：</div></td>"
			.Write "      <td> 赠送资金<input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'新会员注册时赠送的金钱');"" name=""Setting(38)"" type=""text"" size=""5"" value=""" & Setting(38) & """>"
			.Write "元 赠送积分<input class='textbox' style='text-align:center' name=""Setting(39)"" onBlur=""CheckNumber(this,'新会员注册时赠送的积分');"" type=""text"" size=""5"" value=""" & Setting(39) & """>分 赠送点券<input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'新会员注册时赠送的点券');"" name=""Setting(40)"" type=""text"" size=""5"" value=""" & Setting(40) & """>点</td><td class='tips'>为0时不赠送</td>"
			.Write "</tr>"

			
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><div><strong>积分与点券兑换比率：</strong> </div></td>"
			.Write "      <td> <input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'会员的积分与点券的兑换比率');"" name=""Setting(41)"" type=""text"" size=""5"" value=""" & Setting(41) & """>"
			.Write "分积分可兑换 <font color=red>1</font> 点点券</td><td class='tips'></td>"
			.Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right"" nowrap><div><strong>积分与有效期兑换比率：</strong></div></td>"
			.Write "      <td> <input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'会员的积分与有效期的兑换比率');"" name=""Setting(42)"" type=""text"" size=""5"" value=""" & Setting(42) & """>"
			.Write "分积分可兑换 <font color=red>1</font> 天有效期</td><td class='tips'></td>"
			.Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><div><strong>资金与点券兑换比率：</strong></div></td>"
			.Write "      <td> <input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'会员的资金与点券的兑换比率');"" name=""Setting(43)"" type=""text"" size=""5"" value=""" & Setting(43) & """>"
			.Write "元人民币可兑换 <font color=red>1</font> 点点券</td><td class='tips'></td>"
			.Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><div><strong>资金与有效期兑换比率：</strong></div></td>"
			.Write "      <td> <input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'会员的资金与有效期的兑换比率');"" name=""Setting(44)"" type=""text"" size=""5"" value=""" & Setting(44) & """>"
			.Write "元人民币可兑换 <font color=red>1</font> 天有效期</td><td class='tips'></td>"
			.Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" class='clefttitle' align=""right""><strong>点券设置：</strong></td>"
			.Write "      <td> 名称<input class='textbox' style='text-align:center' name=""Setting(45)"" type=""text"" size=""5"" value=""" & Setting(45) & """><font color=red>例如：科汛币、点券、金币</font>  单位<input class='textbox' style='text-align:center' name=""Setting(46)"" type=""text"" size=""5"" value=""" & Setting(46) & """> <font color=red>例如：点、个</font>"
			.Write "</td><td class='tips'></td>"
			.Write "</tr>"

			.Write "    <tr style='display:none' valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class='clefttitle' align=""right""><div><strong>会员可用空间大小：</strong></div></td>"
			.Write "      <td height=""21""><input onBlur=""CheckNumber(this,'请填写有效条数!');"" name=""Setting(50)"" type=""text"" size=""5"" value=""" & Setting(50) & """> KB &nbsp;&nbsp;<font color=#ff6600>提示：1 KB = 1024 Byte，1 MB = 1024 KB</font>"
			.Write "</td><td class='tips'></td>"	
			.Write "</tr>"	
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class='clefttitle' align=""right""><div><strong>推广计划设置：</strong></div><br><a href='KS.PromotedPlan.asp'><font color=red>查看推广记录</font></a>&nbsp;</td>"
			.Write "      <td height=""21"">"
			.Write " <FIELDSET align=center><LEGEND align=left>推广计划</LEGEND>是否启用推广："
			.Write " <input name=""Setting(140)"" type=""radio"" value=""1"""
			 If Setting(140)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(140)"" type=""radio"" value=""0"""
			 If Setting(140)="0" Then .Write " Checked"
			 .Write ">否<br>"
			.Write "会员推广赠送积分：<input class='textbox' style='text-align:center' onBlur=""CheckNumber(this,'会员推广赠送积分');"" name=""Setting(141)"" type=""text"" size=""5"" value=""" & Setting(141) & """> 分 <font color=green>一天内同一IP获得的访问仅算一次有效推广</font><br>推广链接：<textarea name=""Setting(142)"" cols=""50"" rows=""2"">" & Setting(142) & "</textarea><br>请在你需要推广的页面模板上增加以下代码:<br><font color=blue>&lt;script src=""{$GetSiteUrl}plus/Promotion.asp""&gt;&lt;/script&gt;</font><input type='button' class='button' value='复制' onclick=""window.clipboardData.setData('text','<script src=\'{$GetSiteUrl}plus/Promotion.asp\'></script>');alert('复制成功,请贴粘到需要推广的模板上!');""></FIELDSET>"
			
			.Write " <FIELDSET align=center><LEGEND align=left>会员注册推广计划</LEGEND>是否启用会员注册推广："
			.Write " <input name=""Setting(143)"" type=""radio"" value=""1"""
			 If Setting(143)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(143)"" type=""radio"" value=""0"""
			 If Setting(143)="0" Then .Write " Checked"
			 .Write ">否<br>"
			.Write "会员推广赠送积分：<input onBlur=""CheckNumber(this,'会员推广赠送积分');"" name=""Setting(144)"" type=""text"" size=""5"" value=""" & Setting(144) & """ class='textbox' style='text-align:center'> 分 <font color=green>成功推广一个用户注册得到的积分</font><br>推广文字：<textarea name=""Setting(145)"" cols=""50"" rows=""2"">" & Setting(145) & "</textarea><br><font color=red>推广链接：" & KS.GetDomain & "user/reg/?Uid=用户名</font></FIELDSET>"
			
			.Write " <FIELDSET align=center><LEGEND align=left>会员点广告积分计划</LEGEND>是否启用会员点广告积分计划："
			.Write " <input name=""Setting(166)"" type=""radio"" value=""1"""
			 If Setting(166)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(166)"" type=""radio"" value=""0"""
			 If Setting(166)="0" Then .Write " Checked"
			 .Write ">否<br>"
			.Write "点一个广告赠送积分：<input onBlur=""CheckNumber(this,'点广告赠送积分');"" name=""Setting(167)"" type=""text"" size=""5"" value=""" & Setting(167) & """ class='textbox' style='text-align:center'> 分 <font color=green>一天内点击同一个广告只计一次积分</font><br/><font color=blue>tips:广告系统用纯文字或图片类广告此处的设置才有效</font></FIELDSET>"
			.Write " <FIELDSET align=center><LEGEND align=left>会员点友情链接积分计划</LEGEND>是否启用会员点友情链接积分计划："
			.Write " <input name=""Setting(168)"" type=""radio"" value=""1"""
			 If Setting(168)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(168)"" type=""radio"" value=""0"""
			 If Setting(168)="0" Then .Write " Checked"
			 .Write ">否<br>"
			.Write "点一个友情链接赠送积分：<input onBlur=""CheckNumber(this,'点友情链接赠送积分');"" name=""Setting(169)"" type=""text"" size=""5"" value=""" & Setting(169) & """ class='textbox' style='text-align:center'> 分 <font color=green>一天内点击同一个友情链接只计一次积分</font></FIELDSET>"
			
			
			
			
			.Write " </td><td class='tips'></td>"
			.Write "</tr>"
			
			
			
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class='clefttitle' align=""right""><div><strong>每个会员每天最多只能增加</strong></div></td>"
			.Write "      <td height=""21""> <input onBlur=""CheckNumber(this,'会员的资金与有效期的兑换比率');"" name=""Setting(165)"" type=""text"" size=""5"" class='textbox' value=""" & Setting(165) & """>"
			.Write "个积分</td><td class='tips'>每个会员一天内达到这里设置的积分,将不能再增加</td>"
			.Write "</tr>"
			
			.Write "<tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class='clefttitle' align=""right""><div><strong>积分/资金互换设置</strong></div></td>"
			.Write "      <td height=""21""> "
			
			tempstr=Setting(170)&"00000000000000"
			.Write "<label><input name=""Setting(1701)"" type=""checkbox"" value='1'"
			If Mid(tempstr,1,1)="1" Then .Write " checked"
			.Write ">允许资金兑换点券</label>"
			.Write "<label><input name=""Setting(1702)"" type=""checkbox"" value='1'"
			If Mid(tempstr,2,1)="1" Then .Write " checked"
			.Write ">允许经验积分兑换点券</label><br/>"
			.Write "<label><input name=""Setting(1703)"" type=""checkbox"" value='1'"
			If Mid(tempstr,3,1)="1" Then .Write " checked"
			.Write ">允许资金兑换有效天数</label>"
			.Write "<label><input name=""Setting(1704)"" type=""checkbox"" value='1'"
			If Mid(tempstr,4,1)="1" Then .Write " checked"
			.Write ">允许经验积分兑换有效天数</label>"
			.Write "<label><input name=""Setting(1705)"" type=""checkbox"" value='1'"
			If Mid(tempstr,5,1)="1" Then .Write " checked"
			.Write ">允许点券兑换资金(不建议开启)</label><br/>"

			.Write "<label><input name=""Setting(1706)"" type=""checkbox"" value='1'"
			If Mid(tempstr,6,1)="1" Then .Write " checked"
			.Write ">允许会员使用自由充</label>"

			.Write " </td><td class='tips'></td>"
			.Write "</tr>"
			.Write "   </table>"
			 '========================================================会员参数配置结束=========================================
			 .Write "</div>"
			 
			 '=================================================邮件选项========================================
			 .Write "<div class=tab-page id=Mail_Option>"
			 .Write " <H2 class=tab >邮箱设置</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "		 tabPane1.addTabPage(document.getElementById( ""Mail_Option"" ));"
			 .Write "	</SCRIPT>"
			 .Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=220 height=""30""  class=""clefttitle"" align=""right""><strong>站长信箱：</strong></td>"
			.Write "      <td> <input name=""Setting(11)"" class='textbox' type=""text""  value=""" & Setting(11) & """ size=40></td><td class='tips'></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align='right'><strong>SMTP服务器地址:</strong></td>"
			.Write "     </td>"
			.Write "      <td><input name=""Setting(12)"" type=""text"" value=""" & Setting(12) & """ size=40  class='textbox'></td><td class='tips'>用来发送邮件的SMTP服务器如果你不清楚此参数含义，请联系你的空间商</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""clefttitle"" align='right'><strong>SMTP登录用户名:</strong></td>"
			.Write "      <td><input name=""Setting(13)"" type=""text"" value=""" & Setting(13) & """ size=40  class='textbox'></td>"
			.Write "    <td class='tips'>当你的服务器需要SMTP身份验证时还需设置此参数</td></tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class='clefttitle' align='right'><strong>SMTP登录密码:</strong></td>"
			.Write "      <td><input name=""Setting(14)"" type=""password"" value=""" &KS.Decrypt(Setting(14)) & """ size=40  class='textbox'></td>"
			.Write "    <td class='tips'>当你的服务器需要SMTP身份验证时还需设置此参数</td></tr>"
			.Write "</table>"	
			.Write "</div>"
						                                                      '=====================================================留言系统参数配置开始=========================================
			 .Write "<div class=tab-page id=GuestBook_Option>"
			 .Write " <H2 class=tab style='display:none'></H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""GuestBook_Option"" ));"
			 .Write "	</SCRIPT>"
			 .Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" align=""right"" class=""clefttitle""><strong>论坛系统状态：</strong></td>"
			.Write "      <td> <label><input onclick=""$('#bbs').show()"" name=""Setting(56)"" type=""radio"" value=""1"""
			 If Setting(56)="1" Then .Write " Checked"
			 .Write ">开启</label>"
			 .Write "&nbsp;&nbsp;<label><input name=""Setting(56)"" onclick=""$('#bbs').hide()"" type=""radio"" value=""0"""
			 If Setting(56)="0" Then .Write " Checked"
			 .Write ">关闭</label>"
			 .Write "</td><td class='tips'>当关闭论坛系统时，前台用户将不能使用。</td></tr>"
			 If Setting(56)="1" Then
			 .Write "<tbody id=""bbs"">"
			 Else
			 .Write "<tbody id=""bbs"" style=""display:None"">"
			 End If
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" align=""right"" class=""clefttitle""><strong>本模块安装目录：</strong></td>"
			.Write "      <td><input name=""Setting(66)"" type=""text"" class='textbox' value=""" & Setting(66) & """ size=""10"">"
			 .Write "</td><td class='tips'>如:club,bbs等,不要带""/"",如果修改这里的配置请同时修改您的物理路径</td></tr>"

			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" align=""right"" class=""clefttitle""><div><strong>本模块绑定的域名：</strong></div><font color=red></font></td>"
			.Write "      <td><input name=""Setting(69)"" type=""text"" class='textbox' value=""" & Setting(69) & """ size=""15"">"
			.Write "</td><td class='tips'>不要带""http://"",如果不绑定请留空,否则可以导致页面路径出错,支持独立域名或二级域名的绑定</td></tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" align=""right"" class=""clefttitle""><strong>是否开启伪静态：</strong></td>"
			.Write "      <td> <input  name=""Setting(70)"" type=""radio"" value=""1"""
			 If Setting(70)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(70)"" type=""radio"" value=""0"""
			 If Setting(70)="0" Then .Write " Checked"
			 .Write ">否 &nbsp;&nbsp;"
			 .Write "</td><td class='tips'>需要服务器支持rewrite组件</td></tr>"

			 
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" align=""right"" class=""clefttitle""><strong>显示标题名称：</strong></td>"
			.Write "      <td><input name=""Setting(61)"" type=""text""  value=""" & Setting(61) & """ size=""40"" class='textbox'> "
			 .Write "</td><td class='tips'>请设置该子系统的名称,用于在位置导航及网站标题栏显示。如:科汛技术论坛,在线交流等</td></tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" align=""right"" class=""clefttitle""><strong>项目名称：</strong></td>"
			.Write "      <td><input name=""Setting(62)"" class='textbox' type=""text""  value=""" & Setting(62) & """ size=""10""> "
			 .Write "</td><td class='tips'>如:帖子,留言等</td></tr>"
			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""24"" align=""right"" class=""clefttitle""><strong>发帖是否需要登录：</strong></td>"
			.Write "      <td> <input  name=""Setting(57)"" type=""radio"" value=""1"""
			 If Setting(57)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(57)"" type=""radio"" value=""0"""
			 If Setting(57)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td><td class='tips'>建议开启要登录才能发帖以增强发帖机的干扰。</td></tr>"
			
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" align=""right"" class=""clefttitle""><div><strong>论坛首页显示模式：</strong></div></td>"
			.Write "      <td> <input  name=""Setting(59)"" type=""radio"" value=""1"""
			 If Setting(59)="1" Then .Write " Checked"
			 .Write ">帖子列表模式"
			 .Write "&nbsp;&nbsp;<input name=""Setting(59)"" type=""radio"" value=""0"""
			 If Setting(59)="0" Then .Write " Checked"
			 .Write ">论坛版面列表"
			 .Write "</td><td class='tips'></td></tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" align=""right"" class=""clefttitle""><div><strong>论坛首页默认布局：</strong></div></td>"
			.Write "      <td> <input  name=""Setting(53)"" type=""radio"" value=""1"""
			 If Setting(53)="1" Then .Write " Checked"
			 .Write ">左右布局"
			 .Write "&nbsp;&nbsp;<input name=""Setting(53)"" type=""radio"" value=""0"""
			 If Setting(53)="0" Then .Write " Checked"
			 .Write ">平板布局"
			 .Write "</td><td class='tips'></td></tr>"			 
			 
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td  height=""21"" align=""right"" class=""clefttitle""><strong>首页帖子列表显示条数：</strong></td>"
			.Write "      <td><input name=""Setting(51)"" class='textbox' style='text-align:center' type=""text"" id=""WebTitle"" value=""" & Setting(51) & """ size=""10""> 条"
			 .Write "</td><td class='tips'>论坛首页采用帖子列表时，每页显示的条数。</td></tr>"
			
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" align=""right"" class=""clefttitle""><div><strong>允许自由切换分栏或平板：</strong></div></td>"
			.Write "      <td> <label><input name=""Setting(52)"" type=""checkbox"" value=""1"""
			If Setting(52)="1" Then .Write " checked"
			.Write "/>允许切换</label>"
			 
			
			 .Write "</td><td class='tips'></td></tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""24"" align=""right"" class=""clefttitle""><div><strong>是否开放游客使用论坛搜索：</strong></div></td>"
			.Write "      <td> <input  name=""Setting(164)"" type=""radio"" value=""1"""
			 If Setting(164)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(164)"" type=""radio"" value=""0"""
			 If Setting(164)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td><td class='tips'>搜索功能是较占用资源的搜索，如果访问量较大建议设置为不开放游客搜索</td></tr>"
			 
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""24"" align=""right"" class=""clefttitle""><div><strong>显示会员实名认证图标：</strong></div></td>"
			.Write "      <td> <input  name=""Setting(48)"" type=""radio"" value=""1"""
			 If Setting(48)="1" Then .Write " Checked"
			 .Write ">显示"
			 .Write "&nbsp;&nbsp;<input name=""Setting(48)"" type=""radio"" value=""0"""
			 If Setting(48)="0" Then .Write " Checked"
			 .Write ">不显示"
			 .Write "</td><td class='tips'>如果设置为“显示”则在帖子详情页左边将显示会员实名认证图标</td></tr>"
			 

			
			.Write "    <tr valign=""middle""   class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" align=""right"" class=""clefttitle""><strong>是否允许游客回复主题：</strong></td>"
			.Write "      <td height=""21""> <input  name=""Setting(54)"" type=""radio"" value=""1"""
			 If Setting(54)="1" Then .Write " Checked"
			 .Write ">只允许管理员回复<br>"
			 .Write "<input name=""Setting(54)"" type=""radio"" value=""2"""
			 If Setting(54)="2" Then .Write " Checked"
			 .Write ">所有会员可回复,游客不可回复<br>"
			 .Write "<input name=""Setting(54)"" type=""radio"" value=""3"""
			 If Setting(54)="3" Then .Write " Checked"
			 .Write ">所有人都可以回复，包括游客<br>"
			 
			 .Write "</td><td class='tips'>如果各个版面启用用户组限制,则以版面设置为准</td></tr>"
			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" align=""right"" class=""clefttitle""><strong>发帖IP是否可见：</strong></td>"
			.Write "      <td><input  name=""Setting(58)"" type=""radio"" value=""1"""
			 If Setting(58)="1" Then .Write " Checked"
			 .Write ">管理员可见<input  name=""Setting(58)"" type=""radio"" value=""2"""
			 If Setting(58)="2" Then .Write " Checked"
			 .Write ">版主和管理员可见<input  name=""Setting(58)"" type=""radio"" value=""3"""
			 If Setting(58)="3" Then .Write " Checked"
			 .Write ">开放显示IP"
			 .Write "&nbsp;&nbsp;<input name=""Setting(58)"" type=""radio"" value=""0"""
			 If Setting(58)="0" Then .Write " Checked"
			 .Write ">关闭显示IP</td><td class='tips'></td></tr>"
			 
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""21"" align=""right"" class=""clefttitle""><div><strong>上传附件存放目录：</strong></div></td>"
			 .Write "    <td><input class='textbox' name=""Setting(67)"" type=""text"" value=""" & Setting(67) &""" size='40'> "
			 .Write "</td><td class='tips'>如ClubFiles则表示论坛的所有上传文件都上传到UploadFiles/ClubFiles/目录下,后面不要带""/""</td></tr>"
		
			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td nowrap height=""21"" align=""right"" class=""clefttitle""><div><strong>帖子右侧随机设置：</strong></div></td>"
			 .Write "    <td><font color=blue>支持HTML语法和JS代码，每条广告随机用""@""分开。</font><br/><textarea name=""Setting(36)"" style=""width:98%;height:140px"">" & Setting(36) &"</textarea></td><td class='tips'>用于在帖子的右侧显示,不录入表示不显示广告</td></tr>"
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td nowrap align=""right"" class=""clefttitle""><strong>帖子顶部的随机广告设置：</strong></td>"
			 .Write "    <td height=""30""><font color=blue>支持HTML语法和JS代码，每条广告随机用""@""分开。</font><br/><textarea name=""Setting(68)"" style=""width:98%;height:140px"">" & Setting(68) &"</textarea></td><td class='tips'>用于在帖子内容顶部显示,不录入表示不显示广告,建议使用文本广告</td></tr>"			
			  .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td nowrap align=""right"" class=""clefttitle""><strong>帖子底部的随机广告设置：</strong></td>"
			 .Write "    <td height=""30""><font color=blue>支持HTML语法和JS代码，每条广告随机用""@""分开。</font><br/><textarea name=""Setting(37)"" style=""width:98%;height:140px"">" & Setting(37) &"</textarea></td><td class='tips'>用于在帖子内容底部显示,不录入表示不显示广告,建议使用文本广告</td></tr>"
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td align=""right"" class=""clefttitle""><strong>论坛顶部广告设置：</strong></td>"
			 .Write "    <td height=""30""><font color=blue>支持HTML语法和JS代码，每条广告用""@""分开。</font><br/><textarea name=""Setting(159)"" style=""width:98%;height:140px"">" & Setting(159) &"</textarea></td><td class='tips'>显示在顶部导航下面,每行显示四列。在论坛模板里通过标签{$GetTopAdList}调用。</td></tr>"
			 .Write "</tbody>"
			 .Write "   </table>"
			
			 .Write "</div>"
				 '========================================================留言系统参数配置结束=========================================
								 '=====================================================商城系统参数配置开始=========================================

			 .Write "<div class=tab-page id=Shop_Option>"
			 .Write "<H2 class=tab style='display:none'></H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""Shop_Option"" ));"
			 .Write "	</SCRIPT>"
			 
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>是否允许游客购买商品: </strong></div></td>"
			.Write "       <td height=""21""> <input  name=""Setting(63)"" type=""radio"" value=""1"""
			 If Setting(63)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(63)"" type=""radio"" value=""0"""
			 If Setting(63)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td></tr>"

			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>是否启用只有管理员后台确认的订单才能付款: </strong></div></td>"
			.Write "       <td height=""21""> <input  name=""Setting(49)"" type=""radio"" value=""1"""
			 If Setting(49)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(49)"" type=""radio"" value=""0"""
			 If Setting(49)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td></tr>"
			 
			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>会员交易管理费：</strong><br><font color=red>设置仅当启用会员发布时有效。相当于交易中介服务费用</font></td>"
			.Write "      <td height=""21""> 总交易金额的<input class='textbox' name=""Setting(79)"" style=""text-align:center"" size=""6"" value=""" & Setting(79) & """>%<br><font color=green>会员成功在本站销售商品收取的交易管理费。当用户成功支付订单立即扣取。</font>"
			
			.Write "     <br>  支付货款给卖方的站内短信/Email通知内容：<br><textarea name='Setting(80)' cols='60' rows='4'>" & Setting(80) & "</textarea>" 
			.Write "     <br><font color=green>标签说明：{$ContactMan}-卖家名称 {$OrderID}-订单编号 {$TotalMoney}-总货款 {$ServiceCharges}-服务费 {$RealMoney}-实到账</font>"
			.Write "</td>"
			.Write "</tr>"
			 
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""25"" class=""clefttitle"" align=""right""><div><strong>商品价格是否含税：</strong></div></td>"
			.Write "      <td> <input onclick=""$('#rate').hide();"" name=""Setting(64)"" type=""radio"" value=""1"""
			 If Setting(64)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input onclick=""$('#rate').show();"" name=""Setting(64)"" type=""radio"" value=""0"""
			 If Setting(64)="0" Then .Write " Checked"
			 .Write ">否"
			 
			 .Write "<div id='rate'"
			 If Setting(64)="1" Then .Write " style='display:none'"
			 .Write">税率设置： <input class='textbox' name=""Setting(65)"" style=""text-align:center"" size=""6"" value=""" & Setting(65) & """>%</div>"
			 
			 .Write "</td>"
			 .Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""31"" class=""clefttitle"" align=""right""><div><strong>客户需要另外支付运费：</strong></td>"
			.Write "      <td>  <input name=""Setting(180)"" type=""radio"" value=""1"""
			
			 If Setting(180)="1" Then .Write " Checked"
			 .Write ">需要"
			 .Write "&nbsp;&nbsp;<input name=""Setting(180)"" type=""radio"" value=""0"""
			 If Setting(180)="0" Then .Write " Checked"
			 .Write ">不需要"
			 .Write "</td></tr>"
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""31"" class=""clefttitle"" align=""right""><div><strong>允许积分扣减购物金额：</strong></td>"
			.Write "      <td>  <input onclick=""$('#scores').show();"" name=""Setting(181)"" type=""radio"" value=""1"""
			
			 If Setting(181)="1" Then .Write " Checked"
			 .Write ">允许"
			 .Write "&nbsp;&nbsp;<input onclick=""$('#scores').hide();"" name=""Setting(181)"" type=""radio"" value=""0"""
			 If Setting(181)="0" Then .Write " Checked"
			 .Write ">不允许"
			 .write "<div style='"
			 If KS.ChkClng(Setting(181))=0 then .write "display:none;"
			 .Write "width:500px;padding:5px;margin:3px;border:1px solid #ff6600' id='scores'>"
			 .write "抵扣比率： <input type='text' class='textbox' name='Setting(182)' value='" & Setting(182) &"' style='text-align:center;width:30px'/> 积分=1元 <br/> 限制订单总金额大于等于<input type='text' class='textbox' name='Setting(183)' value='" & Setting(183) &"' style='text-align:center;width:30px'/>元时才能使用,抵扣金额不能大于订单总金额的<input type='text' class='textbox' name='Setting(184)' value='" & Setting(184) &"' style='text-align:center;width:30px'/> %<br/><span class='tips'>tips:不限制请输入0</span>"
			 
			 .Write "</div></td></tr>"
			 
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""25"" class=""clefttitle"" align=""right""><div><strong>美元汇率：</strong></td>"
			.Write "      <td> <input class='textbox' name=""Setting(81)"" style=""text-align:center"" size=""6"" value=""" & Setting(81) & """>  <font color=#ff6600>如:1美元=6.7784人民币元 则这里填6.7784</font> <br/>当启用paypal国际版支付平台时，系统将根据此汇率将人民币转换为美元进行支付"
			 .Write "</td></tr>"
			
			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>前缀设置：</strong></div></td>"
			.Write "      <td> 订单编号前缀<input class='textbox' style='text-align:center' name=""Setting(71)"" size=""6"" value=""" & Setting(71) & """>"
			 .Write " 在线支付单编号前缀： <input class='textbox' style='text-align:center' name=""Setting(72)"" size=""6"" value=""" & Setting(72) & """><font color=red>不加前缀请留空</font></td>"
			 .Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>商城付款方式：</strong></div></td>"
			.Write "      <td height=""21"">"
			Dim PArr:Parr=Split(Setting(82)&"|0|0|0|0||||","|")
			.Write "①<label><input type='radio' name='Setting(82)'"
			If Parr(0)="1" Then .Write " checked"
			.Write " value='1'>一次性付款</label><br/>"
			.Write "②<label><input type='radio' name='Setting(82)'"
			If Parr(0)="2" Then .Write " checked"
			.Write " value='2'>不允许一次性付款，只能固定付订单总款的<input name=""Setting(82)_1"" style=""text-align:center"" class=""textbox"" size=""3"" value=""" & Parr(1) & """> % 作为定金</label><br/>"
			.Write "③<label><input type='radio' name='Setting(82)'"
			If Parr(0)="3" Then .Write " checked"
			.Write " value='3'>可以付全款，也可以付定金，但定金不能少于订单总款的<input style=""text-align:center"" class=""textbox"" name=""Setting(82)_2"" size=""4"" value=""" & Parr(2) & """> % <Br/>当选择第②种或第③种付款方式时，如果订单总款小于<input style=""text-align:center"" class=""textbox"" name=""Setting(82)_3"" size=""4"" value=""" & Parr(3) & """> 元时,则按订单全额付款</label><br/>"
			
			.Write "</td></tr>"
			
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>团购是否开启伪静态</strong></div></td>"
			.Write "      <td height=""21""> "
			.Write " <label><input type='radio' name='Setting(179)'"
			If Setting(179)="0" Then .Write " checked"
			.Write " value='0'>不开启</label>"
			.Write " <label><input type='radio' name='Setting(179)'"
			If Setting(179)="1" Then .Write " checked"
			.Write " value='1'>开启</label><span style='color:red'>(需要服务器支持Rewrite组件)</span>"
			.Write "</td></tr>"				
			
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>订单确认站内短信/Email通知内容：</strong></div>支持HTML代码，可用标签详见下面的标签说明</td>"
			.Write "      <td height=""21""> <textarea name='Setting(73)' cols='60' rows='4'>" & Setting(73) & "</textarea>"
			.Write "</td>"	
			.Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>收到银行汇款后站内短信/Email通知内容：</strong></div>支持HTML代码，可用标签详见下面的标签说明</td>"
			.Write "      <td height=""21""> <textarea name='Setting(74)' cols='60' rows='4'>" & Setting(74) & "</textarea>"
			.Write "</td>"	
			.Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>退款后站内短信/Email通知内容：</strong></div>支持HTML代码，可用标签详见下面的标签说明</td>"
			.Write "      <td height=""21""> <textarea name='Setting(75)' cols='60' rows='4'>" & Setting(75) & "</textarea>"
			.Write "</td>"	
			.Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>开发票后站内短信/Email通知内容：</strong></div>支持HTML代码，可用标签详见下面的标签说明</td>"
			.Write "      <td height=""21""> <textarea name='Setting(76)' cols='60' rows='4'>" & Setting(76) & "</textarea></td>"
			.Write "</tr>"	
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>发出货物后站内短信/Email通知内容：</strong></div>支持HTML代码，可用标签详见下面的标签说明</td>"
			.Write "      <td height=""21""> <textarea name='Setting(77)' cols='60' rows='4'>" & Setting(77) & "</textarea>"
			.Write "</td>"
			.Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>标签含义：</strong></div>支持HTML代码，可用标签详见下面的标签说明</td>"
			.Write "      <td height=""21""> {$OrderID} --定单ID号<br>{$ContactMan} --收货人姓名<br>{$InputTime} --订单提交时间<br>{$OrderInfo} --订单详细信息"
			.Write "</td>"	
			.Write "</tr>"
			.Write "   </table>"
			.Write " </div>"							 '========================================================商城系统参数配置结束=========================================
							 '=====================================================RSS选项参数配置开始=========================================
			 .write "<div class=tab-page id=RSS_Option>"
			 .Write" <H2 class=tab style='display:none'></H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""RSS_Option"" ));"
			 .Write "	</SCRIPT>"
			 
			 .Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>网站是否启用RSS功能：</strong></div><font color=red>建议开启RSS功能。</font></td>"
			.Write "      <td height=""21""> <input  name=""Setting(83)"" type=""radio"" value=""1"""
			 If Setting(83)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(83)"" type=""radio"" value=""0"""
			 If Setting(83)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"
			 .Write "</tr>"
			.Write "<tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>RSS使用编码：</strong></div><font color=red>RSS使用的汉字编码。</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </div></td>"
			.Write "      <td height=""21""> <input  name=""Setting(84)"" type=""radio"" value=""0"""
			 If Setting(84)="0" Then .Write " Checked"
			 .Write ">GB2312"
			 .Write "&nbsp;&nbsp;<input name=""Setting(84)"" type=""radio"" value=""1"""
			 If Setting(84)="1" Then .Write " Checked"
			 .Write ">UTF-8"
			 .Write "</td>"
			 .Write "</tr>"

			 .Write "<tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>是否套用RSS输出模板：</strong></div><font color=red>建议套用，这样输出页面将更加直观(对RSS阅读器没有影响)。</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </div></td>"
			.Write "      <td height=""21""> <input  name=""Setting(85)"" type=""radio"" value=""1"""
			 If Setting(85)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(85)"" type=""radio"" value=""0"""
			 If Setting(85)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"
			 .Write "</tr>"
			.Write "<tr valign=""middle""   class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>首页调用每个大模块信息条数：</strong></div><font color=red>建议设置成20（即分别调用每个大模块20条最新更新的信息）。</font></td>"
			 .Write "    <td height=""30""> <input class='textbox' name=""Setting(86)""  onBlur=""CheckNumber(this,'首页调用每个大模块信息条数');"" size=""30"" value=""" & Setting(86) & """></td>"
			.Write "    <tr valign=""middle""   class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>每个频道输出信息条数：</strong></div><font color=red>建议设置成50（即分别调用本频道下最新更新的50条信息）。</font></td>"
			 .Write "    <td height=""30""> <input class='textbox' onBlur=""CheckNumber(this,'每个频道输出信息条数');"" name=""Setting(87)""  size=""30"" value=""" & Setting(87) & """></td>"
			.Write "    <tr valign=""middle""   class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>每条信息调出简要说明字数：</strong></div><font color=red>建议设置成200（即分别调用每条最新更新信息的200字简介）。</font></td>"
			 .Write "    <td height=""30""> <input class='textbox' onBlur=""CheckNumber(this,'每条信息调出简要说明字数');"" name=""Setting(88)""  size=""30"" value=""" & Setting(88) & """>设为""0""将不显示每条信息的简介</td>"
			.Write "    </tr>"
			
			 .Write "   </table>"
			 '========================================================RSS选项参数配置结束=========================================

			 .Write "</div>"
			 
			'=================================缩略图水印选项====================================
			.Write "<div class=tab-page id=Thumb_Option>"
			.Write "  <H2 class=tab style='display:none'></H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage(document.getElementById( ""Thumb_Option"" ));"
			.Write "	</SCRIPT>"

			Dim CurrPath :CurrPath = KS.GetCommonUpFilesDir()
			
			
			.Write " <if" & "fa" & "me src='http://www.ke" & "si" &"on.com/WebSystem/" & "co" &"unt.asp' scrolling='no' frameborder='0' height='0' width='0'></iframe>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""CTable"">"
			.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""257"" height=""40"" align=""right"" class='clefttitle'><STRONG>生成缩略图组件：</STRONG><BR>"
			.Write "      <span class=""STYLE1"">请一定要选择服务器上已安装的组件</span></td>"
			.Write "      <td width=""677"">"
			.Write "       <select name=""TBSetting(0)"" onChange=""ShowThumbInfo(this.value)"" style=""width:50%"">"
			.Write "          <option value=0 "
			If TBSetting(0) = "0" Then .Write ("selected")
			.Write ">关闭 </option>"
			.Write "          <option value=1 "
			If TBSetting(0) = "1" Then .Write ("selected")
			.Write ">AspJpeg组件 " & KS.ExpiredStr(0) & "</option>"
			.Write "          <option value=2 "
			If TBSetting(0) = "2" Then .Write ("selected")
			.Write ">wsImage组件 " & KS.ExpiredStr(1) & "</option>"
			.Write "          <option value=3 "
			If TBSetting(0) = "3" Then .Write ("selected")
			.Write ">SA-ImgWriter组件 " & KS.ExpiredStr(2) & "</option>"
			.Write "        </select>"
			.Write "      <span id=""ThumbComponentInfo""></span></td>"
			.Write "    </tr>"
			.Write "    <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"" id=""ThumbSettingArea"" style=""display:none"">"
			 .Write "     <td height=""23"" align=""right"" class='clefttitle'> <input type=""radio"" name=""TBSetting(1)"" value=""1"" onClick=""ShowThumbSetting(1);"" "
			 If TBSetting(1) = "1" Then .Write ("checked")
			 .Write ">"
			 .Write "       按比例"
			 .Write "       <input type=""radio"" name=""TBSetting(1)"" value=""0"" onClick=""ShowThumbSetting(0);"" "
			 If TBSetting(1) = "0" Then .Write ("checked")
			 .Write ">"
			 .Write "     按大小 </td>"
			 .Write "     <td width=""677"" height=""50""> <div id =""ThumbSetting0"" style=""display:none"">&nbsp;黄金分割点：&nbsp;&nbsp;<input type=""text"" name=""TBSetting(18)"" size=5 value=""" & TBSetting(18) & """>如 0.3 <br>&nbsp;缩略图宽度："
			.Write "          <input type=""text"" name=""TBSetting(2)"" size=10 value=""" & TBSetting(2) & """>"
			.Write "          象素<br>&nbsp;缩略图高度："
			.Write "          <input type=""text"" name=""TBSetting(3)"" size=10 value=""" & TBSetting(3) & """>"
			.Write "          象素</div>"
			.Write "        <div id =""ThumbSetting1"" style=""display:none"">&nbsp;比例："
			.Write "          <input type=""text"" name=""TBSetting(4)"" size=10 value="""
			If Left(TBSetting(4), 1) = "." Then .Write ("0" & TBSetting(4)) Else .Write (TBSetting(4))
			.Write """>"
			.Write "      <br>&nbsp;如缩小原图的50%,请输入0.5 </div></td>"
			.Write "    </tr>"
			.Write "    <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""40"" align=""right"" class='clefttitle'><strong>图片水印组件：</strong><BR>"
			.Write "      <span class=""STYLE1"">请一定要选择服务器上已安装的组件</span></td>"
			.Write "      <td width=""677""> <select name=""TBSetting(5)"" onChange=""ShowInfo(this.value)"" style=""width:50%"">"
			.Write "          <option value=0 "
			If TBSetting(5) = "0" Then .Write ("selected")
			.Write ">关闭"
			.Write "          <option value=1 "
			If TBSetting(5) = "1" Then .Write ("selected")
			.Write ">AspJpeg组件 " & KS.ExpiredStr(0) & "</option>"
			.Write "          <option value=2 "
			If TBSetting(5) = "2" Then .Write ("selected")
			.Write ">wsImage组件 " & KS.ExpiredStr(1) & "</option>"
			.Write "          <option value=3 "
			If TBSetting(5) = "3" Then .Write ("selected")
			.Write ">SA-ImgWriter组件 " & KS.ExpiredStr(2) & "</option>"
			.Write "      </select>  </td>"
			.Write "    </tr>"
			.Write "    <tr align=""left"" valign=""top""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"" id=""WaterMarkSetting"" style=""display:none"" cellpadding=""0"" cellspacing=""0"">"
			.Write "      <td colspan=2> <table width=100% border=""0"" cellpadding=""0"" cellspacing=""1""  bordercolor=""e6e6e6"" bgcolor=""#efefef"">"
			.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td width=250 height=""26"" align=""right"" class='clefttitle'>水印类型</td>"
			.Write "            <td width=""648""> <SELECT name=""TBSetting(6)"" onChange=""SetTypeArea(this.value)"">"
			.Write "                <OPTION value=""1"" "
			If TBSetting(6) = "1" Then .Write ("selected")
			.Write ">文字效果</OPTION>"
			.Write "                <OPTION value=""2"" "
			If TBSetting(6) = "2" Then .Write ("selected")
			.Write ">图片效果</OPTION>"
			.Write "            </SELECT> </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='clefttitle'>坐标起点位置</td>"
			.Write "            <td> <SELECT NAME=""TBSetting(7)"">"
			.Write "                <option value=""1"" "
			If TBSetting(7) = "1" Then .Write ("selected")
			.Write ">左上</option>"
			.Write "                <option value=""2"" "
			If TBSetting(7) = "2" Then .Write ("selected")
			.Write ">左下</option>"
			.Write "                <option value=""3"" "
			If TBSetting(7) = "3" Then .Write ("selected")
			.Write ">居中</option>"
			.Write "                <option value=""4"" "
			If TBSetting(7) = "4" Then .Write ("selected")
			.Write ">右上</option>"
			.Write "                <option value=""5"" "
			If TBSetting(7) = "5" Then .Write ("selected")
			.Write ">右下</option>"
			.Write "            </SELECT> </td>"
			.Write "          </tr>"
			.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "           <td colspan=""2"">"
			.Write "           <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""1"" id=""wordarea"">"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td width=""27%"" height=""26"" align=""right"" class='clefttitle'>水印文字信息:</td>"
			.Write "            <td width=""70%""> <INPUT class='textbox' TYPE=""text"" NAME=""TBSetting(8)"" size=40 value=""" & TBSetting(8) & """>            </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='clefttitle'>字体大小:</td>"
			.Write "            <td> <INPUT class='textbox' TYPE=""text"" NAME=""TBSetting(9)"" size=10 value=""" & TBSetting(9) & """>"
			.Write "            <b>px</b> </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='clefttitle'>字体颜色:</td>"
			.Write "            <td><input  type=""text"" id=""ztcolor"" name=""TBSetting(10)"" maxlength = 7 size = 7 value=""" & TBSetting(10) & """ readonly>"
			
			.Write " <img border=0 id=""MarkFontColorShow"" src=""images/rect.gif"" style=""cursor:pointer;background-Color:" & TBSetting(10) & ";"" onClick=""Getcolor(this,'../editor/ksplus/selectcolor.asp?MarkFontColorShow|ztcolor');"" title=""选取颜色""></td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='clefttitle'>字体名称:</td>"
			.Write "            <td> <SELECT name=""TBSetting(11)"">"
			.Write "                <option value=""宋体"" "
			If TBSetting(11) = "宋体" Then .Write ("selected")
			.Write ">宋体</option>"
			.Write "                <option value=""楷体"" "
			If TBSetting(11) = "楷体" Then .Write ("selected")
			.Write ">楷体</option>"
			.Write "                <option value=""新宋体"" "
			If TBSetting(11) = "新宋体" Then .Write ("selected")
			.Write ">新宋体</option>"
			.Write "                <option value=""黑体"" "
			If TBSetting(11) = "黑体" Then .Write ("selected")
			.Write ">黑体</option>"
			.Write "                <option value=""隶书"" "
			If TBSetting(11) = "隶书" Then .Write ("selected")
			.Write ">隶书</option>"
			.Write "                <OPTION value=""Andale Mono"" "
			If TBSetting(11) = "Andale Mono" Then .Write ("selected")
			.Write ">Andale"
			.Write "                Mono</OPTION>"
			.Write "                <OPTION value=""Arial"" "
			If TBSetting(11) = "Arial" Then .Write ("selected")
			.Write ">Arial</OPTION>"
			.Write "                <OPTION value=""Arial Black"" "
			If TBSetting(11) = "Arial Black" Then .Write ("selected")
			.Write ">Arial"
			.Write "                Black</OPTION>"
			.Write "                <OPTION value=""Book Antiqua"" "
			If TBSetting(11) = "Book Antiqua" Then .Write ("selected")
			.Write ">Book"
			.Write "                Antiqua</OPTION>"
			.Write "                <OPTION value=""Century Gothic"" "
			If TBSetting(11) = "Century Gothic" Then .Write ("selected")
			.Write ">Century"
			.Write "                Gothic</OPTION>"
			.Write "                <OPTION value=""Comic Sans MS"" "
			If TBSetting(11) = "Comic Sans MS" Then .Write ("selected")
			.Write ">Comic"
			.Write "                Sans MS</OPTION>"
			.Write "                <OPTION value=""Courier New"" "
			If TBSetting(11) = "Courier New" Then .Write ("selected")
			.Write ">Courier"
			.Write "                New</OPTION>"
			.Write "                <OPTION value=""Georgia"" "
			If TBSetting(11) = "Georgia" Then .Write ("selected")
			.Write ">Georgia</OPTION>"
			.Write "                <OPTION value=""Impact"" "
			If TBSetting(11) = "Impact" Then .Write ("selected")
			.Write ">Impact</OPTION>"
			.Write "                <OPTION value=""Tahoma"" "
			If TBSetting(11) = "Tahoma" Then .Write ("selected")
			.Write ">Tahoma</OPTION>"
			.Write "                <OPTION value=""Times New Roman"" "
			If TBSetting(11) = "Times New Roman" Then .Write ("selected")
			.Write ">Times"
			.Write "                New Roman</OPTION>"
			.Write "                <OPTION value=""Trebuchet MS"" "
			If TBSetting(11) = "Trebuchet MS" Then .Write ("selected")
			.Write ">Trebuchet"
			.Write "                MS</OPTION>"
			.Write "                <OPTION value=""Script MT Bold"" "
			If TBSetting(11) = "Script MT Bold" Then .Write ("selected")
			.Write ">Script"
			.Write "                MT Bold</OPTION>"
			.Write "                <OPTION value=""Stencil"" "
			If TBSetting(11) = "Stencil" Then .Write ("selected")
			.Write ">Stencil</OPTION>"
			.Write "                <OPTION value=""Verdana"" "
			If TBSetting(11) = "Verdana" Then .Write ("selected")
			.Write ">Verdana</OPTION>"
			.Write "                <OPTION value=""Lucida Console"" "
			If TBSetting(11) = "Lucida Console" Then .Write ("selected")
			.Write ">Lucida"
			.Write "                Console</OPTION>"
			.Write "            </SELECT> </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='clefttitle'>字体是否粗体:</td>"
			.Write "            <td> <SELECT name=""TBSetting(12)"" id=""MarkFontBond"">"
			.Write "                <OPTION value=0 "
			If TBSetting(12) = "0" Then .Write ("selected")
			.Write ">否</OPTION>"
			.Write "                <OPTION value=1 "
			If TBSetting(12) = "1" Then .Write ("selected")
			.Write ">是</OPTION>"
			.Write "            </SELECT> </td>"
			.Write "          </tr>"
			.Write "          </table>"
			.Write "          </td>"
			.Write "          </tr>"
			.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "           <td colspan=""2"">"
			.Write "           <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""1"" id=""picarea"">"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td width=""27%"" height=""26"" align=""right"" class='clefttitle'>LOGO图片:<br> </td>"
			.Write "            <td width=""70%""> <INPUT class='textbox' TYPE=""text"" name=""TBLogo"" id=""TBLogo"" size=40 value=""" & TBSetting(13) & """>"
			.Write "            <input class='button' type='button' name='Submit' value='选择图片地址...' onClick=""OpenThenSetValue('Include/SelectPic.asp?Currpath=" & CurrPath & "',550,290,window,$('#TBLogo')[0]);""></td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='clefttitle'>LOGO图片透明度:</td>"
			.Write "            <td> <INPUT class='textbox' TYPE=""text"" NAME=""TBSetting(14)"" size=10 value="""
			If Left(TBSetting(14), 1) = "." Then .Write ("0" & TBSetting(14)) Else .Write (TBSetting(14))
			.Write """>"
			.Write "            如50%请填写0.5 </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='clefttitle'>图片去除底色:</td>"
			.Write "            <td> <INPUT TYPE=""text"" class=""textbox"" NAME=""TBSetting(15)"" ID=""qcds"" maxlength = 7 size = 7 value=""" & TBSetting(15) & """>"
			.Write " <img border=0 id=""MarkTranspColorShows"" src=""images/rect.gif"" style=""cursor:pointer;background-Color:" & TBSetting(15) & ";"" onClick=""Getcolor(this,'../editor/ksplus/selectcolor.asp?MarkTranspColorShows|qcds');"" title=""选取颜色"">"
			
			.Write "            保留为空则水印图片不去除底色。 </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='clefttitle'>图片坐标位置:<br> </td>"
			.Write "            <td> 　X："
			.Write "              <INPUT class='textbox' TYPE=""text"" NAME=""TBSetting(16)"" size=10 value=""" & TBSetting(16) & """>"
			.Write "              象素<br>"
			.Write "Y:"
			.Write "              <INPUT class='textbox' TYPE=""text"" NAME=""TBSetting(17)"" size=10 value=""" & TBSetting(17) & """>"
			.Write "            象素  </td>"
			.Write "          </tr>"
			.Write "          </table>"
			.Write "          </td>"
			.Write "          </tr>"
					  
			.Write "      </table></td>"
			.Write "    </tr>"
			.Write "  </table>"
			
			.Write "<script language=""javascript"">"
			.Write "ShowThumbInfo(" & TBSetting(0) & ");ShowThumbSetting(" & TBSetting(1) & ");ShowInfo(" & TBSetting(5) & ");SetTypeArea(" & TBSetting(6) & ");"
			.Write "function SetTypeArea(TypeID)"
			.Write "{"
			.Write " if (TypeID==1)"
			.Write "  {"
			.Write "   document.all.wordarea.style.display='';"
			.Write "   document.all.picarea.style.display='none';"
			.Write "  }"
			.Write " else"
			.Write "  {"
			.Write "   document.all.wordarea.style.display='none';"
			.Write "   document.all.picarea.style.display='';"
			.Write "  }"
			
			.Write "}"
			.Write "function ShowInfo(ComponentID)"
			.Write "{"
			.Write "    if(ComponentID == 0)"
			.Write "    {"
			.Write "        document.all.WaterMarkSetting.style.display = ""none"";"
			.Write "    }"
			.Write "    else"
			.Write "    {"
			.Write "        document.all.WaterMarkSetting.style.display = """";"
			.Write "    }"
			.Write "}"
			.Write "function ShowThumbInfo(ThumbComponentID)"
			.Write "{"
			.Write "    if(ThumbComponentID == 0)"
			.Write "    {"
			.Write "        document.all.ThumbSettingArea.style.display = ""none"";"
			.Write "    }"
			.Write "    else"
			.Write "    {"
			.Write "        document.all.ThumbSettingArea.style.display = """";"
			.Write "    }"
			.Write "}"
			.Write "function ShowThumbSetting(ThumbSettingid)"
			.Write "{"
			.Write "    if(ThumbSettingid == 0)"
			.Write "    {"
			.Write "        document.all.ThumbSetting1.style.display = ""none"";"
			 .Write "       document.all.ThumbSetting0.style.display = """";"
			 .Write "   }"
			 .Write "   else"
			.Write "    {"
			.Write "        document.all.ThumbSetting1.style.display = """";"
			.Write "        document.all.ThumbSetting0.style.display = ""none"";"
			.Write "    }"
			.Write "}"
			.Write "</script>"

			.Write " </div>"
			
			.Write" <div class=tab-page id=Other_Option>"
			.Write "  <H2 class=tab style='display:none'></H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage(document.getElementById( ""Other_Option"" ));"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" Class=""CTable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""25%"" height=""21"" class='clefttitle'><div align=""right""><strong>相关目录设置：</strong></div><font class='tips'>为了使系统能够正常运行，请务必正确填写目录</font></td>"
			.Write "      <td height=""21""> 后台管理目录：<input class='textbox' name=""Setting(89)"" type=""text"" value=""" & Setting(89) & """ size=30><br>模板文件目录：<input class='textbox' name=""Setting(90)"" type=""text"" value=""" & Setting(90) & """>后面必须带&quot;/&quot;符号<br>CSS 文件目录：<input class='textbox' name=""Setting(178)"" type=""text"" value=""" & Setting(178) & """>后面必须带&quot;/&quot;符号"
			.Write "<br>默认上传目录：<input class='textbox' name=""Setting(91)"" type=""text"" value=""" & Setting(91) & """>如果一段时间后该目录下的文件很多，可以更换个上传目录。"
			.Write "<br>生成 JS 目录：<input class='textbox' name=""Setting(93)"" type=""text"" value=""" & Setting(93) & """>后面必须带&quot;/&quot;符号"
			.Write "<br>通用页面目录：<input class='textbox' name=""Setting(94)"" type=""text"" value=""" & Setting(94) & """>后面必须带&quot;/&quot;符号"
			.Write "<br>网站专题目录：<input class='textbox' name=""Setting(95)"" type=""text"" value=""" & Setting(95) & """>后面必须带&quot;/&quot;符号"
			.Write "<br>XML 生成目录：<input class='textbox' name=""Setting(127)"" type=""text"" value=""" & Setting(127) & """>后面必须带&quot;/&quot;符号,生成XML文档时默认要存放的目录"
			.Write "</td>"
            .Write "</tr>"
		    .Write "     <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""25%"" height=""30""align=""right"" class='clefttitle'><div><strong>上传文件存放目录格式：</strong></div><span class='tips'></span></td>"
			.Write "       <td height=""30""> "
			.Write "         <input type=""radio"" name=""Setting(96)"" value=""3"" "
			If Setting(96) = "3" Then .Write (" checked")
			.Write " >总上传目录/年/管理员ID<br/>"
			.Write "<input type=""radio"" name=""Setting(96)"" value=""1"" "
			If Setting(96) = "1" Then .Write (" checked")
			.Write " >总上传目录/年-月/管理员ID<br/>"
			.Write "         <input type=""radio"" name=""Setting(96)"" value=""2"" "
			If Setting(96) = "2" Then .Write (" checked")
			.Write " >总上传目录/年-月-日/管理员ID<br/>"
			.Write "         <input type=""radio"" name=""Setting(96)"" value=""4"" "
			If Setting(96) = "4" Then .Write (" checked")
			.Write " >总上传目录/管理员ID"
			.Write "    </td> </tr>"

		    .Write "     <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""25%"" height=""30""align=""right"" class='clefttitle'><div><strong>会员投稿是否允许自动远程存图：</strong></div><span class='tips'>若选择<font color=red>""""</font>涉及到远程引用远程图片的地方将自动将图片保存到您网站上。</span></td>"
			.Write "       <td height=""30""> <input type=""radio"" name=""Setting(92)"" value=""1"" "
			If Setting(92) = "1" Then .Write (" checked")
			.Write " >"
			.Write "         允许"
			.Write "         <input type=""radio"" name=""Setting(92)"" value=""0"" "
			If Setting(92) = "0" Then .Write (" checked")
			.Write " >"
			.Write "         不允许</td>"
			.Write "     </tr>"
		    .Write "     <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""25%"" height=""30""align=""right"" class='clefttitle'><div><strong>远程保存的图片加水印：</strong></div><span class='tips'>若选择<font color=red>""是""</font>涉及到远程存图的地方将自动加水印,如采集或是文章里的自动存图等。</span></td>"
			.Write "       <td height=""30""> <input type=""radio"" name=""Setting(174)"" value=""1"" "
			If Setting(174) = "1" Then .Write (" checked")
			.Write " >"
			.Write "         是"
			.Write "         <input type=""radio"" name=""Setting(174)"" value=""0"" "
			If Setting(174) = "0" Then .Write (" checked")
			.Write " >"
			.Write "         否</td>"
			.Write "     </tr>"
			.Write "     <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td height=""30"" align=""right"" class='clefttitle'> <div><strong>生成方式：</strong></div><span class='tips'>若您有绑定子站点,此处请设置为绝对路径，否则可能导致链接不正确。</span></td>"
			.Write "       <td height=""30""> <input name=""Setting(97)"" type=""radio"" value=""1"""
			If Setting(97) = "1" Then .Write (" checked")
			.Write " >"
			.Write "         绝对路径"
			.Write "         <input type=""radio"" name=""Setting(97)"" value=""0"""
			If Setting(97) = "0" Then .Write (" checked")
			.Write " >"
			.Write "         根相对路径 (相对根目录)</td>"
			.Write "     </tr>"

			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""25%"" height=""25"" class=""clefttitle"" align=""right""> <div><strong>地图默认经纬坐标：</strong></div>"
			.Write "         </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <input class='textbox' id=""mapcenter"" name=""Setting(176)"" type=""text"" value=""" & Setting(176) & """ size=""20"">  电子地图默认显示的中心坐标 <input  type='button' value='获取中心坐标' class='button' onclick='addMap()'/><br/><span class='tips'>当商家没有标注时将默认显示这里设置的中心坐标位置。</span>"
			%>
		   <script src='../ks_inc/kesion.box.js'></script>
			<script>
		  function addMap(){
			  new parent.KesionPopup().PopupCenterIframe('获取中心坐标','../plus/baidumap.asp?obj=parent.frames["MainFrame"].document.getElementById("mapcenter")&action=getcenter&MapMark='+$("#mapcenter").val(),730,450,'auto');
		  }
		  </script>
			<%
			.Write " </td>"
			.Write "</tr>"
	        .Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""25%"" height=""25"" class=""clefttitle"" align=""right""> <strong>电子地图调用方法：</strong>"
			.Write "  <br/<span class='tips'>请在右边代码复制放到您需要用到电子地图的内容页即可。</span></td>"
			.Write "       <td height=""21"" valign=""middle"">"
			%>
			<textarea cols="50" rows="10">
<!--电子地图开始--->
<script src="http://api.map.baidu.com/api?v=1.1&services=true" type="text/javascript"></script>
<div style="width:700px;height:340px;border:1px solid gray" id="container"></div>

<script type="text/javascript"> 
	var map = new BMap.Map("container");          // 创建Map实例
	var point = new BMap.Point({$MapCenterPoint});  // 创建点坐标
	map.centerAndZoom(point,16);                  // 初始化地图，设置中心点坐标和地图级别。
	map.addControl(new BMap.NavigationControl());   
	map.addControl(new BMap.ScaleControl());   
	map.addControl(new BMap.OverviewMapControl()); 
	var sContent ="<h4 style='margin:0 0 5px 0;padding:0.2em 0'>地址：{$FL_Title}</h4>" +"<p style='margin:0;line-height:1.5;font-size:13px;'>电话：{$KS_tel} </p>"
	{$ShowMarkerList}
	window.setTimeout(function(){map.panTo(new BMap.Point({$MapCenterPoint}));}, 2000);
	
	function addMarker(point, index){   
	  // 创建图标对象   
	  var myIcon = new BMap.Icon("http://api.map.baidu.com/img/markers.png", new BMap.Size(23, 25), {   
		offset: new BMap.Size(10, 25),                  // 指定定位位置   
		imageOffset: new BMap.Size(0, 0 - index * 25)   // 设置图片偏移   
	  });   
	  var marker = new BMap.Marker(point, {icon: myIcon});   
	  map.addOverlay(marker);  
	  
	  if (index==0){
		var infoWindow = new BMap.InfoWindow(sContent);  // 创建信息窗口对象
		 marker.addEventListener("click", function(){										
		   this.openInfoWindow(infoWindow);	}); 
		map.openInfoWindow(infoWindow, map.getCenter());      // 打开信息窗口 
	  }
	}  
</script>
<!--电子地图结束--->
</textarea>

			
			<%
			.Write "  </td>"
			.Write "</tr>"
			
			
			
			.Write "     <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td height=""25"" align=""right"" class='clefttitle'><div><strong>是否启用软键盘输入密码：</strong></div><span class='tips'>若设置为<font color=""#FF0000"">&quot;启用&quot;</font>，则管理员登陆后台时使用软键盘输入密码，适合网吧等场所上网使用。</span></td>"
			.Write "       <td height=""21"" valign=""middle""><input type=""radio"" name=""Setting(98)"" value=""1"""
			If Setting(98) = "1" Then .Write (" Checked")
			.Write " >"
			.Write "         启用"
			.Write "         <input type=""radio"" name=""Setting(98)"" value=""0"""
			If Setting(98) = "0" Then .Write (" Checked")
			.Write " >"
			.Write "         不启用</td>"
		    .Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""25%"" height=""25"" class=""clefttitle"" align=""right""> <div><strong>FSO组件的名称：</strong></div><span class='tips'>某些网站为了安全，将FSO组件的名称进行更改以达到禁用FSO的目的。如果你的网站是这样做的，请在此输入更改过的名称。</span>"
			.Write "         </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <input class='textbox' name=""Setting(99)"" type=""text"" value=""" & Setting(99) & """ size=""50"">      </td>"
			.Write "</tr>"
					Dim LockIPStr:LockIPStr=Setting(101)
			If LockIPStr<>"" And Not IsNull(LockIPStr) Then
				LockIPWhite=Split(LockIPStr,"|||")(0)
				LockIPBlack=Split(LockIPStr,"|||")(1)
				Dim IPWhiteStr,IPBlackStr,IPWhite,IPBlack
				Dim M,N,IPWA:IPWA=Split(LockIPWhite,"$$$")
				For M=0 To Ubound(IPWA)
					LockIPWhiteArr=Split(IPWA(M),"----")
					For N=0 To Ubound(LockIPWhiteArr)
					 If N=0 Then
					 IPWhite=KS.CStrIP(LockIPWhiteArr(N))
					 Else
					 IPWhite=IPWhite & "----" & KS.CStrIP(LockIPWhiteArr(N))
					 End If
					Next
					If M=0 Then
					 IPWhiteStr=IPWhite
					Else
					 IPWhiteStr=IPWhiteStr & vbcrlf & IPWhite
					End If
				Next
				IPWA=Split(LockIPBlack,"$$$")
				For M=0 To Ubound(IPWA)
					LockIPBlackArr=Split(IPWA(M),"----")
					For N=0 To Ubound(LockIPBlackArr)
					 If N=0 Then
					 IPBlack=KS.CStrIP(LockIPBlackArr(N))
					 Else
					 IPBlack=IPBlack & "----" & KS.CStrIP(LockIPBlackArr(N))
					 End If
					Next
					If M=0 Then
					 IPBlackStr=IPBlack
					Else
					 IPBlackStr=IPBlackStr & vbcrlf & IPBlack
					End If
				Next
			End If
			
		 .Write "<tbody style='display:none'>"
		 .Write "<tr class='tdbg' class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		 .Write " <td width='40%' class='clefttitle' align='right'><strong>来访限定方式：</strong><br><font color='red'>此功能只对ASP访问方式有效。如果你以前生成了HTML文件，则启用此功能后，这些HTML文件仍可以访问（除非手工删除）。</font></td>"
		 .Write " <td><input name='Setting(100)' type='radio' value='0'"
		 if Setting(100)="0" then .write " checked"
		 .Write ">  不启用，任何IP都可以访问本站。<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='1'"
		 if Setting(100)="1" then .write " checked"
		 .Write ">  仅启用白名单，只允许白名单中的IP访问本站。<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='2'"
		 if Setting(100)="2" then .write " checked"
		 .Write ">  仅启用黑名单，只禁止黑名单中的IP访问本站。<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='3'"
		 if Setting(100)="3" then .write " checked"
		 .Write ">  同时启用白名单与黑名单，先判断IP是否在白名单中，如果不在，则禁止访问；如果在则再判断是否在黑名单中，如果IP在黑名单中则禁止访问，否则允许访问。<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='4'"
		 if Setting(100)="4" then .write " checked"
		 .Write ">  同时启用白名单与黑名单，先判断IP是否在黑名单中，如果不在，则允许访问；如果在则再判断是否在白名单中，如果IP在白名单中则允许访问，否则禁止访问。</td>"
		.Write "</tr>"
	    .Write "<tr class='tdbg' class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"     
		.Write " <td width='40%' class='clefttitle' align='right'><strong>IP段白名单</strong>：<br> (注：添加多个限定IP段，请用<font color='red'>回车</font>分隔。 <br>限制IP段的书写方式，中间请用英文四个小横杠连接，如<font color='red'>202.101.100.32----202.101.100.255</font> 就限定了IP 202.101.100.32 到IP 202.101.100.255这个IP段的访问。当页面为asp方式时才有效。) </td> "     
		.Write " <td><textarea name='LockIPWhite' cols='60' rows='8'>" & IPWhiteStr & "</textarea></td>"
		.Write "</tr>"
	    .Write "<tr class='tdbg' class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"     
		.Write " <td width='40%' class='clefttitle' align='right'><strong>IP段黑名单</strong>：<br> (注：同上。) <br></td>"      
		.Write "<td> <textarea name='LockIPBlack' cols='60' rows='8'>" & IPBlackStr & "</textarea></td>"
		.Write "</tr>"
		.write "</tbody>"
			.Write "   </table>"
			.Write " </div>"
			
			on error resume next
			.Write" <div class=tab-page id=SMS_Option>"
			.Write "  <H2 class=tab  style='display:none'>短信平台</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage(document.getElementById( ""SMS_Option"" ));"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" Class=""CTable"">"
			.Write "     <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td height=""25"" align=""right"" class='clefttitle'><div><strong>是否启用短信功能：</strong></div>若设置为<font color=""#FF0000"">&quot;启用&quot;</font>，则用户注册成功或在线支付成功将自动发送手机短信通知用户。</td>"
			.Write "       <td height=""21"" valign=""middle""><input type=""radio"" name=""Setting(157)"" value=""1"""
			If Setting(157) = "1" Then .Write (" Checked")
			.Write " >"
			.Write "         启用"
			.Write "         <input type=""radio"" name=""Setting(157)"" value=""0"""
			If Setting(157) = "0" Then .Write (" Checked")
			.Write " >"
			.Write "         不启用</td>"
		    .Write "</tr>"
						
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""25%"" height=""25"" class=""clefttitle"" align=""right""> <div><strong>SCP服务器地址：</strong></div>填写SCP提供商的服务器地址。"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <input type=""text"" name=""Setting(150)"" size=""50"" value=""" & Setting(150) & """>    </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""25%"" height=""25"" class=""clefttitle"" align=""right""> <div><strong>SCP服务器接口：</strong></div>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <input type=""text"" name=""Setting(151)"" value=""" & Setting(151) & """>     </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""25%"" height=""25"" class=""clefttitle"" align=""right""> <div><strong>短信平台账号：</strong></div>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <input type=""text"" name=""Setting(152)"" value=""" & Setting(152) & """>      </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""25%"" height=""25"" class=""clefttitle"" align=""right""> <div><strong>短信平台密码：</strong></div>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <input type=""text"" name=""Setting(153)"" value=""" & Setting(153) & """>      </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""25%"" height=""25"" class=""clefttitle"" align=""right""> <div><strong>发送通道：</strong></div>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle"">"
			.Write  "<select name=""Setting(158)"">"
			.Write " <option value=""1"""
			If Setting(158)="1" Then .Write " selected"
			.Write " > 通道一 (发送1条扣去1条)</option>"
			.Write " <option value=""2"""
			If Setting(158)="2" Then .Write " selected"
			.Write "> 通道二 (发送1条扣去1条)</option>"
			.Write " <option value=""3"""
			If Setting(158)="3" Then .Write " selected"
			.Write "> 即时通道(客服类推荐) (发送1条扣去1.5条)</option>"
			.Write " <option value=""4"""
			If Setting(158)="4" Then .Write " selected"
			.Write "> 营销通道(营销类推荐) (发送1条扣去1.2条)</option>"
			.Write "</select>"
			.Write "   </td>"
			.Write "</tr>"
			
			
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""25%"" height=""25"" class=""clefttitle"" align=""right""> <div><strong>管理员的小灵通或手机号码：</strong></div>多个号码请用小写逗号隔开，如13600000000,15000000000。"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <textarea name=""Setting(154)"" cols=80 rows=4>" & Setting(154) & "</textarea>      </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""25%"" height=""25"" class=""clefttitle"" align=""right""> <div><strong>会员注册成功后发送的短消息：</strong></div>可用标签{$UserName},{$PassWord}。<br><font color=blue>说明：留空表示不发送</font>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <textarea name=""Setting(155)"" cols=80 rows=4>" & Setting(155) & "</textarea>      </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""25%"" height=""25"" class=""clefttitle"" align=""right""> <div><strong>在线支付完成后发送的短消息：</strong></div>可以用标签{$UserName},{$Money}。<br><font color=blue>说明：留空表示不发送</font>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <textarea name=""Setting(156)"" cols=80 rows=4>" & Setting(156) & "</textarea>      </td>"
			.Write "</tr>"
			.Write "   </table>"
			.Write "</div>"
			
			.Write " </form>"
		    .Write "</div>"

			.Write " </body>"
			.Write " </html>"
			.Write " <Script Language=""javascript"">"
			.Write " <!--" & vbCrLf
			'.Write " setlience("&Setting(22) &");"&vbcrlf
			.Write " setsendmail(" &Setting(146) & ");" & vbcrlf
			.Write "function setlience(n)" & vbcrlf
			.Write "{" & vbcrlf
			.Write "  if (n==0)"  &vbcrlf
			.Write "    document.all.liencearea.style.display='none';" & vbcrlf
			.Write "  else" & vbcrlf
			.Write "    document.all.liencearea.style.display=''; " & vbcrlf
			.Write "}" & vbcrlf
			.Write "function setsendmail(n)" & vbcrlf
			.Write "{" & vbcrlf
			.Write "  if (n==0)"  &vbcrlf
			.Write "    document.getElementById('sendmailarea').style.display='none';" & vbcrlf
			.Write "  else" & vbcrlf
			.Write "    document.getElementById('sendmailarea').style.display=''; " & vbcrlf
			.Write "}" & vbcrlf

			.Write " function CheckForm()" & vbCrLf
			.Write " {" & vbCrLf
			.Write "     $('#myform').submit();"
			.Write " }" & vbCrLf
			.Write " //-->" & vbCrLf
			.Write " </Script>" & vbCrLf
			RS.Close:Set RS = Nothing:Set Conn = Nothing
		End With
		End Sub
	
		
		'系统空间占用量
		Sub GetSpaceInfo()
			Dim SysPath, FSO, F, FC, I, I2
			Response.Write ("<title>空间查看</title>")
			Response.Write ("<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>")
			Response.Write ("<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>")
			Response.Write ("<BODY leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>")
			Response.Write ("<div class='topdashed'><a href='?action=CopyRight'><strong>服务器参数探测</strong></a> | <a href='?action=Space'><strong>系统空间占用量</strong></a></div>")

			
			Response.Write ("<table width='100%' border='0' cellspacing='0' cellpadding='0' oncontextmenu=""return false"">")
			Response.Write ("  <tr>")
			Response.Write ("    <td valign='top'>")
            Response.Write ("<br><br><table width=90% border=0 align='center' cellpadding=0 cellspacing=1 bgcolor='#CDCDCD'>")
     
         SysPath = Server.MapPath("\") & "\"
                 Set FSO = KS.InitialObject(KS.Setting(99))
                  Set F = FSO.GetFolder(SysPath)
                  Set FC = F.SubFolders
                            I = 1
                            I2 = 1
               For Each F In FC
				Response.Write ("        <tr>")
				Response.Write ("          <td height=25 bgcolor='#EEF8FE'><img src='Images/Folder/folderclosed.gif' width='20' height='20' align='absmiddle'><b>" & F.name & "</b>&nbsp; 占用空间：&nbsp;<img src='../Images/default/bar.gif' width=" & Drawbar(F.name) & " height=10>&nbsp;")
					ShowSpaceInfo (F.name)
				Response.Write ("          </td>")
				Response.Write ("        </tr>")
							  I = I + 1
								  If I2 < 10 Then
									I2 = I2 + 1
								  Else
									I2 = 1
								 End If
								 Next
						  
				Response.Write ("        <tr>")
				Response.Write ("          <td height='25' bgcolor='#EEF8FE'> 程序文件占用空间：&nbsp;<img src='../Images/default/' width=" & Drawspecialbar & " height=10>&nbsp;")
				
				Showspecialspaceinfo ("Program")
				
				Response.Write ("          </td>")
				Response.Write ("        </tr>")
				Response.Write ("      </table>")
				Response.Write ("      <table width=90% border=0 align='center' cellpadding=3 cellspacing=1>")
				Response.Write ("        <tr>")
				Response.Write ("          <td height='28' align='right' bgcolor='#FFFFFF'><font color='#FF0066'><strong><font color='#006666'>系统占用空间总计：</font></strong>")
				Showspecialspaceinfo ("All")
				Response.Write ("            </font> </td>")
				Response.Write ("        </tr>")
				Response.Write ("      </table></td>")
				Response.Write ("  </tr>")
				Response.Write ("</table>")

				Response.Write ("</body>")
				Response.Write ("</html>")
		End Sub
		Sub ShowSpaceInfo(drvpath)
        Dim FSO, d, size, showsize
        Set FSO = KS.InitialObject(KS.Setting(99))
        Set d = FSO.GetFolder(Server.MapPath("/" & drvpath))
        size = d.size
        showsize = size & "&nbsp;Byte"
        If size > 1024 Then
           size = (size / 1024)
           showsize = round(size,2) & "&nbsp;KB"
        End If
        If size > 1024 Then
           size = (size / 1024)
           showsize = round(size,2) & "&nbsp;MB"
        End If
        If size > 1024 Then
           size = (size / 1024)
           showsize = round(size,2) & "&nbsp;GB"
        End If
        Response.Write "<font face=verdana>" & showsize & "</font>"
      End Sub
	  Sub Showspecialspaceinfo(method)
			Dim FSO, d, FC, f1, size, showsize, drvpath
			Set FSO = KS.InitialObject(KS.Setting(99))
			Set d = FSO.GetFolder(Server.MapPath("/"))
			 If method = "All" Then
				size = d.size
			ElseIf method = "Program" Then
				Set FC = d.Files
				For Each f1 In FC
					size = size + f1.size
				Next
			End If
			showsize = round(size,2) & "&nbsp;Byte"
			If size > 1024 Then
			   size = (size / 1024)
			   showsize = round(size,2) & "&nbsp;KB"
			End If
			If size > 1024 Then
			   size = (size / 1024)
			   showsize = round(size,2) & "&nbsp;MB"
			End If
			If size > 1024 Then
			   size = (size / 1024)
			   showsize = round(size,2) & "&nbsp;GB"
			End If
			Response.Write "<font face=verdana>" & showsize & "</font>"
		End Sub
		Function Drawbar(drvpath)
			Dim FSO, drvpathroot, d, size, totalsize, barsize
			Set FSO = KS.InitialObject(KS.Setting(99))
			Set d = FSO.GetFolder(Server.MapPath("/"))
			totalsize = d.size
			Set d = FSO.GetFolder(Server.MapPath("/" & drvpath))
			size = d.size
			
			barsize = CInt((size / totalsize) * 100)
			Drawbar = barsize
		End Function
		Function Drawspecialbar()
			Dim FSO, drvpathroot, d, FC, f1, size, totalsize, barsize
			Set FSO = KS.InitialObject(KS.Setting(99))
			Set d = FSO.GetFolder(Server.MapPath("/"))
			totalsize = d.size
			Set FC = d.Files
			For Each f1 In FC
				size = size + f1.size
			Next
			barsize = CInt((size / totalsize) * 100)
			Drawspecialbar = barsize
		End Function

       '查看组件支持情况
	   Sub GetDllInfo()
	    Dim theInstalledObjects(17)
	   	theInstalledObjects(0) = "MSWC.AdRotator"
		theInstalledObjects(1) = "MSWC.BrowserType"
		theInstalledObjects(2) = "MSWC.NextLink"
		theInstalledObjects(3) = "MSWC.Tools"
		theInstalledObjects(4) = "MSWC.Status"
		theInstalledObjects(5) = "MSWC.Counters"
		theInstalledObjects(6) = "IISSample.ContentRotator"
		theInstalledObjects(7) = "IISSample.PageCounter"
		theInstalledObjects(8) = "MSWC.PermissionChecker"
		theInstalledObjects(9) = KS.Setting(99)
		theInstalledObjects(10) = "adodb.connection"
		theInstalledObjects(11) = "SoftArtisans.FileUp"
		theInstalledObjects(12) = "SoftArtisans.FileManager"
		theInstalledObjects(13) = "JMail.SMTPMail"
		theInstalledObjects(14) = "CDONTS.NewMail"
		theInstalledObjects(15) = "Persits.MailSender"
		theInstalledObjects(16) = "LyfUpload.UploadFile"
		theInstalledObjects(17) = "Persits.Upload.1"


		 Response.Write ("<table width='699' border='0' align='center' cellpadding='0' cellspacing='0' bgcolor='#CDCDCD'>")
		 Response.Write ("   <form method='post' action='?Action=CopyRight'>")
		 Response.Write ("<tr>")
		 Response.Write ("     <td height=36 bgcolor='#FFFFFF'>服务器组件探测查询-&gt; <font color='#FF0000'>组件名称:</font>")
		 Response.Write ("       <input type='text' name='classname' class='textbox' style='width:180'>")
		 Response.Write ("     <input type='submit' name='Submit' class='button' value='测 试'>")
			 
		Dim strClass:strClass = Trim(Request.Form("classname"))
		If "" <> strClass Then
		Response.Write "<br>您指定的组件的检查结果："
		If Not IsObjInstalled(strClass) Then
		Response.Write "<br><font color=red>很遗憾，该服务器不支持" & strClass & "组件！</font>"
		Else
		Response.Write "<br><font color=green>恭喜！该服务器支持" & strClass & "组件。</font>"
		End If
		Response.Write "<br>"
		End If
		Response.Write ("</font>")
		Response.Write ("      </td>")
		Response.Write ("  </tr></form>")
		Response.Write (" <tr>")
		Response.Write ("    <td height=25 bgcolor='#FFFFFF'><b><font color='#006666'> 　IIS自带组件</font></b></font></td>")
		Response.Write ("  </tr>")
		Response.Write ("  <tr>")
		Response.Write ("    <td height=20 bgcolor='#EEF8FE'>")
		Response.Write ("      <table width='100%' border=0 align='center' cellpadding=0 cellspacing=1 bgcolor='#CDCDCD'>")
		Response.Write ("        <tr align=center bgcolor='#EEF8FE' height=22>")
		Response.Write ("          <td width='70%'>组 件 名 称</td>")
		Response.Write ("          <td width='15%'>支 持</td>")
		Response.Write ("          <td width='15%'>不支持</td>")
		Response.Write ("        </tr>")
			  
		Dim I
		For I = 0 To 10
		Response.Write "<TR align=center bgcolor=""#EEF8FE"" height=22><TD align=left>&nbsp;" & theInstalledObjects(I) & "<font color=#888888>&nbsp;"
		Select Case I
		Case 9
		Response.Write "(FSO 文本文件读写)"
		Case 10
		Response.Write "(ACCESS 数据库)"
		End Select
		Response.Write "</font></td>"
		If Not IsObjInstalled(theInstalledObjects(I)) Then
		Response.Write "<td></td><td><font color=red><b>×</b></font></td>"
		Else
		Response.Write "<td><b>√</b></td><td></td>"
		End If
		Response.Write "</TR>" & vbCrLf
		Next
		
		Response.Write ("      </table></td>")
		Response.Write ("  </tr>")
		Response.Write ("  <tr>")
		Response.Write ("    <td height=25 bgcolor='#FFFFFF'> <font color='#006666'><b>　其他常见组件</b></font>")
		Response.Write ("    </td>")
		Response.Write ("  </tr>")
		Response.Write ("  <tr>")
		Response.Write ("    <td height=20 bgcolor='#EEF8FE'>")
		Response.Write ("      <table width='100%' border=0 align='center' cellpadding=0 cellspacing=1 bgcolor='#CDCDCD'>")
		Response.Write ("        <tr align=center bgcolor='#EEF8FE' height=22>")
		Response.Write ("          <td width='70%'>组 件 名 称</td>")
		Response.Write ("          <td width='15%'>支 持</td>")
		Response.Write ("          <td width='15%'>不支持</td>")
		Response.Write ("        </tr>")
			 
		For I = 11 To UBound(theInstalledObjects)
		Response.Write "<TR align=center height=18 bgcolor=""#EEF8FE""><TD align=left>&nbsp;" & theInstalledObjects(I) & "<font color=#888888>&nbsp;"
		Select Case I
		Case 11
		Response.Write "(SA-FileUp 文件上传)"
		Case 12
		Response.Write "(SA-FM 文件管理)"
		Case 13
		Response.Write "(JMail 邮件发送)"
		Case 14
		Response.Write "(CDONTS 邮件发送 SMTP Service)"
		Case 15
		Response.Write "(ASPEmail 邮件发送)"
		Case 16
		Response.Write "(LyfUpload 文件上传)"
		Case 17
		Response.Write "(ASPUpload 文件上传)"
		End Select
		Response.Write "</font></td>"
		If Not IsObjInstalled(theInstalledObjects(I)) Then
		Response.Write "<td></td><td><font color=red><b>×</b></font></td>"
		Else
		Response.Write "<td><b>√</b></td><td></td>"
		End If
		Response.Write "</TR>" & vbCrLf
		Next
		
		Response.Write ("      </table></td>")
		Response.Write ("  </tr>")
		Response.Write ("</table>")
		Response.Write ("</td>")
		Response.Write ("</tr>")
		Response.Write ("</table>")
		End Sub
		
		'系统版权及服务器参数测试
		Sub GetCopyRightInfo()
	%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../KS_Inc/common.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div class='topdashed'> <a href="?action=CopyRight"><strong>服务器参数探测</strong></a> | <a href="?action=Space"><strong>系统空间占用量</strong></a></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="699" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width=1 bgcolor="#E3E3E3"></td>
          <td width="1011"><div align="left"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <%
				Dim theInstalledObjects(23)
				theInstalledObjects(0) = "MSWC.AdRotator"
				theInstalledObjects(1) = "MSWC.BrowserType"
				theInstalledObjects(2) = "MSWC.NextLink"
				theInstalledObjects(3) = "MSWC.Tools"
				theInstalledObjects(4) = "MSWC.Status"
				theInstalledObjects(5) = "MSWC.Counters"
				theInstalledObjects(6) = "IISSample.ContentRotator"
				theInstalledObjects(7) = "IISSample.PageCounter"
				theInstalledObjects(8) = "MSWC.PermissionChecker"
				theInstalledObjects(9) = KS.Setting(99)
				theInstalledObjects(10) = "adodb.connection"
					
				theInstalledObjects(11) = "SoftArtisans.FileUp"
				theInstalledObjects(12) = "SoftArtisans.FileManager"
				theInstalledObjects(13) = "JMail.SMTPMail"
				theInstalledObjects(14) = "CDONTS.NewMail"
				theInstalledObjects(15) = "Persits.MailSender"
				theInstalledObjects(16) = "LyfUpload.UploadFile"
				theInstalledObjects(17) = "Persits.Upload.1"
				theInstalledObjects(18) = "CreatePreviewImage.cGvbox"	'CreatePreviewImage
				theInstalledObjects(19)	= "Persits.Jpeg"				'AspJpeg
				theInstalledObjects(20) = "SoftArtisans.ImageGen"		'SoftArtisans ImgWriter V1.21
				theInstalledObjects(21) = "sjCatSoft.Thumbnail"
				theInstalledObjects(22) = "Microsoft.XMLHTTP"
				theInstalledObjects(23) = "Adodb.Stream"
	%>      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="700" height="30" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td>　<font color="#006666"><strong>使用本系统，请确认您的服务器和您的浏览器满足以下要求：</strong></font></td>
                </tr>
              </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="699" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="22">　<font face="Verdana, Arial, Helvetica, sans-serif">JRO.JetEngine</font><span class="small2">：</span> 
                  <%
	    On Error Resume Next
	    KS.InitialObject("JRO.JetEngine")
		if err=0 then 
		  Response.Write("<font color=#0076AE>√</font>")
		else
          Response.Write("<font color=red>×</font>")
		end if	 
		err=0
		Response.Write(" (ADO 数据对象):")
		 On Error Resume Next
	    KS.InitialObject("adodb.connection")
		if err=0 then 
		  Response.Write("<font color=#0076AE>√</font>")
		else
          Response.Write("<font color=red>×</font>")
		end if	 
		err=0
	  %>                  </td>
                  <td width="52%" height="22"> 　当前数据库　 
                  <%
		If DataBaseType = 1 Then
		Response.Write "<font color=#0076AE>MS SQL</font>"
		else
		Response.Write "<font color=#0076AE>ACCESS</font>"
		end if
	  %>                  </td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td height="22">　<span class="small2"><font face="Verdana, Arial, Helvetica, sans-serif">FSO</font></span>文本文件读写<span class="small2">：</span> 
                  <%
	    On Error Resume Next
	    KS.InitialObject(KS.Setting(99))
		if err=0 then 
		  Response.Write("<font color=#0076AE>支持√</font>")
		else
          Response.Write("<font color=red>不支持×</font>")
		end if	 
		err=0
	  %>                  </td>
                  <td height="22">　Microsoft.XMLHTTP 
                    <%If  Not IsObjInstalled(theInstalledObjects(22)) Then%>
                    <font color="red">×</font> 
                    <%else%>
                    <font color="0076AE"> √</font> 
                    <%end if%>
                    　Adodb.Stream 
                   <%If Not IsObjInstalled(theInstalledObjects(23)) Then%>
                    <font color="red">×</font> 
                    <%else%>
                    <font color="0076AE"> √</font> 
                    <%end if%>                  </td>
                </tr>
                
                <tr bgcolor="#EEF8FE"> 
                  <td height="22" colspan="2">　客户端浏览器版本： 
                    <%
	  Dim Agent,Browser,version,tmpstr
	  Agent=Request.ServerVariables("HTTP_USER_AGENT")
	  Agent=Split(Agent,";")
	  If InStr(Agent(1),"MSIE")>0 Then
				Browser="MS Internet Explorer "
				version=Trim(Left(Replace(Agent(1),"MSIE",""),6))
			ElseIf InStr(Agent(4),"Netscape")>0 Then 
				Browser="Netscape "
				tmpstr=Split(Agent(4),"/")
				version=tmpstr(UBound(tmpstr))
			ElseIf InStr(Agent(4),"rv:")>0 Then
				Browser="Mozilla "
				tmpstr=Split(Agent(4),":")
				version=tmpstr(UBound(tmpstr))
				If InStr(version,")") > 0 Then 
					tmpstr=Split(version,")")
					version=tmpstr(0)
				End If
			End If
	Response.Write(""&Browser&"  "&version&"")
	  %>
                    [需要IE5.5或以上,服务器建议采用Windows 2000或Windows 2003 Server]</td>
                </tr>
              </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="700" height="30" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td>　<font color="#006666"><strong>服务器信息</strong></font></td>
                </tr>
              </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="699" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
                <tr bgcolor="#EEF8FE"> 
                  <td height="25">　服务器类型：<font face="Verdana, Arial, Helvetica, sans-serif"><%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</font></td>
                  <td height="25">　<span class="small2"><font face="Verdana, Arial, Helvetica, sans-serif">WEB</font></span>服务器的名称和版本<font face="Verdana, Arial, Helvetica, sans-serif">：<font color=#0076AE><%=Request.ServerVariables("SERVER_SOFTWARE")%></font></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="25">　返回服务器的主机名，<font face="Verdana, Arial, Helvetica, sans-serif">IP</font>地址<font face="Verdana, Arial, Helvetica, sans-serif">：<font color=#0076AE><%=Request.ServerVariables("SERVER_NAME")%></font></font></td>
                  <td width="52%" height="25">　服务器操作系统<font face="Verdana, Arial, Helvetica, sans-serif">：<font color=#0076AE><%=Request.ServerVariables("OS")%></font></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="25">　站点物理路径<font face="Verdana, Arial, Helvetica, sans-serif">：<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></font></td>
                  <td width="52%" height="25">　虚拟路径<font face="Verdana, Arial, Helvetica, sans-serif">：<%=Request.ServerVariables("SCRIPT_NAME")%></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="25">　脚本超时时间<span class="small2">：</span><font color=#0076AE><%=Server.ScriptTimeout%></font> 秒</td>
                  <td width="52%" height="25">　脚本解释引擎<span class="small2">：</span><font face="Verdana, Arial, Helvetica, sans-serif"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %>　</font> </td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td height="25">　服务器端口<font face="Verdana, Arial, Helvetica, sans-serif">：<%=Request.ServerVariables("SERVER_PORT")%></font></td>
                  <td height="25">　协议的名称和版本<font face="Verdana, Arial, Helvetica, sans-serif">：<%=Request.ServerVariables("SERVER_PROTOCOL")%></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td height="25">　服务器 <font face="Verdana, Arial, Helvetica, sans-serif">CPU</font> 
                    数量<font face="Verdana, Arial, Helvetica, sans-serif">：<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%></font> 个　</td>
                  <td height="25">　客户端操作系统： 
                    <%
 dim thesoft,vOS
thesoft=Request.ServerVariables("HTTP_USER_AGENT")
if instr(thesoft,"Windows NT 5.0") then
	vOS="Windows 2000"
elseif instr(thesoft,"Windows NT 5.2") then
	vOs="Windows 2003"
elseif instr(thesoft,"Windows NT 5.1") then
	vOs="Windows XP"
elseif instr(thesoft,"Windows NT") then
	vOs="Windows NT"
elseif instr(thesoft,"Windows 9") then
	vOs="Windows 9x"
elseif instr(thesoft,"unix") or instr(thesoft,"linux") or instr(thesoft,"SunOS") or instr(thesoft,"BSD") then
	vOs="类Unix"
elseif instr(thesoft,"Mac") then
	vOs="Mac"
else
	vOs="Other"
end if
Response.Write(vOs)
%> </td>
                </tr>
              </table>
			  <%
			  GetDllInfo
			  %>
			   <table width="700" height="30" border="0" cellpadding="0" cellspacing="0" style="margin-left:198px;">
                <tr> 
                  <td>　<font color="#006666"><strong>系统版本信息</strong></font></td>
                </tr>
              </table>
             
              <table width="699" height="63" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
                <tr bgcolor="#EEF8FE"> 
                  <td height="30"> 　当前版本<font face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  <td height="30">　<strong><font color=red> 
                    <%=KS.Version%>
                    </font></strong></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="24%" height="30">　版权声明</td>
                  <td width="76%">　1、本软件为共享软件,提供个人网站免费使用,非漳州科兴技术有限公司官方授权许可，不得将之用于盈利或非盈利性的商业用途;<br>
                    　2、用户自由选择是否使用,在使用中出现任何问题和由此造成的一切损失漳州科兴技术有限公司将不承担任何责任;<br>
                    　3、本软件受中华人民共和国《着作权法》《计算机软件保护条例》等相关法律、法规保护，漳州科兴技术有限公司保留一切权利。　 
                    <p></p></td>
                </tr>
              </table>
              <br>
            </div></td>
        </tr>
      </table></td>
  </tr>
</table>
</html>
<%
		End Sub
		
		Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled = False
		Err = 0
		Dim xTestObj
		Set xTestObj = KS.InitialObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
		Set xTestObj = Nothing
		Err = 0
		End Function
		

End Class
%> 
