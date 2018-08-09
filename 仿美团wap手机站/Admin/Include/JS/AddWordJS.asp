<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="../Label/LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New AddWordJS
KSCls.Kesion()
Set KSCls = Nothing

Class AddWordJS
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		'主体部分
		Public Sub Kesion()
		Dim TempClassList, InstallDir, CurrPath, JSConfig, JSConfigArr, Action, JSID, Str, Descript, FolderID
		Dim JSFileName, WordCss, OpenType, ArticleListNumber, RowHeight, TitleLen, ContentLen, ColNumber, NavType, Navi, MoreLinkType, MoreLink, SplitPic, DateRule, DateAlign, TitleCss, DateCss, ContentCss, BGCss
		
		CurrPath = KS.GetCommonUpFilesDir()
		
		'判断是否编辑
		JSID = Trim(Request.QueryString("JSID"))
		FolderID = Trim(Request.QueryString("FolderID"))
		If JSID = "" Then
		  Action = "Add"
		Else
		  Action = "Edit"
		  Dim JSRS, JSName
		  Set JSRS = Server.CreateObject("Adodb.Recordset")
		  JSRS.Open "Select * From KS_JSFile Where JSID='" & JSID & "'", Conn, 1, 1
		  If JSRS.EOF And JSRS.BOF Then
			 JSRS.Close
			 Set JSRS = Nothing
			 Response.Write ("<Script>alert('参数传递出错!');history.back();</Script>")
			 Response.End
		  End If
			FolderID = JSRS("FolderID")
			JSName = Replace(Replace(JSRS("JSName"), "{JS_", ""), "}", "")
			JSFileName = Trim(JSRS("JSFileName"))
			JSID = JSRS("JSID")
			Descript = Trim(JSRS("Description"))
			JSConfig = Trim(JSRS("JSConfig"))
			JSRS.Close
			Set JSRS = Nothing
			JSConfig = Replace(JSConfig, """", "") '注:去除左右双引号"
			JSConfigArr = Split(JSConfig, ",")
			WordCss = JSConfigArr(1)
			ColNumber = JSConfigArr(2)
			OpenType = JSConfigArr(3)
			ArticleListNumber = JSConfigArr(4)
			RowHeight = JSConfigArr(5)
			TitleLen = JSConfigArr(6)
			ContentLen = JSConfigArr(7)
			NavType = JSConfigArr(8)
			Navi = JSConfigArr(9)
			MoreLinkType = JSConfigArr(10)
			MoreLink = JSConfigArr(11)
			SplitPic = JSConfigArr(12)
			DateRule = JSConfigArr(13)
			DateAlign = JSConfigArr(14)
			TitleCss = JSConfigArr(15)
			DateCss = JSConfigArr(16)
			ContentCss = JSConfigArr(17)
			BGCss = JSConfigArr(18)
		End If
		If WordCss = "" Then WordCss = "A"
		If ArticleListNumber = "" Then ArticleListNumber = 5
		If RowHeight = "" Then RowHeight = 20
		If TitleLen = "" Then TitleLen = 20
		If ContentLen = "" Then ContentLen = 50
		If ColNumber = "" Then ColNumber = 1
		
		Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
		Response.Write "<script src=""../../../ks_inc/jquery.js"" language=""JavaScript""></script>"
		Response.Write "<link href=""../Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		%>
		<script language="javascript">
		function SelectPicStyle(ObjValue)
		{
				document.all.ViewStylePicArea.innerHTML='<img src="../../Images/View/'+ObjValue+'.gif" border="0">';
		}
		function SetNavStatus()
		{
		  if (document.all.NavType.value==0)
		   {document.all.NavWord.style.display="";
			document.all.NavPic.style.display="none";}
		  else
		  {
		   document.all.NavWord.style.display="none";
		   document.all.NavPic.style.display="";}
		}
		function SetMoreLinkStatus()
		{
		if (document.all.MoreLinkType.value==0)
		   {document.all.LinkWord.style.display="";
			document.all.LinkPic.style.display="none";}
		  else
		  {
		   document.all.LinkWord.style.display="none";
		   document.all.LinkPic.style.display="";}
		}
		function CheckForm()
		{   
			if (document.myform.JSName.value=='')
			 {
			  alert('请输入JS名称');
			  document.myform.JSName.focus(); 
			  return false
			  }
			  if (document.myform.JSFileName.value=='')
			  {
			   alert('请输入JS文件名');
			  document.myform.JSFileName.focus(); 
			  return false
			  }
			 if (CheckEnglishStr(document.myform.JSFileName,"JS文件名")==false) 
			   return false;
			 if (!IsExt(document.myform.JSFileName.value,'JS'))
			   { alert('JS文件名的扩展名必须是.js');
				  document.myform.JSFileName.focus(); 
				  return false;
			   }
			var WordCss='"'+document.myform.WordCss.value+'"';
			var NavType=1;
			var ColNumber=document.myform.ColNumber.value;
			var OpenType='"'+document.myform.OpenType.value+'"';
			var ArticleListNumber=document.myform.ArticleListNumber.value;
			var RowHeight=document.myform.RowHeight.value;
			var TitleLen=document.myform.TitleLen.value;
			var ContentLen=document.myform.ContentLen.value;
			var Nav,NavType=document.myform.NavType.value;
			var MoreLink,MoreLinkType=document.myform.MoreLinkType.value;
			var SplitPic='"'+document.myform.SplitPic.value+'"';
			var DateRule=document.myform.DateRule.value;
			var DateAlign='"'+document.myform.DateAlign.value+'"';
			var TitleCss='"'+document.myform.TitleCss.value+'"';
			var DateCss='"'+document.myform.DateCss.value+'"';
			var ContentCss='"'+document.myform.ContentCss.value+'"';
			var BGCss='"'+document.myform.BGCss.value+'"';
		
			if  (ArticleListNumber=='')  ArticleListNumber=5;
			if (RowHeight=='') RowHeight=20
			if  (TitleLen=='') TitleLen=30;
			if  (ColNumber=='') ColNumber=1;
			if  (NavType==0) Nav='"'+document.myform.TxtNavi.value+'"'
			 else  Nav='"'+document.myform.NaviPic.value+'"';
			if  (MoreLinkType==0) MoreLink='"'+document.myform.MoreLinkWord.value+'"'
			else  MoreLink='"'+document.myform.MoreLinkPic.value+'"';
			document.myform.JSConfig.value=	'GetWordJS,'+WordCss+','+ColNumber+','+OpenType+','+ArticleListNumber+','+RowHeight+','+TitleLen+','+ContentLen+','+NavType+','+Nav+','+MoreLinkType+','+MoreLink+','+SplitPic+','+DateRule+','+DateAlign+','+TitleCss+','+DateCss+','+ContentCss+','+BGCss;
			document.myform.submit();
		}
		</script>
		<%
		Response.Write "</head>"
		Response.Write "<body topmargin=""0"" leftmargin=""0"">"
		Response.Write "<div align=""center"">"
		Response.Write "<form  method=""post"" name=""myform"" action=""AddJSSave.asp"">"
		Response.Write " <input type=""hidden"" name=""JSConfig"">"
		Response.Write " <input type=""hidden"" name=""JSType"" value=""1"">"
		Response.Write " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		Response.Write " <input type=""hidden"" name=""Page"" value=""" & Request("Page") & """>"
		Response.Write "  <input type=""hidden"" name=""JSID"" value=""" & JSID & """>"
		Response.Write " <input type=""hidden"" name=""FileUrl"" value=""AddWordJS.asp"">"
		Response.Write ReturnJSInfo(JSID, JSName, JSFileName, FolderID, 3, Descript)
		Response.Write "<br>"
		Response.Write "    <table width=""96%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		Response.Write "      <tr>"
		Response.Write "        <td> <FIELDSET align=center>"
		Response.Write "          <LEGEND align=left>自由文字JS属性设置</LEGEND>"
		Response.Write "          <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		Response.Write "            <tr>"
		Response.Write "            <td width=""69%""><table width=""96%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "                  <tr>"
		Response.Write "                    <td width=""50%"" style='text-align:left' height=""28"">文字样式"
		Response.Write "<select name=""WordCss"" class=""textbox"" id=""WordCss"" onchange=""SelectPicStyle(this.value)"">"
							   Dim SelStr
								   If WordCss = "A" Then
								   SelStr = " Selected"
								   Else
								   SelStr = ""
								   End If
								 Response.Write ("<option value=""A""" & SelStr & ">样式A</option>")
								If WordCss = "B" Then
								   SelStr = " Selected"
								   Else
								   SelStr = ""
								   End If
								 Response.Write ("<option value=""B""" & SelStr & ">样式B</option>")
								 If WordCss = "C" Then
								   SelStr = " Selected"
								   Else
								   SelStr = ""
								   End If
								Response.Write ("<option value=""C""" & SelStr & ">样式C</option>")
								If WordCss = "D" Then
								   SelStr = " Selected"
								   Else
								   SelStr = ""
								   End If
								 Response.Write ("<option value=""D""" & SelStr & ">样式D</option>")
								 If WordCss = "E" Then
								   SelStr = " Selected"
								   Else
								   SelStr = ""
								   End If
								 Response.Write ("<option value=""E""" & SelStr & ">样式E</option>")
							  
		Response.Write "                      </select>"
		Response.Write "                      并排列数"
		 
		 Response.Write "                     <input type=""text"" class=""textbox"" onBlur=""CheckNumber(this,'排列列数');""  style=""width:30px;"" value=""" & ColNumber & """ name=""ColNumber"">"
		 Response.Write "                   </td>"
		 Response.Write "                   <td width=""50%"" nowrap height=""28"">"
		 
		Response.Write ReturnOpenTypeStr(OpenType)
		
		Response.Write "       </td>"
		Response.Write "                  </tr>"
		Response.Write "                  <tr>"
		Response.Write "                    <td width=""50%"" style='text-align:left' height=""28"">文章数量"
		Response.Write "                      <input name=""ArticleListNumber"" class=""textbox"" type=""text"" id=""ArticleListNumber""    style=""width:20%;"" onBlur=""CheckNumber(this,'文章数量');"" value=""" & ArticleListNumber & """>"
		Response.Write "                      取<font color=""#FF0000"">0</font>时,将列出全部文章</td>"
		Response.Write "                    <td width=""50%"" height=""28"">文章行距"
		Response.Write "                      <input name=""RowHeight"" class=""textbox"" type=""text"" id=""RowHeight""    style=""width:70%;"" onBlur=""CheckNumber(this,'文章行距');"" value=""" & RowHeight & """></td>"
		Response.Write "                  </tr>"
		Response.Write "                  <tr>"
		Response.Write "                    <td width=""50%"" style='text-align:left' height=""28"">标题字数"
		Response.Write "                      <input name=""TitleLen"" class=""textbox"" onBlur=""CheckNumber(this,'标题字数');"" type=""text""    style=""width:70%;"" value=""" & TitleLen & """>"
		Response.Write "                    </td>"
		Response.Write "                    <td width=""50%"" height=""28"">内容字数"
		Response.Write "                      <input name=""ContentLen"" class=""textbox"" type=""text"" id=""ContentLen""    style=""width:70%;"" onBlur=""CheckNumber(this,'内容字数');"" value=""" & ContentLen & """></td>"
		Response.Write "                  </tr>"
		Response.Write "                  <tr>"
		Response.Write "                    <td width=""50%"" style='text-align:left' height=""28"">导航类型"
		Response.Write "                      <select name=""NavType"" class=""textbox"" style=""width:70%;"" onchange=""SetNavStatus()"">"
					
					If JSID = "" Or CStr(NavType) = "0" Then
					Response.Write ("<option value=""0"" selected>文字导航</option>")
					Response.Write ("<option value=""1"">图片导航</option>")
				   Else
					Response.Write ("<option value=""0"">文字导航</option>")
					Response.Write ("<option value=""1"" selected>图片导航</option>")
				   End If
		Response.Write "                      </select></td>"
		Response.Write "                    <td width=""50%"" height=""28"">"
				 If JSID = "" Or CStr(NavType) = "0" Then
				  Response.Write ("<div align=""left"" id=""NavWord""> ")
				  Response.Write ("<input type=""text"" class=""textbox"" name=""TxtNavi"" onBlur='CheckBadChar(this,""文字导航"");' style=""width:90%;"" value=""" & Navi & """> ")
				  Response.Write ("</div>")
				  Response.Write ("<div align=""left"" id=NavPic style=""display:none""> ")
				  Response.Write ("<input type=""text"" readonly class=""textbox"" style=""width:100"" id=""NaviPic"" name=""NaviPic"">")
				  Response.Write ("<input type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""选择图片..."">")
				  Response.Write ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  Response.Write ("</div>")
				Else
				  Response.Write ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
				  Response.Write ("<input type=""text"" class=""textbox"" name=""TxtNavi"" onBlur='CheckBadChar(this,""文字导航"");' style=""width:90%;""> ")
				  Response.Write ("</div>")
				  Response.Write ("<div align=""left"" id=NavPic> ")
				  Response.Write ("<input type=""text"" readonly class=""textbox"" style=""width:100"" id=""NaviPic"" name=""NaviPic"" value=""" & Navi & """>")
				  Response.Write ("<input type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""选择图片..."">")
				  Response.Write ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  Response.Write ("</div>")
				End If
		Response.Write "        </td>"
		Response.Write "                  </tr>"
		Response.Write "                  <tr id=""MoreLinkArea"">"
		Response.Write "                    <td width=""50%"" style='text-align:left' height=""28"">详情链接"
		 Response.Write "                     <select name=""MoreLinkType"" style=""width:70%;"" class=""textbox"" onchange=""SetMoreLinkStatus()"">"
					If JSID = "" Or CStr(MoreLinkType) = "0" Then
					Response.Write ("<option value=""0"" selected>文字链接</option>")
					Response.Write ("<option value=""1"">图片链接</option>")
				   Else
					Response.Write ("<option value=""0"">文字链接</option>")
					Response.Write ("<option value=""1"" selected>图片链接</option>")
				   End If
		Response.Write "                      </select></td>"
		Response.Write "                    <td width=""50%"" height=""28"">"
				  
				  If JSID = "" Or CStr(MoreLinkType) = "0" Then
					Response.Write ("<div align=""left"" id=""LinkWord""> ")
					Response.Write ("  <input type=""text"" class=""textbox"" onBlur='CheckBadChar(this,""链接字样"");' name=""MoreLinkWord"" style=""width:90%;"" value=""" & MoreLink & """>")
					Response.Write ("</div>")
					Response.Write ("<div align=""left"" id=""LinkPic"" style=""display:none""> ")
					Response.Write ("<input type=""text"" readonly class=""textbox"" style=""width:100"" id=""MoreLinkPic"" name=""MoreLinkPic"">")
					Response.Write ("<input type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.MoreLinkPic);"" name=""Submit3"" value=""选择图片..."">")
					Response.Write ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.MoreLinkPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
					Response.Write ("</div>")
				Else
				   Response.Write ("<div align=""left"" id=""LinkWord"" style=""display:none""> ")
				   Response.Write ("<input type=""text"" class=""textbox"" onBlur='CheckBadChar(this,""链接字样"");' name=""MoreLinkWord"" style=""width:90%;"">")
				   Response.Write ("</div>")
				   Response.Write ("<div align=""left"" id=""LinkPic""> ")
				   Response.Write ("<input type=""text"" readonly class=""textbox"" style=""width:100"" id=""MoreLinkPic"" name=""MoreLinkPic"" value=""" & MoreLink & """>")
				   Response.Write ("<input type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.MoreLinkPic);"" name=""Submit3"" value=""选择图片..."">")
				   Response.Write ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.MoreLinkPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				   Response.Write ("</div>")
				End If
		 Response.Write "       </td>"
		Response.Write "                  </tr>"
		Response.Write "                  <tr>"
		Response.Write "                    <td height=""28"" style='text-align:left' colspan=""2"">分隔图片"
		Response.Write "                      <input name=""SplitPic"" type=""text"" class=""textbox"" id=""SplitPic2"" style=""width:58%;"" value=""" & SplitPic & """ readonly>"
		Response.Write "                      <input name=""SubmitPic"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.SplitPic);"" type=""button"" id=""SubmitPic"" value=""选择图片..."">"
		Response.Write "                      <span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.SplitPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>"
		Response.Write "                     </td>"
		Response.Write "                  </tr>"
		Response.Write "                  <tr>"
		Response.Write "                    <td height=""28"" style='text-align:left' nowrap>日期格式"
		Response.Write ReturnDateFormat(DateRule)
		Response.Write "         </td>"
		Response.Write "                    <td height=""28""> <div align=""left"">日期对齐"
		Response.Write "                        <select name=""DateAlign"" id=""select4"" style=""width:70%;"">"
					
					If JSID = "" Or CStr(DateAlign) = "left" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 Response.Write ("<option value=""left""" & Str & ">左对齐</option>")
					If CStr(DateAlign) = "center" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 Response.Write ("<option value=""center""" & Str & ">居中对齐</option>")
					If CStr(DateAlign) = "right" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 Response.Write ("<option value=""right""" & Str & ">右对齐</option>")
					
		 Response.Write "                       </select>"
		 Response.Write "                     </div></td>"
		 Response.Write "                 </tr>"
		 Response.Write "                 <tr>"
		 Response.Write "                   <td height=""28"" style='text-align:left'>标题样式"
		 Response.Write "                     <input name=""TitleCss"" type=""text"" class=""textbox"" id=""TitleCss"" onBlur=""CheckBadChar(this,'标题样式');"" style=""width:70%;"" value=""" & TitleCss & """></td>"
		 Response.Write "                   <td height=""28"">日期样式<font color=""#FF0000"">"
		 Response.Write "                     <input name=""DateCss"" type=""text"" class=""textbox""  id=""DateCss"" style=""width:70%;"" onBlur=""CheckBadChar(this,'日期样式');"" value=""" & DateCss & """>"
		 Response.Write "                     </font></td>"
		 Response.Write "                 </tr>"
		 Response.Write "                 <tr>"
		 Response.Write "                   <td height=""28"" style='text-align:left'>内容样式"
		 Response.Write "                     <input name=""ContentCss"" type=""text"" class=""textbox"" id=""ContentCss"" style=""width:70%;"" onBlur=""CheckBadChar(this,'内容样式');"" value=""" & ContentCss & """></td>"
		 Response.Write "                   <td height=""28"">背景样式"
		 Response.Write "                      <input name=""BGCss"" type=""text"" class=""textbox"" id=""BGCss"" style=""width:70%;"" onBlur=""CheckBadChar(this,'背景样式');"" value=""" & BGCss & """></td>"
		 Response.Write "                 </tr>"
		 Response.Write "               </table></td>"
		Response.Write "              <td width=""31%"" align=""center""><table width=""90%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "                  <tr>"
		 Response.Write "                   <td height=""25"" align=""center""><strong>样式预览</strong></td>"
		Response.Write "                  </tr>"
		Response.Write "                  <tr>"
		Response.Write "                    <td height=""100%"" id=""ViewStylePicArea"">&nbsp;</td>"
		Response.Write "                  </tr>"
		Response.Write "                </table></td>"
		Response.Write "            </tr>"
		Response.Write "          </table>"
		Response.Write "          </FIELDSET></td>"
		Response.Write "      </tr>"
		Response.Write "    </table>"
		Response.Write "    </form>"
		Response.Write "</div>"
		Response.Write "</body>"
		Response.Write "</html>"
		Response.Write "<script>"
		Response.Write "SelectPicStyle('" & WordCss & "');"
		Response.Write "</script>"
		End Sub
End Class
%> 
