<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<%

Dim KSCls
Set KSCls = New FriendLinkReg
KSCls.Kesion()
Set KSCls = Nothing
Class FriendLinkReg
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
			Response.Write "<html>" & vbCrLf
			Response.Write "<head>" & vbCrLf
			Response.Write "<title>申请友情链接-KesionCMS</title>" & vbCrLf
			Response.Write "<meta http-equiv=""Content-Language"" content=""zh-cn"">" & vbCrLf
			Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">" & vbCrLf
			Response.Write "<link href=""../../../images/style.css"" rel=""stylesheet"" type=""text/css"">" & vbCrLf
			Response.Write "<script src=""../../../ks_inc/jquery.js""></script>" & vbCrLf
			Response.Write "<script src=""../../../ks_inc/common.js""></script>" & vbCrLf
			Response.Write "</head>" & vbCrLf
			Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">" & vbCrLf
			'Response.Write "<br>" & vbCrLf
			Response.Write "  <table bgcolor=""#ffffff""  width=""960"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
			Response.Write "    <tr>" & vbCrLf
			Response.Write "      <td align=""center""><br>" & vbCrLf
			Response.Write "         <table border=""0"" cellpadding=""2"" cellspacing=""1"" width=""960""  class=""table_border"">" & vbCrLf
			Response.Write "         <tr>" & vbCrLf
			Response.Write "           <td  class=""link_title"" colspan=2>本站链接信息</td>" & vbCrLf
			Response.Write "         </tr>" & vbCrLf
			Response.Write "         <tr>" & vbCrLf
			Response.Write "           <td colspan=""2"" align=""left"" class=""linktext"">申请链接交换，请在您的网站上先做好本站的链接，再提交友情链接申请。</td>"
			Response.Write "         </tr>" & vbCrLf
			Response.Write "         <tr>" & vbCrLf
			Response.Write "           <td width=""176"" height=""30""  class=""link_title"">※本站文字链接代码:</td>"
			Response.Write "           <td width=""313"" height=""30"" class=""link_title""  >演示：<a href=""" & KS.Setting(2) & """ title=""" & KS.Setting(1) & """ target=""_blank"">" & KS.Setting(0) & "</a></td>"
			Response.Write "         </tr>" & vbCrLf
			Response.Write "         <tr align=""center"">" & vbCrLf
			Response.Write "           <td height=""60"" colspan=""2"" style=""padding:8px 0px;""> <textarea name=""textlink"" rows=""4"" onMouseOver=""javascript:this.select();"" style=""width:100%;border:#ccc 1px solid; padding:8px;line-height:22px;color:#999;""><a href=""" & KS.Setting(2) & """ title=""" & KS.Setting(1) & """ target=""_blank"">" & KS.Setting(0) & "</a></textarea></td>"
			Response.Write "         </tr>" & vbCrLf
			Response.Write "         <tr>" & vbCrLf
			Response.Write "           <td width=""176"" height=""30""  class=""link_title"">※本站LOGO链接代码:</td>"
			Response.Write "           <td height=""30""  class=""link_title"" >演示:<a href=""" & KS.Setting(2) & """ title=""" & KS.Setting(1) & """ target=""_blank""><img src=""" & Replace(KS.Setting(4),"{$GetInstallDir}",KS.Setting(3)) & """ width=""88"" height=""31"" border=""0"" align=""absmiddle""></a></td>"
			Response.Write "         </tr>" & vbCrLf
			Response.Write "         <tr align=""center"">" & vbCrLf
			Response.Write "           <td height=""60"" colspan=""2"" style=""padding:8px 0px;""> <textarea name=""logolink"" rows=""5"" onMouseOver=""javascript:this.select();"" style=""width:100%;border:#ccc 1px solid; padding:8px;line-height:22px;color:#999;""><a href=""" & KS.Setting(2) & """ title=""" & KS.Setting(1) & """ target=""_blank""><img src=""" & KS.Setting(4) & """ width=""88"" height=""31"" border=""0""></a></textarea></td>"
			Response.Write "         </tr>" & vbCrLf
			Response.Write "       </table>" & vbCrLf
			Response.Write "         <br>" & vbCrLf
			Response.Write "         <table width=""960"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""table_border"">" & vbCrLf
			Response.Write "           <tr>" & vbCrLf
			Response.Write "             <td  class=""link_title"">申请友情链接</td>" & vbCrLf
			Response.Write "           </tr>" & vbCrLf
			Response.Write "           <tr><td>"
			Response.Write "  <form action=""regsave.asp"" name=""LinkForm"" method=""post"">" & vbCrLf
			Response.Write "   <input name=""Action"" type=""hidden"" id=""Action"" value=""AddLink"">" & vbCrLf
			Response.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
			Response.Write "    <tr>" & vbCrLf
			Response.Write "      <td>" & vbCrLf
			Response.Write "        <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td width=""20%"" height=""30"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"" class=""linktext"">网站名称</div></td>" & vbCrLf
			Response.Write "            <td width=""542"" height=""30"" class=""linktext"">" & vbCrLf
			Response.Write ("<input name=""SiteName"" class=""textbox"" type=""text"" id=""SiteName"" size=""38"" >")
			Response.Write "              <font color=""red"">(*必须填写)</font></td>" & vbCrLf
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""30"" align=""center"" class=""linktext"">所属类别</td>" & vbCrLf
			Response.Write "            <td height=""30""  class=""linktext"">" & vbCrLf
			Response.Write "              <select Name=""FolderID"" >" & vbCrLf
						   
						Dim GRS
						Set GRS = Conn.Execute("Select FolderID,FolderName From KS_LinkFolder Order BY AddDate Desc")
						 Do While Not GRS.EOF
							Response.Write ("<Option value=" & GRS(0) & ">" & GRS(1) & "</OPTION>")
						   GRS.MoveNext
						 Loop
						 GRS.Close
						 Set GRS = Nothing
					   
			 Response.Write "             </Select> </td>" & vbCrLf
			 Response.Write "         </tr>"
			 Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""30"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"" class=""linktext"">网站站长</div></td>" & vbCrLf
			Response.Write "            <td height=""30""  class=""linktext"">" & vbCrLf
			Response.Write "              <input name=""WebMaster"" class=""textbox"" type=""text"" size=""38""> <font color=""red"">(*必须填写)</font></td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""30"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"" class=""linktext"">站长信箱</div></td>" & vbCrLf
			Response.Write "            <td height=""30""  class=""linktext"">" & vbCrLf
			Response.Write "              <input name=""Email"" type=""text"" class=""textbox"" size=""38"" value=""kesioncms@hotmail.com"" ></td>" & vbCrLf
			Response.Write "          </tr>" & vbCrLf
					  
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""30"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"" class=""linktext"">网站密码</div></td>" & vbCrLf
			Response.Write "            <td height=""30"" class=""linktext"">" & vbCrLf
			Response.Write "              <input name=""PassWord"" type=""password"" class=""textbox"" size=""42"" > <font color=""red"">(*必须填写,且不少于6位)</font></td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""30"" align=""center"" class=""linktext"">确定密码</td>" & vbCrLf
			Response.Write "            <td height=""30"" class=""linktext"">" & vbCrLf
			Response.Write "            <input name=""ConPassWord""  class=""textbox"" type=""password"" size=""42"" > <font color=""red"">(*必须填写,同上)</font></td>" & vbCrLf
			Response.Write "          </tr>" & vbCrLf
					 
			
			Response.Write "          <tr>"
			Response.Write "            <td height=""30"" align=""center"" class=""linktext"">网站地址</td>" & vbCrLf
			Response.Write "            <td height=""30"" class=""linktext""><input name=""Url"" class=""textbox"" type=""text""  value=""http://"" id=""Url"" size=""38""> <font color=""red"">(*必须填写,且必须输入http://)</font></td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>"
			Response.Write "            <td height=""30"" align=""center"" class=""linktext"">链接类型</td>" & vbCrLf
			Response.Write "            <td height=""30"" class=""linktext"">"
			Response.Write ("<input type=""radio"" name=""LinkType"" onclick=""SetLogoArea('none')"" value=""0"" checked> 文字链接 ")
			Response.Write ("<input type=""radio"" name=""LinkType"" onclick=""SetLogoArea('')"" value=""1"">  LOGO链接 ")
					   
			Response.Write "             </td>" & vbCrLf
			Response.Write "          </tr>"
			Response.Write "         <tr Style=""display:none"" ID=""LinkArea"">" & vbCrLf
			Response.Write "            <td height=""30"" align=""center"" class=""linktext"">Logo 地址</td>" & vbCrLf
			Response.Write "            <td height=""30""  class=""linktext""><input name=""Logo"" class=""textbox"" type=""text""  value=""http://"" id=""Logo"" size=""38""> <font color=""red"">(*必须填写)</font></td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""30"" align=""center"">"
			Response.Write "              <div align=""center"" class=""linktext"">网站简介</div></td>" & vbCrLf
			Response.Write "            <td height=""30"" class=""linktext"">" & vbCrLf
			Response.Write "              <textarea name=""Description"" rows=""6"" id=""Description"" style=""width:80%;border:#ccc 1px solid;color:#999;line-height:22px;padding:8px;""></textarea></td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""30"" align=""center"" class=""linktext"">认 证 码</td>" & vbCrLf
			Response.Write "            <td height=""30"" class=""linktext"">" & vbCrLf
			Response.Write "            <script>writeVerifyCode('"&KS.Setting(3)&"',1)</script></td>" & vbCrLf
			Response.Write "          </tr>" & vbCrLf
			Response.Write "        </table>" & vbCrLf
			Response.Write "       </td>"
			Response.Write "    </tr>" & vbCrLf
			Response.Write "    </table>" & vbCrLf
			Response.Write "  <table width=""100%"" height=""38"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
			Response.Write "    <tr>" & vbCrLf
			Response.Write "      <td height=""40"" align=""center"">" & vbCrLf
			Response.Write "        <input type=""button"" class=""inputbutton"" name=""Submit"" Onclick=""CheckForm()"" value="" 确 定 "">" & vbCrLf
			Response.Write "        <input type=""reset"" class=""inputbutton"" name=""Submit2""  value="" 重 填 "">" & vbCrLf
			Response.Write "      </td>" & vbCrLf
			Response.Write "    </tr>" & vbCrLf
			Response.Write "  </table>" & vbCrLf
			Response.Write "  </form>" & vbCrLf
			Response.Write "<Script Language=""javascript"">" & vbCrLf
			Response.Write "<!--" & vbCrLf
			Response.Write "function is_email(str)" & vbCrLf
			Response.Write "{ if((str.indexOf('@')==-1)||(str.indexOf('.')==-1)){" & vbCrLf
			Response.Write "    return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "    return true;" & vbCrLf
			Response.Write "}" & vbCrLf
			Response.Write "function SetLogoArea(Value)" & vbCrLf
			Response.Write "{"
			Response.Write "   document.all.LinkArea.style.display=Value;"
			Response.Write "}" & vbCrLf
			Response.Write "function CheckForm()" & vbCrLf
			Response.Write "{ var form=document.LinkForm;" & vbCrLf
			Response.Write "   if (form.SiteName.value=='')" & vbCrLf
			Response.Write "    {"
			Response.Write "     alert(""请输入网站名称!"");" & vbCrLf
			Response.Write "     form.SiteName.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "   if (form.WebMaster.value=='')" & vbCrLf
			Response.Write "    {"
			Response.Write "     alert(""请输入网站站长!"");" & vbCrLf
			Response.Write "     form.WebMaster.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "    if ((form.Email.value!='@')&&(is_email(form.Email.value)==false))" & vbCrLf
			Response.Write "    {"
			Response.Write "    alert('非法电子邮箱!');" & vbCrLf
			Response.Write "     form.Email.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			
			Response.Write "    if (form.PassWord.value=='')" & vbCrLf
			Response.Write "    {"
			Response.Write "     alert(""请输入网站密码!"");" & vbCrLf
			Response.Write "     form.PassWord.focus();" & vbCrLf
			Response.Write "     return false;"
			Response.Write "    }" & vbCrLf
			Response.Write "   else if (form.PassWord.value.length<6)" & vbCrLf
			Response.Write "    {"
			Response.Write "      alert(""网站密码不能少于6位!"");" & vbCrLf
			Response.Write "     form.PassWord.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "   if (form.ConPassWord.value=='')" & vbCrLf
			Response.Write "    {" & vbCrLf
			Response.Write "     alert(""请输入确定密码!"");" & vbCrLf
			Response.Write "     form.ConPassWord.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }"
			Response.Write "   else if(form.ConPassWord.value.length<6)" & vbCrLf
			Response.Write "    {"
			Response.Write "     alert(""确定密码不能少于6位!"");" & vbCrLf
			Response.Write "     form.ConPassWord.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "   if (form.PassWord.value!=form.ConPassWord.value)" & vbCrLf
			Response.Write "    {" & vbCrLf
			Response.Write "     alert(""两次输入的密码不一致!"");" & vbCrLf
			Response.Write "     form.PassWord.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "   if (form.Url.value=='' || form.Url.value=='http://')" & vbCrLf
			Response.Write "    {" & vbCrLf
			Response.Write "     alert(""请输入网站地址"");" & vbCrLf
			Response.Write "     form.Url.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "   if (form.Verifycode.value=='')" & vbCrLf
			Response.Write "    {" & vbCrLf
			Response.Write "     alert(""请输入认证码!"");" & vbCrLf
			Response.Write "     form.Verifycode.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "    form.submit();" & vbCrLf
			Response.Write "    return true;" & vbCrLf
			Response.Write "}" & vbCrLf
			Response.Write "//-->" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			Response.Write " </td></tr></table>" & vbCrLf
			Response.Write "         <br>" & vbCrLf
			Response.Write "       </td>" & vbCrLf
			Response.Write "     </tr>" & vbCrLf
			Response.Write "   </table>" & vbCrLf
			Response.Write " </form>" & vbCrLf
			Response.Write " </body>" & vbCrLf
			Response.Write " </html>" & vbCrLf
			End Sub
End Class
%> 
