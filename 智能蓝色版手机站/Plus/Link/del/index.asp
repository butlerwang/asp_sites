<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../Plus/md5.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<%

Dim KSCls
Set KSCls = New FriendLinkDel
KSCls.Kesion()
Set KSCls = Nothing

Class FriendLinkDel
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
			Dim LinkID, RSCheck, SiteName, WebMaster, Email, OriPassWord, Url, LinkType, Logo, Descript, Action, FolderID
			
			Action = Replace(Replace(Request("Action"), """", ""), "'", "")
			LinkID = KS.ChkClng(Request("id"))
			
			If LinkID=0 Then
			 Set KS = Nothing
			 Response.Write ("<script>alert('参数传递出错!');history.back();</script>")
			End If
			If Action = "Del" Then
			 OriPassWord = MD5(KS.R(Request.Form("OriPassWord")),16)
			 If OriPassWord = "" Then
				  Call KS.AlertHistory("修改友情链接信息密码输入原设密码!", -1)
				  Set KS = Nothing
			End If
			Set RSCheck = Server.CreateObject("Adodb.Recordset")
			   RSCheck.Open " Select LinkID From KS_Link Where PassWord='" & OriPassWord & "'", Conn, 1, 1
			   If RSCheck.EOF And RSCheck.BOF Then
				  RSCheck.Close
				  Set RSCheck = Nothing
				  Call KS.AlertHistory("对不起,你输入的原设密码有误!", -1)
				  Set KS = Nothing
				  Response.End
			  End If
			  Conn.Execute ("Delete From KS_Link Where LinkID=" & LinkID)
			  RSCheck.Close
			  Set RSCheck = Nothing
			  Conn.Close
			  Set Conn = Nothing
			  Response.Write ("<script>alert('友情链接删除成功!');location.href='../';</script>")
			End If
			   Dim RSObj:Set RSObj = Conn.Execute("Select * From KS_Link Where LinkID=" & LinkID)
			  If Not RSObj.EOF Then
				 SiteName = Trim(RSObj("SiteName"))
				 WebMaster = Trim(RSObj("WebMaster"))
				 Email = Trim(RSObj("Email"))
				 Url = Trim(RSObj("Url"))
				 Logo = Trim(RSObj("Logo"))
				 LinkType = Trim(RSObj("LinkType"))
				 Descript = Trim(RSObj("Description"))
				 FolderID = RSObj("FolderID")
			  End If
			   RSObj.Close: Set RSObj = Nothing
			Response.Write ("<html>") & vbCrLf
			Response.Write ("<head>") & vbCrLf
			Response.Write ("<title>删除友情链接</title>") & vbCrLf
			Response.Write ("<meta http-equiv=""Content-Language"" content=""zh-cn"">") & vbCrLf
			Response.Write ("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">") & vbCrLf
			Response.Write ("<link href=""../../../images/style.css"" rel=""stylesheet"" type=""text/css"">") & vbCrLf
			Response.Write ("</head>") & vbCrLf
			Response.Write ("<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">") & vbCrLf
			Response.Write ("<br>") & vbCrLf
			Response.Write ("  <table width=""770"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">") & vbCrLf
			Response.Write ("    <tr>") & vbCrLf
				  
			Response.Write ("    <td align=""center""><br>") & vbCrLf
			Response.Write ("        <table width=""500"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""table_border"">")
			Response.Write ("          <tr class=""link_table_title""> ") & vbCrLf
			Response.Write ("          <td>删除友情链接</td>") & vbCrLf
			Response.Write ("          </tr>") & vbCrLf
			Response.Write ("          <tr><td>") & vbCrLf
			
			Response.Write "  <form action=""?"" name=""LinkForm"" method=""post"">" & vbCrLf
			Response.Write "   <input name=""Action"" type=""hidden"" id=""Action"" value=""Del"">" & vbCrLf
			Response.Write "   <input name=""ID"" type=""hidden"" value=""" & LinkID & """>" & vbCrLf
			Response.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
			Response.Write "    <tr>" & vbCrLf
			Response.Write "      <td>" & vbCrLf
			Response.Write "        <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td width=""20%"" height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">网站名称</div></td>" & vbCrLf
			Response.Write "            <td width=""542"" height=""25"">" & vbCrLf
			Response.Write SiteName & "</td>" & vbCrLf
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">所属类别</td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			 on error resume next			   
			 Response.Write(Conn.Execute("Select FolderName From KS_LinkFolder Where FolderID=" & FolderID)(0))
			 
					   
			 Response.Write "         </td>" & vbCrLf
			 Response.Write "         </tr>"
			 Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">网站站长</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write WebMaster & "</td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">站长信箱</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write Email & "</td>" & vbCrLf
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>"
			Response.Write "            <td height=""25"" align=""center"">网站地址</td>" & vbCrLf
			Response.Write "            <td height=""25"">" & Url & "</td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>"
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">"
			Response.Write "              <div align=""center"">网站简介</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write Descript & "</td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">原设密码</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write "              <input name=""OriPassWord"" type=""password"" size=""42"" > <font color=""red"">* 必须输入</font> </td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "        </table>" & vbCrLf
			Response.Write "       </td>"
			Response.Write "    </tr>" & vbCrLf
			Response.Write "    </table>" & vbCrLf
			Response.Write "  <table width=""100%"" height=""38"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
			Response.Write "    <tr>" & vbCrLf
			Response.Write "      <td height=""40"" align=""center"">" & vbCrLf
			Response.Write "        <input type=""button"" name=""Submit"" Onclick=""CheckForm()"" value="" 确 定 "">" & vbCrLf
			Response.Write "        <input type=""reset"" name=""Submit2""  value="" 重 填 "">" & vbCrLf
			Response.Write "      </td>" & vbCrLf
			Response.Write "    </tr>" & vbCrLf
			Response.Write "  </table>" & vbCrLf
			Response.Write "  </form>" & vbCrLf
			Response.Write "<Script Language=""javascript"">" & vbCrLf
			Response.Write "<!--" & vbCrLf
			Response.Write "function CheckForm()" & vbCrLf
			Response.Write "{ var form=document.LinkForm;" & vbCrLf
			Response.Write "if (form.OriPassWord.value=='')"
			Response.Write "    {"
			Response.Write "     alert(""请输入网站的原设密码!"");" & vbCrLf
			Response.Write "     form.OriPassWord.focus();" & vbCrLf
			Response.Write "     return false;"
			Response.Write "    }" & vbCrLf
			Response.Write " if (confirm('确定删除该站点信息吗?'))"
			Response.Write "  {  form.submit();" & vbCrLf
			Response.Write "    return true;}" & vbCrLf
			Response.Write "else" & vbCrLf
			Response.Write " {location.href='Index.asp'}"
			Response.Write "}" & vbCrLf
			Response.Write "//-->" & vbCrLf
			Response.Write "</Script>"
			Response.Write ("</td></tr></table>") & vbCrLf
			Response.Write ("        <br>") & vbCrLf
			Response.Write ("      </td>") & vbCrLf
			Response.Write ("    </tr>") & vbCrLf
			Response.Write ("  </table>") & vbCrLf
			Response.Write ("</form>") & vbCrLf
			Response.Write ("</body>") & vbCrLf
			Response.Write ("</html>") & vbCrLf
			End Sub
End Class
%>

 
