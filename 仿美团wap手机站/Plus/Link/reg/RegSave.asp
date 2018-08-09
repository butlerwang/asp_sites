<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../Plus/md5.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<%

Dim KSCls
Set KSCls = New FriendLinkRegSave
KSCls.Kesion()
Set KSCls = Nothing

Class FriendLinkRegSave
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<title>保存申请友情链接</title>"
		Response.Write "</head>"
		
		Dim LinkID, FolderID, SiteName, WebMaster, Email, PassWord, ConPassWord, Locked, Url, LinkType, Logo, Hits, Recommend, Descript, TrueIP,Verifycode
		Dim TempObj, LinkRS, LinkSql
		
		LinkID = KS.ChkClng(KS.S("LinkID"))
		
		SiteName = KS.S("SiteName")
		WebMaster = KS.S("Webmaster")
		Email = KS.S("Email")
		FolderID = KS.S("FolderID")
		PassWord = Request.Form("PassWord")
		ConPassWord = Request.Form("ConPassWord")
		Verifycode=KS.S("Verifycode")
		If Trim(PassWord) <> Trim(ConPassWord) Then
					Call KS.AlertHistory("网站密码不一致!!!", -1)
					Set KS = Nothing
					Response.End
		End If
		PassWord = MD5(KS.R(PassWord),16)
		Locked = 0
		Url = Replace(Replace(Request.Form("Url"), """", ""), "'", "")
		LinkType = KS.S("LinkType")
		Logo = Replace(Replace(Request.Form("Logo"), """", ""), "'", "")
		Hits = 0
		Recommend = 0
		Descript = KS.R(KS.S("Description"))
		IF lcase(Trim(Request.Form("Verifycode")))<>lcase(Trim(Session("Verifycode"))) then 
			Call KS.AlertHistory("验证码有误，请重新输入!", -1)
			Set KS = Nothing
			Response.End		
		end if
		If SiteName <> "" Then
				If Len(SiteName) >= 200 Then
					Call KS.AlertHistory("网站名称不能超过100个字符!", -1)
					Set KS = Nothing
					 Response.End
				End If
		 Else
				Call KS.AlertHistory("请输入网站名称!", -1)
				Set KS = Nothing
				 Response.End
		 End If
			  Set LinkRS = Server.CreateObject("adodb.recordset")
			  LinkSql = "select * from [KS_Link] Where 1=0"
			  LinkRS.Open LinkSql, Conn, 1, 3
			  LinkRS.AddNew
			  LinkRS("SiteName") = SiteName
			  LinkRS("WebMaster") = WebMaster
			  LinkRS("Email") = Email
			  LinkRS("FolderID") = FolderID
			  LinkRS("PassWord") = PassWord
			  LinkRS("Locked") = Locked
			  LinkRS("Url") = Url
			  LinkRS("LinkType") = LinkType
			  LinkRS("Logo") = Logo
			  LinkRS("Hits") = Hits
			  LinkRS("Recommend") = Recommend
			  LinkRS("Description") = KS.HtmlEnCode(Descript)
			  LinkRS("AddDate") = Now
			  LinkRS("Verific") = 0
			  LinkRS.Update
			  LinkRS.Close
			  Set LinkRS = Nothing
			  Response.Write ("<script> alert('您的申请已成功提交,请等待网站管理员的审核!'); location.href='../';</script>")
		End Sub
End Class
%> 
