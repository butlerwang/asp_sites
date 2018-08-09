<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../Plus/md5.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<%

Dim KSCls
Set KSCls = New FriendLinkModifySave
KSCls.Kesion()
Set KSCls = Nothing

Class FriendLinkModifySave
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

Dim LinkID, FolderID, SiteName, WebMaster, Email, OriPassWord, PassWord, ConPassWord, Locked, Url, LinkType, Logo, Hits, Recommend, Descript, TrueIP
Dim TempObj, LinkRS, LinkSql, RSCheck

LinkID = KS.ChkClng(KS.S("LinkID"))

OriPassWord = MD5(KS.R(Request.Form("OriPassWord")),16)
If OriPassWord = "" Then
      Call KS.AlertHistory("修改友情链接信息密码输入原设密码!", -1)
      Set KS = Nothing
End If
Set RSCheck = Server.CreateObject("Adodb.Recordset")
   RSCheck.Open " Select LinkID From KS_Link Where PassWord='" & OriPassWord & "' and linkid=" & linkid , Conn, 1, 1
   If RSCheck.EOF And RSCheck.BOF Then
      RSCheck.Close:Set RSCheck = Nothing
      Call KS.AlertHistory("对不起,你输入的原设密码有误!", -1)
      Set KS = Nothing
      Response.End
  End If
SiteName = KS.S("SiteName")
WebMaster = KS.S("Webmaster")
Email = KS.R(Request.Form("Email"))
FolderID = KS.S("FolderID")
PassWord = Request.Form("PassWord")
ConPassWord = Request.Form("ConPassWord")

If Trim(PassWord) <> Trim(ConPassWord) Then
            Call KS.AlertHistory("网站密码不一致!!!", -1)
            Set KS = Nothing
            Response.End
End If
PassWord = MD5(KS.R(PassWord),16)

Url = Replace(Replace(Request.Form("Url"), """", ""), "'", "")
LinkType = KS.S("LinkType")
Logo = Replace(Replace(Request.Form("Logo"), """", ""), "'", "")
Descript = KS.R(KS.S("Description"))

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
      LinkSql = "select * from [KS_Link] Where LinkID=" & LinkID
      LinkRS.Open LinkSql, Conn, 1, 3
      LinkRS("SiteName") = SiteName
      LinkRS("WebMaster") = WebMaster
      LinkRS("Email") = Email
      If KS.S("PassWord") <> "" Then
      LinkRS("PassWord") = PassWord
      End If
      LinkRS("FolderID") = FolderID
      LinkRS("Url") = Url
      LinkRS("LinkType") = LinkType
      LinkRS("Logo") = Logo
      LinkRS("Description") = KS.HtmlEnCode(Descript)
      LinkRS.Update
      LinkRS.Close
      Set LinkRS = Nothing
      Response.Write ("<script>alert('修改友情链接成功!');location.href='../';</script>")
End Sub
End Class
%> 
