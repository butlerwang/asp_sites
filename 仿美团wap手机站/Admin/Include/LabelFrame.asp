<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%

response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Dim KSCls
Set KSCls = New LabelFrame
KSCls.Kesion()
Set KSCls = Nothing

Class LabelFrame
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
		Dim FolderID, Url, FileName, PageTitle, ChannelID, LabelID, Action, LabelType, JSID,TemplateType,sChannelID,JSAction
		Dim QueryParam
		Url = Request.QueryString("Url")
		Action = Request.QueryString("Action")
		JSAction= Request.QueryString("JSAction")
		FolderID = Request.QueryString("FolderID")
		LabelID = Request.QueryString("LabelID")
		LabelType = Request.QueryString("LabelType")
		ChannelID = Request.QueryString("ChannelID")
		JSID = Request.QueryString("JSID")
		PageTitle = Request.QueryString("PageTitle")
		sChannelID=Request.QueryString("sChannelID")
		TemplateType=Request.QueryString("TemplateType")
		QueryParam = "?FolderID=" & FolderID
		If Action <> "" Then QueryParam = QueryParam & "&Action=" & Action
		If JSAction<>"" Then QueryParam = QueryParam & "&JSAction=" & JSAction
		If LabelID <> "" Then QueryParam = QueryParam & "&LabelID=" & LabelID
		If LabelType <> "" Then QueryParam = QueryParam & "&LabelType=" & LabelType
		If ChannelID <> "" Then QueryParam = QueryParam & "&ChannelID=" & ChannelID
		If sChannelID<>"" Then QueryParam = QueryParam & "&sChannelID=" & sChannelID
		If TemplateType<>"" Then QueryParam = QueryParam & "&TemplateType=" & TemplateType
		If JSID <> "" Then QueryParam = QueryParam & "&JSID=" & JSID
		 
		FileName = Url & QueryParam
		
		Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<META HTTP-EQUIV=""pragma"" CONTENT=""no-cache"">" 
        Response.Write "<META HTTP-EQUIV=""Cache-Control"" CONTENT=""no-cache, must-revalidate"">"
        Response.Write "<META HTTP-EQUIV=""expires"" CONTENT=""Wed, 26 Feb 1997 08:21:57 GMT"">"
		Response.Write "<title>" & PageTitle & "</title>"
		Response.Write "</head>"
		Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"" scroll=no>"
		Response.Write "<Iframe src=" & FileName & " style=""width:100%;height:100%;"" frameborder=0 scrolling=""yes""></Iframe>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
End Class
%>
 
