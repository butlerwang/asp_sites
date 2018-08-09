<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%

Dim KSCls
Set KSCls = New Frame
KSCls.Kesion()
Set KSCls = Nothing

Class Frame
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<title>" & Request.QueryString("PageTitle") & "</title>"
		%>
		<script>
   window.onunload=SetReturnValue;
	function SetReturnValue()
	{
		if (typeof(window.returnValue)!='string') window.returnValue='';
	}
</script>
		<%
		Dim RequestItem, ParaList, FileName, Url
		ParaList = ""
		For Each RequestItem In Request.QueryString
			If RequestItem <> "FileName" And RequestItem <> "PageTitle" Then
				If ParaList = "" Then
					ParaList = RequestItem & "=" & Server.URLEncode(Request.QueryString(RequestItem))
				Else
					ParaList = ParaList & "&" & RequestItem & "=" & Server.URLEncode(Request.QueryString(RequestItem))
				End If
			End If
		Next
		FileName = Request("FileName")
		If FileName <> "" Then
			Url = FileName & "?" & ParaList
		Else
			Response.Write ("<script language=""JavaScript"">alert('文件不存在');window.close();</script>")
			Response.End
		End If
		
		Response.Write "</head>"
		Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"" scrolling=no>"
		Response.Write "<iframe src=""" & Url & """ width=""100%"" height=""100%"" frameborder=""0"" scrolling=""auto"" align=""center""></iframe>"
		Response.Write "</body>"
		Response.Write "</html>"
		
		End Sub
End Class
%> 
