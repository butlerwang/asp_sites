<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceApp.asp"-->
<%


Dim KSCls
Set KSCls = New SpaceIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SpaceIndex
        Private KS, KSR
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Call CloseConn()
		End Sub
		Public Sub Kesion()
		    If KS.SSetting(0)=0 Then KS.Die "<script>alert('对不起，本站点关闭空间站点功能!');window.close();</script>"
		    Dim QueryStrings:QueryStrings=Request.ServerVariables("QUERY_STRING")
			'ks.die QueryStrings
			If QueryStrings<>"" Then 
			  QueryStrings=KS.UrlDecode(QueryStrings)
			  Dim SApp:Set SApp=New SpaceApp
			  SApp.Show(QueryStrings)
			  If SApp.FoundSpace=false Then KS.Die "<script>alert('该用户没有开通空间!');top.location.href='" & KS.GetDomain &"';</script>"
			  Set SApp=Nothing
			Else
				response.Redirect("../user/space.asp")
		   End If 
		End Sub
		
End Class
%>
