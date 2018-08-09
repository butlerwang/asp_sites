<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="KS_Cls/Kesion.AppCls.asp"-->
<%

Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing
Const AllowSecondDomain=true       '是否允许开启空间二级域名 true-开启 false-不开启


Class SiteIndex
        Private KS,AppCls
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set AppCls=New KesionAppCls
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set AppCls=Nothing
		 Call CloseConn()
		End Sub
		Public Sub Kesion()
		    If AllowSecondDomain=True And KS.IsNul(Request.QueryString("do")) Then 
			    SecondDomain
			Else
                Call AppCls.HomePage()
			End If
		End Sub
		
		Public Sub SecondDomain()
		dim From,gourl,sdomain,title,username,domain
		From = LCase(Request.ServerVariables("HTTP_HOST"))
		sdomain = LCase(KS.SSetting(15))
		sdomain = Replace(sdomain,"http://","")
		sdomain = Replace(sdomain,"/","")
		
		dim domain1,domain2
		domain = LCase (from)
		domain = Replace (domain,"http://","")
		domain = Replace (domain,"/","")
		If lcase(domain)=lcase(KS.WSetting(1)) or lcase(domain)=lcase(KS.Setting(69)) or lcase(domain)=lcase(KS.JSetting(41)) or (sdomain=domain and sdomain<>"") Then  '论坛
                Call AppCls.Domain(domain)
				Exit Sub
		else
			 domain1= Replace (Left (domain,InStr (domain,".")),".","")
			 if Trim (domain1)="" or (domain1="www" and domain=replace(lcase(KS.Setting(2)),"http://","")) or (Request.ServerVariables("HTTP_HOST")="http://" & KS.Setting(2)) or ("http://" & lcase(Request.ServerVariables("HTTP_HOST"))=lcase(KS.Setting(2))) Then 
			     Call AppCls.HomePage() : Exit Sub
			 Else
 
			 End If
		        Set AppCls=New KesionAppCls
			    if instr(domain,replace(replace(lcase(KS.Setting(2)),"http://",""),"www.",""))=0 and domain1="www" then
                Call AppCls.Domain(domain)
				else
                Call AppCls.Domain(domain1)
				end if
				Exit Sub
			end if
	 End Sub
End Class
%>
