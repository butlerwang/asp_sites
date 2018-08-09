<%


Const ChannelNotOnStr="4,5,6,7,8,9,10"   '定义关闭的模块,请不要随便更改

'获得当前版本号
Function GetVersion()
	Dim Doc:set Doc = CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	Doc.async = false
	Doc.setProperty "ServerHTTPRequest", true 
	Doc.load(Server.MapPath("include/version.xml"))
	if Doc.readystate=4 and Doc.parseError.errorCode=0 Then 
	Dim Node:Set Node= Doc.documentElement.selectSingleNode("//kesioncms/version")
	If Not Node Is Nothing Then GetVersion=Node.text Else GetVersion="8.0"
	end if
End Function

Class KesionCls
	  Private Sub Class_Initialize()
      End Sub
	  Private Sub Class_Terminate()
	  End Sub
	 
	  '系统版本号
	  Public Property Get KSVer
		KSVer="KesionCMS V" & GetVersion &" Free(utf-8)"
	  End Property 
	  
	  '系统缓存名称,如果你的一个站点下安装多套科汛系统，请分别将各个目录下的系统的缓存名称设置成不同
	  Public Property Get SiteSN
	    If EnabledSubDomain Then '如果启用二级域名，则SiteSN必须用固定值
		  SiteSN="KS9"
		Else
		  SiteSN="KS9" & Replace(Replace(LCase(Request.ServerVariables("SERVER_NAME")), "/", ""), ".", "")  
	    End If
	  End Property
	   
End Class
%>