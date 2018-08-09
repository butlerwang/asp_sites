<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New CacheMain
KSCls.Kesion()
Set KSCls = Nothing

Class CacheMain
        Private KS,CacheNum
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
       Sub Kesion()
	     If KS.G("Action")="Clean" Then
		   Call CleanCache()
		 Else
		   Call CacheMain()
		 End If
	   End Sub
	   Sub CacheMain
         With Response
			 If Not KS.ReturnPowerResult(0, "KMST20000") Then
			  Call KS.ReturnErr(1, "")
			  response.End()
			End If
			.Write "<html>"
			.Write"<head>"
			.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write"</head>"
			.Write"<body scroll=no leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			.Write"<div class=""topdashed sort"">更新站点缓存</div>"
			.Write "<table width='100%' height='100%'>"
			.Write  "<tr>"
			.Write " <td> <iframe scrolling=""auto"" frameborder=""0"" src=""KS.CleanCache.asp?Action=Clean"" width=""100%"" height=""100%""></iframe>"
			.Write"</td>"
			.Write " </tr>"
			.Write"</TABLE>"
		End With
      End Sub
	  
	  Sub CleanCache()
		With Response
		  .Write "<html>"
		  .Write "<head>"
		  .Write "<title>缓存更新</title>"
		  .Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		  .Write "<link href=""Include/Admin_Style.Css"" rel=""stylesheet"" type=""text/css"">"
		  .Write "</head>"
		  .Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		  .Write "    <table width=""100%""  border=""0"" align=""center""cellspacing=""0"" cellpadding=""0"">"
		  .Write "    <tr class='sort'><td width='40' align=center><b>序号</b></td><td width='550' align='center'><b>更新对象</b></td><td align=""center""><b>状态</b></td></tr>"
		
					delallcache()
		 
		.Write "<script>function back(){history.back();}setTimeout('back()',1000);</script>"
		.Write "</body>"
		.Write "</html>"
    End With
    End Sub
	
	Sub delallcache()
		Dim CacheList,i
		CacheList=split(KS.GetCacheList(KS.SiteSN),",")
		CacheNum=UBound(CacheList)
		With Response
		If CacheNum>1 Then
			For i=0 to CacheNum-1
				KS.DelCahe CacheList(i)
				.Write "<tr height=""22"" class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'""><td class='splittd' align=""center"">"
				.Write i+1 & "</td><td class='splittd'><font color='#FF6600'>"&Replace(CacheList(i),KS.SiteSN & "","")&"</font></td><td align=""center"" class='splittd'>完成</td></tr>"	
			Next
			.Write "<tr height=""22"" class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'""><td colspan='3' class='splittd' align='center'>共更新了&nbsp;&nbsp;<font color='#FF6600'>"
			.Write CacheNum
			.Write "</font>&nbsp;&nbsp;个缓存对象</td></tr>"	
		Else
			.Write "<tr height=""22"" class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'""><td colspan='3' class='splittd' align='center'>所有缓存对象已经更新。</td></tr>"
		End If
	  End With
	End Sub        
End Class
%> 
