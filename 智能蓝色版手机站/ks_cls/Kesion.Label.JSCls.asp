<%

Dim JSCls:Set JSCls=New JSCommonCls
Class JSCommonCls
        Dim KS,Temps,DomainStr
		Private Sub Class_Initialize()
		Set KS=New PublicCls
		DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		 Set JSCls=Nothing
		 Set KS=Nothing
		End Sub
		'====================================================替换通用JS=============================================
		'说明：这里可以自由扩展，把常用的JS代码进行封装，就可以利用标签调用
		'============================================================================================================
		Sub echo(str)
		  temps=temps & str
		End Sub
		Sub echoln(str)
		  temps=temps & str & vbcrlf
		End Sub
		Sub Run(sTemp,ByRef Templates)
		 dim RCls: set RCls=New Refresh
		  temps=Templates
		 select case Lcase(sTemp)
		   case "js_time1" : echo  "<script src=""" & DomainStr & "ks_inc/time/1.js"" type=""text/javascript""></script>"
		   case "js_time2" : echo  "<script src=""" & DomainStr & "ks_inc/time/2.js"" type=""text/javascript""></script>"
		   case "js_time3" : echo  "<script src=""" & DomainStr & "ks_inc/time/3.js"" type=""text/javascript""></script>"
		   case "js_time4" : echo  "<div id=""kstime""></div><script>setInterval(""kstime.innerHTML=new Date().toLocaleString()+' 星期'+'日一二三四五六'.charAt (new Date().getDay());"",1000);</script>"
		   case "js_language" : echo "<script src=""" & DomainStr & "KS_Inc/language.js"" type=""text/javascript""></script>"
		   case "js_collection": echo "<a href=""#"" onclick=""javascript:window.external.addFavorite('http://'+location.hostname+(location.port!=''?':':'')+location.port,'" & KS.Setting(0) &"');"">加入收藏</a>"
		   case "js_homepage" : echo "<a onclick=""this.style.behavior='url(#default#homepage)';this.setHomePage('http://'+location.hostname+(location.port!=''?':':'')+location.port);"" href=""#"">设为首页</a>"
		   case "js_contactwebmaster" : echo "<a href=""mailto:" & KS.Setting(11) & """>联系站长</a>"
		   case "js_nosave" : echo "<NOSCRIPT><IFRAME SRC=*.html></IFRAME></NOSCRIPT>"
		   case "js_goback" : echo "<a href=""javascript:history.back(-1)"">返回上一页</a>"
		   case "js_windowclose" : echo "<a href=""javascript:window.close();"">关闭窗口</a>"
		   case "js_noiframe": echo "<script type=""text/javascript"">if(self!=top){top.location=self.location;}</script>"
		   case "js_nocopy" : echoln "<script type=""text/javascript"">" 
		                      echoln "document.oncontextmenu=new Function(""event.returnValue=false;"");"  
							  echoln "document.onselectstart=new Function(""event.returnValue=false;"");"
							  echoln "</script>"
		   case "js_dcroll" : echoln "<script type=""text/javascript"">"
		                      echoln "var currentpos,timer; " 
							  echoln "function initialize(){ timer=setInterval(""scrollwindow()"",30);} " 
							  echoln "function sc(){clearInterval(timer);}" 
							  echoln "function scrollwindow(){ "
							  echoln "if (document.documentElement && document.documentElement.scrollTop){"
							  echoln " currentpos=document.documentElement.scrollTop;window.scroll(0,++currentpos); "
							  echoln " if (currentpos != document.documentElement.scrollTop) sc();}"
							  echoln "else if (document.body){"
							  echoln "	currentpos=document.body.scrollTop; window.scroll(0,++currentpos);"
							  echoln "if (currentpos != document.body.scrollTop) sc(); }"
							  echoln "} "
							  echoln "document.onmousedown=sc"
							  echoln "document.ondblclick=initialize"
							  echoln "</script>"
		 end select
		  Templates=temps
		End Sub
		
		Sub Equal(stemp,Param,ByRef Templates)
		  dim RCls: set RCls=New Refresh
		  temps=Templates
		  select case Lcase(stemp)
		    case "js_ad" '对联广告
			  echo "<script>var delta=" & Param(3) & ";var closeSrc='" & Param(2) & "';var rightSrc='" & Param(1) & "';var leftSrc='" & Param(0) & "';</script><script src=""" & DomainStr & "ks_inc/ad/1.js"" type=""text/javascript""></script>" 
		    case "js_status1" '状态栏目打字效果
			  echo "<script type=""text/javascript"">var msg = '" & Param(0) & "' ;var interval = " & Param(1) & ";</script><script src=""" & DomainStr & "ks_inc/status/1.js"" type=""text/javascript""></script>"
			case "js_status2" '文字在状态栏上从右往左循环显示
			  echo "<script>var speed = " & Param(1) &";var m1 = '" & Param(0) & "' ;</script><script src=""" & DomainStr & "ks_inc/status/2.js"" type=""text/javascript""></script>"
			case "js_status3" '文字在状态栏上打字之后移动消失
			  echo "<script>var speed = " & Param(1) &";var Message = '" & Param(0) & "' ;</script><script src=""" & DomainStr & "ks_inc/status/3.js"" type=""text/javascript""></script>"
		  end select
		  Templates=temps
		End Sub
		
		
        '====================================================替换通用JS结束=============================================

End Class
%> 
