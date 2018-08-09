<%

'作用：同步到第三方微博平台等

'初始化appid,调用同步到微博平台时，要先初始化
sub initialOpenId()
   Dim  XslDoc:Set XslDoc = KS.InitialObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
   If XslDoc.Load(request.ServerVariables("APPL_PHYSICAL_PATH") &"api/api.config") Then
     API_QQEnable = Cbool(XslDoc.documentElement.selectSingleNode("rs:data/z:row").getAttribute("api_qqenable"))
     api_qqappid = XslDoc.documentElement.selectSingleNode("rs:data/z:row").getAttribute("api_qqappid")
	 API_SinaEnable=Cbool(XslDoc.documentElement.selectSingleNode("rs:data/z:row").getAttribute("api_sinaenable"))
	 API_SinaId  = XslDoc.documentElement.selectSingleNode("rs:data/z:row").getAttribute("api_sinaid")
	 API_AlipayEnable=Cbool(XslDoc.documentElement.selectSingleNode("rs:data/z:row").getAttribute("api_alipayenable"))
   End If
end sub

'在需要同步的地方，显示同步选项
function ShowSynchronizedOption(ByRef CheckJS)
    Call initialOpenId()
	if API_QQEnable=false and API_SinaEnable=false then exit function
    dim str:str="<script type=""text/javascript"">" & vbcrlf
			str=str &" function setweibo(type,imgname,title){" & vbcrlf
			str=str &"  if (jQuery('#'+type+'img').attr('src').indexOf('no')==-1){" & vbcrlf
			str=str &"	 jQuery('#'+type+'img').attr('src','" & KS.GetDomain & "images/default/'+imgname+'_no.png');" &vbcrlf
			str=str &"	 jQuery('#'+type+'img').attr('title','未设置同步到'+title);" &vbcrlf
			str=str &"	 jQuery('#'+type).val(0);" &vbcrlf
			str=str &"}else{jQuery('#'+type+'img').attr('src','"&KS.GetDomain &"images/default/'+imgname+'.png');" &vbcrlf
			str=str &"	 jQuery('#'+type+'img').attr('title','已设置同步到'+title);" &vbcrlf
			str=str &"	 jQuery('#'+type).val(1);" &vbcrlf
			str=str &" checktoken(type,imgname);} "&vbcrlf
			str=str &"}" &vbcrlf
			str=str & "function checktoken(type,imgname){"&vbcrlf
			str=str &"jQuery.post(""" & KS.GetDomain & "user/UserAjax.asp"",{action:'CheckToken',checktype:type},function(d){" & vbcrlf
			str=str & "   if (d=='nobind'){" & vbcrlf
			str=str & "	  if (type=='qqweibo'){" &vbcrlf
			str=str &"	 jQuery('#'+type+'img').attr('src','" & KS.GetDomain & "images/default/'+imgname+'_no.png');" &vbcrlf
			str=str & "   $.dialog.confirm('对不起，您还没有绑定QQ登录，无法设置同步，是否现在绑定?',function(){" &vbcrlf
			str=str & "   $.dialog({width:'520px',height:'370px',title:'绑定QQ登录',content:'url:" & KS.GetDomain & "api/qq/redirect_to_login.asp'});},function(){});" &vbcrlf
			str=str &"	 }else if(type=='sinaweibo'){" & vbcrlf
			str=str &"	 jQuery('#'+type+'img').attr('src','" & KS.GetDomain & "images/default/'+imgname+'_no.png');" &vbcrlf
			str=str & "   $.dialog.confirm('对不起，您还没有绑定新浪微博，无法设置同步，是否现在绑定?',function(){" &vbcrlf
			'str=str & "   $.dialog({width:'720px',height:'450px',title:'新浪微博同步绑定',content:'url:" & KS.GetDomain & "api/sina/redirect_to_login.asp'});},function(){});" &vbcrlf
			str=str & "  location.href='" & KS.GetDomain & "api/sina/redirect_to_login.asp';},function(){});" &vbcrlf
			str=str &"  }" & vbcrlf
			str=str &"}else if(d=='error'){" &vbcrlf
			str=str &" if (type=='qqweibo'){ "&vbcrlf
			str=str & "   $.dialog.confirm('对不起，QQ登录授权失败，无法设置同步，是否重新授权?',function(){" &vbcrlf
			str=str & "   $.dialog({width:'520px',height:'370px',title:'绑定QQ登录',content:'url:" & KS.GetDomain & "api/qq/redirect_to_login.asp'});},function(){});" &vbcrlf
			str=str &"  }else if(type=='sinaweibo'){" &vbcrlf
			str=str & "   $.dialog.confirm('对不起，新浪微博授权失败，无法设置同步，是否重新授权?',function(){" &vbcrlf
			str=str & "   location.href='" & KS.GetDomain & "api/sina/redirect_to_login.asp';},function(){});" &vbcrlf
			str=str &"  }" &vbcrlf
			str=str &"}});" & vbcrlf
			str=str &" }" & vbcrlf
			str=str &"</script>" & vbcrlf
			 if cbool(API_QQEnable)=true then
				if not ks.isnul(getuserinfo("qqopenid")) then
					 if KS.FoundInArr(GetUserInfo("Synchronization")&",","1",",")=true Then
							CheckJS="checktoken('qqweibo','qq_weibo');"
					        str=str &"<input type='hidden' value='1' name='qqweibo' id='qqweibo'/><label style=""cursor:pointer"" onclick=""setweibo('qqweibo','qq_weibo','腾讯微博');""><img id='qqweiboimg' src='" & KS.GetDomain & "images/default/qq_weibo.png' title='已设置同步到腾讯微博' align='absmiddle'></label>&nbsp;" &vbcrlf
					 else
							str=str &"<input type='hidden' value='0' name='qqweibo' id='qqweibo'/><label style=""cursor:pointer"" onclick=""setweibo('qqweibo','qq_weibo','腾讯微博');""><img id='qqweiboimg' src='" & KS.GetDomain & "images/default/qq_weibo_no.png' title='未设置同步到腾讯微博' align='absmiddle'></label>&nbsp;" &vbcrlf
					 end if
				else
				 str=str &"<label style=""cursor:pointer"" onclick=""setweibo('qqweibo','qq_weibo','腾讯微博');""><img id='qqweiboimg' src='" & KS.GetDomain &"images/default/qq_weibo_no.png' title='未绑定QQ登录,点击绑定' align='absmiddle'></label>&nbsp;" &vbcrlf
				end if
		  end if
		  if cbool(API_SinaEnable)=true then
				if not ks.isnul(getuserinfo("sinaid")) then
					 if KS.FoundInArr(GetUserInfo("Synchronization")&",","2",",")=true Then
						  CheckJS=CheckJS & "checktoken('sinaweibo','icon_sina');"
					      str=str &" <input type='hidden' value='1' name='sinaweibo' id='sinaweibo'/><label style=""cursor:pointer"" onclick=""setweibo('sinaweibo','icon_sina','新浪微博');""><img id='sinaweiboimg' src='" & KS.GetDomain &"images/default/icon_sina.png' title='已设置同步到新浪微博' align='absmiddle'></label>" &vbcrlf
					 else
						  str=str &"<input type='hidden' value='0' name='sinaweibo' id='sinaweibo'/><label style=""cursor:pointer"" onclick=""setweibo('sinaweibo','icon_sina','新浪微博');""><img id='sinaweiboimg' src='" & KS.GetDomain & "images/default/icon_sina_no.png' title='未设置同步到新浪微博' align='absmiddle'></label>" &vbcrlf
					 end if
				else
				   str=str &"<label style=""cursor:pointer"" onclick=""setweibo('sinaweibo','icon_sina','新浪微博');""><img id='sinaweiboimg' src='" & KS.GetDomain & "images/default/icon_sina_no.png' title='未绑定新浪微博,点击绑定' align='absmiddle'></label>&nbsp;" &vbcrlf
				end if
		end if
		ShowSynchronizedOption=str
end function


'同步到腾讯微博
'参数 content 内容  pic 图片地址，可以留空
function add_qq_weibo(content,pic)
    initialOpenId
	if cbool(API_QQEnable)=false Then Exit function
	if GetUserInfo("qqtoken")="" or GetUserInfo("qqopenid")="" or api_qqappid="" then ks.die "<script>alert('qq授权已过期，请重新授权!');top.location.href='" & KS.Setting(3) &"api/qq/redirect_to_login.asp';</script>"
	Dim str,key
	if pic<>"" then
	    if content<>"" then content=replace(content,"=","〓")
		Dim oDic:Set oDic = Server.CreateObject("Scripting.Dictionary")
		oDic.Add "oauth_consumer_key",api_qqappid
		oDic.Add "access_token",getuserinfo("qqtoken")
		oDic.Add "openid",GetUserInfo("qqopenid")
		oDic.Add "format","xml"
		oDic.Add "content",content
		oDic.Add "pic",pic
		For Each key in oDic
		 if str="" then
		   str=key &"=" &oDic(key) 
		 else
		   str=str &"&" &key &"=" & oDic(key)
		 end if
		Next
		add_qq_weibo = do_post("https://graph.qq.com/t/add_pic_t",str,true)  '带图片
	else		 
		str="format=xml&content=" & content &"&access_token="&getuserinfo("qqtoken")&"&oauth_consumer_key="&api_qqappid&"&openid="&GetUserInfo("qqopenid")
		add_qq_weibo = do_post("https://graph.qq.com/t/add_t",str,false)     '纯文本
    end if
end function

'同步到新浪微博
'参数 content 内容  pic 图片地址，可以留空
function add_sina_weibo(content,pic)
    initialOpenId
	if cbool(API_SinaEnable)=false Then Exit function
	if GetUserInfo("sinatoken")="" or GetUserInfo("sinaid")="" or API_SinaId="" then ks.die "<script>alert('新浪微博授权已过期，请重新授权!');top.location.href='" & KS.Setting(3) &"api/sina/redirect_to_login.asp';</script>"
	Dim str,key
	if pic<>"" then
		Dim oDic:Set oDic = Server.CreateObject("Scripting.Dictionary")
		oDic.Add "access_token",getuserinfo("sinatoken")
		oDic.Add "status",content
		oDic.Add "pic",pic
		For Each key in oDic
		 if str="" then
		   str=key &"=" &oDic(key) 
		 else
		   str=str &"&" &key &"=" & oDic(key)
		 end if
		Next
		add_sina_weibo = do_post("https://api.weibo.com/2/statuses/upload.json",str,true)  '带图片
	else		 
		str="status=" & content &"&access_token="&getuserinfo("sinatoken")
		add_sina_weibo = do_post("https://api.weibo.com/2/statuses/update.json",str,false)     '纯文本
    end if
end function

'同步微博评论到新浪
'参数 weiboid 需要评论的微博ID,content 内容
function add_sina_comment(weiboid,content)
    initialOpenId
	if cbool(API_SinaEnable)=false Then Exit function
	Dim str
	str="id=" & weiboid&"&comment=" & content &"&comment_ori=1&access_token="&getuserinfo("sinatoken")
	add_sina_comment = do_post("https://api.weibo.com/2/comments/create.json",str,false)     '纯文本
end function



'图片构造上传核心代码
Function GetPostParam(str,boundary)
	Dim MPboundary,endMPboundary,multipartbody,aItems,i,objFile,arr,pic,content,filename,data
	MPboundary = "--"&boundary
	endMPboundary = MPboundary&"--"
	multipartbody = "" 			
	Set objFile   =   Server.CreateObject( "ADODB.Stream") 
	objFile.Type   =   2  
	objFile.Mode   =   3  
	objFile.Charset  =   "UTF-8" 
	objFile.Open 
	aItems=Split(str,"&")
	For i=0 To Ubound(aItems)
	arr=Split(aItems(i),"=")
	 If lcase(arr(0))="pic" Then 
		pic= arr(1)
		content=GetPicStream(pic)			
		filename=GetPicExt(pic)
		multipartbody = MPboundary&vbCrLf
		multipartbody  =multipartbody&"Content-Disposition: form-data; name="""&arr(0)&"""; filename="""&filename(1)&""""&vbCrLf
		multipartbody  =multipartbody&"Content-Type: "&filename(0)&""&vbCrLf&vbCrLf
		objFile.WriteText multipartbody
		objFile.Position   =   0 
		objFile.Type   =   1  
		objFile.Position   =   objFile.Size 
		objFile.Write   content
		objFile.Position   =   0 
		objFile.Type   =   2  
		objFile.Position   =   objFile.Size 
		objFile.WriteText  vbCrLf	
	 Else 
		multipartbody = MPboundary&vbCrLf
		multipartbody = multipartbody&"Content-Disposition: form-data; name="""&arr(0)&""""&vbCrLf&vbCrLf
		multipartbody = multipartbody&replace(arr(1),"〓","=")&vbCrLf
		objFile.WriteText multipartbody
	 End If 
	Next
		objFile.WriteText  endMPboundary&vbCrLf 
		objFile.Position   =   0
		objFile.Type   =   1 
		data = objFile.Read(-1)	
		objFile.Close 	
		Set objFile=Nothing
		GetPostParam = data	
End Function

'的取图片类型
Function GetPicExt(url)
 Dim arr(1)
 Select Case Right(LCase(url),4)
 Case ".jpg","jpeg" : arr(0)="image/jpeg": arr(1)="tmp.jpg"
  Case ".gif"       : arr(0)="image/gif" : arr(1)="tmp.gif"
  Case ".png"       : arr(0)="image/png" : arr(1)="tmp.png"
  Case ".bmp"       : arr(0)="image/bmp" : arr(1)="tmp.bmp"
  Case Else         : arr(0)="image/jpeg": arr(1)="tmp.jpg"
 End Select 
 GetPicExt=arr
End Function
'获取图片数据流
Function GetPicStream(url)
    Dim objFile,data,xmlhttp
	If  InStr(Lcase(url),"http://")>0 Then
		Set  xmlhttp=Server.CreateObject("MSXML2.ServerXMLHTTP")
		xmlhttp.open "GET",url,false			 
		xmlhttp.send()
		data=xmlhttp.responseBody
		Set xmlhttp=Nothing
   ELSE
    dim objstream
	set objstream=server.createobject("adodb.stream")
	objstream.Type=1'1为2进制,2为文本
	objstream.mode=3
	objstream.open
	objstream.Position=0
	objstream.loadfromfile server.MapPath(url)
	data=objstream.read(objstream.Size)
	objstream.close
	set objstream=nothing
  End If 
	GetPicStream=data
End Function

'上传到指定的URL，并返回服务器应答
'参数据 isPic true有图片 false无图片
Public Function do_post(ByVal strURL,PostData,IsPic)
	    Dim boundary:boundary="------------------"&DateDiff("s","01/01/1970 08:00:00",Now())
        Dim xmlHttp:Set xmlHttp = CreateObject("Msxml2.XMLHTTP")
        xmlHttp.Open "POST", strURL, False
		if IsPic Then
			xmlHttp.setRequestHeader "Content-Type", "multipart/form-data; boundary="&boundary
			xmlHttp.Send GetPostParam(PostData,boundary)	
		Else
		 	xmlHttp.setRequestHeader "Content-Length", Len(PostData)
			xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            xmlHttp.Send PostData
		End If
		If Err.Number <> 0 Then
			  Set xmlHttp = Nothing
			   do_post = "Error"
			   Exit Function
		Else
		       do_post = xmlHttp.responseText
		End If
End Function
			
%>
