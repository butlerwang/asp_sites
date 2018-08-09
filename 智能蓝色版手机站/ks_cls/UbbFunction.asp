<%
Private Function UbbCode_S1(re,strText,uCodeC,tCode)
		Dim s
		s=strText
		re.Pattern="\["&uCodeC&"\][\s\n]*\[\/"&uCodeC&"\]"
		s=re.Replace(s,"")
		re.Pattern="\[\/"&uCodeC&"\]"
		s=re.replace(s, Chr(1)&"/"&uCodeC&"]")
		re.Pattern="\["&uCodeC&"\]([^\x01]*)\x01\/"&uCodeC&"\]"
		s=re.Replace(s,tCode)
		re.Pattern="\x01\/"&uCodeC&"\]"
		s=re.replace(s,"[/"&uCodeC&"]")
		UbbCode_S1=s
End Function
'参数 strcontent 内容  n楼层	
Public Function Ubbcode(ByVal strcontent,n)
	
	
    If Instr(StrContent,"[")=0 and Instr(strcontent,"]")=0 Then Ubbcode=strcontent : Exit Function
	'strcontent=server.HTMLEncode(replace(strcontent,chr(10),"[br]"))
	'strcontent=replace(replace(strcontent,"<iframe","&lt;iframe"),"</iframe","&lt;iframe")
	'strcontent=ks.ClearBadChr(KS.CheckScript(strcontent))
	
    If KS.IsNUL(StrContent) Then Ubbcode=" " : Exit Function
	Dim i,re:Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	
	dim maxw
	if Fcls.CallFrom3G="true" then maxw=300 else maxw=600
	
	'图片UBB
		re.pattern="\[img\](.*?)\[\/img\]"
		strcontent=replace(replace(strcontent,"   ","&nbsp; &nbsp;"),"  ","&nbsp;&nbsp;")
		if re.Test(strcontent) then strcontent=re.replace(strcontent,"<a onfocus=""this.blur()"" href=""javascript:;"" onclick=""showbigpic('$1');""><img src=""$1"" border=""0"" alt=""点击查看原图"" style=""max-width:" & maxw &"px;max-height:" & maxw &"px;"" onload=""if(600<this.offsetWidth)this.width='" & maxw &"';if(" & maxw &"<this.offsetHeight)this.height='" & maxw &"0';""></a>")
	
		re.pattern="\[img=*([0-9]*),*([0-9]*)\](.*?)\[\/img\]"
		if re.Test(strcontent) then strcontent=re.replace(strcontent,"<a onfocus=""this.blur()"" href=""javascript:;"" onclick=""showbigpic('$3');""><img src=""$3"" border=""0""  width=""$1"" heigh=""$2"" alt=""点击查看原图"" style=""max-width:" & maxw &"px;max-height:" & maxw &"px;"" onload=""if(" & maxw &"<this.offsetWidth)this.width='" & maxw &"';if(" & maxw &"<this.offsetHeight)this.height='" & maxw &"';""></a>")
	
	re.pattern="\[p=(\d{1,2}|null), (\d{1,2}), (left|center|right)\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<p style=""line-height: $1px; text-indent: $2em; text-align: $3;"">")
	 

	
	'链接UBB
	re.pattern="\[url=" & replace(ks.getdomain,"/","\/") &"(.*?)\](.*?)\[\/url\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<a href=""" & KS.GetDomain &"$1"" target=""_blank"">$2</a>")
	re.pattern="\[url=" & replace(ks.getdomain,"/","\/") &"(.*?)\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<a href=""$2"" target=""_blank"">$2</a>")

	'strcontent=UbbCode_S1(Re,strcontent,"url","$1")
	strcontent=UbbCode_S1(Re,strcontent,"url","<a href=""javascript:;"" onclick=""openlink('$1');"">$1</a>")
	re.pattern="\[url=(.[^\[]*)\]"
	if re.Test(strcontent) then strcontent= re.replace(strcontent,"<a href=""javascript:;"" onclick=""openlink('$1');"">")
	'if re.Test(strcontent) then strcontent= re.replace(strcontent,"")

	'邮箱UBB
	re.pattern="(\[email\])(.*?)(\[\/email\])"
	if re.Test(strcontent) then strcontent= re.replace(strcontent,"<img align=""absmiddle"" ""src=images/common/bb_email.gif""><a href=""mailto:$2"">$2</a>")
	re.pattern="\[email=(.[^\[]*)\]"
	if re.Test(strcontent) then strcontent= re.replace(strcontent,"<img align=""absmiddle"" src=""images/common/bb_email.gif""><a href=""mailto:$1"" target=""new"">")
	'QQ号码UBB
	re.pattern="\[qq]([0-9]*)\[\/qq\]"
	if re.Test(strcontent) then strcontent= re.replace(strcontent,"<a target=""new"" href=""tencent://message/?v=3&uin=$1&Site=" & KS.Setting(0) &"&Menu=yes""><img border=""0"" src=""http://wpa.qq.com/pa?p=2:$1:47"" alt=""点击这里给我发消息""></a>")
    '水平线
	re.pattern="\[hr\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<hr/>")
	re.pattern="\[hr(.[^\[]*)\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<hr$1/>")
    '颜色UBB
	re.pattern="\[color=(.[^\[]*)\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<font color=""$1"">")
    '背景色UBB
	re.pattern="\[backcolor=(.[^\[]*)\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<font style=""background-color:$1"">")
	'文字字体UBB
	re.pattern="\[font=(.[^\[]*)\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<font face=""$1"">")
	'文字大小UBB
	re.pattern="\[size=(\d+?)\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<font size=""$1"">")
	re.pattern="\[size=(\d+(\.\d+)?(px|pt)+?)\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<font style=""font-size:$1"">")
	'文字对齐方式UBB
	re.pattern="\[align=(center|left|right)\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<p align=""$1"">")

	'表格UBB
	re.pattern="\[table\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<table border=""1"" style=""border-collapse:collapse;"">")
	re.pattern="\[table=(.[^\[]*),(.*?)\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<table width=""$1"" border=""1"" style=""border-collapse:collapse;background:$2"">")


	re.pattern="\[table=(.[^\[]*)\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<table width=""$1"" border=""1"" style=""border-collapse:collapse;"">")
    '表格UBB2
	re.pattern="\[td=([0-9]*),([0-9]*),(.*?)\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<td colspan=""$1"" rowspan=""$2"" width=""$3"">")
    re.pattern="\[td=([0-9]*),([0-9]*)\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<td colspan=""$1"" rowspan=""$2"">")

	'字体倾斜
	re.Pattern="\[i\]((.|\n)*?)\[\/i\]"
	if re.Test(strcontent) then strContent=re.Replace(strContent,"<i>$1</i>")
	'浮动代码
	re.pattern="\[float=(left|right)\]"
	if re.Test(strcontent) then strcontent=re.replace(strcontent,"<div style=""float:$1"" class=""floatcode"">")

    'media
	re.pattern="\[media=(flv),*([0-9]*),*([0-9]*),*([0-1]*)\]([^\[]*)\[\/media\]"
	if re.Test(strcontent) then strcontent= re.replace(strcontent,"<embed allowfullscreen=""true"" allowscriptaccess=""always""  bgcolor=""#ffffff"" flashvars=""file=$5&amp;autostart=$4"" height=""$3"" src=""" & KS.GetDomain & "editor/plugins/flvPlayer/jwplayer.swf"" width=""$2""></embed>")
	
	
	re.pattern="\[media=(rm|rmvb),*([0-9]*),*([0-9]*),*([0-1]*)\]([^\[]*)\[\/media\]"
	if re.Test(strcontent) then strcontent= re.replace(strcontent,"<div><object classid=""clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA"" width=""$2"" height=""$3""><param name=""autostart"" value=""$4""/><param name=""src"" value=""$5""/><param name=""controls"" value=""imagewindow""/><param name=""console"" value=""mediaid""/><embed src=""$5"" type=""audio/x-pn-realaudio-plugin"" controls=""IMAGEWINDOW"" console=""mediaid"" width=""$2"" height=""$3""></embed></object><br/><object classid=""clsid:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA"" width=""$2"" height=""32""><param name=""src"" value=""$5"" /><param name=""controls"" value=""controlpanel"" /><param name=""console"" value=""mediaid"" /><embed src=""$5"" type=""audio/x-pn-realaudio-plugin"" controls=""ControlPanel""  console=""mediaid"" width=""$2"" height=""32""></embed></object></div>")

	re.pattern="\[media=(wma),*([0-9]*),*([0-9]*),*([0-1]*)\]([^\[]*)\[\/media\]"
	if re.Test(strcontent) then strcontent= re.replace(strcontent, "<object classid=""clsid:6BF52A52-394A-11d3-B153-00C04F79FAA6"" width=""$2"" height=""64""><param name=""autostart"" value=""$4"" /><param name=""url"" value=""$5"" /><embed src=""$5"" autostart=""$4"" type=""audio/x-ms-wma"" width=""$2"" height=""64""></embed></object>")

	re.pattern="\[media=(mp3),*([0-9]*),*([0-9]*),*([0-1]*)\]([^\[]*)\[\/media\]"
	if re.Test(strcontent) then strcontent= re.replace(strcontent,"<object classid=""clsid:6BF52A52-394A-11d3-B153-00C04F79FAA6"" width=""$2"" height=""64""><param name=""autostart"" value=""$4"" /><param name=""url"" value=""$5"" /><embed src=""$5"" autostart=""$4"" type=""application/x-mplayer2"" width=""$2"" height=""64""></embed></object>")

	re.pattern="\[media=(wmv),*([0-9]*),*([0-9]*),*([0-1]*)\]([^\[]*)\[\/media\]"
	if re.Test(strcontent) then strcontent= re.replace(strcontent,"<object classid=""clsid:6BF52A52-394A-11d3-B153-00C04F79FAA6"" width=""$2"" height=""$3""><param name=""autostart"" value=""$4"" /><param name=""url"" value=""$5"" /><embed src=""$5"" autostart=""$4"" type=""video/x-ms-wmv"" width=""$2"" height=""$3""></embed></object>")

	're.pattern="\[media=(swf),*([0-9]*),*([0-9]*),*([0-1]*)\](http://.[^\[]*)\[\/media\]"
	re.pattern="\[media=(swf),*([0-9]*),*([0-9]*),*([0-1]*)\]([^\[]*)\[\/media\]"
	if re.Test(strcontent) then strcontent= re.replace(strcontent,"<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" width=""$2"" height=""$3""><param name=""autostart"" value=""$4"" /><param name=""src"" value=""$5"" /><embed controller=""true"" width=""$2"" height=""$3"" src=""$5"" autostart=""1""></embed></object>")


    'strcontent=replace(strcontent,vbcrlf,"<BR>")
	re.pattern="\[code\]((.|\n)*?)\[\/code\]"
	dim tempcodes,searcharray,replacearray,tempcode
	Set tempcodes=re.Execute(strcontent)
	For i=0 To tempcodes.count-1
	  tempcode=tempcodes(i)
	  tempcode=replace(tempcode,"[code]"&chr(10)&"[br]","[code]")
	  tempcode=replace(tempcode,"[br][/code]","[/code]")
	  tempcode=Replace("<div><a href=""javascript:;"" onclick=""CopyCode($('#code" &n & i&"')[0])"" style=""color:#999"">复制代码</a></div><div class=""blockcode""><div id=""code" &N& i & """><ol><li>" & tempcode,"<BR>",vbcrlf&"<li>")
	  tempcode=replace(tempcode,"[br]","<li>")
	  tempcode=replace(tempcode,chr(10),"")
	  
	  re.Pattern="(\<li\>){1,}"
	  if re.Test(tempcode) then tempcode=re.Replace(tempcode,"<li>")

	  strcontent=replace(strcontent,tempcodes(i),tempcode)
	next

    searcharray=Array("[br]","[sup]","[/sup]","[sub]","[/sub]","[strike]","[/strike]","[/url]","[/email]","[/backcolor]","[/color]", "[/size]", "[/font]", "[/align]", "[b]", "[/b]","[u]", "[/u]", "[list]", "[list=1]", "[list=a]","[list=A]", "[*]", "[/list]", "[indent]", "[/indent]","[code]","[/code]","[quote]","[/quote]","[free]","[/free]","[hide]","[/hide]","[tr]","[td]","[/td]","[/tr]","[/table]","[/float]","[/p]")
	replacearray=Array("<br/>","<sup>","</sup>","<sub>","</sub>","<strike>","</strike>","</a>","</a>","</font>","</font>", "</font>", "</font>", "</p>", "<b>", "</b>","<u>", "</u>", "<ul>", "<ol type=1>", "<ol type=a>","<ol type=A>", "<li>", "</ul></ol>", "<blockquote>", "</blockquote>","","</ol></div></div>","<div class=""quote"">","</div>","<div class=""quote""><h5>免费内容:</h5><blockquote>","</blockquote></div>","<div class=""quote""><h5>隐藏内容:</h5><blockquote>","</blockquote></div>","<tr>","<td>&nbsp;","</td>","</tr>","</table>","</div>","</p>")
	For i=0 To UBound(searcharray)
		'strcontent=replace(strcontent,searcharray(i),replacearray(i),1,-1,vbTextCompare)
		strcontent=replace(strcontent,searcharray(i),replacearray(i))
	next
	set re=Nothing
    if err then err.clear
	strcontent=replace(strcontent,chr(10),"<br/>")
	Ubbcode=strcontent
	
End Function
%>