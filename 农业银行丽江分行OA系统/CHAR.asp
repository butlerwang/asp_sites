<%
sub error()
%><br>
    <table cellpadding=0 cellspacing=0 border=0 width=95% bgcolor=#777777 align=center>
        <tr>
            <td>
                <table cellpadding=3 cellspacing=1 border=0 width=100%>
    <tr align="center"> 
      <td width="100%" bgcolor=#EEEEEE>错误信息</td>
    </tr>
    <tr> 
      <td width="100%" bgcolor=#FFFFFF><b>产生错误的可能原因：</b><br><br>
<li>
<%=errmsg%>
      </td>
    </tr>
    <tr align="center"> 
      <td width="100%" bgcolor=#EEEEEE>
<a href="javascript:history.go(-1)"> << 返回上一页</a>
      </td>
    </tr>  
    </table>   </td></tr></table>
<%
end sub

function doCode(fString, fOTag, fCTag, fROTag, fRCTag)
	fOTagPos = Instr(1, fString, fOTag, 1)
	fCTagPos = Instr(1, fString, fCTag, 1)
	while (fCTagPos > 0 and fOTagPos > 0)
		fString = replace(fString, fOTag, fROTag, 1, 1, 1)
		fString = replace(fString, fCTag, fRCTag, 1, 1, 1)
		fOTagPos = Instr(1, fString, fOTag, 1)
		fCTagPos = Instr(1, fString, fCTag, 1)
	wend
	doCode = fString
end function

function HTMLEncode(fString)

	fString = replace(fString, ">", "&gt;")
	fString = replace(fString, "<", "&lt;")

	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	HTMLEncode = fString
end function

function HTMLEncode2(fString)
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	HTMLEncode2 = fString
end function

function UBBCode(strContent)
	on error resume next
	strContent = HTMLEncode(strContent)
	dim objRegExp
	Set objRegExp=new RegExp
	objRegExp.IgnoreCase =true
	objRegExp.Global=True

	objRegExp.Pattern="(\[URL\])(.*)(\[\/URL\])"
	strContent= objRegExp.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$2</A>")

	objRegExp.Pattern="(\[URL=(.*)\])(.*)(\[\/URL\])"
	strContent= objRegExp.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$3</A>")

	objRegExp.Pattern="(\[EMAIL\])(.*)(\[\/EMAIL\])"
	strContent= objRegExp.Replace(strContent,"<A HREF=""mailto:$2"">$2</A>")
	objRegExp.Pattern="(\[EMAIL=(.*)\])(.*)(\[\/EMAIL\])"
	strContent= objRegExp.Replace(strContent,"<A HREF=""mailto:$2"" TARGET=_blank>$3</A>")

	objRegExp.Pattern="(\[FLASH\])(.*)(\[\/FLASH\])"
	strContent= objRegExp.Replace(strContent,"<OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$2""><PARAM NAME=quality VALUE=high><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$2</embed></OBJECT>")

	objRegExp.Pattern="(\[IMG\])(.*)(\[\/IMG\])"
	strContent=objRegExp.Replace(strContent,"<IMG SRC=""$2"" border=0>")

        objRegExp.Pattern="(\[HTML\])(.*)(\[\/HTML\])"
	strContent=objRegExp.Replace(strContent,"<SPAN><IMG src=pic/code.gif align=absBottom> HTML 代码片段如下:<BR><TEXTAREA style=""WIDTH: 94%; BACKGROUND-COLOR: #f7f7f7"" name=textfield rows=10>$2</TEXTAREA><BR><INPUT onclick=runEx() type=button value=运行此代码 name=Button> [Ctrl+A 全部选择   提示:你可先修改部分代码，再按运行]</SPAN><BR>")

	objRegExp.Pattern="(\[color=(.*)\])(.*)(\[\/color\])"
	strContent=objRegExp.Replace(strContent,"<font color=$2>$3</font>")
	objRegExp.Pattern="(\[face=(.*)\])(.*)(\[\/face\])"
	strContent=objRegExp.Replace(strContent,"<font face=$2>$3</font>")
	objRegExp.Pattern="(\[align=(.*)\])(.*)(\[\/align\])"
	strContent=objRegExp.Replace(strContent,"<div align=$2>$3</div>")

	objRegExp.Pattern="(\[QUOTE\])(.*)(\[\/QUOTE\])"
	strContent=objRegExp.Replace(strContent,"<BLOCKQUOTE><font size=1 face=""Verdana, Arial"">quote:</font><HR>$2<HR></BLOCKQUOTE>")
	objRegExp.Pattern="(\[fly\])(.*)(\[\/fly\])"
	strContent=objRegExp.Replace(strContent,"<marquee width=90% behavior=alternate scrollamount=3>$2</marquee>")
	objRegExp.Pattern="(\[move\])(.*)(\[\/move\])"
	strContent=objRegExp.Replace(strContent,"<MARQUEE scrollamount=3>$2</marquee>")
	objRegExp.Pattern="(\[glow=(.*),(.*),(.*)\])(.*)(\[\/glow\])"
	strContent=objRegExp.Replace(strContent,"<table width=$2 style=""filter:glow(color=$3, strength=$4)"">$5</table>")
	objRegExp.Pattern="(\[SHADOW=(.*),(.*),(.*)\])(.*)(\[\/SHADOW\])"
	strContent=objRegExp.Replace(strContent,"<table width=$2 style=""filter:shadow(color=$3, direction=$4)"">$5</table>")
    
	objRegExp.Pattern="(\[i\])(.*)(\[\/i\])"
	strContent=objRegExp.Replace(strContent,"<i>$2</i>")
	objRegExp.Pattern="(\[u\])(.*)(\[\/u\])"
	strContent=objRegExp.Replace(strContent,"<u>$2</u>")
	objRegExp.Pattern="(\[b\])(.*)(\[\/b\])"
	strContent=objRegExp.Replace(strContent,"<b>$2</b>")
	objRegExp.Pattern="(\[fly\])(.*)(\[\/fly\])"
	strContent=objRegExp.Replace(strContent,"<marquee>$2</marquee>")

	objRegExp.Pattern="(\[size=1\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=1>$2</font>")
	objRegExp.Pattern="(\[size=2\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=2>$2</font>")
	objRegExp.Pattern="(\[size=3\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=3>$2</font>")
	objRegExp.Pattern="(\[size=4\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=4>$2</font>")

	strContent = doCode(strContent, "[list]", "[/list]", "<ul>", "</ul>")
	strContent = doCode(strContent, "[list=1]", "[/list]", "<ol type=1>", "</ol id=1>")
	strContent = doCode(strContent, "[list=a]", "[/list]", "<ol type=a>", "</ol id=a>")
	strContent = doCode(strContent, "[*]", "[/*]", "<li>", "</li>")
	strContent = doCode(strContent, "[code]", "[/code]", "<pre id=code><font size=1 face=""Verdana, Arial"" id=code>", "</font id=code></pre id=code>")

	set objRegExp=Nothing
	UBBCode=strContent
end function

public function translate(sourceStr,fieldStr)
rem 处理逻辑表达式的转化问题
  dim  sourceList
  dim resultStr
  dim i,j
  if instr(sourceStr," ")>0 then 
     dim isOperator
     isOperator = true
     sourceList=split(sourceStr)
     '--------------------------------------------------------
     rem Response.Write "num:" & cstr(ubound(sourceList)) & "<br>"
     for i = 0 to ubound(sourceList)
        rem Response.Write i 
	Select Case ucase(sourceList(i))
	Case "AND","&","和","与"
		resultStr=resultStr & " and "
		isOperator = true
	Case "OR","|","或"
		resultStr=resultStr & " or "
		isOperator = true
	Case "NOT","!","非","！","！"
		resultStr=resultStr & " not "
		isOperator = true
	Case "(","（","（"
		resultStr=resultStr & " ( "
		isOperator = true
	Case ")","）","）"
		resultStr=resultStr & " ) "
		isOperator = true
	Case Else
		if sourceList(i)<>"" then
			if not isOperator then resultStr=resultStr & " and "
			if inStr(sourceList(i),"%") > 0 then
				resultStr=resultStr&" "&fieldStr& " like '" & replace(sourceList(i),"'","''") & "' "
			else
				resultStr=resultStr&" "&fieldStr& " like '%" & replace(sourceList(i),"'","''") & "%' "
			end if
        		isOperator=false
		End if	
	End Select
        rem Response.write resultStr+"<br>"
     next 
     translate=resultStr
  else '单条件
     if inStr(sourcestr,"%") > 0 then
     	translate=" " & fieldStr & " like '" & replace(sourceStr,"'","''") &"' "
     else
	translate=" " & fieldStr & " like '%" & replace(sourceStr,"'","''") &"%' "
     End if
     rem 前后各加一个空格，免得连sql时忘了加，而出错。
  end if  
end function

function IsValidEmail(email)

dim names, name, i, c

'Check for valid syntax in an email address.

IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
   IsValidEmail = false
   exit function
end if
for each name in names
   if Len(name) <= 0 then
     IsValidEmail = false
     exit function
   end if
   for i = 1 to Len(name)
     c = Lcase(Mid(name, i, 1))
     if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
       IsValidEmail = false
       exit function
     end if
   next
   if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
   end if
next
if InStr(names(1), ".") <= 0 then
   IsValidEmail = false
   exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
   IsValidEmail = false
   exit function
end if
if InStr(email, "..") > 0 then
   IsValidEmail = false
end if

end function
%>
