<!--#include file="../../conn.asp"-->
<!--#include file="../../ks_cls/kesion.membercls.asp"-->
<!--#include file="../cls_api.asp"-->
<%
'请将下面信息更改成自己申请的信息
Dim appid  : appid   = API_QQAppId  'opensns.qq.com 申请到的appid
Dim appkey : appkey  = API_QQAppKey 'opensns.qq.com 申请到的appkey
Dim callback:callback = API_QQCallBack 'QQ登录成功后跳转的地址


'生成时间戳 
Function ToUnixTime(strTime, intTimeZone)
If IsEmpty(strTime) or Not IsDate(strTime) Then strTime = Now
If IsEmpty(intTimeZone) or Not isNumeric(intTimeZone) Then intTimeZone = 0
ToUnixTime = DateAdd("h",-intTimeZone,strTime)
ToUnixTime = DateDiff("s","1970-1-1 0:0:0", ToUnixTime)
End Function

'生成随机数
Public Function MakeRandom(ByVal maxLen)
	  Dim strNewPass,whatsNext, upper, lower, intCounter
	  Randomize
	 For intCounter = 1 To maxLen
	   upper = 57:lower = 48:strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	 Next
	   MakeRandom = strNewPass
End Function





'将url变成集合
function parse_str(str)
dim objData,aryData,i,aryT
set objData=Server.CreateObject("Scripting.Dictionary")
aryData=split(str,"&")
for i=0 to ubound(aryData)
   aryT=split(aryData(i),"=")
   if ubound(aryT)>0 then 
    objData.add aryT(0),aryT(1)
   else
    objData.add aryT(0),""
   end if
next
set parse_str=objData
end function


%>
