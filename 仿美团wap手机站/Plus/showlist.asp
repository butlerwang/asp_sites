<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file=../"Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%

Dim KS:Set KS=New PublicCLs
Dim KSUser:Set KSUser = New UserCls
Dim LoginTF:LoginTF=KSUser.UserLoginChecked
Dim SQLStr,RS,xml,node,Url,SignUser


'接收参数
Dim Num:Num=KS.ChkClng(Request("num"))                                '列出条数 
Dim TitleLen:TitleLen=KS.ChkClng(Request("titlelen"))                 '标题字数
Dim Tid:Tid=KS.G("Tid")                                               '调用的栏目ID,可留空
Dim ShowClassName:ShowClassName=KS.ChkClng(request("showclassname"))  '显示栏目名称 1显示 0不显示
Dim ShowDate:ShowDate=KS.ChkClng(Request("showdate"))                 '显示时间 1显示 0不显示


If Num=0 Then Num=10
Dim Param:Param=" Where Verific=1"
If Tid<>"" Then
  Param=Param & " and tid='" & tid & "'"
End If
SqlStr= "Select top " &num & " id,tid,title,adddate,fname,issign,signuser From KS_Article " & Param&" order by id desc"
Set RS=Server.CreateObject("adodb.recordset")
RS.Open SQLStr,conn,1,1
If Not RS.Eof Then
  Set xml=KS.RsToXml(rs,"row","")
End If
RS.Close
Set RS=Nothing
If Not IsObject(xml) Then KS.Die ""

For Each Node In Xml.DocumentElement.SelectNodes("row")
  Url=KS.GetItemUrl(1,Node.selectsinglenode("@tid").text,node.selectsinglenode("@id").text,node.selectsinglenode("@fname").text)
  SignUser=Node.selectsinglenode("@signuser").text
  KS.Echo "document.write('<li>"
  If ShowClassName=1 Then    '显示栏目名称
   KS.Echo "<span class=""category"">[" & KS.GetClassNP(Node.SelectSingleNode("@tid").text) &"]</span>"
  End If
  KS.Echo "<a href=""" & url &""" target=""_blank"">" & KS.Gottopic(Node.SelectSingleNode("@title").text,TitleLen) &"</a>"
  If ShowDate=1 Then   '显示日期
    KS.Echo " " & year(node.selectsinglenode("@adddate").text) & "年" &month(node.selectsinglenode("@adddate").text)& "月" &day(node.selectsinglenode("@adddate").text) &"日"
  End If
  
  If node.selectsinglenode("@issign").text="1" and Not KS.IsNul(signuser) then
	  If LoginTF=True Then
	     If KS.FoundInArr(signuser,KSUser.UserName,",")=true Then   '检查当前用户是否在签收用户列表里
		       if conn.execute("select top 1 username from ks_itemsign where username='" & ksuser.username & "' and channelid=1 and infoid=" & node.selectsinglenode("@id").text).eof then
			     KS.Echo " <a href=""" & url & """ target=""_blank""><span class=""qs"">签收</span></a>"
			   end if
		 End If
	  End If
  End If
  KS.Echo "</li>');" &vbcrlf
Next

%>
