<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%

'==============================================================
'请根据你的需要自行修改以下代码
'本文件调用方式:<script src='/plus/browsing.asp?channelid={$ChannelID}&id={$InfoID}&Num=10'><//script>
'=================================================================
Dim KS:Set KS=New PublicCls
Dim ID:ID=KS.ChkClng(KS.S("ID"))
Dim ChannelID:ChannelID=KS.ChkCLng(KS.S("Channelid"))
Dim Num:Num=KS.ChkClng(KS.S("Num"))
Dim RS,SQL,K,Str,Url

If Num=0 Then Num=10
If ChannelID=0 Then ChannelID=5
Dim IDList:IDList=KS.C("View" & ChannelID)
If KS.FoundInArr(IDList,ID,",")=False Then
 If IDList="" Then
  IDList=ID
 Else
  IDList=ID&"," & IDList
 End If
 Dim IDArr,T_Str
 IDArr=Split(IDList,",")
 For I=0 To Ubound(IDArr)
   If I<Num Then
     If T_Str="" Then
	  T_Str=IDArr(i)
	 Else
	  T_Str=T_Str & "," & IDArr(i)
	 End If
   End If 
 Next
   If EnabledSubDomain Then
	  Response.Cookies(KS.SiteSn).domain=RootDomain					
	Else
		Response.Cookies(KS.SiteSn).path = "/"
	End If
  Response.Cookies(KS.SiteSN)("View" & ChannelID)=T_Str
End If

Select Case KS.C_S(ChannelID,6)
 Case 1 Call ArticleList()
 Case 5  Call ProductList()
End Select


Sub ArticleList()
    str=""
	Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "Select Top " & Num & " ID,Title,Tid,Fname,Changes,InfoPurview,ReadPoint From KS_Article Where ID In(" & KS.R(KS.C("View" & ChannelID)) & ") order by id desc",conn,1,1
	If Not RS.Eof Then SQL=RS.GetRows(-1):RS.Close:Set RS=Nothing
	If IsArray(SQL) Then  
	  For K=0 To Ubound(SQL,2)
	   Url=KS.GetItemURL(ChannelID,sql(2,k),sql(0,k),sql(3,k))
	   str=str & "<div class=""Browsing""><ul>"
	   str=str & "<li><a href=""" & URL & """ target=""_blank"">" & KS.Gottopic(SQL(1,K),38) & "</a></li>"
	   str=str & "</ul></div>"
	  Next
	  Erase SQL
	End If
End Sub

Sub ProductList()
    str=""
	Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "Select Top " & Num & " ID,Title,Tid,Fname,PhotoUrl,Price_member From KS_Product Where ID In(" & KS.R(KS.C("View" & ChannelID)) & ") order by id desc",conn,1,1
	If Not RS.Eof Then SQL=RS.GetRows(-1):RS.Close:Set RS=Nothing
	If IsArray(SQL) Then  
	  For K=0 To Ubound(SQL,2)
	   Url=KS.GetItemURL(ChannelID,sql(2,k),sql(0,k),sql(3,k))
	   str=str & "<div class=""sidepd"">"
	   str=str & "<a class=""sidepdleft"" href=""" & Url & """ target=""_blank""><img width=""65"" height=""65"" src=""" & SQL(4,K) & """ border=""0""></a>"
	   dim price:price=SQL(5,K)
	   if price<1 then price= "0"&price
	   str=str & "<h2><a href=""" & URL & """ target=""_blank"">" & KS.Gottopic(SQL(1,K),38) & "</a></h2><h3><span>￥</span>" & price & "</h3>"
	   str=str & "</div>"
	  Next
	  Erase SQL
	End If
End Sub

Response.Write("document.writeln('"& str & "');")

CloseConn()
Set KS=Nothing

%>