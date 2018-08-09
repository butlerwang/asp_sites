<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
Dim JSID,JSRS,Str,JSConfig,FileName,Page
Dim KeyWord,SearchType,StartDate,EndDate
  
'收集搜索参数
KeyWord=Request("KeyWord")
SearchType=Request("SearchType")
StartDate = Request("StartDate")
EndDate = Request("EndDate")
'搜索参数集合
Dim SearchParam
SearchParam="KeyWord=" & KeyWord &"&SearchType=" & SearchType & "&StartDate=" & StartDate& "&EndDate=" & EndDate

JSID=Trim(Request.QueryString("JSID"))
Page=Request.QueryString("Page")
Set JSRS=Server.CreateObject("Adodb.Recordset")
 Str="SELECT JSConfig FROM KS_JSFile Where JSID='" & JSID &"'"
 JSRS.Open Str,Conn,1,1
IF JSRS.Eof and JSRS.Bof THEN
 JSRS.Close
 Set JSRS=Nothing
 Response.Write("<Script>alert('参数传递出错!');history.back();</Script>")
 Response.End
End if 
 JSConfig=JSRS(0)
 JSRS.Close : Set JSRS=Nothing

 If InStr(JSConfig,"{Tag:")<>0 Then
     FileName=Replace(Split(JSConfig," ")(0),"{Tag:","") & ".asp?" & SearchParam &"&Page=" & Page & "&JSID=" & JSID
 Else
	 Str=trim(Split(JSConfig,",")(0))
	 Select Case Str
	   CASE "GetArticleList"
		 FileName="AddClassJS.asp?" & SearchParam &"&Page=" & Page & "&JSID=" & JSID
	   CASE "GetMarqueeArticle"
			  FileName="AddMarqueeArticleJS.asp?" & SearchParam &"&Page=" & Page & "&JSID=" & JSID
	   CASE "GetStripArticle"
			  FileName="AddStripArticleJS.asp?" & SearchParam &"&Page=" & Page & "&JSID=" & JSID
	   CASE "GetPicArticleList"
			 FileName="AddPicArticleJS.asp?" & SearchParam &"&Page=" & Page & "&JSID=" & JSID
	   CASE "GetExtJS"    '扩展JS
			  FileName="AddExtJS.asp?" & SearchParam &"&Page=" & Page & "&JSID=" & JSID		 	  
	   CASE "GetWordJS"    '文字JS
			  FileName="AddWordJS.asp?" & SearchParam &"&Page=" & Page & "&JSID=" & JSID		 	  
	End Select 
 End If
Response.Redirect("JS/AddSysJS.asp?Action=Edit&EditUrl=" & server.urlencode(FileName))
%> 
