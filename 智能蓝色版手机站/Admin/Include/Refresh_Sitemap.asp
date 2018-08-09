<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="Admin_Style.CSS" rel="stylesheet" type="text/css">
<title>Google Sitemap</title>

<%


dim xmlstr,lastmod
dim sql_KS_Class,SqlStr,rs,rsclass,i,Classpath
dim sitemappath
Dim KS:Set KS=New PublicCls


'=========================主程序==========================
If KS.G("Action")<>"" Then
    Dim changefreq:changefreq=KS.G("changefreq")
	Dim prioritynum:prioritynum=KS.ChkCLng(KS.G("prioritynum"))
	dim tmFile,objFso,smw
	sitemappath=KS.Setting(3)&"sitemap.xml"
	Set objFso = KS.InitialObject(KS.Setting(99))

	if KS.G("Action")="creategoogle" then
		If prioritynum=0 then prioritynum=15
		Dim big:big=KS.G("Big")
		Dim SQL,K
		Set RS=KS.InitialObject("ADODB.RECORDSET")
		RS.Open "Select BasicType,ChannelTable,ChannelID From KS_Channel Where ChannelStatus=1 And ChannelID<>6 And Channelid<>9 And ChannelID<>10 Order By ChannelID",Conn,1,1
		SQL=RS.GetRows(-1)
		RS.Close

		xmlstr="<?xml version=""1.0"" encoding=""UTF-8""?>"&vbcrlf
		xmlstr=xmlstr&"<urlset xmlns=""http://www.google.com/schemas/sitemap/0.84"">"&vbcrlf
	
		For K=0 To Ubound(SQL,2)
		 Select Case  SQL(0,K)
		  Case 1 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate"
		  Case 2 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate"
		  Case 3 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate"
		  Case 4 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate"
		  Case 5 :SqlStr="select top " & prioritynum & " ID,Title,Tid,0,0,Fname,0,AddDate"
		  Case 7 :SqlStr="select top " & prioritynum & " ID,Title,Tid,0,0,Fname,0,AddDate"
		  Case 8 :SqlStr="select top " & prioritynum & " ID,Title,Tid,0,0,Fname,0,AddDate"
		 End Select
		
		SqlStr=SqlStr & " from "& SQL(1,K) & " where verific=1 and deltf=0 order by id desc"
		rs.Open SqlStr,conn,1,1
		for i=1 to rs.RecordCount
			xmlstr=xmlstr&"    <url>"&vbcrlf
			xmlstr=xmlstr&"        <loc><![CDATA["&KS.GetItemUrl(SQL(2,K),RS(2),RS(0),RS(5))&"]]></loc>"&vbcrlf
			xmlstr=xmlstr&"        <lastmod>" & GetDate(rs(7)) & "</lastmod>"&vbcrlf
			xmlstr=xmlstr&"        <changefreq>"&changefreq&"</changefreq>"&vbcrlf
			xmlstr=xmlstr&"        <priority>"&big&"</priority>"&vbcrlf
			xmlstr=xmlstr&"    </url>"&vbcrlf
			rs.movenext 
		next
		rs.close
	  Next
	'=sitemap===============================================================================================================
		xmlstr=xmlstr&"</urlset>"
	
	
		'==============写入sitemap======================
		Call KS.WriteTOFile(sitemappath,xmlstr)
	   '===========sitemap================================
	
	response.write("<script language='JavaScript' type='text/JavaScript'>")
	response.write("function yy() {")
	response.write("overstr.innerHTML='<div align=center>恭喜,sitemap.xml生成完毕！<br><br><a href=" & KS.Setting(3) & "sitemap.xml target=_blank>点击查看生成好的sitemap.xml文件</a></div>'; }")
	response.write("</script>")
	
	elseif  KS.G("Action")="createbaidu" then
	
		xmlstr="<?xml version=""1.0"" encoding=""utf-8""?>"&vbcrlf
	    xmlstr=xmlstr & "<document>"
		xmlstr=xmlstr & "<webSite>" & Replace(KS.Setting(2),"http://","") & "</webSite>"
		xmlstr=xmlstr & "<webMaster>" & KS.Setting(11) &"</webMaster>"
		xmlstr=xmlstr & "<updatePeri>" &changefreq & "</updatePeri>"
		Dim Num:Num=0
		Set RS=KS.InitialObject("ADODB.RECORDSET")
		RS.Open "Select BasicType,ChannelTable,ChannelID From KS_Channel Where ChannelStatus=1 And ChannelID<>6 And Channelid<>9 And ChannelID<>10 Order By ChannelID",Conn,1,1
		SQL=RS.GetRows(-1)
		RS.Close

		
		For K=0 To Ubound(SQL,2)
		 Select Case  SQL(0,K)
		  Case 1 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate,Intro,ArticleContent,PhotoUrl,author,origin"
		  Case 2 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,PictureContent,PictureContent,photourl,author,origin"
		  Case 3 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,downcontent,downcontent,photourl,author,origin"
		  Case 4 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,flashcontent,flashcontent,photourl,author,origin"
		  Case 5 :SqlStr="select top " & prioritynum & " ID,Title,Tid,0,0,Fname,0,AddDate,prointro,prointro,photourl,ProducerName,TrademarkName"
		  Case 7 :SqlStr="select top " & prioritynum & " ID,Title,Tid,0,0,Fname,0,AddDate,moviecontent,moviecontent,photourl,MovieAct,MovieDQ"
		  Case 8 :SqlStr="select top " & prioritynum & " ID,Title,Tid,0,0,Fname,0,AddDate,gqcontent,gqcontent,photourl,inputer,ContactMan"
		 End Select

		
		 SqlStr=SqlStr & " from "& SQL(1,K) & " where verific=1 and deltf=0 order by id desc"
		 
		 
		rs.Open SqlStr,conn,1,1
		for i=1 to rs.RecordCount
			xmlstr=xmlstr&"    <item>"&vbcrlf
			xmlstr=xmlstr&"        <title>" & replace(KS.LoseHtml(rs(1)),"&nbsp;"," ") &"</title>"
			xmlstr=xmlstr&"        <link><![CDATA["&KS.GetItemUrl(SQL(2,K),RS(2),RS(0),RS(5))&"]]></link>"&vbcrlf
			xmlstr=xmlstr&"        <description><![CDATA[" & Replace(KS.LoseHtml(rs(8)),"&nbsp;","") & "]]></description>"&vbcrlf
			xmlstr=xmlstr&"        <text><![CDATA[" &Replace(KS.LoseHtml(rs(9)),"&nbsp;","") & "]]></text>"&vbcrlf
			if Not KS.IsNul(RS(10)) Then
			 Dim PhotoUrl:PhotoUrl=RS(10)
			 If Left(Lcase(PhotoUrl),4)<>"http" Then
			   PhotoUrl=KS.Setting(2) & PhotoUrl
			 End If
			xmlstr=xmlstr&"        <image><![CDATA["&PhotoUrl&"]]></image>"&vbcrlf
			End If
			xmlstr=xmlstr&"        <category><![CDATA["&KS.C_C(RS(2),0)&"]]></category>"&vbcrlf
			xmlstr=xmlstr&"        <author><![CDATA["&rs(11)&"]]></author>"&vbcrlf
			xmlstr=xmlstr&"        <source><![CDATA["&rs(12)&"]]></source>"&vbcrlf
			xmlstr=xmlstr&"        <pubDate>"&GetDate(rs(7))&"</pubDate>"&vbcrlf
			xmlstr=xmlstr&"    </item>"&vbcrlf
			Num=Num+1
			If Num>=100 Then Exit For
			rs.movenext 
		next
		rs.close
		If Num>=100 Then Exit For
	  Next
	'=sitemap===============================================================================================================
		
		
		xmlstr=xmlstr & "</document>"
	
		'==============写入news.xml======================
		Dim NewsPath:NewsPath=KS.Setting(3) &"news.xml"
		Call KS.WriteTOFile(NewsPath,xmlstr)
	   '===========sitemap================================

	
	response.write("<script language='JavaScript' type='text/JavaScript'>")
	response.write("function yy() {")
	response.write("overstr.innerHTML='<div align=center>恭喜,news.xml生成完毕！<br><br><a href=" & KS.Setting(3) & "news.xml target=_blank>点击查看生成好的news.xml文件</a></div>'; }")
	response.write("</script>")
	end if
	
	'===================================================
		set rs=nothing
End If


response.write("<script language='JavaScript' type='text/JavaScript'>")
response.write("function ll() { ")
response.write("overstr.innerHTML='<div align=center>正在生成，请耐心等待。。。<br></div>'; } ")
response.write("</script>")

set rs=nothing
conn.Close:set conn=nothing
'===================================================结束
Function GetDate(DateStr)
	if KS.G("Action")="creategoogle" then
	GetDate=Year(DateStr) & "-" & Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2)
	else
	GetDate=Year(DateStr) & "-" & Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2)& " " & Right("0" &hour(DateStr),2) &":" & Right("0" &minute(DateStr),2)& ":" & Right("0" & Second(DateStr),2)
	end if
End Function
%>


</head>

<body onLoad="yy()">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td height="25" class="Sort">
 <div align="center"><strong>XML地图生成操作</strong></div></td>
</tr>
</table>
<table width="600" border="0" align="center" cellpadding="6" cellspacing="0">
  <tr>
    <td><div id="overstr"></div></td>
  </tr>
</table>

<form id="form1" name="bqsitemapform" method="post" action="?action=creategoogle">

<table width="600" border="0" align="center" cellpadding="6" cellspacing="0" class="border">
  <tr class="Title">
    <td>★XML地图生成操作</td>
  </tr>
  <tr class="tdbg">
    <td height="17" align="center">
	<a href='http://www.google.com/webmasters/sitemaps/login' target='_blank'><img border=0 src="../images/GoogleSiteMaplogo.gif" /></a>生成符合GOOGLE规范的XML格式地图页面
	<br /></td>
  </tr>
  <tr class="tdbg">
    <td height="18">更新频率：
      <select name="changefreq" id="changefreq">
        <option value="always ">频繁的更新</option>
        <option value="hourly">每小时更新</option>
        <option value="daily" selected="selected">每日更新</option>
        <option value="weekly">每周更新</option>
        <option value="monthly">每月更新</option>
        <option value="yearly">每年更新</option>
        <option value="never">从不更新</option>
      </select></td>
  </tr>
  <tr class="tdbg">
    <td height="35">每个系统调用：
      <input name="prioritynum" type="text" id="prioritynum" value="15" size="6" />条信息内容为最高注意度
	 </td>
  </tr>
  <tr class="tdbg">
    <td height="35">注 意 度：
      <input name="big" type="text" id="big" value="0.5" size="6" />0-1.0之间,推荐使用默认值

	  <br>
  </tr>
</table>
<table width="600" border="0" align="center" cellpadding="6" cellspacing="0">
  <tr>
    <td height="45" align="center"><input name="Submit1"  class="button" onClick="ll();" type="submit" id="Submit1" value="开始生成sitemap" /></td>
  </tr>
</table>
</form>


<form id="form1" name="bqsitemapform" method="post" action="?action=createbaidu">

<table width="600" border="0" align="center" cellpadding="6" cellspacing="0" class="border">
  <tr class="Title">
    <td>★百度新闻开放协议XML生成操作</td>
  </tr>
  <tr class="tdbg">
    <td height="17" align="center">
	<a href='http://news.baidu.com/newsop.html#kg' target='_blank'><img border=0 src="../images/baidulogo.gif" /></a>生成符合百度XML格式的开放新闻协议
	<br /></td>
  </tr>
  <tr class="tdbg">
    <td height="18">更新周期：      
      <input name="changefreq" type="text" id="changefreq" value="15" size="8"> 
      分钟 </td>
  </tr>
  <tr class="tdbg">
    <td height="35">每个系统调用：
      <input name="prioritynum" type="text" id="prioritynum" value="50" size="6" />
      条信息内容为最高注意度(最多100条)	 </td>
  </tr>
</table>
<table width="600" border="0" align="center" cellpadding="6" cellspacing="0">
  <tr>
    <td height="45" align="center"><input name="Submit1"  class="button" onClick="ll();" type="submit" id="Submit1" value="开始生成sitemap" /></td>
  </tr>
</table>
</form>

<br />
</body>
</html>
