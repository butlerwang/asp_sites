<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/ClubFunction.asp"-->
<!--#include file="../plus/md5.asp"-->
<%

response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"
Dim KS:Set KS=New PublicCls
Set KSUser = New UserCls
Call KSUser.UserLoginChecked()
Dim ChannelID,InfoID,RS,CommentStr,UserIP,Total,TitleStr,TitleLinkStr,TotalPoint,N,DomainStr,Title
Dim totalPut, MaxPerPage,PageNum,SqlStr,PrintOut,CommentXML,PostId,PostTable,Tid,Fname
ChannelID=KS.Chkclng(KS.S("ChannelID"))

IF ChannelID=0 And KS.S("Action")<>"Support" And KS.S("Action")<>"QuoteSave" Then KS.Die ""
PrintOut=KS.S("PrintOut")

InfoID=KS.ChkClng(KS.S("InfoID"))
DomainStr=KS.GetDomain
Const BasicType=1  '定义关联论坛的模型基型
Select Case KS.S("Action")
 Case "Show"  Call Show()
 Case "Write"
  If KS.ChkClng(KS.C_S(ChannelID,12))=0 and channelid<>1000 Then Response.end()
  Call Ajax()
  Response.Write("document.write('" & GetWriteComment(ChannelID,InfoID) & "');")
 Case "WriteSave"  Call WriteSave()
 Case "Support"  
  If PrintOut="js" Then
   Response.Write "ShowSupportMessage('" & Support() & "');"
  Else
   Response.Write Support()
  End If
 Case "QuoteSave" Call QuoteSave()
 Case Else  Call CommentMain()
 End Select
 Set KS=Nothing
 Set KSUser=Nothing
 
Sub Ajax()
 %>
function xmlhttp()
{
	if(window.XMLHttpRequest){
		return new XMLHttpRequest();
	} else if(window.ActiveXObject){
		return new ActiveXObject("Microsoft.XMLHTTP");
	} 
	throw new Error("XMLHttp object could be created.");
}
	
var loader=new xmlhttp;
function ajaxLoadPage(url,request,method,fun)
{
	method=method.toUpperCase();
	if (method=='GET')
	{
		urls=url.split("?");
		if (urls[1]=='' || typeof urls[1]=='undefined')
		{
			url=urls[0]+"?"+request;
		}
		else
		{
			url=urls[0]+"?"+urls[1]+"&"+request;
		}
		
		request=null;
	}
	loader.open(method,url,true);
	if (method=="POST")
	{
		loader.setRequestHeader("Content-Type","application/x-www-form-urlencoded");
	}
	loader.onreadystatechange=function(){
	     eval(fun+'()');
	}
	loader.send(request);
}

function formToRequestString(form_obj)
{
    var query_string='';
    var and='';
    for (var i=0;i<form_obj.length;i++ )
    {
        e=form_obj[i];
        if (e.name) {
            if (e.type=='select-one') {
                element_value=e.options[e.selectedIndex].value;
            } else if (e.type=='select-multiple') {
                for (var n=0;n<e.length;n++) {
                    var op=e.options[n];
                    if (op.selected) {
                        query_string+=and+e.name+'='+escape(op.value);
                        and="&"
                    }
                }
                continue;
            } else if (e.type=='checkbox' || e.type=='radio') {
                if (e.checked==false) {   
                    continue;   
                }   
                element_value=e.value;
            } else if (typeof e.value != 'undefined') {
                element_value=e.value;
            } else {
                continue;
            }
            query_string+=and+e.name+'='+escape(element_value);
            and="&"
        }
    }
    return query_string;
}
function ajaxFormSubmit(form_obj,fun)
{ 
	ajaxLoadPage(form_obj.getAttributeNode("action").value,formToRequestString(form_obj),form_obj.method,fun)
}
 <%
 End Sub
 
 Sub CommentMain
	Dim KSRCls,FileContent
	Set KSRCls = New Refresh
	FCls.RefreshType = "Comment" '设置刷新类型，以便取得当前位置导航等

	if KS.C_S(ChannelID,15)="" then KS.Die "请先到模型设置里绑定评论页模板!"
	FileContent = KSRCls.LoadTemplate(KS.C_S(ChannelID,15))
	If Trim(FileContent) = "" Then FileContent = "模板不存在!"
	FileContent=Replace(FileContent,"{$GetShowComment}","<script src=""" & domainstr & "ks_inc/Comment.page.js"" language=""javascript""></script><script src=""" & domainstr & "ks_inc/Kesion.box.js"" language=""javascript""></script><script language=""javascript"" defer>var from3g=0;Page(1," & ChannelID & ",'" & InfoID & "','Show','"& domainstr & "');</script><div id=""c_" & InfoID & """></div><div id=""p_" & InfoID & """ align=""right""></div>")

	if channelid<>8 then
	 if conn.execute("select count(id) from " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID).eof then 
	 KS.Die "<script>alert('对不起，已删除 ！');window.close();</script>"
	end if
	if conn.execute("select comment from " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID)(0)=0 then KS.Die "<script>alert('对不起，不允许评论 ！');window.close();</script>"
	end if
	
 TitleStr=conn.execute("Select top 1 Title From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID)(0)

  FileContent=Replace(FileContent,"{$GetTitle}",TitleStr)
  FileContent=Replace(FileContent,"{$GetWriteComment}","<script language=""javascript"" src=""?Action=Write&ChannelID=" & ChannelID& "&InfoID=" & InfoID & """></script>")
	FileContent = KSRCls.ReplaceLableFlag(KSRCls.ReplaceAllLabel(FileContent))
	FileContent = KSRCls.ReplaceGeneralLabelContent(FileContent) '替换通用标签
	Set KSRCls = Nothing
   Response.Write(FileContent)
End Sub

Sub Show()
	MaxPerPage=5    '每页显示评论条数
	If Request.ServerVariables("HTTP_REFERER")<>"" Then 
	  If Instr(Lcase(Request.ServerVariables("HTTP_REFERER")),"comment.asp")<>0 Then MaxPerPage=20
	End If
    If KS.FoundInArr(BasicType,KS.C_S(ChannelID,6),",") Then
     SqlStr="Select top 1 ID,Title,Tid,Fname,PostId,PostTable From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID
    ElseIf ChannelID=1000 Then
     SqlStr="Select top 1 ID,subject as Title,classid as tid,0,0,0 From KS_GroupBuy Where ID=" & InfoID
	Else
     SqlStr="Select top 1 ID,Title,Tid,Fname,0,0 From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID
	End If
	Set RS=Server.CreateObject("ADODB.RECORDSET")
	 RS.Open SqlStr,Conn,1,1
	 If Not RS.Eof Then
	   PostTable=RS(5) : PostId=KS.ChkClng(RS(4)) : TitleStr=RS(1)
	   If ChannelID=1000 Then
	   'TitleLinkStr="<a href='" & KS.GetDomain & "shop/groupbuyshow.asp?id=" & InfoID & "' target='_blank'>" & TitleStr & "</a>"
	   TitleLinkStr=TitleStr
	   ElseIf PostID=0 Then
	   TitleLinkStr="<a href='" & KS.GetItemUrl(ChannelID,RS(2),rs(0),rs(3)) & "' target='_blank'>" & TitleStr & "</a>"
	   Else
	   TitleLinkStr="<a href='" & KS.GetClubShowUrl(PostId) & "' target='_blank'>" & TitleStr & "</a>"
	   End If
	 Else
	   RS.Close:Set RS=Nothing
	   KS.Die ""
	 End If
     CurrentPage = KS.ChkClng(KS.S("page"))
	 If CurrentPage<=0 Then CurrentPage=1
	 RS.Close
	 If PostId<>0 Then
        RS.Open "Select b.userface,0 as anonymous,a.* From " & PostTable & " a left join KS_User b on a.username=b.username Where a.Verific=1 And a.TopicID=" & PostId & " and a.parentid<>0 Order By ID Desc",conn,1,1
     Else
	   RS.Open "Select b.userface,b.userid,a.* From KS_Comment a left join KS_User b on a.username=b.username Where ProjectID=0 and a.Verific=1 And a.ChannelID=" & ChannelID & " And a.InfoID=" & InfoID & " Order By ID Desc",conn,1,1
	 end If
	 
  IF Not Rs.Eof Then
  
       If PostId<>0 Then
		 totalPut = Conn.Execute("Select Count(ID) From "& PostTable & " Where Verific=1 And parentid<>0 and TopicId=" & PostID)(0)
	   Else
		 totalPut = Conn.Execute("Select Count(ID) From KS_Comment Where ProjectID=0 and Verific=1 And ChannelID=" & ChannelID & " And InfoID=" & InfoID)(0)
	   End If
				        If CurrentPage < 1 Then	CurrentPage = 1
						If (totalPut Mod MaxPerPage) = 0 Then
									PageNum = totalPut \ MaxPerPage
						Else
									PageNum = totalPut \ MaxPerPage + 1
						End If
		
				         If CurrentPage >1 And (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
				         End If
						 Set CommentXML=KS.ArrayToxml(Rs.GetRows(MaxPerPage),Rs,"row","xml")
						 Call showContent()

  Else
	CommentStr=""
  End If
  Rs.Close:Set Rs=Nothing
  
  If KS.C_S(ChannelID,12)=0 and channelid<>1000 Then TotalPut=0
  If PrintOut="js" Then
   Response.Write "show(""" & replace(replace(CommentStr,vbcrlf,"\n"),"""","\""") & "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|条||2"");"
  Else
   Response.Write CommentStr & "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|条||2"
  End If
End Sub

Sub ShowContent()
   If KS.C_S(ChannelID,12)=0 and channelid<>1000 Then Exit Sub
	Set KSRCls = New Refresh
	If KS.S("from3g")<>"1" Then CommentStr="<br /> &nbsp;以下是对 <strong>[" & TitleLinkStr & "]</strong> 的评论,"
	CommentStr=CommentStr &"总共:<font color=red>" & totalPut & " </font>条评论<br />"
    CommentStr=CommentStr & "<table  width='98%' border='0' align='center' cellpadding='0' cellspacing='1'>"
	   	


  If CurrentPage=1 Then	N=TotalPut	Else N=totalPut-MaxPerPage*(CurrentPage-1)
  Dim FaceStr,Publish,QuoteContentj,Content,Node,UserFace,ID,ReplyContent,ReplayTime,Opposition,Support
  
  If IsObject(CommentXML) Then
   For Each Node In CommentXML.DocumentElement.SelectNodes("row")
		FaceStr= KS.GetDomain &  "images/face/boy.jpg"
		ID=Node.SelectSingleNode("@id").text
		If PostId=0 Then
		   ReplayTime=Node.SelectSingleNode("@adddate").text
		   Opposition=Node.SelectSingleNode("@oscore").text
		   Support=Node.SelectSingleNode("@score").text
		   ReplyContent=Node.SelectSingleNode("@replycontent").text
		   IF Node.SelectSingleNode("@anonymous").text="0" Then
			Publish=Node.SelectSingleNode("@username").text
			UserFace=Node.SelectSingleNode("@userface").text
			Publish="会员:<a href=""" & KS.GetSpaceUrl(Node.SelectSingleNode("@userid").text) & """ target=""_blank"">" & publish & "</a>"
		   Else
			Publish= "游客："& Node.SelectSingleNode("@username").text
		   End IF
		   QuoteContent=Node.SelectSingleNode("@quotecontent").text
	   Else
	       UserFace=Node.SelectSingleNode("@userface").text
		   Publish="<a href=""" & KS.GetSpaceUrl(Node.SelectSingleNode("@userid").text) & """ target=""_blank"">" & Node.SelectSingleNode("@username").text & "</a>"
	       ReplayTime=Node.SelectSingleNode("@replaytime").text
		   Opposition=Node.SelectSingleNode("@opposition").text
		   Support=Node.SelectSingleNode("@support").text
	   End If
		If Not KS.IsNul(UserFace) and Node.SelectSingleNode("@anonymous").text<>"1" Then
				FaceStr=UserFace
				If lcase(left(FaceStr,4))<>"http" and left(facestr,1)<>"/" then FaceStr=KS.GetDomain & FaceStr
		End If
	   If Not KS.IsNUL(QuoteContent) Then
	   QuoteContent=Replace(QuoteContent,"[quote]","<div style='margin:2px;border:1px solid #cccccc;background:#FFFFEE;padding:4px'>")
	   QuoteContent=Replace(QuoteContent,"[/quote]","</div>")
	   QuoteContent=Replace(QuoteContent,"[dt]","<div style='padding-left:10px;color:#999999'>")
	   QuoteContent=Replace(QuoteContent,"[/dt]","</div>")
	   QuoteContent=Replace(QuoteContent,"[dd]","<div style='padding-left:10px;'>")
	   QuoteContent=Replace(QuoteContent,"[/dd]","</div>")
	   End If
	  ' Content = KS.HtmlCode(ReplaceFace(QuoteContent & Node.SelectSingleNode("@content").text))
	   Content = KS.HtmlCode(ReplaceFace(QuoteContent & Node.SelectSingleNode("@content").text))
	   If PostId<>0 Then
	    Content=KSRCls.ScanAnnex(KSRCls.UbbCode(Content,n))
	   End If
	   

	   CommentStr=CommentStr & "<tr>"
	   CommentStr=CommentStr & "<td width='70' rowspan='3' style='margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;'><img width='60' height='60' alt='" & Node.SelectSingleNode("@username").text &"' onerror=this.src='" &  KS.GetDomain &  "user/images/noavatar_middle.gif'; src='" & facestr & "' border='1'></td>"
	   CommentStr=CommentStr & "<td height='25' width='*'>"
	   CommentStr=CommentStr & publish
	   CommentStr=CommentStr  & " <font color='#999999'>(发表时间： " & ReplayTime &")</font> </td><td><font style='font-size:32px;font-family:Arial Black;color:#EEF0EE'> " & N & "</font> </td>"
	   CommentStr=CommentStr & "</tr>"
	   CommentStr=CommentStr & "<tr><td height='25' colspan='2' style='word-break:break-all;'>" & Content
	   If ReplyContent<>"" Then
	   CommentStr=CommentStr & "<div style='padding:4px;color:red;border:1px solid #ccc;background:#FFFFEE;'>""" & Node.SelectSingleNode("@replyuser").text & """回复:" & ReplyContent & "</div>"
	   End If
	   
	   CommentStr=CommentStr & "</td></tr>"
	   CommentStr=CommentStr & "<tr>"
	   If KS.S("from3g")<>"1" Then 
	   CommentStr=CommentStr & "<td style='margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;' height='25' colspan='2'><div style='text-align:right'><a href='javascript:void(0)' onclick=reply("& PostId &","& ChannelID & "," & ID & ",'" & KS.GetDomain & "');>盖楼(回复)</a> <a href='javascript:void(0)' onclick=javascript:Support(" & PostId & "," & ID & ",1,'" &KS.GetDomain & "');><span style='color:brown'>支持</span>[" & Support & "]</a> <a href='javascript:void(0)' onclick=javascript:Support(" & PostId & "," & ID & ",0,'" & KS.GetDomain & "');return false>反对[" & Opposition & "]</a></div> </td>"
	   Else
	   CommentStr=CommentStr & "<td style='margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;' height='25' colspan='2'><div style='text-align:right'><a href='javascript:void(0)' onclick=javascript:Support(" & PostId & "," & ID & ",1,'" &KS.GetDomain & "');><span style='color:brown'>支持</span>[" & Support & "]</a> <a href='javascript:void(0)' onclick=javascript:Support(" & PostId & "," & ID & ",0,'" & KS.GetDomain & "');return false>反对[" & Opposition & "]</a></div> </td>"
	   End If
	   CommentStr=CommentStr & "</tr>"
	   N=N-1
   Next
 End If
   CommentStr=CommentStr & "</table>"
	Set KSRCls=Nothing
End Sub
 
 '发表评论
Function GetWriteComment(ChannelID,InfoID)
%>
function success()
{
	var loading_msg='\n\n\t请稍等，正在提交评论...';
	var C_Content=document.getElementById('C_Content');
	
 	if (loader.readyState==1){C_Content.value=loading_msg;}
	if (loader.readyState==4)
		{   var s=loader.responseText;
			if (s=='ok')
			 {alert('恭喜,你的评论已成功提交！');
			  if (typeof(loadDate)!="undefined") loadDate(1);
			  leavePage();
			 }else{alert(s);
			  C_Content.value=document.getElementById('sC_Content').value;
			 }
		}
}
var OutTimes =11;
function leavePage()
{
	if (OutTimes==0)
	 {
	 document.getElementById('C_Content').disabled=false;
	 document.getElementById('SubmitComment').disabled=false;
	 document.getElementById('C_Content').value=''
	 <%If KS.C_S(ChannelID,13)="1" Then%>
	  document.form1.Verifycode.value='';
	 <%end if%>
	 <%If KS.C_S(ChannelID,14)<>0  Then%>
	 document.getElementById('cmax').value=<%=KS.C_S(ChannelID,14)%>;
	 <%end if%>
	 OutTimes =11;
	 return;
	 }
	else {
	    document.getElementById('C_Content').disabled=true;
		document.getElementById('SubmitComment').disabled=true;
		OutTimes -= 1;
		document.getElementById('C_Content').value ="\n\n\t评论已提交，等待 "+ OutTimes + " 秒钟后您可继续发表...";
		setTimeout("leavePage()", 1000);
		}
	}
function checklength(cobj)
{ 
	var cmax=<%=KS.C_S(ChannelID,14)%>;
	if (cobj.value.length>cmax) {
	cobj.value = cobj.value.substring(0,cmax);
	alert("评论不能超过"+cmax+"个字符!");
	}
	else {
	document.getElementById('cmax').value = cmax-cobj.value.length;
	}
}

   function checkform()
   {
	var anounname=document.getElementById('AnounName');
	var C_Content=document.getElementById('C_Content');
	var sC_Content=document.getElementById('sC_Content');
	var anonymous=document.getElementById('Anonymous');
	var pass=document.getElementById('Pass');
   if (anounname.value==''){
        alert('请填写用户名。');
		anounname.focus();
        return false;
     }
	if (anonymous.checked==false && pass.value==''){
	   alert('请输入密码或选择游客发表！');
	   pass.focus();
	   return false;
	}
	<%If KS.C_S(ChannelID,13)="1" Then%>
   if (document.form1.Verifycode.value==''){
	   alert('请入验证码!');
	   document.form1.Verifycode.focus();
	   return false;
    }
	<%end if%>
   if (C_Content.value==''||C_Content.value=='文明上网，请对您的发言负责！'){
	   alert('请填写评论内容!');
	   C_Content.focus();
	   return false;
    }
	sC_Content.value=C_Content.value;
	try{ajaxFormSubmit(document.form1,'success');
	 }catch(e){
	  document.form1.action="<%=DomainStr%>plus/Comment.asp?Action=WriteSave&flag=NotAjax";
	  document.form1.submit();
	 }
	 
	 
	}
<%
		 GetWriteComment = GetWriteComment & "<style>.comment_write_table,.comment_write_table textarea,.comment_write_table a{color:#666}.comment_write_table textarea,.comment_write_table .textbox{padding:2px;color:#999;border:1px solid #cccccc;}</style><table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""comment_write_table""><form name=""form1"" action=""" & DomainStr &"plus/Comment.asp?Action=WriteSave"" method=""post""><input type=""hidden"" value=""" & ChannelID & """ name=""ChannelID""><input type=""hidden"" value=""" & InfoID & """ name=""InfoID"">"
		GetWriteComment = GetWriteComment & "<tr>"
		GetWriteComment = GetWriteComment & "  <td style=""padding:10px;"">"
		Dim PostNum,PostId
		PostId=KS.ChkClng(Request.QueryString("postId"))
		If PostId<>0 Then
		  PostNum=Conn.Execute("Select top 1 TotalReplay From KS_GuestBook Where ID=" & PostId)(0)
		Else
		  PostNum=Conn.Execute("Select count(1) From KS_Comment Where ProjectID=0 and Verific=1 and ChannelID=" & ChannelID & " And InfoID=" & InfoID)(0)
		End If
		GetWriteComment = GetWriteComment & "  <div style=""font-size:14px;height:30px;line-height:30px;text-align:left;""><strong>已有 <span style=""color:brown;font-weight:bold"">" & PostNum & "</span> 条跟帖</strong>"
		If ChannelID<>1000 and request("from3g")="" Then
			If PostId<>0 Then
			GetWriteComment = GetWriteComment & "<a href=""" & KS.GetClubShowUrl(PostId) & """ style=""color:brown"" target=""_blank"">(点击查看)</a></div>"
			Else
			GetWriteComment = GetWriteComment & "<a href=""" & DomainStr &"plus/Comment.asp?ChannelID=" & ChannelID & "&InfoID=" & InfoID & """ style=""color:brown"">(点击查看)</a></div>"
			End If
		Else
			GetWriteComment = GetWriteComment & "</div>"
		End If
		
		If KS.C_S(ChannelID,14)<>0  Then
		GetWriteComment = GetWriteComment & "<textarea onkeydown=""checklength(this);"" onkeyup=""checklength(this);"" name=""C_Content"" rows=""6"" id=""C_Content"" onfocus=""if(this.value==\'文明上网，请对您的发言负责！\'){this.value=\'\'}"" wrap=""PHYSICAL"" onblur=""if(this.value==\'\'){this.value=\'文明上网，请对您的发言负责！\'}"" style=""overflow:auto;font-size:14px;width:100%"">文明上网，请对您的发言负责！</textarea>"
		Else
		GetWriteComment = GetWriteComment & "<textarea style=""font-size:14px;padding:5px;width:98%;height:90px;overflow:auto;"" onfocus=""if(this.value==\'文明上网，请对您的发言负责！\'){this.value=\'\'}"" wrap=""PHYSICAL"" onblur=""if(this.value==\'\'){this.value=\'文明上网，请对您的发言负责！\'}"" name=""C_Content"" rows=""4"" id=""C_Content"">文明上网，请对您的发言负责！</textarea>"
		End If
		
		GetWriteComment = GetWriteComment & "</td></tr>"
		GetWriteComment = GetWriteComment & "  <tr><td nowrap>"
		GetWriteComment = GetWriteComment & "  <div style=""margin-left:10px;text-align:left;"">"
		If KSUser.UserName="" Then
		GetWriteComment = GetWriteComment & " 用户名：<input onfocus=""if(this.value==\'匿名\'){this.value=\'\';}"" class=""textbox"" maxlength=15 name=""AnounName"" type=""text"" id=""AnounName"" value=""匿名"" style=""width:70px""/> <a href=""" & DomainStr & "user/reg/""><u>注册</u></a>"
		Else
		GetWriteComment = GetWriteComment & "   <span style=""display:none"">用户名：<input class=""textbox"" maxlength=15 name=""AnounName"" type=""text"" id=""AnounName"" value=""" & KSUser.username & """ style=""width:70px""/></span>欢迎您，" & KSUser.UserName &"! <a href="""
		if request("from3g")="1" then 
		GetWriteComment = GetWriteComment & "user.asp"
		Else
		GetWriteComment = GetWriteComment & DomainStr & "user/"
		End If
		GetWriteComment = GetWriteComment & """>[会员中心]</a> <a onclick=""return(confirm(\'确认退出吗？\'));"" href=""" & DomainStr & "user/UserLogout.asp"">[退出]</a>"
		End If
		Dim Style,Check
		If KS.C_S(ChannelID,12)="1" or KS.C_S(ChannelID,12)="2" Then
		 If KS.IsNul(KS.C("UserName"))  Then style="": else Style=" style=""display:none"""
		 checked=""
		Else
		 Style=" style=""display:none""":checked=" checked"
		End If
		
		GetWriteComment = GetWriteComment & "<span id=""pp""" & style & "> 密码：<input class=""textbox"" name=""Pass"" size=""8"" type=""password"" id=""Pass"" value=""" & KSUser.PassWord & """ ></span>"

		If KS.C_S(ChannelID,13)="1" and channelid<>1000 Then
		if request("from3g")="1" then GetWriteComment = GetWriteComment & "<br/>"
		GetWriteComment = GetWriteComment & "&nbsp;认证码：<script>writeVerifyCode(""" & KS.GetDomain &""",0)</script>"
		End IF
		
		If KS.C("UserName")="" Then
		GetWriteComment = GetWriteComment & "<span id=""nm"">"
		Else
		GetWriteComment = GetWriteComment & "<span id=""nm"" style=""display:none"">"
		End If

		If KS.C_S(Channelid,12)="1" Or KS.C_S(Channelid,12)="2" Then
		GetWriteComment = GetWriteComment & "<span style=""display:none"">"
		Else
		GetWriteComment = GetWriteComment & "<span>"
		End iF
		GetWriteComment = GetWriteComment & "<label><input onclick=""if(this.checked==true){document.getElementById(\'Pass\').disabled=true;document.getElementById(\'pp\').style.display=\'none\';}else{if(document.getElementById(\'AnounName\').value==\'匿名\'){document.getElementById(\'AnounName\').value=\'\';}document.getElementById(\'Pass\').disabled=false;document.getElementById(\'pp\').style.display=\'\';}"" type=""checkbox""" & checked & " value=""1"" name=""Anonymous"" id=""Anonymous"">匿名</label></span>"
		GetWriteComment = GetWriteComment & "</span>"

		If KS.C_S(ChannelID,14)<>0  Then
		GetWriteComment = GetWriteComment & "&nbsp;字数：<input disabled class=""textbox"" type=""text"" id=""cmax"" size=""3"" name=""cmax"" value=""" & KS.C_S(ChannelID,14) & """/>"
		End If
		if request("from3g")="1" then GetWriteComment = GetWriteComment & "<br/>"
		GetWriteComment = GetWriteComment & "<input type=""hidden"" name=""sC_Content"" id=""sC_Content""><input type=""submit"" id=""SubmitComment"" name=""SubmitComment"" value=""确认发表"" style=""padding:2px"" onclick=""checkform();return false""/></div>"
		
		GetWriteComment = GetWriteComment & "</td></tr></form></table>"
		
		End Function  
		
		Function ReplaceFace(c)
		 C=Replace(Replace(C,chr(10),"<br/>"),"  ","&nbsp;&nbsp;")
		 Dim str:str=":)|:(|:D|:'(|:@|:o|:P|:$|;P|:L|:Q|:lol|:loveliness:|:funk:|:curse:|:dizzy:|:shutup:|:sleepy:|:hug:|:victory:|:time:|:kiss:|:handshake|:call:|55555|"
		 Dim strArr:strArr=Split(str,"|")
		 Dim K,NS
		 For K=1 To 9
		  c=replace(c,"[e" & k & "]","[e0" & k & "]")
		 Next
		 For K=1 To 24
		  NS=Right("0" & K,2)
		  c=replace(c,"[e"&ns &"]","<img title='" & strarr(k) & "' src='" & DomainStr & "editor/ubb/images/smilies/default/" & NS & ".gif'>")
		 Next
		 C=KS.FilterIllegalChar(C)
		 ReplaceFace=C
		End Function
		
'保存发表
Sub WriteSave()	
		Dim UserName,C_Content,Verific,Anonymous,point,VerifyCode,Pass,Flag,ComeUrl,GroupID,LoginTF,PostId,PostTable
		Flag=KS.S("Flag")
		ComeUrl=Request.ServerVariables("HTTP_REFERER"):If ComeUrl="" Then ComeUrl=KS.GetDomin
		LoginTF=Cbool(KSUser.UserLoginChecked)
		If ChannelID=1000 Then '团购
		 If Conn.Execute("Select top 1 id From KS_GroupBuy Where Comment>=1 and ID=" & InfoID).Eof Then
		 If Flag="NotAjax" Then KS.Die "<script>alert('对不起,本团购不允许评论');location.href='" & ComeUrl & "';</script>" Else  KS.Die "对不起,本团购不允许评论！"
		 End If
		ElseIf KS.ChkClng(KS.C_S(Channelid,12))=0 Then 
		 If Flag="NotAjax" Then KS.Die "<script>alert('对不起,本信息不允许评论');location.href='" & ComeUrl & "';</script>" Else  KS.Die "对不起,本信息不允许评论！"
		End If	  

		AnounName=KS.R(KS.S("AnounName"))
		If LoginTF=false And Len(AnounName)>20 Or Len(AnounName)<2 Then
		 If Flag="NotAjax" Then KS.Die "<script>alert('用户名不符合规范，长度限制在2-20之间!');location.href='" & ComeUrl & "';</script>" Else KS.Die "用户名不符合规范，长度限制在2-20之间!"
		End If
		Pass=KS.R(KS.G("Pass"))
		C_Content=KS.S("C_Content")
		VerifyCode=KS.S("VerifyCode")
		
		Anonymous=KS.ChkClng(KS.S("Anonymous"))
		point=KS.ChkClng(KS.S("point"))
		If ChannelID<>1000 AND KS.C_S(ChannelID,13)="1" and lcase(Trim(Request.Form("Verifycode")))<>lcase(Trim(Session("Verifycode"))) Then
		 If Flag="NotAjax" Then KS.Die "<script>alert('验证码有误，请重新输入!');history.back();</script>" Else KS.Die ("验证码有误，请重新输入！")
		End IF
		  
		IF Anonymous=0 Then
		  if LoginTF=false  then
		     	if Pass="" Then 
				  If Flag="NotAjax" Then KS.Die "<script>alert('请填写登录密码或选择游客发表。');history.back();</script>" Else KS.Die("请填写登录密码或选择游客发表。")
				End if
             Pass=Md5(Pass,16)
		     Dim UserRS:Set UserRS=Server.CreateObject("Adodb.RecordSet")
			 UserRS.Open "Select top 1 UserID,UserName,PassWord,Locked,Score,LastLoginIP,LastLoginTime,LoginTimes,RndPassword,GroupID From KS_User Where UserName='" &AnounName & "' And PassWord='" & Pass & "'",Conn,1,3
			 If UserRS.Eof And UserRS.BOf Then
				  UserRS.Close:Set UserRS=Nothing
				  If Flag="NotAjax" Then KS.Die "<script>alert('你输入的用户名或密码有误，请重新输入!');history.back();</script>"Else  KS.Die("你输入的用户名或密码有误，请重新输入!")
			 ElseIf UserRS("Locked")=1 Then
				  If Flag="NotAjax" Then KS.Die "<script>alert('您的账号已被管理员锁定，请与管理员联系!');history.back();</script>" Else  KS.Die("您的账号已被管理员锁定，请与管理员联系!")
			 Else
			            GroupID=UserRS("GroupID")
			            '登录成功，更新用户相应的数据
						Dim RndPassword:RndPassword=KS.R(KS.MakeRandomChar(20))
						If datediff("n",UserRS("LastLoginTime"),now)>=KS.Setting(36) then '判断时间
						UserRS("Score")=UserRS("Score")+KS.Setting(37)
						end if
						UserRS("LastLoginIP") = KS.GetIP
                        UserRS("LastLoginTime") = Now()
                        UserRS("LoginTimes") = UserRS("LoginTimes") + 1
						UserRS("RndPassWord")=RndPassWord
                        UserRS.Update
						If EnabledSubDomain Then
							 Response.Cookies(KS.SiteSn).domain=RootDomain					
						Else
                             Response.Cookies(KS.SiteSn).path = "/"
						End If
						Response.Cookies(KS.SiteSn)("UserName") = AnounName
						Response.Cookies(KS.SiteSn)("Password") = Pass
						Response.Cookies(KS.SiteSN)("RndPassword")= RndPassword
						Response.Cookies(KS.SiteSn)("UserID") = UserRS("UserID")
			end if
			UserRS.Close : Set UserRS=Nothing
		  Else
		     groupid=KSUser.GroupID
		  end if
		Else
		    Dim RSG:Set RSG=Conn.Execute("select top 1 groupid from KS_User Where UserName='" & AnounName & "'")
			If Not RSG.Eof Then
			  groupID=rsg(0)
			End If
			RSG.Close : Set RSG=Nothing
		End IF
		
		if KS.ChkClng(KS.C_S(Channelid,12))=1 Or KS.ChkClng(KS.C_S(ChannelID,12))=2 then
		  if KS.C("UserName")="" or KS.C("PassWord")=""  then
				  If Flag="NotAjax" Then KS.Die "<script>alert('对不起，系统设置不允许游客发表。');history.back();</script>" Else KS.Die("对不起，系统设置不允许游客发表。")
		  End If
		End If

		IF InfoID="" Then 
			 If Flag="NotAjax" Then KS.Die "<script>alert('参数传递有误!');history.back();</script>" Else KS.Die ("参数传递有误!")
		End if
		if AnounName="" Then
			 If Flag="NotAjax" Then KS.Die "<script>alert('请填写你的昵称!');history.back();</script>" Else KS.Die("请填写你的昵称!")
		End if
		if C_Content="" Then 
			 If Flag="NotAjax" Then KS.Die "<script>alert('请填写评论内容!');history.back();</script>" Else KS.Die("请填写评论内容!")
		End if
		If Len(C_Content)>KS.ChkClng(KS.C_S(ChannelID,14)) and KS.ChkClng(KS.C_S(ChannelID,14))<>0 Then
			 If Flag="NotAjax" Then KS.Die "<script>alert('评论内容必须在" &KS.C_S(ChannelID,14) & "个字符以内!');history.back();</script>" Else KS.Die("评论内容必须在" &KS.C_S(ChannelID,14) & "个字符以内!")
		End if
		
		if ks.c("username")<>"" then Anonymous=0

		Set RS=Server.CreateObject("ADODB.RECORDSET")
		If KS.FoundInArr(BasicType,KS.C_S(ChannelID,6),",") Then
		 RS.Open "Select top 1 Title,PostId,PostTable,Tid,Fname From " & KS.C_S(ChannelID,2) &" Where id=" & InfoID,Conn,1,1
		ElseIF ChannelID=1000 Then
		 RS.Open "Select top 1 subject as Title,0,0,classid  as Tid,id as Fname From KS_GroupBuy Where id=" & InfoID,Conn,1,1
		Else
		 RS.Open "Select top 1 Title,0,0,Tid,Fname From " & KS.C_S(ChannelID,2) &" Where id=" & InfoID,Conn,1,1
		End If
		If RS.Eof And RS.Bof Then
		  RS.Close:Set RS=Nothing
		  If Flag="NotAjax" Then KS.Die "<script>alert('内容不存在!');history.back();</script>" Else KS.Die("内容不存在!")
		End IF
		PostId=KS.ChkClng(RS(1)) : PostTable=RS(2):Title=RS("Title"):Tid=rs(3):Fname=RS(4)
		RS.Close
		Set RS=Nothing
		Call DoWriteSave(0,PostID,InfoID,AnounName,C_Content,"",KSUser,Anonymous)
	    If Flag="NotAjax" Then KS.Die "<script>alert('评论发表成功!');location.href='" & ComeUrl & "';</script>" Else KS.Die "ok"
End Sub

Sub DoWriteSave(IsQuote,PostID,InfoID,AnounName,C_Content,QuoteContent,KSUser,Anonymous)
     Dim BoardID,O_LastPost,N_LastPost,UserID,BSetting,Verific,LoginTF,RS
	 LoginTF=Cbool(KSUser.UserLoginChecked)
     if KS.ChkClng(KS.C_S(Channelid,12))=1 Or KS.ChkClng(KS.C_S(ChannelID,12))=3 then verific=0 else verific=1
	 If KS.ChkClng(KS.C_S(Channelid,12))=5 Then
	  If KS.IsNul(KS.C("UserName")) And KS.IsNul(KS.C("PassWord")) Then verific=0 else verific=1
	 End If
	 if channelid=1000 then
	  dim rsg:set rsg=conn.execute("select top 1 comment from ks_groupbuy where id=" & infoid)
	  if rsg.eof then
	    rsg.close:set rsg=nothing
	    exit sub
	  else
	    if rsg("comment")=0 then
	    rsg.close:set rsg=nothing
	    exit sub
		elseif rsg("comment")=1 then
		 verific=0
		else
		 verific=1
		end if
	  end if
	  rsg.close:set rsg=nothing
	 end if
	 
     If PostId<>0 Then '绑定论坛帖子
	     Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select top 1 BoardID,PostTable From KS_GuestBook Where ID=" & PostId,conn,1,1
		  If RS.Eof And RS.Bof Then
		    RS.CLose:Set RS=Nothing
		    If Flag="NotAjax" Then KS.Die "<script>alert('帖子内容不存在!');history.back();</script>" Else KS.Die("帖子内容不存在!")
		  End If
		  PostTable=RS("PostTable"):BoardID=RS("BoardID")
		  RS.Close
		  If IsQuote=1 Then  '引用
		   RS.Open "Select top 1 * From " & PostTable  & " Where ID=" & KS.ChkClng(KS.S("quoteId")),Conn,1,1
		   If RS.Eof And RS.Bof Then
		    RS.CLose:Set RS=Nothing
		    If Flag="NotAjax" Then KS.Die "<script>alert('引用的帖子内容不存在!');history.back();</script>" Else KS.Die("引用的帖子内容不存在!")
		   End If
		   C_Content="[quote]以下是引用 " & RS("UserName") & " 在" & RS("ReplayTime") & " 的发言：[br]"& RS("Content") &"[/quote]" & C_Content
		   RS.Close
		  End If
		  
		  UserID=KS.ChkClng(KSUser.GetUserInfo("UserID"))
		  If BoardID<>0 Then
			 KS.LoadClubBoard()
			 Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
			 BSetting=Node.SelectSingleNode("@settings").text
		  End If
		  BSetting=BSetting & "$$$0$0$0$0$0$0$1$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
		  BSetting=Split(BSetting,"$")
		  Call InsertReply(PostTable,AnounName,UserID,PostId,C_Content,0,0,PostId,verific,SQLNowString) '写入论坛回复表
		  Conn.Execute("Update KS_GuestBook Set LastReplayTime=" & SqlNowString &",LastReplayUser='" & AnounName &"',LastReplayUserID=" & UserID & ",TotalReplay=TotalReplay+1 where id=" & PostId)
		  N_LastPost=PostId & "$" & now & "$" & Replace(Title,"$","") &"$" & AnounName & "$" &UserID&"$$"
           If KS.ChkClng(BSetting(4))>0 and LoginTF=true Then
				 Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(BSetting(4)),"系统","在论坛回复主题[" & Title & "]所得!",0,0)
		   End If
		  
		   '更新版面数据
			If BoardID<>0 Then
			  KS.LoadClubBoard()
			  O_LastPost=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@lastpost").text
			  Call UpdateBoardPostNum(0,BoardID,Verific,O_LastPost,N_LastPost)
			End If
			UpdateTodayPostNum '更新今日发帖数等
		Else
		     If KS.ChkClng(Split(KS.C_S(KS.G("ChannelID"),46)&"||||","|")(4))<>0 Then 
			  If not Conn.Execute("Select top 1 * From KS_Comment Where ProjectID=0 and InfoID=" & InfoID & " and UserIp='" & KS.GetIP & "' and datediff(" & DataPart_H & ",AddDate," & SqlNowString &")<" & KS.ChkClng(Split(KS.C_S(KS.G("ChannelID"),46)&"||||","|")(4))).eof then
		          If Flag="NotAjax" Then KS.Die "<script>alert('对不起，同一篇文档" &KS.ChkClng(Split(KS.C_S(KS.G("ChannelID"),46)&"||||","|")(4)) & "小时内只能评论一次!');history.back();</script>" Else KS.Die("对不起，同一篇文档" &KS.ChkClng(Split(KS.C_S(KS.G("ChannelID"),46)&"||||","|")(4)) & "小时内只能评论一次!")
			  end if
			 End If
			   
			   groupid=ksuser.getuserinfo("groupid")
			 
			   Conn.Execute("Insert Into KS_Comment(ChannelID,InfoID,UserName,Anonymous,Content,QuoteContent,UserIP,Point,Score,OScore,Verific,AddDate,ProjectID) values(" & ChannelID & "," & InfoID & ",'" & AnounName & "'," & Anonymous & ",'" & Replace(C_Content,"'","''") & "','" & Replace(QuoteContent,"'","''") & "','" & KS.GetIP & "',0,0,0," & Verific & "," & SQLNowString& ",0)")
			  If KS.ChkClng(groupid)<>0 and Verific=1 Then
				  If KS.ChkClng(KS.U_S(GroupID,6))>0 Then
					 Call  KS.ScoreInOrOut(KS.C("UserName"),1,KS.ChkClng(KS.U_S(GroupID,6)),"系统","参与文档[<a href=""" & KS.GetItemUrl(channelid,Tid,infoid,Fname) & """ target=""_blank"">" & Title & "</a>]的评论!",1002,""&ChannelID&""&InfoID)
				  End If
			  End If
	  End If
End Sub

Sub QuoteSave()
 Dim quoteId:quoteId=KS.ChkClng(KS.S("quoteId"))
 Dim Content:Content=KS.S("QuoteContent")
 Dim QuoteArray,AnounName,QuoteContent,Verific,Anonymous,UserName,LoginTF
 PostID=KS.ChkClng(KS.S("PostID"))
 If quoteId=0 Then Response.Write "<script>alert('参数传递出错!');</script>":Exit Sub
 If Content="" Then Response.Write "<script>alert('回复内容必须输入!');</script>":Exit Sub
 If Len(Content)>KS.ChkClng(KS.C_S(ChannelID,14)) and KS.ChkClng(KS.C_S(ChannelID,14))<>0 Then
	 KS.Die "<script>alert('评论内容必须在" &KS.C_S(ChannelID,14) & "个字符以内!');</script>"
 End if
 Anonymous=KS.ChkClng(KS.S("Anonymous"))
 LoginTF=Cbool(KSUser.UserLoginChecked)
 IF LoginTF=false and (KS.ChkClng(KS.C_S(Channelid,12))=1 or KS.ChkClng(KS.C_S(Channelid,12))=2) Then
  Response.Write "<script>alert('对不起,本站只允许注册会员发表!');</script>":Exit Sub
 End If
 
 If Anonymous=1 Then
  AnounName="匿名"
 Elseif Anonymous=0 and LoginTF=false then
  Response.Write "<script>alert('对不起,请先登录!');</script>":Exit Sub
 Else
   AnounName=KSUser.UserName
 End If
 If LoginTF=True Then UserName=KSUser.UserName Else UserName="匿名"
 
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 If PostId=0 Then
	  RS.Open "Select top 1 channelid,infoid,username,Anonymous,adddate,content,quotecontent from ks_comment where ProjectID=0 and id=" & quoteid,conn,1,1
	  if RS.Eof Then
		  RS.Close:Set RS=Nothing
		  Response.Write "<script>alert('参数传递出错!');</script>":Exit Sub
	  End If
	  QuoteArray = RS.GetRows(-1)
	  RS.Close : Set RS=Nothing
	 Dim Qstr:Qstr="[dt]引用 " 
	 If QuoteArray(3,0)=1 Then
	  Qstr=Qstr & "匿名"
	 Else
	  Qstr=Qstr & "会员:" & QuoteArray(2,0)
	 End If 
	 Qstr=Qstr & " 发表于" & QuoteArray(4,0) & "的评论内容[/dt][dd]" & QuoteArray(5,0) & "[/dd]"
	 If QuoteArray(6,0)<>"" Then
	 QuoteContent="[quote]" & QuoteArray(6,0) & Qstr & "[/quote]"
	 Else
	 QuoteContent="[quote]" & Qstr & "[/quote]"
	 End If
	 InfoID=QuoteArray(1,0)
 Else
     InfoID=PostId
 End If
 Call DoWriteSave(1,PostID,InfoID,AnounName,Content,QuoteContent,KSUser,Anonymous)
 
 KS.Die "<script>alert('恭喜,您的评论发表成功!');try{parent.loadDate(1);parent.closeWindow();}catch(e){top.location.replace(document.referrer);}</script>"
End Sub

Function Support()
	 Dim ID,OpType,PostId,RS
	 ID=KS.ChkClng(KS.S("ID")) : OpType=KS.ChkClng(KS.S("Type")) : PostId=KS.ChkClng(KS.S("PostID"))
	 IF Cbool(Request.Cookies(Cstr(ID))("SupportCommentID"))<>true Then
	    If PostID<>0 Then
		   Set RS=Conn.Execute("Select top 1 PostTable From KS_GuestBook Where ID=" & PostId)
		   If Not RS.Eof Then
	        if OpType=1 Then
		       Conn.Execute("Update " & RS("PostTable") & " Set Support=Support+1 Where ID=" & ID)
		    else
		       Conn.Execute("Update " & RS("PostTable") & " Set Opposition=Opposition+1 Where ID=" & ID)
			end if
		   End If
		   RS.Close:Set RS=Nothing
		Else
	        if OpType=1 Then
		       Conn.Execute("Update KS_Comment Set score=score+1 Where ID=" & ID)
		    else
	           Conn.Execute("Update KS_Comment Set OScore=OScore+1 Where ID=" & ID)
			end if
		End If
		Response.Cookies(Cstr(ID))("SupportCommentID")=true
	Else
	 Support="你已投过票了！" : Exit Function
	End If
	if OpType=1 Then Support="good" Else Support="bad"
End Function
%>
 
