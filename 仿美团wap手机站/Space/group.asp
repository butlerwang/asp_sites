<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<%

Dim KSCls
Set KSCls = New Group
KSCls.Kesion()
Set KSCls = Nothing

Class Group
        Private KS,KSBCls,KSUser,KSR
		Private PerPageNumber,CurrPage,totalPut,RS,MaxPerPage
		Private ID,Template,TemplateID,TeamName,groupadmin
		Private Sub Class_Initialize()
		  MaxPerPage =15
		  Set KS=New PublicCls
		  Set KSR=New refresh
		  Set KSBCls=New BlogCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSBCls=Nothing
		 Set KSUser=Nothing
		 Set KSR=Nothing
		End Sub
		Public Sub Kesion()
		 If KS.SSetting(0)=0 Then
		 Response.Write "<script>alert('对不起，本站点关闭个人空间功能!');window.close();</script>"
		 Response.end
		 End If
		ID=KS.ChkClng(KS.S("ID"))
		If ID=0 Then Response.End()
		Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select top 1 * From KS_Team Where ID=" & ID,conn,1,1
		If RS.Eof And RS.Bof Then
		 Response.Write "<script>alert('参数传递出错!');window.close();</script>"
		 Response.end
		End If
		If RS("Verific")=0 Then
		 Response.Write "<script>alert('该圈子尚未审核!');window.close();</script>"
		 response.end
		elseif RS("Verific")=2 then
		 Response.Write "<script>alert('该圈子已被管理员锁定!');window.close();</script>"
		 response.end
		end if
		
		 TeamName=RS("TeamName")
		 groupadmin=rs("username")
		 TemplateID=RS("TemplateID")
		 Template="<html>"&vbcrlf &"<title>" & TeamName & "</title>" &vbcrlf
		 Template=Template & "<meta http-equiv=""Content-Language"" content=""zh-CN"" />" &vbcrlf
         Template=Template & "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />" & vbcrlf
         Template=Template & "<meta name=""generator"" content=""KesionCMS"" />" & vbcrlf
		 Template=Template & "<meta name=""author"" content=""" & RS("UserName") & ","" />" & vbcrlf
		 Template=Template & "<meta name=""keyword"" content=""" & TeamName & """ />"&VBCRLF		 Template=Template & "<meta name=""description"" content=""KS"" />"  & vbcrlf
		 Template=Template & "<link href=""css/css.css"" type=""text/css"" rel=""stylesheet"">" & vbcrlf
		 Template=Template & "<script src=""../ks_inc/jquery.js"" language=""javascript""></script>"  & vbcrlf
		 Template=Template & "<script src=""../ks_inc/kesion.box.js"" language=""javascript""></script>"  & vbcrlf
		 Template=Template & "<script src=""js/ks.space.js"" language=""javascript""></script>"  & vbcrlf
		 Template=Template & "<script src=""js/ks.space.page.js"" language=""javascript""></script>"  & vbcrlf
		 'template=Template & KSBCls.GetTemplatePath(TemplateID,"TemplateMain")
		 template=KSBCls.GetTemplatePath(TemplateID,"TemplateMain")
		 
		 
		 
		 template=KSBCls.ReplaceGroupLabel(RS,Template)
		 
		 Select Case KS.S("Action")
		  case "showtopic"
		   	Template=Replace(Template,"{$GroupMain}","<script language=""javascript"" defer>TeamPage(1,'showtopic&teamid=" & id & "&tid=" & KS.S("tid") & "&groupadmin=" & groupadmin & "')</script><div id=""teammain""></div><div id=""kspage"" align=""right""></div>" &  showtopic)
		  case "replaysave"
		   call replaysave
		  case "users"
		   	Template=Replace(Template,"{$GroupMain}","<script language=""javascript"" defer>TeamPage(1,'users&teamid=" & id & "')</script><div id=""teammain""></div><div id=""kspage"" style=""clear:both"" align=""right""></div>")
		 case "join"
		  		Template=Replace(Template,"{$GroupMain}",showjoin())
		 case "joinsave"
		    call joinsave
		 case "deltopic"
		    call deltopic
		 case "deluser"
		    call deluser()
		 case "settop"
		   call settop()
		 case "setbest"
		   call setbest()
		 case "post"
		  	Template=Replace(Template,"{$GroupMain}",showpost())
		 case "topicsave"
		   call topicsave()
		 case "info"
		  	Template=Replace(Template,"{$GroupMain}",showinfo())
		 case else
		 Template=Replace(Template,"{$GroupMain}","<script language=""javascript"" defer>TeamPage(1,'teamtopic&teamid=" & id & "&isbest=" & KS.ChkClng(KS.S("isbest")) &"')</script><div id=""teammain""></div><div id=""kspage"" align=""right""></div>")
		  end select
		  
		 Template=KSR.KSLabelReplaceAll(Template)
		 Response.Write Template
		  RS.Close
          Set  RS=Nothing
		End Sub
		
		function showtopic()
		 dim tid:tid=KS.chkclng(KS.S("tid"))
		showtopic=showtopic &"<div id=""form_comment""><a name=""add_comment""></a>"
		showtopic=showtopic &"<script type=""text/javascript"">function checkform(){if (CKEDITOR.instances.Content.getData()==''){alert('请输入回复内容!');CKEDITOR.instances.Content.focus();return false;}return true;}</script>"
		showtopic=showtopic &"<br/>"
		showtopic=showtopic &"<table width=""100%"" cellpadding=""0"" cellspacing=""0"">"
		showtopic=showtopic &"<form action='group.asp?action=replaysave&id=" & id & "&tid=" & tid & "' method='post' name='myform' id='myform' onSubmit=""return(checkform())"">"
		showtopic=showtopic &"    <tr>"
		showtopic=showtopic &"	  <td colspan=""2"" bgcolor=""#EDF5F9"" style=""padding-left:15px;height:25px;line-height:25px;""><strong>回复话题</strong></td>"
		showtopic=showtopic &"	</tr>"
			IF Cbool(KSUser.UserLoginChecked)=false Then
		showtopic=showtopic &"    <tr>"
		showtopic=showtopic &"	  <td colspan=""2"" bgcolor=""#FFFFFF"" align=""center"" height=""80""><p>登录后才可以参与该话题的讨论,如要参与讨论请先<a href=""../user/login/"" target=""_blank"">登录</a>到会员中心！</p></td>"
		showtopic=showtopic &"	</tr>"
			else
			on error resume next
		showtopic=showtopic &"	<tr>"
		showtopic=showtopic &"		<td width=""100"" align=""center"" bgcolor=""#FFFFFF"">回复话题：</td>"
		showtopic=showtopic &"		<td bgcolor=""#FFFFFF""><input type=""text"" readonly value=""Re:" & conn.execute("select title from ks_teamtopic where id="& tid )(0) & """ size=""50"" name=""title"">"
        showtopic=showtopic &"        </td>"
		showtopic=showtopic &"	</tr>"
		showtopic=showtopic &"	<tr>"
		showtopic=showtopic &"		<td colspan='2' bgcolor=""#FFFFFF"">"
		showtopic=showtopic &"		<textarea id=""Content"" name=""Content"" style=""display:none""></textarea>"
		
		showtopic=showtopic & "<script type=""text/javascript"" src=""../editor/ckeditor.js"" mce_src=""../editor/ckeditor.js""></script><script type=""text/javascript"">CKEDITOR.replace('Content', {width:""97%"",height:""160px"",toolbar:""Basic"",filebrowserBrowseUrl :""../editor/ksplus/SelectUpFiles.asp"",filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script>"

        showtopic=showtopic &"        </td>"
		showtopic=showtopic &"	</tr>"
		showtopic=showtopic &"	<tr>"
		showtopic=showtopic &"		<td colspan=""2"" bgcolor=""#FFFFFF"">"
		showtopic=showtopic &"		<input type=""submit"" class='btn' style='padding:2px;margin:2px' value="" 提 交 回 复 "">"
        showtopic=showtopic &"        </td>"
		showtopic=showtopic &"	</tr>"
			end if
		showtopic=showtopic &"	</form>"
	    showtopic=showtopic &"</table>"
		showtopic=showtopic &"</div>"
		end function
		

		'保存回复
		Sub replaysave()
		dim tid:tid=KS.chkclng(KS.S("tid"))
		dim title:title=KS.S("title")
		dim content:content=KS.S("content"):if content="" then call KS.alert("请输入回复内容!",""):exit sub
		IF Cbool(KSUser.UserLoginChecked)=false Then  call KS.alert("请先登录!",""):exit sub
		dim username:username=KS.R(KSUser.UserName)
		dim rs:set rs=server.createobject("adodb.recordset")
		rs.open "select top 1 * from ks_teamtopic",conn,1,3
		rs.addnew
		 rs("parentid")=tid
		 rs("teamid")=id
		 rs("title")=title
		 rs("content")=content
		 rs("adddate")=now
		 rs("userip")=KS.getip
		 rs("status")=1
		 rs("username")=username
		 rs("isbest")=0
		  rs("istop")=0
		rs.update
		rs.movelast
		Call KS.FileAssociation(1031,rs("ID"),content,0)
		rs.close:set rs=nothing
		response.redirect request.servervariables("http_referer")
		End Sub
		
		Function showjoin()
		 IF Cbool(KSUser.UserLoginChecked)=false Then
		  showjoin= "对不起，申请加入圈子之前必须先<a href=""../user/login/"" target=""_blank"">登录</a>到会员中心！"
		  exit function
		 end if
		 if not conn.execute("select username from ks_teamusers where username='" & ksuser.username & "' and teamid=" & id).eof then
		  showjoin= "<div><b>您不能再申请，产生的可能原因如下：</b><div style='border:1px solid #efefef;overflow:hidden'></div><li>您已申请过，未得到圈主的审核;</li><li>您已是本圈子的成员，不需要再申请;</li><li>您可能已被圈主邀请，但您还未在会员中心确认;</li></div>"
		  showjoin=showjoin & "<div><b>申请须知：</b><div style='border:1px solid #efefef;overflow:hidden'></div>"
		  showjoin=showjoin & RS("Note")
		  showjoin=showjoin & "</div>"
		  exit function
		 end if
		  showjoin=showjoin & "<script>"
		  showjoin=showjoin & " function checkform()"
		  showjoin=showjoin & " {if (document.myform.username.value==''){"
		  showjoin=showjoin & "	 alert('申请人必须填写!');"
		  showjoin=showjoin & "	 document.myform.username.focus();"
		  showjoin=showjoin & "	 return false"
		  showjoin=showjoin & "	 }"
		  showjoin=showjoin & "	 if (document.myform.reason.value==''){"
		  showjoin=showjoin & "	  alert('请输入加入圈子的理由!');"
		  showjoin=showjoin & "	  document.myform.reason.focus();"
		  showjoin=showjoin & "	  return false"
		  showjoin=showjoin & "	  }"
		  showjoin=showjoin & "	  return true;"
		  showjoin=showjoin & " }"
		  showjoin=showjoin & "</script>"
		  showjoin=showjoin & "<table width=""100%"" cellspacing=""0"" cellspadding=""0"" border=""0"">"
		  showjoin=showjoin & " <form name=""myform"" action=""?id=" & id & "&action=joinsave"" method=""post"" onSubmit=""return(checkform())""> "
		  showjoin=showjoin & "	<tr>"
		  showjoin=showjoin & "	  <td align=""center"" bgcolor=""#f9f9f9"" colspan=2>申 请 加 入 群 组</td>"
		  showjoin=showjoin & "	</tr>"
		  showjoin=showjoin & "	<tr>"
		  showjoin=showjoin & "	  <td width=""100"">申 请 人：</td>"
		  showjoin=showjoin & "	  <td><input name=""username"" type=""textbox"" value=""" & ksuser.username & """ readonly size=10></td>"
		  showjoin=showjoin & "	</tr>"
		  showjoin=showjoin & "	<tr>"
		  showjoin=showjoin & "	  <td>加入理由：</td>"
		  showjoin=showjoin & "	  <td><textarea name=""reason"" cols=""50"" rows=""6""></textarea></td>"
		  showjoin=showjoin & "	</tr>"
		  showjoin=showjoin & "	<tr>"
		  showjoin=showjoin & "	  <td colspan=2 align=""center""><input type=""submit"" value=""提交申请""></td>"
		  showjoin=showjoin & "	</tr>"
		  showjoin=showjoin & "	</form>"
		  showjoin=showjoin & "</table>"
		  showjoin=showjoin & "<div><b>申请须知：</b><div style='border:1px solid #efefef;overflow:hidden'></div>"
		  showjoin=showjoin & RS("Note")
		  showjoin=showjoin & "</div>"
		End Function
		
		'保存申请
		Sub JoinSave()
		dim id:id=KS.chkclng(KS.S("id"))
		dim username:username=KS.R(KS.S("username"))
		dim reason:reason=KS.R(KS.S("reason"))
		dim rs:set rs=server.createobject("adodb.recordset")
		rs.open "select * from ks_teamusers where teamid=" & id & " and username='" & username & "'",conn,1,3
		if rs.eof then
		 rs.addnew
		  rs("teamid")=id
		  rs("username")=username
		  rs("status")=2  '申请加入
		  rs("power")=0   '普通用户
		  rs("reason")=reason
		  rs("Applydate")=now
		 rs.update
		end if
		rs.close:set rs=nothing
		call KS.alert("你的申请已提交，请等待圈主的审核!","?id=" & id)
		End Sub
		
		'发表新贴
		function showpost()
		 IF Cbool(KSUser.UserLoginChecked)=false Then
		  showpost= "对不起，发表新帖之前必须先<a href=""../User/"" target=""_blank"">登录</a>到会员中心！"
		  exit function
		 end if
		 if conn.execute("select username from ks_teamusers where username='"& ksuser.username & "' and teamid=" & id).eof then
		  showpost= "对不起，你不是该圈子的成员，没有权利发表话题！"
		  exit function
		 elseif conn.execute("select username from ks_teamusers where username='"& ksuser.username & "' and status<>2 and teamid=" & id).eof then
		  showpost= "对不起，你提交的申请还未得到确认，没有权利发表话题！"
		  exit function
		 end if

		showpost="<script>"
		showpost=showpost & "function checkform()"
		showpost=showpost & " {"
		showpost=showpost & "  if (document.myform.topic.value=='')"
		showpost=showpost & "  {"
		showpost=showpost & "   alert('请输入讨论话题!');"
		showpost=showpost & "   document.myform.topic.focus();"
		showpost=showpost & "  return false;"
		showpost=showpost & "  }"
		showpost=showpost & "  if (CKEDITOR.instances.Content.getData()=='')"
		showpost=showpost & "  {"
		showpost=showpost & "   alert('请输入讨论内容!');"
		showpost=showpost & "   CKEDITOR.instances.Content.getData().focus();"
		showpost=showpost & "   return false;"
		showpost=showpost & "  }"
		showpost=showpost & "  return true;"
		showpost=showpost & " }"
		showpost=showpost & "</script>"
		showpost=showpost & "<div id=""form_comment"">"
		showpost=showpost & "<form action='group.asp?action=topicsave&id=" & id & "' onSubmit=""return(checkform())"" method='post' name='myform' id='myform'>"
		showpost=showpost & "<div id=""ad_teamcomment""></div><ul><p>" & ksuser.username &" , 欢迎您参与圈子讨论!</p></ul><ul>仅该圈子成员可以发起主题，非成员仅可以回复</ul><ul>昵称：<input name='UserName' type='text' id='UserName' size='15' maxlength='20' value='" & ksuser.username & "' readonly /></ul>"
		showpost=showpost & "<ul>话题：<input name='topic' type='text' id='topic' size='50' maxlength='50' value='' /></ul>"
		showpost=showpost & "<ul>"
		showpost=showpost & "<div><textarea id=""Content"" name=""Content"" style=""width:400px;height:250px; display:none"" ></textarea><script type=""text/javascript"" src=""../editor/ckeditor.js"" mce_src=""../editor/ckeditor.js""></script><script type=""text/javascript"">CKEDITOR.replace('Content', {width:""99%"",height:""200px"",toolbar:""Basic"",filebrowserBrowseUrl :""../editor/ksplus/SelectUpFiles.asp"",filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script></div> "
		showpost=showpost & "</ul>"
		showpost=showpost & "<ul><input type='submit' style='margin:3px;padding:2px' class='btn' value=' OK,发 表 '></ul>"
		showpost=showpost & "</form>"
		showpost=showpost & "</div>"
		End Function
		'保存发表
		Sub topicsave()
		 dim id:id=KS.chkclng(KS.S("id"))
		 dim topic:topic=KS.R(KS.S("topic"))
		 dim content:content=KS.HTMLEncode(KS.S("content"))
		 dim rs:set rs=server.createobject("adodb.recordset")
		 rs.open "select top 1 * from ks_teamtopic",conn,1,3
		 rs.addnew
		  rs("title")=topic
		  rs("content")=content
		  rs("teamid")=id
		  rs("parentid")=0
		  rs("username")=KS.S("username")
		  rs("adddate")=now
		  rs("userip")=KS.getip
		  rs("status")=1
		  rs("isbest")=0
		  rs("istop")=0
		 rs.update
		 rs.movelast
		 Call KS.FileAssociation(1031,rs("ID"),content,0)
		 rs.close:set rs=nothing
		 response.write "<script>alert('您的讨论话题发表成功！');location.href='?id=" & id &"';</script>"
		End Sub	
		
		
		'圈子信息
		function showinfo()
		showinfo="<div id=""ginfo"">"
		showinfo=showinfo &"	<h1>圈子信息</h1>"
		showinfo=showinfo &"<div id=""group_info"">"
		showinfo=showinfo &"	<div id=""group_pic""><img src=""" & rs("photourl") & """ border=""0""></div>"
		showinfo=showinfo &"	<div id=""group_xx"">"
		showinfo=showinfo &"	<li>圈子名称:" & rs("teamname") & "</li>"
		showinfo=showinfo &"	<li>创建者:" & rs("username") & "</li>"
		showinfo=showinfo &"	<li>创建时间:" & rs("adddate") & "</li>"
		showinfo=showinfo &"	<li>成员人数:" & conn.execute("select count(username)  from ks_teamusers where status=3 and teamid=" & rs("id"))(0) & "</li>"
		showinfo=showinfo &"	<li>主题回复:" & conn.execute("select count(*) from ks_teamtopic where parentid=0 and teamid=" & id )(0) & "/" & conn.execute("select count(*) from ks_teamtopic where parentid<>0 and teamid=" & id )(0) & "</li>"
		showinfo=showinfo &"</div></div>"
		showinfo=showinfo &"<div id=""user_list"">"
		showinfo=showinfo &"  <h1>圈子管理员</h1>"
		showinfo=showinfo &"<div>"
		showinfo=showinfo &"  <ul><li class=""u1"">"
			dim rsu:set rsu=server.createobject("adodb.recordset")
			rsu.open "select top 1 * from ks_user where username='" & rs("username") &"'",conn,1,1
			if not rsu.eof then
			  Dim UserFaceSrc:UserFaceSrc=rsu("UserFace")
			  if lcase(left(userfacesrc,4))<>"http" then userfacesrc=KS.Setting(2) & userfacesrc

		showinfo=showinfo &" <img src=""" & UserFaceSrc & """ border=""1"" width=""48"" height=""48""></li>"
		showinfo=showinfo &"	<li class=""u2""><a href=""?" & rsu("username") & """ target=""_blank"">" & rs("username") & "</a></li>"
		showinfo=showinfo &"	<li class=""u3"">(" & rsu("province") & rsu("city") & ")</li>"
			end if
			rsu.close:set rsu=nothing
		showinfo=showinfo &"</ul>"
		showinfo=showinfo &"</div></div></div>"
		End Function
        
		Sub deltopic()
		 IF Cbool(KSUser.UserLoginChecked)=false Then
		  call KS.Alert("对不起，请先登录！","")
		  exit sub
		 end if
		 dim tid:tid=ks.chkclng(ks.s("tid"))
		 if tid=0 then response.end
		 dim rst:set rst=server.createobject("adodb.recordset")
		 rst.open "select * from ks_teamtopic where id=" & tid,conn,1,3
		 if not rst.eof then
		     conn.execute("delete from ks_uploadfiles where channelid=1031 and infoid=" & tid)
		  if rst("username")=KSUser.UserName or KSUser.UserName=groupadmin then
		     conn.execute("delete from ks_uploadfiles where channelid=1031 and infoid in(select id from ks_teamtopic where parentid=" & tid & ")")
		     conn.execute("delete from ks_teamtopic where parentid=" & tid)
			 rst.delete
		  else
		     rst.close:et rst=nothing
		    call ks.alert("对不起，你没有删除的权限","")
		  end if
		 end if
		 rst.close:set rst=nothing
		 if ks.s("flag")="replay" then
		 response.write "<script>alert('删除成功');location.href='"& request.servervariables("http_referer") & "';</script>"
		 else
		 response.write "<script>alert('删除成功');location.href='group.asp?id="& id & "';</script>"
		 end if
		End Sub
		'置顶设置
		Sub Settop()
		 IF Cbool(KSUser.UserLoginChecked)=false Then
		  call KS.Alert("对不起，请先登录！","")
		  exit sub
		 end if
		  dim tid:tid=KS.chkclng(KS.S("tid"))
		  dim rs:set rs=server.createobject("adodb.recordset")
		  rs.open "select top 1 istop from ks_teamtopic where id=" & tid,conn,1,3
		  if not rs.eof then
		   if rs(0)=1 then
			 rs(0)=0
		   else
			 rs(0)=1
		   end if
		   rs.update
		  end if
		  rs.close:set rs=nothing
		  response.redirect request.servervariables("http_referer")
		end sub
		'精华设置
		Sub Setbest()
		 IF Cbool(KSUser.UserLoginChecked)=false Then
		  call KS.Alert("对不起，请先登录！","")
		  exit sub
		 end if
		  dim tid:tid=KS.chkclng(KS.S("tid"))
		  dim rs:set rs=server.createobject("adodb.recordset")
		  rs.open "select top 1 isbest from ks_teamtopic where id=" & tid,conn,1,3
		  if not rs.eof then
		   if rs(0)=1 then
			 rs(0)=0
		   else
			 rs(0)=1
		   end if
		   rs.update
		  end if
		  rs.close:set rs=nothing
		  response.redirect request.servervariables("http_referer")
		end sub
		Sub deluser()
		 IF Cbool(KSUser.UserLoginChecked)=false Then
		  call KS.Alert("对不起，请先登录！","")
		  exit sub
		 end if
		  if KSUser.UserName=groupadmin then
		     conn.execute("delete from ks_teamusers where teamid=" &id & " and username<>'" & ksuser.username & "' and username='" & KS.S("UserName") & "'")
		  else
		    call ks.alert("对不起，你没有此操作的权限","")
		  end if
		 response.write "<script>alert('用户已被成功踢出!');location.href='group.asp?id="& id & "&action=users';</script>"
		End Sub
End Class
%>