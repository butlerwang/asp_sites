<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%

Dim KS,KSUser
Set KS=New PublicCls
Set KSUser = New UserCls
Call KSUser.UserLoginChecked()
Dim ChannelID,RS,CommentStr,Total,UserIP

select case KS.S("Action")
  case "CommentSave"
    call CommentSave()
  case else
    Response.Write("document.write('" & GetWriteComment(KS.ChkClng(KS.S("UserID")),KS.S("ID"),KS.S("Title"),KS.S("UserName")) & "');")
end select


		'*********************************************************************************************************
		'函数名：GetWriteComment
		'作  用：取得发表评论信息
		'参  数：ID -信息ID
		'*********************************************************************************************************
		Function GetWriteComment(UserID,ID,Title,UserName)
		%>
		function insertface(Val)
	      {  
		  if (Val!=''){ document.getElementById('Content').focus();
		  var str = document.selection.createRange();
		  str.text = Val; }
          }
		  function success()
			{
				var loading_msg='\n\n\t请稍等，正在提交评论...';
				var content=document.getElementById('Content');
				
				if (loader.readyState==1)
					{
						content.value=loading_msg;
					}
				if (loader.readyState==4)
					{   var s=loader.responseText;
						if (s=='ok')
						 {
						 alert('恭喜,你的评论已成功提交！');
						  location.reload();
						 }
						else
						 {alert(s);
						 }
					}
			}
		

		   function checkform()
		   { 
		    if (document.getElementById('AnounName').value=='')
			{
			 alert('请输入昵称!');
			 document.getElementById('AnounName').focus();
			 return false;
			}
		    if (document.getElementById('Content').value=='')
			{
			 alert('请输入评论内容!');
			 document.getElementById('Content').focus();
			 return false;
			}
		   ksblog.ajaxFormSubmit(document.form1,'success')
           }
		   
		function ShowLogin()
		{ 
		 new KesionPopup().popupIframe('会员登录','<%=KS.Setting(3)%>user/userlogin.asp?Action=Poplogin',397,184,'no');
		}
		<%
		If KS.SSetting(25)="0" And KS.IsNul(KS.C("UserName")) Then
		  GetWriteComment="<div style=""margin:20px""><strong>温馨提示：</strong>只有会员才可以发表评论,如果是会员请先<a href=""javascript:ShowLogin()"">登录</a>,不是会员请点此<a href=""../user/reg/"" target=""_blank"">注册</a>。</div>"
		Else
		 GetWriteComment = "<table width=""98%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""comment_write_table"">"
		 GetWriteComment = GetWriteComment & "<form name=""form1"" action=""WriteComment.asp?action=CommentSave"" method=""post"">"
		 GetWriteComment = GetWriteComment & "<input type=""hidden"" value=""" & UserID & """ name=""UserID""><input type=""hidden"" value=""" & UserName & """ name=""UserName""><input type=""hidden"" value=""" & ID & """ name=""ID"">"
		 GetWriteComment = GetWriteComment & "<tr><td colspan=""2"" height=""30"" class=""comment_write_title""><strong>发表评论:</strong>"
		 Dim HomePage
		 If KS.C("UserName")<>"" Then
		  HomePage=KS.Setting(2) & "/space/?" & KS.C("UserID")
		 Else
		  HomePage="http://"
		 End If
		GetWriteComment = GetWriteComment & "<br/>昵称："
		GetWriteComment = GetWriteComment & "   <input name=""AnounName"" maxlength=""100"" type=""text"" id=""AnounName"" value=""" & KS.C("username") & """"
		If KS.C("UserName")<>"" Then GetWriteComment = GetWriteComment & " readonly"
		GetWriteComment = GetWriteComment & " style=""height:25px;line-height:25px;color:#999;width:35%;border:1px solid #ccc;background:#FBFBFB;""/><br/>主页："
		GetWriteComment = GetWriteComment & "    <input name=""HomePage"" maxlength=""150"" value=""" & HomePage & """ type=""text"" id=""HomePage"" style=""height:25px;line-height:25px;color:#999;width:55%;border:1px solid #ccc;background:#FBFBFB;"" /><br/>标题："
		GetWriteComment = GetWriteComment & "    <input name=""Title"" maxlength=""150"" value=""Re:" & Title & """ type=""text"" id=""Title"" style=""height:25px;line-height:25px;color:#999;width:55%;border:1px solid #ccc;background:#FBFBFB;"" /><input type=""hidden"" value=""" & Title & """ name=""OriTitle""></td>"
		GetWriteComment = GetWriteComment & "  </tr>"
		
		
		GetWriteComment = GetWriteComment & "  <tr>"
		GetWriteComment = GetWriteComment & "    <td height=""25"" width=""70%"" align=""center""><textarea name=""Content"" rows=""6"" id=""Content"" cols=""70"" style=""color:#999;width:98%;border:1px solid #ccc;background:#FBFBFB;overflow:auto""></textarea></td>"
		
		 Dim str:str="惊讶|撇嘴|色|发呆|得意|流泪|害羞|闭嘴|睡|大哭|尴尬|发怒|调皮|呲牙|微笑|难过|酷|非典|抓狂|吐|"
		 Dim strArr:strArr=Split(str,"|")
		  GetWriteComment = GetWriteComment & "<td width=""140"">"
		 For K=0 to 19
		   GetWriteComment = GetWriteComment & "<img style=""cursor:pointer"" title=""" & strarr(k) & """ onclick=""insertface(\'[e" & k &"]\')""  src=""" & KS.Setting(3) & "images/emot/" & K & ".gif"">&nbsp;"
		   If (K+1) mod 5=0 Then GetWriteComment = GetWriteComment & "<br />"
		 Next

		GetWriteComment = GetWriteComment & "</td>"
		GetWriteComment = GetWriteComment & "  </tr>"
		GetWriteComment = GetWriteComment & "  <tr>"
		
		GetWriteComment = GetWriteComment & "    <td colspan=""2"" style=""text-align:left""><input type=""button"" onclick=""return(checkform())"" name=""SubmitComment"" id=""SubmitComment""class=""btn"" value=""提交评论""/>"
		
		GetWriteComment = GetWriteComment & "    </td>"
		GetWriteComment = GetWriteComment & "  </tr>"
		GetWriteComment = GetWriteComment & "  </form>"
		GetWriteComment = GetWriteComment & "</table>"
		End If
		End Function  
  
        Sub CommentSave()
	    	Dim ID,UserName,HomePage,Content,Anonymous,Title,SinaWeiboID
			ID=KS.ChkClng(KS.S("ID"))
			AnounName=KS.LoseHTML(KS.S("AnounName"))
			HomePage=KS.CheckXSS(KS.S("HomePage"))
			Content=KS.CheckXSS(KS.LoseHtml(KS.S("Content")))
			Title=KS.CheckXSS(KS.LoseHtml(KS.S("Title")))
			If Title="" Then Title="回复本文主题"
			IF ID="0" Then 
			 Response.Write("参数传递有误!")
			 Response.End
			End if
			if AnounName="" Then 
			 Response.Write("请填写你的昵称!'")
			 Response.End
			End if
			
			
			if Content="" Then 
			 Response.Write("请填写评论内容!")
			 Response.End
			End if
			
			Set RS=Conn.Execute("Select top 1 UserName,SinaWeiboID From KS_BlogInfo Where ID=" & ID)
			If RS.Eof And RS.Bof Then
			  RS.Close:Set RS=Nothing
			 Response.Write("参数传递有误!")
			 Response.End
			End If
			UserName=RS(0)
			SinaWeiboID=RS(1)
			RS.Close
			
			If Not KS.IsNul(SinaWeiboID) Then   '同步到新浪微博
			  if KSUser.UserLoginChecked=true then 
			    if KSUser.GetUserInfo("sinatoken")<>"" and KSUser.GetUserInfo("sinaid")<>"" Then '判断有绑定就同步
			      call ksuser.add_sina_comment(SinaWeiboID,content&",来自："&KS.GetDomain &"space/?" & KSUser.GetUserInfo("userid")&"/log/" & ID)
			    end if
			  end if
			End If
			
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_BlogComment",Conn,1,3
			RS.AddNew
			 RS("LogID")=ID
			 RS("AnounName")=AnounName
			 RS("Title")=Title
			 RS("UserName")=UserName
			 RS("HomePage")=HomePage
			 RS("Content")=Content
			 RS("UserIP")=KS.GetIP
			 RS("AddDate")=Now
			RS.UpDate
			 RS.Close:Set RS=Nothing
			 Conn.Execute("Update KS_BlogInfo Set TotalPut=TotalPut+1 Where ID=" & ID)
			 

			 
			  Call CloseConn()
             If KS.S("From")="1" Then
			  Response.Write "<script>alert('你的评论发表成功!');top.location.href='" & Request.ServerVariables("HTTP_REFERER") & "';</script>"
			 Else
			   response.write "ok"
			 End If
			 Set KS=Nothing
		End Sub
  
Call CloseConn
Set KS=Nothing
Set KSUser=Nothing
%>
