<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New admin_Push
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Push
        Dim KS,KSCls,ChannelID,Action,Bsetting
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		Sub Kesion
		    ChannelID=KS.ChkClng(KS.S("ChannelID"))
			Action=KS.G("Action")
			Select Case Action
			 Case "pushToClub" pushToClub
			 Case "doPush" doPush
			End Select
        End Sub
		%>
		<!--#include file="../ks_cls/ClubFunction.asp"-->
		<%
	 Sub pushToClub()
	 %>
	 <html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	 <link href="include/admin_Style.CSS" rel="stylesheet" type="text/css">
     <script language="JavaScript" src="../KS_Inc/JQuery.js"></script>
	 </head>
	 <body>
	 <form name="myform" id="myform" action="KS.Push.asp" method="post">
	 <input name="action" type="hidden" value="doPush">
	 <input name="itemid" type="hidden" value="<%=KS.G("IDS")%>">
	 <input name="channelid" type="hidden" value="<%=channelid%>">
	 <div style="padding:5px;"><strong>选择推送到的论坛版块</strong><span id="navlist1"></span><span id="navlist2"></span><br/><div id="boardlist"><img src="../images/loading.gif" /></div>
	  Tips:收费内容，推送到论坛版块后将失效，建议不要将收费内容推送到论坛版块。
	 </div>
	 </form>
	 <script type="text/javascript">
	  $.get("../plus/ajaxs.asp",{action:"getclubboard",anticache:Math.floor(Math.random()*1000)},function(d){
			setHtml(d);
	   });
	   function loadBoard(v){
		  if (v==''||v=='0') return;
		  var str=$("#pid>option:selected").text();
		   $("#navlist1").html("->"+str);
		   $("#navlist2").html("");
		  $.get("../plus/ajaxs.asp",{action:"getclubboard",pid:v},function(d){
			setHtml(d);
		   });
		}
	  function setHtml(h){
	  	$("#boardlist").html(h);
		$("#btns").html('<div id="showcategory"></div><br/><br/>文档ID号:<%=KS.G("IDS")%><br/><br/><label><input type="checkbox" name="istop" value="1">同时设置为置顶</label><br/><label><input type="checkbox" name="isbest" value="1">同时设置为精华</label><br/><br/><input type="button" value="推送到选定的版块" onclick="return(check())" class="button"/>');
		$("#bid").click(function(){
		  $.get("../plus/ajaxs.asp",{action:"getclubboardcategory",boardid:$(this).val(),anticache:Math.floor(Math.random()*1000)},function(d){
		     $("#showcategory").html(unescape(d));
	       });
		});
	  }
	  function check(){
	   var bid=$('#bid option:selected').val();
		 if (bid!='' && bid!=undefined){
		   $("#myform").submit();
		  }
		 else{
		  alert('请选择要推送到的版面!');
          return false;}
	  }
	 </script>
	 </body>
	 </html>
	 <%
	 End Sub
	 Sub DoPush()
	   Dim ItemID:ItemId=KS.FilterIds(KS.S("ItemID"))
	   If ItemId="" Then KS.AlertHintScript "请选择要推送的文档!"
	   Dim BoardID:BoardID=KS.ChkClng(KS.S("Bid"))
	   Dim CategoryId:CategoryId=KS.ChkClng(KS.S("CategoryId"))
	   If BoardID=0 Then KS.AlertHintScript "请选择论坛版块！"
	    Dim Node,O_LastPost,N,TopicID,boardname,TopicList,Content,i,Arr,PicArr
	   If BoardID<>0 Then
		 KS.LoadClubBoard()
		 Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
		 O_LastPost=Node.SelectSingleNode("@lastpost").text
		 boardname=Node.SelectSingleNode("@boardname").text
		 BSetting=Node.SelectSingleNode("@settings").text&"$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
		 BSetting=Split(BSetting,"$")
	   End If
       Dim PostTable,CommentTable,CommentNum,LastReplayUserID,LastReplayTime,LastReplayUser,RSC
	   Dim IsTop:IsTop=KS.ChkClng(KS.S("IsTop"))
	   Dim IsBest:IsBest=KS.ChkClng(KS.S("IsBest"))
	   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	   RS.Open "Select * From " & KS.C_S(ChannelID,2) &"  Where ID in(" & ItemID &")",conn,1,1
	   N=0:TopicID=0
	   Do While Not RS.Eof
	     dim title:title=rs("title")
		 dim postId:PostId=RS("PostId")
	     If KS.ChkClng(PostId)=0 Then
			CommentTable=RS("PostTable"):If KS.IsNul(CommentTable) Then CommentTable="KS_Comment" '得到评论表
			CommentNum=KS.ChkClng(Conn.Execute("select count(id) From " & CommentTable & " Where ChannelID=" & ChannelID & " And InfoID=" & RS("ID"))(0))
			 content=""
			 Select Case KS.C_S(ChannelID,6)
			  case 1 Content=RS("ArticleContent")
			  case 2 
				arr=split(RS("PicUrls"),"|||")
				for i=0 to ubound(arr)
				  PicArr=split(arr(i),"|")
				  content=content & "[img]" &PicArr(1) & "[/img][br]" & PicArr(0) & "[br]"
				next
				  content=content & RS("PictureContent")
			 end select
			 dim user,userid
			 dim rsu:set rsu=conn.execute("select top 1 PrUserName from ks_admin where username='" & rs("inputer") &"'")
			 if not rsu.eof then
			    user=rsu(0)
			 else
			    user=rs("inputer")
			 end if
			 rsu.close
			 set rsu=conn.execute("select top 1 userid from ks_user where username='" & user & "'")
			 if not rsu.eof then
			   userid=rsu(0)
			 else
			   userid=0
			 end if
			 rsu.close:set rsu=nothing
			 
			 dim infoid:infoid=rs("id")
			 dim inputer:inputer=RS("Inputer")
			 dim hits:hits=rs("hits")
			 dim photourl:photourl=rs("photourl")
			TopicID=InsertPost(BoardID,0,user,userid,Title,Content,photourl,"",0,0,0,0,CategoryId,Hits,IsTop,IsBest,0,O_LastPost,1,PostTable)

			 '转移评论到论坛的回复表
			Set RSC=Conn.Execute("Select a.*,u.UserID From " & CommentTable & " a left join ks_user u on a.username=u.username Where a.ChannelID=" & ChannelID & " And a.InfoID=" & infoid & " Order By a.AddDate")
			If Not RSC.Eof Then
			  Do While Not RSC.Eof 
			   Content=RSC("Content")
			   LastReplayUserID=KS.ChkClng(RSC("UserID"))
			   LastReplayTime=RSC("AddDate")
			   LastReplayUser=RSC("UserName")
			   Call InsertReply(PostTable,LastReplayUser, LastReplayUserID,TopicID,Content,0,1,TopicID,RSC("Verific"),"'" &LastReplayTime &"'")
			  RSC.MoveNext
			  Loop
			  Conn.Execute("Delete From " & CommentTable & " Where ChannelID=" & ChannelID & " And InfoID=" & InfoID)
			End If
			Set RSC=Nothing
			If KS.IsNul(LastReplayTime) Then LastReplayTime=Now
			If KS.IsNul(LastReplayUser) Then LastReplayUser=inputer
			Conn.Execute("Update KS_GuestBook Set LastReplayUserID=" & KS.ChkClng(LastReplayUserID) & ",LastReplayTime='" & LastReplayTime & "',LastReplayUser='"& LastReplayUser & "',Hits=" & hits & ",TotalReplay=" & CommentNum & ",ChannelID=" & ChannelID & ",InfoID=" & InfoID & " Where ID=" & TopicID )
			Conn.Execute("Update " & KS.C_S(ChannelID,2) &" Set PostId=" & TopicID & ",PostTable='" & PostTable & "' Where ID=" & InfoID)
			'如果有启用生成，则重新生成内容页
			If KS.C_S(Channelid,7)<>0 Then
			         Dim RSR:Set RSR=Conn.Execute("select top 1 * From " & KS.C_S(ChannelID,2) &" Where ID=" & InfoID)
					 If Not RSR.Eof Then
						 Dim KSRObj:Set KSRObj=New Refresh
						 Dim DocXML:Set DocXML=KS.RsToXml(RSR,"row","root")
						 Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
						  KSRObj.ModelID=ChannelID
						  KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 
						  Call KSRObj.RefreshContent()
						  Set KSRobj=Nothing
					End if
					Set RSR=Nothing
			End If
			
	   Else
	     TopicID=PostId
	   End If 
		N=N+1
		TopicList=TopicList & n &"、" & Title & "  <a href='" & KS.GetClubShowUrl(TopicID) & "' target='_blank'>浏览帖子</a><br/>"
	   RS.MoveNext
	   Loop
	   RS.Close
	   Set RS=Nothing
	   %>
	   <html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">
		 <link href="include/admin_Style.CSS" rel="stylesheet" type="text/css">
		 <script language="JavaScript" src="../KS_Inc/JQuery.js"></script>
		 </head>
		 <body>
	   <%
	   If IsTop=1 Then MustReLoadTopTopic
	   If N>0 Then
	     KS.Echo "<br/><div class=""attention""><span style=""font-weight:bold"">恭喜，您已成功的推送 <span style='color:red'>" & N & "</span> 篇文档到论坛『" & BoardName &" 』版块!</span><br/>"
		  KS.Echo "<br/><input type='button' value='浏览版块' class=""button"" onclick=""window.open('" &KS.GetClubListUrl(BoardID) &"');""/>"

		 
		  KS.Echo "<br/><br/><strong>您还可以浏览以下推送的文档:</strong><br/>" 
		  KS.echo ("<div style=""*height:160px;overflow-x: hidden; overflow-y: auto;"">")
		  KS.Echo TopicList
		  KS.Echo "</div>"
		 
		 KS.Echo "</div>"
	   End If
	  %>
	   </body>
	   </html>
	  <%
	 End Sub
	  
	
		
End Class
%>