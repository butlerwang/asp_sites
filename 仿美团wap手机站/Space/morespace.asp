<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%


Dim KSCls
Set KSCls = New Spacemore
KSCls.Kesion()
Set KSCls = Nothing

Class Spacemore
        Private KS, KSR,CurrPage,MaxPerPage,TotalPut,str
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		      Dim FileContent
				   FileContent = KSR.LoadTemplate(KS.SSetting(8))
				   FCls.RefreshType = "MoreSpace" '设置刷新类型，以便取得当前位置导航等
				   Application(KS.SiteSN & "RefreshFolderID") = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
				   If Trim(FileContent) = "" Then FileContent = "空间副模板不存在!"
				   FileContent=Replace(FileContent,"{$ShowMain}",GetSpaceList())
				   FileContent=KSR.KSLabelReplaceAll(FileContent)
		           Response.Write FileContent  
		End Sub
	%>
	<!--#Include file="../ks_cls/ubbfunction.asp"-->
	<%	
		
 '空间列表
 Function GetSpaceList()
		 MaxPerPage =KS.ChkClng(KS.SSetting(9))
		 dim classid:classid=ks.chkclng(ks.s("classid"))
		 dim recommend:recommend=ks.chkclng(ks.s("recommend"))
		 CurrPage = KS.ChkClng(KS.G("page"))
		 If CurrPage<=0 Then CurrPage = 1
		 
	    dim rsc:set rsc=conn.execute("select classname,classid from ks_blogclass order by orderid")
	   if not rsc.eof then
	   str="<div class=""categorybox"">" & vbcrlf
	   str=str &"<ul><li>分类查看：</li>"
		   If classid=0 then 
		     str=str &"<li class=""curr""><a href='morespace.asp'>所有分类</a></li>"
		   else
		     str=str &"<li><a href='morespace.asp'>所有分类</a></li>"
		   end if
	    do while not rsc.eof
		 if classid=rsc(1) then
		   str=str & "<li class=""curr""><a href='?classid=" & rsc(1) &"'>" & rsc(0) & "</a></li>"
		 else
		   str=str & "<li><a href='?classid=" & rsc(1) &"'>" & rsc(0) & "</a></li>"
		 end if
		 rsc.movenext
		loop
	   end if
	   rsc.close:set rsc=nothing
   str=str &"</ul>" & vbcrlf
   str=str &"</div>" &vbcrlf	 
		 
  str=str & "<div class=""morespace"">"
 
 dim param:param=" where status=1"
 if classid<>0 then param=param & " and a.classid=" & classid
 if recommend<>0 then param=param & " and recommend=1"

 if ks.s("key")<>"" then param=param & " and blogname like '%" & ks.r(ks.s("key")) &"%'"
    Dim SQLStr:SQLStr="select a.*,b.classname,u.userface,u.realname from (ks_blog a inner join ks_blogclass b on a.classid=b.classid) inner join ks_user u on a.username=u.username " & param & " order by a.hits desc,a.blogid desc"
	Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		rsobj.open SQLStr ,conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				 	str=str & "<ul><p>对不起，没有找到空间! </p></ul>"
				 Else
							  totalPut = conn.execute("select count(1) from (ks_blog a inner join ks_blogclass b on a.classid=b.classid) inner join ks_user u on a.username=u.username " & param)(0)
								If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrPage - 1) * MaxPerPage
								Else
										CurrPage = 1
								End If
								call ShowSpaceList(RSObj)
				           End If
		 
		 str=str &  "            </div>" & vbcrlf
		 str=str & KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
		 RSObj.Close:Set RSObj=Nothing
		 
		 str=str & "<div class=""clear""></div><table class=""spacesear"">"
		  str=str & "<form name=""myform"" action=""morespace.asp"" method=""get""/> <tr height=""22"">"
	   str=str & "<td align=""center"" colspan=2><strong>按空间名称搜索：</strong><input style=""border:1px #000 solid;height:18px;"" type=""text"" size=""12"" name=""key"">&nbsp;&nbsp;<input type=""submit"" value= "" 查 找 "" class=""btn""></td>"
	   str=str & "</form></tr>"
	   str=str & "</table><br/><br/>"
		 GetSpaceList=str
  End Function

     Function GetCurrLogUrl(UserID,ID)
	  If KS.SSetting(21)="1" Then
	  GetCurrLogUrl=KS.GetDomain &"space/list-" & userid & "-" & id&KS.SSetting(22)
	  Else
	  GetCurrLogUrl="../space/?" & userid & "/log/" & id
	  End If
	 End Function

  Sub ShowSpaceList(rs)
   dim i,logo,rss
   do while not rs.eof
     logo=RS("Logo")
	 If KS.IsNul(Logo) Then
	   logo=RS("UserFace")
	 End If
	 If KS.IsNul(Logo) Then Logo="images/face/boy.jpg"
	 If Left(logo,1)<>"/" and Left(lcase(logo),4)<>"http" Then Logo=KS.Setting(3) & Logo
	 str=str & "<ul>"
	 str=str & "<li>"
      str=str & "<div class=""userpic""><img title=""创建时间：" & rs("adddate") & """ src=""" & Logo & """ /></div><div class=""mysplittd"">"
		  dim spacedomain,predomain
		  If KS.SSetting(14)="1"  Then
		   predomain=rs("domain")
		  end if
		  if predomain<>"" then
		   spacedomain="http://" & predomain & "." & KS.SSetting(16)
		  else
		    spacedomain=KS.GetSpaceUrl(rs("userid"))
		  end if

      str=str & "<span><a title=""" & rs("blogname") & """ href=""" & spacedomain  &""" target=""blank""> " & rs("blogname")  &"</a></span>"
	  if rs("recommend")=1 then str=str & "<font color=red>[荐]</font>"
	  str=str &"<br/> 分类：" & rs("classname")
	  str=str & "<div class=""intro""> " & rs("Descript") & "</div>"
	  
	  set rss=conn.execute("select top 1 * From KS_BlogInfo Where UserName='" & RS("UserName") & "' and status=0")
	  If Not RSS.Eof Then
	   str=str &"<div class=""fresh"">" & KS.Gottopic(KS.LoseHtml(KSR.ReplaceEmot(Ubbcode(rss("content"),0))),200)
	   If RSS("Istalk")="1" Then
	    str=str& "<a href='" & GetCurrLogUrl(RS("UserID"),rss("id")) & "' target='_blank'>[博文]</a>"
	   Else
	    str=str& "<a href='" & GetCurrLogUrl(RS("UserID"),rss("id")) & "' target='_blank'>[新鲜事]</a>"
	   End If
	    str=str &"- <a href='" & GetCurrLogUrl(RS("UserID"),rss("id")) & "' target='_blank'>评论(<font color=red>" & rss("totalput") & "</font>)</a></div>"
	  End If
	  RSS.Close
	  str=str & "<div class=""btntips""><a href='javascript:void(0)' onclick=""addF(event,'" & rs("username") & "')"">加为好友</a> | <a href='javascript:void(0)' onclick=""sendMsg(event,'"& rs("username") & "')"">发送消息</a> | <a href='" & SpaceDomain & "' target='_blank'>关注“" & rs("UserName") & "” 的空间</a></div>"
	  str=str & "</div>"
	  str=str & "</li>"
	  str=str & "</ul>"
   rs.movenext
	  	I = I + 1
		  If I >= MaxPerPage Then Exit Do
	  loop
  End Sub		
		
		
End Class
%>
