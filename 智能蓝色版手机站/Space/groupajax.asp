<%@ Language="VBSCRIPT" CODEPAGE="65001" %>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<%

response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"
Dim KSCls
Set KSCls = New AjaxCls
KSCls.Kesion()
Set KSCls = Nothing

Class AjaxCls
      Private KS,KSUser
	  Private Action,Template,id,groupadmin
	  Private CurrentPage,totalPut,MaxPerPage,PageNum
	  Private Sub Class_Initialize()
	   Set KS=New PublicCls
       Set KSUser=New UserCls
      End Sub
	 Private Sub Class_Terminate()
	  Set KS=Nothing
	  Set KSUser=Nothing
	  CloseConn()
	 End Sub

     Sub Kesion()
      Action=KS.S("Action")
      id=ks.chkclng(ks.s("teamid"))
	  if id=0 then response.End()
		 groupadmin=conn.execute("select top 1 username from ks_team where id=" & id)(0)
	   Select Case Action
		Case "teamtopic"
		 Call TeamTopic()
		Case "showtopic"
		 Call ShowTopic()
		Case "users"
		 Call ShowUsers()
	   End Select	
	 End Sub	
	 

  
  '圈子主题列表
  Sub TeamTopic()
		 MaxPerPage =20
		 response.write "  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		  If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
%>
  <table border="0" cellpadding="1" cellspacing="1" width="98%" backcolor="#efefef">
      <tr height="25" bgcolor="#f9f9f9">
      <td align="center"><strong>话题</strong></td>
      <td width="80" align="center"><strong>作者</strong></td>
      <td width="50" align="center"><strong>回复</strong></td>
      <td width="100" align="center"><strong>最后更新</strong></td>
    </tr>

<%
 dim param:param=" where teamid=" & id & " and parentid=0"
 if KS.chkclng(KS.S("isbest"))=1 then param=param & " and isbest=1 "
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			rsobj.open "select * from ks_teamtopic "& param & " order by istop desc,adddate desc" ,conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				 	response.write "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=4><p>没有任何讨论话题! </p></td></tr>"
				 Else
							totalPut = RSObj.RecordCount
                           If CurrentPage < 1 Then	CurrentPage = 1
			
									If (totalPut Mod MaxPerPage) = 0 Then
										pagenum = totalPut \ MaxPerPage
									Else
										pagenum = totalPut \ MaxPerPage + 1
									End If
								If CurrentPage = 1 Then
									call ShowTeamTopicList(RSObj)
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
										call ShowTeamTopicList(RSObj)
									Else
										CurrentPage = 1
										call ShowTeamTopicList(RSObj)
									End If
								End If
				           End If
		 
		 response.write  "            </table>" & vbcrlf
		 Response.Write "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|个||2"
		 RSObj.Close:Set RSObj=Nothing
  End Sub

  Sub ShowTeamTopicList(rs)
   dim i
   do while not rs.eof
    if i mod 2=0 then
	 response.write "<tr style=""background:#fff;height:25px"">"
	else
	 response.write "<tr style=""background:#FBFDFF;height:25px"">"
	end if
	%>
      <td>
	  <%if rs("istop")=1 then response.write "[置顶]"
	     if rs("isbest")=1 then response.write "[精华]"%>
      <a title="<%=rs("title")%>" href="?action=showtopic&id=<%=id%>&tid=<%=rs("id")%>">
	  <%=rs("title")%></a></td>
      <td align="center"><a title="作者:<%=rs("username")%>" target="_blank" href="?<%=rs("username")%>"><%=rs("username")%></a></td>
      <td align="center"><%=conn.execute("select count(id) from ks_teamtopic where parentid=" & rs("id"))(0)%></td>
      <td align="center"><%=rs("adddate")%></td>
    </tr>
  <%

   rs.movenext
	  	I = I + 1
		  If I >= MaxPerPage Then Exit Do
	  loop
  End Sub
  
  '会员列表
  Sub ShowUsers()
		 MaxPerPage =10
		  If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
	%>
	<div id="user_list">
	<%
	dim rs:set rs=server.createobject("adodb.recordset")
	rs.open "select * from ks_teamusers where teamid=" &id & " and status=3",conn,1,1
	if not rs.eof then
			TotalPut = RS.recordcount
			if (TotalPut mod MaxPerPage)=0 then
				PageNum = TotalPut \ MaxPerPage
			else
				PageNum = TotalPut \ MaxPerPage + 1
			end if
			If CurrentPage = 1 Then
					call showuser(rs)
			Else
					If (CurrentPage - 1) * MaxPerPage < totalPut Then
					RS.Move (CurrentPage - 1) * MaxPerPage
					call showuser(rs)
					Else
					CurrentPage = 1
					call showuser(rs)
					End If
			End If
	end if
	rs.close:set rs=nothing			
	response.write "</div>"  
	 Response.Write "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|个||2"
  End Sub
  
Sub ShowUser(rs)
 dim i
do while not rs.eof 
	  dim rsu:set rsu=server.createobject("adodb.recordset")
	   rsu.open "select * from ks_user where username='" & rs("username") & "'",conn,1,1
	   if not rsu.eof then
		  Dim UserFaceSrc:UserFaceSrc=rsu("UserFace")
		  if lcase(left(userfacesrc,4))<>"http" then userfacesrc=KS.setting(2) & userfacesrc
	   %>
		<ul>
		<li class="u1"><img onerror="this.src='<%=KS.GetDomain%>images/face/boy.jpg';" src="<%=userfacesrc%>" border=0 width="60" height="60"></li>
		<li class="u2"><a href="../space/?<%=rsu("username")%>" target=_blank><%=rs("username")%></a>
		<%if KS.C("UserName")=groupadmin and KS.C("UserName")<>rsu("username") then%>
		 <a href="?action=deluser&username=<%=rsu("username")%>&id=<%=id%>" style="color:#ff6600" onclick="return(confirm('确定将此会员踢出吗?'))">X踢出</a>
		<%end if%>
		</li>
		<li class="u3">(<%=rsu("province")%><%=rsu("city")%>)</li>
		</ul>
	<%
	  end if
	  rsu.close:set rsu=nothing
	  rs.movenext
	  	I = I + 1
		  If I >= MaxPerPage Then Exit Do
	 loop
End Sub

	Private Function bbimg(strText)
		Dim s,re
        Set re=new RegExp
        re.IgnoreCase =true
        re.Global=True
		s=strText
		re.Pattern="<img(.[^>]*)([/| ])>"
		s=re.replace(s,"<img$1/>")
		re.Pattern="<img(.[^>]*)/>"
		s=re.replace(s,"<img$1 onload=""resizepic(this);"" onclick=""window.open(this.src)"" style=""cursor:pointer""/>")
		bbimg=s
	End Function	


     '显示帖子列表
		function showtopic()
		 dim tid:tid=KS.chkclng(KS.S("tid"))
		 dim rs:set rs=server.createobject("adodb.recordset")
		 rs.open "select b.username,b.userface,b.userid,a.* from ks_teamtopic a ,ks_user b where a.username=b.username and a.id=" &tid,conn,1,1
		 if rs.eof and rs.bof then
		   rs.close:set rs=nothing
		   call KS.alert("参数传递出错!","")
		   exit function
		 end if
		  Dim UserFaceSrc:UserFaceSrc=rs(1)
		  if lcase(left(userfacesrc,4))<>"http" then userfacesrc=KS.GetDomain & userfacesrc
		 response.write "<table width=""99%"" cellpadding=""1"" cellspacing=""1"" bgcolor=""#efefef"">"
		 response.write "    <tr>"
		 response.write "	  <td colspan=""2"" style=""font-size:14px;font-weight:bold;padding-left:15px;"" bgcolor=""#EDF5F9""><strong>"&rs("title")&"</strong></td>"
		 response.write "	</tr>"
		 response.write "	<tr>"
		 response.write "		<td width=""100"" style=""padding:10px;"" align=""center"" bgcolor=""#FFFFFF"">"
		 response.write "			<img border=""0"" style=""border:#efefef 1px solid; width:56px;height:56px"" src=""" & userfacesrc & """ /><br>"
		 response.write "			<a href=""../space/?" & rs(0) & """ target=""_blank"">" & rs(0) & "</a>"
		 response.write "		</td>"
		 response.write "		<td valign=""top""  style=""padding:10px;""  bgcolor=""#FFFFFF"">发表于<span>" & rs("adddate") & "<br/><div class=""Content""> " & KS.HtmlCode(replace(rs("content"),"alt=","")) & "</div></td>"
		 response.write "	</tr>"
		 response.write "	<tr>"
		 response.write "		<td align=""center"" bgcolor=""#FFFFFF"">楼主</td>"
		 response.write "		<td bgcolor=""#FFFFFF"">IP:" & rs("userip") & "&nbsp;&nbsp;&nbsp;&nbsp;"
		 response.write "		<a href=""#add_comment"">回复(" & conn.execute("select count(id) from ks_teamtopic where parentid=" & tid)(0) & ")</a> "
				if rs(0)=KS.C("UserName") or KS.C("UserName")=groupadmin then
				  		 response.write "<a href='group.asp?action=deltopic&id=" & id & "&tid=" & rs("id") & "' onclick='return(confirm(""确定删除该主题吗?""))'>删除</a>"
				   if KS.C("username")=groupadmin then
					 if rs("istop")=1 then
		                response.write "&nbsp; <a href='group.asp?action=settop&id="&id&"&tid=" & rs("id") &"'>取消置顶</a> "
					 else
					  response.write "&nbsp; <a href='group.asp?action=settop&id="&id&"&tid=" & rs("id") &"'>设为置顶</a> "
					 end if
					 if rs("isbest")=1 then
					 		 response.write "<a href='group.asp?action=setbest&id="&id&"&tid=" & rs("id") &"'>取消精华</a>"
					 else
					 		 response.write "<a href='group.asp?action=setbest&id="&id&"&tid=" & rs("id") &"'>设为精华</a>"
					 end if
				   end if
				  end if

		 response.write "		</td>"
		 response.write "	</tr>"
		 response.write "</table>"
		
		
		
			MaxPerPage=10
		 response.write "<div id=""comment_list"">"
		  CurrentPage=KS.ChkClng(KS.S("Page"))

		  
		 If CurrentPage<=0 Then CurrentPage=CurrentPage+1
		 dim rsp:set rsp=server.createobject("adodb.recordset")
		 rsp.open "select b.username,b.userid,b.userface,a.* from ks_teamtopic a, ks_user b where a.username=b.username and parentid=" & tid & " order by adddate desc",conn,1,1
		 if not rsp.eof then
				TotalPut = RSp.recordcount
				if (TotalPut mod MaxPerPage)=0 then
					PageNum = TotalPut \ MaxPerPage
				else
					PageNum = TotalPut \ MaxPerPage + 1
				end if
				If CurrentPage = 1 Then
					    Call ShowReplayContent(rsp)
				Else
						If (CurrentPage - 1) * MaxPerPage < totalPut Then
						RSp.Move (CurrentPage - 1) * MaxPerPage
						Call ShowReplayContent(rsp)
						Else
						CurrentPage = 1
						Call ShowReplayContent(rsp)
						End If
				End If
		end if
		
		 response.write "</div>"				
		rs.close:set rs=nothing
	    Response.Write "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|个||2"
		
		end function
		
Function ShowReplayContent(rsp)
  dim i,UserFaceSrc
  do while not rsp.eof
   		 UserFaceSrc=rsp("UserFace")
		  if lcase(left(userfacesrc,4))<>"http" then userfacesrc=KS.GetDomain & userfacesrc
				 response.write "<table width=""99%"" cellpadding=""1"" cellspacing=""1"" bgcolor=""#efefef"">"
				 response.write "<tr>"
				 response.write " <td colspan=""2"" style=""padding-left:15px;"" bgcolor=""#EDF5F9""><strong>" & rsp("title") & "</strong></td>"
				 response.write "	</tr>"
				 response.write "	<tr>"
				 response.write "		<td  style=""padding:10px;"" width=""100"" align=""center"" bgcolor=""#FFFFFF"">"
				 response.write "			<img style=""border:#efefef 1px solid; width:56px;height:56px"" src=""" & userfacesrc & """ /><br /><a href=""?" & rsp(0) & """ target=""_blank"">" & rsp(0) & "</a></td>"
				 response.write "		<td valign=""top""  style=""padding:10px;""  bgcolor=""#FFFFFF"">发表于<span>"&rsp("adddate") &"<div class=""Content"">" & rsp("content") & "</div></td>"
				 response.write "	</tr>"
				 response.write "	<tr>"
				 response.write "		<td bgcolor=""#FFFFFF"" align=""center"">第 "&i+1 & " 楼</td>"
				 response.write "		<td bgcolor=""#FFFFFF"">IP:" & rsp("userip") & "&nbsp;&nbsp;&nbsp;&nbsp;"
		if rsp(0)=KS.C("UserName") or KS.C("UserName")=groupadmin then
				 response.write "<a href='group.asp?action=deltopic&flag=replay&id=" & id & "&tid=" & rsp("id") & "' onclick='return(confirm(""确定删除该回复吗?""))'>删除</a>"
		  end if
				 response.write "</td>"
				 response.write "	</tr>"
				 response.write "</table>"
	  rsp.movenext
	  	I = I + 1
		  If I >= MaxPerPage Then Exit Do
   loop
 rsp.close:set rsp=nothing
End Function

 End Class 
%>