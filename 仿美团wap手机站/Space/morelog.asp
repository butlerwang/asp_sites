<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<%


Dim KSCls
Set KSCls = New Spacemore
KSCls.Kesion()
Set KSCls = Nothing

Class Spacemore
        Private KS, KSRFObj,str,CurrPage,TotalPut,MaxPerPage
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRFObj = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		      Dim FileContent
				   FileContent = KSRFObj.LoadTemplate(KS.SSetting(8))
				   FCls.RefreshType = "MorelOG" '设置刷新类型，以便取得当前位置导航等
				   Application(KS.SiteSN & "RefreshFolderID") = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
				   If Trim(FileContent) = "" Then FileContent = "空间副模板不存在!"
				   FileContent=Replace(FileContent,"{$ShowMain}",GetLogList())
				   FileContent=KSRFObj.KSLabelReplaceAll(FileContent)
		   Response.Write FileContent  
		End Sub
		Function GetLogList()
		 dim typeid:typeid=ks.chkclng(ks.s("classid"))
		 dim isbest:isbest=ks.chkclng(ks.s("isbest"))
		 CurrPage = KS.ChkClng(KS.G("page"))
		 If CurrPage<=0 Then CurrPage = 1
	     dim rsc:set rsc=conn.execute("select typename,typeid from ks_blogtype order by orderid")
		 if not rsc.eof then
		   str="<div class=""categorybox"">" & vbcrlf
		   str=str &"<ul><li>分类查看：</li>"
		   If typeid=0 then 
		     str=str &"<li class=""curr""><a href='morelog.asp'>所有分类</a></li>"
		   else
		     str=str &"<li><a href='morelog.asp'>所有分类</a></li>"
		   end if
			do while not rsc.eof
			 if typeid=rsc(1) then
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
	   str=str &"<table border=""0"" cellpadding=""1"" cellspacing=""1"" align=""center"" width=""98%"" backcolor=""#efefef"">" &vbcrlf
       str=str & " <tr height=""22"" bgcolor=""#f9f9f9"">" &vbcrlf
	   str=str &"     <td><strong>日志标题</strong></td>"
	   str=str & "    <td width=""100"" align=""center""><strong>分 类</strong></td>" &vbcrlf
	   str=str &"     <td width=""70"" align=""center""><strong>作者</strong></td>" &vbcrlf
	   str=str & "    <td align=""center""><strong>更新时间</strong></td></tr>" & vbcrlf

	 MaxPerPage=30
	 dim param:param=" where istalk<>1 and status=0"
	 if typeid<>0 then param=param & " and a.typeid=" & typeid
	 if isbest<>0 then param=param & " and best=1"
	 if ks.s("key")<>"" then param=param & " and Title like '%" & ks.r(ks.s("key")) &"%'"
 
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			rsobj.open "select a.*,b.typename from ks_blogInfo a inner join ks_blogType b on a.typeid=b.typeid " & param & " order by adddate desc" ,conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				 	str=str & "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=4><p>对不起，没有找到日志文章! </p></td></tr>"
				 Else
							totalPut = RSObj.RecordCount

								   If CurrPage>1 and (CurrPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrPage - 1) * MaxPerPage
									Else
										CurrPage = 1
								End If
									call ShowlogList(RSObj)

				           End If
		 str=str & "</table>"
		 str=str & KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
		 str=str & "<div class=""clear""></div>"
		 str=str &"<table border=""0"" cellpadding=""1"" cellspacing=""1"" align=""center"" width=""98%"">" &vbcrlf
		  str=str & "<form name=""myform"" action=""morelog.asp"" method=""get""/> <tr height=""22"">"
	   str=str & "<td style=""text-align:left"" colspan=2><strong>按博文标题搜索：</strong><input style=""border:1px #000 solid;height:18px;"" type=""text"" size=""12"" name=""key"">&nbsp;&nbsp;<input type=""submit"" value= "" 查 找 "" class=""btn""></td>"
	   str=str & "</form></tr>"
	   str=str & "</table><br/><br/>"
		 
		 
		 RSObj.Close:Set RSObj=Nothing
		GetLogList=str
  End Function

  Sub ShowLogList(rs)
   dim i,KSBCls
   Set KSBCls=New BlogCls
   do while not rs.eof
    if i mod 2=0 then
	  str=str & "<tr style=""background:#fff;height:22px"">"
	else
	  str=str &  "<tr style=""background:#FBFDFF;height:22px"">"
	end if
	 str=str & "<td><img src=""images/bullet.gif"" align=""absmiddle"" />"
	 str=str &" <a title=""" & rs("Title") & """ href=""" & KSBCls.GetCurrLogUrl(rs("userid"),rs("id")) & """ target=""blank"">"
	 str=str & KS.GotTopic(rs("title"),32) & "</a>"
	  if rs("best")=1 then str=str & "<font color=red>[精]</font>"
	 str=str & "</td>"
	 str=str & "<td align=""center"">" & rs("typename") & "</td>"
	 str=str & "<td align=""center"">" & rs("username") & "</td>"
	 str=str & "<td align=""center"">" & rs("adddate") & "</td>"
	 str=str & "</tr>"

      rs.movenext
	  	I = I + 1
		  If I >= MaxPerPage Then Exit Do
	  loop
	Set KSBCls=Nothing
  End Sub
		 
End Class
%>
