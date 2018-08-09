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
        Private KS, KSRFObj,Str,MaxPerPage,CurrPage,TotalPut
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
			   FCls.RefreshType = "Morexc" '设置刷新类型，以便取得当前位置导航等
			   Application(KS.SiteSN & "RefreshFolderID") = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
			   If Trim(FileContent) = "" Then FileContent = "空间副模板不存在!"
				 PhotoList
			   FileContent=Replace(FileContent,"{$ShowMain}",str)
			   FileContent=KSRFObj.KSLabelReplaceAll(FileContent)
		      Response.Write FileContent  
		End Sub
		
	  '相册列表
  Sub PhotoList()
		 MaxPerPage =KS.ChkClng(KS.SSetting(12))
		 dim classid:classid=ks.chkclng(ks.s("classid"))
		 dim recommend:recommend=ks.chkclng(ks.s("recommend"))
		  CurrPage = KS.ChkClng(KS.G("page"))
		  If CurrPage<=0 Then CurrPage=1
		 
		 dim rsc:set rsc=conn.execute("select classname,classid from ks_PhotoClass order by orderid")
		 if not rsc.eof then
		   str="<div class=""categorybox"">" & vbcrlf
		   str=str &"<ul><li>分类查看：</li>"
		   If classid=0 then 
		     str=str &"<li class=""curr""><a href='morephoto.asp'>所有分类</a></li>"
		   else
		     str=str &"<li><a href='morephoto.asp'>所有分类</a></li>"
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
	   str=str &"<div class=""albumlist""><ul>" &vbcrlf

	 dim param:param=" where status=1"
	 if classid<>0 then param=param & " and  classid=" & classid
	 if recommend<>0 then param=param & " and  recommend=1"
	 if ks.s("key")<>"" then param=param & " and XCName like '%" & ks.r(ks.s("key")) &"%'"
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * from KS_Photoxc " & Param & " order by id desc",Conn,1,1
		         If RS.EOF and RS.Bof  Then
				 str=str& "<div style=""border: #efefef 1px dotted;text-align:center"">没有创建相册！</div>"
				 Else
						totalPut = RS.RecordCount
						If CurrPage>1 and(CurrPage - 1) * MaxPerPage < totalPut Then
							RS.Move (CurrPage - 1) * MaxPerPage
						End If
						call showphoto(RS)
				  End If
		 RS.Close:Set RS=Nothing
		 str=str & "</div>"
		 str=str & KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
		 str=str & "<div class=""clear""></div>"
		 str=str &"<table border=""0"" class=""spacesear"" cellpadding=""1"" cellspacing=""1"" align=""center"" width=""98%"">" &vbcrlf
		  str=str & "<form name=""myform"" action=""morephoto.asp"" method=""get""/> <tr height=""22"">"
	   str=str & "<td style=""text-align:left; padding-left:15px;"" colspan=2><strong>按相册名称搜索：</strong><input style=""border:1px #000 solid;height:18px;"" type=""text"" size=""12"" name=""key"">&nbsp;&nbsp;<input type=""submit"" value= "" 查 找 "" class=""btn""></td>"
	   str=str & "</form></tr>"
	   str=str & "</table><br/><br/>"
  End Sub

	 Sub showphoto(rs)
	 	 Dim I,url
		 Do While Not RS.Eof 
		  
		  str=str & "<li>" &vbcrlf
			If KS.SSetting(21)="1" Then
			   Url="showalbum-" & RS("userid") & "-" & RS("id")
			Else
			   Url="../space/?" & RS("userid") &"/showalbum/" &RS("id")
			End If
			str=str &"<div class=""albumbg""><a href=""" & url &""" target=""_blank""><img style=""margin-left:-4px;margin-top:5px"" src=""" &RS("photourl") &""" width=""120"" height=""90"" border=0></a></div><B><a href=""" & Url &""">" &RS("xcname") &"</a></B> (" & RS("xps") & ")<font color=red>[" & GetStatusStr(RS("flag")) &"]</font>" & vbcrlf
			 str=str &"</li>"
			 RS.movenext
			I = I + 1
			If I >= MaxPerPage Then Exit Do
		Loop
		  
	 End Sub
	 
	 Function GetStatusStr(val)
           Select Case Val
		    Case 1:GetStatusStr="公开"
			Case 2:GetStatusStr="会员"
			Case 3:GetStatusStr="密码"
			Case 4:GetStatusStr="隐私"
		   End Select
			GetStatusStr="<font color=red>" & GetStatusStr & "</font>"
	 End Function
		

End Class
%>
