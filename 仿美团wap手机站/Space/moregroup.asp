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
        Private KS, KSRFObj,str,totalput,maxperpage,currpage
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
			  FCls.RefreshType = "MoreGroup" '设置刷新类型，以便取得当前位置导航等
			  Application(KS.SiteSN & "RefreshFolderID") = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
			  If Trim(FileContent) = "" Then FileContent = "空间副模板不存在!"
			  grouplist
			 FileContent=Replace(FileContent,"{$ShowMain}",str)
			FileContent=KSRFObj.KSLabelReplaceAll(FileContent)
		   Response.Write FileContent  
		End Sub
		
		'圈子列表
	Sub GroupList()
		 MaxPerPage =KS.ChkClng(KS.SSetting(11))
		 dim classid:classid=KS.ChkClng(KS.S("ClassID"))
		 dim recommend:recommend=KS.ChkClng(KS.S("recommend"))
		   CurrPage = KS.ChkClng(KS.G("page"))
		  If CurrPage<=0 Then CurrPage=1
	   
		  dim param:param=" where verific=1"
          if classid<>0 then param=param & " and a.classid=" & classid
			 if recommend<>0 then param=param & " and  recommend=1"
		 if ks.s("key")<>"" then param=param & " and teamname like '%" & ks.r(ks.s("key")) &"%'"
		  str=str & "  <div class=""groupmore"">" & vbcrlf
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "select a.*,b.classname from KS_team a inner join ks_teamclass b on a.classid=b.classid " & Param & " order by id desc",Conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				 str=str & "<ul><li>没有用户创建圈子！</li></ul>"
				 Else
							totalPut = RSObj.RecordCount
							If currpage >1 and (currpage - 1) * MaxPerPage < totalPut Then
								RSObj.Move (currpage - 1) * MaxPerPage
							End If
							call ShowGroup(RSObj)

				 End If
		 
		 str=str &"            </div>" & vbcrlf
		 RSObj.Close:Set RSObj=Nothing
		 str=str & KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
		 str=str & "<div class=""clear""></div>"
		 str=str &"<table border=""0"" cellpadding=""1"" cellspacing=""1"" align=""center"" width=""98%"" class=""spacesear"" >" &vbcrlf
		  str=str & "<form name=""myform"" action=""moregroup.asp"" method=""get""/> <tr height=""22"">"
	   str=str & "<td style=""text-align:left; padding-left:15px;"" colspan=2><strong>按圈子名称搜索：</strong><input style=""border:1px #000 solid;height:18px;"" type=""text"" size=""12"" name=""key"">&nbsp;&nbsp;<input type=""submit"" value= "" 查 找 "" class=""btn""></td>"
	   str=str & "</form></tr>"
	   str=str & "</table><br/><br/>"
	 End Sub
			 
	 Sub ShowGroup(RS)		 
		 Dim I
		 Do While Not RS.Eof 
		     str=str &"<ul>"
		     str=str &"<li>"& vbcrlf
			 str=str & " <a href=""group.asp?id=" & rs("id") & """ title=""" & rs("teamname") & """ target=""_blank""><img align=""left"" src=""" & rs("photourl") & """ border=""0"" style=""border:1px solid #f1f1f1;padding:2px;margin-right:6px;""></a>"
			str=str & "<a class=""teamname"" href=""group.asp?id=" & rs("id") & """ title=""" & rs("teamname") & """ target=""_blank""> " & rs("TeamName") & "</a><br>创建者：" & rs("username") & "<br>创建时间:" &rs("adddate") & "<br>圈子分类：" & rs("classname") & "<br>主题/回复：" & conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "  and parentid=0")(0) & "/" & conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id"))(0) & "&nbsp;&nbsp;&nbsp;成员:" & conn.execute("select count(username)  from ks_teamusers where status=3 and teamid=" & rs("id"))(0) & "人  </li>"
			str=str & "</ul>"
			rs.movenext
			I = I + 1
		  If I >= MaxPerPage Then Exit Do
		 Loop
	 End Sub
		
		
	
End Class
%>
