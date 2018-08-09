<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%

Response.Buffer=true
Dim KSCls
Set KSCls = New RefreshHtmlSave
KSCls.Kesion()
Set KSCls = Nothing
Const PauseNum=100  '生成多少篇时暂停，设置为0不暂停止。设置暂停可以缓解生成时服务器的压力
Const PauseTime=2    '暂停时间，单位：秒
Class RefreshHtmlSave
        Private KS,KSRObj
		Private RefreshFlag,f
		Private ReturnInfo,FsoHtmlList
		Private StartRefreshTime
		Private ChannelID,ItemName,Table
		Private Types
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KSRObj=Nothing
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
			With KS
			Server.ScriptTimeOut=9999999
			Types = Request("Types")             'Content 生成内容页操作 Folder 生成栏目操作
			RefreshFlag = Request("RefreshFlag") '取得是按何种类型刷新,如New只发布最新的指定篇数文章
			ChannelID = Request("ChannelID")     '按频道处理
			FCls.ChannelID=ChannelID
			
			If RefreshFlag<>"IDS" Then
				If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "20005") Then                '检查权限
					  Call KS.ReturnErr(1, "")
					  response.End()
				End If
			End If
			
			
			f=Request("f")
			If KS.S("FsoType")="1" Then
			FCls.FsoListNum=0
			ElseIf KS.S("FsoType")="2" Then
			FCls.FsoListNum=KS.ChkCLng(KS.S("FsoListNum"))
			Else
			FCls.FsoListNum=KS.ChkClng(KS.C_S(ChannelID,35))
			End If
			if f="task" then 
			   FCls.FsoListNum=3
			end if



			FCls.ItemUnit = KS.C_S(ChannelID,4)
			
	
			'刷新时间
			StartRefreshTime = Request("StartRefreshTime")
			If StartRefreshTime = "" Then StartRefreshTime = Timer()
			Table=KS.C_S(ChannelID,2)
			ItemName=KS.C_S(ChannelID,3)
			Select Case Types
			 Case "Content"
			            If KS.C_S(ChannelID,7)<>1 and KS.C_S(ChannelID,7)<>2 Then Call KS.AlertHistory("KesionCMS系统提醒您：\n\n1、此模型内容页没有启用生成静态HTML功能\n\n2、请到模型管理->模型信息设置启用生成静态Html功能",-1):Exit Sub
						Call RefreshContent
			 Case "Folder"
			          
			          If ChannelID<>0 and KS.C_S(ChannelID,7)<>1 Then Call KS.AlertHistory("KesionCMS系统提醒您：\n\n1、此模型栏目页没有启用生成静态HTML功能\n\n2、请到模型管理->模型信息设置启用生成静态Html功能",-1):Exit Sub
  
						Call RefreshFolder
			End Select
			End With
		End Sub
		
		Sub Main()
		  With KS
		  .echo ("<html>")
		  .echo ("<head>")
		  .echo ("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">")
		  .echo ("<title>系统信息</title>")
		  .echo ("<script src='../../ks_inc/jquery.js'></script>")
		  .echo ("<script src='../../ks_inc/kesion.box.js'></script>")
		  .echo ("<script type='text/javascript'>")
		  .echo (" function show()")
		  .echo (" { var p=new KesionPopup();")
		  .echo ("  p.PopupImgDir='../';")
		  .echo ("  var str=""<div style='height:60px;line-height:60px' id='fsotips'>正在整理数据,请稍候!</div>"";")
		  .echo ("  p.popupTips('生成提示',str,510,300);")
		  .echo (" }")
		  .echo ("</script>")
		  .echo ("</head>")
		  .echo ("<link rel=""stylesheet"" href=""Admin_Style.css"">")
		'  If RefreshFlag<>"ID" Then
		'  .echo ("<body oncontextmenu=""return false;"" scroll=no>")
		'  Else
		  .echo ("<body oncontextmenu=""return false;"" scroll=no style='background-color:transparent'>")
		'  End If
		  If RefreshFlag="ID" Then
              .echo "<div style=""display:none"">"
				.echo "<br><br><br><table style=""display:none"" id=""BarShowArea"" width=""400"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1"">"
		 Else
				.echo "<br><br><br><table id=""BarShowArea"" width=""400"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1"">"
		 End iF
				.echo "<tr> "
				.echo "<td bgcolor=000000>"
				.echo " <table width=""400"" border=""0"" cellspacing=""0"" cellpadding=""1"">"
				.echo "<tr> "
				.echo "<td bgcolor=ffffff height=9><img src=""../images/114_r2_c2.jpg"" width=0 height=10 id=img2 name=img2 align=absmiddle></td></tr></table>"
				'.echo "<td bgcolor=ffffff height=9><span width=0 height=16 id=img2 name=img2 align=absmiddle bgcolor='#000000'></span></td></tr></table>"
				.echo "</td></tr></table>"
				.echo "<table width=""550"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1""><tr> "
				.echo "<td align=center> <span id=txt2 name=txt2 style=""font-size:9pt"">0</span><span id=txt4 style=""font-size:9pt"">%</span></td></tr> "
				.echo "<tr><td align=center><span id=txt3 name=txt3 style=""font-size:9pt"">0</span></td></tr>"
				.echo "</table>"
			
			 .echo ("<table width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
			 .echo (" <tr>")
			 .echo ("   <td height=""50"">")
			 .echo ("     <div align=""center""> ")
			 .echo (ReturnInfo)
			 .echo ("       </div></td>")
			 .echo ("   </tr>")
			 .echo ("</table>")
			 .echo ("</div>")
		 

		 .echo ("<table width=""100%""   border=""0"" cellpadding=""0"" cellspacing=""0"">")
		 .echo (" <tr>")
		 .echo ("   <td height=""50"" id=""fsohtml"">")
		 .echo (FsoHtmlList)
		 .echo ("      </td>")
		 .echo ("   </tr>")
		 .echo ("</table>")
		 .echo ("</body>")
		 .echo ("</html>")
		 End With
		End Sub
		
		'================================================================================================================================
'                                                     以下为本模块相应处理的函数		'================================================================================================================================
		'生成栏目的处理过程
		Sub RefreshFolder()
		With KS
		Dim FolderID, R_Sql, RefreshTotalNum, R_RS, NewsTotalNum, NewsNo		  
		 If NewsNo = "" Then NewsNo = 0
		  Select Case RefreshFlag
		    Case "ID"
			    FolderID = Trim(Request("FolderID"))
			    R_Sql = "Select * from KS_Class where ChannelID=" & ChannelID &" and  DelTF=0 And ID ='" & FolderID & "'"
			Case "IDS"
			    FolderID = Replace(Replace(Request("ID")," ",""),",","','")
			    R_Sql = "Select * from KS_Class where ID IN('" & FolderID & "')"
			Case "Folder"
				FolderID = Trim(Request("FolderID"))
				R_Sql = "Select * from KS_Class where ChannelID=" & ChannelID &" and  DelTF=0 And ID IN (" & FolderID & ") Order By FolderOrder ASC"
		   Case "All"
				R_Sql = "Select * from KS_Class where ChannelID=" & ChannelID &" and ClassType<>2 and DelTF=0 Order By FolderOrder ASC"
		   Case Else
			R_Sql = ""
		  End Select
		
		Call Main
		If R_Sql <> "" Then
			Set R_RS = Server.CreateObject("ADODB.RecordSet")
			R_RS.Open R_Sql, Conn, 1, 1
			If R_RS.EOF Then
				.echo "<script>img2.width=""0"";" & vbCrLf
				.echo "txt2.innerHTML=""没有可生成的" & ItemName & "栏目！<br><br><input name='button1' type='button' onclick=javascript:location.href='RefreshHtml.asp?Action=ref&ChannelID=" & ChannelID &"'; class='button' value=' 返 回 '>"";" & vbCrLf
				.echo "txt3.innerHTML="""";" & vbCrLf
				.echo "txt4.innerHTML="""";" & vbCrLf
				.echo "document.all.BarShowArea.style.display='none';" & vbCrLf
				.echo "</script>" & vbCrLf
				R_RS.Close:Set R_RS = Nothing
				.die ""
			Else
			       NewsTotalNum = R_RS.RecordCount
				   For NewsNo=1 to NewsTotalNum
				    ChannelID=R_RS("ChannelID") : FCls.ChannelID=ChannelID
				    If R_RS("ClassPurview")=2 Then
				     FsoHtmlList="<table border=""0"">"_
								& "<tr><td><li><strong>ID号为：</strong></li></td><td> <font color=red>"  & R_RS("ID") & "</font> 的栏目没有生成!</td></tr>"_
								& "<tr><td><li><strong>原 因：</strong></li></td><td>该栏目设置为认证栏目"_
						& "</table>"		
					Else
						Dim FsoHtmlPath:FsoHtmlPath=KS.GetFolderPath(R_RS("ID"))
						FsoHtmlList="<table border=""0"">"_
									& "<tr><td><li><strong>ID 号 为：</strong></li></td><td> <font color=red>"  & R_RS("ID") & "</font> 的栏目已生成</td></tr>"_
									& "<tr><td><li><strong>栏目名称：</strong></li></td><td><font color=red>" & R_RS("FolderName") & "</font></li></td><tr>" _
									& "<tr><td><li><strong>生成路径：</strong></li></td><td><a href=""" & FsoHtmlPath & """ target=""_blank"">" & FsoHtmlPath & "</a></li></td><tr>" _
									& "</table>"				
						Call KSRObj.RefreshFolder(ChannelID,R_RS)  '调用栏目刷新函数
					End If
				
				    If RefreshFlag="ID" Then Call InnerJS(NewsNo,NewsTotalNum,"个栏目"):.Die ""
					
					Call InnerJS(NewsNo,NewsTotalNum,"个栏目")
					R_RS.MoveNext
					if Not Response.IsClientConnected then Exit FOR
				  Next
				.echo "<script>"
				.echo "fsohtml.innerHTML='';" & vbCrLf
				.echo "img2.width=400;" & vbCrLf
				.echo "txt2.innerHTML=""生成" & ItemName & "栏目结束！100"";" & vbCrLf
				.echo "txt3.innerHTML=""总共生成了 <font color=red><b>" & NewsTotalNum & "</b></font> 个" & ItemName & "栏目,总费时:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font> 秒<br><br><input name='button1' type='button' onclick=javascript:location='RefreshHtml.asp?Action=ref&ChannelID=" & ChannelID &"'; class='button' value=' 返 回 '>"";" & vbCrLf
				.echo "img2.title=""(" & NewsNo & ")"";</script>" & vbCrLf
				'定时任务,关闭
				if f="task" then
				 KS.Echo "<script>setTimeout('window.close();',3000);</script>"
				end if
				
				R_RS.Close:Set R_RS = Nothing
			End If
		Else
				.echo "<script>img2.width=""0"";" & vbCrLf
				.echo "txt2.innerHTML=""没有可生成的栏目！<br><br><input name='button1' type='button' onclick=javascript:location='RefreshHtml.asp?Action=ref&ChannelID=" & ChannelID & "'; class='button' value=' 返 回 '>"";" & vbCrLf
				.echo "txt3.innerHTML="""";" & vbCrLf
				.echo "txt4.innerHTML="""";" & vbCrLf
				.echo "document.all.BarShowArea.style.display='none';" & vbCrLf
				.echo "</script>" & vbCrLf
		End If
		End With
		End Sub
		
		'生成内容页的处理过程
		Sub RefreshContent()
		Dim AlreadyRefreshByID, NowNum, R_Sql, R_RS, TotalNum,ID,UpdateSql
		Dim StartDate, EndDate, FolderID, RefreshTotalNum,StartID,EndID
		AlreadyRefreshByID = Request.QueryString("AlreadyRefreshByID")
		RefreshTotalNum = Request.QueryString("RefreshTotalNum")
		NowNum = Request.QueryString("NowNum") '正在刷新第几篇文章
		if KS.G("refreshtf")="1" then
		R_Sql=" Where refreshtf=0 and Verific=1"
		else
		R_Sql=" Where Verific=1"
		end if
		With KS
		If NowNum = "" Then NowNum = 0
		  Select Case RefreshFlag
		   Case "ID"
			    ID=KS.G("ID")
				UpdateSql="Update "& Table & " Set RefreshTF=1 "  & R_SQL&" and ID IN(Select top 2 id from " & Table & R_Sql & " And ID<=" & id & " Order By ID Desc)"
				R_Sql="Select Top 2 * From " & Table & R_SQL&" and ID IN(Select top 2 id from " & Table & R_Sql & " And ID<=" & id & " Order By ID Desc) Order By ID"
				RefreshTotalNum=conn.execute("select count(id) from " & Table  &" where verific=1 and ID<=" & ID)(0)
				If RefreshTotalNum>2 Then RefreshTotalNum=2
		   Case "IDS"
			    ID=KS.FilterIds(KS.G("ID"))
				If ID="" Then KS.Die "err!"
				UpdateSql="Update "& Table & " Set RefreshTF=1 "  &R_SQL&" and ID IN(" & ID & ")"
				R_Sql="Select Top 200 * From " & Table & R_SQL&" and ID IN(" & ID & ") Order By ID desc"
				RefreshTotalNum=conn.execute("select count(id) from " & Table  &" where verific=1 and ID in(" & ID & ")")(0)
		   Case "InfoID"
				 StartID = KS.ChkClng(KS.G("StartID"))
				 EndID = KS.ChkClng(KS.G("EndID"))
				 RefreshTotalNum=conn.execute("select count(id) from " & Table  & R_Sql & " and ID>= " & StartID & " And  ID <=" & EndID)(0)
				 UpdateSql="Update "& Table & " Set RefreshTF=1 " & R_Sql & " and ID>= " & StartID & " And  ID <=" & EndID
				 R_Sql = "Select * from " & Table  & R_Sql & " and ID>= " & StartID & " And  ID <=" & EndID & " order by ID desc"
			Case "New"
			  TotalNum = KS.ChkCLng(Request("TotalNum"))
			  If TotalNum >conn.execute("select count(id) from "& Table & R_SQL )(0) Then TotalNum = conn.execute("select count(id) from "& Table & R_SQL)(0)
			  RefreshTotalNum = TotalNum
			  If TotalNum=0 Then TotalNum=1
			   UpdateSql="Update "& Table & " Set RefreshTF=1 Where ID in(Select Top " & TotalNum & " ID from " & Table  & R_SQL& " Order By ID Desc)"
			  R_Sql="Select Top " & TotalNum & " * from " & Table  & R_SQL& " Order By ID Desc"
		   Case "Date"
			  StartDate = Request("StartDate"):EndDate = DateAdd("d", 1, Request("EndDate"))
			                     'Access
				 RefreshTotalNum=conn.execute("select count(id) from " & Table  & R_Sql & " and AddDate>= #" & StartDate & "# And  AddDate <=#" & EndDate & "#")(0)
				 UpdateSql="Update "& Table & " Set RefreshTF=1 "& R_Sql & " and AddDate>= #" & StartDate & "# And  AddDate <=#" & EndDate & "#"
				 R_Sql = "Select * from " & Table  & R_Sql & " and AddDate>= #" & StartDate & "# And  AddDate <=#" & EndDate & "# order by ID desc"
		   Case "All"
		      RefreshTotalNum=conn.execute("select count(id) from " & Table  & R_Sql)(0)
			  UpdateSql="Update "& Table & " Set RefreshTF=1 "& R_Sql 
			  R_Sql = "Select * from " & Table  & R_Sql & " order by ID desc"
		  Case "Folder"
			 FolderID = Trim(Replace(Request("FolderID")," ",""))
			 TotalNum = KS.ChkCLng(Request("TotalNum"))
			  If (TotalNum >conn.execute("select count(id) from " & Table  & R_Sql& " And Tid IN(" & FolderID & ")")(0)) Or  TotalNum=0 Then TotalNum = conn.execute("select count(id) from " & Table  & R_Sql& " And Tid IN(" & FolderID & ")")(0)
			  RefreshTotalNum = TotalNum
			  If TotalNum=0 Then TotalNum=1
			  If KS.ChkCLng(Request("TotalNum"))<>0 Then
			   UpdateSql="Update "& Table & " Set RefreshTF=1 Where ID In(Select top " & TotalNum & " ID from " & Table  & R_Sql & " And Tid IN(" & FolderID & "))"
			   R_Sql = "Select top " & TotalNum & " * from " & Table  & R_Sql & " And Tid IN(" & FolderID & ") order by ID desc"
			  Else
			  	UpdateSql="Update "& Table & " Set RefreshTF=1 " & R_Sql & " And Tid IN(" & FolderID & ")"

			   R_Sql = "Select * from " & Table  & R_Sql & " And Tid IN(" & FolderID & ") order by ID desc"
			  End If
		  Case "Pause"
		     UpdateSql=Request.QueryString("UpdateSql")
		     R_Sql=Request.QueryString("R_Sql")
			 RefreshTotalNum=KS.ChkClng(KS.G("RefreshTotalNum"))
		Case Else
			R_Sql = ""
			RefreshTotalNum = 0
		End Select
		Call Main
		If R_Sql <> "" Then
		  If DataBaseType=1 Then
			    Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
				Set Cmd.ActiveConnection=conn
				Cmd.CommandText="KS_TSqls"
				Cmd.CommandType=4	
				CMD.Prepared = true 	
				Cmd.Parameters.Append cmd.CreateParameter("@SQL",202,1,8000,R_Sql)
				Set R_RS=Cmd.Execute
				Set Cmd=Nothing
			Else
			    Set R_RS=Conn.Execute(R_Sql)
			End If
			'Set R_RS = Server.CreateObject("ADODB.RecordSet")
			'R_RS.Open R_Sql, Conn, 1, 1
			If R_RS.EOF And R_RS.BOF Then
				.echo "<script>img2.width=""0"";" & vbCrLf
				.echo "txt2.innerHTML=""没有可生成的内容页！<br><br><input name='button1' type='button' onclick=javascript:location='RefreshHtml.asp?Action=ref&ChannelID=" & ChannelID  &"'; class='button' value=' 返 回 '>"";" & vbCrLf
				.echo "txt3.innerHTML="""";" & vbCrLf
				.echo "txt4.innerHTML="""";" & vbCrLf
				.echo "document.all.BarShowArea.style.display='none';" & vbCrLf
				.echo "</script>" & vbCrLf
				Response.Flush
				R_RS.Close:Set R_RS=Nothing
				Exit Sub
			Else
				'On Error Resume Next
				Dim CurrNowNum:CurrNowNum=KS.ChkClng(KS.G("CurrNowNum"))
				If CurrNowNum=0 Then CurrNowNum=1
				R_RS.Move(CurrNowNum-1)
				For NowNum=CurrNowNum To RefreshTotalNum
				     Dim DocXML:Set DocXML=KS.arrayToXml(R_RS.GetRows(1),R_RS,"row","root")
				     Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
					  KSRObj.ModelID=ChannelID
					  KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 
					If KS.C_S(ChannelID,7)=0 Then
					      FsoHtmlList=GetRefreshErr(KSRObj.Node,ItemName)
					Else
						  FsoHtmlList=GetRefreshSucc(KSRObj.Node,ItemName)
						  KSRObj.RefreshContent()
				    End If
				
				If Err.Number <> 0 Then
				 FsoHtmlList = "操作失败!<br><font color=red>" & Err.Description & "</font>"
				 Call InnerJS(NowNum,RefreshTotalNum,KS.C_S(ChannelID,4))
				End If
				If RefreshFlag="ID" and NowNum=2 Then
				 Call InnerJS(NowNum,RefreshTotalNum,KS.C_S(ChannelID,4)):R_RS.Close:Set R_RS=Nothing:.Die ""
				Else
				 Call InnerJS(NowNum,RefreshTotalNum,KS.C_S(ChannelID,4))
				End If
				
				if Not Response.IsClientConnected then Exit FOR
				If PauseNum>0 Then
					If RefreshTotalNum>1 and NowNum Mod PauseNum=0 Then
					    R_RS.Close:Set R_RS=Nothing
						 .echo "<script>"
						 .echo "fsohtml.innerHTML='<div style=""text-align:cdenter""><div style=""margin:10px;height:80px;padding:8px;border:1px dashed #cccccc;text-align:left;""><img src=""../images/succeed.gif"" align=""left""><br>&nbsp;&nbsp;&nbsp;&nbsp;<b>温馨提示：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;以免过度占用服务器资源，系统暂停" & PauseTime & "秒后继续<img src=""../../images/default/wait.gif""><br>&nbsp;&nbsp;&nbsp;&nbsp;如果" & PauseTime & "秒后没有继续，请点此<a href=""RefreshHtmlSave.asp?CurrNowNum=" & NowNum+1 & "&ChannelID=" & ChannelID & "&RefreshFlag=Pause&Types=" & Types & "&StartRefreshTime=" & StartRefreshTime & "&UpdateSql=" & server.URLEncode(UpdateSql) & "&R_Sql=" & Server.UrlEncode(R_Sql) & "&RefreshTotalNum=" & RefreshTotalNum & """><font color=red>继续</font></a>或点此<a href=""refreshhtml.asp?Action=ref&channelid=" & channelid & """><font color=red>停止</font></a>!</div></div>';" & vbCrLf
						 .echo "</script>" &vbcrlf
						 .die "<meta http-equiv=""refresh"" content=""" & PauseTime & ";url=RefreshHtmlSave.asp?f=" & f & "&CurrNowNum=" & NowNum+1 & "&ChannelID=" & ChannelID & "&RefreshFlag=Pause&Types=" & Types & "&StartRefreshTime=" & StartRefreshTime & "&UpdateSql=" & server.URLEncode(UpdateSql) & "&R_Sql=" & Server.UrlEncode(R_Sql) & "&RefreshTotalNum=" & RefreshTotalNum & """>"
					End If
			   End If
			Next
			    If UpdateSql<>"" Then Conn.Execute(UpdateSql)
				.echo "<script>"
				.echo "fsohtml.innerHTML='';" & vbCrLf
				.echo "img2.width=400;" & vbCrLf
				.echo "txt2.innerHTML=""生成内容页结束！100"";" & vbCrLf
				.echo "txt3.innerHTML=""总共生成了 <font color=red><b>" & RefreshTotalNum & "</b></font> 条,总费时:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font> 秒<br><br><input name='button1' type='button' onclick=javascript:location='RefreshHtml.asp?Action=ref&ChannelID=" & ChannelID & "'; class='button' value=' 返 回 '>"";" & vbCrLf
				.echo "img2.title=""(" & NowNum & ")"";</script>" & vbCrLf
				'定时任务,关闭
				if f="task" then
				 KS.Echo "<script>setTimeout('window.close();',3000);</script>"
				end if
			End If
		Else
				.echo "<script>img2.width=""0"";" & vbCrLf
				.echo "txt2.innerHTML=""没有可生成的内容页！<br><br><input name='button1' type='button' onclick=javascript:location='RefreshHtml.asp?Action=ref&ChannelID=" & ChannelID & "'; class='button' value=' 返 回 '>"";" & vbCrLf
				.echo "txt3.innerHTML="""";" & vbCrLf
				.echo "txt4.innerHTML="""";" & vbCrLf
				.echo "document.all.BarShowArea.style.display='none';" & vbCrLf
				.echo "</script>" & vbCrLf
				'定时任务,关闭
				if f="task" then
				 KS.Echo "<script>setTimeout('window.close();',3000);</script>"
				end if
		End If
		End With
		End Sub
		
		Function GetRefreshErr(Node,ItemName)
		GetRefreshErr="<table border=""0"">"_
								& "<tr><td><li><strong>ID 号为：</strong></li></td><td> <font color=red>"  & Node.SelectSingleNode("@id").text & "</font> 的文章没有生成!</td></tr>"_
								& "<tr><td><li><strong>可能原因：</strong></li></td><td>1、" & ItemName & "频道没有启用生成静态HTML生成功能；<br>2、该" & ItemName & "所在的栏目为半开放栏目或是认证栏目；<br>3、该" & ItemName & "设置了需要扣点浏览、游客不能浏览或是设置为转向链接；<br>"_
						& "</table>"	
		End Function
		Function GetRefreshSucc(Node,ItemName)
		 Dim str,FsoHtmlPath:FsoHtmlPath= KS.GetItemURL(ChannelID,Node.SelectSingleNode("@tid").text,Node.SelectSingleNode("@id").text,node.SelectSingleNode("@fname").text)
		 str=""
		 if RefreshFlag<>"ID" Then str="<img src=""../images/succeed.gif"" align=""left""><table border=""0"">"
		 GetRefreshSucc=str & "<table border=""0"">"_
									& "<tr><td><li><strong>ID 号为：</strong></li></td><td> <font color=red>"  & Node.SelectSingleNode("@id").text & "</font> 的" & ItemName & "已生成</td></tr>"_
									& "<tr><td><li><strong>" & ItemName & "标题：</strong></li></td><td><font color=red>" & Node.SelectSingleNode("@title").text & "</font></li></td><tr>" _
									& "<tr><td><li><strong>生成路径：</strong></li></td><td><a href=""" & FsoHtmlPath & """ target=""_blank"">" & FsoHtmlPath & "</a></li></td><tr>" _
									& "</table>"
				  'conn.execute("update " & Table & " set refreshtf=1 where id=" & Node.SelectSingleNode("@id").text)
		End Function
		
		Sub InnerJS(NowNum,TotalNum,itemname)
		  With KS
				.echo "<script>"
				if RefreshFlag<>"ID" Then
				.echo "fsohtml.innerHTML='<div style=""margin:10px;height:80px;padding:8px;border:1px dashed #cccccc;text-align:left;"">" & replace(FsoHtmlList,"'","\'") & "</div>';" & vbCrLf
			    else
				.echo "fsohtml.innerHTML='" & replace(FsoHtmlList,"'","\'") & "';" & vbCrLf
				end if
				.echo "img2.width=" & Fix((NowNum / TotalNum) * 400) & ";" & vbCrLf
				.echo "txt2.innerHTML=""生成进度:" & FormatNumber(NowNum / TotalNum * 100, 2, -1) & """;" & vbCrLf
				.echo "txt3.innerHTML=""总共需要生成 <font color=red><b>" & TotalNum & "</b></font> " & itemname & ",<font color=red><b>在此过程中请勿刷新此页面！！！</b></font> 系统正在生成第 <font color=red><b>" & NowNum & "</b></font> " & itemname & """;" & vbCrLf
				.echo "img2.title=""(" & NowNum & ")"";</script>" & vbCrLf
				Response.Flush
		  End With
		End Sub
		
End Class
%> 