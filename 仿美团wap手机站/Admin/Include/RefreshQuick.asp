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
		Private Action
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
			Action = Request("Action")             'Content 生成内容页操作 Folder 生成栏目操作
			RefreshFlag = Request("RefreshFlag") '取得是按何种类型刷新,如New只发布最新的指定篇数文章
			ChannelID = Request("ChannelID")     '按频道处理
			FCls.ChannelID=ChannelID
			
			If Not KS.ReturnPowerResult(0, "M010007") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
			End if
			
			
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
			
			Select Case Action
			 Case "initial"  Call Initial()
			 Case "startleft" Call startleft()
			 Case "startright" Call startright()
			 Case "start" Call Start()
			 Case "settingsave" Call settingsave()
			 Case Else
			      ReturnInfo="<input type='button' class='button' value='一键快速生成HTML' onclick=""location.href='?action=initial';""/>"
				  ReturnInfo=ReturnInfo & "<br/><form name='myform' action='?action=settingsave' method='post'><strong>生成配置：</strong>每个模型只生成最新添加的 <input type='text' name='num' value='" & KS.ChkClng(KS.ReadSetting(1)) & "' class='textbox' style='text-align:center;height:23px;line-height:23px;width:50px'/> 篇文档 <input type='submit' value=' 保存设置 ' class='button'/><br/><br/><font color=red>tips:如果不限制请输入0,如果您的网站数据量很大，旧的数据一般不需要重新生成，所以可以在此处设置一个值。</font></form>"
			      Call Main()
			End Select
			End With
		End Sub
		
		Sub settingsave()
		  dim num:num=ks.chkclng(request("num"))
		  Call KS.settingsave(1,num)
		  KS.AlertHintScript "恭喜，保存设置成功!"
		End Sub
		
		Sub Initial()
		  Dim XMLStr,i,RS
		        Call Main
		        Response.Write "<div style='text-align:center'>正在初始化需要生成的项目..."
		            XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
					XMLStr=XMLStr&" <item>" &vbcrlf
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "select * from KS_Channel Where ChannelID not in(6,10) And FsoHtmlTF<>0 And ChannelStatus=1 order by ChannelID",conn,1,1
			    i=0
				IF Split(KS.Setting(5),".")(1)<>"asp" Then
						XMLStr=XMLStr & "  <fsoitem isfinish=""0"" id=""1"" channelid=""0"">"&vbcrlf
						XMLStr=XMLStr & "    <name><![CDATA[网站首页]]></name>" &vbcrlf
						XMLStr=XMLStr & "    <fsotype>0</fsotype>" &vbcrlf
					    XMLStr=XMLStr & "  </fsoitem>"&vbcrlf
                End If
				Do While Not RS.EOf
				      if rs("FsoHtmlTF")=1 or rs("channelid")=9 Then
					    if rs("channelid")=9 then
						 i=I+1
					    XMLStr=XMLStr & "  <fsoitem isfinish=""0"" id=""" & i &""" channelid=""" & rs("channelid") & """>"&vbcrlf
						XMLStr=XMLStr & "    <name><![CDATA[发布考试频道首页等]]></name>" &vbcrlf
						XMLStr=XMLStr & "    <fsotype>90</fsotype>" &vbcrlf
					    XMLStr=XMLStr & "  </fsoitem>"&vbcrlf
						end if
				        i=I+1
					    XMLStr=XMLStr & "  <fsoitem isfinish=""0"" id=""" & i &""" channelid=""" & rs("channelid") & """>"&vbcrlf
						XMLStr=XMLStr & "    <name><![CDATA[" & rs("channelname") &"]]></name>" &vbcrlf
						XMLStr=XMLStr & "    <fsotype>1</fsotype>" &vbcrlf
					    XMLStr=XMLStr & "  </fsoitem>"&vbcrlf
				      end if
					  	i=I+1
					    XMLStr=XMLStr & "  <fsoitem isfinish=""0"" id=""" & i &""" channelid=""" & rs("channelid") & """>"&vbcrlf
						XMLStr=XMLStr & "    <name><![CDATA[" & rs("channelname") &"]]></name>" &vbcrlf
						XMLStr=XMLStr & "    <fsotype>2</fsotype>" &vbcrlf
					    XMLStr=XMLStr & "  </fsoitem>"&vbcrlf
			     RS.MoveNext
				Loop
			RS.Close:Set RS=Nothing
			
			 If KS.Setting(78)<>0 Then   '自动生成专题
					  	i=I+1
					    XMLStr=XMLStr & "  <fsoitem isfinish=""0"" id=""" & i &""" channelid=""0"">"&vbcrlf
						XMLStr=XMLStr & "    <name><![CDATA[专题首页]]></name>" &vbcrlf
						XMLStr=XMLStr & "    <fsotype>1000</fsotype>" &vbcrlf
					    XMLStr=XMLStr & "  </fsoitem>"&vbcrlf
					  	i=I+1
					    XMLStr=XMLStr & "  <fsoitem isfinish=""0"" id=""" & i &""" channelid=""0"">"&vbcrlf
						XMLStr=XMLStr & "    <name><![CDATA[专题分类]]></name>" &vbcrlf
						XMLStr=XMLStr & "    <fsotype>1001</fsotype>" &vbcrlf
					    XMLStr=XMLStr & "  </fsoitem>"&vbcrlf
					  	i=I+1
					    XMLStr=XMLStr & "  <fsoitem isfinish=""0"" id=""" & i &""" channelid=""0"">"&vbcrlf
						XMLStr=XMLStr & "    <name><![CDATA[专题页]]></name>" &vbcrlf
						XMLStr=XMLStr & "    <fsotype>1002</fsotype>" &vbcrlf
					    XMLStr=XMLStr & "  </fsoitem>"&vbcrlf
			 End If
			
					XMLStr=XMLStr &" </item>" &vbcrlf
					Call KS.WriteTOFile(KS.Setting(3) & "config/quickrefresh.xml",xmlstr)  
					response.write "<br/>生成任务初始化完毕，两秒钟后自动开始。"
			 response.write "</div>"
			ks.die "<meta http-equiv=""refresh"" content=""2;url=Refreshquick.asp?action=start"">"

		End Sub
		
		
		Sub Start()
		    call main()
			response.write "<table border='0' width='100%' height='100%' cellspacing='0' cellpadding='0'>"
			response.write "<tr><td width='230'>"
		   response.write "<iframe src='refreshquick.asp?action=startleft' frameborder='0' name='myframeleft' width='100%' height='100%'></iframe>"
           response.write "</td><td>"
		   response.write "<iframe src='refreshquick.asp?action=startright' frameborder='0' name='myframeright' width='100%' height='100%'></iframe>"
		   response.write "</td></tr>"
		   response.write "</table>"
		End Sub
		
		Sub StartLeft()
		    Dim TaskXML,TaskNode,Node,N,TaskUrl,Taskid,Action
			set TaskXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			TaskXML.async = false
			TaskXML.setProperty "ServerHTTPRequest", true 
			TaskXML.load(Server.MapPath(KS.Setting(3)&"Config/quickrefresh.xml"))
			Set TaskNode=TaskXML.DocumentElement.SelectNodes("//item/fsoitem")
			Dim TotalNum:TotalNum=tasknode.length
			
			response.write ("<link rel=""stylesheet"" href=""Admin_Style.css"">")
			response.write "<br/>&nbsp;&nbsp;&nbsp;<strong>共找到 <font color='red'>" &  TotalNum & "</font> 个生成的项目。</strong><br/>"
			For Each Node In TaskXML.DocumentElement.SelectNodes("fsoitem[@isfinish=1]")
			  select case node.selectsinglenode("fsotype").text
			     case 0 
				  	response.write "<li><font color=green><b>√</b></font> 网站首页生成完毕</li>"
                 case 1
				  	response.write "<li><font color=green><b>√</b></font> [" & node.selectsinglenode("name").text & "]栏目页生成完毕</li>"
                 case 2
				  	response.write "<li><font color=green><b>√</b></font> [" & node.selectsinglenode("name").text & "]内容页生成完毕</li>"
			     case 90 
				  	response.write "<li><font color=green><b>√</b></font> 考试频道首页等生成完毕</li>"
			     case 1000 
				  	response.write "<li><font color=green><b>√</b></font> 专题首页生成完毕</li>"
			     case 1001 
				  	response.write "<li><font color=green><b>√</b></font> 专题分类生成完毕</li>"
			     case 1002 
				  	response.write "<li><font color=green><b>√</b></font> 专题页生成完毕</li>"
			  end select
			Next
            For Each Node In TaskXML.DocumentElement.SelectNodes("fsoitem[@isfinish=0]")
			   ChannelID=Node.selectsinglenode("@channelid").text
			   		Node.Attributes.getNamedItem("isfinish").text=1
					TaskXML.Save(Server.MapPath(KS.Setting(3)&"Config/quickrefresh.xml"))
			   select case node.selectsinglenode("fsotype").text
			     case 0
				       response.write "<li style='color:red;font-weight:bold'>正在生成网站首页...</li>"
					   response.write "<script>parent.myframeright.location.href='refreshquick.asp?action=startright&fsotype=100';</script>"
					   response.end
				 case 1
				       response.write "<li style='color:red;font-weight:bold'>正在生成[" & Node.SelectSingleNode("name").text & "]的栏目页...</li>"
					   response.write "<script>parent.myframeright.location.href='refreshquick.asp?action=startright&channelid=" & ChannelID & "&fsotype=1';</script>"
					   response.end
				 case 2
				       response.write "<li style='color:red;font-weight:bold'>正在生成[" & Node.SelectSingleNode("name").text & "]的内容页...</li>"
					   response.write "<script>parent.myframeright.location.href='refreshquick.asp?action=startright&channelid=" & ChannelID & "&fsotype=2';</script>"
					  response.end
				 case 90
				       response.write "<li style='color:red;font-weight:bold'>正在生成考试频道首页等...</li>"
					   response.write "<script>parent.myframeright.location.href='refreshquick.asp?action=startright&fsotype=90';</script>"
					   response.end
				 case 1000
				       response.write "<li style='color:red;font-weight:bold'>正在生成专题首页...</li>"
					   response.write "<script>parent.myframeright.location.href='refreshquick.asp?action=startright&fsotype=1000';</script>"
					   response.end
				 case 1001
				       response.write "<li style='color:red;font-weight:bold'>正在生成专题分类页...</li>"
					   response.write "<script>parent.myframeright.location.href='refreshquick.asp?action=startright&fsotype=1001';</script>"
					   response.end
				 case 1002
				       response.write "<li style='color:red;font-weight:bold'>正在生成专题页...</li>"
					   response.write "<script>parent.myframeright.location.href='refreshquick.asp?action=startright&fsotype=1002';</script>"
					   response.end
			   end select
			Next
		    if TaskXML.DocumentElement.SelectNodes("fsoitem[@isfinish=0]").length=0 then
			     response.write "<li style='font-weight:bold;color:green;'>&nbsp;&nbsp;恭喜，所有任务生成完毕!</li>"
				 KS.Die "<script>setTimeout(""parent.myframeright.location.href='refreshquick.asp?action=startright&fsotype=-1';"",500);</script>"
			end if
		End Sub
		

		
		Sub StartRight()
		 Dim Template
		 Dim FsoType:FsoType=KS.ChkClng(KS.G("FsoType"))
		 Select Case FsoType
		   case 100
				Template = KSRObj.LoadTemplate(KS.Setting(110))
				Template=KSRObj.KSLabelReplaceAll(Template)
				Call KS.WriteTOFile(KS.Setting(3)&KS.Setting(5),Template)
				FsoHtmlList="<div style='text-align:center'>网站首页生成完成，<a href='../../' target='_blank'>点此浏览</a>!两秒后执行下一个发布任务...</div>" 
				Call Main
				KS.Die "<script>setTimeout(""parent.myframeleft.location.href='refreshquick.asp?action=startleft';"",2000);</script>"
		  case 1
		      ItemName=KS.C_S(ChannelID,1)
              RefreshFolder()
			  response.End()
		  case 2
		      RefreshFlag="New"
			  ItemName=KS.C_S(ChannelID,1)
		      RefreshContent()
			  Response.End()
		  case 90
              RefreshSjIndex()
			  response.End()
          case 1000
              RefreshSpeicalIndex()
			  response.End()
          case 1001
              RefreshSpeicalClass()
			  response.End()
          case 1002
              RefreshSpeical()
			  response.End()
		  case -1
				Dim TaskXML,TaskNode,Node,N,TaskUrl,Taskid,Action
				set TaskXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				TaskXML.async = false
				TaskXML.setProperty "ServerHTTPRequest", true 
				TaskXML.load(Server.MapPath(KS.Setting(3)&"Config/quickrefresh.xml"))
				Set TaskNode=TaskXML.DocumentElement.SelectNodes("//item/fsoitem")
				Dim TotalNum:TotalNum=tasknode.length

				FsoHtmlList="<div style='text-align:center'><div style='text-align:center;font-weight:bold;color:green;width:400px;'><img src='../images/succeed.gif' align='left'><br/>恭喜，所有生成任务执行完毕,共执行了 <font color=red>" & TotalNum & "</font> 个生成项目！</div><input type='button' value='返回一键发布首页' onclick=""parent.location.href='refreshquick.asp';"" class='button'/></div>" 
				Call Main
		  case else
				FsoHtmlList="<div style='text-align:center'>正在执行生成任务...</div>" 
				Call Main
		 End Select
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
		  .echo ("<body oncontextmenu=""return false;"" scroll=no>")
		if action<>"startright" Then
		 .echo "<div class='topdashed sort'> 一键快速生成HTML管理</div>"
		End If
		
		if action<>"start" then
			.echo "<br><br><br>"
			if action<>"" and KS.G("fsotype")<>"100" and KS.G("fsotype")<>"-1" and action<>"initial" then
				.echo "<table id=""BarShowArea"" width=""400"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1"">"
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
			end if
			
			 .echo ("<table width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
			 .echo (" <tr>")
			 .echo ("   <td height=""50"">")
			 .echo ("     <div style='text-align:center'> ")
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
		 
		end if
		 .echo ("</body>")
		 .echo ("</html>")
		 End With
		End Sub
		
		'生成栏目的处理过程
		Sub RefreshFolder()
		 If ChannelID=9 Then Call RefreshSJFolder():  Exit Sub
		With KS
		Dim FolderID, R_Sql, RefreshTotalNum, R_RS, NewsTotalNum, NewsNo		  
		 If NewsNo = "" Then NewsNo = 0
		 R_Sql = "Select * from KS_Class where ChannelID=" & ChannelID &" and ClassType<>2 and DelTF=0 Order By FolderOrder ASC"
		Call Main
		If R_Sql <> "" Then
			Set R_RS = Server.CreateObject("ADODB.RecordSet")
			R_RS.Open R_Sql, Conn, 1, 1
			If R_RS.EOF Then
				.echo "<script>img2.width=""0"";" & vbCrLf
				.echo "txt2.innerHTML=""没有可生成的" & ItemName & "栏目！两秒后执行下一个发布任务..."";" & vbCrLf
				.echo "txt3.innerHTML="""";" & vbCrLf
				.echo "txt4.innerHTML="""";" & vbCrLf
				.echo "document.all.BarShowArea.style.display='none';" & vbCrLf
				.echo "</script>" & vbCrLf
				R_RS.Close:Set R_RS = Nothing
				.Die "<script>setTimeout(""parent.myframeleft.location.href='refreshquick.asp?action=startleft';"",2000);</script>"
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
				.echo "txt3.innerHTML=""总共生成了 <font color=red><b>" & NewsTotalNum & "</b></font> 个" & ItemName & "栏目,总费时:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font> 秒<br/>两秒后执行下一个发布任务..."";" & vbCrLf
				.echo "img2.title=""(" & NewsNo & ")"";</script>" & vbCrLf
				R_RS.Close:Set R_RS = Nothing
				.die "<script>setTimeout(""parent.myframeleft.location.href='refreshquick.asp?action=startleft';"",2000);</script>"
			End If
		End If
		End With
		End Sub
		
		'生成内容页的处理过程
		Sub RefreshContent()
		Dim AlreadyRefreshByID, NowNum, R_Sql, R_RS, TotalNum,ID,UpdateSql
		Dim StartDate, EndDate, FolderID, RefreshTotalNum,StartID,EndID
		RefreshTotalNum = Request.QueryString("RefreshTotalNum")
		NowNum = Request.QueryString("NowNum") '正在刷新第几篇文章
		Table=KS.C_S(ChannelID,2)
		ItemName=KS.C_S(ChannelID,3)
		If ChannelID=9 Then Call RefreshSJContent():  Exit Sub
		
		R_Sql=" Where Verific=1"
		With KS
		If NowNum = "" Then NowNum = 0
		  Select Case RefreshFlag
			Case "New"
			  TotalNum = KS.ChkClng(KS.ReadSetting(1))
			  If TotalNum<>0 Then
				  If TotalNum >conn.execute("select count(id) from "& Table & R_SQL )(0) Then TotalNum = conn.execute("select count(id) from "& Table & R_SQL)(0)
				  RefreshTotalNum = TotalNum
				  If TotalNum=0 Then TotalNum=1
				   UpdateSql="Update "& Table & " Set RefreshTF=1 Where ID in(Select Top " & TotalNum & " ID from " & Table  & R_SQL& " Order By ID Desc)"
				  R_Sql="Select Top " & TotalNum & " * from " & Table  & R_SQL& " Order By ID Desc"
		      Else
				  RefreshTotalNum=conn.execute("select count(id) from " & Table  & R_Sql)(0)
				  UpdateSql="Update "& Table & " Set RefreshTF=1 "& R_Sql 
				  R_Sql = "Select * from " & Table  & R_Sql & " order by ID desc"
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
				.echo "txt2.innerHTML=""没有可生成的内容页！两秒后执行下一个发布任务..."";" & vbCrLf
				.echo "txt3.innerHTML="""";" & vbCrLf
				.echo "txt4.innerHTML="""";" & vbCrLf
				.echo "document.all.BarShowArea.style.display='none';" & vbCrLf
				.echo "</script>" & vbCrLf
				Response.Flush
				R_RS.Close:Set R_RS=Nothing
				.die "<script>setTimeout(""parent.myframeleft.location.href='refreshquick.asp?action=startleft';"",2000);</script>"
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
						 .echo "fsohtml.innerHTML='<div style=""text-align:cdenter""><div style=""margin:10px;height:80px;padding:8px;border:1px dashed #cccccc;text-align:left;""><img src=""../images/succeed.gif"" align=""left""><br>&nbsp;&nbsp;&nbsp;&nbsp;<b>温馨提示：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;以免过度占用服务器资源，系统暂停" & PauseTime & "秒后继续<img src=""../../images/default/wait.gif""><br>&nbsp;&nbsp;&nbsp;&nbsp;如果" & PauseTime & "秒后没有继续，请点此<a href=""RefreshQuick.asp?fsotype=2&CurrNowNum=" & NowNum+1 & "&ChannelID=" & ChannelID & "&RefreshFlag=Pause&Action=startright&StartRefreshTime=" & StartRefreshTime & "&UpdateSql=" & server.URLEncode(UpdateSql) & "&R_Sql=" & Server.UrlEncode(R_Sql) & "&RefreshTotalNum=" & RefreshTotalNum & """><font color=red>继续</font></a>或点此<a href=""refreshhtml.asp?Action=ref&channelid=" & channelid & """><font color=red>停止</font></a>!</div></div>';" & vbCrLf
						 .echo "</script>" &vbcrlf
						 .die "<meta http-equiv=""refresh"" content=""" & PauseTime & ";url=RefreshQuick.asp?fsotype=2&CurrNowNum=" & NowNum+1 & "&ChannelID=" & ChannelID & "&RefreshFlag=Pause&Action=startright&StartRefreshTime=" & StartRefreshTime & "&UpdateSql=" & server.URLEncode(UpdateSql) & "&R_Sql=" & Server.UrlEncode(R_Sql) & "&RefreshTotalNum=" & RefreshTotalNum & """>"
					End If
			   End If
			Next
			    If UpdateSql<>"" Then Conn.Execute(UpdateSql)
				.echo "<script>"
				.echo "fsohtml.innerHTML='';" & vbCrLf
				.echo "img2.width=400;" & vbCrLf
				.echo "txt2.innerHTML=""生成内容页结束！100"";" & vbCrLf
				.echo "txt3.innerHTML=""总共生成了 <font color=red><b>" & RefreshTotalNum & "</b></font> 条,总费时:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font> 秒<br/>两秒后执行下一个发布任务..."";" & vbCrLf
				.echo "img2.title=""(" & NowNum & ")"";</script>" & vbCrLf
				
				.die "<script>setTimeout(""parent.myframeleft.location.href='refreshquick.asp?action=startleft';"",2000);</script>"

			End If
		Else
				.echo "<script>img2.width=""0"";" & vbCrLf
				.echo "txt2.innerHTML=""没有可生成的内容页！<br><br>"";" & vbCrLf
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
				.echo "fsohtml.innerHTML='<div style=""margin:10px;height:80px;padding:8px;border:1px dashed #cccccc;text-align:left;"">" & FsoHtmlList & "</div>';" & vbCrLf
			    else
				.echo "fsohtml.innerHTML='" & FsoHtmlList & "';" & vbCrLf
				end if
				.echo "img2.width=" & Fix((NowNum / TotalNum) * 400) & ";" & vbCrLf
				.echo "txt2.innerHTML=""生成进度:" & FormatNumber(NowNum / TotalNum * 100, 2, -1) & """;" & vbCrLf
				.echo "txt3.innerHTML=""总共需要生成 <font color=red><b>" & TotalNum & "</b></font> " & itemname & ",<font color=red><b>在此过程中请勿刷新此页面！！！</b></font> 系统正在生成第 <font color=red><b>" & NowNum & "</b></font> " & itemname & """;" & vbCrLf
				.echo "img2.title=""(" & NowNum & ")"";</script>" & vbCrLf
				Response.Flush
		  End With
		End Sub
		
		Sub RefreshSjIndex()
		    response.Redirect("../mnkc/refreshindex.asp?from=quick")
		End Sub
		
		Sub RefreshSJFolder()
		    response.Redirect("../mnkc/refreshallcalss.asp?type=all&from=quick")
		End Sub
		
		Sub RefreshSJContent()
		     response.Redirect("../mnkc/refreshsj.asp?action=refresh&type=all&from=quick")
		End Sub
		
		Sub RefreshSpeicalIndex()
		     response.Redirect("../include/RefreshSpecialSave.asp?types=Index&from=quick")
		End Sub
		
		Sub RefreshSpeicalClass()
		     response.Redirect("../include/RefreshSpecialSave.asp?types=ChannelSpecial&RefreshFlag=All&from=quick")
		End Sub
		Sub RefreshSpeical()
		     response.Redirect("../include/RefreshSpecialSave.asp?types=Special&RefreshFlag=All&from=quick")
		End Sub
End Class
%> 