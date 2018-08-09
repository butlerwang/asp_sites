<%
'================================下载模型开始================================
		   case "getdowntitle"   echo GetNodeText("title") & " " & GetNodeText("downversion")
		   case "getdownaction"  echo "【<A href=""" & DomainStr & "plus/Comment.asp?ChannelID=" & ModelID & "&InfoID=" & ItemID & """ target=""_blank"">我来评论</A>】【<A href=""" & DomainStr & "User/User_Favorite.asp?Action=Add&ChannelID=" & ModelID & "&InfoID=" & ItemID & """ target=""_blank"">我要收藏</A>】【<A href=""javascript:window.close();"">关闭窗口</A>】"
		   case "getdownurl"   echo KS.GetItemURL(ModelID,GetNodeText("tid"),ItemID,GetNodeText("fname"))
		   case "getdownsystem"   echo GetNodeText("downpt")
		   case "getdownsize" echo GetNodeText("downsize")
		   case "getdowntype" echo GetNodeText("downlb")
		   case "getdownlanguage" echo GetNodeText("downyy")
		   case "getdownpower" echo GetNodeText("downsq")
		   case "getdownpoint" echo GetNodeText("readpoint")
		   case "getdowndecpass" echo GetNodeText("jymm")
		   case "getdownintro" echo KS.ReplaceInnerLink(GetNodeText("downcontent"))
		   case "getdownaddress" 
		     Dim UrlArr, I,N,TotalNum, AUrl
			 UrlArr = Split(GetNodeText("downurls"), "|||")
			 TotalNum = UBound(UrlArr)
			 For I = 0 To TotalNum
			    N=N+1: AUrl = Split(UrlArr(I), "|")
				If AUrl(0)=0 Then
				 echoln "<img src="""& DomainStr & "Images/Default/down.gif"" border=""0"" alt="""" align=""absmiddle"" /><a href=""" & DomainStr & "item/downLoad.asp?m=" & ModelID & "&id=" & ItemID & "&downid=" & N & """ target=""_blank"">" & AUrl(1) & "</a>"      
				 If I<>TotalNum Then echoln "<br/>"
				Else
				  Dim RS_S:Set RS_S=Conn.Execute("Select DownloadName,IsDisp,DownloadPath,DownID,SelFont From KS_DownSer Where ParentID=" & AUrl(0))
				  If RS_S.Eof Then
				    If TotalNum=0 Then echo "<li>暂不提供下载地址</li>"
				  Else
				     DO While Not RS_S.Eof
					  IF RS_S(1)=1 Then
				      echoln "<img src="""& domainstr & "Images/Default/down.gif"" border=""0"" align=""absmiddle""><a href=""" & RS_S(2) & Aurl(2) & """ " & RS_S(4)&" target=""_blank"">" & RS_S(0) & "</a>"          
					  Else
				      echoln "<img src="""& domainstr & "Images/Default/down.gif"" border=""0"" align=""absmiddle""><a href=""" & DomainStr & "item/DownLoad.asp?m=" & ModelID & "&id=" & ItemID & "&DownID=" & N & "&Sid=" & RS_S(3) & """ " & RS_S(4)&" target=""_blank"">" & RS_S(0) & "</a>"        
					  End If
					  RS_S.MoveNext
					  IF Not RS_S.Eof Or I<>TotalNum Then echoln "<br/>" 
					 Loop
				  End If
				  RS_S.Close:Set RS_S=Nothing
				End If
			 Next
		   case "getdownlink"
				If Not (LCase(Node.SelectSingleNode("@ysdz").text) = "http://" Or Node.SelectSingleNode("@ysdz").text = "") Then  echo "<a href=""" & Node.SelectSingleNode("@ysdz").text & """ target=""_blank""><u>作者或开发商主页</u></a>"
				If Not (LCase(Node.SelectSingleNode("@zcdz").text) = "http://" Or Node.SelectSingleNode("@zcdz").text = "") Then  echo "&nbsp;&nbsp;<a href=""" & Node.SelectSingleNode("@zcdz").text & """ target=""_blank""><u>注册地址</u></a>"
		   case "getdownysdz"
				If LCase(Node.SelectSingleNode("@ysdz").text) = "http://" Or Node.SelectSingleNode("@ysdz").text = "" Then
				   echo "无"
				Else
				   echo "<a href=""" & Node.SelectSingleNode("@ysdz").text & """ target=""_blank"">" & Node.SelectSingleNode("@ysdz").text & "</a>"
				End If
		   case "getdownzcdz"
				If LCase(Node.SelectSingleNode("@zcdz").text) = "http://" Or Node.SelectSingleNode("@zcdz").text = "" Then
				   echo "无"
				Else
				   echo "<a href=""" & Node.SelectSingleNode("@zcdz").text & """ target=""_blank"">" & Node.SelectSingleNode("@zcdz").text & "</a>"
				End If
		   case "getdownproperty"
		     If GetNodeText("recommend") = "1" Then Echo "<span title=""推荐"" style=""cursor:default;color:green"">荐</span> "
			 If GetNodeText("popular") = "1" Then  echo "<span title=""热门"" style=""cursor:default;color:red"">热</span> "
			 If GetNodeText("strip")="1" Then echo "<span title=""今日头条"" style=""cursor:default;color:#0000ff"">头</span> "
			 If GetNodeText("rolls") = "1" Then echo "<span title=""滚动"" style=""cursor:default;color:#F709F7"">滚</span> "
			 If GetNodeText("slide") = "1" Then echo "<span title=""幻灯片"" style=""cursor:default;color:black"">幻</span>"
'================================下载模型开始================================
%>