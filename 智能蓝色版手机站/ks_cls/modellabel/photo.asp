<%
'================================图片模型开始================================
		   case "getpicturename" echo GetNodeText("title")
		   case "showpictures" echo PageContent
		   case "getpicnums" echo ubound(split(Node.SelectSingleNode("@picurls").text&"","|||"))+1
		   case "getpictureintro" echo KS.ReplaceInnerLink(GetNodeText("picturecontent"))
		   case "getpictureurl"   echo KS.GetItemURL(ModelID,GetNodeText("tid"),ItemID,GetNodeText("fname"))
		   case "getpictureinput"   echo "<a href=""" & DomainStr & "Space/?" & GetNodeText("inputer") &""" target=""_blank"">" & GetNodeText("inputer") & "</a>"
		   case "getpicturesrc","getphotourl"    
		      Dim Purl:Purl=GetNodeText("photourl")
			  If KS.IsNul(Purl) Then echo DomainStr &"images/nopic.gif" Else Echo purl
		   case "getpictureproperty"
		     If GetNodeText("recommend") = "1" Then Echo "<span title=""推荐"" style=""cursor:default;color:green"">荐</span> "
			 If GetNodeText("popular") = "1" Then  echo "<span title=""热门"" style=""cursor:default;color:red"">热</span> "
			 If GetNodeText("strip")="1" Then echo "<span title=""今日头条"" style=""cursor:default;color:#0000ff"">头</span> "
			 If GetNodeText("rolls") = "1" Then echo "<span title=""滚动"" style=""cursor:default;color:#F709F7"">滚</span> "
			 If GetNodeText("slide") = "1" Then echo "<span title=""幻灯片"" style=""cursor:default;color:black"">幻</span>"
		   case "getpicturevotescore" echo "<script type=""text/Javascript"" src=""" & DomainStr & "Item/GetVote.asp?m=" & ModelID & "&ID=" & ItemID & """></script>"
		   case "getpicturevote" echo "<a href=""" & DomainStr & "Item/Vote.asp?m=" & ModelID & "&ID=" & ItemID & """>投它一票</a>"
'================================图片模型结束================================
%>