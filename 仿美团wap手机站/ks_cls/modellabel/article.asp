<%
 '================================文章模型开始================================
	  case "getarticlepagelist"    '分页导航
		 dim URL,PN,PageTitle:PageTitle=GetNodeText("pagetitle")
		 if Not KS.IsNul(PageTitle) Then
		   echo "<div class=""pagenavigation"">" &vbcrlf
		   echo " <h2>分页导航</h2>" &vbcrlf
		   echo " <ul>" &vbcrlf
		    PageTitle=split(PageTitle,"§")
			Dim ShowUrl:ShowUrl=KS.LoadInfoUrl(ModelID,Tid,"",ItemId)
			For PN=0 to Ubound(split(Node.SelectSingleNode("@articlecontent").text,"[NextPage]"))	
			
			  If KS.C_S(ModelID,7)<>1 and KS.C_S(ModelID,7)<>2 Then '动态
			      If KS.C_S(ModelID,48)=0 Then
				  URL="?m=" & ModelID & "&d="& ItemID & "&p="&(PN+1)
				 ElseIf KS.C_S(ModelID,48)=2 Then
				  URL=KS.Setting(3) & GCls.StaticPreContent & "-" & ItemID & "-" & ModelID & "-" & (PN+1) & GCls.StaticExtension
				 Else
				  URL=KS.Setting(3) & "?" & GCls.StaticPreContent & "-" & ItemID & "-" & ModelID & "-" & (PN+1) & GCls.StaticExtension
				 End If
			  else  '生成静态
			          dim sFname:sFname = Trim(Node.SelectSingleNode("@fname").text)
					  dim FExt:FExt   = Mid(sFname, InStrRev(sFname, ".")) '分离出扩展名
					  dim FName:Fname = Replace(sFname, FExt, "")  '分离出文件名 如 2005/9-10/1254ddd
					If PN=0 Then
					  URL = ShowUrl & sFname
					Else
					  URL =ShowUrl & Fname & "_" & (PN+1) & FExt
				    End If
			  end if	
			 echo "<li><a href=""" & url & """>第" & (PN+1) & "页：" & PageTitle(PN) & "</a></li>" &vbcrlf
			Next
		   echo " </ul>" &vbcrlf
		   echo "</div>" &vbcrlf
		 End If
		 case "getarticletitle" echo LFCls.ReplaceDBNull(GetNodeText("fulltitle"),GetNodeText("title"))
		case "getarticlesize"  
				  echoln "<script type=""text/javascript"" language=""javascript"">"
				  echoln  "function ContentSize(size)"
				  echoln  "{document.getElementById('MyContent').style.fontSize=size+'px';}" 
				  echoln  "</script>"
				  echoln "【字体：<a href=""javascript:ContentSize(16)"">大</a> <a href=""javascript:ContentSize(14)"">中</a> <a href=""javascript:ContentSize(12)"">小</a>】"
	    case "getarticlecontent"
			      echoln ReplaceAd(FormatImgLink(KS.ReplaceInnerLink(Replace(Replace(Replace(Replace(PageContent,"{$","{§"),"{LB","{#LB"),"{SQL","{#SQL"),"{=","{#=")),NextUrl,TotalPage),GetNodeText("tid"))
		case "getarticleaction"
			      echo "【<A href=""" & DomainStr & "plus/Comment.asp?ChannelID=" & ModelID & "&InfoID=" & ItemID & """ target=""_blank"">发表评论</A>】【<A href=""" & DomainStr & "item/SendMail.asp?m="&ModelID &"&ID=" & ItemID & """ target=""_blank"">告诉好友</A>】【<A href=""" & DomainStr & "item/Print.asp?m=" & ModelID &"&ID=" & ItemID & """ target=""_blank"">打印此文</A>】【<A href=""" & DomainStr & "User/User_Favorite.asp?Action=Add&ChannelID=" & ModelID & "&InfoID=" & ItemID & """ target=""_blank"">收藏此文</A>】【<A href=""javascript:window.close();"">关闭窗口</A>】"
		case "getarticleintro"
			   Dim myintro:myintro=KS.LoseHtml(GetNodeText("intro"))
			  if instr(myintro,"[UploadFiles]")<>0 then
			   myintro=replace(myintro,KS.CutFixContent(myintro, "[UploadFiles]", "[/UploadFiles]", 1),"")
			  end if
			  echo myintro
		case "getarticleshorttitle" echo GetNodeText("title")
		case "getarticleurl"   echo KS.GetItemURL(ModelID,GetNodeText("tid"),ItemID,GetNodeText("fname"))
		case "getarticlekeyword","getpicturekeyword","getdownkeyword" echo Replace(GetNodeText("keywords"), "|", ",")
		case "getarticleauthor","getpictureauthor","getdownauthor" echo LFCls.ReplaceDBNull(GetNodeText("author"),"佚名")
		case "getarticleinput","getpictureinput","getdowninput" echo "<a href=""" & DomainStr & "Space/?" & GetNodeText("inputer") &""" target=""_blank"">" & GetNodeText("inputer") & "</a>"
		case "getarticleorigin","getpictureorigin","getdownorigin" echo KS.GetOrigin(LFCls.ReplaceDBNull(GetNodeText("origin"),"本站原创"))
		case "getarticleproperty" 
			  If GetNodeText("recommend") = "1" Then echo "<span title=""推荐"" style=""cursor:default;color:green"">荐</span> "
			  If GetNodeText("popular") = "1" Then  echo "<span title=""热门"" style=""cursor:default;color:red"">热</span> "
			  If GetNodeText("strip")="1" Then echo "<span title=""今日头条"" style=""cursor:default;color:#0000ff"">头</span> "
			  If GetNodeText("rolls") = "1" Then echo "<span title=""滚动"" style=""cursor:default;color:#F709F7"">滚</span> "
			  If GetNodeText("slide") = "1" Then echo "<span title=""幻灯片"" style=""cursor:default;color:black"">幻</span>"
		case "getarticledate","getadddate","getpicturedate","getdowndate" echo LFCls.Get_Date_Field(GetNodeText("adddate"), "YYYY年MM月DD日")
		case "getmodifydate" echo LFCls.Get_Date_Field(GetNodeText("modifydate"), "YYYY年MM月DD日")
        case "getprevarticle","getprevpicture","getprevdown","getprevproduct" echo LFCls.ReplacePrevNext(ModelID,ItemID, GetNodeText("tid"), "<")
        case "getnextarticle","getnextpicture","getnextdown","getnextproduct" echo LFCls.ReplacePrevNext(ModelID,ItemID, GetNodeText("tid"), ">")
		case "getpictureaction" echo "【<A href=""" & DomainStr & "plus/Comment.asp?ChannelID=" & ModelID & "&InfoID=" & ItemID & """ target=""_blank"">我来评论</A>】【<A href=""" & DomainStr & "User/User_Favorite.asp?Action=Add&ChannelID=" & ModelID & "&InfoID=" & ItemID & """ target=""_blank"">我要收藏</A>】【<A href=""javascript:window.close();"">关闭窗口</A>】"
        case "getseotitle" if Not KS.IsNul(GetNodeText("seotitle")) Then echo GetNodeText("seotitle") Else Echo GetNodeText("title")
        case "getseokeywords" if Not KS.IsNul(GetNodeText("seokeyword")) Then echo GetNodeText("seokeyword") Else Echo GetNodeText("keywords")
        case "getseodescription" 
		   if Not KS.IsNul(GetNodeText("seodescript")) Then 
		    echo GetNodeText("seodescript") 
		   Else 
		    Select Case KS.C_S(ModelID,6)
			 case 1 Echo KS.LoseHTML(GetNodeText("intro"))
			 case 2 Echo KS.Gottopic(KS.LoseHTML(GetNodeText("picturecontent")),150)
			 case 3 Echo KS.Gottopic(KS.LoseHTML(GetNodeText("downcontent")),150)
			 case 4 Echo KS.Gottopic(KS.LoseHTML(GetNodeText("flashcontent")),150)
			 case 5 Echo KS.Gottopic(KS.LoseHTML(GetNodeText("prointro")),150)
			 case 7 Echo KS.Gottopic(KS.LoseHTML(GetNodeText("moviecontent")),150)
			 case 8 Echo KS.Gottopic(KS.LoseHTML(GetNodeText("gqcontent")),150)
			End Select
		   End If
 '================================文章模型结束=================================
%>