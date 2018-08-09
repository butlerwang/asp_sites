<!--#include file="Kesion.Label.FunctionCls.asp"-->
<!--#include file="Kesion.Label.LocationCls.asp"-->
<!--#include file="Kesion.Label.SearchCls.asp"-->
<!--#include file="Kesion.Label.SQLCls.asp"-->
<!--#include file="Kesion.Label.JSCls.asp"-->
<!--#include file="ModelLabel/JobCls.asp"-->
<!--#include file="ModelLabel/mnkcCls.asp"-->
<!--#include file="Kesion.Label.BaseFunCls.asp"-->
<!--#include file="Kesion.FsoVarCls.asp"-->
<%

Class Refresh
		Private KS,KSLabel,DomainStr  
		public Templates,ModelID,Tid,ItemID          rem  ModelID 模型ID ItemID 文档ID      
		public Node,PageContent,NextUrl,PrevUrl,TotalPage    rem  Node 节点对象,PageContent 分页内容
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSLabel =New RefreshFunction
		  DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSLabel=Nothing
		End Sub
		%>
		<!--#include file="UbbFunction.asp"-->
		<%
		Sub Echo(sStr)
			Templates    = Templates & sStr 
		End Sub
		Sub EchoLn(sStr)
		    Templates    = Templates & sStr & VbNewLine
		End Sub
		
		public Sub Scan(ByVal sTemplate)
		    If Fcls.RefreshType="Content" Then Call ReplaceHits(sTemplate,ModelID,ItemId)  '内容页先替换点击数标签
			Dim iPosLast, iPosCur
			iPosLast = 1
			Dim Tags,Key,yllen
          do  while (true)
                iPosCur = findTags(sTemplate, tags, key, iPosLast,yllen)
                if (iPosCur <>0) then
                    Echo mid(sTemplate,iPosLast, iPosCur - iPosLast)
                    select case (tags)
                        case "{$"
                            Parse sTemplate, key
                        case "{="
                            ParseEqual sTemplate, key
                    end select
                    iPosLast = yllen + 1
                else
                    Echo    Mid(sTemplate, iPosLast)
                    exit do
				end if
          loop
		End Sub 
		
		Function FindTags(sTemplate, byref tags, ByRef key, iPosLast, ByRef yllen)
				dim a:a = array("{$", "{=")   '定义标签开始标记
				dim i, cur, posCur
				cur=0
				for i=0 to ubound(a)
						posCur=instr(iPosLast,sTemplate,a(i))
						if (posCur<>0 and (cur=0 or posCur<cur)) then
							cur=posCur
							tags=a(i)
							yllen=instr(posCur,sTemplate,"}")
							key=mid(sTemplate,posCur+len(a(i)),yllen-posCur-len(a(i)))
							if (cur <= 0) then exit for   '说明已经是最小了,可以退出
						end if
				next
				FindTags=cur
		End Function		
				
		
		Function GetNodeText(NodeName)
		 Dim N
		 If IsObject(Node) Then
		  set N=node.SelectSingleNode("@" & NodeName)
		  If Not N is Nothing Then GetNodeText=N.text
		 End If
		End Function
		
		Function Parse(sTemplate, sTemp)
		    dim MyNode
			ParseChannelLabel sTemp   rem 解释频道标签
			ParseRssLabel     sTemp   rem 解释RSS标签
			select case Lcase(sTemp)
			 '================================网站通用参数开始===========================
			 case "getsiteurl"      echo Domainstr
			 case "getsitename"     echo KS.Setting(0)
			 case "getsitetitle"    echo KS.Setting(1)
			 case "getsitelogo"     echo "<img src=""" & KS.Setting(4) & """ border=""0"" align=""absmiddle"" alt=""logo"" />"
			 case "getsitecountall" echo GetSiteCountAll()
			 case "getsiteonline"   echo "<script src=""" & DomainStr & "plus/online.asp?Referer=""+escape(document.referrer) type=""text/javascript""></script>"
			 case "getpoplogin"
			      Dim LoginStrxml:LoginStrXml=LFCls.GetXMLByNoCache("userlogin","/logintemplate/label","[@name='popup']")
				  LoginStrxml=Replace(Replace(LoginStrxml,"{$GetSiteUrl}",DomainStr),"{$GetInstallDir}",KS.Setting(3))
				  echo LoginStrxml
			 case "getuserloginbyscript" echo "<script src=""" & domainstr & "user/userlogin.asp?action=script"" type=""text/javascript""></script>"
             case "gettopuserlogin" 
			      LoginStrXml=LFCls.GetConfigFromXML("userlogin","/logintemplate/label","top")
				  LoginStrxml=Replace(Replace(LoginStrxml,"{$GetSiteUrl}",DomainStr),"{$GetInstallDir}",KS.Setting(3))
				  echo LoginStrxml
             case "getuserlogin"    echo "<iframe width=""180"" height=""122"" id=""loginframe"" name=""loginframe"" src=""" & KS.Setting(3) & "user/userlogin.asp"" frameborder=""0"" scrolling=""no"" allowtransparency=""true""></iframe>"
             case "getspecial"
			      Dim SpecialIndexUrl,SpecialDir:SpecialDir = KS.Setting(95)
				  If Split(KS.Setting(5),".")(1)<>"asp" Then SpecialIndexUrl=DomainStr & SpecialDir Else SpecialIndexUrl=DomainStr & "item/SpecialIndex.asp"
				  echo "<a href=""" & SpecialIndexUrl & """ target=""_blank"">专题首页</a>"
             case "getfriendlink"   echo "<a href=""" & DomainStr & "plus/Link/"" target=""_blank"">友情链接</a>"
			 case "getinstalldir"   echo KS.Setting(3)
			 case "getmanagelogin"  echo "<a href=""" & DomainStr & KS.Setting(89) & "Login.asp"" target=""_blank"">管理登录</a>"
			 case "getcopyright"    echo KS.Setting(18)
			 case "getmetakeyword"  echo KS.Setting(19)
			 case "getmetadescript" echo KS.Setting(20)
			 case "getwebmaster"    echo "<a href=""mailto:" & KS.Setting(11) & """>" & KS.Setting(10) & "</a>"
			 case "getwebmasteremail" echo KS.Setting(11)
			 case "getsiteurl"         echo DomainStr
			 case "getclubinstalldir"     echo KS.Setting(66)
			 case "gettemplatedir"     echo KS.Setting(90)
			 case "getcssdir"     echo KS.Setting(178)
			 case "gettopadlist"  echo KS.GetClubTopAdList
			 case "todaygroupbuylink"  If KS.ChkClng(KS.Setting(179))=1 Then Echo DomainStr & "groupbuy/" Else Echo DomainStr & "shop/groupbuy.asp"
			case "historygroupbuylink" If KS.ChkClng(KS.Setting(179))=1 Then Echo DomainStr & "groupbuy/history/" Else Echo DomainStr & "shop/groupbuy.asp?flag=history"
			 '================================网站通用参数结束===========================
             '====================百度电子地图开始========================
			 case "mapkey" echo KS.Setting(175)
			 case "mapcenterpoint" 
			   Dim MapMarker,MarkerArr
			   MapMarker=GetNodeText("mapmarker")
			   if Not KS.IsNul(MapMarker) Then
			     MarkerArr=Split(MapMarker,"|")
				 echo MarkerArr(0)
			   Else
			    echo KS.Setting(176)
			   End If
			 case "showmarkerlist"
			   MapMarker=GetNodeText("mapmarker")
			   if Not KS.IsNul(MapMarker) Then
			     MarkerArr=Split(MapMarker,"|")
				   For i=0 to Ubound(MarkerArr)
				     echo "point = new BMap.Point(" & MarkerArr(i) & "); " & vbcrlf
				     echo "addMarker(point, " & i & ");" &vbcrlf
				   Next
			   end if
			 '====================地图结束=====================================			 
			 
			 case "channelid"     echo ModelID
			 case "infoid"        echo ItemID
             case "itemname"      echo KS.C_S(ModelID,3)
			 case "itemunit"      echo KS.C_S(ModelID,4)
			 case "getusername"   echo GetNodeText("inputer")
		     case "getrank"       echo Replace(GetNodeText("rank"),"★","<img src=""" & DomainStr & "Images/Star.gif"" border=""0"">")
		     case "getdate"       echo GetNodeText("adddate")
			 case "getkeytags" echo ReplaceKeyTags(GetNodeText("keywords"))
			 case "getshowcomment" 
			   If GetNodeText("comment")="1" Then  
			    echo "<script src=""" & DomainStr & "ks_inc/Comment.page.js"" type=""text/javascript""></script><script src=""" & DomainStr & "ks_inc/kesion.box.js"" type=""text/javascript""></script><script type=""text/javascript"" defer>" 
				if Fcls.CallFrom3G="true" then echo "var from3g=1;" else echo "var from3g=0;"
				echo "Page(1," & ModelID & ",'" & ItemID & "','Show','"& DomainStr & "');</script><div id=""c_" & ItemID & """></div><div id=""p_" & ItemID & """ align=""right""></div>"
			   end if
			 case "getwritecomment" 
			   If GetNodeText("comment")="1" Then 
			     echo "<script tyle=""text/Javascript"" src=""" & DomainStr & "plus/Comment.asp?"
				 if Fcls.CallFrom3G="true" then echo "from3g=1&"
				 echo "Action=Write&PostId=" & GetNodeText("postid") &"&ChannelID=" & ModelID & "&InfoID=" & ItemID & """></script>"
			   End If
           case "getprevurl" echo LFCls.GetPrevNextURL(ModelID,ItemID, GetNodeText("tid"), "<","")
           case "getnexturl" echo LFCls.GetPrevNextURL(ModelID,ItemID, GetNodeText("tid"), ">","")
           %>
		   <!--#include file="modellabel/article.asp"-->
		   <!--#include file="modellabel/photo.asp"-->
		   <!--#include file="modellabel/download.asp"-->
		   <!--#include file="modellabel/flash.asp"-->
		   <!--#include file="modellabel/shop.asp"-->
		   <!--#include file="modellabel/movie.asp"-->
		   <!--#include file="modellabel/supply.asp"-->
		   <%		   
		   case else
		       echo ShCls.run(sTemp)
		      If lcase(left(sTemp,3))="ks_" Then '输出自定义字段
			   if isnumeric(GetNodeText(Lcase(sTemp))) then
			    if GetNodeText(Lcase(sTemp))<1 then  '小于1前面加输出0
				 echo "0" & GetNodeText(Lcase(sTemp))     
				else
				 echo GetNodeText(Lcase(sTemp))    
				end if
			   else
			    echo GetNodeText(Lcase(sTemp))    
			   end if
			  ElseIf lcase(left(sTemp,3))="fl_" Then
			   echo GetNodeText(Lcase(right(sTemp,len(sTemp)-3)))     '输出任意字段
			  elseIf left(lcase(sTemp),3)="js_" then
			   Call JsCls.Run(sTemp,Templates)
			  End If
		 end select
			'Parse    = iPosBegin
			Set MyNode=Nothing
		End Function 
		
		'解释等号标签
		Function ParseEqual(sTemplate, sTemp)
			Dim MyNode,TagName,TagParam,Param,PosTag,I
			PosTag       = InStr(sTemp,"(")
			If PosTag>0 Then
				TagName      = Mid(sTemp,1,PosTag-1)
				TagParam     = Replace(Replace(sTemp,")",""),TagName&"(","")
				'response.write (sTemp & "=" & tagParam)
				'response.end
				Param        = Split(TagParam,",")
				select case Lcase(TagName)
				 case "getfieldvalue"
				   dim G_fieldname:G_fieldname = Param(2)
				   dim G_table:G_table         = KS.C_S(KS.ChkClng(Param(0)),2)
				   dim G_id:G_id               = KS.ChkClng(Param(1))
				   dim rsobj:set rsobj=server.CreateObject("adodb.recordset")
				   rsobj.open "select top 1 " & G_fieldname &" from " & G_table & " where id=" & G_id,conn,1,1
				   if not rsobj.eof then
				      echo rsobj(0)
				   end if
				   rsobj.close:Set rsobj=nothing
				 case "getlogo" echo "<img src=""" & KS.Setting(4) & """ border=""0"" width=""" & Param(0) & """ height=""" & Param(1) & """ align=""absmiddle"" alt=""logo"" />"
				 case "getadvertise" echo "<script src=""" & DomainStr & KS.Setting(93) & ""& TagParam & ".js"" type=""text/javascript""></script>"
				 case "gettopuser" GetTopUser Param(0),Param(1)
				 case "getvote" echo GetVote(TagParam)
				 case "getpk" echo GetPK(TagParam)
				 case "getclassurl" echo KS.GetFolderPath(Param(0))
				 case "gettags" echo GetTags(Param(0),Param(1))
				 case "getuserdynamic" GetUserDynamic TagParam
				 
				 case "getphoto" echo "<div align=""center""><img src=""" & LFCls.ReplaceDBNull(GetNodeText("photourl"), DomainStr & "images/nopic.gif") & """  width=""" & Param(0) & """ height=""" & Param(1) & """ border=""0"" alt=""" & GetNodeText("title") &"""/></div>"
				 case "getcategory" echo GetCategory(param(0))
				 
				 case "getdownphoto" ,"getmoviephoto","getsupplyphoto","getflashphoto"
				  Dim DownPhotoUrl:DownPhotoUrl=GetNodeText("photourl") : If DownPhotoUrl="" Or IsNull(DownPhotoUrl) Then DownPhotoUrl=DomainStr & "images/nopic.gif"
				  if Lcase(left(DownPhotoUrl,7))<>"http://" then DownPhotoUrl=KS.Setting(2) &DownPhotoUrl
				  echo "<img src=""" & DownPhotoUrl & """ height=""" & Param(1) & """ width=""" & Param(0) & """ alt=""" & GetNodeText("title") & """/>"
				
				 case "getflashbyplayer"  '动漫播放器
			       echo  GetFlashPlayer(TagParam)
				 case "getflash"
				   echo GetFlashContent(TagParam)
                 case "getgroupphotos"    '商城组图新增
				       Dim ProImgTotal:ProImgTotal=Conn.Execute("Select count(id) From KS_ProImages Where ProID=" & ItemID)(0)
				       ProPhotoUrl=GetNodeText("bigphoto") : If KS.IsNul(ProPhotoUrl) Then ProPhotoUrl="/images/nopic.gif" : If Left(LCase(ProPhotoUrl),4)<>"http" then ProPhotoUrl=KS.Setting(2) & ProPhotoUrl
					   echo "<div class=""defaultpic""><a href='" & DomainStr & "shop/showpic.asp?id=" & ItemID & "&u=" & ProPhotoUrl & "' target=""_blank""><img title=""查看全部图片"" src=""" & ProPhotoUrl &"""/></a></div>"
					   echo "<div style=""text-align:center"" class=""defaultpictext"">共有美图 <span style=""font-weight:bold;color:#72578C"">"& ProImgTotal & "</span> 张</div>"
					   Dim ImgRS,SelectTop:SelectTop=Param(0) : If Not IsNumeric(SelectTop) Then SelectTop=6
					   Set ImgRS=Conn.Execute("Select top " & SelectTop & " SmallPicUrl,BigPicUrl,GroupName From KS_ProImages Where ProID=" & ItemID & " Order By Id Desc")
					   If Not ImgRS.Eof Then SQL=ImgRS.GetRows(-1) Else SQL=""
					   ImgRS.Close : Set ImgRS=Nothing
					   If IsArray(SQL) Then
					     echo "<div class=""proimglist""><ul>"
					     For i=0 to Ubound(SQL,2)
						  ProPhotoUrl=SQL(0,I) : If Left(LCase(ProPhotoUrl),4)<>"http" then ProPhotoUrl=KS.Setting(2) & ProPhotoUrl
						  echo "<li><a title=""" & SQL(2,I) & """ href='" & DomainStr & "shop/showpic.asp?id=" & ItemID & "&u=" & SQL(1,I) & "' target=""_blank""><img src='" & ProPhotoUrl & "' alt=""" & SQL(2,I) & """ /></a></li>"
						 Next
						 echo "<ul></div>"
					   End If				 
				case "getgroupphoto"     '商城组图
				
						 Dim SQL,DefaultGroupName,DefaultBigPic,DefaultSmallPic,GroupImgList,Spic,Bpic
						 Dim RSG:Set RSG=Conn.Execute("Select ID,ProID,SmallPicUrl,BigPicUrl,GroupName From ks_proimages where ProID=" & ItemID & "  order by id")
						If Not RSG.Eof Then SQL=RSG.GetRows(-1) Else SQL=""
						RSG.Close:Set RSG=Nothing
						If IsArray(SQL) Then
						  For I=0 To Ubound(SQL,2)
						    Spic=SQL(2,I) : If lcase(left(Spic,4))<>"http" Then Spic=KS.Setting(2) & Spic
							Bpic=SQL(3,I) : If lcase(left(Bpic,4))<>"http" Then Bpic=KS.Setting(2) & Bpic
							If I=0 Then
							 DefaultBigPic=Bpic
							 DefaultSmallPic=Spic
							 GroupImgList="<li><img class=""curr"" title=""" &sql(4,i) & """ src=""" & Spic &""" alt=""" & Bpic & """ /></li>"
							Else
							 GroupImgList=GroupImgList & "<li><img title=""" &sql(4,i) & """ src=""" & Spic &""" alt=""" & Bpic & """ /></li>"
							End If
						  Next
						 Else
						    DefaultSmallPic=GetNodeText("photourl")
							DefaultBigPic=GetNodeText("bigphoto")
						 End If
						 
						If KS.IsNul(DefaultBigPic) Then DefaultBigPic=DomainStr & "images/nopic.gif"
						If KS.IsNul(DefaultSmallPic) Then DefaultBigPic=DomainStr & "images/nopic.gif"
						If lcase(left(DefaultBigPic,4))<>"http" Then DefaultBigPic=KS.Setting(2) & DefaultBigPic
						If lcase(left(DefaultSmallPic,4))<>"http" Then DefaultSmallPic=KS.Setting(2) & DefaultSmallPic
						if GroupImgList="" then
							 GroupImgList="<li><img class=""curr""  src=""" & DefaultSmallPic &""" alt=""" & DefaultBigPic & """ /></li>"
						end if
						 Dim G_T:G_T=LFCls.GetConfigFromXML("ProImages","/labeltemplate/label","proimages")
						 G_T = Replace(G_T,"{$GetInstallDir}",KS.Setting(3))
						 G_T = Replace(G_T,"{$GetSiteUrl}",DomainStr)
						 G_T = Replace(G_T,"{$DefaultBigPic}",DefaultBigPic)
						 G_T = Replace(G_T,"{$DefaultSmallPic}",DefaultSmallPic)
						 G_T = Replace(G_T,"{$GroupImgList}",GroupImgList)
						 G_T = Replace(G_T,"{$BigWidth}",Param(0))
						 G_T = Replace(G_T,"{$BigHeight}",Param(1))
						 G_T = Replace(G_T,"{$BigHeight1}",KS.ChkClng(Param(1))-9)
						 G_T = Replace(G_T,"{$BigWidth1}",KS.ChkClng(Param(0))-39)
						 G_T = Replace(Replace(G_T,"{$GetProductName}",GetNodeText("title")),"{$InfoID}",ItemID)
						 Echo G_T
			     case "getproductphoto"
				      Dim ProPhotoUrl:ProPhotoUrl=GetNodeText("photourl")
					  If IsNull(ProPhotoUrl) Or ProPhotoUrl = "" Then ProPhotoUrl=DomainStr & "images/nopic.gif"
					  Dim TempBigPhoto:TempBigPhoto=GetNodeText("bigphoto")
					  If lcase(left(TempBigPhoto,4))<>"http" Then TempBigPhoto=KS.Setting(2) & TempBigPhoto
					  echo "<div align=""center""><img src=""" & TempBigPhoto & """  width=""" & Param(0) & """ height=""" & Param(1) & """ border=""0"" alt=""商品图片""/></div><div style=""text-align:center;margin:8px""><a href=""" & DomainStr & "shop/showpic.asp?id=" & itemid & "&u=" & Server.UrlEncode(TempBigPhoto) & """ target=""_blank""><img src=""" & DomainStr &"images/v.gif"" border=""0"" alt=""""/> 点击看大图</a></div>"
				case "getmovieplaylist"          '影视播放器
					  Dim PerLineNum: PerLineNum = Param(0)
					  Dim NaviPicStr: NaviPicStr = Param(1)
					  Dim PlayListStr,MovieNum,MovieUrlsArr:MovieUrlsArr=Split(GetNodeText("movieurls"),"|||")
					 MovieNum=Ubound(MovieUrlsArr)+1
					 If MovieNum=1 Then
					    echo "<img src=""" & NaviPicStr & """ border=""0"" align=""absmiddle"" alt=""""/> <a href='javascript:Play(" & ItemID & ",0);'>" & GetNodeText("title") & "</a>&nbsp;&nbsp;"
					 Else
					   For I=0 To MovieNum-1
						Dim Marr:Marr=Split(MovieUrlsArr(i),"§")
						If I=0 Then
						  echo "<img src=""" & NaviPicStr & """ border=""0"" align=""absmiddle"" /> <a href='javascript:Play(" & ItemID & "," & I & ");'>" & Marr(0) & "</a>&nbsp;&nbsp;"
						Else
						  echo "<img src=""" & NaviPicStr & """ border=""0"" align=""absmiddle""> <a href='javascript:Play(" & ItemID & "," & I & ");'>" & Marr(0) & "</a>&nbsp;&nbsp;"
						End If
						If Cint(I+1) Mod PerLineNum = 0 Then  echo "<br />"
					   Next
					 End If
			   case "getmoviedownlist" 		 '影片下载 
				  If GetNodeText("downtf")="0" Then
					echo "---"
				  Else
					  PerLineNum = Replace(Replace(Param(0),"""",""),"'","")
					  NaviPicStr = Replace(Replace(Param(1),"""",""),"'","")
					  MovieUrlsArr=Split(GetNodeText("movieurls"),"|||")
					  MovieNum=Ubound(MovieUrlsArr)+1
					 If MovieNum=1 Then
					   echo "<img src=""" & NaviPicStr & """ border=""0"" align=""absmiddle"" alt=""""/> <a href='javascript:Down(" & ItemID & ",0);'>" & GetNodeText("title") & "</a>&nbsp;&nbsp;"
					 Else
					   For I=0 To MovieNum-1
						Marr=Split(MovieUrlsArr(I),"§")
						If I=0 Then
						  echo "<img src=""" & NaviPicStr & """ border=""0"" align=""absmiddle"" alt=""""/> <a href='" & KS.GetDomain & "Movie/down.asp?id=" & ItemID & "&downid="  & I & "' target='_blank'>" & Marr(0) & "</a>&nbsp;&nbsp;"
						Else
						  echo "<img src=""" & NaviPicStr & """ border=""0"" align=""absmiddle"" alt=""""/> <a href='" & KS.GetDomain & "Movie/down.asp?id=" & ItemID & "&downid="  & I & "' target='_blank'>" & Marr(0) & "</a>&nbsp;&nbsp;"
						End If
						If Cint(I+1) Mod PerLineNum = 0 Then  echo "<br />"
					   Next
					 End If
				  End If
			 case "getmoviepageplay" echo GetMoviePagePlay(Param)
			 case "getwritecomments" 
				   Dim CommentID:CommentID=KS.ChkClng(Param(0))
				   If CommentID<>0 Then
				       Dim RS:Set RS=Conn.Execute("Select top 1 ProjectContent,MaxLen From KS_MoodProject Where ID=" & CommentID)
					   If Not RS.Eof Then
					       Dim ProjectContentArr:ProjectContentArr=Split(RS(0),"$$$")
						   Dim PMaxLen:PMaxLen=KS.ChkClng(RS(1))
						   Dim Sstr:Sstr="<table border=""0"" cellspacing=""0"" cellpadding=""0"">"
						   Dim NN,MM,KK:KK=0
						   For NN=0 To Ubound(ProjectContentArr)
						    If Split(ProjectContentArr(kk),"|")(0)<>"" Then
								 Sstr=Sstr & "<tr>"
								 For MM=1 To 2
								 if Split(ProjectContentArr(kk),"|")(0)<>"" then
								 Sstr=Sstr & " <td width=""270"" valign=""top""><input type=""hidden"" name=""score" & KK &""" id=""score" & kk & """>"
								 Sstr=Sstr & "  <div class=""rate-text"">"  & Split(ProjectContentArr(kk),"|")(0) & "：</div>"
								 Sstr=Sstr & "		<div id=""Commentdemo" & KK & """ class=""add_comment_start""></div><div id=""score" & KK & "_desc"" class=""add_comment_start_desc""></div>"
								 Sstr=Sstr & "<script language=""javascript"">"
								 Sstr=Sstr  &"	$('#Commentdemo" & KK & "').rater(null, {maxvalue:5,curvalue:0}, function(el , value) {setRateValue(value, ""score" & KK & """);});"
								 Sstr=Sstr & "</script>"
								 Sstr=Sstr & "</td>"
								 end if
								 KK=KK+1
								 If KK>=Ubound(ProjectContentArr) Then Exit For
							   Next
							   Sstr=Sstr & "</tr>"
						   End If
						   If KK>=Ubound(ProjectContentArr) Then Exit For
						  Next
						  Sstr=Sstr & "</table>"
						   
						   Dim CommentStr:CommentStr=LFCls.GetXMLByNoCache("comments","/posttemplate/label","[@name='post']")

						   CommentStr=Replace(CommentStr,"{$GetSiteUrl}",KS.GetDomain)
						   CommentStr=Replace(CommentStr,"{$ProjectID}",Param(0))
						   CommentStr=Replace(CommentStr,"{$ChannelID}",ModelID)
						   CommentStr=Replace(CommentStr,"{$ItemID}",ItemID)
						   CommentStr=Replace(CommentStr,"{$ScoreItem}",Sstr)
						   CommentStr=Replace(CommentStr,"{$Title}",GetNodeText("title"))
						   If PMaxLen<>"0" Then
						   CommentStr=Replace(CommentStr,"{$MaxLen}",PMaxLen & "个字")
						   CommentStr=Replace(CommentStr,"{$MaxLenNum}",PMaxLen)
						   Else
						   CommentStr=Replace(CommentStr,"{$MaxLen}","不限制")
						   CommentStr=Replace(Replace(CommentStr,"{$MaxLenNum}",0),"{$DisplayZS}","display:none;")
						   End If
						   If EnabledSubDomain Then
						   CommentStr=Replace(CommentStr,"{$Domain}","<script>document.domain=""" & RootDomain &""";</script>")
						   Else
						   CommentStr=Replace(CommentStr,"{$Domain}","")
						   End If
						   echo CommentStr
					   End If
					   RS.Close:Set RS=Nothing
				   End If
				 case "getshowcomments" 
				    CommentID=KS.ChkClng(Param(0))
				    If CommentID<>0 Then
					 Set RS=Conn.Execute("Select top 1 ProjectContent,IsRewrite From KS_MoodProject Where ID=" & CommentID)
					 If Not RS.Eof Then
					     ProjectContentArr=Split(RS(0),"$$$")
						 Dim IsRewrite:IsRewrite=RS(1)
						 Set RS=Conn.Execute("select top 10 * From KS_Comment Where ProjectID=" & CommentID & " and ChannelID=" & ModelID & " and InfoID=" & ItemID & " and verific=1 order by id desc")
						 DiM Floor:Floor=Conn.Execute("Select count(1) From KS_Comment Where ProjectID=" & CommentID & " and ChannelID=" & ModelID & " and InfoID=" & ItemID & " and verific=1")(0)
						 If Not RS.Eof Then
						     Dim AvgStr:AvgStr=""
							For NN=0 To Ubound(ProjectContentArr)
							 if Split(ProjectContentArr(NN),"|")(0)<>"" then
							   AvgStr=AvgStr & Split(ProjectContentArr(NN),"|")(0) & ":<span style='color:red'>" & round(conn.execute("select avg(m" & nn&") from ks_comment where channelid=" & ModelID & " and infoid=" & ItemId &" and projectid=" & Commentid)(0),2) &"</span> 分&nbsp;&nbsp;"
							 End If
							Next
						     CommentStr=LFCls.GetXMLByNoCache("comments","/posttemplate/label","[@name='showtitle']")
							 Dim RewriteUrl
							 If IsRewrite="1" Then
							  RewriteUrl=DomainStr & "plus/rating-" & modelid & "-" & itemid &"-" & param(0) & ".html"
							 Else
							  RewriteUrl=DomainStr & "plus/rating.asp?channelid=" & modelid & "&infoid="& itemid & "&projectid=" &param(0)
							 End If
						     echo Replace(Replace(Replace(CommentStr,"{$CommentNum}",floor),"{$MoreRateUrl}",RewriteUrl),"{$GetAgvStr}",AvgStr)
							 CommentStr=LFCls.GetXMLByNoCache("comments","/posttemplate/label","[@name='show']")
							 Do While Not RS.Eof
							   Dim StarStr:StarStr=""
							   For NN=0 To Ubound(ProjectContentArr)
							    if Split(ProjectContentArr(NN),"|")(0)<>"" then
							     StarStr=StarStr & Split(ProjectContentArr(NN),"|")(0) & "：<img src='" & KS.GetDomain & "images/star/star-" & RS("M"&nn)& ".jpg'/>&nbsp;&nbsp;"
								end if
							   Next
							   Dim ContentStr:ContentStr=rs("content")
							   If Not KS.IsNul(rs("ReplyContent")) Then
							    ContentStr=ContentStr &"<div style='margin:10px;padding:4px;color:red;border:1px dashed #ccc;background:#FFFFEE;'>管理员回复：" &rs("ReplyContent") &"</div>" 
							   End If
							   Dim UserIP:UserIP=left(rs("userip"), InStrRev(rs("userip"), ".")) & "*"
							   echo Replace(Replace(Replace(Replace(Replace(Replace(Replace(commentstr,"{$Title}",rs("title")),"{$Content}",ContentStr),"{$UserName}",rs("username")),"{$PostTime}",rs("AddDate")),"{$UserIP}",UserIP),"{$ShowStar}",StarStr),"{$Floor}",Floor)
							   Floor=Floor-1
							  RS.MoveNext
							 Loop
							 RS.Close
							 Set RS=Nothing
						 End If
					 End If
					End If
			 case else
			   If left(lcase(TagName),3)="js_" then
			    Call JSCls.Equal(TagName,Param,Templates)	  
			   end if  
			 end select
		    End If
	   End Function

	
	  '替换频道专用标签
		Sub ParseChannelLabel(ByVal sTemp)
		   on error resume next
		   sTemp = Lcase(sTemp)
		  select case sTemp
		     case "getchannelid" echo Fcls.ChannelID
			 case "getchannelname" echo KS.C_S(FCls.ChannelID,1)
			 case "getitemname" echo KS.C_S(FCls.ChannelID,3)
			 case "getitemurl" echo KS.C_S(FCls.ChannelID,4)
		  End Select
		  
		   If FCls.RefreshFolderID="0" Or FCls.RefreshFolderID="" Then Exit Sub
		   	Dim I,ClassBasicInfoArr,ClassDefineContentArr
			ClassBasicInfoArr    = Split(KS.C_C(FCls.RefreshFolderID,6),"||||")
			ClassDefineContentArr= Split(KS.C_C(FCls.RefreshFolderID,7),"||||")

		    
		   select case sTemp
			 case "getclassid" echo FCls.RefreshFolderID
			 case "getsmallclassid" echo KS.C_C(FCls.RefreshFolderID,9)
			 case "gettopclassname" echo KS.C_C(Split(KS.C_C(FCls.RefreshFolderID,8),",")(0),1)
			 case "gettopclassurl" echo KS.GetFolderPath(KS.C_C(Split(KS.C_C(FCls.RefreshFolderID,8),",")(0),0))
			 case "getparentid" 
			  if FCls.RefreshType="Content" Then
			   echo KS.C_C(FCls.RefreshFolderID,13)
			 Else
			   echo FCls.RefreshParentID
			 End If
			 case "getparentclassid" 
			  if FCls.RefreshType="Content" Then
			   echo KS.C_C(KS.C_C(FCls.RefreshFolderID,13),9)
			 Else
			   echo KS.C_C(FCls.RefreshParentID,9)
			 End If
			 case "getparenturl"  
			 if FCls.RefreshType="Content" Then
			   echo KS.GetFolderPath(KS.C_C(FCls.RefreshFolderID,13))
			 else
			   If FCls.RefreshParentID="0" Then echo KS.Setting(2) else echo KS.GetFolderPath(FCls.RefreshParentID)
			 end if
			 case "getparentclassname" 
			 if FCls.RefreshType="Content" Then
			 echo KS.C_C(KS.C_C(FCls.RefreshFolderID,13),1)
			 Else
			 echo KS.C_C(FCls.RefreshParentID,1)
			 End If
			 case "getclassname" echo KS.C_C(FCls.RefreshFolderID,1)
			 case "getclassurl" echo KS.GetFolderPath(FCls.RefreshFolderID)
		  end select
		  
		  If IsArray(ClassBasicInfoArr) Then
		    select case sTemp
		     case "getclasspic" echo "<img src=""" & ClassBasicInfoArr(0) & """ border=""0"" alt="""" />"
			 case "getclassintro" echo ClassBasicInfoArr(1)
			 case "getclass_meta_keyword" echo ClassBasicInfoArr(2)
			 case "getclass_meta_description" echo ClassBasicInfoArr(3)
			end select
		  End If
		    
		  If IsArray(ClassDefineContentArr) Then
		     For I=1 To Ubound(ClassDefineContentArr)+1
			   if sTemp="getclassdefinecontent" & I  then echo ClassDefineContentArr(I-1)
			 Next
		  End If
		  if err then err.clear
		End Sub
		
		'替换RSS标签
		Sub ParseRssLabel(sTemp)
		   IF KS.Setting(83)=0 Then Exit Sub
		   Dim CurrentClassID:CurrentClassID=FCls.RefreshFolderID
		   Dim ChannelID:ChannelID=FCls.ChannelID
		   select case Lcase(sTemp)
		      case "rss" 
			    select case Lcase(FCls.RefreshType)
				 case "index" echo GetRssLink("rss.asp")
				 case "folder" echo GetRssLink("Rss.asp?ChannelID=" & ChannelID & "&ClassID=" &CurrentClassID & "")
			    end select
			 case "rsselite"
			    select case Lcase(FCls.RefreshType)
				 case "index" echo GetRssLink("Rss.asp?Elite=1")
				 case "folder" echo GetRssLink("Rss.asp?ChannelID=" & ChannelID & "&ClassID=" &CurrentClassID & "&Elite=1")
			    end select
			 case "rsshot"
			    select case Lcase(FCls.RefreshType)
				 case "index" echo GetRssLink("Rss.asp?Hot=1")
				 case "folder" echo GetRssLink("Rss.asp?ChannelID=" & ChannelID & "&ClassID=" &CurrentClassID & "&Hot=1")
			    end select
		   end select
		End Sub
		'取得每个频道的RSS链接，结合ParseRssLabel调用
		Function GetRssLink(LinkStr)
		   GetRssLink="<a href=""" & DomainStr &"plus/" &  LinkStr & """ target=""_blank""><img src=""" & DomainStr & "Images/Rss.gif" & """ border=""0""></a>"
		End Function
		
		'扫描并替换附件信息
		Function ScanAnnex(sTemplate)
		If Instr(sTemplate,"[UploadFiles]")=0 or Instr(sTemplate,"[/UploadFiles]")=0  Then ScanAnnex=ReplaceEmot(sTemplate) : Exit Function
		Dim TempStr,iPosLast, iPosCur,iPosBegin
		iPosLast    = 1
		Do While True 
			iPosCur    = InStr(iPosLast, sTemplate, "[UploadFiles]") 
			If iPosCur>0 Then
					TempStr=TempStr & Mid(sTemplate, iPosLast, iPosCur-iPosLast)
					
					Dim iPosCur1, sToken, sTemp,FileInfoArr,FileSize,Ext,Title
					iPosBegin=iPosCur+13
					iPosCur1      = InStr(iPosBegin, sTemplate, "[/UploadFiles]")
					sTemp        = Mid(sTemplate,iPosBegin,iPosCur1-iPosBegin)
					FileInfoArr  = split(sTemp,",")
					iPosBegin    = iPosCur1+14
					If Ubound(FileInfoArr)>=1 Then
					  FileSize=KS.ChkClng(FileInfoArr(1))
					  If FileSize<1 Then
						FileSize=FormatNumber(FileSize,2,-1,0,-1) & " bytes"
					  ElseIf FileSize>1024*1024 Then
						FileSize=FormatNumber(round(FileSize/1024/1024,2),2,-1,0,-1) & " MB"
					  Else
						FileSize=FormatNumber(round(FileSize/1024,2),2,-1,0,-1) & " KB"
					  End If
					End If
					If Ubound(FileInfoArr)>=2 Then Ext=FileInfoArr(2) Else Ext="rar"
					If Ubound(FileInfoArr)>=3 Then Title="点击下载文件:" & FileInfoArr(3) Else Title="点击下载该文件"
					tempstr=tempstr  & "<table border=""0"" class=""annex"" cellspacing=""2""><tr><td height=""20"" class=""annextitle"">&nbsp;<B>下载信息</B>&nbsp;&nbsp;[文件大小："  & FileSize &" 下载次数：<script src=""" & KS.GetDomain & "item/filedown.asp?action=hits&id=" & FileInfoArr(0) & """></script> 次]"
					tempstr=tempstr & "<tr><td><img src=""" & KS.GetDomain & "editor/ksplus/fileicon/" & Ext & ".gif"" /> <a href=""" & KS.GetDomain & "item/filedown.asp?id=" & FileInfoArr(0) & "&Ext=" & Ext & "&fname=" & server.URLEncode(FileInfoArr(3)) & """ target=""downframe"">" & Title & "</a></td></tr>"
					tempstr=tempstr & "</table><iframe name='downframe' src='about:blank' style='display:none' width='0' height='0'></iframe>" 
					iPosLast=iPosBegin
		
			Else 
					TempStr=TempStr &Mid(sTemplate, iPosLast)
				   Exit Do  
			End If 
		Loop
		 ScanAnnex=ReplaceEmot(TempStr)
		End Function
		'替换表情
		Function ReplaceEmot(c)
		 Dim str:str=":)|:(|:D|:'(|:@|:o|:P|:$|;P|:L|:Q|:lol|:loveliness:|:funk:|:curse:|:dizzy:|:shutup:|:sleepy:|:hug:|:victory:|:time:|:kiss:|:handshake|:call:|55555|不是我|不要啊|亲一亲|加油|向前进|吓死你|呐喊|鸣哇|呵呵|呸|哈哈|哼|嗯|嘿嘿|困死了|天打雷劈|好闷啊|对不起|开心|很忙|抓狂|放电|无聊|汗一个|看我历害|脑残|飞吻|good|不妙啊|不是啦|交出来|亲亲|偷笑|哭|喜欢|嗯|坏笑|太好啦|好主意|好同志|悄悄走|我爱你|打你|晕菜|没良心"
		 Dim strArr:strArr=Split(str,"|")
		 Dim K,NS
		 For K=1 To 70
		  NS=Right("0" & K,2)
		  c=replace(c,"[em"&NS &"]","<img title='" & strarr(k-1) & "' alt='" & strarr(k-1) & "' src='" & KS.Setting(2) &KS.Setting(3) & "editor/ubb/images/smilies/default/" & NS & ".gif'/>")
		 Next
		 ReplaceEmot=C
	   End Function
		
		
		'*******************************************************************************************************
		'函数名：KSLabelReplaceAll
		'作  用：替换所有标签
		'参  数：F_C 模板内容
		'返回值：替换过的模板内容
		'********************************************************************************************************
		Public Function KSLabelReplaceAll(F_C)
		          F_C = ReplaceAllLabel(F_C)                    
				  F_C = ReplaceLableFlag(F_C)                   '替换函数标签
				  F_C = ReplaceGeneralLabelContent(F_C)        '替换通用标签 如{$GetWebmaster}
				  F_C = ReplaceRA(F_C, "")
				  KSLabelReplaceAll=F_C
	    End Function
		'*******************************************************************************************************
		'函数名：LoadTemplate
		'作 用：取出模板内容
		'参 数：TemplateFname模板地址
		'返回值：模板内容
		'********************************************************************************************************
		Function LoadTemplate(TemplateFname)
					  on error resume next
					  TemplateFname=Replace(TemplateFname,"{@TemplateDir}",KS.Setting(3) & KS.Setting(90))
					  TemplateFname =Replace(TemplateFname, "//", "/")
						dim str,stm
						set stm=server.CreateObject("adodb.stream")
						stm.Type=2 '以本模式读取
						stm.mode=3 
						stm.charset="UTF-8"
						stm.open
						stm.loadfromfile server.MapPath(TemplateFname)
						str=stm.readtext
						stm.Close
						set stm=nothing
						if err then
						LoadTemplate = "模板文档不存在或模板路径填写不正确！"
						else
						LoadTemplate=str
						End if
						LoadTemplate=Replace(LoadTemplate,"{$UID}",Request("Uid"))
						LoadTemplate=LoadTemplate & Published
		
		End Function
		'**************************************************
		'函数名：ReplaceLableFlag
		'作  用：替换并执行系统函数标签
		'参  数： Content  ----待替换内容
		'返回值：返回用","隔开的字符串
		'**************************************************
		Function ReplaceLableFlag(Content)
			Dim regEx, Matches, Match, TempStr
			Set regEx = New RegExp
			ReplaceLableFlag = Content
			Set regEx = New RegExp
			regEx.Pattern = "{Tag([\s\S]*?):(.+?)}([\s\S]*?){/Tag\1}"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			For Each Match In Matches
				ReplaceLableFlag = Replace(ReplaceLableFlag,Match.Value,KSLabel.GetLabel(Match.Value))
			Next
		End Function
		
		
		'扫描系统函数标签
		Function ScanSysLabel(Content)
		  Dim iPosLast, iPosCur,Tstr
			iPosLast    = 1
			Do While True 
				iPosCur    = InStr(iPosLast, Content, "{LB_") 
				If iPosCur>0 Then
					Tstr=tstr &  Mid(Content, iPosLast, iPosCur-iPosLast)
					iPosLast  = ParseSysLabel(Content, iPosCur+4,Tstr)
				Else 
					Tstr=tstr & Mid(Content, iPosLast)
					Exit do
				End If 
		   Loop 
		   ScanSysLabel=Tstr
		End Function
		Function ParseSysLabel(sTemplate, iPosBegin,Tstr)
			Dim iPosCur, sToken, sTemp,MyNode
			iPosCur      = InStr(iPosBegin, sTemplate, "}")
			sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
			iPosBegin    = iPosCur+1
			Set MyNode   = Application(KS.SiteSN&"_labellist").documentElement.SelectSingleNode("labellist[@labelname='{LB_" & sTemp & "}']")
			If Not MyNode Is Nothing Then Tstr=Tstr &  MyNode.text 
			
			ParseSysLabel= iPosBegin
		End Function
		
		
		'*********************************************************************************************************
		'函数名：ReplaceAllLabel
		'作  用：将标签名称转换成对应标签内容
		'参  数： Content需转换的内容
		'*********************************************************************************************************
		Function ReplaceAllLabel(Content)
			dim Node
			Call LoadLabelToCache()    '加载标签
			
		    Content=ScanSysLabel(Content)

			Call LoadJSFileToCache()   '加载JS
			For Each Node in Application(KS.SiteSN&"_jslist").documentElement.SelectNodes("jslist")
				Content=Replace(Content,Node.selectSingleNode("@jsname").text,Node.text)
			Next
			If Lcase(Fcls.RefreshType)<>"content" Then Content=ReplaceSQLLabel(Content)

			ReplaceAllLabel=Content
		End Function
		
		Function ReplaceSQLLabel(Content)
			'替换自定义函数标签 
			Dim DCls:Set Dcls=New DIYCls
			ReplaceSQLLabel=DCls.ReplaceUserFunctionLabel(Content)
			Set DCls=nothing
		End Function

	
		'加载数据库的所有标签到缓存	
		 Sub LoadLabelToCache()
			If Not IsObject(Application(KS.SiteSN&"_labellist")) Then
					Set  Application(KS.SiteSN&"_labellist")=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
					Application(KS.SiteSN&"_labellist").appendChild(Application(KS.SiteSN&"_labellist").createElement("xml"))
						Dim i,SQL,Node
						Dim RS:Set RS = Server.CreateObject("ADODB.Recordset")
						RS.Open "Select ID,LabelType,LabelName,LabelContent from KS_Label Where LabelType<>5 and LabelType<>7", Conn, 1, 1
						If Not RS.Eof Then SQL=RS.GetRows(-1)
						RS.Close:Set RS = Nothing
						If IsArray(SQL) Then
							for i=0 to Ubound(SQL,2)
								 Set Node=Application(KS.SiteSN&"_labellist").documentElement.appendChild(Application(KS.SiteSN&"_labellist").createNode(1,"labellist",""))
								 Node.attributes.setNamedItem(Application(KS.SiteSN&"_labellist").createNode(2,"labelname","")).text=SQL(2,I)
								 Node.attributes.setNamedItem(Application(KS.SiteSN&"_labellist").createNode(2,"labelid","")).text=SQL(0,I)
								If SQL(1,I) = 1 Then
								 Node.text=ReplaceFreeLabel(SQL(3,I))
								Else
								 Node.text=Replace(SQL(3,I),"labelid=""0""","labelid=""" & SQL(0,I) & """")
								End IF
							next
						End If
			End if
		End Sub
		
		'加载数据库的所有JS到缓存
		Sub LoadJSFileToCache()
			If Not IsObject(Application(KS.SiteSN&"_jslist")) Then
					Set  Application(KS.SiteSN&"_jslist")=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
					Application(KS.SiteSN&"_jslist").appendChild( Application(KS.SiteSN&"_jslist").createElement("xml"))
						Dim i,SQL,Node
						Dim RS:Set RS = Server.CreateObject("ADODB.Recordset")
						RS.Open "Select JSID,JSName,JSFileName from KS_JSFile", Conn, 1, 1
						If Not RS.Eof Then SQL=RS.GetRows(-1)
						RS.Close:Set RS = Nothing
						If IsArray(SQL) Then
							for i=0 to Ubound(SQL,2)
								 Set Node=Application(KS.SiteSN&"_jslist").documentElement.appendChild(Application(KS.SiteSN&"_jslist").createNode(1,"jslist",""))
								 Node.attributes.setNamedItem(Application(KS.SiteSN&"_jslist").createNode(2,"jsname","")).text=SQL(1,I)
								 Node.text="<script charset=""utf-8"" language=""javascript"" src=""" & Replace(KS.Setting(3) & KS.Setting(93),"//","/") & Trim(SQL(2,I)) & """></script>"
							next
						End If
			End if
		End Sub

	'替换自由标签为内容,仅替换一级
	Function ReplaceFreeLabel(sTrC)
			dim node
			If not IsObject(Application(KS.SiteSN&"_ReplaceFreeLabel")) then
					Dim RS:Set RS = Server.CreateObject("ADODB.Recordset")
					RS.Open "Select LabelName,LabelContent,ID from KS_Label", Conn, 1, 1
					if Not RS.eof then
						'KS.Value=RS.GetString(,,"^||^","^%%%^","")
						Set Application(KS.SiteSN&"_ReplaceFreeLabel")=KS.ArrayToXml(RS.GetRows(-1),rs,"row","")
					end if
					RS.Close:Set RS = Nothing

			End if
			For Each Node In Application(KS.SiteSN&"_ReplaceFreeLabel").documentElement.SelectNodes("row")
					sTrC = Replace(sTrC,trim(Node.SelectSingleNode("@labelname").text),Replace(Node.SelectSingleNode("@labelcontent").text,")}","," & Node.SelectSingleNode("@id").text &")}"))
			next
			'ReplaceFreeLabel = ReplaceGeneralLabelContent(sTrC)
			ReplaceFreeLabel = ScanSysLabel(sTrC)
		End Function

		'*********************************************************************************************************
		'函数名：FSOSaveFile
		'作  用：生成文件
		'参  数： Content内容,路径 注意虚拟目录
		'*********************************************************************************************************
		Sub FSOSaveFile(Content, FileName)
			dim stm:set stm=server.CreateObject("adodb.stream")
			stm.Type=2 '以文本模式读取
			stm.mode=3
			stm.charset="utf-8"
			stm.open
			stm.WriteText content
			stm.SaveToFile server.MapPath(FileName),2 
			stm.flush
			stm.Close
			set stm=nothing
		End Sub
		
		'*********************************************************************************************************
		'函数名：RefreshJS
		'作  用：发布JS
		'参  数：JSName JS名称
		'*********************************************************************************************************
		Sub RefreshJS(JSName)
			Dim JSRS, SqlStr, JSContent
			Set JSRS = Server.CreateObject("ADODB.Recordset")
			SqlStr = "Select * From KS_JSFile Where JSName='" & Trim(JSName) & "'"
			JSRS.Open SqlStr, Conn, 1, 1
			If JSRS.EOF And JSRS.BOF Then
			 JSRS.Close:Set JSRS = Nothing:Exit Sub
			End If
			  Dim JSConfig, JSFileName, SaveFilePath, JSDir, JSType
			  JSFileName = Trim(JSRS("JSFileName"))
			  JSDir = Trim(KS.Setting(93))
			  JSType = Trim(JSRS("JSType"))
			  If Left(JSDir, 1) = "/" Or Left(JSDir, 1) = "\" Then JSDir = Right(JSDir, Len(JSDir) - 1)
			  SaveFilePath = KS.Setting(3) & JSDir
			  Call KS.CreateListFolder(SaveFilePath)
			   
			   JSConfig = Trim(JSRS("JSConfig"))
			  If JSType = "0" Then
				JSContent=Replace(Replace(Replace(Replace(KSLabel.GetLabel(JSConfig), Chr(13)& Chr(10), ""),"'","\'"),"""","\"""),vbcrlf,"")             
				JSContent=Replace(JSContent,Chr(13) ,"")
				JSContent = "document.write('" & JSContent & "');"
			  Else
				Dim FreeType
				FreeType = Left(JSConfig, InStr(JSConfig, ",") - 1) '取出自由JS的类型
				JSConfig = Replace(JSConfig, FreeType & ",", "")
				
				Select Case FreeType      '根据函数做相应的操作
				  Case "GetExtJS"          '扩展JS
					 JSConfig = Replace(JSConfig, "'", """")
					 JSConfig = ReplaceLableFlag(ReplaceAllLabel(JSConfig))
					 JSConfig = ReplaceGeneralLabelContent(JSConfig)
					 JSConfig = Replace(Replace(Replace(JSConfig, Published, ""),"'","\'"),"""","\""")
					 JSContent = ReplaceJsBr(JSConfig)
				  Case "GetWordJS"
					 JSConfig = Replace(Trim(JSConfig), """", "")   '替换原参数的双引号为空
					 JSContent = RefreshWordJS(Trim(JSRS("JSID")), JSConfig)           '替换文字JS
				  Case Else
					 JSContent = ""
				End Select
			End If
			  Call FSOSaveFile(JSContent, SaveFilePath & JSFileName)
			 JSRS.Close:Set JSRS = Nothing
		End Sub
		Function ReplaceJsBr(Content)
		 Dim i
		 Dim JsArr:JSArr=Split(Content,Chr(13) & Chr(10))
		 For I=0 To Ubound(JsArr)
		   ReplaceJsBr=ReplaceJsBr & "document.writeln('" & JsArr(I) &"')" & vbcrlf 
		 Next
		End Function
		'*********************************************************************************************************
		'函数名：RefreshWordJS
		'作  用：发布文字JS
		'参  数：JSID JSID,JSConfig JS参数
		'*********************************************************************************************************
		Function RefreshWordJS(JSID, JSConfig)
		     Dim JSConfigArr:JSConfigArr = Split(JSConfig, ",")
			 If UBound(JSConfigArr) = 17 Then
					RefreshWordJS = KSLabel.RefreshCss(JSID, UCase(JSConfigArr(0)), JSConfigArr(1), JSConfigArr(2), JSConfigArr(3), JSConfigArr(4), JSConfigArr(5), JSConfigArr(6), JSConfigArr(7), JSConfigArr(8), JSConfigArr(9), JSConfigArr(10), JSConfigArr(11), JSConfigArr(12), JSConfigArr(13), JSConfigArr(14), JSConfigArr(15), JSConfigArr(16), JSConfigArr(17))
					RefreshWordJS = Replace(RefreshWordJS, "'", """")
					RefreshWordJS = "document.write('" & RefreshWordJS & "');"
			 Else
					RefreshWordJS = "document.write('标签参数溢出！');"
			 End If
		End Function
		
		'=================================以下为相关栏目,内容页,频道首页等的刷新函数=====================================
		
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名：RefreshContent
		'作  用：刷新内容页面
		'参  数： 无
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function RefreshContent()
			 Dim TFileContent, F_C, FilePath, FilePathAndName, FilePathAndNameTemp, sFname,Fname, FExt, TempFileContent, Content, ContentArr, I, N, CurrPage, PageStr, Flag
			 Dim TemplateID
			   TID = Trim(Node.SelectSingleNode("@tid").text)
			   Call FCls.SetContentInfo(ModelID,Tid,ItemID,Node.SelectSingleNode("@title").text)
			    If ModelID=8 Then
			     TemplateID      = KS.C_C(Tid,5)
				Else
				 TemplateID      = Node.SelectSingleNode("@templateid").text
				End If
			   TempFileContent = LoadTemplate(TemplateID)
			   TempFileContent = ReplaceAllLabel(TempFileContent)
			   If InStr(TempFileContent, "{Tag:GetRelativeList") <> 0 Then TempFileContent = Replace(TempFileContent, "{Tag:GetRelativeList", "{UnTag:GetRelativeList"):Flag = True  Else Flag = False
			   If InStr(TempFileContent, "{Tag:GetLocation") <> 0 and instr(TempFileContent,"showtitle=""true""}{/Tag}")<>0 Then TempFileContent = Replace(TempFileContent, "{Tag:GetLocation", "{UnTag:GetLocation"):Flag = True  Else Flag = False

			   If Flag = True Then
				TFileContent = ReplaceLableFlag(TempFileContent)
			   ElseIf (TemplateID <> FCls.RefreshTemplateID) Or (Tid <> FCls.RefreshCurrTid) Or FCls.RefreshTempFileContent = "" Then
				FCls.RefreshCurrTid = Tid
				FCls.RefreshTemplateID = TemplateID
				FCls.RefreshTempFileContent = ReplaceLableFlag(TempFileContent)  '替换函数标签
				TFileContent = FCls.RefreshTempFileContent
			   Else
				TFileContent = FCls.RefreshTempFileContent
			   End If
			  

			  on error resume next
			  sFname = Trim(Node.SelectSingleNode("@fname").text)
			  FExt   = Mid(sFname, InStrRev(sFname, ".")) '分离出扩展名
			  Fname = Replace(sFname, FExt, "")  '分离出文件名 如 2005/9-10/1254ddd
			  
			  FilePathAndNameTemp =KS.LoadFsoContentRule(ModelID,Tid,ItemId)
			  Dim ShowUrl:ShowUrl=KS.LoadInfoUrl(ModelID,Tid,"",ItemId)
			  FilePathAndName = FilePathAndNameTemp & sFname
			  FilePath = Replace(FilePathAndName, Mid(FilePathAndName, InStrRev(FilePathAndName, "/")), "")
			  
			  Call KS.CreateListFolder(FilePath)
			  
			  '判断是不是转向链接
			  If KS.C_S(ModelID,6)=1 Then
			    if node.SelectSingleNode("@changes").text="1" then
				 Templates=""
				  echoln "<span style=""display:none""><Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?Action=Count&GetFlag=0&m=1&ID=" & ItemId & """></Script></span>"
				  echoln "<script type=""text/javascript"">"
				  echoln "<!--"
				  echoln " location.href='" & Node.SelectSingleNode("@articlecontent").text & "';"
				  echoln "//-->"
				  echoln "</script>"
				 Call FSOSaveFile(Templates, FilePathAndName)
				 Exit Function
				end If
			  End If
			  '判断是不是收费信息
			  IF KS.C_S(ModelID,6)=1 or KS.C_S(ModelID,6)=2 or KS.C_S(ModelID,6)=4 Then
			    If Node.SelectSingleNode("@readpoint").text>0 or Node.SelectSingleNode("@infopurview").text="2" Or (Node.SelectSingleNode("@infopurview").text=0 And (KS.C_C(Tid,3)=1 Or KS.C_C(Tid,3)=2)) Then
				  Templates=""
				  echoln "<script type=""text/javascript"">"
				  echoln "<!--"
				  echoln "  location.href='" & KS.Setting(3) & "item/show.asp?m=" & ModelID & "&d=" & ItemID &"';"
				  echoln "//-->"
				  echoln "</script>"
				  Call FSOSaveFile(Templates, FilePathAndName)
				 Exit Function
				End If
			  End If
			  
			  
			  
			  Dim StartPage,K
			  Select Case Cint(KS.C_S(ModelID,6))
			  Case 1   '文章模型
					  Content = Node.SelectSingleNode("@articlecontent").text
					  Content =Replace(Content,"ジ","")  '过滤掉，不然会乱码
					  If Node.SelectSingleNode("@postid").text<>"0" Then Content=UbbCode(Content,1)
					  If KS.IsNul(Content) Then Content = " "
					  ContentArr = Split(Content, "[NextPage]")
					  TotalPage = UBound(ContentArr) + 1
					  For I = 0 To UBound(ContentArr)
					     CurrPage = I + 1
					     GetPrevNextUrl TotalPage,CurrPage,ShowUrl,sFname,Fname,FExt,Tid    '得到上一页及下一页URL
						 PageStr=GetContentPage(TotalPage,CurrPage,ShowUrl,sFname,Fname,FExt) '取得分页
						 F_C = TFileContent
					    If CurrPage <> 1 Then FilePathAndName = FilePathAndNameTemp & Fname & "_" & CurrPage & FExt
						Dim PageTitleArr,PageTitle
						PageTitle=Node.SelectSingleNode("@pagetitle").text
						If Not KS.IsNul(PageTitle) Then
							  PageTitleArr=Split(PageTitle,"§")
							  If CurrPage-1<=Ubound(PageTitleArr) Then
							   F_C=Replace(F_C,"{$GetArticleTitle}",PageTitleArr(CurrPage-1))
							  End If
						ElseIF Currpage>1 Then
							   F_C=Replace(F_C,"{$GetArticleTitle}",GetNodeText("title") & "(" & currpage & ")")
						End IF
						
					   If InStr(F_C, "{UnTag:GetRelativeList") <> 0 Or InStr(F_C, "{UnTag:GetLocation") <> 0 Then  F_C = ReplaceLableFlag(Replace(F_C, "{UnTag:", "{Tag:"))
					   PageContent="<div id=""MyContent"">" & ContentArr(I) & "</div>" & PageStr
					   Templates = ""
					   Scan F_C
					   F_C = Templates
					   F_C = Replace(Replace(F_C,"[KS_Charge]",""),"[/KS_Charge]","")
					   
					   If Instr(F_C,"[KS_ShowIntro]")<>0 Then
							  If CurrPage=1 Then
								F_C=Replace(Replace(F_C,"[KS_ShowIntro]",""),"[/KS_ShowIntro]","")
							  Else
								F_C=Replace(F_C,KS.CutFixContent(F_C, "[KS_ShowIntro]", "[/KS_ShowIntro]", 1),"")
							  End If
					   End If

					   F_C = ReplaceGeneralLabelContent(F_C)
					   F_C = ReplaceRA(F_C, Trim(KS.C_C(Tid,4))) 
					   F_C = Replace(Replace(Replace(Replace(F_C,"{§","{$"),"{#LB","{LB"),"{#SQL","{SQL"),"{#=","{=")
					   Call FSOSaveFile(F_C, FilePathAndName)
					Next
			case 2  '图片模型
					  Content=Node.SelectSingleNode("@picurls").text
					  If IsNull(Content) Then Content = "" 
					  ContentArr = Split(Content, "|||") : TotalPage  = UBound(ContentArr) + 1
					  Dim ShowStyle,PageNum,Tp
					  ShowStyle=KS.ChkClng(Node.SelectSingleNode("@showstyle").text) : If ShowStyle=0 Then ShowStyle=1
					  PageNum=KS.ChkClng(Node.SelectSingleNode("@pagenum").text) : If PageNum=0 Then PageNum=10
					  If (ShowStyle=1 or ShowStyle=2 Or ShowStyle=4) And TotalPage<=1 Then ShowStyle=3
					  Select Case ShowStyle
					   Case 5
							 Tp=LFCls.GetConfigFromXML("picturelabel","/labeltemplate/label","style5")
							 For i=1 To TotalPage
							   if i=1 then DefaultImageSrc=Split(ContentArr(i-1),"|")(1)
							   ThumbList=ThumbList & "<LI class=""sel"" onclick='picchang1(""" & Split(ContentArr(i-1),"|")(1) & """,""" & Split(ContentArr(i-1),"|")(1) & """)'><IMG width=""119"" height=""90"" src=""" & Split(ContentArr(i-1),"|")(2) & """><LI>" &vbcrlf
							Next
							Tp=Replace(Tp,"{$ShowThumbList}",ThumbList)
							Tp=Replace(Tp,"{$DefaultImageSrc}",DefaultImageSrc)
							PageContent=Tp
							  F_C = TFileContent
							  If InStr(F_C, "{UnTag:GetRelativeList") <> 0 Or InStr(F_C, "{UnTag:GetLocation") <> 0 Then  F_C = ReplaceLableFlag(Replace(F_C, "{UnTag:", "{Tag:"))
							  Templates = "" : Scan F_C
							  F_C = Templates
							  F_C = ReplaceGeneralLabelContent(F_C)
							  F_C = ReplaceRA(F_C, Trim(KS.C_C(Tid,4))) 
							  Call FSOSaveFile(F_C, FilePathAndName)

					   case 1
					       Tp=LFCls.GetConfigFromXML("picturelabel","/labeltemplate/label","style1")
					       Dim ThumbList,DefaultImageSrc,DefaultImageIntro,r,Tpage,ThumbPic
						   For I = 0 To TotalPage - 1
							 CurrPage = I + 1
							 GetPrevNextUrl TotalPage,CurrPage,ShowUrl,sFname,Fname,FExt,Tid    '得到上一页及下一页URL
								ThumbList=""
								For n=1 To TotalPage
								 ThumbPic=Split(ContentArr(n-1),"|")(2) : If lcase(Left(ThumbPic,4))<>"http" Then ThumbPic=KS.Setting(2) & ThumbPic
								 If N=1 Then
								  If CurrPage = N Then
									ThumbList=ThumbList &"<li><a class=""currthumb"" href=""" & ShowUrl & sFname &""" target=""_self""><img src=""" & ThumbPic &""" border=""0""/></a></li>"
								  Else
								   ThumbList=ThumbList &"<li><a class=""normalthumb"" href=""" & ShowUrl & sFname &""" target=""_self""><img src=""" & ThumbPic &""" border=""0""/></a></li>"
								  End If
								 Else
								  If CurrPage = N Then
									ThumbList=ThumbList &"<li><a class=""currthumb"" href=""" & ShowUrl & Fname & "_" & N & FExt &""" target=""_self""><img src=""" & ThumbPic &""" border=""0""/></a></li>"
								  Else
								    ThumbList=ThumbList &"<li><a class=""normalthumb"" href=""" & ShowUrl & Fname & "_" & N & FExt &""" target=""_self""><img src=""" & ThumbPic &""" border=""0""/></a></li>"
								  End If
								 End If
								Next
								DefaultImageSrc=Split(ContentArr(CurrPage-1), "|")(1) :If lcase(Left(DefaultImageSrc,4))<>"http" Then DefaultImageSrc=KS.Setting(2) & DefaultImageSrc
								DefaultImageIntro=Split(ContentArr(CurrPage-1), "|")(0)
						    If CurrPage <> 1 Then FilePathAndName = FilePathAndNameTemp & Fname & "_" & CurrPage & FExt
							F_C = TFileContent
							If InStr(F_C, "{UnTag:GetRelativeList") <> 0 Or InStr(F_C, "{UnTag:GetLocation") <> 0 Then  F_C = ReplaceLableFlag(Replace(F_C, "{UnTag:", "{Tag:"))
							Dim PicSrc :PicSrc=Split(ContentArr(I), "|")(1)
							If (Lcase(Left(PicSrc,4))<>"http") Then PicSrc=KS.Setting(2) & PicSrc
							  PageContent=Replace(Tp,"{$PrevUrl}",PrevUrl)
							  PageContent=Replace(PageContent,"{$NextUrl}",NextUrl)
							  PageContent=Replace(PageContent,"{$CurrPage}",CurrPage)
							  PageContent=Replace(PageContent,"{$TotalPage}",TotalPage)
							  PageContent=Replace(PageContent,"{$ShowThumbList}",ThumbList)
							  PageContent=Replace(PageContent,"{$DefaultImageSrc}",DefaultImageSrc)
							  PageContent=Replace(PageContent,"{$DefaultImageIntro}",DefaultImageIntro)
							  If TotalPage>1 Then F_C=Replace(F_C,"{$GetPictureName}",GetNodeText("title") & "(" & currpage & ")")
							  Templates = "" : Scan F_C
							  F_C = Templates
							  F_C = ReplaceGeneralLabelContent(F_C)
							  F_C = ReplaceRA(F_C, Trim(KS.C_C(Tid,4))) 
							  Call FSOSaveFile(F_C, FilePathAndName)
						 Next
					   case 4 '不分页
					    Dim BigImgSrc,IntroList,BigPic
						For n=1 To TotalPage
						  ThumbPic=Split(ContentArr(n-1),"|")(2) : If lcase(Left(ThumbPic,4))<>"http" Then ThumbPic=KS.Setting(2) & ThumbPic
						  BigPic=Split(ContentArr(n-1),"|")(1) : If lcase(Left(BigPic,4))<>"http" Then BigPic=KS.Setting(2) & BigPic
						  IntroList=IntroList & Split(ContentArr(n-1),"|")(0) &"|"
						  BigImgSrc=BigImgSrc & BigPic &"|"
						  If CurrPage = N Then
						  	ThumbList=ThumbList &"<li><a id=""t" & n & """ class=""currthumb"" href=""javascript:void(0)"" onclick=""showImg(" & n & ");""><img src=""" & ThumbPic &""" border=""0""/></a></li>"
						  Else
						   ThumbList=ThumbList &"<li><a id=""t" & n & """ class=""normalthumb"" href=""javascript:void(0)"" onclick=""showImg(" & n & ");""><img src=""" &ThumbPic &""" border=""0""/></a></li>"
						  End If
						Next
						  DefaultImageSrc=Split(ContentArr(0), "|")(1)
						  DefaultImageIntro=Split(ContentArr(0), "|")(0) 
						  Tp=LFCls.GetConfigFromXML("picturelabel","/labeltemplate/label","style4")
						  Tp=Replace(Tp,"{$TotalPage}",TotalPage)
						  Tp=Replace(Tp,"{$ImgArr}",BigImgSrc)
						  Tp=Replace(Tp,"{$IntroArr}",Replace(Replace(IntroList,"'","\'"),chr(10),"<br/>"))
						  Tp=Replace(Tp,"{$ShowThumbList}",ThumbList)
						  Tp=Replace(Tp,"{$DefaultImageSrc}",DefaultImageSrc)
						  Tp=Replace(Tp,"{$DefaultImageIntro}",DefaultImageIntro)
						  PageContent=Tp
						  F_C = TFileContent
						  If InStr(F_C, "{UnTag:GetRelativeList") <> 0 Or InStr(F_C, "{UnTag:GetLocation") <> 0 Then  F_C = ReplaceLableFlag(Replace(F_C, "{UnTag:", "{Tag:"))
						  Templates = "" : Scan F_C
						  F_C = Templates
						  F_C = ReplaceGeneralLabelContent(F_C)
						  F_C = ReplaceRA(F_C, Trim(KS.C_C(Tid,4))) 
						  Call FSOSaveFile(F_C, FilePathAndName)
					   case 2
							Tp=LFCls.GetConfigFromXML("picturelabel","/labeltemplate/label","style2")
							if ((ubound(ContentArr)+1) mod pagenum)=0 then
									Tpage=(ubound(ContentArr)+1)\pagenum
							else
									Tpage=(ubound(ContentArr)+1)\pagenum + 1
							end if
							For I = 0 To Tpage - 1
								CurrPage = I + 1 : ThumbList=""
								if CurrPage<=1 then  n=0 else n=pagenum*(CurrPage-1)
								For r=1 to pagenum
									  if n<=ubound(ContentArr) Then
									  ThumbPic=Split(ContentArr(n),"|")(2) : If lcase(Left(ThumbPic,4))<>"http" Then ThumbPic=KS.Setting(2) & ThumbPic
									  BigPic=Split(ContentArr(n),"|")(1) : If lcase(Left(BigPic,4))<>"http" Then BigPic=KS.Setting(2) & BigPic
									  ThumbList=ThumbList&"<li><a id="""" href=""" & BigPic & """  class=""highslide"" onclick=""return hs.expand(this)"" title=""""><img alt='" & Split(ContentArr(n), "|")(0) & "' src='" & ThumbPic  & "' style='border:1px #999999 solid' border='0'></a><div style='text-align:center'>" & KS.Gottopic(Split(ContentArr(n), "|")(0),15) & "</div></li>"
									  else 
									   exit for
									  end if
									  n=n+1
								Next
		                    If CurrPage <> 1 Then FilePathAndName = FilePathAndNameTemp & Fname & "_" & CurrPage & FExt
							F_C = TFileContent
							If InStr(F_C, "{UnTag:GetRelativeList") <> 0 Or InStr(F_C, "{UnTag:GetLocation") <> 0 Then  F_C = ReplaceLableFlag(Replace(F_C, "{UnTag:", "{Tag:"))
							
							 PageContent=Replace(Tp,"{$ShowGroupList}",ThumbList)
							 If Tpage>1 Then
							   F_C=Replace(F_C,"{$GetPictureName}",GetNodeText("title") & "(" & currpage & ")")
							   GetPrevNextUrl Tpage,CurrPage,ShowUrl,sFname,Fname,FExt,Tid    '得到上一页及下一页URL
							   PageContent=Replace(PageContent,"{$ShowPage}",GetContentPage(Tpage,CurrPage,ShowUrl,sFname,Fname,FExt))
							 End If
							  Templates = "" : Scan F_C
							  F_C = Templates
							  F_C = ReplaceGeneralLabelContent(F_C)
							  F_C = ReplaceRA(F_C, Trim(KS.C_C(Tid,4))) 
							  Call FSOSaveFile(F_C, FilePathAndName)
						   Next
					   case 3
					        Tp=LFCls.GetConfigFromXML("picturelabel","/labeltemplate/label","style3")
							if ((ubound(ContentArr)+1) mod pagenum)=0 then
									Tpage=(ubound(ContentArr)+1)\pagenum
							else
									Tpage=(ubound(ContentArr)+1)\pagenum + 1
							end if
							For I = 0 To Tpage - 1
								CurrPage = I + 1 : ThumbList=""
								if CurrPage<=1 then  n=0 else n=pagenum*(CurrPage-1)
								   For r=1 to pagenum
										  if n<=ubound(ContentArr) Then	
										   BigPic=Split(ContentArr(n),"|")(1) : If lcase(Left(BigPic,4))<>"http" Then BigPic=KS.Setting(2) & BigPic
							               ThumbList=ThumbList & "<img title=""" &Split(ContentArr(n), "|")(0) & """ class=""scrollLoading"" onclick=""javascript:window.open(this.src)"" alt='" & Split(ContentArr(n), "|")(0)& "' style=""cursor:hand;background:url(" & DomainStr &"images/default/loading.gif) no-repeat center;""  data-url=""" &BigPic &""" src=""" & DomainStr &"images/default/pixel.gif"" border='0'><br/><span>" &Split(ContentArr(n), "|")(0) & "</span>"
										  Else 
										   Exit For
										  End If
										  n=n+1
								   Next
								If CurrPage <> 1 Then FilePathAndName = FilePathAndNameTemp & Fname & "_" & CurrPage & FExt
								F_C = TFileContent
								If InStr(F_C, "{UnTag:GetRelativeList") <> 0 Or InStr(F_C, "{UnTag:GetLocation") <> 0 Then  F_C = ReplaceLableFlag(Replace(F_C, "{UnTag:", "{Tag:"))
								
								 PageContent=Replace(Tp,"{$ShowImgList}",ThumbList)
								 If Tpage>1 Then
								   F_C=Replace(F_C,"{$GetPictureName}",GetNodeText("title") & "(" & currpage & ")")
								   GetPrevNextUrl Tpage,CurrPage,ShowUrl,sFname,Fname,FExt,Tid    '得到上一页及下一页URL
								   PageContent=Replace(PageContent,"{$ShowPage}",GetContentPage(Tpage,CurrPage,ShowUrl,sFname,Fname,FExt))
								 End If
								  Templates = "" : Scan F_C
								  F_C = Templates
								  F_C = ReplaceGeneralLabelContent(F_C)
								  F_C = ReplaceRA(F_C, Trim(KS.C_C(Tid,4))) 
								  Call FSOSaveFile(F_C, FilePathAndName)
						Next
					  End Select
					  
					    
			 case Else   
			 	F_C = TFileContent
				If InStr(F_C, "{UnTag:GetRelativeList") <> 0 Or InStr(F_C, "{UnTag:GetLocation") <> 0 Then  F_C = ReplaceLableFlag(Replace(F_C, "{UnTag:", "{Tag:"))
				Templates = ""
								'供求系统替换权限标签
				If Fcls.ChannelID=8 And Instr(F_C,"[KS_Charge]")<>0 Then
				 Dim ChargeContent:ChargeContent=KS.CutFixContent(F_C, "[KS_Charge]", "[/KS_Charge]", 1)
				 F_C=Replace(F_C,ChargeContent,LFCls.GetConfigFromXML("supply","/labeltemplate/label","divajax"))
				End If

				Scan F_C
				F_C = Templates
				F_C = ReplaceRA(F_C, Trim(KS.C_C(TID,5))) '如果采用根相对路径,则替换绝对路径为根相对路径
				Call FSOSaveFile(F_C, FilePathAndName)
			end select
			
		End Function
		
		Sub GetPrevNextUrl(TotalPage,CurrPage,ShowUrl,sFname,Fname,FExt,Tid)
		    If TotalPage > 1 Then
				 If CurrPage=1 Then
					 NextUrl = ShowUrl & Fname & "_" & (CurrPage + 1) & FExt : PrevUrl="#"
				 ElseIf CurrPage = 2 And CurrPage <> TotalPage Then '对于最后一页刚好是第二页的要做特殊处理
					 NextUrl = ShowUrl & Fname & "_" & (CurrPage + 1) & FExt : PrevUrl = ShowUrl & sFname
				 ElseIf CurrPage = 2 And CurrPage = TotalPage Then
					 NextUrl=KS.GetFolderPath(Tid): PrevUrl = ShowUrl & sFname
				 ElseIf CurrPage = TotalPage Then
					 NextUrl=KS.GetFolderPath(Tid): PrevUrl = ShowUrl & Fname & "_" & (CurrPage - 1) & FExt
				 Else
				     NextUrl = ShowUrl & Fname & "_" & (CurrPage + 1) & FExt : PrevUrl = ShowUrl & Fname & "_" & (CurrPage - 1) & FExt
				 End If
			Else
				NextUrl=KS.GetFolderPath(Tid):PrevUrl="#"
			End If	
		End Sub
		Function GetContentPage(TotalPage,CurrPage,ShowUrl,sFname,Fname,FExt)
		   If TotalPage<=1 Then Exit Function
		   Dim PageStr,StartPage,K,N
		   PageStr = "<div id=""pageNext"" style=""text-align:center""><table align=""center""><tr><td>"
		   If CurrPage > 1 And PrevUrl<>"#" Then PageStr = PageStr & "<a class=""prev"" href=""" & PrevUrl & """>上一页</a> "
		   startpage=1:k=0: if (CurrPage>=10) then startpage=(CurrPage\10-1)*10+CurrPage mod 10+2
		    For N = startpage To TotalPage
				If CurrPage = N Then
					PageStr = PageStr & ("<a class=""curr"" href=""#""><span style=""color:red"">" & N & "</span></a> ")
				Else
					If N=1 Then
					  PageStr = PageStr & ("<a class=""num"" href="""  & ShowUrl & sFname & """>" & N & "</a> ")
					Else
					  PageStr = PageStr & ("<a class=""num"" href=""" &  ShowUrl & Fname & "_" & N & FExt & """>" & N & "</a> ")
				    End If
				End If
				K=K+1 : If K>=10 Then Exit For
			Next

			If CurrPage<>TotalPage Then PageStr = PageStr & "<a class=""next"" href=""" & NextUrl & """>下一页</a>"
			PageStr = PageStr & "</td></tr></table></div>"
			GetContentPage=PageStr
		End Function
		
		
		
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名：RefreshFolder
		'作  用：刷新栏目页面
		'参  数：RS Recordset数据集
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function RefreshFolder(ChannelID,RS)
			Dim F_C, FolderDir, FilePath, Index
			Call FCls.SetClassInfo(RS("ChannelID"),RS("ID"),RS("TN"))
			F_C = LoadTemplate(RS("FolderTemplateID"))
			F_C = ReplaceAllLabel(F_C)
			F_C = ReplaceLableFlag(F_C)   '替换函数标签
            F_C = ReplaceGeneralLabelContent(F_C)          '替换网站通用标签
			 If KS.C_S(ChannelID,44)="1" Or (KS.C_S(ChannelID,44)="3" And Trim(RS("TN")) = "0") Then 
			 Index = RS("FolderFsoIndex")
			 ElseIf KS.C_S(ChannelID,44)="4" Then
             Index=KS.C_S(ChannelID,45) &Mid(Trim(RS("FolderFsoIndex")), InStrRev(Trim(RS("FolderFsoIndex")), ".")) '分离出扩展名
			 Else
             Index=KS.C_S(ChannelID,45) & "_" & rs("classid")&Mid(Trim(RS("FolderFsoIndex")), InStrRev(Trim(RS("FolderFsoIndex")), ".")) '分离出扩展名
			 End If
			 If RS("ClassType")<>"3" Then
			  'If RS("TN")="0" Then
			  '  Index=Split(RS("Folder"),"/")(0) & "/" & RS("FolderFsoIndex")
			  'ELSE
			 Index=Replace(Index,"{$TopClassEname}",Split(RS("Folder"),"/")(0))
			 Index=Replace(Index,"{$ClassEname}",Split(RS("Folder"),"/")(ubound(split(RS("Folder"),"/"))-1))
			 Index=Replace(Index,"{$ClassID}",RS("ClassID"))
			 Index=Replace(Index,"{$BigClassID}",RS("ID"))
			  'End If
			 End If
			 
			 FolderDir = KS.C_S(ChannelID,8)
			 If Left(FolderDir, 1) = "/" Or Left(FolderDir, 1) = "\" Then FolderDir = Right(FolderDir, Len(FolderDir) - 1)
			
			 If KS.C_S(ChannelID,44)="1"  Or RS("ClassType")="3" Then 
			   FilePath = KS.Setting(3) & FolderDir & RS("Folder")
			 ElseIf KS.C_S(ChannelID,44)="2" Or KS.C_S(ChannelID,44)="4" Then
			   FilePath = KS.Setting(3) & FolderDir
			 Else
			   FilePath = KS.Setting(3) & FolderDir & Split(RS("Folder"),"/")(0) & "/"
			 End If
             
			 
			 If RS("ClassType")="3" Then
			  Dim FsoName:FsoName = Mid(FilePath, InStrRev(FilePath, "/")) '分离出扩展名
			  Call KS.CreateListFolder(Replace(FilePath,FsoName,""))
			 Else
				 Dim FsoFolder:FsoFolder=FilePath 
				 If Instr(Index,"/")<>0 Then
				   FsoFolder=FsoFolder & Replace(Trim(Index), Mid(Trim(Index), InStrRev(Trim(Index), "/")), "")
				 End If
			     Call KS.CreateListFolder(FsoFolder)
			 End If
			If (FCls.PageList <> "") Then
			  Call GetPageStr(FCls.PageList, "", Index, F_C, FilePath, Trim(RS("FolderDomain")))
			  FCls.PageList=""
			Else
			 F_C = Replace(F_C, "{PageListStr}", "")
			 F_C = ReplaceRA(F_C, Trim(RS("FolderDomain")))
			 If RS("ClassType")="3" Then Index=""
			 Call FSOSaveFile(F_C, FilePath & Index)
		   End If
		End Function
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名：RefreshSpecials
		'作  用：刷新专题页面
		'参  数：RS Recordset数据集
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function RefreshSpecials(RS)
			Dim F_C, SpecialDir, FilePath,Index,TempStr
			'设置刷新类型,以取得当前导航位置
			Call FCls.SetSpecialInfo(RS("ClassID"),RS("SpecialID"))                       
			'读出专题页对应的模板
			  F_C = LoadTemplate(RS("TemplateID"))
  			  F_C = ReplaceSpecialContent(F_C,RS)
			  F_C = KSLabelReplaceAll(F_C)
			  Index = Trim(RS("FsoSpecialIndex"))
			  SpecialDir = KS.Setting(95)
			  If Left(SpecialDir, 1) = "/" Or Left(SpecialDir, 1) = "\" Then SpecialDir = Right(SpecialDir, Len(SpecialDir) - 1)
			  FilePath = KS.Setting(3) & SpecialDir & RS("SpecialEName") & "/"
			  Call KS.CreateListFolder(FilePath)
			  F_C = ReplaceLableFlag(F_C)                    '替换函数标签
			  If (FCls.PageList <> "") Then
				Call GetPageStr(FCls.PageList, Trim(DomainStr & SpecialDir & RS("SpecialEname") & "/"), Index, F_C, FilePath, "")
				FCls.PageList = ""
			  Else
				   F_C = Replace(F_C, "{PageListStr}", "")
				   Call FSOSaveFile(F_C, FilePath & Index)
			  End If
		End Function
		Function ReplaceSpecialContent(F_C,RS)
		 F_C=Replace(F_C,"{$GetSpecialID}",RS("SpecialID"))
		 F_C=Replace(F_C,"{$GetSpecialName}",RS("SpecialName"))
		 If Not Isnull(RS("PhotoUrl")) And RS("PhotoUrl")<>"" Then
		 F_C=Replace(F_C,"{$GetSpecialPic}","<img src=""" & RS("PhotoUrl") & """ border=""0"">")
		 Else
		 F_C=Replace(F_C,"{$GetSpecialPic}","<img src=""" & DomainStr & "images/nopic.gif"" border=""0"">")
		 End If
		 F_C=Replace(F_C,"{$GetSpecialNote}",RS("SpecialNote"))
		 F_C=Replace(F_C,"{$GetSpecialDate}",RS("SpecialAddDate"))
		 F_C=Replace(F_C,"{$GetSpecialMetaKey}",LFCls.ReplaceDBNull(RS("MetaKey"),""))
		 F_C=Replace(F_C,"{$GetSpecialMetaDescript}",LFCls.ReplaceDBNull(RS("MetaDescript"),""))
		 ReplaceSpecialContent=ReplaceSpecialClass(F_C)
		End Function
		Function ReplaceSpecialClass(F_C)
		 If FCls.RefreshType="Special" Or FCls.RefreshType="ChannelSpecial" Then
		   F_C=Replace(F_C,"{$GetSpecialClassName}",KS.GetSpecialClass(FCls.RefreshFolderID,"classname"))
		   F_C=Replace(F_C,"{$GetSpecialClassURL}",KS.GetFolderSpecialPath(FCls.RefreshFolderID, True))
		 End If
		 ReplaceSpecialClass=F_C
		End Function
		
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名：RefreshSpecialClass
		'作  用：刷新频道专题汇总页
		'参  数：RS Recordset数据集
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function RefreshSpecialClass(RS)
			 Dim F_C, SpecialDir, Index, FilePath
			  FCls.RefreshType = "ChannelSpecial"    
			  FCls.RefreshFolderID = RS("ClassID")
			  FCls.ItemUnit="个"
			 If RS("TemplateID")="" Then
			 RefreshSpecialClass="请先绑定专题分类模板!":exit function
			 Else
			 F_C = LoadTemplate(RS("TemplateID"))
			 End If
			
			 F_C = ReplaceSpecialClass(F_C)  
			 F_C = KSLabelReplaceAll(F_C)
			 
			  SpecialDir = KS.Setting(95)
			  If Left(SpecialDir, 1) = "/" Or Left(SpecialDir, 1) = "\" Then SpecialDir = Right(SpecialDir, Len(SpecialDir) - 1)
			   
			  Index = RS("FsoIndex")
			  FilePath = KS.Setting(3) & SpecialDir & RS("ClassEname") & "/"
			  Call KS.CreateListFolder(FilePath)

			  If (FCls.PageList <> "") Then
				Call GetPageStr(FCls.PageList, Trim(DomainStr & SpecialDir & RS("ClassEname") & "/"), Index, F_C, FilePath, "")
				FCls.PageList=""
			  Else
				F_C = ReplaceRA(F_C, "")
				Call FSOSaveFile(F_C, FilePath & Index)
			 End If
		End Function
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名：RefreshCommonPage
		'作  用：刷新通用页面
		'参  数：RS Recordset数据集
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function RefreshCommonPage(ByVal FileName,FsoFileName)
		  Dim F_C, CommonDir, FilePath
		      F_C = LoadTemplate(FileName)
			  F_C = KSLabelReplaceAll(F_C) 
			  F_C = Replace(Replace(F_C,"{$InfoID}","0"),"{$GetClassID}","0")
			  
			  '如果采用根相对路径,则替换绝对路径为根相对路径
			  F_C = ReplaceRA(F_C, "")
			  CommonDir = Replace(KS.Setting(94), "\", "")
			  If Left(CommonDir, 1) = "/" Then CommonDir = Right(CommonDir, Len(CommonDir) - 1)
			  'FilePath = KS.Setting(3) & CommonDir
			   FilePath=Replace(FsoFileName,Split(FsoFileName,"/")(Ubound(Split(FsoFileName,"/"))),"")
			  
			  Call KS.CreateListFolder(KS.Setting(3) & CommonDir & FilePath)
			  Call FSOSaveFile(F_C, KS.Setting(3) & CommonDir & FsoFileName)
		End Function
		
		'*********************************************************************************************************
		'函数名：ReplaceRA
		'作  用：自动判断系统是否用相对路径或绝对路径并转换
		'参  数：FileContent原文件,FolderDomain 是否有绑定二级域名
		'*********************************************************************************************************
		Function ReplaceRA(F_C, FolderDomain)
		     if instr(FCls.RefreshType,"guest")<>0 then ReplaceRA=F_C : exit function
		     If Lcase(Fcls.RefreshType)="content" Then  F_C=ReplaceSQLLabel(F_C)
			 If CStr(KS.Setting(97)) = "0" Then
				 If FolderDomain <> "" Then
				   F_C = Replace(F_C, FolderDomain, "/")
				 Else
					  If Trim(KS.Setting(3)) = "/" Then
						F_C = Replace(F_C, DomainStr, "/")
					  Else
						F_C = Replace(F_C, Replace(DomainStr, Trim(KS.Setting(3)), ""), "")
					  End If
				End If
			  End If
			F_C=Replace(F_C,"{#GetFullDomain}",KS.Setting(2))
			F_C=ScanAnnex(F_C)
			ReplaceRA = F_C
		End Function
		'*********************************************************************************************************
		'函数名：GetPageStr
		'作  用：取得分页的通用函数
		'参  数：PageContent--分页内容,LinkUrl--链接地址,Index-首页名称
		'        F_C--待保存的文件内容,FilePath---待保存路径,SecondDomain --二级域名
		'*********************************************************************************************************
		Sub GetPageStr(PageContent, LinkUrl, Index, F_C, FilePath, SecondDomain)
			Dim PageStr, FileStr, I, PageContentArr,LoopEnd,TotalPage,Fname,FExt ,LinkUrlFname
			  FExt = Mid(Trim(Index), InStrRev(Trim(Index), ".")) '分离出扩展名
			  Fname = Replace(Trim(Index), FExt, "")              '分离出文件名
			  LinkUrlFname = LinkUrl & Fname
			  Dim HomeLink:HomeLink=LinkUrl & Index
			 
			  If Instr(HomeLink,"/")<>0 Then
			   HomeLink = Mid(HomeLink, InStrRev(HomeLink, "/")+1)   '分离出文件名
			  End If
			  If Instr(LinkUrlFname,"/")<>0 Then
			   LinkUrlFname = Mid(LinkUrlFname, InStrRev(LinkUrlFname, "/")+1)   '分离出文件名
			  End If
			  
			  PageContentArr = Split(PageContent, "{KS:PageList}")
			  TotalPage = FCls.TotalPage
			  If KS.ChkClng(FCls.FsoListNum)<>0 and KS.ChkClng(FCls.FsoListNum)<FCls.TotalPage Then LoopEnd=KS.ChkClng(FCls.FsoListNum) Else LoopEnd=FCls.TotalPage
			  I=0
			  Do While I<LoopEnd
			   I=I+1  

			   '=========以下为分页静态化======================
			   Select Case FCls.PageStyle
			    Case 1
				   If I=1 and I<>TotalPage Then
				   PageStr = "首页  上一页 <a href=""" & LinkUrlFname & "_" & TotalPage -1 & FExt & """>下一页</a>  <a href= """ & LinkUrlFname & "_1" & FExt & """>尾页</a>"
				   ElseIf I=1 And I=TotalPage Then
					PageStr ="首页  上一页 下一页 尾页"
				   ElseIf (I=TotalPage And I <> 2) Then
					 PageStr="<a href=""" &  HomeLink  & """>首页</a>  <a href=""" &  LinkUrlFname  & "_"  &  TotalPage-I+2 & FExt  & """>上一页</a> 下一页  尾页"
				   ElseIf(I = TotalPage And I = 2) Then
					 PageStr="<a href=""" & HomeLink & """>首页</a>  <a href=""" &  HomeLink  & """>上一页</a> 下一页  尾页"
				   ElseIf(I = 2) Then
					 PageStr="<a href=""" &  HomeLink  & """>首页</a>  <a href=""" & HomeLink & """>上一页</a> <a href=""" & LinkUrlFname  & "_" & TotalPage-I & FExt  & """>下一页</a>  <a href= """ &  LinkUrlFname  & "_1" & FExt  & """>尾页</a>"
				   Else
					 PageStr="<a href=""" & HomeLink  & """>首页</a>  <a href=""" & LinkUrlFname & "_" & TotalPage-I+2 & FExt  & """>上一页</a> <a href=""" & LinkUrlFname  & "_" & TotalPage -I & FExt & """>下一页</a>  <a href= """ & LinkUrlFname & "_1" & FExt &""">尾页</a>"
				   End If
			       PageStr="共 <span id=""totalrecord"">" & Fcls.TotalPut & "</span> " & FCls.ItemUnit &"  页次:<span id=""currpage"" style=""color:red""> " & I & "</span>/<span id=""totalpage"">" & TotalPage & "</span>页  <span id=""perpagenum"">" & FCls.PerPageNum & "</span>" & FCls.ItemUnit &"/页 " & PageStr
				Case 2,3
						PageStr="第<span id=""currpage"" style=""color:red"">" &  I  & "</span>页 共<span id=""totalpage"">" & TotalPage & "</span>页 "
						If I=1 Then
						 PageStr=PageStr & "<span style=""font-family:webdings;font-size:14px"">9</span> <span style=""font-family:webdings;font-size:14px"">7</span> "
						ElseIf I=2 Then
						 PageStr=PageStr & "<a href=""" & HomeLink & """ title=""首页""><span style=""font-family:webdings;font-size:14px"">9</span></a> <a href=""" &  HomeLink & """ title=""上一页""><span style=""font-family:webdings;font-size:14px"">7</span></a> "
						Else
						 PageStr=PageStr & "<a href=""" & HomeLink  & """ title=""首页""><span style=""font-family:webdings;font-size:14px"">9</span></a> <a href=""" & LinkUrlFname & "_" &  TotalPage-I+2 & FExt & """ title=""上一页""><span style=""font-family:webdings;font-size:14px"">7</span></a> "
						End If
						
						If FCls.PageStyle=2 Then
						     PageStr=PageStr & " <span id=""pagelist"">"
							 Dim startpage:startpage=1
							 Dim P,n:n=1
							 If I>10 Then  startpage=(I/10-1)*10+(i mod 10)+1
							 for p=startpage to TotalPage
								If P=1 Then
									 If (P=I) Then
										PageStr=PageStr & "<a href=""" & HomeLink & """><font color=""#ff0000"">[" & P & "]</font></a>&nbsp;"
									 else
										PageStr=PageStr & "<a href=""" & HomeLink & """>[" & P & "]</a>&nbsp;"
									 End If
								Else
									  if (p=i) Then
										PageStr=PageStr & "<a href=""" & LinkUrlFname & "_" &  TotalPage-P+1 & FExt & """><font color=""#ff0000"">[" & i & "]</font></a>&nbsp;"
									  else
										PageStr=PageStr & "<a href=""" & LinkUrlFname & "_" & TotalPage-p+1 & FExt  &""">[" & p & "]</a>&nbsp;"
									  end if
								End If
								  n=n+1
								  if n>10 Then Exit For
							 Next
						  PageStr=PageStr & "</span>"
					  End If
					  
					If I=TotalPage Then
					  PageStr=PageStr & "<span style=""font-family:webdings;font-size:14px"">8</span> <span style=""font-family:webdings;font-size:14px"">:</span>"
					Else
					  PageStr=PageStr & "<a href=""" & LinkUrlFname &"_" & TotalPage -I & FExt &""" title=""下一页""><span style=""font-family:webdings;font-size:14px"">8</span></a> <a href=""" & LinkUrlFname & "_1" & FExt  & """><span style=""font-family:webdings;font-size:14px"">:</span></a>"
					End If
			 case 4  '新增样式
				 n=0:startpage=1
				 PageStr="<table border=""0"" align=""right""><tr><td id=""pagelist"">" & vbcrlf
				 If I=2 Then
				  PageStr=PageStr & "<a href=""" &  HomeLink & """ title=""上一页"">"
				 ElseIf I>1 Then
				  pageStr=PageStr & "<a href=""" & LinkUrlFname & "_" &  TotalPage-I+2 & FExt & """ class=""prev"">上一页</a>"
				 End If
				 
				 if (I<>TotalPage) then pageStr=PageStr & "<a href=""" & LinkUrlFname &"_" & TotalPage -I & FExt & """ class=""next"">下一页</a>"
				 pageStr=pageStr & "<a href=""" & HomeLink & """ class=""prev"">首 页</a>"
				 if (I>=7) then startpage=I-5
				 if TotalPage-I<5 Then startpage=TotalPage-10
				 If startpage<=0 Then startpage=1
				 For p=startpage To TotalPage
				    If p= I Then
				     PageStr=PageStr & " <a href=""#"" class=""curr""><font color=red>" & p &"</font></a>"
				    Else
					 If P=1 Then
				     PageStr=PageStr & " <a class=""num"" href=""" & HomeLink &""">" & p &"</a>"
					 Else
				     PageStr=PageStr & " <a class=""num"" href=""" & LinkUrlFname & "_" & TotalPage-p+1 & FExt &""">" & p &"</a>"
					 End If
					End If
					n=n+1
					if n>=10 then exit for
				 Next
				 If TotalPage=1 Then
				 pageStr=pageStr & "<a href=""" & LinkUrlFname & FExt &""" class=""prev"">末页</a>"
				 Else
				 pageStr=pageStr & "<a href=""" & LinkUrlFname & "_1" & FExt &""" class=""prev"">末页</a>"
				 End If
				 pageStr=PageStr & " <span>总共<span id=""totalpage"">" & TotalPage & "</span>页</span></td></tr></table>"
			   End Select
			   
			   
			   If FCls.PageStyle<>4 Then
			   PageStr=PageStr &" 转到：<select id=""turnpage"" size=""1"" onchange=""javascript:window.location=this.options[this.selectedIndex].value""><option value=""1"">1</option></select>"
			   End If
			   Dim FFName:FFName=Fname
			   If Instr(FFName,"/")<>0 Then FFname=mid(FFname, InStrRev(Trim(FFname), "/")+1)
			   PageStr=PageStr & vbcrlf & "<script src=""page" & FCls.RefreshFolderID &".html"" type=""text/javascript"" language=""javascript""></script>"&vbcrlf &"<script language=""javascript"" type=""text/javascript"">pageinfo("&FCls.PageStyle&"," &FCls.PerPageNum &",'"&FExt&"','"&FFName&"');</script>"
			   
			  FileStr = Replace(F_C, "{PageListStr}",  PageContentArr(I-1)& "<div id=""fenye"" class=""plist"" style=""margin-top:6px;text-align:right;"">" & PageStr & "</div>")
			  
			  '===============分页静态化结束=====================================================
			   
			   
			   			  

			   FileStr = ReplaceRA(FileStr, SecondDomain)
			   if (TotalPage-I+1>0) Then
				   Dim TempFilePath
				   If I = 1  Then
					  TempFilePath = FilePath & Index
				   Else
					  TempFilePath = FilePath & Fname & "_" & TotalPage-I+1 & FExt
				   End If
				   Call FSOSaveFile(FileStr, TempFilePath)
               End If
 			  Loop
			   If FCls.RefreshType="Folder" And LoopEnd>5 Then KS.Echo "<script>closeWindow();</script>"
			   FilePath=mid(TempFilePath, 1,InStrRev(Trim(TempFilePath), "/"))
          	   Dim JSStr
			   JSStr="var TotalPage=" & TotalPage & ";"&vbcrlf & "var TotalPut=" & KS.ChkClng(Fcls.TotalPut) & ";" &vbcrlf
               JSStr=JSStr & "document.write(""<script language='javascript' src='" & KS.Setting(2) & KS.Setting(3)&"ks_inc/kesion.page.js'></script>"");"&vbcrlf
			   Call FSOSaveFile(JSStr,FilePath&"page" & FCls.RefreshFolderID & ".html")
		End Sub
			
		
		'*********************************************************************************************************
		'函数名：ReplaceGeneralLabelContent
		'作  用：替换通用标签为内容
		'参  数：FileContent原文件
		'*********************************************************************************************************
		Function ReplaceGeneralLabelContent(F_C)
		 
			    %>
				<!--#include file="modellabel/common.asp"-->
				<%
				Templates=""
				Scan F_C
				ReplaceGeneralLabelContent = Templates
		End Function
		
		Function GetTags(TagType,Num)
		  if not isnumeric(num) then exit function
		  dim sqlstr,sql,i,n,str,turl
		  select case cint(tagtype)
		   case 1:sqlstr="select top 500 keytext,hits from ks_keywords where IsSearch=0 order by hits desc"
		   case 2:sqlstr="select top 500 keytext,hits from ks_keywords where IsSearch=0 order by lastusetime desc,id desc"
		   case 3:sqlstr="select top 500 keytext,hits from ks_keywords where IsSearch=0 order by Adddate desc,id desc"
		   case else 
		    GetTags="":exit function
		  end select
		  
		  dim rs:set rs=conn.execute(sqlstr)
		  if rs.eof then rs.close:set rs=nothing:exit function
		  sql=rs.getrows(-1)
		  rs.close:set rs=nothing
		  for i=0 to ubound(sql,2)
		   if KS.FoundInArr(str,sql(0,i),",")=false then
		    n=n+1:str=str & "," & sql(0,i)
			If KS.Setting(185)="1" Then turl=DomainStr & "tags/" & server.URLEncode(sql(0,i)) Else TUrl=DomainStr & "plus/tags.asp?n=" & server.URLEncode(sql(0,i))
		    gettags=gettags & "<a href=""" & TUrl & """ target=""_blank"" title=""TAG:" & sql(0,i) & "&#10;被搜索了" & SQL(1,I) &"次"">" & sql(0,i) & "</a> "
		   end if
		   if n>=cint(num) then exit for
		  next
		  
		End Function
		'*********************************************************************************************************
		'函数名：GetSiteCountAll
		'作  用：替换网站统计标签为内容
		'参  数：Flag-0总统计，1-文章统计 2-图片统计
		'*********************************************************************************************************
		Function GetSiteCountAll()
		   Dim ChannelTotal: ChannelTotal = Conn.Execute("Select Count(*) From KS_Class Where TN='0'")(0)
		   Dim MemberTotal:MemberTotal=Conn.Execute("Select Count(*) From KS_User")(0)
		   Dim CommentTotal: CommentTotal = Conn.Execute("Select Count(*) From KS_Comment")(0)
		   Dim GuestBookTotal:GuestBookTotal=Conn.Execute("Select Count(ID) From KS_GuestBook")(0)
		   GetSiteCountAll="<div class=""sitetotal"">" & vbcrlf
			  GetSiteCountAll = GetSiteCountAll & "<li>频道总数： " & ChannelTotal & " 个</li>" & vbcrlf
			  dim rsc:set rsc=conn.execute("select channelid,ItemName,Itemunit,channeltable from ks_channel where channelstatus=1 and channelid<>6 And ChannelID<>9 and channelid<>10  and channelid<>11")
			  dim k,sql:sql=rsc.getrows(-1)
			  rsc.close:set rsc=nothing
			  for k=0 to ubound(sql,2)
			  GetSiteCountAll = GetSiteCountAll & "<li>" & sql(1,k) & "总数： " & Conn.Execute("Select Count(id) From " & sql(3,k))(0) & " " & sql(2,k)&"</li>" & vbcrlf
			  next
			  GetSiteCountAll = GetSiteCountAll & "<li>注册会员： " & MemberTotal & " 位</li>" & vbcrlf
			  GetSiteCountAll = GetSiteCountAll & "<li>留言总数： " & GuestBookTotal &" 条</li>" & vbcrlf
			  GetSiteCountAll = GetSiteCountAll & "<li>评论总数： " & CommentTotal & " 条</li>" & vbcrlf
			  GetSiteCountAll = GetSiteCountAll & "<li>在线人数： <script language=""javascript"" src=""" & DomainStr & "plus/online.asp?ID=1""></script> 人</li>" & vbcrlf
		   GetSiteCountAll = GetSiteCountAll & "</div>" & vbcrlf
		End Function
		

		
		
		Function ReplaceKeyTags(KeyStr)
		  Dim I,Turl,K_Arr:K_Arr=Split(KeyStr&"",",")
		  For I=0 To Ubound(K_Arr)
		     If KS.Setting(185)="1" Then turl=KS.GetDomain & "tags/" & server.URLEncode(K_Arr(i)) Else TUrl=KS.GetDomain & "plus/tags.asp?n=" & server.URLEncode(K_Arr(i))
		    ReplaceKeyTags=ReplaceKeyTags & "<a href=""" & Turl & """ target=""_blank"">" & K_Arr(i) & "</a> "
		  Next
		End Function
		'替换画中画广告
		Function ReplaceAD(ByVal Content,ClassID)
		 Dim ShowADTF,CLen,Dir,Width,Height,AdUrl,AdLinkUrl,LC,RC,AdStr,ADType
		 Dim ClassBasicInfo:ClassBasicInfo=KS.C_C(ClassID,6)
		 If ClassBasicInfo="" Then Exit Function
		 Dim AdP:AdP = Split(Split(ClassBasicInfo,"||||")(4),"%ks%")
		 ShowADTF=KS.ChkClng(Adp(0))
		 If ShowADTF=0 Then ReplaceAD=Content:Exit Function
		 Dim Param:Param=Split(AdP(1),",")

		 CLen=KS.ChkClng(Param(0)):Dir=Param(1):Width=KS.ChkClng(Param(2)):Height=KS.ChkClng(Param(3)):AdUrl=Adp(3):AdLinkUrl=Adp(4):ADType=KS.ChkClng(ADP(2))

		 If CLen<>0 Then LC=InterceptString(Content,Clen)
		 RC=Right(Content,Len(Content)-Len(LC))		 		 
               If ADType=2 Then
			     Adstr="<table border=""0"" width="""& Width & """ height=""" & height & """ align="""&Dir&"""><tr><td>" & AdUrl & "</td></tr></table>"
			   Else
                    If Lcase(Right(AdUrl,3))="swf" Then'判断是否Swf图片
						AdStr="<table width=""0"" border=""0"" align="""&Dir&"""><tr><td><object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0""  height=""" & height & """ width="""&width&""" ><param name=""movie"" value="""&AdUrl&"""><param name=""quality"" value=""high""><embed src="""&AdUrl&""" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" height=""" & height & """  width="""&Width&"""></embed></object></td></tr></table>"
					Else
						If AdLinkUrl="" Then AdLinkUrl="#"
						AdStr="<table width=""0"" border=""0"" align="""&Dir&"""><tr><td><a href="""&AdLinkUrl&"""  target=""_blank""><img border=""0"" src="""&AdUrl&""" height=""" & height & """ width="""&Width&"""></a></td></tr></table>"
					End If
				End If	

		 ReplaceAD=LC & AdStr & RC
	   End Function
	   '截取字符串
		Function InterceptString(ByVal txt,length)
			Dim x,y,ii,c,ischines,isascii,tempStr
			length=Cint(length)
			txt=trim(txt):x = len(txt):y = 0
			if x >= 1 then
				for ii = 1 to x
					c=asc(mid(txt,ii,1))
					if  c< 0 or c >255 then
						y = y + 2:ischines=1:isascii=0
					else
						y = y + 1:ischines=0:isascii=1
					end if
					if y >= length then
						if ischines=1 and StrCount(left(trim(txt),ii),"<a")=StrCount(left(trim(txt),ii),"</a>") then
							txt = left(txt,ii) '"字符串限长
							exit for
						else
							if isascii=1 then x=x+1
						end if
					end if
				next
				InterceptString = txt
			else
				InterceptString = ""
			end if
		End Function
		
		'判断字符串出现的次数
		Public Function StrCount(Str,SubStr)        
			Dim iStrCount,iStrStart,iTemp
			iStrCount = 0:iStrStart = 1:iTemp = 0:Str=LCase(Str):SubStr=LCase(SubStr)
			Do While iStrStart < Len(Str)
				iTemp = Instr(iStrStart,Str,SubStr,vbTextCompare)
				If iTemp <=0 Then
					iStrStart = Len(Str)
				Else
					iStrStart = iTemp + Len(SubStr)
					iStrCount = iStrCount + 1
				End If
			Loop
			StrCount = iStrCount
		End Function
		
		
		Sub ReplaceHits(F_C,ChannelID,Id)
			If InStr(F_C, "{$GetHits}") <> 0 Then           '总浏览数
				 F_C = Replace(F_C, "{$GetHits}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?Action=Count&GetFlag=0&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByDay}", "<Script Language=""Javascript"" Src=""" &DomainStr & "item/GetHits.asp?GetFlag=1&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByWeek}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=2&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByMonth}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=3&m=" & ChannelID &"&ID=" & ID & """></Script>")
			ElseIf InStr(F_C, "{$GetHitsByDay}") <> 0 Then  '本日浏览数
				 F_C = Replace(F_C, "{$GetHits}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=0&m=" & ChannelID &"&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByDay}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?Action=Count&GetFlag=1&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByWeek}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=2&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByMonth}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=3&m=" & ChannelID & "&ID=" & ID & """></Script>")
			ElseIf InStr(F_C, "{$GetHitsByWeek}") <> 0 Then '本周浏览数
				 F_C = Replace(F_C, "{$GetHits}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=0&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByDay}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=1&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByWeek}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?Action=Count&GetFlag=2&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByMonth}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=3&m=" & ChannelID & "&ID=" & ID & """></Script>")
			ElseIf InStr(F_C, "{$GetHitsByMonth}") <> 0 Then '本月浏览数
				 F_C = Replace(F_C, "{$GetHits}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=0&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByDay}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=1&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByWeek}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=2&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByMonth}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?Action=Count&GetFlag=3&m=" & ChannelID & "&ID=" & ID & """></Script>")
			End If
		End Sub
		
		
		
		Function GetMoviePagePlay(Param)
		 Dim Str,url
		 dim  MovieUrlsArr:MovieUrlsArr = Split(GetNodeText("movieurls"),"|||")
		 Dim Marr:Marr=Split(MovieUrlsArr(0),"§")
		      IF GetNodeText("serverid")=9999 and lcase(left(marr(1),4)="http") then
			   url=marr(1)
			  ElseIf GetNodeText("serverid")=0  Then
			   url= marr(1)
			  Else
			  url= Conn.Execute("Select Url1 From KS_MediaServer Where ID=" & KS.ChkClng(GetNodeText("serverid")))(0)&marr(1)
			  End If

		select case  lcase(Mid(Trim(Marr(1)), InStrRev(Trim(Marr(1)), ".")))
		 case ".flv"
			  Str="<script type=""text/javascript"">" & vbcrlf & _
			  "var texts='" & GetNodeText("title") &"'" & vbcrlf & _
			  "var files='" & url & "'" & vbcrlf & _
			
			  "document.write('<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0"" width=""" & param(0) & """ height=""" & param(1) & """>');" & vbcrlf & _
			  "document.write('<param name=""movie"" value=""" & domainstr & "editor/plugins/flvPlayer/jwplayer.swf""><param name=""quality"" value=""high"">');" & vbcrlf & _
			  "document.write('<param name=""menu"" value=""false""><param name=""allowFullScreen"" value=""true"" />');" & vbcrlf & _
			  "document.write('<param name=""FlashVars"" value=""file='+files+'&amp;autostart=true"">');" & vbcrlf & _
			  "document.write('<embed src=""" & domainstr & "editor/plugins/flvPlayer/jwplayer.swf"" allowFullScreen=""true"" FlashVars=""file='+files+'&amp;autostart=true"" menu=""false"" quality=""high"" width=""" & Param(0) & """ height=""" & param(1) & """ type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" />'); " & vbcrlf & _
			"document.write('</object>'); " & vbcrlf &_
			"</script>" &vbcrlf
		 case ".rm",".rmvb",".rt",".ra",".rp",".rv"
		    str="<embed src=""" & url & """ type=""audio/x-pn-realaudio-plugin"" console=""Clip1"" controls=""ImageWindow"" height=""" & Param(1) & """ width=""" & Param(0) & """ autostart=""1""><br /><embed width=""" & Param(0) & """ height=""36"" controls=""ControlPanel"" console=""Clip1"" type=""audio/x-pn-realaudio-plugin"" autostart=""1"">"
		 case ".swf"
		   str = "<object id=""Flash1"" classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" width=""" & Param(0) & """ height=""" & Param(1) & """>" & vbCrLf
   str = str & "<param name=""movie"" value=""" & url & """ />" & vbCrLf
		  str = str & "<param name=""quality"" value=""high"" />" & vbCrLf
		  str = str & "<embed src=""" & url & """ quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" width=""" & Param(0) & """ height=""" & Param(1) & """></embed>" & vbCrLf
		  str = str & "</object>" & vbCrLf
		 case else
	        str="<embed width=""" & Param(0) & """ height=""" & Param(1) & """ autostart=""1"" src=""" & url & """>"
		 end select

		 GetMoviePagePlay=str
		End Function
	
	   

        '有效期
		Function GetValidTime(Days)
		 Select Case  KS.ChkClng(Days)
		   Case 3:GetValidTime="三天"
		   Case 7:GetValidTime="一周"
		   Case 15:GetValidTime="半个月"
		   Case 30:GetValidTime="一个月"
		   Case 90:GetValidTime="三个月"
		   Case 180:GetValidTime="半年"
		   Case 365 :GetValidTime="一年"
		   Case Else:GetValidTime="长期"
		 End Select
		End Function
	
	
		Function GetProductType(TypeID)
		 Select Case TypeID
		  Case 1:GetProductType="正常销售"
		  Case 2:GetProductType="涨价销售"
		  Case 3:GetProductType="降价销售"
		End Select
		End Function
	
		'**************************************************
		'函数名：Published
		'作  用：取得发布时间及版权信息
		'参  数：无
		'**************************************************
		'取得Flash播放内容,param=宽，高
		Function GetFlashContent(Param)
		  Dim FlashWidth: FlashWidth = Split(Param, ",")(0)
		  Dim FlashHeight: FlashHeight = Split(Param, ",")(1)
		  Dim FlashUrl: FlashUrl = node.SelectSingleNode("@flashurl").text
		  If LCase(Left(FlashUrl, 5)) <> "http:" Then  FlashUrl = Left(DomainStr, Len(DomainStr) - 1) & FlashUrl
		  If Instr(lcase(flashurl),".flv")<>0 then
		  GetFlashContent=GetFlashFlvPlayer(flashurl,FlashWidth,FlashHeight)
		  else
		  GetFlashContent = "<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" width=""" & FlashWidth & """ height=""" & FlashHeight & """>" & vbCrLf
		  GetFlashContent = GetFlashContent & "<param name=""movie"" value=""" & FlashUrl & """ />" & vbCrLf
		  GetFlashContent = GetFlashContent & "<param name=""quality"" value=""high"" />" & vbCrLf
		  GetFlashContent = GetFlashContent & "<embed src=""" & FlashUrl & """ quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" width=""" & FlashWidth & """ height=""" & FlashHeight & """></embed>" & vbCrLf
		  GetFlashContent = GetFlashContent & "</object>" & vbCrLf
		  end if
		End Function
		'根据参数取得Flash播放器,param=宽，高
		Function GetFlashPlayer(Param)
		  Dim FlashWidth: FlashWidth = Split(Param, ",")(0)
		  Dim FlashHeight: FlashHeight = Split(Param, ",")(1)
		  Dim FlashUrl: FlashUrl = node.SelectSingleNode("@flashurl").text
		  If LCase(Left(FlashUrl, 5)) <> "http:" Then   FlashUrl = Left(DomainStr, Len(DomainStr) - 1) & FlashUrl
		  If Instr(lcase(flashurl),".flv")<>0 then
		   GetFlashPlayer=GetFlashFlvPlayer(flashurl,FlashWidth,FlashHeight)
		  else
		  GetFlashPlayer = LFCls.GetConfigFromXML("Label","/labeltemplate/label","flashplayer")
		  GetFlashPlayer = Replace(GetFlashPlayer,"{$Width}",FlashWidth)
		  GetFlashPlayer = Replace(GetFlashPlayer,"{$Height}",FlashHeight)
		  GetFlashPlayer = Replace(GetFlashPlayer,"{$WebUrl}",DomainStr)
		  GetFlashPlayer = Replace(GetFlashPlayer,"{$FlashUrl}",FlashUrl)
		  end if
		End Function
		Function GetFlashFlvPlayer(flashurl,FlashWidth,FlashHeight)
		      GetFlashFlvPlayer="<script type=""text/javascript"">" & vbcrlf & _
			  "document.write('<embed src=""" & domainstr & "editor/plugins/flvPlayer/jwplayer.swf"" allowFullScreen=""true"" FlashVars=""file=" & flashurl &"&amp;autostart=true"" menu=""false"" quality=""high"" width=""" & FlashWidth & """ height=""" & FlashHeight & """ type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" />'); " & vbcrlf & _
			"</script>" &vbcrlf
		End Function
		
		'按父ID返回二级分类结构
		Function GetCategory(tn)
		    Call KS.LoadClassConfig()
			Dim Node,ClassXML,TreeStr
			Set ClassXML=Application(KS.SiteSN&"_class")
				If IsOBject(ClassXml) Then
				  For Each Node In ClassXML.DocumentElement.SelectNodes("class[@ks27=1][@ks13=" & tn & "]")
				      TreeStr = TreeStr  & "<dl><dt>" & KS.GetClassNP(Node.SelectSingleNode("@ks0").text)&" </dt>"&vbcrlf
					  TreeStr = TreeStr  & SubTreeList(ClassXml,Node.SelectSingleNode("@ks0").text)
				  Next
				End If
		 GetCategory=TreeStr
		End Function
		Function SubTreeList(ClassXml,tn)
			Dim Node,K,TJ,TreeStr
		    If IsOBject(ClassXml) Then
			          TreeStr = TreeStr & "<dd>"
					  For Each Node In ClassXML.DocumentElement.SelectNodes("class[@ks27=1][@ks13='" & TN &"']")
							 TreeStr = TreeStr & KS.GetClassNP(Node.SelectSingleNode("@ks0").text) 
					  Next
					  TreeStr = TreeStr & "</dd></dl>"
			End If
			SubTreeList=TreeStr
		End Function
				'pk
		Function GetPK(PKID)
		   Dim SqlStr
		   If KS.ChkClng(PKID)=0 Then
			sqlStr="select top 1 * from KS_PKZT  Where Status=1 Order By ID Desc"
		   Else
			sqlStr="select top 1 * from KS_PKZT where ID=" & PKID & ""
		   End If
		   Dim RS:Set RS=Conn.Execute(SQLStr)
		   if RS.Eof And RS.BOF Then
		     RS.Close
			 Set RS=NOthing
			 GetPK=""
			 Exit Function
		   End If
		   Dim template:template=LFCls.GetConfigFromXML("pktmp","/labeltemplate/label","tmp")
		   template=replace(template,"{$PKID}",rs("id")&"")
		   template=replace(template,"{$GetPKTitle}",rs("title")&"")
		   template=replace(template,"{$GetZFTips}",rs("zftips")&"")
		   template=replace(template,"{$GetFFTips}",rs("fftips")&"")
		   template=replace(template,"{$ZFNum}",rs("zfvotes")&"")
		   template=replace(template,"{$FFNum}",rs("ffvotes")&"")
		   template=replace(template,"{$NewsLink}",rs("NewsLink")&"")
		   rs.close
		   set rs=nothing
		   GetPK=template
		End Function
		
		'=================================================
		'函数名：GetVote
		'作  用：显示网站调查
		'=================================================
		Function GetVote(VoteID)
			dim sqlVote,rsVote,i,XML,Node
			If KS.ChkClng(VoteID)=0 Then
			sqlVote="select top 1 * from KS_Vote Order By NewestTF Desc"
			Else
			sqlVote="select top 1 * from KS_Vote where ID=" & VoteID & " Order By NewestTF Desc"
			End If
			Set rsVote= conn.execute(sqlvote)
			if rsVote.bof and rsVote.eof then 
				GetVote= "没有任何调查!"
			else
			    VoteID=RSVote("ID")
				GetVote=GetVote & "<div class=""vote"">" & vbcrlf 
				GetVote=GetVote & "<form name='VoteForm" & VoteID &"' method='post' action='" & DomainStr & "?do=Vote&id=" & VoteID&"' target='_blank'>" &vbcrlf
				GetVote=GetVote &  "<input name='Action' type='hidden' value='dovote'>"&vbcrlf
				GetVote=GetVote & "<div class=""votetitle"">"& rsVote("Title") &"</div>"&vbcrlf
				
				Set XML=LFCls.GetXMLFromFile("voteitem/vote_"&VoteID)
				If IsObject(XML) Then
				   if XML.readystate=4 and XML.parseError.errorCode=0 Then 
					if rsVote("VoteType")="Single" then
						for each node in Xml.DocumentElement.SelectNodes("voteitem")
							GetVote=GetVote & "<input type='radio' name='VoteOption' value='"& Node.getAttribute("id") &"'>" & trim(Node.childNodes(0).text) &"<br/>"&vbcrlf
						Next
						
					else
						for each node in Xml.DocumentElement.SelectNodes("voteitem")
							GetVote=GetVote &  "<input type='checkbox' name='VoteOption' value='"& Node.getAttribute("id") &"'>"& trim(Node.childNodes(0).text) &"<br/>"&vbcrlf
						next
					end if
				  else
				      GetVote=GetVote &  "加载投票项出错，请检查config/voteitem/vote_"&VoteID &".xml是否存在!"
				  end if
			    End If
				GetVote=GetVote &  "<input name=""VoteType"" type=""hidden"" value="""& rsVote("VoteType") &""">"&vbcrlf
				GetVote=GetVote &  "<input name='ID' type='hidden' value='"& rsVote("ID") &"'>"&vbcrlf
				GetVote=GetVote &  "<div style=""clear:both; margin-top:10px; text-align:center"">"&vbcrlf
				GetVote=GetVote &  "<input type='image' src='" & domainStr & "Images/Default/voteSubmit.jpg' border='0' align=""absmiddle"">&nbsp;"&vbcrlf
				GetVote=GetVote &  "<a href='" & domainStr & "index.asp?do=vote&id=" & VoteID &"' target='_blank'><img align='absmiddle' src='" & domainStr & "Images/Default/voteView.jpg' border='0' align=""absmiddle""></a>"&vbcrlf
				GetVote=GetVote &  "</div></form>"&vbcrlf
				GetVote=GetVote & "</div>"&vbcrlf
			end if
			rsVote.close:set rsVote=nothing
		End Function
		'显示会员登录排行
		Sub GetTopUser(Num,MoreStr)
		 Dim Sql,XML,Node,UserFace,UserName
		 Dim RSObj:Set RSObj=Conn.execute("Select Top " & Num &" UserID,UserName,UserFace,LoginTimes,sex From KS_User where groupid<>1 Order BY LoginTimes Desc,UserID Desc")
		 If Not RSObj.Eof Then Set Xml=KS.RsToXml(RSObj,"row","")
		 RSObj.Close : Set RSObj = Nothing
		 If IsObject(Xml) Then
			For each Node In Xml.DocumentElement.SelectNodes("row")
			  userface=Node.SelectSingleNode("@userface").text  : UserName=Node.SelectSingleNode("@username").text
			  if userface="" then
			   userface="user/images/noavatar_small.gif"
			  End If
			  If Left(Lcase(userface),4)<>"http" and Left(userface,1)<>"/" Then userface=KS.GetDomain & userface
			  echoln "<li><a href=""" & KS.GetSpaceUrl(Node.SelectSingleNode("@userid").text) & """ target=""_blank"" class=""b""><img src=""" & userface & """ onerror=""this.src='" & KS.GetDomain & "user/images/noavatar_small.gif';"" border=""0"" alt=""用户:" & UserName & "&#13;&#10;登录:" & Node.SelectSingleNode("@logintimes").text & "次&#13;&#10;性别:" & Node.SelectSingleNode("@sex").text & """/></a><a class=""u"" href=""" & KS.GetSpaceUrl(Node.SelectSingleNode("@userid").text) & """ target=""_blank"">" & UserName & "</a></li>"
			Next
			If MoreStr<>"" Then Echo "<div style=""text-align:center""><a href=""" & KS.GetDomain & "user/userlist.asp"" target=""_blank"">" & MoreStr & "</a></div>"
			Xml=Empty : Set Node=Nothing
		 End If
		End Sub
		
		'显示会员动态
		Sub GetUserDynamic(Num)
		 Dim RS,XML,Node
		 Set RS=Conn.Execute("Select Top " & Num & " id,userid,username,Note,adddate From KS_UserLog Order By Id Desc")
		 If Not RS.Eof Then Set XML=KS.RsToXml(RS,"row","")
		  RS.Close:Set RS=Nothing
		 If IsObject(XML) Then
		  for each Node In XML.DocumentElement.SelectNodes("row")
		    echoln "<li><a href=""" & KS.GetDomain & "user/weibo.asp?userid=" & Node.SelectSingleNode("@userid").text & """ target=""_blank"">" & Node.SelectSingleNode("@username").text & "</a> " & Replace(Replace(Replace(Replace(ubbcode(Node.SelectSingleNode("@note").text,1),"{$GetSiteUrl}",KS.GetDomain),vbcrlf,""),"<p>",""),"</p>","") & "&nbsp;"& KS.GetTimeFormat(Node.SelectSingleNode("@adddate").text) & "</li>"
		  next
		 XML=Empty : Set Node=Nothing
		 End If
		End Sub
		
		
				
		Function FormatImglink(content,url,totalpage)
          dim re:Set re=new RegExp
           re.IgnoreCase =true
           re.Global=True
		   '去除onclick,onload等脚本 
            're.Pattern = "\s[on].+?=([\""|\'])(.*?)\1" 
            'Content = re.Replace(Content, "") 
			Dim LinkStr
		    If TotalPage=1 Then
			 LinkStr="href=""javascript:;"" onclick=""showimg('$2');"""
			Else
			 LinkStr="href=""" & Url & """"
			End If
			
		   '将SRC不带引号的图片地址加上引号 
            re.Pattern = "<img.*?\ssrc=([^\""\'\s][^\""\'\s>]*).*?>" 
            Content = re.Replace(Content, "<a " & LinkStr & "><img src=""$2"" alt=""" & GetNodeText("title") & """  onmousewheel=""return bbimg(this)"" onload=""javascript:resizepic(this)"" border=""0""/></a>") 
		   '正则匹配图片SRC地址 
		   re.Pattern = "<img.*?\ssrc=([\""\'])([^\""\']+?)\1.*?>" 
           Content = re.Replace(Content, "<a " & LinkStr & "><img src=""$2"" alt=""" & GetNodeText("title") & """ onmousewheel=""return bbimg(this)"" onload=""javascript:resizepic(this)"" border=""0""/></a>") 

		  set re = nothing
          FormatImglink = content
		end function 
End Class
%> 
