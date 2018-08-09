<!--#include file="Kesion.MemberCls.asp"-->
<%


Dim StaticCls
Set StaticCls=New KesionStaticCls
Class KesionStaticCls
        Private KS,KSUser, KSR,QueryParams,ChannelID,ThreadType,G_P_Arr
		Private FileContent,RS,SqlStr,Content,InfoPurview,ClassPurview,ReadPoint,ChargeType,PitchTime,ReadTimes
		Private DomainStr,ID,UserLoginTF,CurrPage,PayTF,UserName,UrlsTF
		Private ModelChargeType,ChargeTableName,DateField,ChargeStr,ChargeStrUnit,CurrPoint,IncomeOrPayOut  
        Private PreListTag,PreContentTag,Extension
		Private DocXML
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		  Set KSR = New Refresh
		  DomainStr=KS.GetDomain
		  PreContentTag=GCls.StaticPreContent
		  PreListTag=GCls.StaticPreList
		  Extension=GCls.StaticExtension
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing:Set KSUser=Nothing
		End Sub
		Public Sub Run()
		   ChannelID=KS.ChkClng(KS.S("M"))
		   ID=KS.ChkClng(KS.S("D")) : If ID=0 Then ID=KS.ChkClng(KS.S("ID"))
		   If ChannelID<>0 And ID<>0 Then
		     if KS.C_S(ChannelID,48)=1 Then 
			  Response.Redirect (KS.Setting(3) & "?" & PreContentTag & "-" & ID & "-" & ChannelID & Extension)
			 end if
			 PayTF=KS.ChkClng(KS.S("Pt"))
			 CurrPage=KS.ChkClng(KS.S("P"))
			 If CurrPage<=0 Then CurrPage=1
			 Call StaticContent()
		   ElseIf ID<>0 Then
		     CurrPage=KS.ChkClng(KS.S("Page")): If CurrPage<=0 Then CurrPage=1
		     Call StaticList()
		   Else
			   QueryParams=Replace(Lcase(Request.ServerVariables("QUERY_STRING")),Extension,"")
			   G_P_Arr=Split(QueryParams,"-")
			   If Ubound(G_P_Arr)<1 Then 
				 Response.Redirect("index.asp")
				 Response.End()
			   End If
			   ThreadType=G_P_Arr(0)
		   
			   ID=KS.ChkClng(G_P_Arr(1))
			   If ID=0 Then 
				 Response.Redirect("index.asp")
				 Response.End()
			   End If
			  
			   If ThreadType=PreContentTag Then
				   ChannelID=KS.ChkClng(G_P_Arr(2))
				   If ChannelID=0 Then  Response.Redirect("index.asp"): Response.End()
	
				 If Ubound(G_P_Arr)>=3 Then  CurrPage=KS.ChkClng(G_P_Arr(3))  Else  CurrPage=1
				 If Ubound(G_P_Arr)>=4 Then  PayTF=G_P_Arr(4) 
				 If CurrPage<=0 Then CurrPage=CurrPage+1
				 
				 Call StaticContent()
			   ElseIf ThreadType=PreListTag Then
				 If Ubound(G_P_Arr)>=2 Then  CurrPage=KS.ChkClng(G_P_Arr(2))  Else  CurrPage=1
				 If CurrPage<=0 Then CurrPage=CurrPage+1
				 Call StaticList()
			   End If
		  End If
		End Sub
		'静态化列表
		Sub StaticList()
		 UserLoginTF=Cbool(KSUser.UserLoginChecked)
		 Dim RSObj
		  Set RSObj=Conn.Execute("Select top 1 ID,ClassPurview,TN,FolderTemplateID,FolderDomain,DefaultArrGroupID,ChannelID From KS_Class Where ClassID=" & ID)
		 IF RSObj.Eof And RSObj.Bof Then  RSObj.Close:Set RSObj=Nothing:Call KS.Alert("非法参数!",""):Exit Sub

		  If RSObj("ClassPurview")=2 and  RSObj("channelid")<>8 Then
		    If Cbool(KSUser.UserLoginChecked)=false Then 
			 Call KS.Alert("本栏目为认证栏目，至少要求本站的注册会员才能浏览!",KS.GetDomain & "user/login/"):Response.End
		    elseIF KS.FoundInArr(RSObj("DefaultArrGroupID"),KSUser.GroupID,",")=false Then
		     Call KS.Alert("对不起，你所在的用户级没有权限浏览!",Request.ServerVariables("http_referer")):Response.End
		    End If
		  End If
		  	 ChannelID=RSObj("ChannelID")
			 Dim BigID:BigID=RSObj("ID")
			 If KS.ChkClng(KS.C_S(Channelid,7))=1 Then RSObj.Close:Set RSOBj=Nothing : Response.Redirect(KS.GetFolderPath(BigID))
		     Call FCls.SetClassInfo(ChannelID,RSObj("ID"),RSObj("TN"))
               
			 FileContent = KSR.KSLabelReplaceAll(KSR.LoadTemplate(RSObj("FolderTemplateID")))
			Dim LabelParamStr:LabelParamStr=Application("PageParam")

			If Not KS.IsNul(LabelParamStr) And Instr(FileContent,"{KS:PageList}")=0 Then
				 Dim XMLDoc,XMLSql,LabelStyle,KMRFOBJ
				 Dim ParamNode,IncludeSubClass,ModelID,OrderStr,PrintType,PageStyle,PicStyle,ShowPicFlag,FieldStr,Param
				 Dim PerPageNumber,TotalPut,PageNum,TempStr,TableName
				 Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				 If XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
					 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
					 ModelID         = ParamNode.getAttribute("modelid") : If Not IsNumeric(ModelID) Then ModelID=1
					 IncludeSubClass = ParamNode.getAttribute("includesubclass"):If KS.IsNul(IncludeSubClass) Then IncludeSubClass=true 
					 PrintType       = ParamNode.getAttribute("printtype") : If Not IsNumeric(PrintType) Then PrintType=1
					 PageStyle       = ParamNode.getAttribute("pagestyle") : If PageStyle="" Or IsNull(PageStyle) Then PageStyle=1
					 PicStyle        = ParamNode.getAttribute("picstyle")
					 OrderStr        = ParamNode.getAttribute("orderstr") : If OrderStr="" Or IsNull(OrderStr) Then OrderStr="ID Desc"
					 ShowPicFlag     = ParamNode.getAttribute("showpicflag") : If ShowPicFlag="" Or IsNull(ShowPicFlag) Then ShowPicFlag=false
					 PerPageNumber   = ParamNode.getAttribute("num") : If Not IsNumeric(PerPageNumber) Then PerPageNumber=10
					 
					 Param = " Where I.Verific=1 And I.DelTF=0"
					 If CBool(IncludeSubClass) = True Then 
					 Param= Param & " And I.Tid In (" & KS.GetFolderTid(RSObj("ID")) & ")" 
					 Else 
					 Param= Param & " And I.Tid='" & RSObj("ID") & "'"
					 End If
					 
					 Set KMRFObj= New RefreshFunction
					 Set KMRFObj.ParamNode=ParamNode
				     Call KMRFObj.LoadField(ChannelID,PrintType,PicStyle,ShowPicFlag,FieldStr,TableName,Param)
				
					If Lcase(Left(Trim(OrderStr),2))<>"id" Then  OrderStr=OrderStr & ",I.ID Desc"			
					SqlStr = "SELECT " & FieldStr & " FROM " & KS.C_S(ChannelID,2) & " I " & Param & " ORDER BY I.IsTop Desc," & OrderStr
					Set RS=Server.CreateObject("ADODB.RECORDSET")
					RS.Open SqlStr, Conn, 1, 1
					If RS.EOF And RS.BOF Then
						TempStr = "<p>此栏目下没有" & KS.C_S(ChannelID,3) & "</p>"
					Else
						PerPageNumber=cint(PerPageNumber)
						TotalPut = Conn.Execute("select Count(id) from " & KS.C_S(ChannelID,2) & " I " & Param)(0)
						if (TotalPut mod PerPageNumber)=0 then
								PageNum = TotalPut \ PerPageNumber
						else
								PageNum = TotalPut \ PerPageNumber + 1
						end if
						If CurrPage >1 and (CurrPage - 1) * PerPageNumber < totalPut Then
							RS.Move (CurrPage - 1) * PerPageNumber
						Else
							CurrPage = 1
						End If
						Set XMLSQL=KS.ArrayToXml(RS.GetRows(PerPageNumber),RS,"row","root")
						Call KMRFObj.LoadPageParam(XMLSQL,ParamNode,ChannelID)
						LabelStyle=Application("LabelStyle")
						TempStr = KMRFObj.ExplainGerericListLabelBody(LabelStyle)
						XMLSql=Empty
						
						FCls.PageStyle=PageStyle       '分页样式
						FCls.TotalPage=PageNum         '总页数
						TempStr = TempStr & KS.GetPrePageList(FCls.PageStyle,KS.C_S(ChannelID,4),FCls.TotalPage,CurrPage,TotalPut,PerPageNumber) & "{KS:PageList}" 
						
					End If
				
					RS.Close:Set RS=Nothing					
					XMLDoc= Empty : Set ParamNode=Nothing
				End If	
				
			End If
			
			FileContent=Replace(FileContent,"{Tag:Page}",TempStr)
			
			If Instr(FileContent,"{KS:PageList}")<>0 Then
			  If KS.C_S(ChannelID,48)=0 Then
			   FileContent=Replace(FileContent,"{KS:PageList}",KS.GetPageList("?ID=" & ID,FCls.PageStyle,CurrPage,FCls.TotalPage, True))
			  ElseIf KS.C_S(ChannelID,48)=2 Then
			   FileContent=Replace(FileContent,"{KS:PageList}",KS.GetStaticPageList(GCls.StaticPreList & "-" & ID & "-",FCls.PageStyle,CurrPage,FCls.TotalPage,true,GCls.StaticExtension)&"</div>") 
			  Else
			   FileContent=Replace(FileContent,"{KS:PageList}",KS.GetStaticPageList("?" & GCls.StaticPreList & "-" & ID & "-",FCls.PageStyle,CurrPage,FCls.TotalPage,true,GCls.StaticExtension)& "</div>")
			  End If
			End If
			 

		 RSObj.Close:Set RSObj=Nothing
		 Set KMRFObj=Nothing
		 KS.Echo FileContent
		End Sub
		
		
		'静态化内容页
		Sub StaticContent()
		  UserLoginTF=Cbool(KSUser.UserLoginChecked)
		  Select Case (KS.C_S(Channelid,6))
		   Case 1 Call StaticArticleContent()
		   Case 2 Call StaticPhotoContent()
		   Case 3 Call StaticDownContent()
		   Case 4 Call StaticFlashContent()
		   Case 5 Call StaticProductContent()
		   Case 7 Call StaticMovieContent()
		   Case 8 Call StaticSupplyContent()
		  End Select
		End Sub
		
		Function GetPageStr(Page)
		 If KS.C_S(ChannelID,48)=0 Then
		  GetPageStr="?m=" & ChannelID & "&d="& ID & "&p="&Page
		 ElseIf KS.C_S(ChannelID,48)=2 Then
		  GetPageStr=KS.Setting(3) & PreContentTag & "-" & ID & "-" & ChannelID & "-" & Page & Extension
		 Else
		  GetPageStr=KS.Setting(3) & "?" & PreContentTag & "-" & ID & "-" & ChannelID & "-" & Page & Extension
		 End If
		End Function
		
		
		Sub GetRecords()
			    Set RS=Conn.Execute("Select top 1 a.*,ClassPurview,DefaultArrGroupID,DefaultReadPoint,DefaultChargeType,DefaultPitchTime,DefaultReadTimes From " & KS.C_S(ChannelID,2) & " a inner join KS_Class b on a.tid=b.id Where a.ID=" & ID)
		End Sub
		
		'检查收费及权限
		Sub CheckCharge(KSR)
		 With KSR
		     InfoPurview = Cint(.Node.SelectSingleNode("@infopurview").text)
			 ReadPoint   = Cint(.Node.SelectSingleNode("@readpoint").text)
			 ChargeType  = Cint(.Node.SelectSingleNode("@chargetype").text)
			 PitchTime   = Cint(.Node.SelectSingleNode("@pitchtime").text)
			 ReadTimes   = Cint(.Node.SelectSingleNode("@readtimes").text)
			 ClassPurview= Cint(.Node.SelectSingleNode("@classpurview").text)
			 UserName    = .Node.SelectSingleNode("@inputer").text
			 If InfoPurview=2 or ReadPoint>0 Then
			   IF UserLoginTF=false Then
				 Call GetNoLoginInfo(Content)
			   Else
					 IF KS.FoundInArr(.Node.SelectSingleNode("@arrgroupid").text,KSUser.GroupID,",")=false and readpoint=0 Then
					   Content="<div style=""text-align:center"">对不起，你所在的用户组没有查看本" & KS.C_S(ChannelID,3) & "的权限!</div>"
					 Else
						  Call PayPointProcess()
					 End If
			   End If
			  ElseIF InfoPurview=0 And (ClassPurview=1 or ClassPurview=2) Then 
				  If UserLoginTF=false Then
					Call GetNoLoginInfo(Content)
				  Else     
					 '============继承栏目收费设置时,读取栏目收费配置===========
					 ReadPoint  = Cint(.Node.SelectSingleNode("@defaultreadpoint").text)   
					 ChargeType = Cint(.Node.SelectSingleNode("@defaultchargetype").text)
					 PitchTime  = Cint(.Node.SelectSingleNode("@defaultpitchtime").text)
					 ReadTimes  = Cint(.Node.SelectSingleNode("@defaultreadtimes").text)
					 '============================================================
					 If ClassPurview=2 Then
						 IF KS.FoundInArr(.Node.SelectSingleNode("@defaultarrgroupid").text,KSUser.GroupID,",")=false Then
							Content="<div style=""text-align:center"">对不起，你所在的用户组没有查看的权限!</div>"
						 Else
							Call PayPointProcess()
						 End If
					Else    
					 Call PayPointProcess()
					End If
				  End If
			 Else
			   Call PayPointProcess()
			 End If
		 End With
		End Sub
		
		Sub StaticArticleContent()
		 Call GetRecords()
		 IF RS.Eof And RS.Bof Then
		  RS.Close:Set RS=Nothing
		  KS.ShowTips "error","您要查看的" & KS.C_S(ChannelID,3) & "已删除。或是您非法传递注入参数!"
		 ElseIF Cint(RS("Changes"))=1 Then 
		   Dim ClassID:ClassID=RS("Tid")
		   Dim Fname:Fname=RS("articlecontent")
		   RS.Close:Set RS=Nothing
		   Response.Redirect Fname
		 End IF
		  Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
		  With KSR 
			 Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		      .Tid=.Node.SelectSingleNode("@tid").text
			 If .Node.SelectSingleNode("@verific").text<>1 And UserLoginTF=False And KSUser.UserName<>.Node.SelectSingleNode("@inputer").text Then
			   KS.ShowTips "error","对不起，该" & KS.C_S(ChannelID,3) & "还没有通过审核!"
			 End If
			 Call FCls.SetContentInfo(ChannelID,.Tid,ID,.Node.SelectSingleNode("@title").text)

			 Call CheckCharge(KSR)
		   
		   '无需权限的文章转到静态地址上
		   If .Node.SelectSingleNode("@refreshtf").text="1" and Cint(KS.C_S(ChannelID,7))<>0 and not (readpoint>0) and not (infopurview=2) and not (InfoPurview=0 And ClassPurview>0) Then
		     Response.Redirect KS.GetItemUrl(ChannelID,.Tid,ID,.Node.SelectSingleNode("@fname").text)
		   End If
		   
			
		 FileContent = KSR.LoadTemplate(.Node.SelectSingleNode("@templateid").text)
		 If InStr(FileContent,"[KS_Charge]")=0 Then
		   FileContent = Replace(FileContent,"{$GetArticleContent}","[KS_Charge]{$GetArticleContent}[/KS_Charge]")
		 End If
		 Dim ContentArr:

		' Dim Content:Content=(.Node.SelectSingleNode("@articlecontent").text)
		 If .Node.SelectSingleNode("@postid").text<>"0" Then
		  ContentArr=Split(.UbbCode(.Node.SelectSingleNode("@articlecontent").text,1),"[NextPage]")
		 Else
		  ContentArr=Split(.Node.SelectSingleNode("@articlecontent").text,"[NextPage]")
		 End If
		 'If KS.IsNul(Content) Then Content=" "
		 Dim TotalPage,N,K,PageStr,NextUrl,PrevUrl
			TotalPage = Cint(UBound(ContentArr) + 1)
			   If TotalPage > 1 Then
					   If CurrPage = 1 Then
					     PrevUrl="" : NextUrl=GetPageStr(CurrPage + 1)
					   ElseIf CurrPage = TotalPage Then
					     NextUrl = KS.GetFolderPath(.Tid) : PrevUrl = GetPageStr(CurrPage - 1)
					   Else
					     NextUrl = GetPageStr(CurrPage + 1) :PrevUrl = GetPageStr(CurrPage - 1)
					   End If
					   PageStr =  "<div id=""pageNext"" style=""text-align:center""><table align=""center""><tr><td>"
					   If CurrPage > 1 And PrevUrl<>"" Then PageStr = PageStr & "<a class=""prev"" href=""" & PrevUrl & """>上一页</a> "
					 Dim StartPage:StartPage=1
					 if (CurrPage>=10) then StartPage=(CurrPage\10-1)*10+CurrPage mod 10+2
				     For N = StartPage To TotalPage
						 If CurrPage = N Then
						  PageStr = PageStr & ("<a class=""curr"" href=""#""><span style=""color:red"">" & N & "</span></a> ")
						 Else
						  PageStr = PageStr & ("<a class=""num"" href=""" & GetPageStr(N) & """>" & N & "</a> ")
						 End If
						 K=K+1
						 If K>=10 Then Exit For
					 Next
					 PageStr = "<div id=""MyContent"">" & ContentArr(CurrPage-1) & "</div>" & PageStr 
					 If CurrPage<>TotalPage Then PageStr = PageStr & " <a class=""next"" href=""" & NextUrl & """>下一页</a>"
					 PageStr = PageStr & "</td></tr></table></div>"

					 Dim PageTitleArr,PageTitle
					 PageTitle=	.Node.SelectSingleNode("@pagetitle").text
					 
					 If Not KS.IsNul(PageTitle) Then
					  PageTitleArr=Split(PageTitle,"§")
					  If CurrPage-1<=Ubound(PageTitleArr) Then FileContent=Replace(FileContent,"{$GetArticleTitle}",PageTitleArr(CurrPage-1))
					 ElseIF Currpage>0 Then
					   FileContent=Replace(FileContent,"{$GetArticleTitle}",.Node.SelectSingleNode("@title").text & "(" & currpage & ")")
					 End IF
				 Else
				  NextUrl=KS.GetFolderPath(.Tid)
				  PageStr = "<div id=""MyContent"">" & ContentArr(0) & "</div>"
				 End If

				 .ModelID = ChannelID
				 .ItemID  = ID
				 .PageContent=PageStr
				 .NextUrl=NextUrl
				 .TotalPage=TotalPage
				 .Templates=""
				 .Scan FileContent
		 		 FileContent = .Templates
				 				 					 
		  If Content<>"True" Then
		   Dim ChargeContent:ChargeContent=KS.CutFixContent(FileContent, "[KS_Charge]", "[/KS_Charge]", 0)
		   FileContent=Replace(FileContent,"[KS_Charge]" & ChargeContent &"[/KS_Charge]",Content)
		  Else
		   FileContent=Replace(Replace(FileContent,"[KS_Charge]",""),"[/KS_Charge]","")
		  End If
		  If Instr(FileContent,"[KS_ShowIntro]")<>0 Then
			  If CurrPage=1 Then
		        FileContent=Replace(Replace(FileContent,"[KS_ShowIntro]",""),"[/KS_ShowIntro]","")
			  Else
		        FileContent=Replace(FileContent,KS.CutFixContent(FileContent, "[KS_ShowIntro]", "[/KS_ShowIntro]", 1),"")
			  End If
		  End If

		  FileContent = .KSLabelReplaceAll(FileContent)
		End With
          FileContent=Replace(Replace(Replace(Replace(FileContent,"{§","{$"),"{#LB","{LB"),"{#SQL","{SQL"),"{#=","{=")
		  KS.Echo FileContent
		 
	   End Sub
	   
	   Sub StaticPhotoContent()
		 Call GetRecords()
		 IF RS.Eof And RS.Bof Then
		  RS.Close : Set RS=Nothing
		  KS.ShowTips "error","对不起,您要查看的" & KS.C_S(ChannelID,3) & "已删除。或是您非法传递注入参数!"
		 End IF
		 Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
		 With KSR 
			Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		    .Tid=.Node.SelectSingleNode("@tid").text

		 If .Node.SelectSingleNode("@verific").text<>1 And UserLoginTF=False And KSUser.UserName<>.Node.SelectSingleNode("@inputer").text Then
		   KS.ShowTips "error","对不起，该" & KS.C_S(ChannelID,3) & "还没有通过审核!"
		   Response.End
		 End If
		 Call FCls.SetContentInfo(ChannelID,.Tid,ID,.Node.SelectSingleNode("@title").text)
         Dim ShowStyle,PageNum
		 PageNum     = KS.ChkClng(.Node.SelectSingleNode("@pagenum").text) : If PageNum=0 Then PageNum=10
		 ShowStyle   = KS.ChkClng(.Node.SelectSingleNode("@showstyle").text) : If ShowStyle=0 Then ShowStyle=1
		 
		 
		  Call CheckCharge(KSR)
		  
		   '无需权限的转到静态地址上
		   If .Node.SelectSingleNode("@refreshtf").text="1" and Cint(KS.C_S(ChannelID,7))<>0 and not (readpoint>0) and not (infopurview=2) and not (InfoPurview=0 And ClassPurview>0) Then
		     Response.Redirect KS.GetItemUrl(ChannelID,.Tid,ID,.Node.SelectSingleNode("@fname").text)
		   End If

		  
		 	Dim KSLabel:Set KSLabel =New RefreshFunction
			FileContent = KSR.LoadTemplate(.Node.SelectSingleNode("@templateid").text)
			 Dim PicUrlsArr,N,PageStr,TotalPage,NextUrl,Tp
			 If Cbool(UrlsTF)=true Then
				  PicUrlsArr = Split(Content, "|||")
				  TotalPage = Cint(UBound(PicUrlsArr) + 1)
				  If (ShowStyle=1 or ShowStyle=2 Or ShowStyle=4) And TotalPage=1 Then ShowStyle=3
				  Dim r,c,str,TPage,thumbsphoto
				  Select Case ShowStyle
				   Case 5
				         Tp=LFCls.GetConfigFromXML("picturelabel","/labeltemplate/label","style5")
						 For n=1 To TotalPage
						   if n=1 then DefaultImageSrc=Split(PicUrlsArr(n-1),"|")(1)
						   ThumbList=ThumbList & "<LI class=""sel"" onclick='picchang1(""" & Split(PicUrlsArr(n-1),"|")(1) & """,""" & Split(PicUrlsArr(n-1),"|")(1) & """)'><IMG width=""119"" height=""90"" src=""" & Split(PicUrlsArr(n-1),"|")(2) & """><LI>" &vbcrlf
						Next
						Tp=Replace(Tp,"{$ShowThumbList}",ThumbList)
						Tp=Replace(Tp,"{$DefaultImageSrc}",DefaultImageSrc)
						
				   Case 2
						 Tp=LFCls.GetConfigFromXML("picturelabel","/labeltemplate/label","style2")
					    if ((ubound(PicUrlsArr)+1) mod pagenum)=0 then
							Tpage=(ubound(PicUrlsArr)+1)\pagenum
						else
							Tpage=(ubound(PicUrlsArr)+1)\pagenum + 1
						end if
						If CurrPage>Tpage Then CurrPage=Tpage
						 if CurrPage<=1 then  n=0 else n=pagenum*(CurrPage-1)
						For r=1 to pagenum
							  if n<=ubound(PicUrlsArr) Then
							  thumbsphoto=thumbsphoto&"<li><a id="""" href=""" & Split(PicUrlsArr(n), "|")(1) & """  class=""highslide"" onclick=""return hs.expand(this)"" title=""""><img alt='" & KS.LoseHtml(Split(PicUrlsArr(n), "|")(0)) & "' src='" & Split(PicUrlsArr(n), "|")(2)  & "' style='border:1px #999999 solid' border='0'></a><div style='text-align:center'>" & KS.Gottopic(Split(PicUrlsArr(n), "|")(0),18) & "</div></li>"
                              else 
							   exit for
							  end if
							  n=n+1
						Next
					 Tp=Replace(Tp,"{$ShowGroupList}",thumbsphoto)
					 Tp=Replace(Tp,"{$ShowPage}",GetPicturePage(TPage,CurrPage))
				Case 1  '按分页方式生成图片内容页
				       Dim PrevUrl,ThumbList,DefaultImageSrc,DefaultImageIntro
					   Tpage=TotalPage : If CurrPage>Tpage Then CurrPage=Tpage
					   If TotalPage > 1 Then
							If CurrPage = 1 Then
							  PrevUrl="#" : NextUrl=GetPageStr(CurrPage+1)
							ElseIf CurrPage = TotalPage Then
							  PrevUrl=GetPageStr(CurrPage - 1) :NextUrl=GetPageStr(1)
							Else
							  PrevUrl=GetPageStr(CurrPage - 1) : NextUrl=GetPageStr(CurrPage+1)
							End If
						For n=1 To TotalPage
						  If CurrPage = N Then
						  	ThumbList=ThumbList &"<li class=""currthumb""><a href=""" & GetPageStr(N) &""" target=""_self""><img src=""" & Split(PicUrlsArr(n-1),"|")(2) &""" border=""0""/></a></li>"
						  Else
						   ThumbList=ThumbList &"<li class=""normalthumb""><a href=""" & GetPageStr(N) &""" target=""_self""><img src=""" & Split(PicUrlsArr(n-1),"|")(2) &""" border=""0""/></a></li>"
						  End If
						Next
						DefaultImageSrc=Split(PicUrlsArr(CurrPage-1), "|")(1)
						DefaultImageIntro=Split(PicUrlsArr(CurrPage-1), "|")(0) 
					  Else 
					    ThumbList=ThumbList &"<li class=""currthumb""><a href=""" & GetPageStr(1) &""" target=""_self""><img src=""" & Split(PicUrlsArr(0),"|")(2) &""" border=""0""/></a></li>"
					    DefaultImageSrc=Split(PicUrlsArr(CurrPage-1), "|")(1)
						DefaultImageIntro=Split(PicUrlsArr(CurrPage-1), "|")(0) 
					  End If
					  
                      Tp=LFCls.GetConfigFromXML("picturelabel","/labeltemplate/label","style1")
					  Tp=Replace(Tp,"{$PrevUrl}",PrevUrl)
					  Tp=Replace(Tp,"{$NextUrl}",NextUrl)
					  Tp=Replace(Tp,"{$CurrPage}",CurrPage)
					  Tp=Replace(Tp,"{$TotalPage}",TotalPage)
					  Tp=Replace(Tp,"{$ShowThumbList}",ThumbList)
					  Tp=Replace(Tp,"{$DefaultImageSrc}",DefaultImageSrc)
					  Tp=Replace(Tp,"{$DefaultImageIntro}",DefaultImageIntro)
					  
				 Case 3
				       Tp=LFCls.GetConfigFromXML("picturelabel","/labeltemplate/label","style3")
                        if ((ubound(PicUrlsArr)+1) mod pagenum)=0 then
							Tpage=(ubound(PicUrlsArr)+1)\pagenum
						else
							Tpage=(ubound(PicUrlsArr)+1)\pagenum + 1
						end if
						If CurrPage>Tpage Then CurrPage=Tpage
					   if CurrPage<=1 then  n=0 else n=pagenum*(CurrPage-1)
					   For r=1 to pagenum
							  if n<=ubound(PicUrlsArr) Then
							    ThumbList=ThumbList & "<img title=""" &Split(PicUrlsArr(n), "|")(0) & """ class=""scrollLoading"" onclick=""javascript:window.open(this.src)"" alt='" & Split(PicUrlsArr(n), "|")(0) & "' style=""cursor:hand;background:url(" & KS.GetDomain &"images/default/loading.gif) no-repeat center;"" data-url=""" & Split(PicUrlsArr(n), "|")(1) &""" src=""" & KS.GetDomain &"images/default/pixel.gif"" border='0'><br/><span>" & Split(PicUrlsArr(n), "|")(0) & "</span>"
							  Else 
							    Exit For
							  End If
							  n=n+1
					   Next
					   Tp=Replace(Tp,"{$ShowPage}",GetPicturePage(TPage,CurrPage))
					   Tp=Replace(Tp,"{$ShowImgList}",ThumbList)
					   
				Case 4  '不分页
					    Dim BigImgSrc,IntroList
						For n=1 To TotalPage
						  IntroList=IntroList & Split(PicUrlsArr(n-1),"|")(0) &"|"
						  BigImgSrc=BigImgSrc & Split(PicUrlsArr(n-1),"|")(1) &"|"
						  If CurrPage = N Then
						  	ThumbList=ThumbList &"<li><a id=""t" & n & """ class=""currthumb"" href=""javascript:void(0)"" onclick=""showImg(" & n & ");""><img src=""" & Split(PicUrlsArr(n-1),"|")(2) &""" border=""0""/></a></li>"
						  Else
						   ThumbList=ThumbList &"<li><a id=""t" & n & """ class=""normalthumb"" href=""javascript:void(0)"" onclick=""showImg(" & n & ");""><img src=""" & Split(PicUrlsArr(n-1),"|")(2) &""" border=""0""/></a></li>"
						  End If
						Next
						DefaultImageSrc=Split(PicUrlsArr(CurrPage-1), "|")(1)
						DefaultImageIntro=Split(PicUrlsArr(CurrPage-1), "|")(0) 
                      Tp=LFCls.GetConfigFromXML("picturelabel","/labeltemplate/label","style4")
					  Tp=Replace(Tp,"{$TotalPage}",TotalPage)
					  Tp=Replace(Tp,"{$ImgArr}",BigImgSrc)
					  Tp=Replace(Tp,"{$IntroArr}",Replace(Replace(IntroList,"'","\'"),chr(10),"<br/>"))
					  Tp=Replace(Tp,"{$ShowThumbList}",ThumbList)
					  Tp=Replace(Tp,"{$DefaultImageSrc}",DefaultImageSrc)
					  Tp=Replace(Tp,"{$DefaultImageIntro}",DefaultImageIntro)
				End Select
				FileContent=Replace(FileContent,"{$ShowPictures}",Tp)
                If Tpage>1 Then FileContent=Replace(FileContent,"{$GetPictureName}",.Node.SelectSingleNode("@title").text & "(" & currpage & ")")
			Else
			    PageStr = Content
			End If
			     
				 .ModelID = ChannelID
				 .ItemID  = ID
				 .PageContent=PageStr
				 .NextUrl=NextUrl
				 .TotalPage=TotalPage
				 .Templates=""
				 .Scan FileContent
		 		  FileContent = .Templates
				
			FileContent = KSR.KSLabelReplaceAll(FileContent)
		  End With
		  KS.Echo FileContent
		  Set KSLabel=Nothing
	   End Sub
	   
	   '得到图片分页
	   Function GetPicturePage(TotalPage,CurrPage)
	        If TotalPage<=1 Then Exit Function
	        PageStr =  "<div id=""pageNext""><table><tr><td>"
			If CurrPage > 1 Then PageStr = PageStr & "<a class=""prev"" href=""" & GetPageStr(CurrPage-1) & """>上一页</a> "
			Dim StartPage,N,K,PageStr
			StartPage=1
			if (CurrPage>=10) then StartPage=(CurrPage\10-1)*10+CurrPage mod 10+2
				     For N = StartPage To TotalPage
						 If CurrPage = N Then
						  PageStr = PageStr & ("<a class=""curr"" href=""#"">" & N & "</a> ")
						 Else
						  PageStr = PageStr & ("<a class=""num"" href=""" & GetPageStr(N) & """>" & N & "</a> ")
						 End If
						 K=K+1
						 If K>=10 Then Exit For
					 Next
					 If CurrPage<>TotalPage Then PageStr = PageStr & " <a class=""next"" href=""" & GetPageStr(CurrPage+1) & """>下一页</a>"
					 PageStr = PageStr & "</td></tr></table></div>"
		GetPicturePage=PageStr
	   End Function
	   
	   Sub StaticDownContent()
	       SqlStr= "Select top 1 * From " & KS.C_S(ChannelID,2) & " Where ID=" & ID
			 Set RS=Conn.Execute(SqlStr)
		 IF RS.Eof And RS.Bof Then
		  RS.Close : Set RS=Nothing
		  KS.ShowTips "error","您要查看的" & KS.C_S(ChannelID,3) & "已删除。或是您非法传递注入参数!"
		 End IF
		 Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
		 With KSR 
			Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		      .Tid=.Node.SelectSingleNode("@tid").text
			  Call FCls.SetContentInfo(ChannelID,.Tid,ID,.Node.SelectSingleNode("@title").text)
			  FileContent = .LoadTemplate(.Node.SelectSingleNode("@templateid").text)
			  .ModelID = ChannelID
			  .ItemID  = ID
			  .PageContent=""
			  .NextUrl=""
			  .TotalPage=0
			  .Templates=""
			  .Scan FileContent
			  FileContent = .Templates 
			  FileContent = .KSLabelReplaceAll(FileContent)
		 End With
		 KS.Echo FileContent
		 
	   End Sub
	   Sub StaticFlashContent()
		 Call GetRecords()
		 IF RS.Eof And RS.Bof Then
		  RS.Close:Set RS=Nothing
		  KS.ShowTips "error","您要查看的" & KS.C_S(ChannelID,3) & "已删除。或是您非法传递注入参数!"
		 End IF
		 If RS("Verific")<>1 And UserLoginTF=False And KSUser.UserName<>RS("Inputer") Then
		   KS.ShowTips "error","对不起,该" & KS.C_S(ChannelID,3) & "还没有经过审核!"
		 End If
		 Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
		 With KSR 
			Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		      .Tid=.Node.SelectSingleNode("@tid").text
			  Call FCls.SetContentInfo(ChannelID,.Tid,ID,.Node.SelectSingleNode("@title").text)
			 
			  Call CheckCharge(KSR)
			
			 FileContent =.LoadTemplate(.Node.SelectSingleNode("@templateid").text)
			 
		  If Content<>"True" Then
		   Dim ChargeContent:ChargeContent=KS.CutFixContent(FileContent, "{=GetFlash", "}", 1)
		   If KS.IsNul(ChargeContent) Then ChargeContent=KS.CutFixContent(FileContent, "{=GetFlashByPlayer", "}", 1)
		   FileContent=Replace(FileContent,ChargeContent,Content)
		  End If
			 
			  .ModelID = ChannelID
			  .ItemID  = ID
			  .PageContent=""
			  .NextUrl=""
			  .TotalPage=0
			  .Templates=""
			  .Scan FileContent
			  FileContent = .Templates 
			  
		      FileContent = .KSLabelReplaceAll(FileContent)
		End With
		 KS.Echo FileContent
	   End Sub
	   Sub StaticProductContent()
	       SQLStr="Select top 1 * From " & KS.C_S(ChannelID,2) & "  Where verific=1 And ID=" & ID
			    Set RS=Conn.Execute(SqlStr)
		 IF RS.Eof And RS.Bof Then
		    RS.Close:Set RS=Nothing
		    KS.ShowTips "error","您要查看的" & KS.C_S(ChannelID,3) & "已删除或是暂停销售!"
		 End IF
		 Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
		 With KSR 
			Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		      .Tid=.Node.SelectSingleNode("@tid").text
			  Call FCls.SetContentInfo(ChannelID,.Tid,ID,.Node.SelectSingleNode("@title").text)
			  FileContent = .LoadTemplate(.Node.SelectSingleNode("@templateid").text)
			  .ModelID = ChannelID
			  .ItemID  = ID
			  .PageContent=""
			  .NextUrl=""
			  .TotalPage=0
			  .Templates=""
			  .Scan FileContent
			  FileContent = .Templates 
			  FileContent = .KSLabelReplaceAll(FileContent)
		 End With
		 KS.Echo FileContent
	   End Sub
	   
	   
	   '收费扣点处理过程
	   Sub PayPointProcess()
	      ModelChargeType=KS.ChkClng(KS.C_S(ChannelID,34))
	       Select Case ModelChargeType
			case 1 ChargeStr="资金" : ChargeStrUnit="元人民币": ChargeTableName="KS_LogMoney" : DateField="PayTime": IncomeOrPayOut="IncomeOrPayOut" : CurrPoint=KSUser.GetUserInfo("Money")
			case 2 ChargeStr="积分" : ChargeStrUnit="分积分": ChargeTableName="KS_LogScore": DateField="AddDate":IncomeOrPayOut="InOrOutFlag": CurrPoint=KSUser.GetScore
			case else   '按点券
			 ChargeStr=KS.Setting(45) : ChargeStrUnit=KS.Setting(46)&KS.Setting(45) : ChargeTableName="KS_LogPoint" : DateField="AddDate" :IncomeOrPayOut="InOrOutFlag": CurrPoint=KSUser.GetUserInfo("Point")
			End Select
	   
	       Dim UserChargeType:UserChargeType=KSUser.ChargeType
	        If (Cint(ReadPoint)>0 or InfoPurview=2 or (InfoPurview=0 And (ClassPurview=1 Or ClassPurview=2))) and KSUser.UserName<>UserName Then
					 
					     If UserChargeType=1 Then
							 Select Case ChargeType
							  Case 0:Call CheckPayTF("1=1")
							  Case 1:Call CheckPayTF("datediff(" & DataPart_H &"," & DateField & "," & SqlNowString & ")<" & PitchTime)
							  Case 2:Call CheckPayTF("Times<" & ReadTimes)
							  Case 3:Call CheckPayTF("datediff(" & DataPart_H &"," & DateField & "," & SqlNowString & ")<" & PitchTime & " or Times<" & ReadTimes)
							  Case 4:Call CheckPayTF("datediff(" & DataPart_H &"," & DateField & "," & SqlNowString & ")<" & PitchTime & " and Times<" & ReadTimes)
							  Case 5:Call PayConfirm()
							  End Select
						Elseif UserChargeType=2 Then
				          If KSUser.GetEdays <=0 Then
						     Content="<div align=center>对不起，你的账户已过期 <font color=red>" & KSUser.GetEdays & "</font> 天,此文需要在有效期内才可以查看，您可以<a href='../user/user_payonline.asp' target='_blank'>点此在线充值</a>或与我们联系充值！</div>"
						  Else
						   Call KSUser.UseLogConsum(KS.C_S(ChannelID,6),ChannelID,ID,KSR.Node.SelectSingleNode("@title").text)
						   Call GetContent()
						  End If
						Else
						 Call KSUser.UseLogConsum(KS.C_S(ChannelID,6),ChannelID,ID,KSR.Node.SelectSingleNode("@title").text)
						 Call GetContent()
						end if
					   Else
						  Call GetContent()
					   End IF
	   End Sub
	   '检查是否过期，如果过期要重复扣点券
	   '返回值 过期返回 true,未过期返回false
	   Sub CheckPayTF(Param)
	   
	    Dim SqlStr:SqlStr="Select top 1 Times From " & ChargeTableName & " Where ChannelID=" & ChannelID & " And InfoID=" & ID & " And " & IncomeOrPayOut & "=2 and UserName='" & KSUser.UserName & "' And (" & Param & ")"
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open SqlStr,conn,1,3

		IF RS.Eof And RS.Bof Then
			Call PayConfirm()	
		Else
		       RS.Movelast
			   RS(0)=RS(0)+1
			   RS.Update
			   Call KSUser.UseLogConsum(KS.C_S(ChannelID,6),ChannelID,ID,KSR.Node.SelectSingleNode("@title").text)
			   Call GetContent()
		End IF
		 RS.Close:Set RS=nothing
	   End Sub
	   
	   Sub PayConfirm()
	     If UserLoginTF=false Then Call GetNoLoginInfo(Content):Exit Sub
		 If ReadPoint<=0 Then Call GetContent():Exit Sub

			 If KS.ChkClng(CurrPoint)<ReadPoint Then
					 Content="<div style=""text-align:center"">对不起，你的可用" & ChargeStr & "不足!阅读本文需要 <span style=""color:red"">" & ReadPoint & "</span> " & ChargeStrUnit &",你还有 <span style=""color:green"">" & CurrPoint & "</span> " & ChargeStrUnit & "</div>,请及时与我们联系！" 
			 Else
					If PayTF="1" Then
					 Dim PayPoint : PayPoint=(ReadPoint*KS.C_C(KSR.Tid,11))/100
					 Dim Descript:Descript="阅读收费" & KS.C_S(ChannelID,3) & "“" & KSR.Node.SelectSingleNode("@title").text & "”"
					 Dim TcMsg:TcMsg=KS.C_S(ChannelID,3) & "“" & KSR.Node.SelectSingleNode("@title").text & "”的提成"
					 Dim ClientName:ClientName=KSUser.GetUserInfo("realname")
					 If KS.IsNul(ClientName) Then ClientName=KSUser.UserName
					 Select Case ModelChargeType
					   case 1
					     If PayPoint>0 Then Call KS.MoneyInOrOut(KSR.Node.SelectSingleNode("@inputer").text,KSR.Node.SelectSingleNode("@inputer").text,PayPoint,4,1,now,0,"系统",KS.C_S(ChannelID,3) & TcMsg,ChannelID,ID,1)
					     Call KS.MoneyInOrOut(KSUser.UserName,ClientName,ReadPoint,4,2,now,0,"系统",Descript,ChannelID,ID,1)
						 Call GetContent()
					   case 2
					     If KS.ChkClng(PayPoint)>0 Then Call KS.ScoreInOrOut(KSR.Node.SelectSingleNode("@inputer").text,1,KS.ChkClng(PayPoint),"系统",TcMsg,0,0)
						 Session("ScoreHasUse")="+" '设置只累计消费积分
					     Call KS.ScoreInOrOut(KSUser.UserName,2,KS.ChkClng(ReadPoint),"系统",Descript,ChannelID,ID)
						 Call GetContent()
					   case else
					        If PayPoint>0 Then Call KS.PointInOrOut(ChannelID,ID,KSR.Node.SelectSingleNode("@inputer").text,1,PayPoint,"系统",TcMsg,0)
							 Call KS.PointInOrOut(ChannelID,ID,KSUser.UserName,2,ReadPoint,"系统",Descript,0)
							 Call GetContent()
					 End Select
					 Call KSUser.UseLogConsum(KS.C_S(ChannelID,6),ChannelID,ID,KSR.Node.SelectSingleNode("@title").text)
					Else
					    Dim PayUrl
						if KS.C_S(ChannelID,48)=0 Then
						 PayUrl=DomainStr & "Item/Show.asp?m=" & ChannelID & "&d=" &ID&"&pt=1"
						Else
						 PayUrl=DomainStr & "?"& PreContentTag & "-"&ID & "-" & ChannelID & "-" & CurrPage &"-" &"1"& Extension
						End If
						Content="<div style=""text-align:center"">阅读本文需要消耗 <span style=""color:red"">" & ReadPoint & "</span> " & ChargeStrUnit &",你目前尚有 <span style=""color:green"">" & CurrPoint & "</span> " & ChargeStrUnit &"可用,阅读本文后，您将剩下 <span style=""color:blue"">" & CurrPoint-ReadPoint & "</span> " & ChargeStrUnit &"</div><div style=""text-align:center"">你确实愿意花 <span style=""color:red"">" & ReadPoint & "</span> " & ChargeStrUnit & "来阅读此文吗?</div><div>&nbsp;</div><div align=center><a href=""" & PayUrl & """>我愿意</a>    <a href=""" &DomainStr & """>我不愿意</a></div>"
					End If
			 End If
	   End Sub
	   Sub GetNoLoginInfo(ByRef content)
	       GCls.ComeUrl=GCls.GetUrl()
		   Content="<div style='text-align:center'><script src='../ks_inc/lhgdialog.js' language=""JavaScript""></script><script type='text/javascript'>function ShowLogin(){ $.dialog({title:""<img src='../user/images/icon18.png' align='absmiddle'>会员登录"",content:""url:../user/userlogin.asp?action=PoploginStr"",width:450,height:200});}</script>对不起，你还没有登录，本文至少要求本站的注册会员才可查看!</div><div style='text-align:center'>如果你还没有注册，请<a href=""../user/reg/""><span style='color:red'>点此注册</span></a>吧!</div><div style='text-align:center'>如果您已是本站注册会员，赶紧<a href=""javascript:ShowLogin();""><span style='color:red'>点此登录</span></a>吧！</div>"
	   End Sub
	   Sub GetContent()
	     Select Case KS.ChkClng((KS.C_S(Channelid,6)))
		  Case 1 Content="True"
		  Case 2 Content=KSR.Node.SelectSingleNode("@picurls").text
		  Case 4 Content="True"
		 End Select
		 UrlsTF=true
	   End Sub
End Class
%>