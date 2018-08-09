<%

Class RefreshFunction
        Private KS,DomainStr,CurrModelID,LabelStyle
		Private Templates,AjaxOut,ModelID,ClassID,IncludeSubClass,SpecialID,N  rem 定义本类全局变量 其中 ModelID 模型ID ClassID 栏目ID templates临时变量 N行号
		Private regEx, Matches, Match                                       rem 正则全局变量 
		Private XMLDoc,LabelParamStr,LoadSucceed,LabelID                    rem 标签参数XML对象
		Public ParamNode
	    public  XMLSql,Node,DocNode,Num,Param                               rem 文档XML对象
		Private LabelFunName                                                rem 标签函数名称
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set regEx = New RegExp
		  regEx.IgnoreCase = True
		  regEx.Global = True
		  Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		  DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set XMLDoc=Nothing  : Set XMLSql=Nothing
		 Set DocNode=Nothing : Set Node=Nothing
		End Sub
		
		Sub Echo(sStr)
			Templates    = Templates & sStr 
		End Sub 
		Sub EchoLn(sStr)
		    Templates    = Templates & sStr & VbNewLine
		End Sub
		
		Sub Scan(ByVal sTemplate)
			Dim iPosLast, iPosCur
			iPosLast    = 1
			While True 
				iPosCur    = InStr(iPosLast, sTemplate, "{@")
				If iPosCur>0 Then
					Echo    Mid(sTemplate, iPosLast, iPosCur-iPosLast)
					iPosLast  = Parse(sTemplate, iPosCur+2)
				Else 
					Echo    Mid(sTemplate, iPosLast)
					Templates=ParseIF(Templates)        '2012-4-29号增加循环支持IF判断 
					Exit Sub  
				End If 
		   Wend 
		End Sub 
		Function DoIf(byval condition,byval yes,byval no) 
            if(Eval(condition)) then DoIf=yes else DoIf=no 
        end function 
         
        '增加的代码  使用标签格式 {$IF 1=1 } {True} {False} {/$IF}  按照以下的正则 是不允许在条件分支中出现大括号内容的 
        '使用举例如    {$IF {@autoid} mod 2=0 } {</tr><tr>} {{$IF {@autoid}=1 }{<font color="Red">第一名</font>}{}{/$IF}} {/$IF}  这个是说明 使用中这个例子中的第二级IF应该是用不到吧呵呵 
        '条件中的判断格式： 原来asp该怎么写就怎么写 
        Function ParseIF(sTemplate) 
            dim condition,yes,no            
            regEx.Pattern = "\{\$IF([^\}]*)\}[^\{]*\{([^\}]*)\}[^\{]*\{([^\}]*)\}[^\{]*{/\$IF\}" 
            Set Matches = regEx.Execute(sTemplate) 
            On Error Resume Next 
            '不断的检查自身 用来替换嵌套的IF 
			dim n:n=0
            do while(Matches.Count<>0) 
                    set Match=Matches(0) 
                    condition=Match.SubMatches.Item(0) 
                    yes=Match.SubMatches.Item(1) 
                    no=Match.SubMatches.Item(2)                  
                    sTemplate=replace(sTemplate,Match.Value,DoIf(condition,yes,no))                   
                    set Matches=regEx.Execute(sTemplate)                   
					n=n+1
					if n>=5 then exit do
            loop            
            set Matches=nothing 
            ParseIF=sTemplate 
        end Function


		Function Parse(sTemplate, iPosBegin)
		    on error resume next
			Dim iPosCur, sValue, sTemp,MyNode
			iPosCur      = InStr(iPosBegin, sTemplate, "}")
			sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
			iPosBegin    = iPosCur+1
			select case Lcase(sTemp)
			 case "autoid" echo N+1
			 case "validdate" if Node.SelectSingleNode("@validdate").text="0" then echo "长期" else echo Node.SelectSingleNode("@validdate").text & "天"
			 case "photo","photourl","userface","bigphoto","logo" 
			  Dim PhotoUrl:PhotoUrl=Node.SelectSingleNode("@" & Lcase(sTemp)).text
			  if KS.IsNul(PhotoUrl) then PhotoUrl=DomainStr & "images/nopic.gif"
			  If Left(PhotoUrl,1)<>"/" And Left(Lcase(PhotoUrl),4)<>"http" Then PhotoUrl=DomainStr & PhotoUrl
			  echo PhotoUrl
			 case "groupbuyurl" If KS.ChkClng(KS.Setting(179))=1 Then echo KS.GetDomain & "groupbuy/show-" & Node.SelectSingleNode("@id").text & ".html" Else Echo KS.GetDomain & "shop/groupbuyshow.asp?id=" & Node.SelectSingleNode("@id").text
			 case "groupbuycarturl" If KS.ChkClng(KS.Setting(179))=1 Then echo KS.GetDomain & "groupbuy/cart-" & Node.SelectSingleNode("@id").text & ".html" Else Echo KS.GetDomain & "shop/GroupBuycart.asp?id=" & Node.SelectSingleNode("@id").text
			 case "groupbuyclassname" If KS.ChkClng(Node.SelectSingleNode("@classid").text)<>0 Then echo LFCls.GetSingleFieldValue("select top 1 CategoryName from KS_GroupBuyClass Where id=" & KS.ChkClng(Node.SelectSingleNode("@classid").text))
			 case "groupbuysold" echo "<script src="""  &KS.GetDomain & "shop/GetStock.asp?id=" & Node.SelectSingleNode("@id").text & "&action=GroupBuyHasSold""></script>"
			 case "sjurl" 
			  if Fcls.CallFrom3G="true" Then
			   echo "exam/index.asp?id=" & Node.SelectSingleNode("@id").text 
			  Else
			  if Node.SelectSingleNode("@dtfs").text="1" or Node.SelectSingleNode("@dtfs").text="2" then echo DomainStr & KS.C_S(9,8) & "sj/" & Node.SelectSingleNode("@id").text & ".htm" else echo DomainStr & "mnkc/exam/?id=" & Node.SelectSingleNode("@id").text
			  End If
			 case "sjsq" Dim ChargeStr,ModelChargeType:ModelChargeType=KS.ChkClng(KS.C_S(9,34)) : If ModelChargeType=0 Then ChargeStr=KS.Setting(45) Else If ModelChargeType=1 Then ChargeStr="元" Else ChargeStr="分"
			      if KS.ChkClng(Node.SelectSingleNode("@sq").text)>0 then echo "<span style='color:#f90'><img src='" & DomainStr & KS.C_S(9,10) & "/images/money.gif' align='absmiddle'>" &Node.SelectSingleNode("@sq").text & ChargeStr & "</span>" Else echo "<img src='" & DomainStr & KS.C_S(9,10) & "/images/free.gif' align='absmiddle'>"
			 case "sjtypename" echo split(Node.SelectSingleNode("@tname").text,"|")(ubound(split(Node.SelectSingleNode("@tname").text,"|"))-1)
			 case "sjtypeurl"
			 if Fcls.CallFrom3G="true" Then
			  echo "list.asp?tid=" & Node.SelectSingleNode("@tid").text
			 Else
			  echo DomainStr & KS.C_S(9,8) &"sjlb/" &Node.SelectSingleNode("@tid").text & "_1.htm"
			 End If
			 case "typeid" echo KS.GetGQTypeName(Node.SelectSingleNode("@typeid").text)
			 case "gqtypeid" echo Node.SelectSingleNode("@typeid").text
			 case "refreshtime" echo formatdatetime(Node.SelectSingleNode("@refreshtime").text,2)
			 case "adddate","lastposttime","expiredtime","joindate","activedate"
			 if ModelID=0 Then
			  echo LFCls.Get_Date_Field(Node.SelectSingleNode("@" & Lcase(sTemp)).text,ParamNode.getAttribute("daterule"))
			 Else
			  If DateDiff("d",Node.SelectSingleNode("@" & Lcase(sTemp)).text,Now())-KS.ChkClng(KS.C_S(ModelID,47))<0 Then
			   echo "<font style=""color:red"">" & LFCls.Get_Date_Field(Node.SelectSingleNode("@" & Lcase(sTemp)).text,ParamNode.getAttribute("daterule")) & "</font>"
			  Else
			  echo LFCls.Get_Date_Field(Node.SelectSingleNode("@" & Lcase(sTemp)).text,ParamNode.getAttribute("daterule"))
			  End If
			 End If
			 case "intro" echo KS.Gottopic(KS.LoseHtml(replace(replace(replace(replace(replace(Node.SelectSingleNode("@intro").text&"",chr(32),""),"&nbsp;",""),"	",""),"	",""),"　　","")),KS.ChkClng(ParamNode.getAttribute("introlen")))
			 case "classname" echo KS.C_C(Node.SelectSingleNode("@tid").text,1)
			 case "classurl" echo KS.GetFolderPath(Node.SelectSingleNode("@tid").text)
			 case "classename" dim carr:carr=split(Node.SelectSingleNode("@folder").text&"","/") : if isarr(carr) then echo carr(ubound(carr)-1)
			 case "topclassname" echo KS.C_C(split(KS.C_C(Node.SelectSingleNode("@tid").text,8),",")(0),1)
			 case "topclassurl" echo KS.GetFolderPath(split(KS.C_C(Node.SelectSingleNode("@tid").text,8),",")(0))
			 case "classimg" 
			   dim ci,ciarr
			   ci=Node.SelectSingleNode("@classbasicinfo").text
			   if not ks.isnul(ci) then  ciarr=split(ci,"||||"):if ks.isnul(ciarr(0)) then echo  DomainStr & "images/nopic.gif" else echo ciarr(0)
			 case "classintro"
			   ci=Node.SelectSingleNode("@classbasicinfo").text
			   if not ks.isnul(ci) then  ciarr=split(ci,"||||"):echo ciarr(1)
			 case "specialurl" echo KS.GetSpecialPath(Node.SelectSingleNode("@specialid").text,Node.SelectSingleNode("@specialename").text,Node.SelectSingleNode("@fsospecialindex").text)
			 case "specialclassurl" echo KS.GetFolderSpecialPath(Node.SelectSingleNode("@classid").text, True)
			 case "jobcompanyurl" echo JLCls.GetCompanyUrl(Node.SelectSingleNode("@id").text)
			 case "jobzwlist" echo JLCls.GetZWList(Node.SelectSingleNode("@username").text,ParamNode.getAttribute("zwlen"),KS.G_O_T_S(ParamNode.getAttribute("opentype")),ParamNode.getAttribute("showjipinflag"))
			 case "jipin" if Node.SelectSingleNode("@jipin").text="1" and lcase(ParamNode.getAttribute("showjipinflag"))="true" then echo "<span style='color:red'>急</span>"
			 case "companyname" echo KS.Gottopic(Node.SelectSingleNode("@companyname").text,KS.ChkClng(ParamNode.getAttribute("titlelen")))
			 case "zpnum" if Node.SelectSingleNode("@zpnum").text="0" then echo "不限" else echo Node.SelectSingleNode("@zpnum").text
			 case "jobzwurl" echo JLCls.GetZWUrl(Node.SelectSingleNode("@jobid").text)
			 case "jobresumeurl" echo JLCls.GetJianLiUrl(Node.SelectSingleNode("@id").text)
			 case "resumeage"  echo year(now)-KS.ChkClng("19" & Node.SelectSingleNode("@birth_y").text) & "岁"
			 case "price_member","price_market","limitbuyprice" echo formatnumber(Node.SelectSingleNode("@" & lcase(sTemp)).text,2,-1,-1)
			 case "price" if isnumeric(Node.SelectSingleNode("@price").text) then echo formatnumber(Node.SelectSingleNode("@price").text,2,-1,-1) else echo Node.SelectSingleNode("@price").text
			 case "price_vip" echo "<script src=""" & DomainStr & "shop/GetGroupPrice.asp?l=1&t=p&proid=" & Node.SelectSingleNode("@id").text & """></script>"
			 case "score_vip" echo "<script src=""" & DomainStr & "shop/GetGroupPrice.asp?l=1&t=s&proid=" & Node.SelectSingleNode("@id").text & """></script>"
			 case "brandname" sValue=KS.C_B(Node.SelectSingleNode("@brandid").text,"brandname") : If KS.IsNUL(sValue) Then echo "---" Else Echo sValue
			 case "brandename" sValue=KS.C_B(Node.SelectSingleNode("@brandid").text,"brandename") : If KS.IsNUL(sValue) Then Echo "---" Else Echo sValue
			 
			 case "spaceurl" echo GetSpaceUrl
			 case "newsurl" echo GetEnterpriseNewsUrl
			 case "blogname" echo KS.Gottopic(Node.SelectSingleNode("@blogname").text,KS.ChkClng(ParamNode.getAttribute("titlelen")))
			 case "logclassname" echo LFCls.GetSingleFieldValue("select top 1 typename from ks_blogtype where typeid=" & KS.ChkClng(Node.SelectSingleNode("@typeid").text))
			 case "logurl"  echo GetLogUrl
			 case "albumsurl" echo GetxcUrl
			 case "teamurl" echo GetGroupUrl
			 case "aqurl" echo GetAqUrl
			 case "askclassname","asksubclassname"
			   If IsObject(Application(KS.SiteSN&"_askclasslist")) Then
			    if Lcase(sTemp)="askclassname" then
			     dim anode:set anode=Application(KS.SiteSN&"_askclasslist").documentElement.selectSingleNode("row[@classid="&Node.SelectSingleNode("@bigclassid").text&"]")
				 if not anode is nothing then
			     echo aNode.SelectSingleNode("@classname").text
				 end if
				end if
			    if Lcase(sTemp)="asksubclassname" then
			     set anode=Application(KS.SiteSN&"_askclasslist").documentElement.selectSingleNode("row[@classid="&Node.SelectSingleNode("@smallclassid").text&"]")
				 if not anode is nothing then
			     echo aNode.SelectSingleNode("@classname").text
				 end if
				end if
			   End If
			 case "askzjurl" echo DomainStr & KS.ASetting(1) & "a.asp?askuserid=" & Node.SelectSingleNode("@userid").text
			 case "aqclassurl" echo GetAQClassUrl
			 case "rewardbyimg" if KS.ChkClng(Node.SelectSingleNode("@reward").text)>0 then echo "<img src=""" & domainstr & KS.ASetting(1) & "images/ask_xs.gif"" align=""absmiddle"" />" & Node.SelectSingleNode("@reward").text
			 case "cluburl" echo getcluburl
			 case "boardname" echo getboardinfo("boardname")
			 case "boardurl" echo KS.GetClubListUrl(node.selectsinglenode("@boardid").text)
			 case "subject" echo KS.Gottopic(KS.LoseHtml(replace(Node.SelectSingleNode("@subject").text,"{","｛#")),KS.ChkClng(ParamNode.getAttribute("titlelen")))
			 case "title","xcname","specialname"
			  Dim temptitle
			 If KS.C_S(ModelID,6)="1" Then
				temptitle=GetItemTitle(Node.SelectSingleNode("@"&Lcase(sTemp)).text,KS.ChkClng(ParamNode.getAttribute("titlelen")),true,Node.SelectSingleNode("@titletype").text, Node.SelectSingleNode("@titlefontcolor").text, Node.SelectSingleNode("@titlefonttype").text)
				If Node.SelectSingleNode("@isvideo").text="1" then TempTitle="<img border=""0"" src=""" & DomainStr &"images/video.gif"">" & TempTitle
			 End If
			 If KS.IsNul(TempTitle) Then temptitle=KS.Gottopic(Node.SelectSingleNode("@"&Lcase(sTemp)).text,KS.ChkClng(ParamNode.getAttribute("titlelen")))
			 If Instr(TempTitle,"[url]")<>0 Then TempTitle=Replace(Replace(TempTitle,"[url]",""),"[/url]","")
             echo temptitle
			 case "fulltitle" echo Node.SelectSingleNode("@title").text
			 case "fulltitles" if KS.IsNul(Node.SelectSingleNode("@fulltitles").text) Then echo Node.SelectSingleNode("@title").text Else echo Node.SelectSingleNode("@fulltitles").text
			 case "newimg"
			  	If ModelID<>0 Then
				  If ModelID=10 Then
				   If DateDiff("d",Node.SelectSingleNode("@lastjobtime").text,Now())-KS.ChkClng(KS.JSetting(24))<0 Then echo "<img src=""" & DomainStr &"images/new.gif"" border=""0""/>"
				  Else 
				   If DateDiff("d",Node.SelectSingleNode("@adddate").text,Now())-KS.ChkClng(KS.C_S(ModelID,47))<0 Then echo "<img src=""" & DomainStr &"images/new.gif"" border=""0""/>"
				  End If
				End If
			 case "hotimg"
			    If Node.SelectSingleNode("@popular").text=1 Then echo "<img src=""" & DomainStr & "images/hot.gif"" border=""0""/>"
			 case "link3gurl" '3g版内容链接URL
			     If ModelID=0 Then
					echo KS.Setting(3)&"3g/show.asp?m=" & Node.SelectSingleNode("@channelid").text &"&d=" & Node.SelectSingleNode("@id").text
				 Else
					echo KS.Setting(3)&"3g/show.asp?m=" & ModelID &"&d=" & Node.SelectSingleNode("@id").text
				 End If
			 case "linkurl" 
			     If ModelID=0 Then
					echo KS.GetItemURL(Node.SelectSingleNode("@channelid").text,Node.SelectSingleNode("@tid").text,Node.SelectSingleNode("@id").text,Node.SelectSingleNode("@fname").text)
				 Else
				    If ParamNode.getAttribute("callbyspace")="true" And (ModelID=2 Or ModelID=5) And Not KS.IsNul(Session("SpaceUserID")) Then
					 If Modelid=2 Then
					  If KS.SSetting(21)="1" Then
					   echo "show-photo-" &Session("SpaceUserID") & "-" & Node.SelectSingleNode("@id").text & KS.SSetting(22) 
					  Else
					   echo KS.Setting(3) & "space/?" & Session("SpaceUserID")& "/showphoto/" &Node.SelectSingleNode("@id").text 
					  End If
					 Else
					  If KS.SSetting(21)="1" Then 
					   echo "show-product-" &Session("SpaceUserID") & "-" & Node.SelectSingleNode("@id").text & KS.SSetting(22) 
					  Else
					   echo "?" & Session("SpaceUserID") & "/showproduct/" & Node.SelectSingleNode("@id").text
                      End If
					 End If
					Else
				    echo KS.GetItemURL(ModelID,Node.SelectSingleNode("@tid").text,Node.SelectSingleNode("@id").text,Node.SelectSingleNode("@fname").text)
					End If
				 End If
			 case else 
			   Dim TArr,FieldName,FieldParamArr,Length,FieldType,FieldValue
			   If Instr(sTemp,"(")<>0 Then
				   Tarr=Split(sTemp,"(")
				   FieldName=lcase(Tarr(0))
				   Set MyNode=Node.SelectSingleNode("@" & FieldName)
				   If Not MyNode Is Nothing Then 
				       FieldValue=MyNode.Text
					   FieldParamArr=split(replace(Tarr(1),")",""),",")
					   FieldType=FieldParamArr(0)
					   Select Case Lcase(FieldType)
					     case "text" echo KS.HTMLCode(LFCls.Get_Text_Field(FieldValue,FieldParamArr(1),FieldParamArr(2),FieldParamArr(3),FieldParamArr(4)))
						 case "num" echo LFCls.Get_Num_Field(FieldValue,FieldParamArr(1),FieldParamArr(2))
						 case "date" echo LFCls.Get_Date_Field(FieldValue,FieldParamArr(1))
					   End Select
					   
				  End If
			   Else
				   Set MyNode=Node.SelectSingleNode("@" &lcase(sTemp))
				   If Not MyNode Is Nothing Then echo KS.Gottopic(MyNode.Text,Length)
			   End If
			end select
			Parse    = iPosBegin
			Set MyNode=Nothing
			if err then err.clear
		End Function 
		
		
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:GetLabel
		'作 用: 解释执行系统函数标签
		'参 数: Content标签内容,如{Tag:GetGenericList modelid="1"}<li>循环体</li>{/Tag}
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function GetLabel(Content)
		 LabelStyle     = KS.GetTagLoop(Content)
		 LabelFunName   = Split(Content," ")(0)
   		 LabelParamStr  = Replace(Replace(Content, LabelFunName, ""),"}" & LabelStyle&"{/Tag}", "")
         LabelFunName   = Replace(LabelFunName,"{Tag:","")
		 ModelID = 0
		 
		 on error resume next
		 Execute("GetLabel= " & LabelFunName & "(LabelStyle)")
		 if err then err.clear
		End Function
		
		'加载标签必要参数
		Sub LoadLabelParam()
		     If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				 LoadSucceed = false : Exit Sub
			 Else
			     LoadSucceed = true 
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
		     LabelID   = ParamNode.getAttribute("labelid")
			 ClassID   = ParamNode.getAttribute("classid")
			 AjaxOut   = ParamNode.getAttribute("ajaxout")
			 ModelID         = ParamNode.getAttribute("modelid")         : If Not IsNumeric(ModelID) Then ModelID=1
			 SpecialID       = ParamNode.getAttribute("specialid")
			 IncludeSubClass = ParamNode.getAttribute("includesubclass") 
			 Num             = ParamNode.getAttribute("num")             : If Not Isnumeric(Num) Then Num=10
             If LabelFunName <>"GetPageList" Then LoadSQLParam
		End Sub

		'加载SQL查询条件
		Sub LoadSQLParam()
		    Dim DocProperty
		    If ClassID = "-1" Then ClassID = FCls.RefreshFolderID
			
			If LabelFunName="GetRelativeList" Then 
			 Dim RelativeType:RelativeType=KS.ChkClng(ParamNode.getAttribute("relativetype"))
			 Select Case RelativeType
			   Case 0  '按手工关联
				 If ModelID<>"0" Then
				   Param = " Inner Join KS_ItemInfoR R On I.ID=R.RelativeID Where I.Verific=1 And I.DelTF=0 And R.InfoID=" & FCls.RefreshInfoID & " And R.RelativeChannelID=" & ModelID&""
				 Else
				   Param = " Inner Join KS_ItemInfoR R On I.InfoID=R.RelativeID Where I.Verific=1 And I.DelTF=0 And R.ChannelID=" & Fcls.ChannelID & " and R.InfoID=" & FCls.RefreshInfoID&" and R.RelativeChannelID=I.ChannelID"
				 End If
				 Dim relativeText:relativeText=ParamNode.getAttribute("relativetext")
                 If relativeText<>"" Then Param=Param & " And relativeText='" & relativeText & "'"
			   Case 1  '按关键词关联
			      Dim KeyWords
				 If FCls.ChannelID<>"0" Then
				  KeyWords=LFCls.GetSingleFieldValue("Select top 1 KeyWords From " & KS.C_S(FCls.ChannelID,2) &" Where ID=" & KS.ChkClng(FCls.RefreshInfoID))
				  Else
				  KeyWords=LFCls.GetSingleFieldValue("Select top 1 KeyWords From [KS_ItemInfo] Where InfoID=" & KS.ChkClng(FCls.RefreshInfoID) &" And ChannelID=" & FCls.ChannelID)
				  End If
				  
				  If KS.IsNul(KeyWords) Then 
				   Param=" Where 1=0"
				  Else
					   Dim KeyWordsArr, I, SqlKeyWordStr
					   KeyWordsArr = Split(KeyWords, ",")
					   For I = 0 To UBound(KeyWordsArr)
									 If SqlKeyWordStr = "" Then
											SqlKeyWordStr = " instr(KeyWords,'" & KeyWordsArr(I) & "')>0 "
									 Else
											SqlKeyWordStr = SqlKeyWordStr & "or instr(KeyWords,'" & KeyWordsArr(I) & "')>0 "
									 End If
								
					  Next
					 If ModelID<>"0" Then
					   Param = " Where I.Verific=1 and ("&SqlKeyWordStr&") And I.ID<>" & FCls.RefreshInfoID
					 Else
					   Param = " Where I.Verific=1 and ("&SqlKeyWordStr&") And I.InfoID<>" & FCls.RefreshInfoID
					 End If				  
				End If 
				  
			   Case 2  '录入者关联
			     If ModelID<>"0" Then
				   Param = " Inner Join " & KS.C_S(ModelID,2) & " R On I.Inputer=R.Inputer Where I.Verific=1 And I.DelTF=0 And R.ID=" & FCls.RefreshInfoID
				 Else
				   Param = " Inner Join KS_ItemInfo R On I.Inputer=R.Inputer Where I.Verific=1 And I.DelTF=0 And R.ChannelID=" & Fcls.ChannelID & " and R.InfoID=" & FCls.RefreshInfoID
				 End If
			 End Select
			Else
			 If ModelID<>"0" Then
			 Param = " Where I.Verific=1 And I.DelTF=0"
			 Else
			 Param = " Inner Join KS_Channel C On I.ChannelID=C.ChannelID Where C.ChannelStatus=1 And I.Verific=1 And I.DelTF=0"
			 End If
			End If
		
			If ClassID = "" Then ClassID = "0"
			If ClassID <> "0" Then 
				If Instr(ClassID,",")<>0 Then 
				 Param= Param & " And I.Tid in('" & Replace(ClassID,",","','")& "')" 
				ElseIf CBool(IncludeSubClass) = True Then 
				 Param= Param & " And I.Tid In (" & KS.GetFolderTid(ClassID) & ")" 
				Else 
				 Param= Param & " And I.Tid='" & ClassID & "'"
				End If
			End If
			DocProperty = ParamNode.getAttribute("docproperty") : DocProperty=DocProperty & "0000000000"
			If Mid(DocProperty,1,1)=1 Then Param = Param & " And I.Recommend=1"
			If Mid(DocProperty,2,1)=1 Then Param = Param & " And I.Rolls=1"
			If Mid(DocProperty,3,1)=1 Then Param = Param & " And I.Strip=1"
			If Mid(DocProperty,4,1)=1 Then Param = Param & " And I.Popular=1"
			If Mid(DocProperty,5,1)=1 Then Param = Param & " And I.Slide=1"
			If Mid(DocProperty,6,1)=1 And KS.C_S(ModelID,6)=1 Then Param = Param & " And I.IsVideo=1"
			Dim ai,attr:attr = ParamNode.getAttribute("attr")
			If ModelID<>0 and Not KS.IsNul(attr) Then
			   attr=split(attr,"|")
			   for ai=0 to ubound(attr) 
			     Param=Param & " and I." & attr(ai) & "=1"
			   next
			End If
			
			
			Dim IsPic:IsPic=ParamNode.getAttribute("ispic")&""
			If Lcase(IsPic)="true" Then
			  If ModelID=0 Then
			   Param = Param & " And I.PhotoUrl<>''"
			  ElseIf KS.C_S(ModelID,6)=1 Then 
			   Param = Param & " And I.PicNews=1"
			  End If
			End If
			
			Dim Province,City
			If KS.C_S(ModelID,6)=1 or KS.C_S(ModelID,6)=8 Then
			Province=ParamNode.getAttribute("province")
			City=ParamNode.getAttribute("city")
			 If Not KS.IsNul(Province) Then Param=Param & " and I.Province='" & Province & "'"
			 If Not KS.IsNul(City) Then Param=Param & " and I.City='" & City & "'"
			End If
			If ParamNode.getAttribute("callbyspace")="true" And Not KS.IsNul(Session("SpaceUserName")) Then Param=Param & " And I.inputer='" & Session("SpaceUserName") & "'"
			Param = Param & KS.GetSpecialPara(ModelID,SpecialID)
		End Sub
		'取排序方式
		Function GetOrderParam()
		 Dim OrderStr
		 OrderStr  = ParamNode.getAttribute("orderstr")  : If OrderStr="" Then OrderStr=" I.ID Desc"
		 OrderStr=Lcase(OrderStr)
		 If trim(OrderStr)="rnd" Then
			  Randomize : OrderStr="Rnd(-(I.ID+"&Rnd()&"))"
		  ElseIf Lcase(Left(Trim(OrderStr),2))<>"id" Then  
			 OrderStr=OrderStr & ",I.ID Desc"
		  End If
		  If lcase(left(orderstr,2))="id" and LabelFunName="GetRelativeList" then orderstr="i." & orderstr
		  GetOrderParam = OrderStr
		End Function
		
		'加载模型通用查询字段
		Public Sub LoadField(ByVal ModelID,ByVal PrintType,ByVal PicStyle,ByVal ShowPicFlag,ByRef FieldStr,ByRef TableName,ByRef Param)
		   If ModelID="0" Then 
			 TableName = "[KS_ItemInfo]" 
			 FieldStr  = "I.ChannelID,I.InfoID as ID,I.Title,I.Tid,I.Intro,I.PhotoUrl,I.AddDate,I.Inputer,I.Popular,I.Fname,I.Hits"
		   Else 
			 TableName=KS.C_S(ModelID,2)
			 Select Case KS.C_S(ModelID,6) 
			  Case 1 
			   FieldStr  = "I.ID,I.Title,I.Tid,I.Inputer,I.Fname,I.AddDate,I.Popular,I.Hits,I.IsVideo"
			   FieldStr=FieldStr & ",I.TitleType,I.TitleFontColor,I.TitleFontType"
			   If PrintType>=2 Then  FieldStr=FieldStr & ",I.fulltitle as fulltitles,I.PhotoUrl,I.Intro" 
			   If PrintType>=3 Then  FieldStr=FieldStr & ",I.ReadPoint"
			  Case 2  
			   FieldStr  = "I.ID,I.Title,I.Tid,I.Inputer,I.Fname,I.AddDate,I.Popular,I.Hits,I.Score"
			   If PrintType>=2 Then FieldStr=FieldStr & ",I.PhotoUrl,I.PictureContent As Intro"
			   If PrintType>=3 Then FieldStr=FieldStr & ",I.ReadPoint"
			  Case 3  
			   FieldStr  = "I.ID,I.Title,I.Tid,I.Inputer,I.Fname,I.AddDate,I.Popular,I.Hits"
			   If PrintType>=2 Then FieldStr=FieldStr & ",I.PhotoUrl,I.DownContent As Intro,I.DownSize,I.Rank"
			   If PrintType>=3 Then FieldStr=FieldStr & ",I.ReadPoint"
			  Case 4  
			   FieldStr  = "I.ID,I.Title,I.Tid,I.Inputer,I.Fname,I.AddDate,I.Popular,I.Hits"
			   If PrintType>=2 Then FieldStr=FieldStr & ",I.PhotoUrl,I.FlashContent As Intro,I.Author,I.Rank"
			   If PrintType>=3 Then FieldStr=FieldStr & ",I.ReadPoint"
			  Case 5  
			   FieldStr  = "I.ID,I.Title,I.Tid,I.Inputer,I.Fname,I.AddDate,I.Popular,I.Hits,I.unit"
			   If PrintType>=2 Then FieldStr=FieldStr & ",I.PhotoUrl,I.ProIntro As Intro,I.BigPhoto,I.Price_member,I.Price_Member as Price,I.Price as Price_Market,I.Promodel,I.BrandID,I.ProID,I.LimitBuyPrice"
			  Case 7  
			   FieldStr  = "I.ID,I.Title,I.Tid,I.Inputer,I.Fname,I.AddDate,I.Popular,I.Hits"
			   If PrintType>=2 Then FieldStr=FieldStr & ",I.PhotoUrl,I.MovieContent As Intro"
			   If PrintType>=3 Or PicStyle=13 Or PicStyle=14 Or PicStyle=15 Then FieldStr=FieldStr & ",I.MovieAct,I.MovieDY,I.MovieDQ,I.MovieTime,I.MovieYY,I.ReadPoint,I.Rank"
			   If PrintType=2 And PicStyle=15 Then FieldStr=FieldStr & ",I.MovieDy"
			  Case 8  
			   FieldStr  = "I.ID,I.Title,I.Tid,I.Inputer,I.Fname,I.AddDate,I.Popular,I.Hits,I.TypeID,I.Price"
			   If PrintType>=2 Then FieldStr=FieldStr & ",I.PhotoUrl,I.GQContent As Intro"
			   If PrintType>=3 Or PicStyle=16 Or PicStyle=17 Then FieldStr=FieldStr & ",I.ValidDate,I.ContactMan,I.Tel,I.Address,I.Province,I.City,I.CompanyName"
			   If KS.ChkClng(ParamNode.getAttribute("typeid"))<>0 Then Param =Param & " And I.TypeID="&KS.ChkClng(ParamNode.getAttribute("typeid"))
			  Case Else  
			   FieldStr  = "I.ID,I.Title,I.Tid,I.PhotoUrl,I.AddDate,I.Inputer,I.Popular,I.Fname,I.Hits"
			   If PrintType=2 or (instr(LabelStyle,"{@photourl}")>0 and PrintType>2) Then Param = Param & " And I.PhotoUrl<>''"
			 End Select
			 If PrintType=4 Then FieldStr=FieldStr & GetDiyFieldStr(ModelID)
			End If
		End Sub
		
		'加载动态分页的标签参数
		Public Sub LoadPageParam(xml,pnode,channelid)
		  Set XMLSql    = xml
		  Set ParamNode = pnode
		  ModelID=ChannelID
		End Sub
		
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:GetGenericList
		'作 用: 通用列表标签函数
		'参 数: LabelStyle 标签样式
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function GetGenericList(LabelStyle)
		     LoadLabelParam
			 If LoadSucceed = false Then GetGenericList="标签加载出错!":Exit Function 
			 If LabelID<>"ajax" and Cbool(AjaxOut)=true Then 
			  GetGenericList="<span id=""ks" & LabelID & "_" & ParamNode.getAttribute("classid") & "_" & FCls.RefreshFolderID & "_" &FCls.RefreshInfoID &"_" & FCls.ChannelID & """></span>":Exit Function
			 End If
			 Dim ShowPicFlag,ShowHotFlag
			 Dim SqlStr, M_L_S, O_T_S,C_F_T,TableName,FieldStr,PrintType,PicStyle
			 ShowPicFlag     = ParamNode.getAttribute("showpicflag") 
			 ShowHotFlag     = ParamNode.getAttribute("showhotflag") 
			 PrintType       = ParamNode.getAttribute("printtype")       : If Not IsNumeric(PrintType) Then PrintType=1
			 PicStyle        = ParamNode.getAttribute("picstyle")        : If Not IsNumeric(PicStyle) Then PicStyle=1
			
			LoadField ModelID,PrintType,PicStyle,ShowPicFlag,FieldStr,TableName,Param

			Dim SQLParam:SQLParam=ParamNode.getAttribute("sqlparam")
			If Not KS.IsNul(SQLParam) Then
			 'Param="Where ("&split(lcase(Param),"where ")(1) & ") " & SQLParam
			 Param=Param & " " & SQLParam
			End If

			SqlStr = "SELECT TOP " & Num & " " & FieldStr & " FROM " & TableName & " I " & Param & " ORDER BY I.IsTop Desc," & GetOrderParam()
			Dim RS:Set RS=Conn.Execute(SqlStr)
			If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			RS.Close:Set RS=Nothing
			If IsObject(XMLSql) Then
			 GetGenericList=ExplainGerericListLabelBody(LabelStyle)
			End If 
			Set Node=Nothing
		End Function
		
		'解释通用列表标签体
		Function ExplainGerericListLabelBody(ByVal LabelStyle)
		 Dim TotalNum,Col,K,I,S_C_N,C_N_Link,Title,TempTitle,PicTF,NewTF,HotTF,T_Len,NewImgStr,HotImgStr,LinkUrl
		 Dim DateStr,DateRule,DateAlign,DateCss,ColSpanNum,T_CssStr,O_T_S,R_H,NaviStr,SplitPic,MoreType,MoreLink,C_F_T,M_L_S
		 Dim PrintType
		 PrintType = KS.ChkClng(ParamNode.getAttribute("printtype"))
		 Col       = KS.ChkClng(ParamNode.getAttribute("col")) : If Col=0 Then Col=1
		 S_C_N     = ParamNode.getAttribute("showclassname") : If KS.IsNul(S_C_N) Then S_C_N=false
		 T_Len     = KS.ChkClng(ParamNode.getAttribute("titlelen"))
		 PicTF     = ParamNode.getAttribute("showpicflag") :If KS.IsNul(PicTF) Then PicTF=false Else PicTF=Cbool(PicTF)
		 NewTF     = ParamNode.getAttribute("shownewflag") :If KS.IsNul(NewTF) Then NewTF=false Else NewTF=Cbool(NewTF)
		 HotTF     = ParamNode.getAttribute("showhotflag") :If KS.IsNul(HotTF) Then HotTF=false Else HotTF=Cbool(HotTF)
		 DateRule  = ParamNode.getAttribute("daterule")
		 DateAlign = ParamNode.getAttribute("datealign")
		 DateCss   = ParamNode.getAttribute("datecss")
		 SplitPic  = ParamNode.getAttribute("splitpic")
		 MoreType  = ParamNode.getAttribute("morelinktype")
		 MoreLink  = ParamNode.getAttribute("morelink")
		 
		 O_T_S     = KS.G_O_T_S(ParamNode.getAttribute("opentype"))
		 T_CssStr  = KS.GetCss(ParamNode.getAttribute("titlecss"))
		 R_H       = KS.G_R_H(ParamNode.getAttribute("rowheight"))
		 NaviStr   = KS.GetNavi(ParamNode.getAttribute("navtype"), ParamNode.getAttribute("nav"))
		 If ClassID = "-1" Or Instr(ClassID,",")<>0 Then C_F_T = True Else C_F_T = False
		 If MoreLink <> "" And ClassID <> "0" And C_F_T = False Then M_L_S = KS.GetMoreLink(1,Col, R_H, MoreType, MoreLink, KS.GetFolderPath(ClassID), O_T_S)
		 Set DocNode=XMLSql.DocumentElement.SelectNodes("row")
		 TotalNum=DocNode.length
		 If PrintType=1 Then  
		     Templates="" : N=0 : ColSpanNum=Col
			 echoln "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
			 For K=0 To TotalNum-1
				 echo "<tr>" & vbCrLf
				 For I = 1 To Col
						 Set Node=DocNode.Item(n)
						 If CBool(S_C_N) = True Then C_N_Link = "<span class=""category"">[" & KS.GetClassNP(Node.SelectSingleNode("@tid").text) & "]</span>"			
						  Title = Node.SelectSingleNode("@title").text
						  If ModelID=0 Then CurrModelID=Cint(Node.SelectSingleNode("@channelid").text) Else CurrModelID=ModelID
						  If KS.C_S(ModelID,6)=1 And PicTF=true Then
						   TempTitle = GetItemTitle(Title, T_Len, PicTF, Node.SelectSingleNode("@titletype").text, Node.SelectSingleNode("@titlefontcolor").text, Node.SelectSingleNode("@titlefonttype").text)
						  Else
						   TempTitle= KS.Gottopic(Title,T_Len)
						  End If
						  
				  
						 If Cbool(NewTF)=True And DateDiff("d",Node.SelectSingleNode("@adddate").text,Now())-KS.ChkClng(KS.C_S(CurrModelID,47))<0 Then NewImgStr="<img src=""" & DomainStr &"images/new.gif"" border=""0""/>" Else NewImgStr=""
						 If Cbool(HotTF)=True And Node.SelectSingleNode("@popular").text=1 Then HotImgStr="<img src=""" & DomainStr & "images/hot.gif"" border=""0""/>" Else HotImgStr=""
						  DateStr=GetDateStr(ModelID,Node.SelectSingleNode("@adddate").text,DateRule,DateAlign,KS.GetCss(DateCss),Col,ColSpanNum)
						  LinkUrl=KS.GetItemURL(CurrModelID,Node.SelectSingleNode("@tid").text,Node.SelectSingleNode("@id").text,Node.SelectSingleNode("@fname").text)
						  TempTitle = "<a" & T_CssStr & " href=""" & LinkUrl & """" & O_T_S & " title=""" & Title & """>" & TempTitle & "</a>"
						  If KS.C_S(ModelID,6)=1 Then
						    If Node.SelectSingleNode("@isvideo").text="1" then TempTitle="<img border=""0"" src=""" & DomainStr &"images/video.gif"">" & TempTitle
						  End If
						  
						  If CurrModelID=8 Then
						   Dim ShowGQType:ShowGQType=ParamNode.getAttribute("showgqtype") : If KS.IsNul(ShowGQType) Then ShowGQType=false
						   If Cbool(ShowGQType)=true Then TempTitle=KS.GetGQTypeName(Node.SelectSingleNode("@typeid").text) & TempTitle
						  End If
						  
						  If Col=1 Then
							 echoln ("  <td height=""" & R_H & """>" & (NaviStr & C_N_Link & TempTitle &NewImgStr&HotImgStr& DateStr) & "</td>")
						  Else
							 echoln ("<td width=""" & CInt(100 / CInt(Col)) & "%"" height=""" & R_H & """>")
							 echoln ("<table width=""100%"" height=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"">")
							 echoln ("<tr><td> " & NaviStr & C_N_Link & TempTitle &NewImgStr&HotImgStr & DateStr)
							 echoln ("</td></tr>" & vbcrlf &"   </table>" & vbCrLf & "  </td>")
						  End If
						  N=N+1 : If N>=TotalNum Then Exit For
					  Next
					  echoln "</tr>"
					  echo KS.GetSplitPic(SplitPic,ColSpanNum)
					  If N>=TotalNum Then Exit For
					Next
					echoln M_L_S & ("</table>")
		  ElseIf PrintType=2 Then
		        Templates=ExplainPic(TotalNum,Col,T_Len,O_T_S,T_CssStr)
		  Else
		        Templates=ExplainDiyStyle(LabelStyle,TotalNum)
		  End If

		  Set DocNode=Nothing
		  Set Node=Nothing
		   ExplainGerericListLabelBody=Templates
		End Function
		
		'---------------------------------------------------------------
		'函数名:GetDateStr
		'作用：取日期的样式
		'参数：AddDate,DateRule,DateAlign,DateCssStr,ByRef ColSpanNum
		'---------------------------------------------------------------
		Function GetDateStr(ChannelID,AddDate,DateRule,DateAlign,DateCssStr,ByVal ColNumber,ByRef ColSpanNum)
		       If CStr(DateRule) <> "0" And CStr("DateRule") <> "" Then
					  	Dim NowDate,NowFormatStr
						If DateDiff("d",AddDate,Now())-KS.ChkClng(KS.C_S(ChannelID,47))<0 Then NowFormatStr=" style=""color:red""" Else  NowFormatStr=""
						If Lcase(DateAlign)="left" Then
							GetDateStr="&nbsp;<span" & NowFormatStr & DateCssStr &">" & LFCls.Get_Date_Field(AddDate, DateRule) & "</span>"
							ColSpanNum = ColNumber+1
						Else
							GetDateStr="</td><td width=""*"" nowrap align=" & DateAlign & "><span" & NowFormatStr & DateCssStr & ">" &LFCls.Get_Date_Field(AddDate, DateRule) & "</span>"
							ColSpanNum = ColNumber+2
						End If
				Else
				 GetDateStr="":ColSpanNum = ColNumber+1
				End If
		End Function
		
		'解释自定义样式	
		Function ExplainDiyStyle(LabelStyle,TotalNum)
		        If Instr(LabelStyle,"[loop")=0 Then LabelStyle="[loop={@num}]" & LabelStyle
				If Instr(LabelStyle,"[/loop]")=0 Then LabelStyle=LabelStyle & "[/loop]"
                LabelStyle  = Replace(LabelStyle,"{@num}",TotalNum)
				LabelStyle  = Replace(LabelStyle,Chr(10) ,"$KS:Break$")
		        Dim LoopTimes,CirLabelContent,TempCirContent,K
				'regEx.Pattern =  "\[loop(=\d*){0,1}].+?\[/loop]"
				regEx.Pattern = "\[loop=\d*][\s\S]*?\[/loop]"
				Set Matches = regEx.Execute(LabelStyle) : N=0
			   'For K=0 To TotalNum-1
					For Each Match In Matches
					    LoopTimes=GetLoopNum(Match.Value)   '循环次数
						CirLabelContent = Replace(Replace(Replace(Match.value,"[loop=" & LoopTimes&"]",""),"[/loop]",""),"[loop]","")					
						LabelStyle=Replace(LabelStyle,Match.Value,GetCirLabelContent(CirLabelContent,LoopTimes,N,TotalNum),1,1)
						If N>=KS.ChkClng(TotalNum) Then Exit For
					 Next
			   '	 If N>=KS.ChkCLng(TotalNum) Then Exit For
		       'Next
		       ExplainDiyStyle=CleanLabel(LabelStyle)
		End Function	

		'解释自定义样式循环体
		Function GetCirLabelContent(ByVal CirStyle,ByVal LoopTimes,N,ByVal TotalNum)
			 Dim I:Templates="" 
			 If Not Isnumeric(LoopTimes) Or LoopTimes=0 Then LoopTimes=TotalNum
			 For I=1 To LoopTimes
			  Set Node=DocNode.Item(N)
			  Scan CirStyle
			  N=N+1 : If N>=TotalNum Then Exit For
			 Next
			 GetCirLabelContent=Templates
		End Function
		'返回自定义字段
		Public Function GetDiyFieldStr(ByVal ChannelID)
		  If ChannelID=0 Then Exit Function
		    Dim FieldXML,FieldNode,KSCls,TStr,N
			set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			FieldXML.async = false
			FieldXML.setProperty "ServerHTTPRequest", true 
			FieldXML.load(Server.MapPath(KS.Setting(3)&"Config/fielditem/field_" & ChannelID&".xml"))
			If Not IsObject(FieldXML) Then Exit Function
			if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
						  Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype!=0]")
						  If DiyNode.Length>0 Then
								For Each N In DiyNode
									Tstr=Tstr & ",I." & N.SelectSingleNode("@fieldname").text
			                        If N.SelectSingleNode("showunit").text="1" then Tstr=Tstr &",I." & N.SelectSingleNode("@fieldname").text&"_unit"
								Next
								Set N=nothing
						  End If
			End If
		  GetDiyFieldStr=Tstr
		End Function
		
		'返回循环次数
		Function GetLoopNum(Content)
			 regEx.Pattern="\[loop=\d*]"
			 Set Matches = regEx.Execute(Content)
			 If Matches.count > 0 Then
			  GetLoopNum=Replace(Replace(Matches.item(0),"[loop=",""),"]","")
			 Else
			  GetLoopNum=0
			 end if
		End Function
		'消除多余的循环体
		Function CleanLabel(Content)
			'regEx.Pattern = "\[loop=\d*][^\[\]]*\[/loop]"
			regEx.Pattern = "\[loop=\d*].*\[/loop]"
			Set Matches = regEx.Execute(Content)
			For Each Match In Matches
				Content=Replace(Content,Match.value,"")
			Next
			CleanLabel=Replace(Content,"$KS:Break$",vbcrlf)
		End Function
		
		'解释图片标签
		Function ExplainPic(TotalNum,Col,T_Len,O_T_S,T_CssStr)
		  on error resume next
		  Dim PrintType,PicStyle,PicWidth,PicHeight,PicWidthStr,PicHeightStr,PicBorderColor,PicSpacing,TempTitleStr
		  Dim K,I,LinkAndPicStr,LinkUrl,Title,PhotoUrl,C_Len,C_N_Link,Rank,PicBorderColorStr
		  PrintType = KS.ChkClng(ParamNode.getAttribute("printtype"))
		  PicStyle  = ParamNode.getAttribute("picstyle") : If Not IsNumeric(PicStyle) Then PicStyle=1
		  PicWidth  = KS.ChkClng(ParamNode.getAttribute("picwidth")):If PicWidth<>0 Then PicWidthStr=" width=""" & PicWidth & """" Else PicWidthStr=""
		  PicHeight = KS.ChkClng(ParamNode.getAttribute("picheight")):If PicHeight<>0 Then PicHeightStr=" height=""" & PicHeight & """" Else PicHeightStr=""
		  PicBorderColor = ParamNode.getAttribute("picbordercolor")
		  PicSpacing     = KS.ChkClng(ParamNode.getAttribute("picspacing")) : If PicSpacing<>0 Then PicSpacing="padding-bottom:" & PicSpacing & "px;"
		  C_Len          = KS.ChkClng(ParamNode.getAttribute("introlen"))
		  Templates = "" : N=0
		   echoln "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
		   For K=0 To TotalNum-1
				 echoln  "<tr>"
				 For I = 1 To Col
				      Set Node=DocNode.Item(n)
					   If ModelID = 0 Then CurrModelID=Cint(Node.SelectSingleNode("@channelid").text) Else CurrModelID=ModelID
						LinkUrl=KS.GetItemURL(CurrModelID,Node.SelectSingleNode("@tid").text,Node.SelectSingleNode("@id").text,Node.SelectSingleNode("@fname").text)
					   Title         = Node.SelectSingleNode("@title").text
					   PhotoUrl      = Node.SelectSingleNode("@photourl").text : PhotoUrl= GetPicUrl(PhotoUrl)
					   If PicBorderColor<>"" Then PicBorderColorStr=" style=""border:1px solid " & PicBorderColor & """"
				       LinkAndPicStr = "<a href=""" & LinkUrl & """" & O_T_S & " title=""" & Title & """><Img Src=""" & PhotoUrl & """ border=""0""" & PicWidthStr & PicHeightStr & PicBorderColorStr & " class=""pic"" alt=""" & Title & """ align=""absmiddle""/></a>"                   
					   If T_Len<>0 Then
                       TempTitleStr  = "<a" & T_CssStr & " href=""" & LinkUrl & """" & O_T_S & " title=""" & Title & """>" & KS.GotTopic(Title, T_Len) & "</a>"
					   End If
					   If ModelID = 5 Then
					    Dim ButtonStr,ButtonType,PriceType,PriceStr
						ButtonType = ParamNode.getAttribute("buttontype") : If Not IsNumeric(ButtonType) Then ButtonType = 1
						PriceType  = ParamNode.getAttribute("pricetype")  : If Not IsNumeric(PriceType) Then PriceType = 1
					    ButtonStr  = GetButtonStr(ButtonType,Node.SelectSingleNode("@id").text,LinkUrl,O_T_S)
						PriceStr   = GetPriceStr(PriceType,Node.SelectSingleNode("@price_market").text,Node.SelectSingleNode("@price_member").text)
					   ElseIf ModelID=8 And (PrintType>=3 Or PicStyle=16 Or PicStyle=17) Then
					     Dim ProCity,Province,City
						 Province = Node.SelectSingleNode("@province").text
						 City     = Node.SelectSingleNode("@city").text
						 IF Not KS.IsNul(Province) Then ProCity=Province & "/" & City Else ProCity="地区不限"
					   End If
					   
					  echoln ("<td width=""" & CInt(100 / CInt(Col)) & "%"" style=""text-align:center;" & PicSpacing & """>")
					  select case PicStyle
						case 1
						   echo LinkAndPicStr
						case 2
						   echoln "<div class=""image"">" & LinkAndPicStr & "</div>"
						   echoln "<div class=""t"">" & TempTitleStr & "</div>"
						case 3
						   echoln "<table style=""margin:3px;width:100%"" cellSpacing=""0"" cellPadding=""0"" border=""0"">"
						   echoln "  <tr>"
						   echoln "   <td class=""image"" style=""text-align:center"" width=""" & PicWidth+10 & """>"
						   echoln "    " & LinkAndPicStr
						   echoln "   </td>"
						   echoln "   <td>"
						   echoln "    <div class=""t"">" & TempTitleStr & "</div>"
						   echoln "    <div class=""text"">" & KS.GotTopic(Replace(Replace(KS.LoseHtml(Node.SelectSingleNode("@intro").text), vbCrLf, ""), "&nbsp;", ""), C_Len) & "……</div>"
						   echoln "   </td>"
						   echoln "  </tr>"
						   echoln " </table>"
						case 4
						   echoln "<table style=""margin:3px;width:100%"" cellSpacing=""0"" cellPadding=""0"" border=""0"">"
						   echoln "  <tr>"
						   echoln "   <td>"
						   echoln "    <div class=""t"">" & TempTitleStr & "</div>"
						   echoln "    <div class=""text"">" & KS.GotTopic(Replace(Replace(KS.LoseHtml(Node.SelectSingleNode("@intro").text), vbCrLf, ""), "&nbsp;", ""), C_Len) & "……</div>"
						   echoln "   </td>"
						   echoln "   <td class=""image"" style=""text-align:center"" width=""" & PicWidth+10 & """>"
						   echoln "    " & LinkAndPicStr
						   echoln "   </td>"
						   echoln "  </tr>"
						   echoln " </table>"
					   '<!-动漫系统开始->
                       case 5  
					       C_N_Link = KS.GetClassNP(Node.SelectSingleNode("@tid").text)
						   Rank     = Replace(Node.SelectSingleNode("@rank").text,"★","<img src=""" & DomainStr & "Images/Star.gif"" />")
					       echoln "<table cellspacing=""0"" cellpadding=""2"" width=""100%"" border=""0"">"
						   echoln "  <tr>"
						   echoln "   <td width="""&PicWidth+10&""" class=""image"" style=""text-align:center"">" & LinkAndPicStr & "</td>"
						   echoln "   <td style=""line-height: 150%;text-align:left"" valign=""top""><div class=""t"">" & TempTitleStr & "</div>"
						   echoln "    <div class=""lb"">类别：" & C_N_Link & "</div>"
						   echoln "    <div class=""time"">时间：<span>" & LFCls.Get_Date_Field(Node.SelectSingleNode("@adddate").text,"YYYY-MM-DD") & "</span></div>"
						   echoln "    <div class=""hits"">人气：<span>" & Node.SelectSingleNode("@hits").text & "</span></div>"
						   echoln "    <div class=""tj"">推荐：" & Rank & "</div>"
						   echoln "   </td>"
						   echoln "  </tr>"
						   echoln "</table>"
					   case 6
					       C_N_Link = KS.GetClassNP(Node.SelectSingleNode("@tid").text)
						   Rank     = Replace(Node.SelectSingleNode("@rank").text,"★","<img src=""" & DomainStr & "Images/Star.gif"" />")
					       echoln "<table cellspacing=""0"" cellpadding=""2"" width=""100%"" border=""0"">"
					       echoln "  <tr>"
					       echoln "   <td width=""" & PicWidth+10 &""" style=""text-align:center"">" & LinkAndPicStr & "</td>"
					       echoln "   <td style=""line-height: 150%;text-align:left"" vAlign=""top""><div class=""t"">" & TempTitleStr & "</div>"
					       echoln "   <div class=""text"">" & KS.GotTopic(Replace(Replace(KS.LoseHtml(Node.SelectSingleNode("@intro").text), vbCrLf, ""), "&nbsp;", ""), C_Len) & "……</div>"
					       echoln "   <div class=""info"">作者：" & Node.SelectSingleNode("@author").text & " | 人气：<span class=""rq"">" & Node.SelectSingleNode("@hits").text & "</span> | 推荐：" & Rank  & "</div>"
					       echoln "   </td>"
					       echoln " </tr>"
					       echoln "</table>"
					   '<!-动漫系统结束->
					   '<!-商城系统开始->
					   case 7 
						   echoln "<div class=""image"">" & LinkAndPicStr & "</div>"
						   echoln "<div class=""btn"">" & ButtonStr & "</div>"
					   case 8
						   echoln "<div class=""image"">" & LinkAndPicStr & "</div>"
						   echoln "<div class=""t"">" & TempTitleStr & "</div>"
						   echoln "<div class=""btn"">" & ButtonStr & "</div>"
					   case 9
						   echoln "<div class=""image"">" & LinkAndPicStr & "</div>"
						   echoln "<div class=""t"">" & TempTitleStr & "</div>"
						   echoln "<div class=""price"">" & PriceStr & "</div>"
						   echoln "<div class=""btn"">" & ButtonStr & "</div>"
					   case 10,11,12
					       echoln "<table cellSpacing=""0"" cellPadding=""0"" width=""100%"" border=""0"">"
						   echoln  " <tr>"
						   echoln  "  <td style=""text-align:center"" width=""" &PicWidth+10 & """>" 
						   echoln "<div class=""image"">" & LinkAndPicStr & "</div>"
						   If PicStyle=11 Then  echoln "<div class=""t"">" &TempTitleStr & "</div>"
						   echoln "  </td>"
						   echoln "  <td>"
						   If PicStyle=12 Then  echoln "<div class=""t"">" &TempTitleStr & "</div>"
						   echoln "   <div class=""price"">" & PriceStr & "</div>"
						   echoln "   <div class=""btn"">" & ButtonStr  & "</div>"
						   echoln "  </td>"
						   echoln " </tr>"
						   echoln "</table>"
					   '<!-商城系统结束->
					   '<!-影视系统开始->
					   case 13,14,15
					       echoln "<table cellSpacing=""0"" cellPadding=""0"" width=""100%"" border=""0"">"
						   echoln  " <tr>"
						   echoln  "  <td style=""text-align:center"" width=""" &PicWidth+10 & """>" 
						   echoln "   <div class=""image"">" & LinkAndPicStr & "</div>"
						   echoln "  </td>"
						   echoln "  <td>"
						   echoln "   <div class=""t"">" & TempTitleStr & "</div>"
						   If PicStyle=14 Then  
						    C_N_Link = KS.GetClassNP(Node.SelectSingleNode("@tid").text)
						    Rank     = Replace(Node.SelectSingleNode("@rank").text,"★","<img src=""" & DomainStr & "Images/Star.gif"" />")
						    echoln "  <div class=""text"">" &KS.GotTopic(Replace(Replace(KS.LoseHtml(Node.SelectSingleNode("@intro").text), vbCrLf, ""), "&nbsp;", ""), C_Len) & "</div>"
							echoln "  <div class=""pro"">主演：<span class=""act"">" & KS.Gottopic(Node.SelectSingleNode("@movieact").text,10) & "</span> | 类别：<span class=""lb"">" & C_N_Link & "</span> | 语言：<span class=""yy"">" & Node.SelectSingleNode("@movieyy").text & "</span> 推荐：" & Rank & "</div>"
						   ElseIf PicStyle=13 Then
						   echoln "   <div class=""pro"">主演：<span class=""act"">" & KS.Gottopic(Node.SelectSingleNode("@movieact").text,10) & "</span> 产地：<span class=""cd"">" & Node.SelectSingleNode("@moviedq").text & "</span></div>"
						   echoln "   <div class=""text"">简介：<span class=""intro"">" & KS.GotTopic(Replace(Replace(KS.LoseHtml(Node.SelectSingleNode("@intro").text), vbCrLf, ""), "&nbsp;", ""), C_Len) & "……</span></div>"
						   echoln "   <div class=""pro"">人气：<span class=""hits"">" & Node.SelectSingleNode("@hits").text & "</span> 上传者：<span class=""inputer"">" & Node.SelectSingleNode("@inputer").text & "</span></div>"
						   echoln "   <div class=""btn""><a href=""" & DomainStr & "movie/play/?" & Node.SelectSingleNode("@id").text & """ target=""_blank""><img src=""" & DomainStr & "images/guankan.gif"" border=""0"" alt=""观看"" /></a> <a href=""" & LinkUrl & """ target=""_blank""><img src=""" & DomainStr & "images/xianqin.gif"" border=""0"" alt=""详情"" /></a></div>"
						   ElseIf PicStyle=15 Then
						   echoln "  <div class=""zy"">主演：<span>" & KS.Gottopic(Node.SelectSingleNode("@movieact").text,50) & "</></div>"
						   echoln "  <div class=""dy"">导演：<span>" & Node.SelectSingleNode("@moviedy").text & "</span></div>"
						   echoln "  <div class=""lb"">类别：<span>" & C_N_Link & "</span></div>"
						   echoln "  <div class=""text"">简介：<span class=""intro"">" & KS.GotTopic(Replace(Replace(KS.LoseHtml(Node.SelectSingleNode("@intro").text), vbCrLf, ""), "&nbsp;", ""), C_Len) & "……</span></div>"
						   echoln "   <div class=""btn""><a href=""" & DomainStr & "movie/play/?" & Node.SelectSingleNode("@id").text & """ target=""_blank""><img src=""" & DomainStr & "images/guankan.gif"" border=""0"" alt=""观看"" /></a> <a href=""" & LinkUrl & """ target=""_blank""><img src=""" & DomainStr & "images/xianqin.gif"" border=""0"" alt=""详情"" /></a></div>"
						   End If
						   echoln "  </td>"
						   echoln " </tr>"
						   echoln "</table>"
					   
					   '<!-影视系统结束->
					   '<!-供求系统开始->
					   case 16
					       echoln " <table cellSpacing=""0"" cellPadding=""0"" width=""100%"" border=""0"">"
						   echoln  "  <tr>"
						   echoln  "   <td rowspan=""2"" style=""text-align:center"" width=""" &PicWidth+10 & """><div class=""image"">" & LinkAndPicStr & "</div></td>"
						   echoln "    <td>"
						   echoln "    <span class=""t"">" & TempTitleStr & "</span></td><td width=""150"" style=""text-align:center"" class=""area"">" & ProCity & "</td><td style=""text-align:center"" width=""150"" class=""pubtime"">" & KS.GetTimeFormat(Node.SelectSingleNode("@adddate").text) & "</td>"
						   echoln "</tr>" 
						   echoln  "  <tr><td colspan=""3"">" & KS.GotTopic(Replace(Replace(KS.LoseHtml(Node.SelectSingleNode("@intro").text), vbCrLf, ""), "&nbsp;", ""), C_Len) & "……</td></tr>"
						   echoln " </table>"
					  case 17
					       echoln " <table cellSpacing=""0"" cellPadding=""0"" width=""100%"" border=""0"">"
						   echoln  "  <tr>"
						   echoln  "   <td style=""text-align:center"" width=""" &PicWidth+10 & """><div class=""image"">" & LinkAndPicStr & "</div></td>"
						   echoln "    <td>"
						   echoln "    <div class=""t"">" & TempTitleStr & "</div><div class=""text"">" & KS.GotTopic(Replace(Replace(KS.LoseHtml(Node.SelectSingleNode("@intro").text), vbCrLf, ""), "&nbsp;", ""), C_Len) & "……</div>"
						   echoln "    <div class=""pro"">更新时间:<span class=""pubtime"">" & KS.GetTimeFormat(Node.SelectSingleNode("@adddate").text) & "</span> 有效天数:<span class=""validday"">" & Node.SelectSingleNode("@validdate").text & "</span> 发布者: <span class=""pub"">"& Node.SelectSingleNode("@companyname").text & "</span></div>"
						   echoln "</td></tr>"
						   echoln " </table>"
					   '<!-供求系统结束->
					  end select
					   echoln "</td>"
					   N = N+1 : If N>=TotalNum Then Exit For
				 Next
				 echoln "</tr>"
				 If N>=TotalNum Then Exit For
		   Next
		   echoln "</table>"
		  ExplainPic = Templates
		  if err then err.clear
		End Function
		
		'得到正确的图片地址
		Function GetPicUrl(PicUrl)
		    PicUrl=trim(PicUrl)
			If KS.IsNul(PicUrl) Then PicUrl = DomainStr & "images/nopic.gif"	
			if Lcase(left(PicUrl,7))<>"http://" then GetPicUrl=KS.Setting(2) &PicUrl else GetPicUrl=PicUrl
		End Function
		
		'商城价格样式
		Function GetPriceStr(PriceType,Price_Market,Price_Member)
		      If Price_Market=0 Then Price_Market="—" Else Price_Market="￥"&FormatNumber(Price_Market,2,-1)
			  If Price_Member=0 Then Price_Member="—" Else Price_Member="￥"&FormatNumber(Price_Member,2,-1)
			  Dim VipPrice:VipPrice="<script src=""" & DomainStr & "shop/GetGroupPrice.asp?l=1&t=p&proid=" & Node.SelectSingleNode("@id").text & """></script>"
		     Select Case PriceType
			  Case 0:GetPriceStr="市场价："&Price_Market &"<br />商城价："&Price_Member &"<br />VIP 价：" & VipPrice
			  Case 1:GetPriceStr="商城价："&Price_Member
			  Case 2:GetPriceStr="参考价："&Price_Market
			  Case 3:GetPriceStr="VIP 价："&VipPrice
			  Case 4:GetPriceStr="参考价："&Price_Market & "<br />商城价：" & Price_Member
			  Case 5:GetPriceStr="参考价："&Price_Market & "<br />商城价：" & Price_Member & "<br/>VIP 价：" & VipPrice
			  Case 6:GetPriceStr="参考价："&Price_Market & "<br/>VIP 价：" & VipPrice
			  Case 7:GetPriceStr="商城价："&Price_Member & "<br/>VIP 价：" & VipPrice
			 End Select
		End Function
		'商城按钮样式
		Function GetButtonStr(ButtonType,ID,Url,O_T_S)
		          Dim BuyButton:BuyButton="<a href=""" & DomainStr & "Shop/ShoppingCart.asp?Action=Add&ID=" &ID &""" "&O_T_S &"><img src=""" & DomainStr & "images/productbuy.gif"" alt=""购买"" border=""0""/></a>"
				  Dim FavButton:FavButton="<a href=""" & DomainStr & "user/user_Favorite.asp?Action=Add&ChannelID=5&InfoID=" & ID &""" target=""_blank""><img src=""" & DomainStr & "images/productfav.gif"" alt=""收藏"" border=""0""/></a>"
				  Dim XQButton:XQButton="<a href=""" & Url&""""&O_T_S &"><img alt=""详情"" src=""" & DomainStr & "images/productxq.gif"" border=""0""/></a>"
		     Select Case ButtonType
					  Case 1:GetButtonStr=BuyButton
                      Case 2:GetButtonStr=FavButton
					  Case 3:GetButtonStr=XQButton
					  Case 4:GetButtonStr=BuyButton&" " & FavButton
					  Case 5:GetButtonStr=BuyButton&" " & XQButton
					  Case 6:GetButtonStr=FavButton&" " & XQButton
					  Case 7:GetButtonStr=BuyButton&" " & XQButton&" " & FavButton
					  Case Else:GetButtonStr=""
			 End Select
	  End Function

		
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:GetSlide
		'作 用: 通用幻灯函数
		'参 数: LabelStyle 标签样式
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function GetSlide(LabelStyle)
		     LoadLabelParam
			 If LoadSucceed = false Then GetSlide="标签加载出错!":Exit Function 
			 Dim SqlStr,TableName,FieldStr
			 
			 If ModelID="0" Then 
			  TableName = "[KS_ItemInfo]" : FieldStr="I.InfoID as ID,I.Title,I.ChannelID,I.Tid,I.PhotoUrl,I.Fname"
			  Param = Param & " and PhotoUrl<>''"
			 ElseIf ModelID="-1000" Then
			  TableName = "[KS_GuestBook]" : FieldStr="id,Subject as title,-1000 as channelid,0 as  tid,face as photourl,id as fname" : Param=" Where verific=1 and isslide=1 and deltf=0"
			 Else
			  TableName = KS.C_S(ModelID,2) : FieldStr="I.ID,I.Title,I.Tid,I.PhotoUrl,I.Fname"
			  If KS.C_S(ModelID,6)=1 Then Param=Param & " and I.PicNews=1"
			  If KS.C_S(ModelID,6)=5 Then FieldStr="I.ID,I.Title,I.Tid,I.BigPhoto as PhotoUrl,I.Fname"
			 End If
			 SqlStr = "SELECT TOP " & Num & " " & FieldStr & " FROM " & TableName & " I " & Param & " ORDER BY I.ID Desc"
			 Dim RS:Set RS=Conn.Execute(SqlStr)
			 If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			 RS.Close:Set RS=Nothing
			 If IsObject(XMLSql) Then
				 GetSlide=ExplainSlideLabelBody()
			 End If 
			 Set Node=Nothing
		End Function
        '解释幻灯标签体
		Function ExplainSlideLabelBody()
		     Dim SlideType,T_Len,Width,Height,ShowTitle,ChangeTime,O_T_S,T_CssStr,T_Css
		     Dim DocNode,TotalNum,K,Title,TempTitle,PhotoUrl,LinkUrl,ArrLength
			 Set DocNode=XMLSql.DocumentElement.SelectNodes("row")
		     TotalNum=DocNode.length
			 SlideType = ParamNode.getAttribute("slidetype") : If Not IsNumeric(SlideType) Then SlideType=1
			 T_Len     = KS.ChkClng(ParamNode.getAttribute("titlelen")) 
			 Width     = ParamNode.getAttribute("picwidth")     : If Not IsNumeric(Width) Then Width=200
			 Height    = ParamNode.getAttribute("picheight")    : If Not IsNumeric(Height) Then Height=200
			 ShowTitle = ParamNode.getAttribute("showtitle") 
			 ChangeTime= ParamNode.getAttribute("changetime")
			 O_T_S     = KS.G_O_T_S(ParamNode.getAttribute("opentype"))
			 T_Css     = ParamNode.getAttribute("titlecss")
			 Templates = ""
			 IF Cint(SlideType)<>1 Then 
					 Dim ImgArrStr,LinkArrStr,TextArrStr
					 N=0
					 For K=0 To TotalNum-1
					      Set Node=DocNode.Item(n)
					      Title         = Node.SelectSingleNode("@title").text
					      PhotoUrl      = Node.SelectSingleNode("@photourl").text : PhotoUrl= GetPicUrl(PhotoUrl)
						  If Cint(ModelID)=-1000 Then
						      LinkUrl= KS.GetClubShowUrl(Node.SelectSingleNode("@id").text)
						  Else
							  If ModelID=0 Then CurrModelID=Cint(Node.SelectSingleNode("@channelid").text) Else CurrModelID=ModelID
							  LinkUrl       = KS.GetItemURL(CurrModelID,Node.SelectSingleNode("@tid").text,Node.SelectSingleNode("@id").text,Node.SelectSingleNode("@fname").text)
						  End If
						  TempTitle     = KS.Gottopic(Title,T_Len)
                          if N=0 Then
						    ImgArrStr=PhotoUrl : LinkArrStr=LinkUrl :  TextArrStr=TempTitle 
					      Else
						   ImgArrStr=ImgArrStr & "@@@" & PhotoUrl
						   LinkArrStr=LinkArrStr & "@@@" & LinkUrl
						   TextArrStr=TextArrStr & "@@@" & TempTitle
						  End if
						  N=N+1
                    Next
					 Dim ImgArr:ImgArr=Split(ImgArrStr,"@@@")
					 Dim LinkArr:LinkArr=Split(LinkArrStr,"@@@")
					 Dim TextArr:TextArr=Split(TextArrStr,"@@@")
				Select Case Cint(SlideType)
				  case 2
					echoln "<script src=""" & DomainStr &"ks_inc/loadflash.js"" type=""text/javascript""></script>"
					echoln "<script language=""JavaScript"" type=""text/javascript"">"
					echoln "<!--"
					echoln "var focus_width=" & Width & ";" 
					echoln "var focus_height=" & Height & ";" 
					If Cbool(ShowTitle)=True Then
					echoln "var text_height=22;"
					Else
					echoln "var text_height=0;"
					End If
					 ArrLength=Ubound(ImgArr)
					 If ArrLength>5 Then ArrLength=5
					 Dim I,PicStr,LinkStr,TextStr
					 For I=0 To ArrLength
					   If I=0 Then
						PicStr="var pics='" & ImgArr(0) : LinkStr="var links=escape('" & LinkArr(0) : TextStr="var texts='" & TextArr(0)
					   Else
					    PicStr=PicStr & "|" & ImgArr(I) : LinkStr=LinkStr&"|"&LinkArr(I) : TextStr=TextStr & "|" & TextArr(I)
					   End IF
					 Next
					echoln PicStr &"';"&vbcrlf&LinkStr&"');" &vbcrlf & TextStr &"';"
					echoln "LoadFlash('" & DomainStr & "KS_Inc/Slideviewer.swf','transparent',focus_width,focus_height+text_height,'pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height);"
					echoln "//-->"
					echoln "</script>"
			     case 3
					 ArrLength=Ubound(ImgArr)
					 If ArrLength>5 Then ArrLength=5
					 For I=0 To ArrLength
					   If I=0 Then
						PicStr=ImgArr(0) & "#" & TextArr(0) & "#" & LinkArr(0)
					   Else
					    PicStr=PicStr & "|" & ImgArr(i) & "#" & TextArr(i) & "#" & LinkArr(i)
					   End IF
					 Next
					echoln "<script src=""" & DomainStr &"ks_inc/sinaflash.js"" type=""text/javascript""></script>"
					echoln "<div id=""flashcontent"">"
					echoln "<script type=""text/javascript"">"
					echoln "<!--"
					echoln "var sinaFlash2 = new sinaFlash(""" & DomainStr & "KS_Inc/sinaflash.swf"",""demo"", "& width &", "& height &", """", ""#ffffff"");"
					echoln "sinaFlash2.addParam(""quality"", ""best"");" 
					echoln "sinaFlash2.addParam(""wmode"", ""transparent"");"
					echoln "sinaFlash2.addVariable(""picaddress"",escape("""&PicStr&"""));"
					echoln "sinaFlash2.addVariable(""speed"","""& ChangeTime &""");" 
					echoln "sinaFlash2.write(""flashcontent"");"
					echoln "//-->"
					echoln "</script>"
					echoln "</div>"
				case 4
					echoln "<script src=""" & DomainStr &"ks_inc/sohuflash.js"" type=""text/javascript""></script>"
					echoln "<div id=""sasFlashFocus3"">"
					echoln "<script type=""text/javascript"">"
					echoln "<!--"
					echoln "var pics3="""", mylinks3="""", texts3="""";" 
					echoln "var focus_width=" & Width & ";"
					echoln "var focus_height=" & Height & ";"
					If Cbool(ShowTitle)=True Then
					echoln "var text_height=22;"
					Else
					echoln "var text_height=0;"
					End If
					 ArrLength=Ubound(ImgArr)
					 If ArrLength>5 Then ArrLength=5
					 For I=0 To ArrLength
					   If I=0 Then
						 PicStr  = "var pics='" & ImgArr(0)
						 LinkStr = "var mylinks='" & LinkArr(0)
						 TextStr = "var texts='" & TextArr(0) &""
					   Else
					     PicStr  = PicStr & "|" & ImgArr(I)
					     LinkStr = LinkStr&"|" & LinkArr(I)
					     TextStr = TextStr & "|" & TextArr(I)
					  End IF
					 Next
					echoln PicStr &"';"&vbcrlf&LinkStr&"';" &vbcrlf & TextStr &"';"
					echoln "var easytool2 = new easytool(""" & DomainStr & "KS_Inc/sohuflash.swf"",""sasFlashFocus3"", "& width &", "& height &", ""6"");"
					echoln "easytool2.addParam(""quality"", ""high"");"
					echoln "easytool2.addParam(""wmode"", ""opaque"");"
					echoln "easytool2.addVariable(""pics2"",pics);"
					echoln "easytool2.addVariable(""links2"",mylinks);"
					echoln "easytool2.addVariable(""texts2"",texts);"
					echoln "easytool2.write(""sasFlashFocus3"");"
					echoln "//-->"
					echoln "</script>"
					echoln "</div>"
				 case 5
				    echoln "<script src=""" & DomainStr &"ks_inc/sohuflash_1.js"" type=""text/javascript""></script>"
					echoln "<div id=""flashcontent01""></div>"
					echoln "<script type=""text/javascript"">"
					 
					 ArrLength=Ubound(ImgArr)
					 If ArrLength>5 Then ArrLength=5
					 For I=0 To ArrLength
					   If I=0 Then
						PicStr="var pics='" & ImgArr(0) : LinkStr="var links=escape('" & LinkArr(0) : TextStr="var texts='" & TextArr(0)
					   Else
					    PicStr=PicStr & "|" & ImgArr(I) : LinkStr=LinkStr&"|"&LinkArr(I) : TextStr=TextStr & "|" & TextArr(I)
					   End IF
					 Next
					echoln PicStr &"';"&vbcrlf&LinkStr&"');" &vbcrlf & TextStr &"';" 
					echoln "var speed = 4000;"
					echoln "var sohuFlash2 = new sohuFlash(""" & DomainStr &"ks_inc/focus0414a.swf"",""flashcontent01"",""" & Width & """,""" & Height & """,""8"",""#ffffff"");"
					echoln "sohuFlash2.addParam(""quality"", ""medium"");"
					echoln "sohuFlash2.addParam(""wmode"", ""opaque"");"
					echoln "sohuFlash2.addVariable(""speed"",speed);"
					echoln "sohuFlash2.addVariable(""p"",pics);	"
					echoln "sohuFlash2.addVariable(""l"",links);"
					echoln "sohuFlash2.addVariable(""icon"",texts);"
					echoln "sohuFlash2.write(""flashcontent01"");"
					echoln "</script>"
				
				End Select
					
						 
			 Else
				echoln "<script language=""javascript"" type=""text/javascript"">"
				echoln ("<!--")
				echoln ("function SlidePic1(ID) {this.ID=ID; this.Width=0;this.Height=0; this.TimeOut=5000; this.Effect=23; this.T_Len=0; this.PicNum=-1; this.Img=null; this.Url=null; this.Title=null; this.AllPic=new Array(); this.Add=AddSlidePic1; this.Show=ShowSlidePic1; this.LoopShow=LoopShowSlidePic1;}")
				echoln ("function NewSlide1() {this.ImgUrl=""""; this.LinkUrl=""""; this.Title="""";}")
				echoln ("function AddSlidePic1(SP) {this.AllPic[this.AllPic.length] = SP;}")
				echoln ("function ShowSlidePic1() {")
				echoln ("if(this.AllPic[0] == null) return false;")
				echoln ("document.write('<div align=""center""><a id=""Url' + this.ID + '"" href=""""" & O_T_S & "><img id=""Img' + this.ID + '"" width=' + this.Width + '  height=' + this.Height + ' style=""filter: revealTrans(duration=2,transition=23);"" src=""javascript:null"" border=""0""></a>');") 
				echoln ("if(this.T_Len != 0) document.write(""<br><Div id='Title"" + this.ID + ""'></Div></div>"");")
				echoln ("this.Img = document.getElementById(""Img"" + this.ID);")
				echoln ("this.Url = document.getElementById(""Url"" + this.ID);")
				echoln ("this.Title = document.getElementById(""Title"" + this.ID);")
				echoln ("this.LoopShow();")
				echoln ("}")
				echoln ("function LoopShowSlidePic1() {")
				echoln ("if(this.PicNum<this.AllPic.length-1) this.PicNum++ ;")
				echoln ("else this.PicNum=0;")
				echoln ("this.Img.src=this.AllPic[this.PicNum].ImgUrl;")
				echoln ("if (document.all){")
				echoln ("this.Img.filters.revealTrans.Transition=this.Effect;")
				echoln ("if(navigator.userAgent.indexOf('MSIE 6.0')==-1){this.Img.filters.revealTrans.apply();}")
				echoln ("this.Img.filters.revealTrans.play();}")
				echoln ("this.Url.href=this.AllPic[this.PicNum].LinkUrl;")
				echoln ("if(this.Title) this.Title.innerHTML='<a href=""'+this.AllPic[this.PicNum].LinkUrl+'"" " & O_T_S & ">'+this.AllPic[this.PicNum].Title+'</a>';")
				echoln ("this.Img.timer=setTimeout(this.ID+"".LoopShow()"",this.TimeOut);")
				echoln ("}")
    
					   '新建幻灯片图片对象
					echoln ("var SlidePic1 = new SlidePic1(""SlidePic1"");")
					echoln ("SlidePic1.Width    = " & Width & ";")
					echoln ("SlidePic1.Height   = " & Height & ";")
					echoln ("SlidePic1.TimeOut  = " & ChangeTime & ";")
					echoln ("SlidePic1.Effect   = 23;")
					   If CBool(ShowTitle) = False Then
						 echoln ("SlidePic1.T_Len = 0;")
					   Else
						 echoln ("SlidePic1.T_Len = 1;")
					   End If
						T_CssStr = KS.GetCss(T_Css)
					   For K=0 To TotalNum-1
					    Set Node=DocNode.Item(k)
						PhotoUrl      = Node.SelectSingleNode("@photourl").text : PhotoUrl= GetPicUrl(PhotoUrl)
						If Cint(ModelID)=-1000 Then
						    LinkUrl= KS.GetClubShowUrl(Node.SelectSingleNode("@id").text)
						Else
						If ModelID=0 Then CurrModelID=Cint(Node.SelectSingleNode("@channelid").text) Else CurrModelID=ModelID
						LinkUrl       = KS.GetItemURL(CurrModelID,Node.SelectSingleNode("@tid").text,Node.SelectSingleNode("@id").text,Node.SelectSingleNode("@fname").text)
						End If
						TempTitle = "<span" & T_CssStr & ">" &KS.GotTopic(KS.LoseHtml(Node.SelectSingleNode("@title").text) ,T_Len)& "</span>"
						echoln "var NewItem = new NewSlide1();"
						echoln "NewItem.ImgUrl = '" & PhotoUrl & "';"
						echoln "NewItem.LinkUrl= '" & LinkUrl & "';"
						echoln "NewItem.Title = '" & TempTitle & "';"
						echoln "SlidePic1.Add(NewItem);"
                       Next		
					   echoln ("SlidePic1.Show();")
					   echoln ("//-->")
					   echoln ("</Script>")

				End if
				ExplainSlideLabelBody = Templates
		End Function		
		
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:GetRolls
		'功能:取得连续滚动图片
		'参数:LabelStyle 标签样式
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function GetRolls(LabelStyle)
		     LoadLabelParam
			 If LoadSucceed = false Then GetSlide="标签加载出错!":Exit Function 
			 Dim TableName,FieldStr,SqlStr
			If ModelID="0" Then 
			 TableName = "[KS_ItemInfo]" 
			 FieldStr  = "I.ChannelID,I.InfoID as id,I.Title,I.Tid,I.PhotoUrl,I.Fname"
			 Param = Param & " and I.PhotoUrl<>''"
			Else 
			 TableName=KS.C_S(ModelID,2)
			 FieldStr  = "I.ID,I.Title,I.Tid,I.PhotoUrl,I.Fname"
			 If KS.C_S(ModelID,6)=1 Then Param=Param & " and I.PicNews=1"
			End If
			SqlStr="Select Top " & Num & " " & FieldStr & " From " & TableName & " I " & Param & " Order By " & GetOrderParam
			Dim RS:Set RS=Conn.Execute(SqlStr)
			 If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			 RS.Close:Set RS=Nothing
			 If IsObject(XMLSql) Then
				 GetRolls=ExplainRollsLabelBody()
			 End If 
			 Set Node=Nothing
		End Function
		
		'作 用: 解释通用连续滚动函数体
	   Function ExplainRollsLabelBody()
	         Dim TempPicStr, T_CssStr,Title,T_Css,TempTitleStr,T_Len, O_T_S,M_Dir,LinkUrl,ShowTitle,LinkAndPicStr,Marqueebgcolor
			 Dim K,TotalNum,PicBorderColor,PicBorderColorStr,PicWidth,PicHeight,PicWidthStr,PicHeightStr,M_Width,M_Height,M_Speed,MarqueeType,MarqueeStyle,TemplateFromXml,DateStr,NaviStr
			 Set DocNode=XMLSql.DocumentElement.SelectNodes("row")
		     TotalNum=DocNode.length
			 MarqueeType = ParamNode.getAttribute("marqueetype") : If KS.IsNul(MarqueeType) Then MarqueeType="pic"
			 T_Css   = ParamNode.getAttribute("titlecss") : T_CssStr = KS.GetCss(T_Css)
			 O_T_S   = KS.G_O_T_S(ParamNode.getAttribute("opentype"))
			 M_Dir	 = ParamNode.getAttribute("marqueedirection")
			 PicBorderColor = ParamNode.getAttribute("picbordercolor")
			 MarqueeBgColor = ParamNode.getAttribute("marqueebgcolor")
			 PicWidth  = KS.ChkClng(ParamNode.getAttribute("picwidth")):If PicWidth<>0 Then PicWidthStr=" width=""" & PicWidth & """" Else PicWidthStr=""
		     PicHeight = KS.ChkClng(ParamNode.getAttribute("picheight")):If PicHeight<>0 Then PicHeightStr=" height=""" & PicHeight & """" Else PicHeightStr=""
			 M_Width   = ParamNode.getAttribute("marqueewidth")
			 M_Height  = ParamNode.getAttribute("marqueeheight")
			 M_Speed   = ParamNode.getAttribute("marqueespeed")
             ShowTitle = ParamNode.getAttribute("showtitle")
			 T_Len     = ParamNode.getAttribute("titlelen")
			 MarqueeStyle = ParamNode.getAttribute("marqueestyle")
			 NaviStr   = KS.GetNavi(ParamNode.getAttribute("navtype"), ParamNode.getAttribute("nav"))
			 templates = "" : N=0
			    If MarqueeStyle="1" Then  '纵向不间断
				   For K=0 To TotalNum-1
					 Set Node=DocNode.Item(n)
					 Title      = Node.SelectSingleNode("@title").text
					 If ModelID=0 Then CurrModelID=Cint(Node.SelectSingleNode("@channelid").text) Else CurrModelID=ModelID
					 LinkUrl    = KS.GetItemURL(CurrModelID,Node.SelectSingleNode("@tid").text,Node.SelectSingleNode("@id").text,Node.SelectSingleNode("@fname").text)
					 TempTitleStr = "<a" & T_CssStr & " href=""" & LinkUrl & """" & O_T_S & " title=""" & Title & """>" & KS.GotTopic(Title, T_Len) & "</a>"
					 DateStr="<span" & KS.GetCss(ParamNode.getAttribute("datecss")) & ">" & LFCls.Get_Date_Field(Node.SelectSingleNode("@adddate").text, ParamNode.getAttribute("daterule")) & "</span>"
					  echoln "<li style=""height:" & M_Height & "px;line-height:" & M_Height & "px"">" & NaviStr & TempTitleStr & DateStr & "</li>"
					  n=n+1
				  Next 
				  TemplateFromXml=LFCls.GetConfigFromXML("Label","/labeltemplate/label","rollvertical")
				  TemplateFromXml=Replace(Replace(Replace(Replace(Replace(TemplateFromXml,"{$Width}",M_Width),"{$Height}",M_Height),"{$LoopStr}",Templates),"{$Speed}",M_Speed),"{$LabelID}",LabelID)
				Else
					 If LCase(M_Dir) = "left" Or LCase(M_Dir) = "right" Then
							   echoln "<table width=""100%"" height=""100%"" border=""0"">"
							   echoln " <tr>"
							  For K=0 To TotalNum-1
								 Set Node=DocNode.Item(n)
								 Title      = Node.SelectSingleNode("@title").text
								 If ModelID=0 Then CurrModelID=Cint(Node.SelectSingleNode("@channelid").text) Else CurrModelID=ModelID
								 LinkUrl    = KS.GetItemURL(CurrModelID,Node.SelectSingleNode("@tid").text,Node.SelectSingleNode("@id").text,Node.SelectSingleNode("@fname").text)
	
								 TempTitleStr = "<a" & T_CssStr & " href=""" & LinkUrl & """" & O_T_S & " title=""" & Title & """>" & KS.GotTopic(Title, T_Len) & "</a>"
								If MarqueeType="pic" Then
									 TempPicStr = GetPicUrl(Node.SelectSingleNode("@photourl").text)
									 If PicBorderColor<>"" Then PicBorderColorStr=" style=""border:1px solid "& PicBorderColor & """"
									 LinkAndPicStr = "<a href=""" & LinkUrl & """" & O_T_S & " title=""" & Title & """><Img Src=""" & TempPicStr & """ border=""0"" alt=""" & Title & """" & PicWidthStr & PicHeightStr & PicBorderColorStr&" align=""absmiddle""/></a>"
								
									 echoln "<td style=""text-align:center"">"
									 echoln " <div class=""img"">" & LinkAndPicStr & "</div>"
									If Cbool(ShowTitle) = True Then
									 echoln " <div class=""t"">" & TempTitleStr & " </div>"
									End If
									 echoln "</td>"
								Else
									 DateStr="<span" & KS.GetCss(ParamNode.getAttribute("datecss")) & ">" & LFCls.Get_Date_Field(Node.SelectSingleNode("@adddate").text, ParamNode.getAttribute("daterule")) & "</span>"
									 echoln "<td nowrap=""nowrap"" class=""rolltext"">" & NaviStr & TempTitleStr & DateStr & "</td>"
								End If
								 n=n+1
						   Next
								 echoln "</tr></table>"
					Else
							If MarqueeType="pic" Then echoln "<table width=""100%"" height=""100%"" border=""0"">"
							For K=0 To TotalNum-1
								 Set Node=DocNode.Item(n)
								 Title      = Node.SelectSingleNode("@title").text
								 
								 If ModelID=0 Then CurrModelID=Cint(Node.SelectSingleNode("@channelid").text) Else CurrModelID=ModelID
								 LinkUrl    = KS.GetItemURL(CurrModelID,Node.SelectSingleNode("@tid").text,Node.SelectSingleNode("@id").text,Node.SelectSingleNode("@fname").text)
								
								 TempTitleStr = "<a" & T_CssStr & " href=""" & LinkUrl & """" & O_T_S & " title=""" & Title & """>" & KS.GotTopic(Title, T_Len) & "</a>"
							  If MarqueeType="pic" Then
								 TempPicStr = GetPicUrl(Node.SelectSingleNode("@photourl").text)
								 If PicBorderColor<>"" Then PicBorderColorStr=" style=""border:1px solid "& PicBorderColor & """"
								 LinkAndPicStr = "<a href=""" & LinkUrl & """" & O_T_S & " title=""" & Title & """><Img Src=""" & TempPicStr & """ border=""0"" alt=""" & Title & """" & PicWidthStr & PicHeightStr & """" & PicBorderColorStr &" align=""absmiddle""/></a>"
								echoln "<tr><td style=""text-align:center"">"
								echoln "<div class=""image"">" & LinkAndPicStr & "</div>"
								If Cbool(ShowTitle) = True Then echoln "<div class=""t"">" & TempTitleStr & " </div>"
								echoln "</td></tr>"
							  Else
								DateStr="<span" & KS.GetCss(ParamNode.getAttribute("datecss")) & ">" & LFCls.Get_Date_Field(Node.SelectSingleNode("@adddate").text, ParamNode.getAttribute("daterule")) & "</span>"
								 echoln "<div class=""rolltext"">" & NaviStr & TempTitleStr & DateStr & "</div>"
							  End If
								n=n+1
							Next
							If MarqueeType="pic" Then echoln "</table>"
				   End If		
					TemplateFromXml=LFCls.GetConfigFromXML("Label","/labeltemplate/label","roll"&M_Dir)
					If Not KS.IsNul(MarqueeBgColor) Then
					TemplateFromXml=Replace(TemplateFromXml,"{$BackGround}","background:" & MarqueeBgColor & ";")
					Else
					TemplateFromXml=Replace(TemplateFromXml,"{$BackGround}","")
					End If
					TemplateFromXml=Replace(Replace(Replace(Replace(Replace(TemplateFromXml,"{$Width}",M_Width),"{$Height}",M_Height),"{$ImgStr}",Templates),"{$Speed}",M_Speed),"{$LabelID}",LabelID)
				End If
				ExplainRollsLabelBody=TemplateFromXml
		End Function
		
		
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:GetMarquee
		'功能:取得连续文字滚动
		'参数:LabelStyle 标签样式
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function GetMarquee(LabelStyle)
		    LoadLabelParam
			If LoadSucceed = false Then GetMarquee="标签加载出错!":Exit Function 
			
			Dim TableName,FieldStr,SqlStr
			If ModelID="0" Then 
			 TableName = "[KS_ItemInfo]" 
			 FieldStr  = "I.ChannelID,I.InfoID as id,I.Title,I.Tid,I.AddDate,I.Fname"
			 Param = Param & " and I.PhotoUrl<>''"
			Else 
			 TableName=KS.C_S(ModelID,2)
			 FieldStr  = "I.ID,I.Title,I.Tid,I.AddDate,I.Fname"
			End If
			SqlStr="Select Top " & Num & " " & FieldStr & " From " & TableName & " I " & Param & " Order By " & GetOrderParam
			 Dim RS:Set RS=Conn.Execute(SqlStr)
			 If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			 RS.Close:Set RS=Nothing
			 If IsObject(XMLSql) Then
				GetMarquee=ExplainRollsLabelBody()
			 End If 
			 Set Node=Nothing
		End Function
		
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:GetNotRuleList
		'功  能:取得不规则列表
		'参  数:LabelStyle 标签样式
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function GetNotRuleList(LabelStyle)
		     LoadLabelParam
			 If LoadSucceed = false Then GetNotRuleList="标签加载出错!":Exit Function 
			 If LabelID<>"ajax" and Cbool(AjaxOut)=true Then 
			  GetNotRuleList="<span id=""ks" & LabelID & "_" & ParamNode.getAttribute("classid") & "_" & FCls.RefreshFolderID & "_0_" & FCls.ChannelID & """></span>":Exit Function
			 End If
			 
			Dim TableName,FieldStr,SqlStr
			If ModelID="0" Then 
			 TableName = "[KS_ItemInfo]" 
			 FieldStr  = "I.ChannelID,I.InfoID as id,I.Title,I.Tid,I.comment,I.Fname"
			Else 
			 TableName=KS.C_S(ModelID,2)
			 FieldStr  = "I.ID,I.Title,I.Tid,I.comment,I.Fname"
			 If KS.C_S(ModelID,6)="1" Then FieldStr=FieldStr & ",I.TitleType,I.TitleFontColor,I.TitleFontType"
			End If
			Dim AllowMaxNum:AllowMaxNum=200   '限定允许在200条，内调用
			SqlStr="Select Top " & AllowMaxNum & " " & FieldStr & " From " & TableName & " I " &  Param & " Order By " & GetOrderParam
			 Dim RS:Set RS=Conn.Execute(SqlStr)
			 If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			 RS.Close:Set RS=Nothing
			 If IsObject(XMLSql) Then
				GetNotRuleList=ExplainNotRuleLabelBody()
			 End If 
			 Set Node=Nothing
		End Function
		
		'解释不规则标签体
		Function ExplainNotRuleLabelBody()
		  	 Dim I,P_T,O_T_S ,C_N_Link,K,TotalNum,T_Css,T_CssStr,R_H,NavType,Nav,NaviStr,RowNumber,ShowNumPerRow
			 Dim PreComment,PreShowComment,PreClassID,PreInfoID,SplitPic,MoreLink,C_F_T,M_L_S,MoreType
			 Set DocNode=XMLSql.DocumentElement.SelectNodes("row")
		     TotalNum=DocNode.length
			 P_T      = ParamNode.getAttribute("printtype")  : If Not IsNumeric(P_T) Then P_T = 1
			 T_Css    = ParamNode.getAttribute("titlecss")
			 R_H      = ParamNode.getAttribute("rowheight")
			 O_T_S    = KS.G_O_T_S(ParamNode.getAttribute("opentype"))
			 NavType  = ParamNode.getAttribute("navtype")
			 Nav      = ParamNode.getAttribute("nav")
			 RowNumber= ParamNode.getAttribute("rownumber")  : If Not IsNumeric(RowNumber) Then RowNumber=10
			 ShowNumPerRow = ParamNode.getAttribute("shownumperrow") : If Not IsNumeric(ShowNumPerRow) Then ShowNumPerRow=50
			 SplitPic = ParamNode.getAttribute("splitpic")
			 MoreLink = ParamNode.getAttribute("morelink")
			 MoreType = ParamNode.getAttribute("morelinktype")
			 
			 If ClassID = "-1" Or Instr(ClassID,",")<>0 Then C_F_T = True Else C_F_T = False
		     If MoreLink <> "" And ClassID <> "0" And C_F_T = False Then M_L_S = KS.GetMoreLink(1,1, R_H, MoreType, MoreLink, KS.GetFolderPath(ClassID), O_T_S)
			    Dim CurrTid,LinkStr,Title,EndStr
				T_CssStr = KS.GetCss(T_Css):R_H = KS.G_R_H(R_H):NaviStr = KS.GetNavi(NavType, Nav)
				Templates = "" : N=0

				If Cint(P_T)=2 Then
				 echoln "<li>" : EndStr="</li>"
				Else
				 echoln "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" align=""center"">" & vbCrLf & "<tr><td height=""" & R_H &""">" : EndStr="</td></tr>"
			   End If
			   
				Dim II:ii=0:Dim CC:cc=0:Dim Row,str
				RowNumber=Cint(RowNumber):ShowNumPerRow=Cint(ShowNumPerRow)
				echo NaviStr

				For K=0 To TotalNum-1
				    Set Node=DocNode.Item(n)
				    CurrTid = Node.SelectSingleNode("@tid").text:Title = Trim(Node.SelectSingleNode("@title").text)
					If ModelID=0 Then CurrModelID=Cint(Node.SelectSingleNode("@channelid").text) Else CurrModelID=ModelID
					LinkStr=T_CssStr & " href=""" & KS.GetItemURL(CurrModelID,Node.SelectSingleNode("@tid").text,Node.SelectSingleNode("@id").text,Node.SelectSingleNode("@fname").text) & """" & O_T_S & " title=""" & Title & """"
					ii=ii + KS.strLength(Title)
					if ii>=ShowNumPerRow then
					  cc=ii - ShowNumPerRow:cc=KS.strLength(Title) - cc:row=row+1:ii=0
					  if cc=0 then cc=1
					  IF Cint(row)=Cint(RowNumber) or n>=TotalNum-1 then
					         If KS.C_S(ModelID,6)="1" Then
							 echo "<a" & LinkStr &">"& GetItemTitle(Title,cc,true,Node.SelectSingleNode("@titletype").text, Node.SelectSingleNode("@titlefontcolor").text, Node.SelectSingleNode("@titlefonttype").text)&"</a>"&EndStr
							 Else
							 echo "<a" & LinkStr &">"& KS.GotTopic(Title,cc)&"</a>"&EndStr
							 End If
					  Else
					        If KS.C_S(ModelID,6)="1" Then
							 echo "<a" & LinkStr &">"& GetItemTitle(Title,cc,true,Node.SelectSingleNode("@titletype").text, Node.SelectSingleNode("@titlefontcolor").text, Node.SelectSingleNode("@titlefonttype").text)&"</a>"&EndStr
							Else
						     echo "<a" & LinkStr &">"& KS.GotTopic(Title,cc)&"</a>"&EndStr
							End If
						  If Cint(P_T)=2 Then
						    echo "<li>" & NaviStr
						  else
						    echo (KS.GetSplitPic(SplitPic, 1))
							echoln ""
						    echo  "<tr><td height=""" & R_H &""">" & NaviStr
						  end if
					  End If
					Else
					   If KS.C_S(ModelID,6)="1" Then
					   echoln "<a" & LinkStr &">"&GetItemTitle(Title,0,true,Node.SelectSingleNode("@titletype").text, Node.SelectSingleNode("@titlefontcolor").text, Node.SelectSingleNode("@titlefonttype").text)&"</a> "
					   Else
					   echoln "<a" & LinkStr &">"& Title&"</a> "
					   End If
					   ii=ii + 1
					End IF
					n=n+1
					if cint(row)>=cint(RowNumber) or n>=TotalNum then exit For
				Next
				 If Cint(P_T)=2 Then
				  echo M_L_S
				 Else
				  echo (KS.GetSplitPic(SplitPic, 1))
				  echo M_L_S
				  echoln "</table>"
				 End if
		  ExplainNotRuleLabelBody=Templates
		End Function
		
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:GetCirClassList
		'作 用: 循环栏目列表标签函数
		'参 数: LabelStyle 标签样式
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function GetCirClassList(LabelStyle)
		     If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				GetCirClassList = "标签加载出错!" : Exit Function
			 Else
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
 			 Dim ClassParam,SQLStr,ClassXml,ClassNode,I,ClassStyle,DocStyle,ClassStr,LoopClassStyle,ID,ClassBasicInfoArr,ClassPrintType

		     LabelID   = ParamNode.getAttribute("labelid")
             ClassID   = ParamNode.getAttribute("classid")
			 ClassPrintType=ParamNode.getAttribute("classprinttype") : If Not IsNumeric(ClassPrintType) Then ClassPrintType=1
			 
			 If LabelID<>"ajax" and Cbool(ParamNode.getAttribute("ajaxout"))=true Then 
			  GetCirClassList="<span id=""ks" & LabelID & "_" & ParamNode.getAttribute("classid") & "_" & FCls.RefreshFolderID & "_0_0""></span>":Exit Function
			 End If
			 
			 ClassParam  =" Where ClassType=1"
			
			 If InStr(ClassID,",")<>0 Then
			  ClassParam = ClassParam & " And ID in('" & Replace(ClassID,",","','")& "')" 
			 ElseIf ClassID="-1" Then
			  ClassParam = ClassParam & " And TN='" & FCls.RefreshFolderID & "'" 
			 ElseIf CLassID="-2" THEN
			  ClassParam = ClassParam & " And tj=1 and channelid=" & KS.ChkClng(Fcls.ChannelID)
			 Else
			  ClassParam = ClassParam & " And TN='" & ClassID & "'" 
			 End If
			 
			 SQLStr="Select Top 50 ID From KS_Class " & ClassParam & " Order By root,folderorder"
			 
			 SQLStr="Select Top 50 ID From KS_Class " & ClassParam & " Order By root,folderorder"
			 
			 Dim classObj:Set classObj=Conn.Execute(SQLStr)
             If Not classObj.Eof Then
			  Set ClassXml=KS.RsToXml(classObj,"row","root")
			 End If
			 classObj.Close:Set classObj=Nothing
		     If IsObject(classXml) Then
			   	 regEx.Pattern="\bajaxout="".*?"""
				 Set Matches = regEx.Execute(LabelParamStr)
				 If Matches.count > 0 Then LabelParamStr=Replace(LabelParamStr,Matches.item(0),"ajaxout=""false""")
			   ClassStyle = Split(LabelStyle,"§")(0)
			   DocStyle = Split(LabelStyle,"§")(1)
			   If ClassPrintType=2 Then
			       Dim ClassRow:ClassRow=0
				   For Each ClassNode In classXml.DocumentElement.SelectNodes("row")
					 ID = ClassNode.SelectSingleNode("@id").text
					 regEx.Pattern="\bclassid="".*?"""
					 Set Matches = regEx.Execute(LabelParamStr)
					 If Matches.count > 0 Then LabelParamStr=Replace(LabelParamStr,Matches.item(0),"classid=""" & ID & """")
						 LoopClassStyle = ClassStyle
						 ClassRow=ClassRow+1
						 ClassBasicInfoArr = Split(KS.C_C(ID,6),"||||")
						 LoopClassStyle = Replace(LoopClassStyle,"{@autoid}",ClassRow)
						 LoopClassStyle = Replace(LoopClassStyle,"{@tid}",ID)
						 LoopClassStyle = Replace(LoopClassStyle,"{@classid}",KS.C_C(ID,9))
						 LoopClassStyle = Replace(LoopClassStyle,"{@classname}",KS.C_C(ID,1))
						 LoopClassStyle = Replace(LoopClassStyle,"{@classurl}",KS.GetFolderPath(ID))
						 LoopClassStyle = Replace(LoopClassStyle,"{@classimg}",ClassBasicInfoArr(0))
						 LoopClassStyle = Replace(LoopClassStyle,"{@classintro}",KS.Gottopic(ClassBasicInfoArr(1),200))
						 LoopClassStyle = Replace(LoopClassStyle,"{$InnerText}",GetGenericList(DocStyle))
						 ClassStr = ClassStr & ParseIF(LoopClassStyle)
				   Next
			   Else
			      Dim ClassDocNode,ClassTotalNum,K,ClassCol,MenuBgStr,MenuBgType,MenuBg,O_T_S,M,DocList
				  ClassCol   = ParamNode.getAttribute("classcol")  : If Not IsNumeric(ClassCol) Then ClassCol=2
				  MenuBgType = ParamNode.getAttribute("menubgtype") : If Not IsNumeric(MenuBgType) Then MenuBgType=0
				  MenuBg     = ParamNode.getAttribute("menubg")  
				  MenuBgStr = KS.GetMenuBg(MenuBgType, MenuBg, ClassCol)
                  O_T_S     = KS.G_O_T_S(ParamNode.getAttribute("opentype")) : M=0
			      Set ClassDocNode=classXml.DocumentElement.SelectNodes("row")
		          ClassTotalNum=ClassDocNode.length
			      ClassStr = "<table border=""0"" cellpadding=""0"" cellspacing=""2"" width=""100%"">" & vbCrLf
				  For K=0 To ClassTotalNum-1
				    ClassStr = ClassStr & "<tr>" & vbcrlf
					For I = 1 To ClassCol
					      Set ClassNode= ClassDocNode.Item(M)
	                      ID = ClassNode.SelectSingleNode("@id").text
						  regEx.Pattern="\bclassid="".*?"""
						  Set Matches = regEx.Execute(LabelParamStr)
						  If Matches.count > 0 Then LabelParamStr=Replace(LabelParamStr,Matches.item(0),"classid=""" & ID & """")
							ClassStr = ClassStr & "<td valign=""top"" style=""width:" & CInt(100 / CInt(ClassCol)) & "%;"">" & vbCrLf
							ClassStr = ClassStr & "<table width=""100%"" border=""0"" align=""center"" cellPadding=""0"" cellSpacing=""0""><tr><td>"
							ClassStr = ClassStr & "<div style=""text-align:left;height: 30px;line-height:30px;border-top: 1px solid #d2d3d9;border-left: 1px solid #d2d3d9;border-right: 1px solid #d2d3d9;padding-left:10px;background:" & MenuBgStr & """><div style=""float:left;font-weight:bold;"">" & KS.GetClassNP(ID) & "</div><div style=""float:right;""><a href=""" & KS.GetFolderPath(ID) & """>更多...</a></div></div>" & vbCrLf
							ClassStr = ClassStr & "<div style=""border: 1px solid #D2D3D9;line-height: 150%;text-align: left;padding:0px 5px 0px 5px;"">" & vbCrLf
							DocList  = GetGenericList(DocStyle)								   
							If DocList="" Then
							 ClassStr = ClassStr & "此栏目下没有添加信息"
							Else
							 ClassStr = ClassStr & DocList
							End If
							ClassStr = ClassStr & "</div>" & vbCrLf
							ClassStr = ClassStr & "</td></tr></table></td>" & vbCrLf
							
							M=M+1
							If M>=ClassTotalNum Then Exit For
					Next
					ClassStr = ClassStr  & "</tr>" & vbcrlf
					If M>=ClassTotalNum Then Exit For
				  Next
				  ClassStr = ClassStr  & "</table>" & vbcrlf
			   End If
			 End If
		   GetCirClassList=ClassStr
		End Function
		
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:GetRelativeList
		'作 用: 链接列表标签函数
		'参 数: LabelStyle 标签样式
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function GetRelativeList(LabelStyle)
		' If FCls.RefreshType="Content" Then
		  GetRelativeList=GetGenericList(LabelStyle)
		' Else
		'  GetRelativeList="相关链接标签只能放在内容页模板!"
		' End If
        End Function
		
		
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:GetPageList
		'作 用: 终级分页列表标签函数
		'参 数: LabelStyle 标签样式
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function GetPageList(LabelStyle)
		 If FCls.RefreshType = "Folder" Or FCls.RefreshType="Special" Then
				   LoadLabelParam
			      If LoadSucceed = false Then GetPageList="标签加载出错!":Exit Function 
				  If cbool(AjaxOut) = true  Then GetPageList=GetPageStr(LabelID):Exit Function
				  If (FCls.RefreshType = "Folder" And (KS.C_S(FCls.ChannelID,7)="0" or KS.C_S(FCls.ChannelID,7)="2" Or KS.C_C(FCls.RefreshFolderID,3)="2")) Or (FCls.RefreshType = "Special" And KS.Setting(78)="0") Then 
				   Application("PageParam")=LabelParamStr:Application("LabelStyle")=LabelStyle : GetPageList="{Tag:Page}": Exit Function
				  End If

				 
				  Dim FolderID,SqlStr,TableName,FieldStr,PrintType,PicStyle,ShowPicFlag,IncludeSubClass,PageStyle,RS
				  Dim TotalPut,PerPageNum
				  ShowPicFlag     = ParamNode.getAttribute("showpicflag") 
				  PrintType       = ParamNode.getAttribute("printtype")       : If Not IsNumeric(PrintType) Then PrintType=1
				  PicStyle        = ParamNode.getAttribute("picstyle")        : If Not IsNumeric(PicStyle) Then PicStyle=1
				  IncludeSubClass = ParamNode.getAttribute("includesubclass") 
				  PerPageNum = ParamNode.getAttribute("num") : If Not IsNumeric(PerPageNum) Then PerPageNum=10
				  PageStyle  = ParamNode.getAttribute("pagestyle") : If KS.IsNul(PageStyle) Then PageStyle=1
                  FolderID = FCls.RefreshFolderID
                  Param = " Where I.Verific=1 And I.DelTF=0"
				  If FCls.RefreshType="Special" Then
				      Param= Param & KS.GetSpecialPara(ModelID,FCls.CurrSpecialID)
				  Else
				      ModelID  = FCls.ChannelID
					  If CBool(IncludeSubClass) = True Then 
						 Param= Param & " And I.Tid In (" & KS.GetFolderTid(FolderID) & ")" 
					  Else 
						 Param= Param & " And I.Tid='" & FolderID & "'"
					  End If
				  End If
				  
				  Dim SQLParam:SQLParam=ParamNode.getAttribute("sqlparam")
					If Not KS.IsNul(SQLParam) Then
					 'Param="Where ("&split(lcase(Param),"where ")(1) & ") " & SQLParam
					 Param=Param & " " & SQLParam
				  End If

				  
			      LoadField ModelID,PrintType,PicStyle,ShowPicFlag,FieldStr,TableName,Param
						
			      SqlStr = "SELECT " & FieldStr & " FROM " & TableName & " I " & Param & " ORDER BY I.IsTop Desc," & GetOrderParam()
                  Set RS=Server.CreateObject("ADODB.RECORDSET")
				  RS.Open SQLStr,Conn,1,1
				  
				  If RS.EOF And RS.Bof Then	GetPageList="<p>此栏目下没有信息!</p>":RS.Close:Set RS=Nothing:FCls.PageList = "":Exit Function
				  TotalPut = Conn.Execute("select Count(id) from " & TableName & " I " & Param)(0)
				  PerPageNum=cint(PerPageNum)
				   Dim N,PageNum, CurrPage,TempStr
				    if (TotalPut mod PerPageNum)=0 then
							PageNum = TotalPut \ PerPageNum
				    else
							PageNum = TotalPut \ PerPageNum + 1
				    end if
					Dim EndPageNum:EndPageNum=PageNum
					If KS.ChkClng(FCls.FsoListNum)<>0 And KS.ChkClng(FCls.FsoListNum)<PageNum Then EndPageNum=KS.ChkClng(FCls.FsoListNum)				
					  If FCls.RefreshType="Folder" And EndPageNum>5 Then KS.Echo "<script>show();</script>"
					  For CurrPage = 1 To EndPageNum
						 RS.Move (CurrPage - 1) * PerPageNum,1
						 Set XMLSQL=KS.ArrayToXml(RS.GetRows(PerPageNum),rs,"row","root")
						 TempStr = TempStr &  ExplainGerericListLabelBody(LabelStyle)
					     TempStr = TempStr & "{KS:PageList}" '加上分页符
						 If FCls.RefreshType="Folder" And EndPageNum>5 And CurrPage Mod 2=0 Then
							KS.Echo "<script>$('#fsotips').html('正在生成栏目""<font color=red>" & KS.C_C(FolderID,1) & "</font>"",本栏目共有<font color=red>" & EndPageNum & "</font>个分页需要生成,正在获取第<font color=red>" & CurrPage & "</font>个分页数据...');</script>"
							Response.Flush()
						 End If
						 
					   If RS.Eof Then Exit For
					 Next
					 If FCls.RefreshType="Folder" And EndPageNum>5 Then KS.Echo "<script>$('#fsotips').html('获取分页数据完毕,分页生成中...');</script>"
					 RS.Close:Set RS = Nothing
					 FCls.PageList=TempStr
					 FCls.PageStyle=PageStyle
					 FCls.PerPageNum=PerPageNum
					 FCls.TotalPage=PageNum
					 FCls.TotalPut=TotalPut
					 GetPageList="{PageListStr}"
		 Else
		   GetPageList="分页标签只能放在栏目列表页及专题页模板!"
		 End If
        End Function
		
		'取得Ajax分页函数
		Function GetPageStr(LabelID)
			Templates = ""
			echoln "<script src=""" & DomainStr & "ks_inc/page.js"" type=""text/javascript""></script>"
			echoln "<script type=""text/javascript"" defer>"
			echoln "   Page(1,'"& LabelID & "','" & FCls.RefreshFolderID & "','" & KS.Setting(3) &"','item/ajaxpage.asp','" & FCls.RefreshType & "','" & FCls.CurrSpecialID & "','" & FCls.RefreshInfoID &"');"
			echoln "</script>"
			echoln "  <div id=""pagecontent""><div id=""c_" & LabelID & """></div></div>"
			echoln "  <div id=""fenye""  class=""fenye""><div id=""p_" & LabelID & """ align=""right""></div></div>"
            GetPageStr = Templates
		End Function
		
	
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:GetSpecialList
		'作 用: 专题列表标签函数
		'参 数: LabelStyle 标签样式
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function GetSpecialList(ByVal LabelStyle)
             If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				 GetSpecialList="标签加载出错" : Exit Function
			 Else
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
		     LabelID   = ParamNode.getAttribute("labelid")
			 ClassID   = ParamNode.getAttribute("classid") : If Not IsNumeric(ClassID) Then ClassID=0
			 AjaxOut   = ParamNode.getAttribute("ajaxout")
			 Num       = ParamNode.getAttribute("num")
			 
			Dim OrderStr
		    OrderStr  = ParamNode.getAttribute("orderstr")  : If OrderStr="" Then OrderStr=" SpecialID Desc"
		   OrderStr=Lcase(OrderStr)
		  If trim(OrderStr)="rnd" Then
			 
			  Randomize : OrderStr="Rnd(-(SpecialID+"&Rnd()&"))"
		  ElseIf Lcase(Left(Trim(OrderStr),9))<>"specialid" Then  
			 OrderStr=OrderStr & ",SpecialID Desc"
		  Else 
		     OrderStr=" SpecialID Desc"
		  End If
			If ClassID<>0 Then Param=" Where S.ClassID=" & ClassID Else Param=""
			Dim SqlStr:SqlStr="Select TOP " & num & " S.specialid,S.classid,S.SpecialName,S.SpecialEname,S.FsoSpecialIndex,S.SpecialAddDate as AddDate,S.PhotoUrl,S.SpecialNote As Intro,S.creater,C.ClassName as SpecialClassName From KS_Special S Inner Join KS_SpecialClass C On S.ClassID=C.ClassID" & Param & " Order By " & OrderStr
			Dim RS:Set RS=Conn.Execute(SqlStr)
			If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			RS.Close:Set RS=Nothing
			If IsObject(XMLSql) Then
			 GetSpecialList=ExplainSpecialListLabelBody(LabelStyle)
			End If 
			Set Node=Nothing
		End Function
		
		Function ExplainSpecialListLabelBody(ByVal LabelStyle)
		  Dim PrintType,ShowStyle,NaviStr,T_CssStr,PhotoCssStr,TotalNum,K,I,Col,TempTitle,SpecialUrl,O_T_S,T_Len,I_Len,DateRule,DateAlign,ColSpanNum,SplitPic,R_H,TempPicStr,PicWidth,PicHeight,PicWidthStr,PicHeightStr,MoreLink,MoreLinkStr,MoreType
		  PrintType  = ParamNode.getAttribute("printtype") : If Not IsNumeric(PrintType) Then PrintType=1
		  ShowStyle  = ParamNode.getAttribute("showstyle") : If Not IsNumeric(ShowStyle) Then ShowStyle=1
		  Col        = ParamNode.getAttribute("col")       : If Not IsNumeric(Col) Then Col=1
		  NaviStr    = KS.GetNavi(ParamNode.getAttribute("navtype"), ParamNode.getAttribute("nav"))
		  T_CssStr   = KS.GetCss(ParamNode.getAttribute("titlecss"))
		  PhotoCssStr = KS.GetCss(ParamNode.getAttribute("photocss"))
		  O_T_S      = KS.G_O_T_S(ParamNode.getAttribute("opentype"))
		  T_Len      = ParamNode.getAttribute("titlelen")  : If Not IsNumeric(T_Len) Then T_Len=0
		  I_Len      = ParamNode.getAttribute("introlen")  : If Not IsNumeric(T_Len) Then T_Len=0
		  DateRule   = ParamNode.getAttribute("daterule")
		  DateAlign  = ParamNode.getAttribute("datealign")
		  SplitPic   = ParamNode.getAttribute("splitpic")
		  R_H        = ParamNode.getAttribute("rowheight")
		  PicWidth   = KS.ChkClng(ParamNode.getAttribute("picwidth")):If PicWidth<>0 Then PicWidthStr=" width=""" & PicWidth & """"
		  PicHeight  = KS.ChkClng(ParamNode.getAttribute("picheight")):If PicHeight<>0 Then PicHeightStr=" height="""&PicHeight&""""
		  MoreLink   = ParamNode.getAttribute("morelink")
		  MoreType   = ParamNode.getAttribute("morelinktype")
		  
		  If ClassID<>0 And MoreLink <> "" Then MoreLinkStr= KS.GetMoreLink(1,Col, 20, MoreType, MoreLink, KS.GetFolderSpecialPath(ClassID, True), O_T_S)
		  
		  Set DocNode=XMLSql.DocumentElement.SelectNodes("row")
		  TotalNum=DocNode.length
		  Templates  = "" : N = 0 
		  If PrintType=1 Then
		     Templates="" : N=0
			 echoln "<table border=""0"" cellpadding=""0"" align=""center"" cellspacing=""0"" width=""99%"">"
			 For K=0 To TotalNum-1
				 echo "<tr>" & vbCrLf
				 For I = 1 To Col
					 Set Node=DocNode.Item(n)
		              TempTitle = Node.SelectSingleNode("@specialname").text 
					  SpecialUrl=KS.GetSpecialPath(Node.SelectSingleNode("@specialid").text,Node.SelectSingleNode("@specialename").text,Node.SelectSingleNode("@fsospecialindex").text)
					  TempPicStr=GetPicUrl(Node.SelectSingleNode("@photourl").text)
					  TempPicStr="<a href=""" & SpecialUrl & """" & O_T_S & """ title=""" & TempTitle & """><img src=""" & TempPicStr & """" & PicWidthStr & PicHeightStr &""" alt=""" & TempTitle & """ border=""0""" & PhotoCssStr &"/></a>"
					  
					  TempTitle = "<a" & T_CssStr & " href=""" & SpecialUrl & """" & O_T_S & " title=""" & TempTitle & """>" & KS.GotTopic(TempTitle,T_Len) & "</a>"
					  
					  
					If Col=1 Then
						 echoln ("  <td height=""" & R_H & """>")
					Else
						 echoln ("<td width=""" & CInt(100 / CInt(Col)) & "%"" height=""" & R_H & """>")
					End If
					ColSpanNum=Col
					select case ShowStyle
					  case 1 echo NaviStr & TempTitle  & GetDateStr(1,Node.SelectSingleNode("@adddate").text,DateRule,DateAlign,"",Col, ColSpanNum)
					  
					  case 2 echo TempPicStr
					  case 3 echo "<div style=""text-align:center"">" &TempPicStr&"<br />"&TempTitle & "</div>"
					  case 4 
					     echoln "<table cellSpacing=""0"" cellpadding=""0"" style=""margin:3px;width:100%"" border=""0"">"
						 echoln " <tr>"
						 echoln "  <td align=""center"" width=" & PicWidth+10 & ">"
						 echoln "  " & TempPicStr
						 echoln "  </td>"
						 echoln " <td>" & KS.GotTopic(KS.LoseHtml(Node.SelectSingleNode("@intro").text),I_len)&"</td>"
						 echoln " </tr>"
						 echoln "</table>"
					 Case 5
					     echoln "<table cellspacing=""0"" cellpadding=""0"" style=""margin:3px;width:100%"" border=""0"">"
						 echoln " <tr>"
						 echoln " <td align=""center"" width=""" & PicWidth+10 & """>"
						 echoln "  " & TempPicStr
						 echoln " </td>"
						 echoln"  <td>" & TempTitle &"<br />" & KS.GotTopic(KS.LoseHtml(Node.SelectSingleNode("@intro").text),I_len)&"</td>"
						 echoln " </tr>"
						 echoln "</table>"
					end select
					echoln "</td>"
					N = N+1 : If N>=TotalNum Then Exit For
				 Next
				 echoln "</tr>"
				 echoln KS.GetSplitPic(SplitPic, ColSpanNum)
				 If N>=TotalNum Then Exit For
		    Next
			echoln MoreLinkStr
		    echoln "</table>"
		 Else 
		    Templates = ExplainDiyStyle(LabelStyle,TotalNum)
		 End If
		 ExplainSpecialListLabelBody=Templates
		End Function
		
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:GetCirSpecialList
		'作 用: 循环分类专题列表标签函数
		'参 数: LabelStyle 标签样式
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function GetCirSpecialList(LabelStyle)
		     If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				 GetCirSpecialList = "标签加载出错!" : Exit Function
			 Else
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
 			 Dim ClassParam,SQLStr,ClassXml,ClassNode,I,ClassStyle,DocStyle,ClassStr,LoopClassStyle,ID,ClassPrintType

		     LabelID   = ParamNode.getAttribute("labelid")
			 ClassPrintType=ParamNode.getAttribute("classprinttype") : If Not IsNumeric(ClassPrintType) Then ClassPrintType=1
			 
			 If LabelID<>"ajax" and Cbool(ParamNode.getAttribute("ajaxout"))=true Then 
			  GetCirSpecialList="<span id=""ks" & LabelID & "_0_0_0_0""></span>":Exit Function
			 End If
		     
			 SQLStr="Select ClassID,ClassName,Descript From KS_SpecialClass ORDER BY OrderID asc"
			 Dim classObj:Set classObj=Conn.Execute(SQLStr)
             If Not classObj.Eof Then
			  Set ClassXml=KS.RsToXml(classObj,"row","root")
			 End If
			 classObj.Close:Set classObj=Nothing
		     If IsObject(classXml) Then
			   	 regEx.Pattern="\bajaxout="".*?"""
				 Set Matches = regEx.Execute(LabelParamStr)
				 If Matches.count > 0 Then LabelParamStr=Replace(LabelParamStr,Matches.item(0),"ajaxout=""false""")
			     ClassStyle = Split(LabelStyle,"§")(0)
			     DocStyle = Split(LabelStyle,"§")(1)
				
				 If ClassPrintType=2 Then
				   Dim ClassRow:ClassRow=0
				   For Each ClassNode In classXml.DocumentElement.SelectNodes("row")
					 ID = ClassNode.SelectSingleNode("@classid").text
					 regEx.Pattern="\bclassid="".*?"""
					 Set Matches = regEx.Execute(LabelParamStr)
					 If Matches.count > 0 Then LabelParamStr=Replace(LabelParamStr,Matches.item(0),"classid=""" & ID & """")
						 LoopClassStyle = ClassStyle
						 ClassRow=ClassRow+1
                         LoopClassStyle = Replace(LoopClassStyle,"{@autoid}",ClassRow)
						 LoopClassStyle = Replace(LoopClassStyle,"{@classid}",ID)
						 LoopClassStyle = Replace(LoopClassStyle,"{@specialclassname}",ClassNode.SelectSingleNode("@classname").text)
						 LoopClassStyle = Replace(LoopClassStyle,"{@specialclassurl}",KS.GetFolderSpecialPath(ID, True))
						 LoopClassStyle = Replace(LoopClassStyle,"{@specialclassintro}",KS.Gottopic(ClassNode.SelectSingleNode("@descript").text,200))
						 LoopClassStyle = Replace(LoopClassStyle,"{$InnerText}",GetSpecialList(DocStyle))
						 ClassStr = ClassStr & ParseIF(LoopClassStyle)
				   Next
				 Else
				  Dim ClassDocNode,ClassTotalNum,K,ClassCol,MenuBgStr,MenuBgType,MenuBg,O_T_S,M,DocList
				  ClassCol   = ParamNode.getAttribute("classcol")  : If Not IsNumeric(ClassCol) Then ClassCol=2
				  MenuBgType = ParamNode.getAttribute("menubgtype") : If Not IsNumeric(MenuBgType) Then MenuBgType=0
				  MenuBg     = ParamNode.getAttribute("menubg")  
				  MenuBgStr = KS.GetMenuBg(MenuBgType, MenuBg, ClassCol)
                  O_T_S     = KS.G_O_T_S(ParamNode.getAttribute("opentype")) : M=0
			      Set ClassDocNode=classXml.DocumentElement.SelectNodes("row")
		          ClassTotalNum=ClassDocNode.length
			      ClassStr = "<table border=""0"" cellpadding=""0"" cellspacing=""2"" width=""100%"">" & vbCrLf
				  For K=0 To ClassTotalNum-1
				    ClassStr = ClassStr & "<tr>" & vbcrlf
					For I = 1 To ClassCol
					      Set ClassNode= ClassDocNode.Item(M)
	                      ID = ClassNode.SelectSingleNode("@classid").text
						  regEx.Pattern="\bclassid="".*?"""
						  Set Matches = regEx.Execute(LabelParamStr)
						  If Matches.count > 0 Then LabelParamStr=Replace(LabelParamStr,Matches.item(0),"classid=""" & ID & """")
							ClassStr = ClassStr & "<td valign=""top"" style=""width:" & CInt(100 / CInt(ClassCol)) & "%;"">" & vbCrLf
							ClassStr = ClassStr & "<table width=""100%"" border=""0"" align=""center"" cellPadding=""0"" cellSpacing=""0""><tr><td>"
							ClassStr = ClassStr & "<div style=""text-align:left;height: 30px;line-height:30px;border-top: 1px solid #d2d3d9;border-left: 1px solid #d2d3d9;border-right: 1px solid #d2d3d9;padding-left:10px;background:" & MenuBgStr & """><div style=""float:left;font-weight:bold;""><a href=""" & KS.GetFolderSpecialPath(ClassNode.SelectSingleNode("@classid").text, True) & """ target=""_blank"">" & ClassNode.SelectSingleNode("@classname").text & "</a></div><div style=""float:right;""><a href=""" & KS.GetFolderSpecialPath(ID, True) & """>更多...</a></div></div>" & vbCrLf
							ClassStr = ClassStr & "<div style=""border: 1px solid #D2D3D9;line-height: 150%;"">" & vbCrLf
							DocList  = GetSpecialList(DocStyle)								   
							If DocList="" Then
							 ClassStr = ClassStr & "此分类下没有添加专题"
							Else
							 ClassStr = ClassStr & DocList
							End If
							ClassStr = ClassStr & "</div>" & vbCrLf
							ClassStr = ClassStr & "</td></tr></table></td>" & vbCrLf
							
							M=M+1
							If M>=ClassTotalNum Then Exit For
					Next
					ClassStr = ClassStr  & "</tr>" & vbcrlf
					If M>=ClassTotalNum Then Exit For
				  Next
				  ClassStr = ClassStr  & "</table>" & vbcrlf
				 
				 End If
			 End If
			 GetCirSpecialList = ClassStr
		End Function
		
		'取得分页分类下的专题
		Function GetLastSpecialList(LabelStyle)
			 LoadLabelParam
			 If LoadSucceed = false Then GetLastSpecialList="标签加载出错!":Exit Function 
			 If cbool(AjaxOut) = true  Then GetLastSpecialList=GetPageStr(LabelID):Exit Function
			 If FCls.FromAspPage=True Then 
			  	   Application("PageParam")=LabelParamStr
				   Application("LabelStyle")=LabelStyle
				   GetLastSpecialList="{Tag:Page}"
                   FCls.FromAspPage=false:Exit Function
			 End If
			 
			 If FCls.RefreshType = "ChannelSpecial" Then 
			        Dim SqlStr,RS,TotalPut,PerPageNum,PrintType,PageStyle
			   		PrintType       = ParamNode.getAttribute("printtype") : If Not IsNumeric(PrintType) Then PrintType=1
					PageStyle       = ParamNode.getAttribute("pagestyle") : If KS.IsNul(PageStyle) Then PageStyle=1
					PerPageNum      = ParamNode.getAttribute("num") : If Not IsNumeric(PerPageNum) Then PerPageNum=10

				  SqlStr="Select S.specialid,S.classid,S.SpecialName,S.SpecialEname,S.FsoSpecialIndex,S.SpecialAddDate as AddDate,S.PhotoUrl,S.SpecialNote As Intro,S.creater,C.ClassName as SpecialClassName From KS_Special S Inner Join KS_SpecialClass C On S.ClassID=C.ClassID Where S.ClassID=" & KS.ChkClng(FCls.RefreshFolderID) & " Order By S.SpecialID Desc"
					
                  Set RS=Server.CreateObject("ADODB.RECORDSET")
				  RS.Open SQLStr,Conn,1,1
				  
				  If RS.EOF And RS.Bof Then	GetLastSpecialList="<p>此分类下没有专题!</p>":RS.Close:Set RS=Nothing:FCls.PageList = "":Exit Function
				  TotalPut = Conn.Execute("Select count(S.specialid) From KS_Special S Inner Join KS_SpecialClass C On S.ClassID=C.ClassID Where S.ClassID=" & KS.ChkClng(FCls.RefreshFolderID))(0)
				  PerPageNum=cint(PerPageNum)
				   Dim N,PageNum, CurrPage,TempStr
				    if (TotalPut mod PerPageNum)=0 then
							PageNum = TotalPut \ PerPageNum
				    else
							PageNum = TotalPut \ PerPageNum + 1
				    end if
					Dim EndPageNum:EndPageNum=PageNum
					  For CurrPage = 1 To EndPageNum
						 RS.Move (CurrPage - 1) * PerPageNum,1
						 Set XMLSQL=KS.ArrayToXml(RS.GetRows(PerPageNum),rs,"row","root")
						 TempStr = TempStr & ExplainSpecialListLabelBody(LabelStyle)
					     TempStr = TempStr & "{KS:PageList}" '加上分页符
					   If RS.Eof Then Exit For
					 Next
					 RS.Close:Set RS = Nothing
					 FCls.PageList=TempStr
					 FCls.PageStyle=PageStyle
					 FCls.PerPageNum=PerPageNum
					 FCls.TotalPage=PageNum
					 FCls.TotalPut=TotalPut
					 GetLastSpecialList="{PageListStr}"
		   Else
		    GetLastSpecialList="此标签只能放在专题分页页模板!"
		   End If
		End Function
		
		'位置导航
		Function GetLocation(LabelStyle)
		     If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				GetLocation="" : Exit Function
			 Else 
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
		     Dim KSLCls:Set KSLCls=New RefreshLocationCls
			 GetLocation = KSLCls.GetLocation(ParamNode)
			 Set KSLCls=Nothing
		End Function
		
		'取得顶部栏目导航
		Function GetNavigation(LabelStyle)
			 If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				GetNavigation="" : Exit Function
			 Else 
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
             Dim ChannelID,SQLStr,NavType,Nav,SplitPic,Col, OpenType,O_P_T,T_Css,P_T,DivID,DivCss,UlID,ULCss,LiID,LICss,FieldStr
			 ChannelID   = ParamNode.getAttribute("channelid")
			 NavType     = ParamNode.getAttribute("navtype")
			 Nav         = ParamNode.getAttribute("nav")
			 SplitPic    = ParamNode.getAttribute("splitpic")
			 Col         = ParamNode.getAttribute("col") : If Not IsNumeric(Col) Then Col=1
			 OpenType    = ParamNode.getAttribute("opentype") : O_P_T=KS.G_O_T_S(OpenType)
			 T_Css       = ParamNode.getAttribute("titlecss")
			 P_T         = ParamNode.getAttribute("printtype")
			 DivID       = ParamNode.getAttribute("divid")
			 DivCss      = ParamNode.getAttribute("divclass")
			 ULID        = ParamNode.getAttribute("ulid")
			 ULCss       = ParamNode.getAttribute("ulclass")
			 LIID        = ParamNode.getAttribute("liid")
			 licss       = ParamNode.getAttribute("liclass")
			 If Cint(P_T)=3 Then
			  FieldStr="ClassID,ID,ID AS Tid, FolderName,TS,ClassBasicInfo,folder"
			 Else
			  FieldStr="ClassID,ID,FolderName,TS,folder"
			 End If
			 select case channelid
			   case "0" SqlStr = "Select " & FieldStr & ",TN,FolderOrder From KS_Class A Inner Join KS_Channel B On A.ChannelID=B.ChannelID Where  B.ChannelStatus=1 and TN='0' AND TopFlag=1 And DelTF=0 Order By root,FolderOrder"
			   case "9999" 
			       if FCls.RefreshFolderID="0" then
					   SqlStr = "Select " & FieldStr & ",TN,FolderOrder From KS_Class a inner join KS_Channel b on a.channelid=b.channelid Where  B.ChannelStatus=1 and TN='0' AND TopFlag=1 And DelTF=0 Order By root,FolderOrder"
				   else
					    SqlStr = "Select " & FieldStr & " From KS_Class A Inner Join KS_Channel B On A.ChannelID=B.ChannelID Where B.ChannelStatus=1 And TN='" & FCls.RefreshFolderID & "' And TopFlag=1 And DelTF=0  Order BY root,FolderOrder"
					end if
			   case "9998"
					  Dim Rst ,ParentID
					  Set Rst=Conn.Execute("Select top 1 TN From KS_Class Where ID='" & FCls.RefreshFolderID & "'")
					  If Not Rst.EOF Then ParentID=Rst(0)  Else  ParentID=FCls.RefreshFolderID
					  if parentid="0" then parentid=FCls.RefreshFolderID
					  Rst.close:Set Rst=Nothing
					  SqlStr = "Select " & FieldStr & " From KS_Class A Inner Join KS_Channel B On A.ChannelID=B.ChannelID Where B.ChannelStatus=1 And TN='" & ParentID & "' And TopFlag=1 and DelTF=0  Order BY root,FolderOrder"

				case "9997" GetNavigation=GetExtNav(1,NavType, Nav, SplitPic, Col, OpenType, T_Css,P_T,DivID,DivCss,UlID,ULCss,LiID,LICss):Exit Function 
				case "9996" GetNavigation=GetExtNav(2,NavType, Nav, SplitPic, Col, OpenType, T_Css,P_T,DivID,DivCss,UlID,ULCss,LiID,LICss):Exit Function 
				case "9995" GetNavigation=GetExtNav(3,NavType, Nav, SplitPic, Col, OpenType, T_Css,P_T,DivID,DivCss,UlID,ULCss,LiID,LICss):Exit Function 
				case "9994" GetNavigation=GetExtNav(4,NavType, Nav, SplitPic, Col, OpenType, T_Css,P_T,DivID,DivCss,UlID,ULCss,LiID,LICss):Exit Function 
			    case else
			        If Len(ChannelID)<=3 Then
						SqlStr = "Select " & FieldStr & " From KS_Class where TN='0' And ChannelID=" & ChannelID & " AND TopFlag=1 And DelTF=0  Order BY root,FolderOrder"
					Else
					    SqlStr = "Select " & FieldStr & " From KS_Class A Inner Join KS_Channel B On A.ChannelID=B.ChannelID Where B.ChannelStatus=1 And TopFlag=1 And TN='" & ChannelID & "' And DelTF=0 Order BY root,FolderOrder"

					End If
			 end select
			 Dim RS,XML,I,K,Node,TotalNum,EndDIV,EndUL,NavStr,ClassID
			 Set RS=Conn.Execute(SQLStr)
			 If Not RS.Eof Then Set XML=KS.RsToXml(RS,"row","root")
			 Templates = "" : N = 0
			 If IsObject(XML) Then
			         If Cint(P_T)=3 Then
					  Set DocNode=XML.DocumentElement.SelectNodes("row")
		              Templates=ExplainDiyStyle(LabelStyle,DocNode.length)
                     ElseIf Cint(P_T)=2 Then
					      If DivID<>"" Or DivCss<>"" Then echoln "<div" & KS.GetCssID(DivID)&KS.GetCss(DivCss) &">" : EndDIv="</div>" Else EndDiv=""
						  If UlID<>"" Or ULCss<>"" Then echoln "<ul"&KS.GetCssID(UlID)&KS.GetCss(ULCss) &">" : EndUL="</ul>" Else EndUL=""
						  Dim KK:KK=0
						  For Each Node In XML.DocumentElement.SelectNodes("row")
						     ClassID=Node.SelectSingleNode("@id").text
							 If FCls.RefreshFolderID=ClassID or (UCase(FCls.RefreshType) = "INDEX" and kk=0) Then
							 echo  "  <li class=""currclass"""&KS.GetCssID(LIID)&KS.GetCss(LICss)&">"
							 '增加第一级导航判断
							 ElseIf  Split(KS.C_C(Fcls.RefreshFolderID,8)&",",",")(0)=classid then
							    echo  "  <li class=""currclass"""&KS.GetCssID(LIID)&KS.GetCss(LICss)&">"
							  '===========
							 Else
							   echo  "  <li"&KS.GetCssID(LIID)&KS.GetCss(LICss)&">"
							 End If
							 echo "<a " & KS.GetCss(T_Css) & " href=""" & KS.GetFolderPath(ClassID) & """" & O_P_T & ">" & Trim(Node.SelectSingleNode("@foldername").text) & "</a></li>"
							 kk=kk+1
						  Next
						   If EndUL<>"" Then echoln EndUL
						   If EndDiv<>"" Then echoln EndDiv
					  Else
					      	Set DocNode=XML.DocumentElement.SelectNodes("row")
		                    TotalNum=DocNode.length
							echoln "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
						   For I=0 To TotalNum-1
							   echoln "<tr><td height=""22"" align=""center"">"
							  For K = 1 To Col
							    Set Node=DocNode.Item(n)
								ClassID= Node.SelectSingleNode("@id").text
								 If K=1 Then NavStr="" Else  NavStr=KS.GetNavi(NavType, Nav)
								 If FCls.RefreshFolderID=ClassID  Then
								  echo NavStr & "<a class=""currclass""" & KS.GetCss(T_Css) & " href=""" & KS.GetFolderPath(ClassID) & """" & O_P_T & ">" & Trim(Node.SelectSingleNode("@foldername").text) & "</a>"
								 Else
								  echo NavStr & "<a" & KS.GetCss(T_Css) & " href=""" & KS.GetFolderPath(ClassID) & """" & O_P_T & ">" & Trim(Node.SelectSingleNode("@foldername").text) & "</a>"
								 End If
								 N = N+1 : If N>=TotalNum Then Exit For
							  Next
							  echoln "</td></tr>"
							  echoln KS.GetSplitPic(SplitPic, Col)
							  If N>=TotalNum Then Exit For
						   Next
						   echoln "</table>"
					   End If
			 End If
		     GetNavigation=Templates
		End Function
		
		'取得外部频道导航
		Function GetExtNav(Flag,NavType, Nav, SplitPic, Col, OpenType, T_Css,P_T,DivID,DivCss,UlID,ULCss,LiID,LICss)
					Dim SQL,RS, I,K,TotalNum,N,Url,SQLStr
					Select Case Flag
					 Case 1:SQLStr="Select ClassID,ClassName From KS_BlogClass Order By OrderID asc"
					 Case 2:SQLStr="Select TypeID,TypeName From KS_BlogType Order By OrderID asc"
					 Case 3:SQLStr="Select ClassID,ClassName From KS_TeamClass Order By OrderID asc"
					 Case 4:SQLStr="Select ClassID,ClassName From KS_PhotoClass Order By OrderID asc"
					 Case Else:Exit Function
					End Select
					Set RS = Conn.Execute(SqlStr)
					If RS.Eof And RS.Bof Then GetExtNav="":RS.Close:Set RS=Nothing:Exit Function
					SQL=RS.GetRows(-1):TotalNum=Ubound(SQL,2)
					
					  If Cint(P_T)=2 Then
						GetExtNav = "<div"&KS.GetCssID(DivID)&KS.GetCss(DivCss) &">" & vbCrLf & " <ul"&KS.GetCssID(UlID)&KS.GetCss(ULCss) &">" & vbCrLf
					    For K=0 To TotalNum
						    Url=SQL(0,K)
						    Select Case Flag
							 Case 1:Url=DomainStr &"space/morespace.asp?classID=" & Url
							 Case 2:Url=DomainStr &"space/morelog.asp?classID=" & Url
							 Case 3:Url=DomainStr &"space/moregroup.asp?classID=" & Url
							 Case 4:Url=DomainStr &"space/morephoto.asp?classID=" & Url
							 End Select
							GetExtNav = GetExtNav & "  <li"&KS.GetCssID(LIID)&KS.GetCss(LICss)&">" & "<a " & KS.GetCss(T_Css) & " href=""" & Url & """" & KS.G_O_T_S(OpenType) & ">" & Trim(SQL(1,K)) & "</a></li>" & vbCrLf
					    Next
					    GetExtNav = GetExtNav & "  </ul>" & vbcrlf & "  </div>" & vbCrLf
					  Else
						  GetExtNav = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & vbCrLf
					      For K=0 To TotalNum
						  GetExtNav = GetExtNav & "<tr><td height=""22"">" & vbCrLf
						  For I = 1 To Col
						    Url=SQL(0,N)
						    Select Case Flag
							 Case 1:Url=DomainStr &"space/morespace.asp?classID=" & Url
							 Case 2:Url=DomainStr &"space/morelog.asp?classID=" & Url
							 Case 3:Url=DomainStr &"space/moregroup.asp?classID=" & Url
							 Case 4:Url=DomainStr &"space/morephoto.asp?classID=" & Url
							 End Select
							  GetExtNav = GetExtNav & KS.GetNavi(NavType, Nav) & "<a " & KS.GetCss(T_Css) & " href=""" & Url & """" & KS.G_O_T_S(OpenType) & ">" & Trim(SQL(1,N)) & "</a>" & vbCrLf
							  N=N+1
							  If N>=TotalNum+1 Then Exit For
						  Next
						  GetExtNav = GetExtNav & "</td></tr>" & vbCrLf
						  GetExtNav = GetExtNav & KS.GetSplitPic(SplitPic, Col)
						  If N>=TotalNum+1 Then Exit For
					   Next
					   GetExtNav = GetExtNav & "</table>" & vbCrLf
					  End If
		End Function
		'取得网站公告列表
		Function GetAnnounceList(LabelStyle)
			 If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				GetAnnounceList="" : Exit Function
			 Else 
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
			 Dim AnnounceType, OWidth, OHeight, Width, Height, Speed, ShowStyle, OpenType, num, T_Len, ShowAuthor, C_Len,ChannelID,AjaxOut
			 Dim SqlStr, NaviStr, T_CssStr, Title, Content,AddDate,ID
			 Dim Param,Xml,Node
			 LabelID       = ParamNode.getAttribute("labelid")
			 AjaxOut       = ParamNode.getAttribute("ajaxout"):If Not KS.IsNul(AjaxOut) Then AjaxOut=Cbool(AJaxout)
			 AnnounceType  = KS.ChkClng(ParamNode.getAttribute("announcetype"))
			 If LabelID<>"ajax"and AnnounceType<>1 and AjaxOut=true Then GetAnnounceList="<span id=""ks" & LabelID & "_0_0_0_0""></span>":Exit Function
			 Channelid     = ParamNode.getAttribute("channelid")
			 Num           = ParamNode.getAttribute("listnumber") : If Not IsNumeric(Num) Then Num=10
			 ShowStyle     = KS.ChkClng(ParamNode.getAttribute("showstyle"))
			 OpenType      = KS.ChkClng(ParamNode.getAttribute("opentype"))
			 NaviStr       = KS.GetNavi(ParamNode.getAttribute("navtype"), ParamNode.getAttribute("nav"))
			 T_CssStr      = KS.GetCss(ParamNode.getAttribute("titlecss"))
			 T_Len         = KS.ChkClng(ParamNode.getAttribute("titlelen"))
			 C_Len         = KS.ChkClng(ParamNode.getAttribute("contentlen"))
			 ShowAuthor    = KS.ChkClng(ParamNode.getAttribute("showauthor"))
			 Speed         = KS.ChkClng(ParamNode.getAttribute("speed"))
			 OWidth        = KS.ChkClng(ParamNode.getAttribute("owidth"))
			 OHeight       = KS.ChkClng(ParamNode.getAttribute("oheight"))
			 Width         = KS.ChkClng(ParamNode.getAttribute("width"))
			 Height        = KS.ChkClng(ParamNode.getAttribute("height"))
			 Param = " Where 1=1"
             If ChannelID=9999 Then
			   Param=Param & " And ChannelID=" & KS.ChkClng(FCls.ChannelID)
			 ElseIf ChannelID=9998  Then
			   Param= Param & " And ChannelID=0"
			 ElseIf ChannelID<>0 Then 
			   Param= Param & " and ChannelID=" & ChannelID
			 End If
			 If num = 0 Then
			  SqlStr = "Select * From KS_Announce " & Param & " Order BY NewestTF Desc,AddDate Desc"
			 Else
			  SqlStr = "Select Top " & num & " * From KS_Announce " & Param & " Order BY NewestTF Desc,AddDate Desc"
			 End If
			 Templates="" : N=0
			 Dim RS:Set RS=Conn.Execute(SQLStr)
			 If Not RS.Eof Then Set Xml=KS.RsToXml(RS,"row","root")	  
			 RS.Close:Set RS=Nothing
			 If Not IsObject(Xml) Then Exit Function
			 select case AnnounceType
			        case 0  '普通
                      If ShowStyle = 1 Then          '纵向显示
					    echoln "<table cellspacing=""0"" cellpadding=""0"" border=""0"" width=""100%"">"
					    For Each Node in Xml.documentelement.SelectNodes("row")
						  Title = Trim(Node.SelectSingleNode("@title").text) : Content = Trim(node.SelectSingleNode("@content").text) : AddDate = Node.SelectSingleNode("@adddate").text : ID=Node.SelectSingleNode("@id").text
						  echoln "<tr><td>"
						  If OpenType = 0 Then
						   echo "<a" & T_CssStr & " href=""#"" onclick=""javascript:window.open('" & DomainStr & "plus/Announce/?" & ID & "','NewWin','height=" & OHeight & ", width=" & OWidth & ", toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no');"" title=""" & Title & """> "
						  Else
						   echo "<a" & T_CssStr & " href=""" & DomainStr & "plus/Announce/?" & ID & """ title=""" & Title & """ target=""_blank""> "
						  End If
						   echoln NaviStr & KS.GotTopic(Title, T_Len) & "</a></td></tr>"
						  If C_Len <> 0 Then
						   echoln "<tr><td style=""padding-left:10px"">" & KS.GotTopic(Replace(Replace(KS.LoseHtml(Content), vbCrLf, ""), "&nbsp;", ""), C_Len) & "…</td></tr>"
						  End If
						  If ShowAuthor = 1 Then
						   echoln "<tr><td align=""right"">" & Node.SelectSingleNode("@author").text & "</td></tr>"
						   echoln "<tr><td align=""right"">" & Year(AddDate) & "年" & Month(AddDate) & "月" & Day(AddDate) & "日</td></tr>"
						  End If
					   Next
					   echoln "</table>"
					ElseIf ShowStyle = 2 Then   '横向显示
					    For Each Node in Xml.documentelement.SelectNodes("row")
						  Title = Trim(Node.SelectSingleNode("@title").text) : AddDate = Node.SelectSingleNode("@adddate").text : ID=Node.SelectSingleNode("@id").text
						  If OpenType = 0 Then
						   echo "<a" & T_CssStr & " href=""#"" onclick=""javascript:window.open('" & DomainStr & "plus/Announce/?" & ID & "','NewWin','height=" & OHeight & ", width=" & OWidth & ", toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no');"" title=""" & Title & """> "
						  Else
						   echo "<a" & T_CssStr & " href=""" & DomainStr & "plus/Announce/?" & ID & """ title=""" & Title & """ target=""_blank""> "
						  End If
						   echo NaviStr & KS.GotTopic(Title, T_Len) & "</a>"
						 If ShowAuthor = 1 Then
						   echo "[" & Node.SelectSingleNode("@author").text & "&nbsp;&nbsp;" & Year(AddDate) & "年" & Month(AddDate) & "月" & Day(AddDate) & "日]"
						 End If
						 echo "&nbsp;&nbsp;"
						Next
					End If
				Case 1                   '弹出
				     ID=Xml.documentelement.SelectNodes("row").item(0).SelectSingleNode("@id").text
					 echoln "<script type=""text/javascript"">"
					 echoln "<!--"
					 echoln "window.open('" & DomainStr & "plus/Announce/?" & ID & "','NewWin','height=" & OHeight & ", width=" & OWidth & ",  toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no');"
					 echoln "-->"
					 echoln "</script>"
					 
				Case 2                   '滚动
					If ShowStyle = 1 Then       '纵向显示
					   echo "<marquee direction=""up"" onmouseover=""this.stop()"" onmouseout=""this.start()"" scrollamount=""" & Speed & """ scrollDelay=""4"" width=""" & Width & """ height=""" & Height & """>"
					   echoln "<table border=0>"
					    For Each Node in Xml.documentelement.SelectNodes("row")
						  Title = Trim(Node.SelectSingleNode("@title").text) : Content = Trim(node.SelectSingleNode("@content").text) : AddDate = Node.SelectSingleNode("@adddate").text : ID=Node.SelectSingleNode("@id").text
						  
						  echo "<tr><td>"
						  If OpenType = 0 Then
						  
						   echo "<a" & T_CssStr & " href=""#"" onclick=""javascript:window.open('" & DomainStr & "plus/Announce/?" & ID& "','NewWin','height=" & OHeight & ", width=" & OWidth & ", toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no');"" title=""" & Title & """> "
						  Else
						   echo "<a" & T_CssStr & " href=""" & DomainStr & "plus/Announce/?" & ID & """ title=""" & Title & """ target=""_blank""> "
						  End If
						  
						  echoln NaviStr & KS.GotTopic(Title, T_Len) & "</a></td></tr>"
						 If C_Len <> 0 Then
						  echo "<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;" & KS.GotTopic(Replace(Replace(KS.LoseHtml(Content), vbCrLf, ""), "&nbsp;", ""), C_Len) & "…"
						 End If
						 echoln "</td></tr>"
						 If ShowAuthor = 1 Then
						 echoln "<tr><td align=""right"">" & Node.SelectSingleNode("@author").text & "</td></tr>" & vbCrLf & "<tr><td align=""right"">" & Year(AddDate) & "年" & Month(AddDate) & "月" & Day(AddDate) & "日</td></tr>"
						 End If
					  Next
					   echoln "</table>"
					   echoln "</marquee>"
					ElseIf ShowStyle = 2 Then   '横向显示
					   echo "<marquee onmouseover=""this.stop()"" onmouseout=""this.start()"" scrollamount=""" & Speed & """ scrollDelay=""4"" width=""" & Width & """ Height=""" & Height & """ align=""left"">"
					   For Each Node in Xml.documentelement.SelectNodes("row")
						  Title = Trim(Node.SelectSingleNode("@title").text) : AddDate = Node.SelectSingleNode("@adddate").text : ID=Node.SelectSingleNode("@id").text
						  If OpenType = 0 Then
						  echo "<a" & T_CssStr & " href=""#"" onclick=""javascript:window.open('" & DomainStr & "plus/Announce/?" & ID & "','NewWin','height=" & OHeight & ", width=" & OWidth & ", toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no');"" title=""" & Title & """> "
						  Else
						  echo "<a" & T_CssStr & " href=""" & DomainStr & "plus/Announce/?" & ID & """ title=""" & Title & """ target=""_blank""> "
						  End If
						  echo NaviStr & KS.GotTopic(Title, T_Len) & "</a>"
						 If ShowAuthor = 1 Then
						  echo "[" & Node.SelectSingleNode("@author").text & "&nbsp;&nbsp;" & Year(AddDate) & "年" & Month(AddDate) & "月" & Day(AddDate) & "日]"
						 End If
						  echo "&nbsp;&nbsp;"
					  Next
					      echo "</marquee>"
					End If
	 		 end select
			 GetAnnounceList = Templates
		End Function
		
		'取得友情链接列表函数
		Function GetLinkList(LabelStyle)
			 Dim show,FolderID, LinkType, ShowStyle, LogoWidth, LogoHeight, num, T_Len,M,Col,RollWidth,RollHeight,RollSpeed
			 Dim SqlStr, Para,Xml,Node, SiteName,TopStr, URL,TitleStr, WidthStr, FriendLinkRegStr,TemplateFromXml,Recommend
			 If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				GetLinkList="" : Exit Function
			 Else 
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
             LabelId = ParamNode.getAttribute("labelid")
             FolderId = ParamNode.getAttribute("classid")
			 Col      = ParamNode.getAttribute("col")
			 Num      = KS.ChkClng(ParamNode.getAttribute("num")) : If Num<>0 Then TopStr=" top " & num
			 T_Len    = ParamNode.getAttribute("titlelen")
			 LinkType = ParamNode.getAttribute("linktype")
			 Show     = ParamNode.getAttribute("show")
			 ShowStyle = ParamNode.getAttribute("showstyle")
			 LogoWidth = ParamNode.getAttribute("logowidth")
			 LogoHeight= ParamNode.getAttribute("logoheight")      
			 RollWidth = ParamNode.getAttribute("rollwidth")
			 RollHeight= ParamNode.getAttribute("rollheight")  
			 RollSpeed = ParamNode.getAttribute("rollspeed")    
			 Recommend = ParamNode.getAttribute("recommend")

			 Dim k, I, NoLinkRowNumber,TotalNum
			 FriendLinkRegStr = DomainStr & "plus/link/reg" '注册链接
			 WidthStr = CInt(100 / CInt(Col)) & "%"
			 
			 FolderID = CInt(FolderID):LinkType = CInt(LinkType)
			 Para = " Where Locked=0 And Verific=1"
			 If FolderID <> 0 Then  Para = Para & " And FolderID=" & FolderID
			 If Recommend="1" Then Para = Para & " And Recommend=1"
			 If LinkType = 2 Then
			   Para = Para & " Order BY LinkType Desc,OrderID"
			 Else
			   Para = Para & " And LinkType=" & LinkType & " Order BY OrderID,linkid"
			 End If
			 SqlStr = "Select " & TopStr & " LinkID,LinkType,SiteName,Description ,Logo,AddDate,FolderID,Url From KS_Link" & Para
			 Dim RSObj:Set RSObj=Conn.Execute(SqlStr)
			 If Not RSObj.Eof Then Set XML=KS.RsToXml(RSObj,"row","root")
			 RSObj.Close:Set RSObj=nothing
			 Templates = "" : N=0
			 Select Case (CInt(ShowStyle))
                Case 4
				   If IsObject(XML) Then
				 	  Set DocNode=XML.DocumentElement.SelectNodes("row")
					  Templates=ExplainDiyStyle(LabelStyle,DocNode.length)
				   End If				
				 Case 1                '向上滚动
				  echoln " <table width=""100%"" cellSpacing=""2"">"
				  If Not IsObject(XML) Then
					 If FolderID = 0 Then                  '当显示所有类别的友情链接时,显示点击申请
					   For I = 1 To num
						 echoln "<tr align=""center"" height=""22"">"
						 If LinkType = 0 Then
						   echoln "<td><a href=""" & FriendLinkRegStr & """ target=""_blank"" title=""点击申请"">您的位置</a></td>"
						 Else
						   echoln "<td><a href=""" & FriendLinkRegStr & """ target=""_blank"" title=""点击申请""><Img src=""" & DomainStr & "Images/Default/nologo.gif"" border=""0""/></a></td>"
						 End If
						  echoln "</tr>"
					  Next
					End If
				  Else
				    For Each Node In xml.documentelement.SelectNodes("row")
					 echoln "<tr align=""center"" height=""22"">"
					 SiteName = Node.SelectSingleNode("@sitename").text
  				     If Show=0 Then Url=Node.SelectSingleNode("@url").text Else Url=DomainStr & "plus/link/To?" & Node.SelectSingleNode("@linkid").text
					 TitleStr = " title=""网站名称:" & SiteName & "&#13;&#10;网站描述:" & Node.SelectSingleNode("@description").text & """"
						  If Node.SelectSingleNode("@linktype").text = "0" Then
						   echoln "<td><a id=""link" & Node.SelectSingleNode("@linkid").text & """ href=""" & Url & """ target=""_blank""" & TitleStr & ">" & KS.GotTopic(SiteName, T_Len) & "</a></td>"
						  Else
						   echoln "<td><a id=""link" & Node.SelectSingleNode("@linkid").text & """ href=""" & Url & """ target=""_blank""><img src=""" & Node.SelectSingleNode("@logo").text & """" & TitleStr & " alt=""" & SiteName & """  width=""" & LogoWidth & """ height=""" & LogoHeight & """ border=""0""/></a></td>"
						  End If
						  echoln "</tr>"
						I = I + 1
					Next
				  End If
				   echoln "</table>"
				    TemplateFromXml=LFCls.GetConfigFromXML("Label","/labeltemplate/label","rollup")
					TemplateFromXml=Replace(TemplateFromXml,"{$BackGround}","")
					TemplateFromXml=Replace(Replace(Replace(Replace(Replace(TemplateFromXml,"{$Width}",RollWidth),"{$Height}",RollHeight),"{$ImgStr}",Templates),"{$Speed}",RollSpeed),"{$LabelID}",LabelID)
				    Templates = TemplateFromXml
				Case 2                '横向列表
				   echoln " <table width=""100%"" cellspacing=""2""> "
				  If Not IsObject(XML) Then
						If FolderID = 0 Then
						   If num = 0 Then NoLinkRowNumber = 1 Else NoLinkRowNumber = num \ Col
						   For I = 1 To NoLinkRowNumber
							  echoln "<tr align=""center"">"
							  For k = 1 To Col
								If LinkType = 1 Then
								  echo "<td width=""" & WidthStr & """><a href=""" & FriendLinkRegStr & """ target=""_blank"" title=""点击申请""><Img src=""" & DomainStr & "Images/Default/nologo.gif"" alt=""点击申请"" border=""0""/></a></td>"
								Else
								  echo "<td width=""" & WidthStr & """ nowrap=""nowrap""><a href=""" & FriendLinkRegStr & """ target=""_blank"" title=""点击申请"">您的位置</a></td>"
								End If
							  Next
							  echoln "</tr>"
						   Next
						End If
				Else
				    Set DocNode=XML.DocumentElement.SelectNodes("row")
		            TotalNum=DocNode.length
					if TotalNum>Num And Num<>0 Then TotalNum=Num 
				    For M=0 To TotalNum-1
					  If Col = 1 Then echoln "<tr align=""center"">" Else echoln "<tr>"
					  For k = 1 To Col
					      Set Node=DocNode.item(n)
						  SiteName = Node.SelectSingleNode("@sitename").text
						  If Show=0 Then Url=Node.SelectSingleNode("@url").text Else Url=DomainStr & "plus/link/To?" & Node.SelectSingleNode("@linkid").text
						  
						  TitleStr = " title=""网站名称:" & SiteName & "&#13;&#10;网站描述:" & Node.SelectSingleNode("@description").text & """"
						  
						  If Node.SelectSingleNode("@linktype").text = "0" Then
							echo "<td width=""" & WidthStr & """ nowrap=""nowrap""><a id=""link" & Node.SelectSingleNode("@linkid").text & """ href=""" & URL & """ target=""_blank""" & TitleStr & ">" & KS.GotTopic(SiteName, T_Len) & "</a></td>"
						  Else
							echo "<td width=""" & WidthStr & """><a id=""link" & Node.SelectSingleNode("@linkid").text & """ href=""" & URL & """ target=""_blank""><img src=""" & Node.SelectSingleNode("@logo").text & """" & TitleStr & " alt=""" & SiteName & """ width=""" & LogoWidth & """ height=""" & LogoHeight & """ border=""0""/></a></td>"
						  End If
						  N = N+1 : If N>=TotalNum Then Exit For
					  Next
					  '不到Col个单元格,则进行补空
					  for  k=k+1 to Col
							If LinkType = 1 Then
								   echo "<td width=""" & WidthStr & """><a href=""" & FriendLinkRegStr & """ target=""_blank""  title=""点击申请""><Img src=""" & DomainStr & "Images/Default/nologo.gif"" alt=""点击申请"" border=""0""/></a></td>"
							 Else
								   echo "<td width=""" & WidthStr & """ nowrap=""nowrap""><a href=""" & FriendLinkRegStr & """ target=""_blank"" title=""点击申请"">您的位置</a></td>"
							End If
					 next
					 echoln "</tr>"
					 If N>=TotalNum Then Exit For
				   Next
				  End If
				echo "</table>"
				Case 3                '下拉列表
				  echoln "<select name=""FriendLink"" onchange=""if(this.options[this.selectedIndex].value!=''){window.open(this.options[this.selectedIndex].value,'_blank');}"">"
				 If Not IsObject(XML) Then
				  echoln  "<option value=''>---没有任何链接---</option>"
				 Else
					 For Each Node In XML.DocumentElement.SelectNodes("row")
					   If N=0 Then
						 echoln  "<option value=''>---" & Conn.Execute("Select FolderName From KS_LinkFolder Where FolderID=" & Node.SelectSingleNode("@folderid").text)(0) & "---</option>"
					   End If
					   N=N+1
					   If Show=0 Then Url=Node.SelectSingleNode("@url").text Else Url=DomainStr & "plus/link/to?" & Node.SelectSingleNode("@linkid").text
					   echoln "<option value='" & Url & "'>" & KS.GotTopic(Node.SelectSingleNode("@sitename").text, T_Len) & "</option>"
					 Next
				End If
				  echoln "</select>"
			 End Select
			 XML=Empty
			 GetLinkList = Templates
		End Function
		
		
		'==============================求职系统==============================
		Function GetJobList(LabelStyle)
		     If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				GetJobList="" : Exit Function
			 Else 
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
		  GetJobList=JLCls.GetJobList(ParamNode,LabelStyle)
		End Function
		Function GetJobZWList(LabelStyle)
		     If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				GetJobZWList="" : Exit Function
			 Else 
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
		  GetJobZWList=JLCls.GetJobZWL(ParamNode,LabelStyle)
		End Function
		Function GetJobResumeList(LabelStyle)
		     If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				GetJobResumeList="" : Exit Function
			 Else 
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
		  GetJobResumeList=JLCls.GetJobResume(ParamNode,LabelStyle)
		End Function
		'===============================求职标签结束========================================
		
		'===============================问答系统=============================================
		Function GetAQList(LabelStyle)
             If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				 GetAQList="标签加载出错" : Exit Function
			 Else
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
		     LabelID   = ParamNode.getAttribute("labelid")
			 ClassID   = ParamNode.getAttribute("classid") : If Not IsNumeric(ClassID) Then ClassID=0
			 AjaxOut   = ParamNode.getAttribute("ajaxout")
			 Num       = ParamNode.getAttribute("num")

			 If LabelID<>"ajax" and Cbool(ParamNode.getAttribute("ajaxout"))=true Then 
			  GetAQList="<span id=""ks" & LabelID & "_0_0_0_0""></span>":Exit Function
			 End If
		
 			 Dim SqlStr,RecommendTF,RewardTF,ZeroTF,OrderStr,JoinType,JJTF
			 RecommendTF= ParamNode.getAttribute("recommendtf") : If KS.IsNul(RecommendTF) Then RecommendTF=False
			 RewardTF   = ParamNode.getAttribute("rewardtf") : If KS.IsNul(RewardTF) Then RewardTF=false
			 ZeroTF     = ParamNode.getAttribute("zerotf")   : If KS.IsNul(ZeroTF) Then ZeroTF=false
			 JJTF       = ParamNode.getAttribute("jjtf")   : If KS.IsNul(JJTF) Then JJTF=false
			 OrderStr   = ParamNode.getAttribute("infosort") : If KS.IsNul(OrderStr) Then OrderStr="topicid desc"

			 Param= " Where a.locktopic=0"
			 If Cbool(RecommendTF)=true Then Param=Param & " And a.recommend=1"
			 If Cbool(RewardTF)=true Then Param=Param & " And A.reward>0"
			 If Cbool(ZeroTF)=true Then Param=Param & " And a.postnum=0"
			 If Cbool(JJTF)=true Then Param=Param & " And a.TopicMode=1"
			 If ClassID<>0 Then Param=Param & " And a.classid in (SELECT classid FROM KS_AskClass WHERE ','+parentstr+'' like '%,"&classid&",%')"
			 Param=Param & " Order By " & OrderStr
			 
			 If ParamNode.getAttribute("printtype")="2" or Lcase(ParamNode.getAttribute("showuserface"))="true" Then
			  If KS.ASetting(47)="1" Then JoinType="Left Join" Else JoinType="Inner Join"
			  SqlStr="Select TOP " & num & " a.TopicID,a.ClassID as AqClassId,a.UserName,a.ClassName as AQClassName,a.pclassname,a.Title,a.bestusername,a.TopicMode,a.DateAndTime as AddDate,a.hits,a.reward,a.Anonymous,a.LastPostTime,a.ExpiredTime,b.userface From KS_AskTopic a " & JoinType &" KS_User b On A.userName=B.userName" & Param 
			 Else
			  SqlStr="Select TOP " & num & " a.TopicID,a.ClassID as AqClassId,a.UserName,a.ClassName as AQClassName,a.pclassname,a.Title,a.bestusername,a.TopicMode,a.DateAndTime as AddDate,a.reward,a.Anonymous From KS_AskTopic a " & Param
			 End If
			
			Dim RS:Set RS=Conn.Execute(SqlStr)
			If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			RS.Close:Set RS=Nothing
			If IsObject(XMLSql) Then
			 GetAQList=ExplainAQListLabelBody(LabelStyle,1)
			End If 
			Set Node=Nothing
		End Function
		Function GetAqUrl()
		  If KS.ASetting(16)="1" Then
		  GetAqUrl=DomainStr & KS.ASetting(1) & "show-" & Node.SelectSingleNode("@topicid").text & KS.ASetting(17)
		  Else
		  GetAqUrl=DomainStr & KS.ASetting(1) & "q.asp?id=" & Node.SelectSingleNode("@topicid").text
		  End If
		End Function
		Function GetAQClassUrl()
		 If KS.ASetting(16)="1" Then
		  GetAQClassUrl=DomainStr & KS.ASetting(1) & "list-" & Node.SelectSingleNode("@aqclassid").text & KS.ASetting(17)
		 Else
		  GetAQClassUrl=DomainStr & KS.ASetting(1) & "showlist.asp?id=" & Node.SelectSingleNode("@aqclassid").text
		 End If
		End Function
		Function ExplainAQListLabelBody(LabelStyle,GetType)
		  Dim PrintType,RowHeight,NaviStr,T_CssStr,OpenType,T_Len,SplitPic,DateStr,DateRule,ShowClass,ShowUserName,ShowUserFace,ShowReward,OpenTypeStr,userface,username,SpaceUrl
		  PrintType = ParamNode.getAttribute("printtype")
		  RowHeight = ParamNode.getAttribute("rowheight") : If Not IsNumeric(RowHeight) Then RowHeight=20
		  NaviStr   = KS.GetNavi(ParamNode.getAttribute("navtype"), ParamNode.getAttribute("nav"))
		  T_CssStr  = KS.GetCss(ParamNode.getAttribute("titlecss"))
		  OpenType  = ParamNode.getAttribute("opentype") : OpenTypeStr=KS.G_O_T_S(opentype)
		  T_Len     = ParamNode.getAttribute("titlelen")
		  SplitPic  = ParamNode.getAttribute("splitpic")
		  DateRule  = ParamNode.getAttribute("daterule")
		  ShowClass = ParamNode.getAttribute("showclass") : If KS.IsNul(ShowClass) Then ShowClass=false
		  ShowUserName=ParamNode.getAttribute("showusername") : If KS.IsNul(ShowUserName) Then ShowUserName=false
		  ShowUserFace=ParamNode.getAttribute("showuserface") : If KS.IsNul(ShowUserFace) Then ShowUserFace=False
		  ShowReward  =ParamNode.getAttribute("showreward") : If KS.IsNul(ShowReward) Then ShowReward=false
		  
		  Set DocNode=XMLSql.DocumentElement.SelectNodes("row")
		  Templates  = "" : N = 0 
		  If PrintType="1" Then
		         echoln "<table border=""0"" width=""100%"" cellpadding==""0"" cellspacing=""0"">"
				  For Each Node In XMLSql.DocumentElement.SelectNodes("row")
				    DateStr=GetDateStr(1,Node.SelectSingleNode("@adddate").text,DateRule,"left","",1,1)
					echoln "<tr><td height=""" & RowHeight & """>"
					echo NaviStr
					If Cbool(ShowUserName)=true Then  
					UserName=Node.SelectSingleNode("@username").text : SpaceUrl="href=""" &DomainStr & "space/?" & UserName & """ target=""_blank"""
					If GetType=1 Then
					 If Node.SelectSingleNode("@anonymous").text="1" Then UserName="匿名" : SpaceUrl="href=""#"""
					 End If
					End If
					If Cbool(ShowUserFace)=true Then 
   				     UserFace=Node.SelectSingleNode("@userface").text : If KS.IsNul(UserFace) Then UserFace=DomainStr & "images/nopic.gif" 
					 If Left(Lcase(userface),4)<>"http" then UserFace=DomainStr & UserFace
					 
					 If GetType=1 Then
					 If Node.SelectSingleNode("@anonymous").text="1" Then UserFace=DomainStr & "images/face/0.gif" 
					 End If
					 echo "<a " &SpaceUrl & " class=""face""><img src=""" & UserFace & """ border=""0"" width=""32"" align=""absmiddle"" height=""32""></a> "
					End If
					If Cbool(ShowUserName)=true Then  
					 echo "<a " &SpaceUrl & ">" & username & "</a> " 
					 If GetType=1 Then	echo " 提问了 " else echo " 回答了 "
					End If
					If Cbool(ShowClass)=true And Cbool(ShowUserName)=false And Cbool(ShowUserFace)=False Then echo "<span class=""category""><a href=""" & GetAQClassUrl & """" & OpenTypeStr&">[" & Node.SelectSingleNode("@aqclassname").text & "]</a></span>"
					echo "<a href=""" & GetAqurl & """" &T_CssStr& OpenTypeStr & ">" & KS.GotTopic(Node.SelectSingleNode("@title").text,T_Len) &"</a>"
					If Cbool(ShowClass)=true And (Cbool(ShowUserName)=true Or Cbool(ShowUserFace)=true) Then echo " <span class=""category""><a href=""" & GetAQClassUrl & """" & OpenTypeStr&">[" & Node.SelectSingleNode("@aqclassname").text & "]</a></span>"
					If GetType=1 Then
					If Cbool(ShowReward)=true and Not KS.IsNul(Node.SelectSingleNode("@reward").text) Then If KS.ChkClng(Node.SelectSingleNode("@reward").text)>0 Then echo "<img src=""" & DomainStr & KS.ASetting(1) & "images/ask_xs.gif"">" & Node.SelectSingleNode("@reward").text
					End If
					echo datestr
					echoln "</td></tr>"
					echoln KS.GetSplitPic(SplitPic,1)
				  Next
				  echoln "</table>"
		  Else
		   Templates=ExplainDiyStyle(LabelStyle,DocNode.length)
		  End If
		  ExplainAQListLabelBody=Templates
		End Function
		
		Function GetAAList(LabelStyle)
             If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				 GetAAList="标签加载出错" : Exit Function
			 Else
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
		     LabelID   = ParamNode.getAttribute("labelid")
			 ClassID   = ParamNode.getAttribute("classid") : If Not IsNumeric(ClassID) Then ClassID=0
			 AjaxOut   = ParamNode.getAttribute("ajaxout")
			 Num       = ParamNode.getAttribute("num")

			 If LabelID<>"ajax" and Cbool(ParamNode.getAttribute("ajaxout"))=true Then 
			  GetAAList="<span id=""ks" & LabelID & "_0_0_0_0""></span>":Exit Function
			 End If
		
 			 Dim SqlStr,SolutionTF,OrderStr
			 SolutionTF = ParamNode.getAttribute("SolutionTF") : If KS.IsNul(SolutionTF) Then SolutionTF=2
			 OrderStr   = ParamNode.getAttribute("infosort") : If KS.IsNul(OrderStr) Then OrderStr="a.AnswerID desc"
			 

			 Param= " Where 1=1"
			 If KS.ChkClng(SolutionTF)<>2 Then Param=Param & " And a.TopicMode=" & KS.ChkClng(SolutionTF)
			 If ClassID<>0 Then Param=Param & " And a.classid in (SELECT classid FROM KS_AskClass WHERE ','+parentstr+'' like '%,"&classid&",%')"
			 Param=Param & " Order By " & OrderStr
			 
			 If ParamNode.getAttribute("printtype")="2" or Lcase(ParamNode.getAttribute("showuserface"))="true" Then
			  SqlStr="Select TOP " & num & " a.TopicID,a.ClassID as AqClassId,a.UserName,a.ClassName as AQClassName,a.pclassname,a.Title,a.AnswerTime as AddDate,b.userface From KS_AskAnswer a Inner Join KS_User b On A.userName=B.userName" & Param
			 Else
			  SqlStr="Select TOP " & num & " a.TopicID,a.ClassID as AqClassId,a.UserName,a.ClassName as AQClassName,a.pclassname,a.Title,a.AnswerTime as AddDate From KS_AskAnswer a " & Param
			 End If
			Dim RS:Set RS=Conn.Execute(SqlStr)
			If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			RS.Close:Set RS=Nothing
			If IsObject(XMLSql) Then
			 GetAAList=ExplainAQListLabelBody(LabelStyle,2)
			End If 
			Set Node=Nothing
		End Function
		
		Function GetAskZJList(LabelStype)
             If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				 GetAskZJList="标签加载出错" : Exit Function
			 Else
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
		     LabelID   = ParamNode.getAttribute("labelid")
			 AjaxOut   = ParamNode.getAttribute("ajaxout")
			 Num       = KS.ChkClng(ParamNode.getAttribute("num")) : If Num<=0 Then Num=10
			
			 If LabelID<>"ajax" and Cbool(ParamNode.getAttribute("ajaxout"))=true Then 
			  GetAskZJList="<span id=""ks" & LabelID & "_0_0_0_0""></span>":Exit Function
			 End If
			 
			 Dim OrderStr,SqlStr,BigClassID,SmallClassID,SmallerClassID,recommend,RS,classtype 
			 BigClassID   = KS.ChkClng(ParamNode.getAttribute("bigclassid"))
			 SmallClassID = KS.ChkClng(ParamNode.getAttribute("smallclassid"))
			 SmallerClassID = KS.ChkClng(ParamNode.getAttribute("smallerclassid"))
			 recommend = KS.ChkClng(ParamNode.getAttribute("recommend"))
			 classtype = KS.ChkClng(ParamNode.getAttribute("classtype"))
			 OrderStr   = ParamNode.getAttribute("orderstr") : If KS.IsNul(OrderStr) Then OrderStr="id desc"
			 If Lcase(Left(Trim(OrderStr),2))<>"id" Then  OrderStr=OrderStr & ",ID Desc"
			 
			 Param= " Where status=1"
			 If classtype=2 Then
				 If BigClassID<>0 Then Param=Param & " and BigClassID=" &BigClassID
				 If SmallClassID<>0 Then Param=Param & " and SmallClassID=" & SmallClassID
				 If SmallerClassID<>0 Then Param=Param & " and SmallerClassID=" & SmallClassID
			 ElseIf KS.ChkClng(FCls.RefreshFolderID)<>0 Then
			    Param=Param & " and (BigClassID=" &KS.ChkClng(FCls.RefreshFolderID) & " or smallclassid=" & KS.ChkClng(FCls.RefreshFolderID) & " or smallerclassid=" & KS.ChkClng(FCls.RefreshFolderID) &")"
			 End If
			 If recommend=1 Then Param=Param & " and recommend=1"
			 Param=Param & " Order By " & OrderStr
			 
			 SqlStr="Select TOP " & num & " UserID,UserName,RealName,QQ,AddDate,AskDoneNum,UserFace,Tel,Province,City,Intro, '' as [Domain],BigClassID,SmallClassID From KS_AskZJ " & Param
			If DataBaseType=1 Then
			    Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
				Set Cmd.ActiveConnection=conn
				Cmd.CommandText="KS_TSqls"
				Cmd.CommandType=4	
				CMD.Prepared = true 	
				Cmd.Parameters.Append cmd.CreateParameter("@SQL",202,1,8000,SQLStr)
				Set Rs=Cmd.Execute
				Set Cmd=Nothing
			Else
			    Set RS=Conn.Execute(SqlStr)
			End If
			If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			RS.Close:Set RS=Nothing
			If IsObject(XMLSql) Then
			 Set DocNode=XMLSql.DocumentElement.SelectNodes("row")
			 GetAskZJList=ExplainDiyStyle(LabelStyle,DocNode.length)
			End If 
			Set Node=Nothing
		End Function
		'===============================问答结束=============================================
		
		'论坛帖子
		Function GetClubList(LabelStyle)
             If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				 GetClubList="标签加载出错" : Exit Function
			 Else
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
		     LabelID   = ParamNode.getAttribute("labelid")
			 ClassID   = ParamNode.getAttribute("classid") : If Not IsNumeric(ClassID) Then ClassID=0
			 AjaxOut   = ParamNode.getAttribute("ajaxout")
			 Num       = ParamNode.getAttribute("num") : IF Not IsNumeric(Num) Then Num=10

			 If LabelID<>"ajax" and Cbool(ParamNode.getAttribute("ajaxout"))=true Then 
			  GetClubList="<span id=""ks" & LabelID & "_0_0_0_0""></span>":Exit Function
			 End If
			 
			 Dim OrderStr,SqlStr,BN,Bids,ShowJh,RS
			 ShowJh     = ParamNode.getAttribute("showjh") : If KS.IsNul(ShowJh) Then ShowJh=false
			 OrderStr   = ParamNode.getAttribute("infosort") : If KS.IsNul(OrderStr) Then OrderStr="id desc"
			 
			  Param= " Where verific=1 and deltf=0"
			 If Instr(KS.FilterIds(ClassID),",")<>0 Then
			   Param=Param & " And a.boardid in ("&KS.FilterIds(ClassID)&")" 
			 ElseIf ClassID<>0 Then 
			  KS.LoadClubBoard
			  For Each BN In Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectNodes("row[@parentid=" & ClassID & "]")
			   If Bids="" Then
			    Bids=BN.SelectSingleNode("@id").text
			   Else
			    Bids=Bids & "," & BN.SelectSingleNode("@id").text
			   End If
			  Next
			  If Not KS.IsNul(bids) Then  Param=Param & " And a.boardid in ("&bids&")" Else Param=Param & " And a.boardid=" & ClassID
			 End If
			 
			 if cbool(ShowJh)=true then Param=Param & " and a.isbest=1"
			 Param=Param & " Order By " & OrderStr
			 If instr(LabelStyle,"{@userface}")<>0 or Lcase(ParamNode.getAttribute("showuserface"))="true" Then
			  SqlStr="Select TOP " & num & " a.ID,a.subject,a.boardid,a.addtime as AddDate,a.hits,a.totalreplay,a.username,a.LastReplayTime as lastposttime,b.userface From KS_GuestBook a inner join KS_User b On A.userName=B.userName" & Param
			 Else
			  SqlStr="Select TOP " & num & " ID,subject,boardid,addtime as AddDate,a.LastReplayTime as lastposttime,hits,totalreplay,username From KS_GuestBook a " & Param
			 End If
			If DataBaseType=1 Then
			    Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
				Set Cmd.ActiveConnection=conn
				Cmd.CommandText="KS_TSqls"
				Cmd.CommandType=4	
				CMD.Prepared = true 	
				Cmd.Parameters.Append cmd.CreateParameter("@SQL",202,1,8000,SQLStr)
				Set Rs=Cmd.Execute
				Set Cmd=Nothing
			Else
			    Set RS=Conn.Execute(SqlStr)
			End If
			If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			RS.Close:Set RS=Nothing
			If IsObject(XMLSql) Then
			 GetClubList=ExplainClubListLabelBody(LabelStyle)
			End If 
			Set Node=Nothing
        End Function
		Function GetClubUrl()
		   GetClubUrl=KS.GetClubShowUrl(Node.SelectSingleNode("@id").text)
		End Function
		Function GetBoardInfo(FieldName)
		   KS.LoadClubBoard
		   Dim BN:Set BN=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & Node.SelectSingleNode("@boardid").text &"]")
		  If Not BN Is Nothing Then
		    GetBoardInfo=BN.SelectSingleNode("@" & FieldName).text
		  End If
		  Set BN=Nothing
		End Function
		Function ExplainClubListLabelBody(LabelStyle)
		 Dim PrintType,RowHeight,NaviStr,T_CssStr,OpenType,OpenTypeStr,T_Len,SplitPic,DateRule,DateStr,ShowClass,ShowUserName,ShowUserFace,UserFace
		  PrintType= ParamNode.getAttribute("printtype")
		  RowHeight = ParamNode.getAttribute("rowheight") : If Not IsNumeric(RowHeight) Then RowHeight=20
		  NaviStr   = KS.GetNavi(ParamNode.getAttribute("navtype"), ParamNode.getAttribute("nav"))
		  T_CssStr  = KS.GetCss(ParamNode.getAttribute("titlecss"))
		  OpenType  = ParamNode.getAttribute("opentype") : OpenTypeStr=KS.G_O_T_S(opentype)
		  T_Len     = ParamNode.getAttribute("titlelen")
		  SplitPic  = ParamNode.getAttribute("splitpic")
		  DateRule  = ParamNode.getAttribute("daterule")
		  ShowClass = ParamNode.getAttribute("showclass") : If KS.IsNul(ShowClass) Then ShowClass=false
		  ShowUserName=ParamNode.getAttribute("showusername") : If KS.IsNul(ShowUserName) Then ShowUserName=false
		  ShowUserFace=ParamNode.getAttribute("showuserface") : If KS.IsNul(ShowUserFace) Then ShowUserFace=False

		 Set DocNode=XMLSql.DocumentElement.SelectNodes("row")
		  Templates  = "" : N = 0 
		  If PrintType="1" Then
		         echoln "<table border=""0"" width=""100%"" cellpadding==""0"" cellspacing=""0"">"
				  For Each Node In XMLSql.DocumentElement.SelectNodes("row")
				    DateStr=GetDateStr(1,Node.SelectSingleNode("@adddate").text,DateRule,"left","",1,1)
				    echoln "<tr>"
					echoln "<td height=""" & RowHeight & """>" 
					echo NaviStr
					
					If Cbool(ShowUserFace) Then
					   UserFace=Node.SelectSingleNode("@userface").text : If KS.IsNul(UserFace) Then UserFace=DomainStr & "images/nopic.gif" 
					   echo "<a href=""" &DomainStr & "space/?" & Node.SelectSingleNode("@username").text & """ target=""_blank"" class=""face""><img src=""" & UserFace & """ border=""0"" width=""32"" align=""absmiddle"" height=""32""></a> "
					End If
					
					If Cbool(ShowUserName)=true Then 
					 If KS.IsNul(Node.SelectSingleNode("@username").text) Then
					 echo "<a href=""#"">游客</a> 发表了"
					 Else 
					 echo "<a href=""" & DomainStr & "space/?" & Node.SelectSingleNode("@username").text & """ target=""_blank"">" & Node.SelectSingleNode("@username").text & "</a> 发表了 " 
					 End If
					End If
					If Cbool(ShowClass)=true And Cbool(ShowUserName)=false And Cbool(ShowUserFace)=False and Node.SelectSingleNode("@boardid").text<>"0" Then echo "<span class=""category""><a href=""" & DomainStr & "club/index.asp?boardid=" & Node.SelectSingleNode("@boardid").text & """" & OpenTypeStr&">[" & GetBoardInfo("boardname") & "]</a></span>"
					echo  "<a href=""" & GetClubUrl & """"&OpenTypeStr&T_CssStr &">" & KS.Gottopic(KS.LoseHtml(Node.SelectSingleNode("@subject").text),T_Len) & "</a> "
					If Cbool(ShowClass)=true And (Cbool(ShowUserName)=true Or Cbool(ShowUserFace)=true) and Node.SelectSingleNode("@boardid").text<>"0" Then echo " <span class=""category""><a href=""" & DomainStr & "club/index.asp?boardid=" & Node.SelectSingleNode("@boardid").text & """" & OpenTypeStr&">[" & GetBoardInfo("boardname") & "]</a></span>"
					echo DateStr & "</td>"
					echoln "</tr>"
					echoln KS.GetSplitPic(SplitPic,1)
				  Next
				 echoln "</table>"
		  Else 
		      Templates=ExplainDiyStyle(LabelStyle,DocNode.length)
		  End If
		 ExplainClubListLabelBody=Templates
		End Function
		
		Function GetSJList(LabelStyle)
		     If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				 GetSJList="标签加载出错" : Exit Function
			 Else
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
			 Dim SmallClassID,DateRule,OrderStr,Dtfs,Recommend,Popular
			 ClassID   = ParamNode.getAttribute("bigclassid") : If Not IsNumeric(ClassID) Then ClassID=0
			 SmallClassID = ParamNode.getAttribute("smallclassid") : If Not IsNumeric(SmallClassID) Then SmallClassID=0
		     LabelID   = ParamNode.getAttribute("labelid")
			 AjaxOut   = ParamNode.getAttribute("ajaxout")
			 DateRule  = ParamNode.getAttribute("daterule")
			 Num       = ParamNode.getAttribute("num")
			 Dtfs      = KS.ChkClng(ParamNode.getAttribute("dtfs"))
			 Recommend = KS.ChkClng(ParamNode.getAttribute("recommend"))
			 Popular   = KS.ChkClng(ParamNode.getAttribute("popular"))
			 OrderStr  =ParamNode.getAttribute("orderstr") : If OrderStr="" Then OrderStr="id desc"
			 Param="Where verific=1"
			 If FCls.CallFrom3G="true" Then Param=Param & " and dtfs<>0"
			 If SmallClassID<>0 Then 
			  param=param & " and tid in(select id from ks_sjclass where ts like '%" & SmallClassID &",%')"
			 ElseIf ClassID<>0 Then
			  param=param & " and tid in(select id from ks_sjclass where ts like '%" & ClassID &",%')"
			 End If
			 If Dtfs<>0 Then Param=Param & " and dtfs=" & dtfs
			 If Recommend<>0 Then Param=Param & " and recommend=1"
			 If Popular<>0 Then Param=Param & " and Popular=1"

			 If Instr(lcase(orderStr),"hits")<>0 Then Param=Param & " order by " & OrderStr & ",a.id desc" Else Param=Param &" order by a." & OrderStr
			 If LabelID<>"ajax" and Cbool(ParamNode.getAttribute("ajaxout"))=true Then 
			  GetSJList="<span id=""ks" & LabelID & "_0_0_0_0""></span>":Exit Function
			 End If
			 Dim SqlStr
			 If Instr(LabelStyle,"{@sjtypename}")<>0 Then
			 SqlStr="Select TOP " & num & " b.tname,a.id,a.tid,a.title,a.kssj,a.form_user,a.form_url,a.hits,a.dtfs,[date] as adddate,a.sq,[user],a.sjzf From KS_SJ a inner join ks_sjclass b on a.tid=b.id " & Param
			 Else
			 SqlStr="Select TOP " & num & " a.tid,a.id,a.title,a.kssj,a.form_user,a.form_url,a.hits,a.dtfs,[date] as adddate,a.sq,[user],a.sjzf From KS_SJ a " & Param
			 End If
			Dim RS:Set RS=Conn.Execute(SqlStr)
			If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			RS.Close:Set RS=Nothing
			If IsObject(XMLSql) Then
			 Set DocNode=XMLSql.DocumentElement.SelectNodes("row")
			 GetSJList=ExplainDiyStyle(LabelStyle,DocNode.length)
			End If 
			Set Node=Nothing
		End Function
		
		Function GetGroupBuyList(LabelStyle)
			 If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				 GetGroupBuyList="标签加载出错" : Exit Function
			 Else
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
			 Dim ClassID,DateRule,OrderStr,SQLStr,Recommend
			 ClassID   = ParamNode.getAttribute("classid") : If Not IsNumeric(ClassID) Then ClassID=0

		     LabelID   = ParamNode.getAttribute("labelid")
			 AjaxOut   = ParamNode.getAttribute("ajaxout")
			 If LabelID<>"ajax" and Cbool(AjaxOut)=true Then 
			  GetGroupBuyList="<span id=""ks" & LabelID & "_0_0_0_0""></span>":Exit Function
			 End If
			 DateRule  = ParamNode.getAttribute("daterule")
			 Num       = KS.ChkClng(ParamNode.getAttribute("num")): If Num=0 Then Num=10
			 Recommend = KS.ChkClng(ParamNode.getAttribute("recommend"))
			 OrderStr  =ParamNode.getAttribute("orderstr") : If OrderStr="" Then OrderStr="id desc"
			 Param="Where Endtf=0 and Locked=0"
			 If ClassID<>0 Then Param=Param & " and classid=" & ClassID
			 If Recommend<>0 Then Param=Param & " and recommend=1"
			 If Instr(lcase(orderStr),"adddate")<>0 Then Param=Param & " order by " & OrderStr & ",id desc" Else Param=Param &" order by " & OrderStr
			 SqlStr="Select Top " & Num & " * From KS_GroupBuy " & Param
            Dim RS:Set RS=Conn.Execute(SqlStr)
			If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			RS.Close:Set RS=Nothing
			If IsObject(XMLSql) Then
			 Set DocNode=XMLSql.DocumentElement.SelectNodes("row")
			 GetGroupBuyList=ExplainDiyStyle(LabelStyle,DocNode.length)
			End If 
			Set Node=Nothing
		End Function
		
		Function GetEnterpriseNewsList(LabelStyle)
		     If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				 GetEnterpriseNewsList="标签加载出错" : Exit Function
			 Else
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
			 Dim SmallClassID,DateRule,OrderStr
		     LabelID   = ParamNode.getAttribute("labelid")
			 ClassID   = ParamNode.getAttribute("bigclassid") : If Not IsNumeric(ClassID) Then ClassID=0
			 SmallClassID   = ParamNode.getAttribute("smallclassid") : If Not IsNumeric(SmallClassID) Then SmallClassID=0
			 AjaxOut   = ParamNode.getAttribute("ajaxout")
			 DateRule  = ParamNode.getAttribute("daterule")
			 Num       = ParamNode.getAttribute("num")
			 OrderStr  =ParamNode.getAttribute("orderstr") : If OrderStr="" Then OrderStr="id desc"
			  Param="Where status=1"
			 If ClassID<>0 Then Param=Param & " And BigClassID=" & ClassID
			 If SmallClassID<>0 Then Param=Param & " And SmallClassID=" & SmallClassID
             If ParamNode.getAttribute("callbyspace")="true" And Not KS.IsNul(Session("SpaceUserName")) Then Param=Param & " And username='" & Session("SpaceUserName") & "'"
			 If Instr(lcase(orderStr),"hits")<>0 Then Param=Param & " order by " & OrderStr & ",id desc" Else Param=Param &" order by " & OrderStr
			 If LabelID<>"ajax" and Cbool(ParamNode.getAttribute("ajaxout"))=true Then 
			  GetEnterpriseNewsList="<span id=""ks" & LabelID & "_0_0_0_0""></span>":Exit Function
			 End If
			 Dim SqlStr:SqlStr="Select TOP " & num & " * From KS_EnterPriseNews " & Param
			Dim RS:Set RS=Conn.Execute(SqlStr)
			If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			RS.Close:Set RS=Nothing
			If IsObject(XMLSql) Then
			 Set DocNode=XMLSql.DocumentElement.SelectNodes("row")
			 GetEnterpriseNewsList=ExplainDiyStyle(LabelStyle,DocNode.length)
			End If 
			Set Node=Nothing
		End Function
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:GetSpaceList
		'作 用: 空间列表标签函数
		'参 数: LabelStyle 标签样式
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function GetSpaceList(LabelStyle)
             If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				 GetSpaceList="标签加载出错" : Exit Function
			 Else
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
		     LabelID   = ParamNode.getAttribute("labelid")
			 ClassID   = ParamNode.getAttribute("classid") : If Not IsNumeric(ClassID) Then ClassID=0
			 AjaxOut   = ParamNode.getAttribute("ajaxout")
			 Num       = ParamNode.getAttribute("num")
			 
			 If LabelID<>"ajax" and Cbool(ParamNode.getAttribute("ajaxout"))=true Then 
			  GetSpaceList="<span id=""ks" & LabelID & "_0_0_0_0""></span>":Exit Function
			 End If
			 
			Dim ShowType:ShowType = ParamNode.getAttribute("showtype")
			Dim recommend:recommend=ParamNode.getAttribute("recommend"):If KS.IsNul(recommend) Then recommend=false
			Dim logo:logo=ParamNode.getAttribute("logo") : If KS.IsNul(logo) Then logo=false
			Dim banner:banner=ParamNode.getAttribute("banner") : If KS.IsNul(banner) Then banner=false
			Dim OrderStr:OrderStr  =ParamNode.getAttribute("orderstr") : If OrderStr="" Then OrderStr="b.blogid desc"
             
			If ShowType=2 Then
			 Param=" inner join KS_Enterprise U on B.UserName=U.UserName Where u.Status=1"
			ElseIf ShowType=1 Then
			 Param=" inner join KS_User U on B.UserName=U.UserName Where B.Status=1 And U.UserType=1" 
			Else
			 Param=" Where B.Status=1"
			End If
			
			If ClassID<>0 Then Param=Param & " And b.ClassID=" & ClassID
			If Cbool(recommend)=true Then Param=Param & " And B.recommend=1"
			If Cbool(banner)=true Then Param=Param & " And b.banner<>''"
			If Cbool(logo)=true Then Param=Param & " And b.logo<>''"
			
			If Instr(lcase(orderStr),"hits")<>0 Then Param=Param & " order by " & OrderStr & ",b.blogid desc" Else Param=Param &" order by " & OrderStr
			Dim SqlStr
			If ShowType=2 Then
			SqlStr="Select TOP " & num & " b.userid,b.username,u.companyname as blogname,[Domain],b.logo,b.banner,b.hits From KS_Blog B " & Param
			Else
			SqlStr="Select TOP " & num & " b.userid,b.username,b.blogname,[Domain],b.logo,b.banner,b.hits From KS_Blog B " & Param
			End If
			
			Dim RS:Set RS=Conn.Execute(SqlStr)
			If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			RS.Close:Set RS=Nothing
			If IsObject(XMLSql) Then
			 GetSpaceList=ExplainSpaceListLabelBody(LabelStyle)
			End If 
			Set Node=Nothing
		End Function
		
		Function GetSpaceUrl()
		  If KS.SSetting(14)<>"1" and  Node.SelectSingleNode("@domain").text<>"" then 
		     if instr(Node.SelectSingleNode("@domain").text,".")=0 then
			 GetSpaceUrl="http://" & Node.SelectSingleNode("@domain").text & "." & KS.SSetting(16) 
			 else
			 GetSpaceUrl="http://" & Node.SelectSingleNode("@domain").text
			 end if
		  ElseIf KS.SSetting(21)="1" Then
			GetSpaceUrl=DomainStr & "space/" & server.URLEncode(Node.SelectSingleNode("@userid").text)
		  Else
			GetSpaceUrl=DomainStr & "space/?" & server.URLEncode(Node.SelectSingleNode("@userid").text)
		  End If
		End Function
		Function GetEnterpriseNewsUrl()
		  If KS.SSetting(14)<>"0" and  Node.SelectSingleNode("@domain").text<>"" then 
		     If KS.SSetting(21)="1" Then
			 GetEnterpriseNewsUrl="http://" & Node.SelectSingleNode("@domain").text & "." & KS.SSetting(16) & "/show-news-" & Node.SelectSingleNode("@userid").text & "-" & Node.SelectSingleNode("@id").text & KS.SSetting(22)
			 Else
			 GetEnterpriseNewsUrl="http://" & Node.SelectSingleNode("@domain").text & "." & KS.SSetting(16) & "/?" & Node.SelectSingleNode("@userid").text & "/shownews/" & Node.SelectSingleNode("@id").text 
			 End If
		  ElseIf KS.SSetting(21)="1" Then
			GetEnterpriseNewsUrl=DomainStr & "space/show-news-" & Node.SelectSingleNode("@userid").text & "-" & Node.SelectSingleNode("@id").text & KS.SSetting(22)
		  Else
			GetEnterpriseNewsUrl=DomainStr & "space/?" & Node.SelectSingleNode("@userid").text & "/shownews/" & Node.SelectSingleNode("@id").text
		  End If
		End Function
		
		Function ExplainSpaceListLabelBody(LabelStyle)
		  Dim TotalNum,PrintType
		  Dim T_CssStr,NaviStr,Param,RStr,RowHeight,T_Len,SplitPic,OpenType,MoreStr
		  T_CssStr  = KS.GetCss(ParamNode.getAttribute("titlecss"))
		  NaviStr   = KS.GetNavi(ParamNode.getAttribute("navtype"), ParamNode.getAttribute("nav"))
		  PrintType = ParamNode.getAttribute("printtype")
		  RowHeight = ParamNode.getAttribute("rowheight") : If Not IsNumeric(RowHeight) Then RowHeight=20
		  T_Len     = KS.ChkClng(ParamNode.getAttribute("titlelen"))
		  SplitPic  = ParamNode.getAttribute("splitpic")
		  OpenType  = ParamNode.getAttribute("opentype")
		  MoreStr   = ParamNode.getAttribute("morestr")
		  If ParamNode.getAttribute("recommend")="true" Then Rstr="?recommend=1"
		  Set DocNode=XMLSql.DocumentElement.SelectNodes("row")
		  Templates  = "" : N = 0 
		  If PrintType="1" Then
		       echoln "<table border=""0"" width=""100%"" cellpadding==""0"" cellspacing=""0"">"
				  For Each Node In XMLSql.DocumentElement.SelectNodes("row")
					echoln "<tr><td height=""" & RowHeight & """>"
					echoln NaviStr &"<a href=""" & GetSpaceUrl() &"""" &T_CssStr& KS.G_O_T_S(opentype) & ">" & KS.GotTopic(Node.SelectSingleNode("@blogname").text,T_Len) &"</a>"
					echoln "</td></tr>"
					echoln KS.GetSplitPic(SplitPic,1)
				  Next
				  if morestr<>"" then
					 echoln "<tr><td height=""" & RowHeight & """ align=""right""><a href=""" & DomainStr & "space/morespace.asp" & RStr &"""" &T_CssStr& KS.G_O_T_S(OpenType) & ">" & morestr &"</a></td></tr>"
			      end if
				  echoln "</table>"
		  Else
		   Templates=ExplainDiyStyle(LabelStyle,DocNode.length)
		  End If

		 ExplainSpaceListLabelBody= Templates
		End Function
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:GetBlogInfoList
		'作 用: 空间日志列表标签函数
		'参 数: LabelStyle 标签样式
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function GetBlogInfoList(LabelStyle)
             If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				 GetBlogInfoList="标签加载出错" : Exit Function
			 Else
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
		     LabelID   = ParamNode.getAttribute("labelid")
			 ClassID   = ParamNode.getAttribute("classid") : If Not IsNumeric(ClassID) Then ClassID=0
			 AjaxOut   = ParamNode.getAttribute("ajaxout")
			 Num       = ParamNode.getAttribute("num")
			 
			 If LabelID<>"ajax" and Cbool(ParamNode.getAttribute("ajaxout"))=true Then 
			  GetBlogInfoList="<span id=""ks" & LabelID & "_0_0_0_0""></span>":Exit Function
			 End If
			 
			 Dim UserName,TypeID,OrderStr,isbest,ShowFlag
			 UserName  = ParamNode.getAttribute("username")
			 TypeID    = ParamNode.getAttribute("typeid")
			 OrderStr  = ParamNode.getAttribute("orderstr") : If OrderStr="" Then OrderStr=" id desc"
			 isbest    = ParamNode.getAttribute("isbest") : If isbest="" Then isbest=false
			 
			 Param=" Where Status=0"
		     If UserName<>"" Then Param=Param & " And UserName='" & UserName & "'"
		     If TypeID<>"0" Then Param=Param & " And TypeID=" & TypeID
			 If cbool(IsBest)=true Then Param=Param & " And Best=1"
			 If Instr(Lcase(OrderStr),"hits")<>0 Then OrderStr=OrderStr & ",id desc"
			 Param=Param & " order by " & OrderStr
			 Dim SqlStr
			 SqlStr="Select TOP " & num & " ID,UserID,Title,UserName,AddDate,TypeID,Hits From KS_BlogInfo " & Param
			Dim RS:Set RS=Conn.Execute(SqlStr)
			If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			RS.Close:Set RS=Nothing
			If IsObject(XMLSql) Then
			 GetBlogInfoList=ExplainBlogInfoListLabelBody(LabelStyle)
			End If 
			Set Node=Nothing
	   End Function
	
	   Function GetLogUrl()
	    If KS.SSetting(21)="1" Then
		  GetLogUrl = DomainStr & "space/list-" & Node.SelectSingleNode("@userid").text & "-" & Node.SelectSingleNode("@id").text & KS.SSetting(22)
		Else
		  GetLogUrl = DomainStr & "space/?" & Node.SelectSingleNode("@userid").text & "/log/" & Node.SelectSingleNode("@id").text
		End If
	   End Function	
		
	   Function ExplainBlogInfoListLabelBody(LabelStyle)
		  Set DocNode=XMLSql.DocumentElement.SelectNodes("row")
		  Dim PrintType,RowHeight,NaviStr,T_CssStr,OpenType,T_Len,DateRule,DateAlign,SplitPic,MoreStr,Rstr
		  PrintType = ParamNode.getAttribute("printtype")
		  RowHeight=KS.ChkClng(ParamNode.getAttribute("rowheight"))
		  NaviStr   = KS.GetNavi(ParamNode.getAttribute("navtype"), ParamNode.getAttribute("nav"))
		  T_CssStr  = KS.GetCss(ParamNode.getAttribute("titlecss"))
		  OpenType  = ParamNode.getAttribute("opentype")
		  T_Len     = KS.ChkClng(ParamNode.getAttribute("titlelen"))
		  DateRule  = ParamNode.getAttribute("daterule")
		  DateAlign = ParamNode.getAttribute("datealign")
		  SplitPic  = ParamNode.getAttribute("splitpic")
		  MoreStr   = ParamNode.getAttribute("morestr")
		  If lcase(ParamNode.getAttribute("isbest"))="true" Then Rstr="?isbest=1"
		  
		  Templates  = "" : N = 0 
		  If PrintType="1" Then
		         echoln "<table border=""0"" width=""100%"" cellpadding==""0"" cellspacing=""0"">"
				 For Each Node in DocNode
					echoln "<tr><td height=""" & RowHeight & """>"
					echoln NaviStr &"<a href=""" & GetLogUrl &"""" &T_CssStr& KS.G_O_T_S(OpenType) & ">" & KS.GotTopic(Node.SelectSingleNode("@title").text,T_Len) &"</a>"
					echoln GetDateStr(1,Node.SelectSingleNode("@adddate").text,DateRule,DateAlign,"",1,1)& "</tr>"
					echoln KS.GetSplitPic(SplitPic,1)
				  Next
				  if morestr<>"" then
					echoln "<tr><td height=""" & RowHeight & """ align=""right""><a href=""" & DomainStr & "space/morelog.asp" & RStr &"""" &T_CssStr& KS.G_O_T_S(OpenType) & ">" & morestr &"</a></td></tr>"
			      end if
				  echoln "</table>"
		  Else
		   Templates=ExplainDiyStyle(LabelStyle,DocNode.length)
		  End If
		  ExplainBlogInfoListLabelBody=Templates
		End Function
		
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:GetXCList
		'作 用: 相册列表标签函数
		'参 数: LabelStyle 标签样式
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function GetXCList(LabelStyle)
             If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				 GetXCList="标签加载出错" : Exit Function
			 Else
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
		     LabelID   = ParamNode.getAttribute("labelid")
			 ClassID   = ParamNode.getAttribute("classid") : If Not IsNumeric(ClassID) Then ClassID=0
			 AjaxOut   = ParamNode.getAttribute("ajaxout")
			 Num       = ParamNode.getAttribute("num")
			 
			 If LabelID<>"ajax" and Cbool(ParamNode.getAttribute("ajaxout"))=true Then 
			  GetXCList="<span id=""ks" & LabelID & "_0_0_0_0""></span>":Exit Function
			 End If
			 
			 Dim UserName,TypeID,OrderStr,Recommend,Rstr
			 UserName  = ParamNode.getAttribute("username")
			 TypeID    = ParamNode.getAttribute("typeid")
			 Recommend = ParamNode.getAttribute("recommend") : IF KS.IsNul(Recommend) Then Recommend=false
			 OrderStr  = ParamNode.getAttribute("orderstr") : If OrderStr="" Then OrderStr=" id desc"
			 
			 Dim Param:Param=" Where Status=1"
			 If UserName<>"" Then Param=Param & " And UserName='" & UserName & "'"
			 If ClassID<>"0" Then Param=Param & " And ClassID=" & ClassID
			 If Cbool(Recommend)=true Then Param=Param & " and recommend=1"


			 If Instr(Lcase(OrderStr),"hits")<>0 Then OrderStr=OrderStr & ",id desc"
			 Param=Param & " order by " & OrderStr
			 Dim SqlStr:SqlStr="Select TOP " & num & " id,xcname,userid,username,photourl,flag,xps,hits,descript From KS_PhotoXC " & Param
			Dim RS:Set RS=Conn.Execute(SqlStr)
			If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			RS.Close:Set RS=Nothing
			If IsObject(XMLSql) Then
			 GetXCList=ExplainXCListLabelBody(LabelStyle)
			End If 
			Set Node=Nothing
	   End Function
	   Function GetXCUrl()
	      If KS.SSetting(21)="1" Then
	      GetXCUrl=DomainStr & "space/showalbum-" & Node.SelectSingleNode("@userid").text & "-"& Node.SelectSingleNode("@id").text
		  Else
	      GetXCUrl=DomainStr & "space/?" & Node.SelectSingleNode("@userid").text & "/showalbum/"& Node.SelectSingleNode("@id").text
		  End If
	   End Function
	   Function ExplainXCListLabelBody(LabelStyle)
		  Dim PrintType,ShowStyle,Recommend,Rstr,k,i,Col,Width,Height,OpenType,T_Len,morestr
		  PrintType = ParamNode.getAttribute("printtype")
		  ShowStyle = KS.ChkClng(ParamNode.getAttribute("showstyle"))
		  Col       = KS.ChkClng(ParamNode.getAttribute("col"))
		  Width     = ParamNode.getAttribute("xcwidth")
		  Height    = ParamNode.getAttribute("xcheight")
		  OpenType  = ParamNode.getAttribute("opentype")
		  T_Len     = ParamNode.getAttribute("titlelen")
		  morestr   = ParamNode.getAttribute("morestr")
	      Recommend = ParamNode.getAttribute("recommend") : IF KS.IsNul(Recommend) Then Recommend=false
		  If Cbool(Recommend)=true Then RStr="?recommend=1"
		  Set DocNode=XMLSql.DocumentElement.SelectNodes("row")
		  Dim TotalNum:TotalNum=DocNode.length
		  
		  Templates="" : N=0
		  If PrintType="1" Then
		       echoln "<table border=""0"" width=""100%"">"
				  For K=0 To TotalNum-1
					 echoln "<tr>"
					 For I=1 to Col
					   Set Node=DocNode.Item(n)
						echoln "<td width=""" & CInt(100 / CInt(Col)) & "%"">"
					    Select Case ShowStyle
					      Case 1
						   echoln "  <table width=""100%"" cellspacing=""1"" cellpadding=""2"">"
						   echoln "   <tr>"
						   echoln "     <td width=""" &Width & """>"
						   echoln "<a href=""" & GetXCUrl & """" & KS.G_O_T_S(OpenType) & "><img style=""border:3px solid #f1f1f1"" src=""" & Node.SelectSingleNode("@photourl").text & """ border=""0"" width=""" &Width & """ height=""" & Height & """ alt=""" &Node.SelectSingleNode("@xcname").text & """ /></a> </td><td>名 称：" & KS.GotTopic(Node.SelectSingleNode("@xcname").text,T_Len) & "<br/>作 者：" &  Node.SelectSingleNode("@username").text & "<br/>照 片：" &  Node.SelectSingleNode("@xps").text & "<br/>人 气：" & Node.SelectSingleNode("@hits").text & "<br/>状 态：" & GetStatusStr(Node.SelectSingleNode("@flag").text) & "</td></tr>"
						   echoln "   </table>"
						   
					 Case 2
						   echoln "    <table cellSpacing=""0"" cellPadding=""6"" width=""100"" border=""0"">"
						   echoln "        <tr>"
						   echoln "         <td align=""center"" width=""" & Width & """><a href=""" & GetXCUrl & """" & KS.G_O_T_S(OpenType) & "><img style=""border:3px solid #f1f1f1"" src=""" &Node.SelectSingleNode("@photourl").text & """ border=""0"" width=""" &Width & """ height=""" & Height & """ alt=""" & Node.SelectSingleNode("@xcname").text & """/></a><br />"
						   echoln "     <a href=""" & GetXCUrl & """" & KS.G_O_T_S(OpenType) & ">" &KS.GotTopic(Node.SelectSingleNode("@xcname").text,T_Len) &"</a></td>"
						   echoln "      </tr>"
						   echoln "    </table>"
					 End Select
						   echoln "</td>"
						N = N+1 : If N>=TotalNum Then Exit For
					Next
					 echoln "</tr>"
					If N>=TotalNum Then Exit For
				 Next
				    if morestr<>"" then
					 echoln "<tr><td height=""20"" colspan=""" & col & """ align=""right""><a href=""" & DomainStr & "space/morephoto.asp" & RStr &"""" &KS.G_O_T_S(OpenType) & ">" & morestr &"</a></td></tr>" & vbcrlf
			        end if
				   echoln "</table>"
			Else
			  Templates=ExplainDiyStyle(LabelStyle,TotalNum)
			End If
			ExplainXCListLabelBody=Templates
		End Function
		
		Function GetStatusStr(val)
           Select Case Val
		    Case 1:GetStatusStr="公开"
			Case 2:GetStatusStr="会员"
			Case 3:GetStatusStr="密码"
			Case 4:GetStatusStr="隐私"
		   End Select
			GetStatusStr="<span style=""color:red"">" & GetStatusStr & "</span>"
		End Function
		
		
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:GetGroupList
		'作 用: 相册列表标签函数
		'参 数: LabelStyle 标签样式
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function GetGroupList(LabelStyle)
             If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				 GetGroupList="标签加载出错" : Exit Function
			 Else
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
		     LabelID   = ParamNode.getAttribute("labelid")
			 ClassID   = ParamNode.getAttribute("classid") : If Not IsNumeric(ClassID) Then ClassID=0
			 AjaxOut   = ParamNode.getAttribute("ajaxout")
			 Num       = ParamNode.getAttribute("num")
			 
			 If LabelID<>"ajax" and Cbool(ParamNode.getAttribute("ajaxout"))=true Then 
			  GetGroupList="<span id=""ks" & LabelID & "_0_0_0_0""></span>":Exit Function
			 End If
			 
			 Dim UserName,TypeID,OrderStr,Recommend,Rstr
			 UserName  = ParamNode.getAttribute("username")
			 TypeID    = ParamNode.getAttribute("typeid")
			 Recommend = ParamNode.getAttribute("recommend") : IF KS.IsNul(Recommend) Then Recommend=false
			 
			 Dim Param:Param=" Where verific=1"
			 If UserName<>"" Then Param=Param & " And UserName='" & UserName & "'"
			 If ClassID<>"0" Then Param=Param & " And ClassID=" & ClassID
			 If Cbool(Recommend)=true Then Param=Param & " and recommend=1"


			 Dim SqlStr:SqlStr="Select top " & num &" (select count(id) from ks_teamtopic where teamid=a.id and parentid=0) as teamtopicnum,(select count(id) from ks_teamtopic where teamid=a.id) as teamreplynum,(select count(id) from ks_teamusers where status=3 and teamid=a.id) as teamusernum,id,teamname,username,photourl,adddate,point From KS_Team a " & Param & " Order By ID Desc " 
			Dim RS:Set RS=Conn.Execute(SqlStr)
			If Not RS.Eof Then Set XMLSql=KS.RsToXml(RS,"row","root") Else XMLSql=Empty
			RS.Close:Set RS=Nothing
			If IsObject(XMLSql) Then
			 GetGroupList=ExplainGroupListLabelBody(LabelStyle)
			End If 
			Set Node=Nothing
	   End Function
	   
	   Function GetGroupUrl()
	     GetGroupUrl=DomainStr & "space/group.asp?id=" & Node.SelectSingleNode("@id").text
	   End Function
	   
	   Function ExplainGroupListLabelBody(LabelStyle)
		  Dim PrintType,ShowStyle,Recommend,Rstr,k,i,Col,Width,Height,OpenType,T_Len,morestr
		  PrintType = ParamNode.getAttribute("printtype")
		  ShowStyle = KS.ChkClng(ParamNode.getAttribute("showstyle"))
		  Col       = KS.ChkClng(ParamNode.getAttribute("col"))
		  Width     = ParamNode.getAttribute("width")
		  Height    = ParamNode.getAttribute("height")
		  OpenType  = ParamNode.getAttribute("opentype")
		  T_Len     = ParamNode.getAttribute("titlelen")
		  morestr   = ParamNode.getAttribute("morestr")
	      Recommend = ParamNode.getAttribute("recommend") : IF KS.IsNul(Recommend) Then Recommend=false
		  If Cbool(Recommend)=true Then RStr="?recommend=1"
		  Set DocNode=XMLSql.DocumentElement.SelectNodes("row")
		  Dim TotalNum:TotalNum=DocNode.length
		  
		  Templates="" : N=0
		  If PrintType="1" Then
	             echoln "<table border=""0"" width=""100%"">"
				  For K=0 To TotalNum-1
					 echoln "<tr>"
					 For I=1 to Col
					    Set Node=DocNode.Item(n)
						echoln "<td width=""" & CInt(100 / CInt(Col)) & "%"">"
					  Select Case ShowStyle
					    Case 1
						 echoln "<table class=""border"" cellSpacing=""0"" cellPadding=""0"" style=""margin:3px"" width=""99%"" border=0>"
						 echoln "  <tr>"
						 echoln "	 <td width=""30%"" align=""center""><a href=""" & GetGroupUrl & """ title=""" & Node.SelectSingleNode("@teamname").text & """" & KS.G_O_T_S(OpenType) & "><img style=""border:1px solid #ccc"" src=""" & Node.SelectSingleNode("@photourl").text & """ width=""80"" height=""70"" border=""0""></a></td>"
						 echoln "	 </td>"
						 echoln "	 <td width=""70%""><a class=""teamname"" href=""" & GetGroupUrl & """ title=""" & Node.SelectSingleNode("@teamname").text & """" & KS.G_O_T_S(OpenType) & "> " & Node.SelectSingleNode("@teamname").text & "</a><br><font color=""#a7a7a7"">创建者：" & Node.SelectSingleNode("@username").text & "</font><br><font color=""#a7a7a7"">创建时间:" &Node.SelectSingleNode("@adddate").text & "</font><br>主题/回复：" & Node.SelectSingleNode("@teamtopicnum").text & "/" & Node.SelectSingleNode("@teamreplynum").text & "&nbsp;&nbsp;&nbsp;成员:" & Node.SelectSingleNode("@teamusernum").text & "人                             </td>"
						 echoln "	   </tr>"
						 echoln "	</table>"
					 Case 2
						 echoln "    <table cellSpacing=""0"" cellPadding=""6"" width=""100"" border=""0"">"
						 echoln "        <tr>"
						 echoln "         <td align=""center"" width=""" & Width & """ bgColor=#ffffff height=" & Height & ">"
						 echoln "         <a href=""" & GetGroupUrl & """" & KS.G_O_T_S(OpenType) & "><img src=""" &Node.SelectSingleNode("@photourl").text& """ border=""0"" style=""border:2px solid #f1f1f1"" width=""" &Width & """ height=""" & Height & """></a>"
						 echoln "       <br/><a href=""" & GetGroupUrl & """" & KS.G_O_T_S(OpenType) & ">" &KS.GotTopic(Node.SelectSingleNode("@teamname").text,T_Len) &"</a></td>"
						 echoln "      </tr>"
						 echoln "    </table>"
					 End Select
						echoln "</td>"
			
						N = N+1 : If N>=TotalNum Then Exit For
					
					Next
					 echoln "</tr>"
				  
					If N>=TotalNum Then Exit For
				 Next
				 if morestr<>"" then
					echoln "<tr><td height=""20"" colspan=""" & col & """ align=""right""><a href=""" & DomainStr & "space/moregroup.asp" & RStr &"""" &KS.G_O_T_S(OpenType) & ">" & morestr &"</a></td></tr>" & vbcrlf
			      end if
				   echoln "</table>"
		 Else
			  Templates=ExplainDiyStyle(LabelStyle,TotalNum)
		 End If
		 ExplainGroupListLabelBody=Templates
	   End Function
		
		
		

		Function C_L_C(LabelID,FieldID)
		  on error resume next
		  If not IsObject(Application(KS.SiteSN&"_cirlabellist")) Then
			 Application.Lock
			 Dim RS:Set Rs=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "select ID,Description,LabelContent From KS_Label Where LabelType=6 Order by adddate",conn,1,1
			 Set Application(KS.SiteSN&"_cirlabellist")=KS.RecordsetToxml(rs,"cirlabel","cirlabellist")
			 RS.Close:Set Rs=Nothing
			 Application.unLock
		  End If
	       C_L_C=Application(KS.SiteSN&"_cirlabellist").documentElement.selectSingleNode("cirlabel[@ks0='" & LabelID & "']/@ks" & FieldID & "").text
	      if err then C_L_C="":err.Clear
		End Function
		
		
		Public Function TransformXSLTemplate(iXMLDom,strXslt)
			Dim proc,XMLStyle,node,cnode,XSLTemplate
			If strXslt = "" Then TransformXSLTemplate="" : Exit Function
			Set XSLTemplate=KS.InitialObject("Msxml2.XSLTemplate" & MsxmlVersion )
			Set XMLStyle=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion )
			If XMLStyle.loadxml(strXslt) Then
				XSLTemplate.stylesheet=XMLStyle
				Set proc = XSLTemplate.createProcessor()
				proc.input = iXMLDom
				proc.transform()
				Dim procstr
				procstr = proc.output
				Set proc=Nothing
				TransformXSLTemplate = procstr
			Else
				TransformXSLTemplate = "标签语法错误，检查是否符合XSLT标准"
			End If
			Set XMLStyle=Nothing
			Set XSLTemplate=Nothing
		End Function
		'替换内外sql关系值
		Function GetRelations(ClassNode,Content)
		     Dim regEx, Matches, Match,TempStr,QStr,ReqType
			 Set regEx = New RegExp
			 regEx.Pattern= "{(r)[^{}]*}"
			 regEx.IgnoreCase = True
			 regEx.Global = True
			 Set Matches = regEx.Execute(Content)
			 For Each Match In Matches
				Content=Replace(Content,Match.Value,ClassNode.selectSingleNode("@" & replace(replace(Lcase(Match.Value),"{r:",""),"}","")).text)
			Next
			GetRelations=Content
		End Function
 
        Function CutText(Content)
		     Dim regEx, Matches, Match,TempStr,QStr,TLen,CutLen,Text,TextArr
			 Set regEx = New RegExp
			 regEx.Pattern= "{(KS:CutText\()[^{}]*}"
			 regEx.IgnoreCase = True
			 regEx.Global = True
			 Set Matches = regEx.Execute(Content)
			 For Each Match In Matches
			    Tempstr=Replace(Replace(Match.Value,"{KS:CutText(",""),")}","")
				TextArr=Split(TempStr,",")
				TLen=Ubound(TextArr)
				CutLen=KS.ChkClng(TextArr(Tlen-1))
				Text=Replace(Replace(Tempstr,"," & TextArr(Tlen),""),"," & TextArr(Tlen-1),"")
				If Len(text)>CutLen Then
				Text=KS.GotTopic(Text,CutLen) & Replace(Replace(TextArr(Tlen),"""",""),"'","")
				End If
				Content=Replace(Content,Match.Value,text)
			Next
			CutText=Content
		End Function
		
		Function GetCirList(LabelStyle)
		     If Not XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
				 GetCirList="标签加载出错" : Exit Function
			 Else
				 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
			 End If
		     LabelID   = ParamNode.getAttribute("labelid")
			 Dim ChannelID,DateRule
			 ChannelID = ParamNode.getAttribute("channelid")
			 DateRule  = ParamNode.getAttribute("daterule")
		
			 Dim LBParam:LBParam=C_L_C(LabelID,1)
			 If LBParam="" Then Exit Function
			  LBParam=Replace(LBParam,"{$CurrClassID}",FCls.RefreshFolderID)
			  LBParam=Replace(LBParam,"{$CurrChannelID}",FCls.ChannelID)
			  If Instr(LBParam,"{$CurrClassChildID}")<>0 Then
			   LBParam=Replace(LBParam,"{$CurrClassChildID}",KS.GetFolderTid(FCls.RefreshFolderID))
			  End If
			   LBParam=Replace(LBParam,"{$CurrInfoID}",FCls.RefreshInfoID)

			 LBParam=Split(LBParam,"@@@")
			 Dim OutSql,InnerSQL,iXMLDom,ClassList,ClassNode,RS,DataList,Node
			 OutSQL=LBParam(0)
			 InnerSQL=LBParam(1)
			  Set iXMLDom = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			  iXMLDom.appendChild(iXMLDom.createElement("xml"))
			  'iXMLDom.documentElement.setAttribute "showdate",1
			  Dim RSC:Set RSC=Conn.Execute(OutSQL)
			  If RSC.Eof And RSC.Bof Then RSC.Close:Set RSC=Nothing:Exit Function
			  Set ClassList=KS.RsToxml(rsc,"outerrow","outerlist")
			  RSC.Close:Set RSC=Nothing
			 If Not ClassList Is Nothing Then 
				For Each ClassNode in classlist.documentElement.SelectNodes("outerrow")
				  ClassNode.setAttribute "classlink",KS.GetFolderPath(ClassNode.selectSingleNode("@id").text)
				  
				  Dim ISQL:ISQL=InnerSQL
				  ISQL=GetRelations(ClassNode,ISQL)
				  Set RS=Conn.Execute(ISQL)
				  Set datalist=KS.RsToxml(rs,"innerrow","innerlist")
				  Classnode.appendChild(datalist.documentElement.cloneNode(True))
	           Next

			 For Each Node in classlist.documentElement.SelectNodes("outerrow/innerlist/innerrow")
			   If ChannelID<>0 and instr(LBParam(2),"@linkurl")<>0 Then
				   Dim SqlCls:Set SqlCls=New DIYCls
				   Node.setAttribute "linkurl",SqlCls.Get_InfoUrl_Field("id",Node.selectSingleNode("@id").text,ChannelID,1)
				   Set SqlCls=Nothing
			   End If
			   If Instr(lcase(InnerSQL),"ks_class")<>0 Then  
				   Node.setAttribute "classlink",KS.GetFolderPath(Node.selectSingleNode("@id").text)
			   End If
			   If Instr(Lcase(InnerSQL),"adddate")<>0 Then
			   Node.selectSingleNode("@adddate").text=LFCls.Get_Date_Field(Node.selectSingleNode("@adddate").text,DateRule)
			   End If
			 Next
	         iXMLDom.documentElement.appendChild(ClassList.documentElement.cloneNode(True))
			 
			End If 
	        GetCirList= CutText(TransformXSLTemplate(iXMLDom,LBParam(2)))
	
		End Function
		

		
		
				
		
		
		
'============================================================================================================================
'                                                         以下为相关刷新通用函数
'============================================================================================================================
		
		'功 能:取得信息标题
		Function GetItemTitle(Byval Title, T_Len, PicTF, TitleType, TitleFontColor, TitleFontType)
			Dim DecoratesTitle
			If IsNumeric(T_Len) Then
			  Title = KS.GotTopic(Title, T_Len)
			End If
			If CBool(PicTF) = True Then
			
			 Dim TitleTypeXml:Set TitleTypeXml=LFCls.GetXMLFromFile("TitleType")
			 If IsObject(TitleTypeXml) Then
   			     On Error Resume Next
				 Dim Icon:Icon=TitleTypeXml.documentElement.selectSingleNode("//TitleTypeRule/Field[@Name='" & TitleType & "']/@Icon").Text
				 If Err or KS.IsNul(Icon) Then 
				  err.clear
				  Dim Color:Color=TitleTypeXml.documentElement.selectSingleNode("//TitleTypeRule/Field[@Name='" & TitleType & "']/@Color").Text
				  if err then
				  DecoratesTitle=TitleType
				  Err.Clear
				  else
				   DecoratesTitle = "<span style=""color:" & Color & """>" & TitleType & "</span>"
				  end if
				 Else
				  DecoratesTitle = "<img src=""" & icon & """ border=""0"" alt=""" & TitleType  &"""/>" 
				 End If
			 End If
			 Set TitleTypeXml=Nothing
		   End If
		  If TitleFontColor <> "" Then
				DecoratesTitle = DecoratesTitle & "<span style=""color:" & TitleFontColor & """>" & Title & "</span>"
		  Else
				DecoratesTitle = DecoratesTitle & Title
		  End If
		  If TitleFontType <> "" Then
				 Select Case (TitleFontType)
				  Case 1:DecoratesTitle = "<strong>" & DecoratesTitle & "</strong>"
				  Case 2:DecoratesTitle = "<I>" & DecoratesTitle & "</I>"
				  Case 3:DecoratesTitle = "<strong><I>" & DecoratesTitle & "</I></strong>"
				  Case Else
					DecoratesTitle = DecoratesTitle
				 End Select
		  End If
		  GetItemTitle = DecoratesTitle
		End Function
			
		
		'以下按样式刷新自由JS函数
		Function RefreshCss(JSID, WordCss, Col, OpenType, num, R_H, T_Len, C_Len, NavType, Nav, MoreType, MoreLink, SplitPic, DateRule, DateAlign, T_Css, DateCss, ContentCss, BGCss)
			   If JSID = "" Then
				RefreshCss = ""
				Exit Function
			   End If
			   Dim SqlStr,RS,ChannelID
			   ChannelID=1
              Set RS=Server.CreateObject("ADODB.RECORDSET")
			   If num = 0 Then
				   SqlStr = "Select * From " & KS.C_S(ChannelID,2) &" Where JSID like '%" & JSID & "%' AND Verific=1 AND DelTF=0 Order BY  IsTop Desc,ID Desc "
			   Else
				   SqlStr = "Select TOP " & num & " * From " & KS.C_S(ChannelID,2) &" Where JSID like '%" & JSID & "%' AND Verific=1 AND DelTF=0 Order BY  IsTop Desc,ID Desc "
			   End If
			   RS.Open SqlStr, Conn, 1, 1
			   If Not RS.EOF Then
				  Dim TempStr, TempTitle, NaviStr,ArticleContent, I, ColSpanNum
				  TempStr = "<table " & KS.GetCss(BGCss) & " border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" align=""center"">"
				 Do While Not RS.EOF
				   TempStr = TempStr & "<tr>"
				   For I = 1 To Col
					  TempTitle = GetItemTitle(RS("Title"), T_Len, False, RS("TitleType"), RS("TitleFontColor"), RS("TitleFontType"))
					 
					  TempTitle = "<a " & KS.GetCss(T_Css) & " href=""" & KS.GetItemUrl(channelid,rs("tid"),rs("id"),rs("fname")) & """" & KS.G_O_T_S(OpenType) & " title=""" & RS("Title") & """>" & TempTitle & "</a>"
					  R_H = KS.G_R_H(R_H)
					  NaviStr = KS.GetNavi(NavType, Nav)
					  TempStr = TempStr & "<td width=""" & CInt(100 / CInt(Col)) & "%"">"
					  If RS("Intro")="" Then ArticleContent=RS("ArticleContent") Else ArticleContent=RS("Intro")
					  
					 Select Case WordCss
						Case "A"
							TempStr = TempStr & "<table width=""100%"" height=""" & R_H & """ cellpadding=""0"" cellspacing=""0"" border=""0"">"
							TempStr = TempStr & "<tr><td> " & NaviStr & TempTitle & "</td>"
							If DateRule <> "0" And DateRule <> "" Then
							   TempStr = TempStr & "<td width=""20%"" nowrap align=" & DateAlign & "><span " & KS.GetCss(DateCss) & ">" & LFCls.Get_Date_Field(RS("AddDate"), DateRule) & "</span></td></tr>"
							   ColSpanNum = 2
						   Else
							   TempStr = TempStr & "</tr>"
							   ColSpanNum = 1
						   End If
						   If SplitPic <> "" Then
						   TempStr = TempStr & KS.GetSplitPic(SplitPic, ColSpanNum)
						   End If
						   TempStr = TempStr & "</table>"
					   Case "B"
						   TempStr = TempStr & "<table width=""100%"" height=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
							If DateRule <> "0" And DateRule <> "" Then
							   TempStr = TempStr & "<tr><td height=""" & R_H & """> " & NaviStr & TempTitle & "&nbsp;&nbsp;<span align=" & DateAlign & KS.GetCss(DateCss) & ">" & LFCls.Get_Date_Field(RS("AddDate"), DateRule) & "</span></td></tr>"
							   ColSpanNum = 2
						   Else
							   TempStr = TempStr & "<tr><td height=""" & R_H & """> " & NaviStr & TempTitle & "</td></tr>"
							   ColSpanNum = 1
						   End If
						   TempStr = TempStr & "<tr><td><table border=0 align=center width=""100%""><tr><td><span " & KS.GetCss(ContentCss) & ">&nbsp;&nbsp;&nbsp;&nbsp;" & KS.GotTopic(Replace(Replace(Replace(KS.LoseHtml(ArticleContent), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), C_Len) & "……</span></td></tr>" & KS.GetMoreLink(1,1, R_H, MoreType, MoreLink, KS.GetItemUrl(channelid,rs("tid"),rs("id"),rs("fname")), KS.G_O_T_S(OpenType)) & "</table></td></tr>"
						   If SplitPic <> "" Then
						   TempStr = TempStr & KS.GetSplitPic(SplitPic, ColSpanNum)
						   End If
						   TempStr = TempStr & "</table>"
					 Case "C"
						   TempStr = TempStr & "<table width=""100%"" height=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
						   TempStr = TempStr & "<tr><td width=""100%""><table border=0 align=center width=""100%""><tr><td><span " & KS.GetCss(ContentCss) & ">&nbsp;&nbsp;&nbsp;&nbsp;" & KS.GotTopic(Replace(Replace(Replace(KS.LoseHtml(ArticleContent), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), C_Len) & "……</span></td></tr>" & KS.GetMoreLink(1,1, R_H, MoreType, MoreLink, KS.GetItemUrl(channelid,rs("tid"),rs("id"),rs("fname")), KS.G_O_T_S(OpenType)) & "</table></td></tr>"
						   If DateRule <> "0" And DateRule <> "" Then
							   TempStr = TempStr & "<tr><td width=""100%"" height=""" & R_H & """> " & NaviStr & TempTitle & "&nbsp;&nbsp;<span align=" & DateAlign & KS.GetCss(DateCss) & ">" & LFCls.Get_Date_Field(RS("AddDate"), DateRule) & "</span></td></tr>"
						   Else
							   TempStr = TempStr & "<tr><td width=""100%"" height=""" & R_H & """> " & NaviStr & TempTitle & "</td></tr>"
						   End If
						   TempStr = TempStr & KS.GetSplitPic(SplitPic, 1)
						   TempStr = TempStr & "</table>"
					 Case "D"
							TempStr = TempStr & "<table width=""100%"" height=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
							TempStr = TempStr & "<tr><td> " & NaviStr & "<br><span " & KS.GetCss(T_Css) & ">" & KS.ListTitle1(Trim(RS("Title")), T_Len) & "</span></td>"
						   TempStr = TempStr & "<td><table width=""100%"" height=""100%""><tr><td><span " & KS.GetCss(ContentCss) & ">&nbsp;&nbsp;&nbsp;&nbsp;" & KS.GotTopic(Replace(Replace(Replace(KS.LoseHtml(ArticleContent), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), C_Len) & "……</span></tr>" & KS.GetMoreLink(1,1, R_H, MoreType, MoreLink, KS.GetItemUrl(channelid,rs("tid"),rs("id"),rs("fname")), KS.G_O_T_S(OpenType)) & "</table></td></tr>"
						   TempStr = TempStr & KS.GetSplitPic(SplitPic, ColSpanNum)
						   TempStr = TempStr & "</table>"
					 Case "E"
						   TempStr = TempStr & "<table width=""100%"" height=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
						   TempStr = TempStr & "<tr><td><table width=""100%"" height=""100%""><tr><td><span " & KS.GetCss(ContentCss) & ">&nbsp;&nbsp;&nbsp;&nbsp;" & KS.GotTopic(Replace(Replace(Replace(KS.LoseHtml(ArticleContent), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), C_Len) & "……</span></tr>" & KS.GetMoreLink(1,1, R_H, MoreType, MoreLink, KS.GetItemUrl(channelid,rs("tid"),rs("id"),rs("fname")), KS.G_O_T_S(OpenType)) & "</table></td>"
						   TempStr = TempStr & "<td> " & NaviStr & "<br><span" & KS.GetCss(T_Css) & " >" & KS.ListTitle1(Trim(RS("Title")), T_Len) & "</span></td></tr>"
						   TempStr = TempStr & KS.GetSplitPic(SplitPic, ColSpanNum)
						   TempStr = TempStr & "</table>"
					End Select
					  TempStr = TempStr & "</td>"
					  RS.MoveNext
					  If RS.EOF Then Exit For
				  Next
				 TempStr = TempStr & "</tr>"
				 Loop
				 TempStr = TempStr & "</table>"
				 RefreshCss = TempStr
			   Else
			   RefreshCss = "":RS.Close:Set RS = Nothing
			   End If
		End Function
End Class
%> 