<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../KS_Cls/Template.asp"-->
<%


Dim KSCls
Set KSCls = New SearchCls
KSCls.Kesion()
Set KSCls = Nothing

Const FuzzySearch = 1  '设为1支持模糊查找，但会加大系统资源的开销，如比如搜索“xp 2003”，包含xp和2003两者的、只包含其中一个的，都能搜索出来。
Const RefreshTime = 0  '设置防刷新时间,不限制设置为0

Class SearchCls
        Private OrderOptionList,OrderArr,OrderListStr,TopMenu,TopMenuArr
        Private KS,ChannelID,ModelTable,Param,XML,Node,StartTime,leavetime,I,FieldMenu,QueryParam
		Private CurrPage,MaxPerPage,TotalPut,PageNum,Key,KeyWordArr,SqlStr,OrderStr,currclassid,ClassXML
		Private FieldXML,FieldNode,TemplateFile,IsRewrite,ModelClassID,SEOTitle,AreaXML,Province,City,BrandXML,BrandNodes,HasBrand
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		

		Public Sub Kesion()
		
		If RefreshTime>0 Then
			If DateDiff("s", Session("SearchTime"), Now()) < RefreshTime Then
				Response.Write "<META http-equiv=Content-Type content=text/html; chaRset=utf-8><meta HTTP-EQUIV=REFRESH CONTENT="&RefreshTime&"><br>本页面起用了防刷新机制，请不要在"&RefreshTime&"秒内连续刷新本页面<BR>正在打开页面，请稍后……"
				Response.End
			End If
		End If
		Session("SearchTime")=Now()
		
		
		  ChannelID=KS.ChkClng(GetParam("c"))  '模型ID
		  If ChannelID=0 Then ChannelID=KS.ChkClng(Request("ChannelID"))
		  If ChannelID=0 Then ChannelID=1
		
		 Dim Template,KSR
		 set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		 FieldXML.async = false
		 FieldXML.setProperty "ServerHTTPRequest", true 
		 FieldXML.load(Server.MapPath(KS.Setting(3)& "config/filtersearch/s" & ChannelID & ".xml"))
		 if FieldXML.parseError.errorCode<>0 Then
			KS.Die "对不起，该频道没有开启筛选！"
		 End If
		    if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set FieldNode=FieldXML.DocumentElement.SelectSingleNode("template")
				 If Not FieldNode Is Nothing Then
				  TemplateFile=FieldNode.Text
				 End If
				 Set FieldNode=FieldXML.DocumentElement.SelectSingleNode("isrewrite")
				 If Not FieldNode Is Nothing Then
				  isrewrite=FieldNode.Text
				 End If
				 Set FieldNode=FieldXML.DocumentElement.SelectSingleNode("maxperpage")
				 If Not FieldNode Is Nothing Then
				  MaxPerPage=FieldNode.Text
				 End If
			Else
				KS.Die "对不起，读取频道筛选配置参数出错！"
		    end if
			
		   If KS.IsNul(isrewrite) Then isrewrite=0
		   If KS.ChkClng(MaxPerPage)=0 Then MaxPerPage=20
		   FCls.RefreshType = "searchIndex"   
		   Set KSR = New Refresh
		   Template = KSR.LoadTemplate(TemplateFile)
		   Template = KSR.KSLabelReplaceAll(Template)
		   Set KSR = Nothing
		   StartTime = Timer()
		   InitialSearch
		   Scan Template
	   End Sub
	   Sub ParseArea(sTokenName, sTemplate)
			Select Case sTokenName
				Case "loop"
				      If IsObject(XML) Then
					   For Each Node In Xml.DocumentElement.SelectNodes("row")
						Scan sTemplate
					   Next
					  Else
					   echo "<div class='border' style='text-align:center'>对不起,根据您的查找条件,找不到任何相关记录!</div>"
					  End If
			End Select 
        End Sub 
		Sub ParseNode(sTokenType, sTokenName)
			Select Case lcase(sTokenType)
			    case "item" EchoItem sTokenName
				case "search" 
				          select case sTokenName
						    case "fieldmenu"  echo FieldMenu
							case "orderoptionlist" echo OrderOptionList
							case "orderliststr" echo OrderListStr
							case "topmenu" echo TopMenu
							case "sqlstr" echo sqlstr
							case "queryparam" 
							 echo Replace(GetFieldLink("","","key,p"),".html","")
							case "key" if key<>"0" and key<>"" then echo key
						    case "showpage" 
							If KS.S("Key")<>"" or KS.S("t")<>"" Then  '搜索表单来的
							 echo KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
							Else
							 echo ShowPage()
							End If
							case "totalput" echo KS.ChkClng(TotalPut)
							case "leavetime" 
							  leavetime=FormatNumber((timer-starttime),5)
							  if leavetime<1 then echo "0"
							  echo FormatNumber((timer-starttime),5)
							case "showkey" if key<>"0" and key<>"" then echo "关键字：<span class=""key"">" & KS.R(key) &"</span>"
							case "seotitle" echo SEOTitle
							case "channelid" echo channelid
						  end select
			End Select
		End Sub
		Sub EchoItem(sTokenName)
		  Select Case sTokenName
		    case "id" echo GetNodeText("id")
			case "linkurl"    
			 If ChannelID=9 Then 
			    if GetNodeText("dtfs")="1" then
					  echo Replace(KS.Setting(3) & KS.C_S(9,8),"//","/") & "sj/" & GetNodeText("id") & ".htm"
				else
					  echo KS.Setting(2) & KS.Setting(3) & "mnkc/exam/?id=" & GetNodeText("id")
				end if
			 Else 
			  echo KS.GetItemURL(ChannelID,GetNodeText("tid"),GetNodeText("id"),GetNodeText("fname"))
			 End If
			case "classname"  echo KS.C_C(GetNodeText("tid"),1)
			case "classurl"   echo KS.GetFolderPath(GetNodeText("tid"))
			case "photourl"   if KS.IsNul(GetNodeText("photourl")) Then echo "/images/nopic.gif" else echo GetNodeText("photourl")
			case "typename" if ks.c_s(channelid,6)=8 then echo KS.GetGQTypeName(GetNodeText("typeid"))
			case "intro" 
			 Dim Intro:intro=KS.Gottopic(KS.LoseHtml(GetNodeText("intro")),160)
			 Intro=Replace(Intro,"&nbsp;","")
			 If Not KS.IsNul(Key) Then
			  echo ReplaceKeyWordRed(Intro)
			 Else
			 echo intro
			 End If
			case else
			  echo GetNodeText(sTokenName)
		  End Select
		End Sub
		
		Function ReplaceKeyWordRed(Content)
		   Dim I
		   For I=0 To Ubound(KeyWordArr)
			Content=Replace(Content,KeyWordArr(i),"<span style='color:red'>" &KeyWordArr(i) & "</span>")
		   Next
		   ReplaceKeyWordRed=Content
		End Function
  
		Function GetNodeText(NodeName)
		 Dim N,Str
		 NodeName=Lcase(NodeName)
		 If IsObject(Node) Then
		  set N=node.SelectSingleNode("@" & NodeName)
		  If Not N is Nothing Then Str=N.text
		  If Not KS.IsNul(Key)  And NodeName="title" Then
		   Str=ReplaceKeyWordRed(Str)
		  End If
		  GetNodeText=Str
		 End If
		End Function
		
		'从xml中加载模型字段
		Sub LoadModelField(ChannelID,ByRef FXML)
			set FXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			FXML.async = false
			FXML.setProperty "ServerHTTPRequest", true 
			FXML.load(Server.MapPath(KS.Setting(3)&"Config/fielditem/field_" & ChannelID&".xml"))
		End Sub

		Sub InitialSearch()
		  Dim FieldStr,TopStr,TopNum,QArr
		  
		  key=ks.URLDecode(GetParam("key")) '需要解码，否则firefox乱码
		  If Key="" Then Key=KS.CheckXSS(KS.R(KS.S("Key")))
		  'If Key<>"" Then Key=UnEscape(Key) 
		  CurrPage=KS.ChkClng(GetParam("p"))
		  If CurrPage<=0 Then CurrPage=KS.ChkClng(Request("page"))
		  If CurrPage<=0 Then CurrPage=1
		  
		  if channelid=9 then
		  	   set ClassXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			   ClassXML.async = false
			   ClassXML.setProperty "ServerHTTPRequest", true 
			   ClassXML.load(Server.MapPath(KS.Setting(3)& "config/sjclass.xml"))
		  end if
		  
		  
		  If KS.S("tid")<>"" Then
		   'ModelClassID=KS.ChkClng(KS.C_C(KS.S("tid"),9))
		   dim nn,tts
		   If ChannelID=9 Then  '考试分类
			   tts=ClassXML.DocumentElement.SelectSingleNode("item[@id=" & KS.S("tid") &"]/@ts").text
			   if tts<>"" then
				tts=split(tts,",")
				  for nn=0 to ubound(tts)-1
					if ModelClassID="" Then
						ModelClassID=tts(nn)
					else
						ModelClassID=ModelClassID & ":" & tts(nn)
					end if
				 next
			   end if
		   Else
			   tts=KS.C_C(KS.S("tid"),8)
			   if tts<>"" then
				tts=split(tts,",")
				  for nn=0 to ubound(tts)-1
					if ModelClassID="" Then
						ModelClassID=KS.C_C(tts(nn),9)
					else
						ModelClassID=ModelClassID & ":" & KS.C_C(tts(nn),9)
					end if
				 next
			   end if
			End If
		  End If
		  If ModelClassID="" Then ModelClassID=GetParam("tid")

		  Dim FXML
		  Call LoadModelField(ChannelID,Fxml)
		  Set Application(KS.SiteSN&"_field")=FXMl
		   
		   
		  '排序select选项
		  Dim OrderField,OrderNode,ssel
		  SET FieldNode=FieldXML.DocumentElement.SelectNodes("orderitem[@enabled='true']")
		  I=1
		  OrderOptionList=OrderOptionList & "<option value='" & GetFieldLink("o",0,"") & "'>默认排序</option>" &vbcrlf
		  For Each OrderNode in FieldNode
		    If KS.ChkClng(GetParam("o"))=I Then 
			 ssel=" selected"
			 OrderField=" Order By " & OrderNode.SelectSingleNode("@name").text &" asc"
			 if OrderNode.SelectSingleNode("@name").text<>"id" then OrderField=OrderField &",ID"
			Else 
			 ssel=""
			End If
		    OrderOptionList=OrderOptionList & "<option value='" & GetFieldLink("o",I,"") & "'" & ssel &">" & OrderNode.SelectSingleNode("uptitle").text & "</option>" &vbcrlf
			I=I+1
		    If KS.ChkClng(GetParam("o"))=I Then 
			 ssel=" selected" 
			 OrderField=" Order By " & OrderNode.SelectSingleNode("@name").text &" desc"
			 if OrderNode.SelectSingleNode("@name").text<>"id" then OrderField=OrderField &",ID"
			Else 
			ssel=""
			End If
		    OrderOptionList=OrderOptionList & "<option value='" & GetFieldLink("o",I,"") & "'" & ssel &">" & OrderNode.SelectSingleNode("downtitle").text & "</option>" &vbcrlf
			I=I+1
		  Next
		  
		  '搜索结果上面的全部,推荐,24小时信息等
		  Dim TopMenuLinkParam,TopMenuValue,OONode,TopParam
		  SET FieldNode=FieldXML.DocumentElement.SelectNodes("optionitem")
		  If FieldNode.Length>0 Then
		    I=0
		    For Each OONode In FieldNode
			   If KS.ChkClng(GetParam("x"))=I Then 
			    TopMenu=TopMenu & "<LI class=slt><span><a href='" & GetFieldLink("x",I,"") & "'>" &OONode.SelectSingleNode("title").text & "</a></span></li>"
				TopParam=OONode.SelectSingleNode("sqlparam").text
			   Else
	            TopMenu=TopMenu & "<LI><span><a href='" & GetFieldLink("x",I,"") & "'>" & OONode.SelectSingleNode("title").text & "</a></span></li>"
			   End If
			  I=I+1
			Next
		  End If

		  Dim FieldArr,K,OptionArr,CurrNode
		  SET FieldNode=FieldXML.DocumentElement.SelectNodes("item[@enabled='true']")
		  SEOTitle=""
		  For I=0 To FieldNode.length-1
		    Set CurrNode=FieldNode.Item(i)
			 
			 HasBrand=false
			 If lcase(CurrNode.SelectSingleNode("@name").text)="brandid" Then  '品牌根据栏目ID关联，特殊处理
			  if modelclassid<>"" then
					 set BrandXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
					 BrandXML.async = false
					 BrandXML.setProperty "ServerHTTPRequest", true 
					 BrandXML.load(Server.MapPath(KS.Setting(3)& "config/shopbrand.xml"))
					 if BrandXML.parseError.errorCode<>0 Then
						Call KS.CreateBrandCache()
					 End If
					 dim brandclassid:brandclassid=split(modelclassid,":")(ubound(split(modelclassid,":")))
					 SET BrandNodes=BrandXML.DocumentElement.SelectNodes("item[@classid='" & C_C(brandclassid,0) &"']")
					 If BrandNodes.length>0 THEN HasBrand=TRUE
				 End If
			 End If
			
				'最后条件加上外层样式
				If I=FieldNode.length-1 Then  FieldMenu=FieldMenu & "<DIV class=""condition_append condition_append_bottom"">" 
				If lcase(CurrNode.SelectSingleNode("@name").text)<>"brandid" or (lcase(CurrNode.SelectSingleNode("@name").text)="brandid" and HasBrand) then
				FieldMenu=FieldMenu & "<DIV class=condition_title>" &  CurrNode.SelectSingleNode("title").text & ":</DIV><DIV class=container>"
				end if
			Select Case lcase(CurrNode.SelectSingleNode("@name").text)
			  case "tid"   '分类
				Dim Node,currk
				currk=KS.ChkClng(split(ModelClassID&":",":")(0))
				If currk=0 Then
				 If SEOTitle="" Then SEOTitle=CurrNode.SelectSingleNode("title").text &":不限" Else SEOTitle=SEOTitle & " " &CurrNode.SelectSingleNode("title").text & ":不限"
				 FieldMenu=FieldMenu & "<strong>不限</strong>"
				Else
				 FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("tid","0","") & """>不限</a>"
				End If
			  K=0
			  If ChannelID=9 Then   '考试分类
				   For Each Node In ClassXML.DocumentElement.SelectNodes("item[@tn=0]")
				    If KS.ChkClng(Node.SelectSingleNode("@id").text)=currk Then
					   If SEOTitle="" Then SEOTitle=CurrNode.SelectSingleNode("title").text &":" & Node.SelectSingleNode("@tname").text Else SEOTitle=SEOTitle & " " &CurrNode.SelectSingleNode("title").text & ":" & Node.SelectSingleNode("@tname").text
	
					   FieldMenu=FieldMenu & "<strong>" & Node.SelectSingleNode("@tname").text & "</strong>"
					 Else
					   FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("tid",Node.SelectSingleNode("@id").text,"") & """ title=""" &Node.SelectSingleNode("@tname").text&""">" & Node.SelectSingleNode("@tname").text & "</a>"
					 End If
					 k=K+1
				   Next
				   Call GetSubSJTidMenu(KS.ChkClng(split(ModelClassID&":",":")(0)))
                   FieldMenu=FieldMenu & "</div>"
			  Else 
				  KS.LoadClassConfig()
				  For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks12=" & ChannelID &" and @ks10=1 and @ks25=1 and @ks14=1]")
					 If KS.ChkClng(Node.SelectSingleNode("@ks9").text)=currk Then
					   If SEOTitle="" Then SEOTitle=CurrNode.SelectSingleNode("title").text &":" & Node.SelectSingleNode("@ks1").text Else SEOTitle=SEOTitle & " " &CurrNode.SelectSingleNode("title").text & ":" & Node.SelectSingleNode("@ks1").text
	
					   FieldMenu=FieldMenu & "<strong>" & Node.SelectSingleNode("@ks1").text & "</strong>"
					 Else
					   FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("tid",Node.SelectSingleNode("@ks9").text,"") & """ title=""" &Node.SelectSingleNode("@ks1").text&""">" & Node.SelectSingleNode("@ks1").text & "</a>"
					 End If
					 k=K+1
				  Next
				   FieldMenu=FieldMenu & "</div>"
				   Dim parentid:parentid=C_C(KS.ChkClng(split(ModelClassID&":",":")(0)),0)
				   Call GetSubTidMenu(parentid)
			 End If
				   FieldMenu=FieldMenu & "<div class=""clear""></div>"
			Case "area" '地区
			   	 set AreaXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				 AreaXML.async = false
				 AreaXML.setProperty "ServerHTTPRequest", true 
				 AreaXML.load(Server.MapPath(KS.Setting(3)& "config/area.xml"))
				 if AreaXML.parseError.errorCode<>0 Then
					Call KS.CreateAreaCache()
				 End If
				 currk=KS.ChkClng(split(GetParam("area")&":",":")(0))
				 If currk=0 Then
				  If SEOTitle="" Then SEOTitle=CurrNode.SelectSingleNode("title").text &":不限" Else SEOTitle=SEOTitle & " " &CurrNode.SelectSingleNode("title").text & ":不限"
				  Province=""
				  FieldMenu=FieldMenu & "<strong>不限</strong>"
				 Else
				  FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("area","0","") & """>不限</a>"
				 End If

                 For Each Node In AreaXML.DocumentElement.SelectNodes("item[@parentid=0 and @filtertf=1]")
					 If KS.ChkClng(Node.SelectSingleNode("@id").text)=currk Then
					   If SEOTitle="" Then SEOTitle=CurrNode.SelectSingleNode("title").text &":" & Node.SelectSingleNode("@city").text Else SEOTitle=SEOTitle & " " &CurrNode.SelectSingleNode("title").text & ":" & Node.SelectSingleNode("@city").text
	                   Province=Node.SelectSingleNode("@city").text
					   FieldMenu=FieldMenu & "<strong>" & Node.SelectSingleNode("@city").text & "</strong>"
					 Else
				       FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("area",Node.SelectSingleNode("@id").text,"") & """ title=""" &Node.SelectSingleNode("@city").text&""">" & Node.SelectSingleNode("@city").text & "</a>"
					 End If
				 Next
				 FieldMenu=FieldMenu & "</div>" 
				 parentid=KS.ChkClng(split(GetParam("area")&":",":")(0))
			     Call GetSubAreaMenu(parentid)
			     FieldMenu=FieldMenu & "<div class=""clear""></div>"
			Case "brandid"   '商城品牌
				If HasBrand Then
				 currk=KS.ChkClng(GetParam("brandid"))
				 If currk=0 Then
				  If SEOTitle="" Then SEOTitle=CurrNode.SelectSingleNode("title").text &":不限" Else SEOTitle=SEOTitle & " " &CurrNode.SelectSingleNode("title").text & ":不限"
				  Province=""
				  FieldMenu=FieldMenu & "<strong>不限</strong>"
				 Else
				  FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("brandid","0","") & """>不限</a>"
				 End If
				  For Each Node In BrandNodes
					 If KS.ChkClng(Node.SelectSingleNode("@id").text)=currk Then
					   If SEOTitle="" Then SEOTitle=CurrNode.SelectSingleNode("title").text &":" & Node.SelectSingleNode("brandname").text Else SEOTitle=SEOTitle & " " &CurrNode.SelectSingleNode("title").text & ":" & Node.SelectSingleNode("brandname").text
					   FieldMenu=FieldMenu & "<strong>" & Node.SelectSingleNode("brandname").text & "</strong>"
					 Else
				       FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("brandid",Node.SelectSingleNode("@id").text,"") & """ title=""" &Node.SelectSingleNode("brandname").text&""">" & Node.SelectSingleNode("brandname").text & "</a>"
					 End If
				  Next	
				  FieldMenu=FieldMenu & "</div>"   
				End If
			case "typeid"  '供求的交易类别
			  call KS.LoadGQTypeToXml()
			     currk=KS.ChkClng(GetParam("typeid"))
				 If currk=0 Then
				  If SEOTitle="" Then SEOTitle=CurrNode.SelectSingleNode("title").text &":不限" Else SEOTitle=SEOTitle & " " &CurrNode.SelectSingleNode("title").text & ":不限"
				  Province=""
				  FieldMenu=FieldMenu & "<strong>不限</strong>"
				 Else
				  FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("typeid","0","") & """>不限</a>"
				 End If
				 For Each Node In Application(KS.SiteSN & "_SupplyType").DocumentElement.SelectNodes("row")
				   If KS.ChkClng(Node.SelectSingleNode("@typeid").text)=currk Then
					   If SEOTitle="" Then SEOTitle=CurrNode.SelectSingleNode("title").text &":" & Node.SelectSingleNode("@typename").text Else SEOTitle=SEOTitle & " " &CurrNode.SelectSingleNode("title").text & ":" & Node.SelectSingleNode("@typename").text
					   FieldMenu=FieldMenu & "<strong>" & Node.SelectSingleNode("@typename").text & "</strong>"
					 Else
				       FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("typeid",Node.SelectSingleNode("@typeid").text,"") & """ title=""" &Node.SelectSingleNode("@typename").text&"""  style='color:" & Node.SelectSingleNode("@typecolor").text & "'>" & Node.SelectSingleNode("@typename").text & "</a>"
					 End If
				 Next
				FieldMenu=FieldMenu & "</div>"   
			Case Else
				If CurrNode.SelectSingleNode("showvalue").text="0" Then   '自定义字段读数据库的值
				    Dim ONode:Set ONode=Application(KS.SiteSN&"_field").documentElement.selectsinglenode("fielditem[@fieldname='" & CurrNode.SelectSingleNode("@name").text &"']/options")
					If Not ONode Is Nothing Then
					 OptionArr=Split(ONode.text,"\n")
					End If
				Else
			        OptionArr=Split(CurrNode.SelectSingleNode("showvalue").text,",")
				End If
					If KS.ChkClng(GetParam("v"&i))=0 Then
				     If SEOTitle="" Then SEOTitle=CurrNode.SelectSingleNode("title").text &":不限" Else SEOTitle=SEOTitle & " " &CurrNode.SelectSingleNode("title").text & ":不限"
					 FieldMenu=FieldMenu & "<strong>不限</strong>"
					Else
					 FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("v"&i,"0","") & """>不限</a>"
					End If
				
					For K=0 To Ubound(OptionArr)
					 Dim v:v=OptionArr(k)
					 if instr(v,"|")<>0 then v=split(v,"|")(1)
					 If KS.ChkClng(GetParam("v"&i))=KS.ChkClng(k+1) Then
					 	If SEOTitle="" Then SEOTitle=CurrNode.SelectSingleNode("title").text &":" & v Else SEOTitle=SEOTitle & " " &CurrNode.SelectSingleNode("title").text & ":" & v
					   FieldMenu=FieldMenu & "<strong>" & v & "</strong>"
					 Else
					   FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("v"&i,K+1,"") & """ title=""" &v&""">" & v & "</a>"
					 End If
					Next
			    FieldMenu=FieldMenu & "</div><div class=""clear""></div>"
			End Select
		  Next
		  If FieldNode.length-1>0 Then FieldMenu=FieldMenu & "</div>"  
		  
		  If ChannelID=9 Then
		  Param=" Where Verific=1"
		  Else
		  Param=" Where Verific=1 and deltf=0"
		  End If

			  '构造条件
			  Dim FieldVal,ValArr,QV,searchvalue,FWValue,FieldName
			 ' SET FieldNode=FieldXML.DocumentElement.SelectNodes("item[@enabled='true']")
			  For K=0 To FieldNode.length-1
			      QV=KS.ChkClng(GetParam("v"&k))
			      Set CurrNode=FieldNode.Item(k)
				  FieldName=CurrNode.SelectSingleNode("@name").text
				  searchvalue=CurrNode.SelectSingleNode("searchvalue").text
				  If lcase(FieldName)="tid" Then  '分类
				      dim tid:tid=ModelClassID
					  if tid<>"" and tid<>"0" then
					    tid=split(tid,":")(ubound(split(tid,":")))
						if ChannelID=9 Then '考试分类
						    Param=Param & " And tid in(Select ID From KS_SJClass Where TS LIKE '%" & KS.ChkClng(tid)& ",%')"
						Else
							tid=c_c(tid,0)
							Param=Param & " And tid in(" & KS.GetFolderTid(tid) & ")"
						End If
					  end if
				  ElseIf CurrNode.SelectSingleNode("showvalue").text="0" Then   '自定义字段读数据库的值
				    Set ONode=Application(KS.SiteSN&"_field").documentElement.selectsinglenode("fielditem[@fieldname='" & FieldName &"']/options")
					If Not ONode Is Nothing Then
					 OptionArr=Split(ONode.text,"\n")
				      if qv-1<=Ubound(OptionArr) and qv-1>=0 Then
					   v=OptionArr(qv-1)
				       if instr(v,"|")<>0 then v=split(v,"|")(0)
					    if CurrNode.SelectSingleNode("fieldtype").text="7" then
					     Param=Param & " and " & FieldName & " like '%" & v &"%'"
						else
					     Param=Param & " and " & FieldName & "='" & v &"'"
						end if
					  End If
					End If
				 ElseIf searchvalue<>"" and searchvalue<>"0" Then   '按设置的选项值搜索
				    OptionArr=Split(searchvalue,",")
					if qv-1<=Ubound(OptionArr) and qv-1>=0 Then
					 v=OptionArr(qv-1)
					else
					 v=""
					end if
					if v<>"" then
						select case lcase(CurrNode.SelectSingleNode("condition").text)
							case "dy" Param=Param & " and " & FieldName & "='" & v &"'"
							case "dys" Param=Param & " and " & FieldName & "=" & v &""
							case "like" Param=Param & " and " & FieldName & " like '%" & v &"%'"
							case "fw" 
							   FWValue=split(v&"-","-")
							   Param=Param & " And (" & FieldName & ">=" & FWValue(0) & " and " & FieldName & "<=" & FWValue(1) &")"                        
							case else
                                 
						end select
					end if
				 End If
			  Next
		  

		 If Not KS.IsNul(Key) And Key<>"0" Then
		     If SEOTitle="" Then SEOTitle="关键字：" & Key  Else SEOTitle=SEOTitle & " 关键字:" & KEY
			 Dim II
			 KeyWordArr=Split(Key," ")
			 Select Case KS.ChkClng(Request("t"))
			  case 1
			       If (FuzzySearch=1) Then
					   For II=0 To Ubound(KeyWordArr)
						   If II=0 Then
						   Param=Param & " And (Title Like '%" & KeyWordArr(II) & "%'"
						   Else
						   Param = Param & " or Title Like '%" & KeyWordArr(II) & "%'"
						   End If
					  Next
				   Else
				    Param=Param & " And (Title Like '%" & Key & "%'"
				   End If
				   Param=Param & ")"
			   Case 2
			    Select Case KS.ChkClng(KS.C_S(ChannelID,6))
				 Case 1:Param=Param & " And ArticleContent Like '%" & Key & "%'"
				 Case 2:Param=Param & " And PictureContent Like '%" & Key & "%'"
				 Case 3:Param=Param & " And DownContent Like '%" & Key & "%'"
				 Case 4:Param=Param & " And FlashContent Like '%" & Key & "%'"
				 Case 5:Param=Param & " And ProIntro Like '%" & Key & "%'"
				 Case 7:Param=Param & " And MovieContent Like '%" & Key & "%'"
				 Case 8:Param=Param & " And GQContent Like '%" & Key & "%'"
				End Select
			   Case 3
			    If KS.ChkClng(KS.C_S(ChannelID,6))<=5 Then
			    Param=Param & " And Author Like '%" & Key & "%'"
				ElseIf KS.ChkClng(KS.C_S(ChannelID,6))=7 Then
			    Param=Param & " And MovieAct Like '%" & Key & "%'"
				End If
			   Case 4:Param=Param & " And Inputer Like '%" & Key & "%'"
			   Case 5:Param=Param & " And KeyWords Like '%" & Key & "%'"
			   Case 6
			     If KS.ChkClng(KS.C_S(ChannelID,6))=5 Then
			       Param=Param & " And ProID Like '%" & Key & "%'"
				 ElseIf KS.ChkClng(KS.C_S(ChannelID,6))=7 Then
			       Param=Param & " And MovieDY Like '%" & Key & "%'"
			     End If
			  case else
				if (FuzzySearch=1) Then
				  For I=0 To Ubound(KeyWordArr)
				   If I=0 Then
				   Param=Param & " And (Title Like '%" & KeyWordArr(i) & "%'"
				   Else
				   Param = Param & " or Title Like '%" & KeyWordArr(i) & "%'"
				   End If
				  Next
				Else
				 Param=Param & " And (Title Like '%" & Key & "%'"
				End If
				 Param=Param & ")"
			 End Select  
         End If
		 
		 if Not KS.IsNUL(TopParam) Then
		    If TopParam<>"1=1" Then
		     Param=Param & " AND (" & TopParam & ")"
			End If
		 End If
		 If KS.ChkClng(KS.C_S(ChannelID,6))=1 Then
			 If Not KS.IsNul(Province) Then
				 Param=Param & " And Province='" & KS.DelSQL(Province) &"'"
			 End If
			 If Not KS.IsNul(City) Then
				 Param=Param & " And City='" & KS.DelSQL(City) &"'"
			 End If
		 ElseIf KS.ChkClng(GetParam("brandid"))<>0 And KS.ChkClng(KS.C_S(ChannelID,6))=5 Then
				 Param=Param & " and brandid=" & KS.ChkClng(GetParam("brandid"))
		 ElseIf KS.ChkClng(GetParam("typeid"))<>0 And KS.ChkClng(KS.C_S(ChannelID,6))=8 Then
				 Param=Param & " and typeid=" & KS.ChkClng(GetParam("typeid"))
		 End If
		 
		  
		   ModelTable=KS.C_S(ChannelID,2)
		   Select Case KS.C_S(ChannelID,6)           rem 查询的字段,方便调用这里把一张表的都查询出来了,你可以只列出要查询的字段
		    case 1 FieldStr="*"
			case 2 FieldStr="*,PictureContent As Intro"
			case 3 FieldStr="*,DownContent As Intro"
			case 4 FieldStr="*,FlashContent As Intro"
			case 5 FieldStr="*,ProIntro As Intro"
			case 7 FieldStr="*,MovieContent As Intro"
			case 8 FieldStr="*,GqContent As Intro"
			Case 9 FieldStr="*,sj As Intro"
		   End Select
		  
		  If OrderField<>"" Then
		   OrderStr=OrderField
		  Else
		   OrderStr=" Order by ID Desc"
		  End If
		  
		  
		  SqlStr="Select " & FieldStr & " From " & ModelTable & Param & OrderStr
		 
		 ' ks.echo sqlstr
		  
		  
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open SqlStr,conn,1,1
		  If RS.Eof And RS.Bof Then
		  Else
		     TotalPut = Conn.Execute("select Count(1) from " & ModelTable & " " & Param)(0)
			 If TotalPut>TopNum And TopNum<>0 Then TotalPut=TopNum
			 if (TotalPut mod MaxPerPage)=0 then
				PageNum = TotalPut \ MaxPerPage
			 else
				PageNum = TotalPut \ MaxPerPage + 1
			 end if
			 If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
					RS.Move (CurrPage - 1) * MaxPerPage
			 Else
					CurrPage = 1
			 End If
			 Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","root")
		  End If
		 RS.Close
		 Set RS=Nothing

		End Sub
		
		Function C_C(ClassID,FieldID)
		   If KS.IsNul(ClassID) Then Exit Function
		    KS.LoadClassConfig()
		   Dim Node:Set Node=Application(KS.SiteSN&"_class").documentElement.selectSingleNode("class[@ks9=" & classID & "]/@ks" & FieldID & "")
		   If Not Node Is Nothing Then C_C=Node.text
		   Set Node=Nothing
	   End Function
	  '分类子栏目
	  Sub GetSubTidMenu(parentid)
			   Dim Node,k,nodes,tj,ts,nn,pparam
			   Set Nodes=Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks12=" & ChannelID &" and @ks13='" & parentid & "']")
			  if Nodes.length>0 Then
			      FieldMenu=FieldMenu & "<DIV class=""condition_title"">&nbsp;</DIV><DIV class=""container"">"
				  nn=0
				  For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks12=" & ChannelID &" and @ks13='" & parentid & "' and @ks25=1 and @ks14=1]")
				     param=""
					 tj=Node.SelectSingleNode("@ks10").text
					 ts=split(Node.SelectSingleNode("@ks8").text,",")
					 for k=0 to ubound(ts)-1
					   if param="" then
					    param=KS.C_C(ts(k),9)
					   else
					    param=param & ":" & KS.C_C(ts(k),9)
					   end if
					   if k<> ubound(ts)-1 then  pparam=param
					 next
					 currclassid=split(ModelClassID&":",":")(tj-1)
					 if nn=0 then
					  If ks.chkclng(currclassid)=0 Then
						 FieldMenu=FieldMenu & "<strong>不限</strong>"
					  Else
						 FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("tid",pparam,"") & """>不限</a>"
					  End If
					 end if
					 nn=nn+1
					 If KS.ChkClng(currclassid)=KS.ChkClng(Node.SelectSingleNode("@ks9").text) Then
					 	SEOTitle=SEOTitle & "," &Node.SelectSingleNode("@ks1").text
					    FieldMenu=FieldMenu & "<strong>" & Node.SelectSingleNode("@ks1").text & "</strong>"
					 Else
					   FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("tid",param,"") & """ title=""" &Node.SelectSingleNode("@ks1").text&""">" & Node.SelectSingleNode("@ks1").text & "</a>"
					 End If
				Next
				FieldMenu=FieldMenu & "</div>"
				Parentid=C_C(currclassid,0)
				Call GetSubTidMenu(parentid)
			End If
		End Sub
		
	  '试卷分类子栏目
	  Sub GetSubSJTidMenu(parentid)
	          if KS.ChkClng(parentid)=0 then exit sub
			   Dim Node,k,nodes,tj,ts,nn,pparam
			   Set Nodes=ClassXML.DocumentElement.SelectNodes("item[@tn=" & KS.ChkClng(parentid) &"]")
			  if Nodes.length>0 Then
			      FieldMenu=FieldMenu & "<DIV class=""condition_title"">&nbsp;</DIV><DIV class=""container"">"
				  nn=0
				  For Each Node In ClassXML.DocumentElement.SelectNodes("item[@tn=" & KS.ChkClng(parentid) &"]")
				     param=""
					 tj=Node.SelectSingleNode("@tj").text
					 ts=split(Node.SelectSingleNode("@ts").text,",")
					 for k=0 to ubound(ts)-1
					   if param="" then
					    param=ts(k)
					   else
					    param=param & ":" & ts(k)
					   end if
					   if k<> ubound(ts)-1 then  pparam=param
					 next
					 currclassid=split(ModelClassID&":",":")(tj-1)
					 if nn=0 then
					  If ks.chkclng(currclassid)=0 Then
						 FieldMenu=FieldMenu & "<strong>不限</strong>"
					  Else
						 FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("tid",pparam,"") & """>不限</a>"
					  End If
					 end if
					 nn=nn+1
					 If KS.ChkClng(currclassid)=KS.ChkClng(Node.SelectSingleNode("@id").text) Then
					 	SEOTitle=SEOTitle & "," &Node.SelectSingleNode("@tname").text
					    FieldMenu=FieldMenu & "<strong>" & Node.SelectSingleNode("@tname").text & "</strong>"
					 Else
					   FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("tid",param,"") & """ title=""" &Node.SelectSingleNode("@tname").text&""">" & Node.SelectSingleNode("@tname").text & "</a>"
					 End If
				Next
				FieldMenu=FieldMenu & "</div>"
				Parentid=currclassid
				Call GetSubSJTidMenu(parentid)
			End If
		End Sub
	   
	   '子地区
	   Sub GetSubAreaMenu(parentid)
	     If IsObject(AreaXML) and parentid<>0 Then
             Dim Node,currareaid,k,nodes,tj,ts,nn,pparam
			   Set Nodes=AreaXML.DocumentElement.SelectNodes("item[@parentid=" & parentid &" and @filtertf=1]")
			  if Nodes.length>0 Then
			      FieldMenu=FieldMenu & "<DIV class=""condition_title"">&nbsp;</DIV><DIV class=""container"">"
				  nn=0
				  For Each Node In Nodes
				     param=parentid & ":" & Node.SelectSingleNode("@id").text
					 currareaid=split(GetParam("area")&":",":")(1)
					 if nn=0 then
					  If ks.chkclng(currareaid)=0 Then
					     City=""
						 FieldMenu=FieldMenu & "<strong>不限</strong>"
					  Else
						 FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("area",parentid&":0","") & """>不限</a>"
					  End If
					 end if
					 nn=nn+1
					 If KS.ChkClng(currareaid)=KS.ChkClng(Node.SelectSingleNode("@id").text) Then
					 	SEOTitle=SEOTitle & "," &Node.SelectSingleNode("@city").text
						City=Node.SelectSingleNode("@city").text
					    FieldMenu=FieldMenu & "<strong>" & Node.SelectSingleNode("@city").text & "</strong>"
					 Else
					   FieldMenu=FieldMenu & "<a href=""" & GetFieldLink("area",param,"") & """ title=""" &Node.SelectSingleNode("@city").text&""">" & Node.SelectSingleNode("@city").text & "</a>"
					 End If
				Next
				FieldMenu=FieldMenu & "</div>"
			End If		 
		 
		 End If
	   End Sub
		
		Function GetParam(key)
		  If KS.IsNUL(Request.QueryString) Then GetParam="":Exit Function
		  Dim PArr:Parr=Split(Replace(Request.QueryString,".html",""),",")
		  Dim Pkey,Pval,i
		  For I=0 To Ubound(Parr)
		     if instr(Parr(i),"-")<>0 Then
			   Pkey=Split(Parr(i),"-")(0)
		       if lcase(key)=lcase(Pkey) Then
			     Pval=Split(Parr(i),"-")(1)
				 Exit For
			   End If
			 End If
		  Next
		  GetParam=Pval
		End Function
		

		
		'获取字段选项的链接,nocollectkey 不收集的字段，多个用英文逗号隔开
		Function GetFieldLink(f,v,nocollectkey)
		 Dim r,Param,i
		 Param="" 
		 Dim QPArr,PPArr,Pkey,Pval,QP
		 QP=Request.QueryString
		 If Not KS.IsNul(QP) Then
		   QPArr=Split(Replace(QP,".html",""),",")
		   For i=0 To Ubound(QPArr)
		     If instr(QParr(i),"-")<>0 Then
			   PPArr=Split(QParr(i),"-")
			   Pkey=trim(PPArr(0))
			   Pval=PPArr(1)
			  if KS.FoundInArr(lcase(nocollectkey),lcase(pkey),",")=false Then
				   If lcase(pkey)<>lcase(f) Then
					 If Param="" Then
					   Param=pkey&"-"&Pval
					 Else
					   Param=Param & ","&pkey&"-"&Pval
					 End If
				   ElseIf lcase(pkey)=lcase(f) Then
					 If Param="" Then
					   Param=pkey&"-" & v
					 Else
					   Param=Param & ","&pkey&"-"&v
					 End If
				   End If
			  End If
			 End If
		   Next
		 End If
        
		if f<>"" then
		 if instr(lcase(param),lcase(f)&"-")=0 then 
				If Param="" Then
				 Param=f & "-" & v
				Else
				 Param=Param & "," & f & "-" & v
				End If
		 end if
		end if
		if (instr(lcase(param),"c-")=0) then
		       If Param="" Then
				 Param="c-" & channelid
				Else
				 Param=Param & ",c-" & channelid
				End If
		end if
		 if isrewrite<>"1" then Param="?" & Param Else Param=KS.Setting(3) & "search/" & Param
		 GetFieldLink=Param&".html"
		End Function
		
		
		'伪静态分页
		Public Function ShowPage()
		           Dim I, pageStr
				   pageStr= ("<div id=""fenye"" class=""fenye""><table border='0' align='right'><tr><td>")
					if (CurrPage>1) then pageStr=PageStr & "<a href=""" & GetFieldLink("p",CurrPage-1,"")& """ class=""prev"">上一页</a>"
				   if (CurrPage<>PageNum) then pageStr=PageStr & "<a href=""" & GetFieldLink("p",CurrPage+1,"") & """ class=""next"">下一页</a>"
				    pageStr=pageStr & "<a href=""" & GetFieldLink("p",1,"") & """ class=""prev"">首 页</a>"
				 
					Dim startpage,n,j
					 if (CurrPage>=7) then startpage=CurrPage-5
					 if PageNum-CurrPage<5 Then startpage=PageNum-10
					 If startpage<=0 Then startpage=1
					 n=0
					 For J=startpage To PageNum
						If J= CurrPage Then
						 PageStr=PageStr & " <a href=""#"" class=""curr""><font color=red>" & J &"</font></a>"
						Else
						 PageStr=PageStr & " <a class=""num"" href=""" & GetFieldLink("p",j,"")&""">" & J &"</a>"
						End If
						n=n+1 : if n>=10 then exit for
					 Next
					
					 PageStr=PageStr & " <a class=""next"" href=""" & GetFieldLink("p",pagenum,"")&""">末页</a>"
					 pageStr=PageStr & " <span>共" & totalPut & "条记录,分" & PageNum & "页</span></td></tr></table>"
				     PageStr = PageStr & "</td></tr></table></div>"
			         ShowPage = PageStr
	     End Function

		
End Class
%>

 
