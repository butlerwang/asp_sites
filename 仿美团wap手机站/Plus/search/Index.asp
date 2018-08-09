<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Template.asp"-->
<%

Dim KSCls
Set KSCls = New Link
KSCls.Kesion()
Set KSCls = Nothing

Class Link
        Private KS,ChannelID,ModelTable,Param,XML,Node,StartTime
		Private CurrPage,MaxPerPage,TotalPut,PageNum,Key,stype,OrderStr
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  MaxPerPage=10
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		
		Public Sub Kesion()
		
		Dim RefreshTime:RefreshTime = 1  '设置防刷新时间
		If DateDiff("s", Session("SearchTime"), Now()) < RefreshTime Then
			Response.Write "<META http-equiv=Content-Type content=text/html; chaRset=utf-8><meta HTTP-EQUIV=REFRESH CONTENT="&RefreshTime&"><br>本页面起用了防刷新机制，请不要在"&RefreshTime&"秒内连续刷新本页面<BR>正在打开页面，请稍后……"
			Response.End
		End If
		Session("SearchTime")=Now()

		
		 Dim Template,KSR
		 FCls.RefreshType = "searchIndex"   
		 Set KSR = New Refresh
		   If KS.Setting(139)="" Then KS.Die "请先到""基本信息设置->模板绑定""进行模板绑定操作!"
		   Template = KSR.LoadTemplate(KS.Setting(139))
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
						    case "menu"  SearchMenu
						    case "showpage" echo KS.ShowPage(totalput, MaxPerPage, "", CurrPage,false,false)
							case "totalput" echo TotalPut
							case "leavetime" 
							   dim leavetime:leavetime=FormatNumber((timer-starttime),5)
							   if leavetime<1 then leavetime="0"&leavetime
							   echo leavetime
							case "keyword" echo KS.R(key)
							case "channelid" echo channelid
							case "stype" echo stype
							case "records" SearchRecords
							case "relative" Searchrelative
						  end select
			End Select
		End Sub
		Sub EchoItem(sTokenName)
		  Select Case sTokenName
		    case "id" echo GetNodeText("id")
			case "linkurl" 
			   If ChannelID=0 Then
					echo KS.GetItemURL(GetNodeText("channelid"),GetNodeText("tid"),GetNodeText("infoid"),GetNodeText("fname"))
			   ElseIf ChannelID=102 Then
				  If KS.ASetting(16)="1" Then
				      echo KS.GetDomain & "ask/show-" & GetNodeText("id") & KS.ASetting(17)
				  Else
				      echo KS.GetDomain & "ask/q.asp?id=" & GetNodeText("id")
				  End If
			   Else
				    echo KS.GetItemURL(ChannelID,GetNodeText("tid"),GetNodeText("id"),GetNodeText("fname"))
			   End If 
			 
			case "classname" 
			  If ChannelID=102 Then
			   echo GetNodeText("pclassname") & GetNodeText("classname")
			  Else
			   echo KS.C_C(GetNodeText("tid"),1)
			  End If
			case "classurl" 
			 If ChannelID=102 Then
			  echo KS.GetDomain & "ask/showlist.asp?id=" & Node.SelectSingleNode("@classid").text
			 Else
			  echo KS.GetFolderPath(GetNodeText("tid"))
			 End If
			case "intro" 
			 Dim Intro:intro=KS.Gottopic(KS.LoseHtml(GetNodeText("intro")),160)
			 Intro=Replace(Intro,"&nbsp;","")
			 If Not KS.IsNul(Key) Then
			  echo Replace(Intro,key,"<span style='color:red'>" & key & "</span>")
			 Else
			 echo intro
			 End If
			case else
			  echo GetNodeText(sTokenName)
		  End Select
		End Sub
		Function GetNodeText(NodeName)
		 Dim N,Str
		 NodeName=Lcase(NodeName)
		 If IsObject(Node) Then
		  set N=node.SelectSingleNode("@" & NodeName)
		  If Not N is Nothing Then Str=N.text
		  If Not KS.IsNul(Key)  And NodeName="title" Then
		   Str=Replace(Str,key,"<span style='color:red'>" & key & "</span>")
		  End If
		  GetNodeText=Str
		 End If
		End Function
		
		
		Sub InitialSearch()
		  Dim FieldStr,SqlStr,TopStr,TopNum
		  ChannelID=KS.ChkClng(Request("M"))
		  CurrPage=KS.ChkClng(Request("Page"))
		  If CurrPage<=0 Then CurrPage=1
		  Key=KS.CheckXSS(KS.R(KS.S("Key")))
		  stype=KS.ChkClng(Request("stype"))
		  
		  If ChannelID=102 Then
		   Param=" Where LockTopic=0"
		  Else
		   Param=" Where Verific=1 and deltf=0"
		  End If
		  
		  If Not KS.IsNul(Key) Then
				select case stype
				 case 100 
				  if IsDate(Key) Then
					   Param=Param & " And AddDate>=#" & Key & " 00:00:00# and AddDate<=#" &Key& " 23:59:59#"
				  End If
				 case 2 
				   If ChannelID=102 Then
					Param=Param & " And Title Like '%" & Key & "%'"
				   Else
					  Select Case KS.C_S(ChannelID,6)
						case 1 Param=Param & " And ArticleContent Like '%" & Key & "%'"
						case 2 Param=Param & " And PictureContent Like '%" & Key & "%'"
						case 3 Param=Param & " And DownContent Like '%" & Key & "%'"
						case 4 Param=Param & " And FlashContent Like '%" & Key & "%'"
						case 5 Param=Param & " And ProIntro Like '%" & Key & "%'"
						case 7 Param=Param & " And MovieContent Like '%" & Key & "%'"
						case 8 Param=Param & " And GQContent Like '%" & Key & "%'"
					  End Select
				  End If	  
				 case 3 
				   If ChannelID=102 Then
				     Param=Param & " And UserName Like '%" & Key & "%'"
				   Else
				     Param=Param & " And inputer Like '%" & Key & "%'"
				   End If
				 case else 
				   Param=Param & " and title like '%" & Key & "%'"
				   'Dim KeyParam
				   'KeyParam=AutoKey(key,"Title")
				   'If KeyParam<>"" Then
				   'Param=Param & " And " & KeyParam
				   'End If
				end select
		  TopNum=0
		 Else
		 TopNum=1000  rem 没有输入关键词只列表最新1000条记录
         End If
		 
		 if request("classid")<>"" and request("classid")<>"0" then
		   If ChannelID<>102 Then
		     Param=Param & " And Tid In(" & KS.GetFolderTid(KS.S("ClassID")) & ")"
		   end if
		 end if
		 
		If TopNum<>0 Then TopStr=" Top " & TopNum
		  
		  If ChannelID=0 Then
		   ModelTable="KS_ItemInfo"
		   FieldStr="ID,Tid,Title,ChannelID,InfoID,Intro,AddDate,Fname"
		  ElseIf ChannelID=102 Then
		   ModelTable="KS_AskTopic"
		   FieldStr="topicid as id,classid,pclassname,classname,Title,title as Intro,DateAndTime as AddDate"
		  Else
		   ModelTable=KS.C_S(ChannelID,2)
		   Select Case KS.C_S(ChannelID,6)
		    case 1 FieldStr="ID,Tid,Title,Intro,AddDate,Fname"
			case 2 FieldStr="ID,Tid,Title,PictureContent As Intro,AddDate,Fname"
			case 3 FieldStr="ID,Tid,Title,DownContent As Intro,AddDate,Fname"
			case 4 FieldStr="ID,Tid,Title,FlashContent As Intro,AddDate,Fname"
			case 5 FieldStr="ID,Tid,Title,ProIntro As Intro,AddDate,Fname"
			case 7 FieldStr="ID,Tid,Title,MovieContent As Intro,AddDate,Fname"
			case 8 FieldStr="ID,Tid,Title,GqContent As Intro,AddDate,Fname"
		   End Select
		  End If
		  
		  If ChannelID=102 Then
		  Else
		   OrderStr=" Order by ID Desc"
		  End If
		  
		  
		  SqlStr="Select " & TopStr & " " & FieldStr & " From " & ModelTable & Param & OrderStr
		  'ks.echo sqlstr
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open SqlStr,conn,1,1
		  If RS.Eof And RS.Bof Then
		  Else
		     TotalPut = Conn.Execute("select Count(1) from " & ModelTable & " " & Param)(0)
			 If TotalPut>TopNum And TopNum<>0 Then TotalPut=TopNum
			 If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
					RS.Move (CurrPage - 1) * MaxPerPage
			 Else
					CurrPage = 1
			 End If
			 Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","root")
		  End If
		 RS.Close
		 Set RS=Nothing
		 KeyToDataBase()
		End Sub
		
		Sub KeyToDataBase()
		  If KS.IsNul(Trim(Key)) or CurrPage>1 Then Exit Sub
		  Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		  RS.Open "Select top 1 * From KS_KeyWords Where IsSearch=1 and KeyText='" & Key &"'",conn,1,3
		  If RS.Eof Then
		    RS.AddNew
			RS("AddDate")=Now
			RS("IsSearch")=1
			RS("KeyText")=Key
			RS("Hits")=1
		  End If
		    RS("Hits")=RS("Hits")+1
			RS("LastUseTime")=Now
		  RS.Update
		  RS.Close:Set RS=Nothing
		End Sub
		
		Sub SearchRecords()
		  Dim RecordList,Rarr,K,I,InArray,KArr
		  RecordList=Application(KS.SiteSN&"_SearchRecords")
		  If Not KS.IsNul(Key) Then
		    Rarr=Split(RecordList,"§")
			InArray=false
			For I=0 To Ubound(Rarr)
			 If lcase(Split(Rarr(i),"@")(0))=lcase(Key) Then
			  InArray=true : Exit For
			 End If
			Next
		    If InArray=false Then
			  If RecordList="" Then
			    RecordList=Key & "@" & TotalPut & "@" & ChannelID & "@" & stype
			  Else
			    RecordList= Key & "@" & TotalPut & "@" & ChannelID & "@" & stype & "§" & RecordList
			  End If
			  Rarr=Split(RecordList,"§")
			  For I=0 To Ubound(Rarr) 
			    If I>30 Then Exit For
			    If i=0 Then
				  RecordList=Rarr(I)
				Else
				  RecordList=RecordList & "§" & Rarr(i)
				End If
			  Next
			  
			  Application.Lock
			   Application(KS.SiteSN&"_SearchRecords")=RecordList
			  Application.unLock
			End If
		  End If
		  RecordList=Application(KS.SiteSN&"_SearchRecords")
		  If Not KS.IsNul(RecordList) Then
		    Rarr=Split(RecordList,"§")
			For Each K in Rarr
			  Karr=Split(K,"@")
			  If Karr(0)<>"" Then
			   echo "<li><a href=""?m=" & Karr(2)&"&key=" & server.URLEncode(Karr(0)) &"&stype=" & karr(3) & """>" & Karr(0) &" (结果:<font style='color:red'>" & KS.ChkClng(Karr(1)) & "</font>条)</a></li>"
			  End If
			Next
		  Else
		     echo "没有任何搜索记录!"
		  End If
		End Sub
		
		Sub Searchrelative()
		  If KS.IsNul(Key) Then Exit Sub
		  Dim I,RS,XML,N,SQLStr,Param,KeyParam
		  KeyParam=AutoKey(key,"keytext")
		  Param="Where IsSearch=1" 
		  If KeyParam<>"" Then Param=Param & " and " & KeyParam
		  SQLStr="Select Top 10 KeyText From KS_KeyWords " & Param
		  'KS.Echo SQLSTR
		  Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open SQLStr,conn,1,1
		  If Not RS.Eof Then Set XML=KS.RSToXml(RS,"row","")
		  If IsObject(XML) Then
		    For Each N In XML.DocumentElement.SelectNodes("row")
			 If N.SelectSingleNode("@keytext").text<>Key Then
			 echo "<li><a href=""?key=" & server.URLEncode(N.SelectSingleNode("@keytext").text) &"&"& KS.QueryParam("key,page") & """>" & N.SelectSingleNode("@keytext").text & "</a></li>"
			 End If
			Next
		  End If
		  XML=Empty
		End Sub
		
		Function AutoKey(ByVal strKey,FieldName) 
			CONST lngSubKey=2 
			Dim lngLenKey, Param, i, strSubKey 
			strKey=Replace(strKey," ","")
			lngLenKey=Len(strKey)
			If lngLenKey <=1 Then AutoKey="(" & FieldName & " like '%" & strKey & "%')": Exit Function
			'若长度大于1，则从字符串首字符开始，循环取长度为2的子字符串作为查询条件 
			For i=1 To lngLenKey-(lngSubKey-1) 
			  strSubKey=Mid(strKey,i,lngSubKey) 
			  If Param="" Then
			   Param = "(" & FieldName & " like '%" & strSubKey & "%'" 
			  Else
			   Param=Param & " or " & FieldName & " like '%" & strSubKey & "%'"
			  End If
			Next
			If Param<>"" Then Param=Param & ")"
			AutoKey=Param	
		End Function 

		
		
		
		Sub SearchMenu()
		 If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
		Dim ModelXML,Node
		Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
		If ChannelID=0 Then
		  echo "<li class=""curr""><a href=""?stype="&stype&"&key="&server.URLEncode(key) &""">全部</a></li>"
		Else
		  echo "<li><a href=""?stype="&stype&"&key="&server.URLEncode(key) &""">全部</a></li>"
		End If
		For Each Node In ModelXML.documentElement.SelectNodes("channel")
		 if Node.SelectSingleNode("@ks21").text="1" and Node.SelectSingleNode("@ks0").text<>"6" and Node.SelectSingleNode("@ks0").text<>"9" and Node.SelectSingleNode("@ks0").text<>"10" Then
		  If Trim(ChannelID)=Trim(Node.SelectSingleNode("@ks0").text) Then
		  echo "<li class=""curr""><a href=""?stype="&stype&"&key="&server.URLEncode(key) &"&m="&Node.SelectSingleNode("@ks0").text &""">" & Node.SelectSingleNode("@ks3").text & "</a></li>"
		  Else
		  echo "<li><a href=""?stype="&stype&"&key="&server.URLEncode(key) &"&m="&Node.SelectSingleNode("@ks0").text &""">" & Node.SelectSingleNode("@ks3").text & "</a></li>"
		  End If
		 End If
		next
		If  KS.ASetting(0)="1" then
			If ChannelID=102 Then
			  echo "<li class=""curr""><a href=""?stype="&stype&"&key="&server.URLEncode(key) &"&m=102"">问吧</a></li>"
			Else
			  echo "<li><a href=""?stype="&stype&"&key="&server.URLEncode(key) &"&m=102"">问吧</a></li>"
			End If
		End If
		End Sub
		
End Class
%>

 
