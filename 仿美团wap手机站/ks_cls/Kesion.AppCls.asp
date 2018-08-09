<!--#include file="../Conn.asp"-->
<!--#include file="Kesion.Label.CommonCls.asp"-->
<!--#include file="Kesion.StaticCls.asp"-->
<!--#include file="Kesion.ClubCls.asp"-->
<!--#include file="Kesion.SpaceApp.asp"-->
<%


Class KesionAppCls
        Private KS,KSUser, KSR,Tp
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		  Set KSR = New Refresh
		End Sub
		Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSR=Nothing
		 Set KSUser=Nothing
		End Sub
        
		
		Public Sub HomePage()
			   Dim QueryStrings:QueryStrings=Request.ServerVariables("QUERY_STRING")
			   If QueryStrings<>"" And Ubound(Split(QueryStrings,"-"))>=1 Then
				 Call StaticCls.Run()
			   ElseIf Not KS.IsNul(Request.QueryString("do")) Then
			       Select Case lcase(KS.S("DO"))
					  case "vote" vote
				   End Select
			   Else
				  Dim Template,FsoIndex:FsoIndex=KS.Setting(5)
				  FCls.RefreshType = "INDEX" '设置刷新类型，以便取得当前位置导航等
				  FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签

				  IF Split(FsoIndex,".")(1)<>"asp" Then
				    Dim PerTime:PerTime=KS.ChkClng(KS.Setting(130))
				    Dim IsExistIndex:IsExistIndex=KS.CheckFile(KS.Setting(3) & KS.Setting(5))
				    If IsExistIndex= False Then
					   Template=KSR.KSLabelReplaceAll(KSR.LoadTemplate(KS.Setting(110)))
					   Call KS.WriteTOFile(KS.Setting(5),Template)
					   KS.Die Template
					ElseIf PerTime>0 Then
					     Dim fs:set fs=KS.InitialObject(KS.Setting(99)) 
						 Dim f:set f=fs.getfile(server.MapPath(KS.Setting(5))) 
						 Dim LastModified:LastModified=f.DateLastModified 
					    if datediff("n",LastModified,Now)>=PerTime Then 
						 Template=KSR.KSLabelReplaceAll(KSR.LoadTemplate(KS.Setting(110)))
						 Call KS.WriteTOFile(KS.Setting(5),Template) 
						 KS.Die Template
						else
						 KS.Die KS.ReadFromFile(KS.Setting(3) & KS.Setting(5))
					    End If 
					Else
				       KS.Die KS.ReadFromFile(KS.Setting(3) & KS.Setting(5))
					End If
				  Else
					  Template=KSR.KSLabelReplaceAll(KSR.LoadTemplate(KS.Setting(110)))
				 End IF
				 Response.Write Template  
			  End If
			  Set StaticCls=Nothing
		End Sub
		
		'二级域名
		Public Sub Domain(S)
		   Select Case Lcase(S)
		     case lcase(KS.WSetting(1))
			       response.Redirect(KS.WSetting(4))
		     case lcase(KS.Setting(69))     '论坛
				 dim Club
				 if instr(lcase(request.querystring),lcase(GCls.ClubPreContent))<>0 then
				  set Club=new ClubDisplayCls
				 else
				  set Club=new ClubCls
				 end if
				 Club.kesion
				 Set Club=Nothing

		   End Select
		End Sub
		
		
		
		
		'投票系统
		Private Sub Vote()
		   Dim ID:ID=KS.ChkClng(KS.S("ID"))
		   If Id=0 Then KS.Die "error!"
		   Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		   RS.Open "Select Top 1 * From KS_Vote Where id=" & id,CONN,1,1
		   If RS.Eof And RS.Bof Then
		     RS.Close:Set RS=Nothing
			 KS.Die "error!"
		   End If
		   
		   Dim LoopStr,XML,Node,Str,LC,Xstr,TotalVote
		   
		   '投票操作
		   If KS.S("Action")="dovote" Then
		     If RS("Status")="0" Then
			   KS.Die "<script>alert('该投票已关闭!');location.href='?do=vote&id=" & id&"';</script>"
			 End If
			 Dim LoginTF:LoginTF=KSUser.UserLoginChecked()
			 Dim GroupIds:GroupIds=RS("GroupIDs")
			 If RS("nmtp")="1" and LoginTF=false Then
	            KS.Die "<script>alert('对不起，只会登录会员才能投票!');location.href='?do=vote&id=" & id&"';</script>"
			 End If
			 If Not KS.IsNul(GroupIDs) And KS.FoundInArr(GroupIDs, KSUser.GroupID, ",")=False Then
			 	KS.Die "<script>alert('对不起，您所在的会员组不允许投票!');location.href='?do=vote&id=" & id&"';</script>"
			 End If
			 If RS("TimeLimit")="1" Then
			 	if now<RS("TimeBegin") then KS.Die "<script>alert('对不起，该投票于" & RS("TimeBegin") & "开放！');location.href='?do=vote&id=" & id&"';</script>"
		        if now>RS("TimeEnd") then KS.Die "<script>alert('对不起，该投票已在" & RS("TimeBegin") & "停止！');location.href='?do=vote&id=" & id&"';</script>"
			 End If
			 
			 
		     Dim VoteOption,ItemArr,I,UserName
			 VoteOption=KS.FilterIds(KS.S("VoteOption"))
			 If KS.IsNul(VoteOption) Then
			   KS.Die "<script>alert('您没有选择投票项目!');location.href='?do=vote&id=" & id&"';</script>"
			 End If
			 
			 Dim IPNum:IPNum=KS.ChkClng(RS("IpNum"))
			 Dim IPNums:IPNums=RS("IPNums")
			 If IpNums<>0 Then
			 	If Conn.Execute("Select Count(ID) From KS_PhotoVote Where UserIp='" & KS.GetIP & "' and ChannelID=-1 And InfoID='" & ID & "'")(0)>=IPNums  Then
	             KS.Die "<script>alert('对不起，最多只能投" & IPNums & "次!');location.href='?do=vote&id=" & id&"';</script>"
	             End If
			 End If
			 If IpNum<>0 Then
			 	If Conn.Execute("Select Count(ID) From KS_PhotoVote Where Year(VoteTime)=" & Year(Now) & " and Month(VoteTime)=" & Month(Now) & " and Day(VoteTime)=" & Day(Now) & " and UserIp='" & KS.GetIP & "' and ChannelID=-1 And InfoID='" & ID & "'")(0)>=IPNum  Then
	             KS.Die "<script>alert('对不起，一天最多只能投" & IPNum & "次!');location.href='?do=vote&id=" & id&"';</script>"
	             End If
			 End If
			 
			 ItemArr=Split(VoteOption,",")
		     set XML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			 XML.async = false
			 XML.setProperty "ServerHTTPRequest", true 
			 XML.load(Server.MapPath(KS.Setting(3)&"Config/voteitem/vote_" &id&".xml"))
			 For I=0 To Ubound(ItemArr)
				 Set Node=XML.DocumentElement.SelectSingleNode("voteitem[@id=" & KS.ChkClng(ItemArr(i)) & "]")
				 Node.childNodes(1).text=KS.ChkClng(Node.childNodes(1).text)+1
				 XML.Save(Server.MapPath(KS.Setting(3)&"Config/voteitem/vote_" &id&".xml"))
			 Next
			  Application(KS.SiteSN&"_Configvoteitem/vote_"&ID)=""
			 If LoginTF=False Then UserName="游客" Else UserName=KSUser.UserName
			 Conn.Execute("Insert Into [KS_PhotoVote]([ChannelID],[ClassID],[InfoID],[VoteTime],[UserName],[UserIP]) Values(-1,'-1','" & ID & "'," & SqlNowString & ",'" & UserName & "','" & KS.GetIP() & "')")

		   End If
		   
		   Dim Tp:Tp = KSR.LoadTemplate(RS("TemplateID"))
		   If KS.IsNul(Tp) Then 
		     KS.Die "您绑定的模板没有内容,请检查!"
		   End If
		   LoopStr=KS.CutFixContent(Tp, "[loop]", "[/loop]", 0)
		   If Not IsObject(XML) Then
		   Set XML=LFCls.GetXMLFromFile("voteitem/vote_"&ID)
		   End If
		   For Each Node In Xml.DocumentElement.SelectNodes("voteitem")
		       Xstr=Xstr & "<set label='" & Node.childNodes(0).text &"' value='" &Node.childNodes(1).text &"' />"
			   TotalVote=TotalVote+KS.ChkClng(Node.childNodes(1).text)
		   Next
		   For Each Node In Xml.DocumentElement.SelectNodes("voteitem")
		       LC=LoopStr
			   If RS("VoteType")="Single" Then
			   LC=Replace(LC,"{@ItemType}","<input type='radio' name='VoteOption' value='"& Node.getAttribute("id") &"' />")
			   Else
			   LC=Replace(LC,"{@ItemType}","<input type='checkbox' name='VoteOption' value='"& Node.getAttribute("id") &"' />")
			   End If
			   LC=Replace(LC,"{@ItemID}",Node.getAttribute("id"))
			   LC=Replace(LC,"{@ItemName}",Node.childNodes(0).text)
			   LC=Replace(LC,"{@Num}",Node.childNodes(1).text)
            
			dim perVote,pstr
			if totalVote=0 Then TotalVote=0.00000001
			perVote=round(Node.childNodes(1).text/totalVote,4)
			pstr="<img src='../images/Default/bar.gif' width='" & round(360*perVote) & "' height='15' align='absmiddle' />"
			perVote=perVote*100
			if perVote<1 and perVote<>0 then
				pstr=pstr & "&nbsp;0" & perVote & "%"
			else
				pstr=pstr & "&nbsp;" & perVote & "%"
			end if			   
			   LC=Replace(LC,"{@Percent}",Pstr)

			   Str=Str & LC
		   Next
		   Tp=Replace(Tp,KS.CutFixContent(Tp, "[loop]", "[/loop]", 1),str)
		   Tp=Replace(Tp,"{$VoteName}",rs("title"))
		   Tp=Replace(Tp,"{$VoteID}",id)
		   Tp=Replace(Tp,"{$VoteItemXML}",Xstr)
		   RS.Close:Set RS=Nothing
		   Tp=KSR.KSLabelReplaceAll(Tp)
           KS.Die Tp
		End Sub
		
End Class
%>