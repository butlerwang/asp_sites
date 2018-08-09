<%

Dim LFCls:Set LFCls=New LabelBaseFunCls
Class LabelBaseFunCls
		Private KS     
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set LFCls=Nothing
		End Sub
				
		'显示自定义字段的表单验证
		Public Sub ShowDiyFieldCheck(FieldXML,flag)
			  Dim Node,FieldName,FieldType,XTitle
			  If Not IsObject(FieldXML) Then Exit Sub
			  if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
			      Dim PStr:Pstr="fieldtype!=0&&fieldtype!=13"
				  if flag=0 then Pstr=Pstr & "&&showonuserform=1"
				  Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem["& Pstr &"]")
				  If DiyNode.Length>0 Then
						For Each Node In DiyNode
						     FieldName = Node.SelectSingleNode("@fieldname").text
							 FieldType = KS.ChkClng(Node.SelectSingleNode("fieldtype").text)
							 XTitle    = Node.SelectSingleNode("title").text
							If FieldType=10 And Node.SelectSingleNode("mustfilltf").text="1" Then
								 if flag=1 then
								 Response.Write "if (CKEDITOR.instances." & fieldName& ".getData()==''){alert('" & Xtitle & "必须输入内容！'); CKEDITOR.instances." & fieldName& ".focus(); return false;}" &vbcrlf
								 else
								 Response.Write "if (CKEDITOR.instances." & fieldName& ".getData()==''){$.dialog.alert('" & Xtitle & "必须输入内容！',function(){CKEDITOR.instances." & fieldName& ".focus();}); return false;}" &vbcrlf
								 end if
							ElseIf Node.SelectSingleNode("mustfilltf").text="1" Then 
								 if flag=1 then
								  Response.Write "if (jQuery('#" & FieldName & "').val()==''){alert('" & XTitle & "必须填写!');jQuery('#" & FieldName & "').focus();return false;}" & vbcrlf
								 else
								  Response.Write "if (jQuery('#" & FieldName & "').val()==''){$.dialog.alert('" & XTitle & "必须填写!',function(){jQuery('#" & FieldName & "').focus();});return false;}" & vbcrlf
								 end if
							End If
					        If (FieldType=4 or FieldType=12) Then 
								  if flag=1 then
								   Response.Write "if (jQuery('#" & FieldName &"').val()!=''&& !is_number(jQuery('#" & FieldName & "').val())){alert('" & Xtitle & "必须填写数字!');jQuery('#" & FieldName & "').focus();return false;}"& vbcrlf
								  else
								   Response.Write "if (jQuery('#" & FieldName &"').val()!=''&& !is_number(jQuery('#" & FieldName & "').val())){$.dialog.alert('" & Xtitle & "必须填写数字!',function(){jQuery('#" & FieldName & "').focus();});return false;}"& vbcrlf
								  end if
							end if
					        If FieldType=5 Then 
							  if flag=1 then
						   	   Response.Write "if (jQuery('#" & FieldName & "').val()!=''&&is_date(jQuery('#" & FieldName & "').val())==false){alert('" & XTitle & "必须填写正确的日期!');jQuery('#" & FieldName & "').focus();return false;}" & vbcrlf
							  else
						   	   Response.Write "if (jQuery('#" & FieldName & "').val()!=''&&is_date(jQuery('#" & FieldName & "').val())==false){$.dialog.alert('" & XTitle & "必须填写正确的日期!',function(){jQuery('#" & FieldName & "').focus();});return false;}" & vbcrlf
							  end if
							end if
					        If FieldType=8  and Node.SelectSingleNode("mustfilltf").text="1" Then 
							 if flag=1 then
							  Response.Write "if (is_email(jQuery('#" & FieldName & "').val())==false){alert('" & XTitle & "必须填写正确的邮箱!');jQuery('#" & FieldName & "').focus();return false;}" & vbcrlf
							 else
							  Response.Write "if (is_email(jQuery('#" & FieldName & "').val())==false){$.dialog.alert('" & XTitle & "必须填写正确的邮箱!',function(){jQuery('#" & FieldName & "').focus();});return false;}" & vbcrlf
							 end if
							End If
						Next
				  End If
			  End If
		End Sub
		
		'前台向主数据表插入数据
		Sub InserItemInfo(ChannelID,InfoID,Title,Tid,Intro,KeyWords,PhotoUrl,Inputer,Verific,Fname)
		 Call AddItemInfo(ChannelID,InfoID,Title,Tid,Intro,KeyWords,PhotoUrl,now,Inputer,0,0,0,0,0,0,0,0,0,0,1,Verific,Fname)
		End Sub
		'前台向主数据表修改数据
		Sub ModifyItemInfo(ChannelID,InfoID,Title,Tid,Intro,KeyWords,PhotoUrl,Verific)
		 Gcls.Execute("Update [KS_ItemInfo] Set Title='" & replace(Title,"'","''") & "',Tid='" & Tid & "',Intro='" & Replace(left(KS.LoseHtml(Intro),255),"'","''")  & "',KeyWords='" & Replace(KeyWords,"'","''") & "',PhotoUrl='" & PhotoUrl & "',AddDate=" & SQLNowString &",ModifyDate=" & SQLNowString &",Verific=" & Verific & " Where  ChannelID=" & ChannelID & " and InfoID=" & InfoID)
		End Sub		
		
        '后台向系统主数据表添加数据
        Sub AddItemInfo(ByVal ChannelID,InfoID,Title,Tid,Intro,KeyWords,PhotoUrl,AddDate,Inputer,Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific,Fname)
		 If KS.IsNul(Intro) Then Intro=" "
		 If KS.IsNul(KeyWords) Then KeyWords=" "
		 AddDate=replace(replace(replace(replace(AddDate,"PM ",""),"AM ",""),"上午 ",""),"下午 ","")
		 Gcls.Execute("Insert Into [KS_ItemInfo](ChannelID,InfoID,Title,Tid,Intro,KeyWords,PhotoUrl,AddDate,ModifyDate,Inputer,Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific,DelTF,Fname) values(" & Channelid & "," & InfoID & ",'" & replace(Title,"'","''") & "','" & Tid & "' ,'" & left(Replace(KS.CheckScript(KS.LoseHtml(Intro)),"'","''"),200) & "','" & Replace(KeyWords,"'","''") & "' ,'" & PhotoUrl & "' ,'" & AddDate & "','" & AddDate&"','" & Inputer & "' ," & KS.ChkClng(Hits) & "," & KS.ChkClng(HitsByDay) & ", " & KS.ChkClng(HitsByWeek) & "," & KS.ChkClng(HitsByMonth) & "," & KS.ChkClng(Recommend) & "," & KS.ChkClng(Rolls) & "," & KS.ChkClng(Strip) & "," & KS.ChkClng(Popular) & "," & KS.ChkClng(Slide) & "," & KS.ChkClng(IsTop) & "," & KS.ChkClng(Comment) & "," & KS.ChkClng(Verific)& ",0,'" & Fname & "')")
		End Sub
		'后台修改数据表数据
		Sub UpdateItemInfo(ChannelID,InfoID,Title,Tid,Intro,KeyWords,PhotoUrl,AddDate,Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific)
		 Gcls.Execute("Update [KS_ItemInfo] Set Title='" & replace(Title,"'","''") & "',Tid='" & Tid & "',Intro='" & Replace(left(KS.LoseHtml(Intro),255),"'","''")  & "',KeyWords='" & Replace(KeyWords,"'","''") & "',PhotoUrl='" & PhotoUrl & "',AddDate='" & AddDate & "',ModifyDate=" & SQLNowString & ",Hits=" & KS.ChkClng(Hits) & ",HitsByDay=" & KS.ChkClng(HitsByDay) & ",HitsByWeek=" & KS.ChkClng(HitsByWeek) & ",HitsByMonth=" & KS.ChkClng(HitsByMonth) & ",Recommend=" & KS.ChkClng(Recommend) & ",Rolls=" & KS.ChkClng(Rolls) & ",Strip=" & KS.CHkClng(Strip) & ",Popular=" & KS.ChkClng(Popular) & ",Slide=" & KS.ChkClng(Slide) & ",IsTop=" & KS.ChkClng(IsTop)  &",Comment=" & KS.ChkClng(Comment) & ",Verific=" & KS.CHkClng(Verific) & " Where  ChannelID=" & KS.ChkClng(ChannelID) & " and InfoID=" & KS.ChkClng(InfoID))
		End Sub		
		'*********************************************************************************************************
		'函数名：GetAbsolutePath
		'作  用：返回数据库的绝对路径
		'参  数：RelativePath 数据库连接字段串
		'*********************************************************************************************************
		Function GetAbsolutePath(RelativePath)
			dim Exp_Path,Matches,tempStr
			tempStr=Replace(RelativePath,"\","/")
			if instr(tempStr,":/")>0 then
				GetAbsolutePath=RelativePath
				Exit Function
			End if
			set Exp_Path=new RegExp
			Exp_Path.Pattern="(Data Source=|dbq=)(.)*"
			Exp_Path.IgnoreCase=true
			Exp_Path.Global=true
			Set Matches=Exp_Path.Execute(tempStr)
			If instr(LCase(tempStr),"*.xls")<>0 Then
			GetAbsolutePath="driver={microsoft excel driver (*.xls)};dbq="&Server.MapPath(split(Matches(0).value,"=")(1))
			ElseIf Instr(Lcase(tempstr),"*.dbf")<>0 Then
			GetAbsolutePath="driver={microsoft dbase driver (*.dbf)};dbq="&Server.MapPath(split(Matches(0).value,"=")(1))
			Else
			GetAbsolutePath="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(split(Matches(0).value,"=")(1))
			End If
		End Function

		'*********************************************************************************************************
		'函数名：ReplaceDBNull
		'作  用：替换数据库空值通用函数
		'参  数：DBField 字段值,DefaultValue 空间替换的值
		'*********************************************************************************************************
		Function ReplaceDBNull(DBField,DefaultValue)
		    If IsNull(DBField) Or (DBField="") then
			ReplaceDBNull=DefaultValue
			Else
			ReplaceDBNull = DBField
			end if
		End Function
		'*********************************************************************************************************
		'函数名：GetSingleFieldValue
		'作  用：取单字段值
		'参  数：SQLStr SQL语句
		'*********************************************************************************************************
		Function GetSingleFieldValue(SQLStr)
		    If DataBaseType=0 then
			On Error Resume Next
			GetSingleFieldValue=Conn.Execute(SQLStr)(0)
			If Err Then GetSingleFieldValue=""
			Else
			 Dim RS:Set RS=Conn.Execute(SQLStr)
			 If Not RS.Eof Then
			  GetSingleFieldValue=RS(0)
			 Else
			  GetSingleFieldValue=""
			 End If
			 RS.Close:Set RS=Nothing
			end if
		End Function
		
		'*********************************************************************************************************
		'函数名：GetConfigFromXML
		'作  用：取xml节点配置信息
		'参  数：FileName xml文件名(不含扩展名),Path 节点路径 ,NodeName 节点Name属性值
		'*********************************************************************************************************
		Function GetConfigFromXML(FileName,Path,NodeName)
		  If Not IsObject(Application(KS.SiteSN&"_Config"&FileName)) Then
			  Set Application(KS.SiteSN&"_Config"&FileName)=GetXMLFromFile(FileName)
		  End If
		  Dim Node:Set Node= Application(KS.SiteSN&"_Config"&FileName).documentElement.selectSingleNode(Path & "[@name='" & NodeName & "']")
		  If Not Node Is Nothing Then
		   GetConfigFromXML=Node.text
		  End If
		End Function
		'从config中读取xml配置，不缓存
		'FileName 文件名,Path 节能点路径,Condition 得到节点条件
		Function GetXMLByNoCache(FileName,Path,Condition)
		      Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			  Doc.async = false
			  Doc.setProperty "ServerHTTPRequest", true 
			  Doc.load(Server.MapPath(KS.Setting(3)&"Config/" & FileName &".xml"))
			  dim loadnum:loadnum=0
			  do while Doc.parseError.errorCode<>0   '出错重新加载
				 Doc.load(Server.MapPath(KS.Setting(3)&"Config/" & FileName &".xml"))
				 loadnum=loadnum+1
				 if loadnum>5 then exit do
			  loop
			  Dim Node:Set Node= Doc.documentElement.selectSingleNode(Path & Condition)
			  If Not Node Is Nothing Then  GetXMLByNoCache=Node.text
		End Function
		'*********************************************************************************************************
		'函数名：GetXMLFromFile
		'作  用：取xml文件到Application
		'参  数：FileName xml文件名(不含扩展名)
		'*********************************************************************************************************
		Function GetXMLFromFile(FileName)
		 	If Not IsObject(Application(KS.SiteSN&"_Config"&FileName)) Then
			  Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			  Doc.async = false
			  Doc.setProperty "ServerHTTPRequest", true 
			  Doc.load(Server.MapPath(KS.Setting(3)&"Config/" & FileName &".xml"))
			  dim loadnum:loadnum=0
			  do while Doc.parseError.errorCode<>0   '出错重新加载
				 Doc.load(Server.MapPath(KS.Setting(3)&"Config/" & FileName &".xml"))
				 loadnum=loadnum+1
				 if loadnum>5 then exit do
			  loop
			  if loadnum=0 then  Set Application(KS.SiteSN&"_Config"&FileName)=Doc
		   End If  
            Set GetXMLFromFile=Application(KS.SiteSN&"_Config"&FileName)
		End Function
		
		
		'*********************************************************************************************************
		'函数名：GetFileIDFromContent
		'作  用：从内容里取附件ID号
		'参  数：Content内容
		'*********************************************************************************************************
		Function GetFileIDFromContent(Content)
		        Dim re,FileIDs,fileid,Upfile
				Set re = New RegExp
				re.IgnoreCase = True
				re.Global = True
				re.Pattern = "\[UploadFiles\][0-9]*"
				Set UpFile = re.Execute(Content)
				Set re = Nothing
				For Each fileid In UpFile
				  fileid=replace(lcase(fileid),"[uploadfiles]","")
				  If FileIDs="" Then
				   FileIDs=fileid
				  Else
				   FileIDs=FileIDs & "," & fileid
				  End If
				Next
				GetFileIDFromContent=KS.FilterIds(FileIDs)
		End Function
		
		'*********************************************************************************************************
		'函数名：ReplacePrevNext
		'作  用：上一篇、下一篇
		'参  数：NowID 现在ID,Tid 目录ID,TypeStr类型
		'*********************************************************************************************************
		Function GetPrevNextURL(ChannelID,NowID, Tid, TypeStr,ByRef Title)
		     If Fcls.RefreshType<>"Content" Then GetPrevNextURL="#":Title="此标签只能放内容页模板!" : Exit Function
		     Dim SqlStr,LinkUrl
		     SqlStr="SELECT Top 1 ID,Title,Tid,Fname From " & KS.C_S(ChannelID,2) & " Where Tid='" & Tid & "' And ID" & TypeStr & NowID & " And Verific=1 and  DelTF=0 Order By ID"
			 If TypeStr=">" Then SqlStr=SqlStr & " asc" else SqlStr=SqlStr & " desc"
			 Dim RS:Set RS=Conn.Execute(SqlStr)
			 If RS.EOF And RS.BOF Then
			  GetPrevNextURL = "#" : Title = "没有了"
			 Else
			  LinkUrl = KS.GetItemURL(ChannelID,RS(2),RS(0),RS(3))
			  GetPrevNextURL = LinkUrl : Title= "<a href=""" & LinkUrl & """ title=""" & RS(1) & """>" & RS(1) & "</a>"
			 End If
			 RS.Close:Set RS = Nothing
		End Function
		Function ReplacePrevNext(ChannelID,NowID, Tid, TypeStr)
		     Dim Title
			 Call GetPrevNextURL(ChannelID,NowID, Tid, TypeStr,Title)
			 ReplacePrevNext=Title
		End Function
		
        '取文本字段的值
		'参数说明：字段值,截段字数,未尾输出的字符,HTML处理方式
		Function Get_Text_Field(FieldValue,CutNum,EndTag,HtmlTag,DefaultChar)
		 Dim TempStr:TempStr=FieldValue
		 If KS.IsNul(FieldValue) Then Get_Text_Field=DefaultChar : Exit Function
		 If Not IsNumeric(HtmlTag) Or Not IsNumeric(CutNum) Then Exit Function
		 If HtmlTag=0 Then
		  TempStr=KS.HtmlCode(TempStr)
		 ElseIf HtmlTag=1 Then
		  TempStr=TempStr
		 ElseIF HtmlTag=2 Then
		  TempStr=Replace(KS.LoseHtml(KS.HtmlCode(TempStr))," ","")
		 End If
          If EndTag="0" Then EndTag=""
		  if KS.strLength(TempStr)>cint(CutNum) and CutNum<>0 then TempStr = KS.GotTopic(TempStr, Cint(CutNum)) & EndTag
		 Get_Text_Field=TempStr
		End Function
		'取数字字段的值
		'参数说明：FieldValue-字段值,OutType-输出方式0、原数，1、小数，2百分数,XSWS-小数位数
		Function Get_Num_Field(FieldValue,OutType,XSWS)
		 If Not IsNumeric(FieldValue) Then Get_Num_Field=FieldValue:Exit Function
		 If Not IsNumeric(OutType) Then OutType=0
		 If Not IsNumeric(XSWS) Then XSWS=0
         If OutType=1 Then
		   Get_Num_Field=FormatNumber(FieldValue,XSWS)
		 ElseIf OutType=2 Then
		   Get_Num_Field=FormatPercent(FieldValue)
		 Else
		   Get_Num_Field=FieldValue
		 End if  
		End Function
		'取日期字段的值
		'参数说明：FieldValue-字段值,DateMB-输出日期模板
		Function Get_Date_Field(FieldValue,ByVal DateMB)
		  IF Not IsDate(FieldValue) Then Get_Date_Field=FieldValue:Exit Function
		  If Instr(DateMB,"YYYY")<>0 Then DateMB=Replace(DateMB,"YYYY",Year(FieldValue))
		  If Instr(DateMB,"YY")<>0 Then   DateMB=Replace(DateMB,"YY",Right("0" & Year(FieldValue), 2))
		  If Instr(DateMB,"MM")<>0 Then   DateMB=Replace(DateMB,"MM",Right("0" & Month(FieldValue), 2))
		  If Instr(DateMB,"DD")<>0 Then   DateMB=Replace(DateMB,"DD",Right("0" & Day(FieldValue), 2))
		  If Instr(DateMB,"hh")<>0 Then   DateMB=Replace(DateMB,"hh",Right("0" & hour(FieldValue), 2))
		  If Instr(DateMB,"mm")<>0 Then   DateMB=Replace(DateMB,"mm",Right("0" & minute(FieldValue), 2))
		  If Instr(DateMB,"ss")<>0 Then   DateMB=Replace(DateMB,"ss",Right("0" & second(FieldValue), 2))
		  Get_Date_Field=DateMB
		End Function		
		
		'替换自定义字段
		Function ReplaceUserDefine(ChannelID,F_C,ByVal RS)
		   If Not IsObject(Application(KS.SiteSN&"_userfiledlist"&channelid)) Then
		     Set  Application(KS.SiteSN&"_userfiledlist"&channelid)=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			 Application(KS.SiteSN&"_userfiledlist"&channelid).appendChild(Application(KS.SiteSN&"_userfiledlist"&channelid).createElement("xml"))
				Dim D_F_Arr,K,Node,FieldName
				Dim KS_RS_Obj:Set KS_RS_Obj=Conn.Execute("Select FieldName From KS_Field Where ChannelID=" & ChannelID &" And ShowOnForm=1 Order By OrderID Asc")
				If Not KS_RS_Obj.Eof Then D_F_Arr=KS_RS_Obj.GetRows(-1)
			    KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
				If IsArray(D_F_Arr) Then
					  For K=0 To Ubound(D_F_Arr,2)
						Set Node=Application(KS.SiteSN&"_userfiledlist"&channelid).documentElement.appendChild(Application(KS.SiteSN&"_userfiledlist"&channelid).createNode(1,"userfiledlist"&channelid,""))
						Node.attributes.setNamedItem(Application(KS.SiteSN&"_userfiledlist"&channelid).createNode(2,"fieldname","")).text=D_F_Arr(0,K)
					 Next
				 End If
		 End If
		 
		 For Each Node in Application(KS.SiteSN&"_userfiledlist"&channelid).documentElement.SelectNodes("userfiledlist"&channelid)
		     FieldName=Node.selectSingleNode("@fieldname").text
			If Not IsNull(RS(FieldName)) Then
			  F_C=Replace(F_C,"{$" & FieldName & "}",RS(FieldName))
			Else
			  F_C=Replace(F_C,"{$" & FieldName & "}","")
			End If
		 Next
		ReplaceUserDefine=F_C
	End Function
End Class
%> 
