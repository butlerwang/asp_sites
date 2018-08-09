<!--#include file="Kesion.IfCls.asp"-->
<%
Sub Echo(sStr)
	 Response.Write sStr 
	 'Response.Flush()
End Sub
  
public Sub Scan(ByVal sTemplate)
	Dim iPosLast, iPosCur
	iPosLast    = 1
	Do While True 
		iPosCur    = InStr(iPosLast, sTemplate, "[#") 
		If iPosCur>0 Then
			Echo    Mid(sTemplate, iPosLast, iPosCur-iPosLast)
			iPosLast  = Parse(sTemplate, iPosCur+2)
		Else 
			Echo    Mid(sTemplate, iPosLast)
			Exit Do  
		End If 
	Loop
End Sub

Function Parse(sTemplate, iPosBegin)
	Dim iPosCur, sToken, sTemp,MyNode
	iPosCur      = InStr(iPosBegin, sTemplate, "]")
	sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
	iPosBegin    = iPosCur+1
	select case Lcase(sTemp)
		case "pubtips"  
		  If Action="Edit" Then
		    echo "修改" & KS.C_S(Channelid,3)
		  Else
		    echo "发布" & KS.C_S(Channelid,3)
		  End If
		case "selectclassid"
		   If KS.C("UserName")="" Then  '游客投稿
		    echo "[" & KS.GetClassNP(KS.S("ClassID")) & "] <a href=""Contributor.asp""><<重新选择>></a>"
			echo "<input type=""hidden"" name=""ClassID"" value=""" & KS.S("ClassID") & """>"
		   Else
		   Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton) 
		   End If
		case "status"
		   if action="Edit" Then
		     If RS("Verific")<>1 Then
			  if rs("verific")=2 Then
		       echo "<label style='color:#999999'><input type='checkbox' name='status' value='2' onclick='if(!this.checked){return(confirm(""确定立即投稿发布吗?""));}' checked>放入草稿</label>"
			  else
		       echo "<label style='color:#999999'><input type='checkbox' name='status' onclick='if(!this.checked){return(confirm(""确定立即投稿发布吗?""));}' value='2'>放入草稿</label>"
			  end if
			 Else
			  echo "<input type=""hidden"" name=""okverific"" value=""1""><input type=""hidden"" name=""verific"" value=""1"">"
			 End If
		   Else
		    echo "<label style='color:#999999'><input type='checkbox' name='status' value='2'>放入草稿</label>"
		   End If
		case "readpoint" 
		   if action="Edit" Then echo rs("readpoint") else echo "0"
		case "showsetthumb"
		   if action<>"Edit" Then echo "<label><input type='checkbox' name='autothumb' id='autothumb' value='1' checked>使用图集的第一幅图</label>"
		case "showphotourl"
			If KS.C("UserName")="" Then%>
				  <td width="240"><input class="textbox" name='PhotoUrl'  type='text' style="width:230px;" id='PhotoUrl' maxlength="100" /></td>
			 <%Else
			   if action="Edit" Then PhotoUrl=rs("PhotoUrl")
			   %>
				<td width="340"><input class="textbox" name='PhotoUrl' value="<%=PhotoUrl%>" type='text' style="width:230px;" id='PhotoUrl' maxlength="100" />
                 <input class="uploadbutton1" type='button' name='Submit3' value='选择图片' onClick="OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("选择" & KS.C_S(ChannelID,3))%>&ChannelID=4',500,360,window,document.myform.PhotoUrl);" />
			 </td>
		 <%End If
		case "showselecttk"
		  If KS.C("UserName")<>"" Then echo "<button type=""button""  class=""pn"" onClick=""AddTJ();"" style=""margin: -6px 0px 0 0;""><strong>图片库...</strong></button>"
		case "showquestionandverify"
			If KS.C("UserName")="" Then
			Call PubQuestion()
			%>
				<tr class="tdbg">
						<td  height="25" align="center"><span>验证码：</span></td>
						<td>
						 <script type="text/javascript">writeVerifyCode('<%=KS.GetDomain%>',1,'textbox')</script>
						</td>
				</tr>
		<% End If
		case "showstyle"
		   if action="Edit" Then
		    ShowStyle=RS("ShowStyle"): PageNum=RS("PageNum")
		   Else
		    ShowStyle=4 : PageNum=10
		   End If
		   %>
		   <table width='80%'><tr><td>
								  <input type='radio' onClick="$('#pagenums').hide();" name='showstyle' value='4'<%If ShowStyle="4" Then response.Write " checked"%>><img src='../images/default/p4.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td><td>
								  <input type='radio' onClick="$('#pagenums').hide();" name='showstyle' value='1'<%If ShowStyle="1" Then response.Write " checked"%>><img src='../images/default/p1.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td>
		   <td><input type='radio' onClick="$('#pagenums').show();" name='showstyle' value='2'<%If ShowStyle="2" Then Response.Write " checked"%>><img src='../images/default/p2.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td><td><input type='radio' onClick="$('#pagenums').show();" name='showstyle' value='3'<%If ShowStyle="3" Then Response.Write " checked"%>><img src='../images/default/p3.gif'>
		   </td><td><input type='radio' onClick="$('#pagenums').hide();" name='showstyle' value='5'<%If ShowStyle="5" Then Response.Write " checked"%>><img src='../images/default/p5.gif'>
		   </td></tr></table><div style="margin:5px" id="pagenums"
			<%If ShowStyle="1" or ShowStyle="4" Then Response.Write " style='display:none'"%>
			>每页显示<input type="text" name="pagenum" value="<%=PageNum%>" style="text-align:center;width:30px">张</div>
		<%
		case "downlblist" echo DownLBList
		case "downyylist" echo DownYYList
		case "downsqlist" echo DownSQList
		case "downptlist" echo DownPTList
		case "sizeunit"
		   Dim SizeUnit      
		   If Action="Edit" Then SizeUnit = Right(rs("DownSize"), 2) Else SizeUnit="KB"
			Response.Write "<input name=""SizeUnit"" type=""radio"" value=""KB"" id=""kb"""
			If SizeUnit = "KB" Then response.write "checked"
			Response.Write "><label for=""kb"">KB</label> " & vbCrLf
			Response.Write "<input type=""radio"" name=""SizeUnit"" value=""MB"" id=""mb"""
			If SizeUnit = "MB" Then response.write "checked"
			Response.Write "><label for=""mb"">MB</label> " & vbCrLf
		case "content"
		  If Action="Edit" Then
		   select case KS.ChkClng(KS.C_S(ChannelID,6))
		    case 1 if not KS.IsNul(rs("ArticleContent")) then echo Server.HtmlEncode(rs("ArticleContent"))
			case 2 if not KS.IsNul(rs("PictureContent")) then echo Server.HtmlEncode(rs("PictureContent"))
			case 3 if not KS.IsNul(rs("DownContent")) then echo Server.HtmlEncode(rs("DownContent"))
		   end select
		  End If
		case else
		   Dim II,DV,XNode
		   if instr(sTemp,"|select")<>0 then  '下拉及联动
		     Set Xnode=FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='" & split(sTemp,"|")(0) &"']")
			 if Not Xnode is nothing then
				 If Action="Edit" Then DV=RS(split(sTemp,"|")(0)) Else DV=xnode.selectsinglenode("defaultvalue").text
				 KS.Echo KSUser.GetSelectOption(ChannelID,FieldDictionary,FieldXML,xnode.selectsinglenode("fieldtype").text,split(sTemp,"|")(0),xnode.selectsinglenode("width").text,xnode.selectsinglenode("options").text,DV) 
			 end if
		   elseif instr(sTemp,"|radio")<>0 then  '单选
		     Set Xnode=FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='" & split(sTemp,"|")(0) &"']")
			 if Not Xnode is nothing then
			      If Action="Edit" Then DV=RS(split(sTemp,"|")(0)) Else DV=xnode.selectsinglenode("defaultvalue").text
			       KS.Echo  KSUser.GetRadioOption(split(sTemp,"|")(0),xnode.selectsinglenode("options").text,DV)
			 End If
		   elseif instr(sTemp,"|checkbox")<>0 then  '多选
		     Set Xnode=FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='" & split(sTemp,"|")(0) &"']")
			 if Not Xnode is nothing then
			      If Action="Edit" Then DV=RS(split(sTemp,"|")(0)) Else DV=xnode.selectsinglenode("defaultvalue").text
			       KS.Echo  KSUser.GetCheckBoxOption(split(sTemp,"|")(0),xnode.selectsinglenode("options").text,DV)
			 End If
		   elseif instr(sTemp,"|unit")<>0 then  '单位
		     Set Xnode=FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='" & split(sTemp,"|")(0) &"']")
			 if Not Xnode is nothing then
			      If Action="Edit" Then DV=RS(split(sTemp,"|")(0)&"_unit") Else DV=xnode.selectsinglenode("defaultvalue").text
			       KS.Echo  KSUser.GetUnitOption(split(sTemp,"|")(0),xnode.selectsinglenode("unitoptions").text,DV)
			 End If
		   elseif action="Add" then
		     if lcase(trim(stemp))="author" and Not KS.IsNul(KS.C("UserName")) then
			   echo KSUser.GetUserInfo("RealName")
			 end if
			 if lcase(trim(stemp))="origin" and Not KS.IsNul(KS.C("UserName")) then
			   echo LFCls.GetSingleFieldValue("SELECT top 1 CompanyName From KS_EnterPrise Where UserName='" & KSUser.UserName & "'")
			 end if
		   elseif action="Edit" Then
		     echo rs(trim(stemp))
		   Elseif left(lcase(sTemp),3)="ks_" then
		     echo server.htmlencode(GetDiyFieldValue(FieldXML,sTemp))
		   End If
	end select
	Parse    = iPosBegin
 End Function
 
 
'=========================扫描会员中心主体框架 增加于2010年6月========================================

Public Sub Kesion()
         Dim LoginTF:LoginTF=Cbool(KSUser.UserLoginChecked)
		 Dim FileContent,MainUrl,RequestItem,TemplateFile
		 Dim KSR,ParaList
		 FCls.RefreshType = "Member"   '设置当前位置为会员中心
		 Set KSR = New Refresh
		 TemplateFile=KS.Setting(116)
		 If LoginTF=True Then  TemplateFile=KS.U_G(KSUser.GroupID,"templatefile")
		 If trim(TemplateFile)="" Then TemplateFile=KS.Setting(116)
         If trim(TemplateFile)="" Then Response.Write "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
		 FileContent = KSR.LoadTemplate(TemplateFile)
		 If Trim(FileContent) = "" Then FileContent = "模板不存在!"
		  FileContent = KSR.KSLabelReplaceAll(FileContent)
		 Set KSR = Nothing
		 ScanTemplate RexHtml_IF(FileContent)
End Sub	
 
public Sub ScanTemplate(ByVal sTemplate)
	Dim iPosLast, iPosCur
	iPosLast    = 1
	Do While True 
		iPosCur    = InStr(iPosLast, sTemplate, "{#") 
		If iPosCur>0 Then
			Echo    Mid(sTemplate, iPosLast, iPosCur-iPosLast)
			iPosLast  = ParseTemplate(sTemplate, iPosCur+2)
		Else 
			Echo    Mid(sTemplate, iPosLast)
			Exit Do  
		End If 
	Loop
End Sub

Function ParseTemplate(sTemplate, iPosBegin)
		Dim iPosCur, sToken, sTemp,MyNode,CheckJS
		iPosCur      = InStr(iPosBegin, sTemplate, "}")
		sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
		iPosBegin    = iPosCur+1
		select case Lcase(sTemp)
			case "showusermain"  loadMain
			case "showmymenu"  ShowMyMenu
			case "userid"  echo ks.c("userid")
			case "username" echo ksuser.username
			case "groupname" echo KS.U_G(KSUser.GroupID,"groupname")
			case "showsynchronizedoption"  echo KSUser.ShowSynchronizedOption(CheckJS)
			case "checkjs" echo checkjs
			
			case "userface"
			  Dim UserFaceSrc:UserFaceSrc=KSUser.GetUserInfo("UserFace")
			  if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc=KS.Setting(3) & userfacesrc
			  response.write userfacesrc
			case else
			  response.write ksuser.getuserinfo(sTemp)
		end select
		 ParseTemplate=iPosBegin
End Function


 
 Sub ShowMyMenu()
   %>
		<h3>我的面板</h3>
		<div class="left02">
		  <ul>
		     <li><img src="images/icon10.png" align="absmiddle"/> <a href="user_editinfo.asp">会员资料</a>
			 <span><a href="user_rz.asp" title="实名认证" style="color:red;">实名认证</a></span>
			 </li>
		     <li><img src="images/money.jpg" align="absmiddle"/><a href="user_logmoney.asp">消费明细</a>
			 <span><a href="user_payonline.asp">充值</a></span></li>
		 <%
		
		
		 If KS.C_S(5,21)=1 Then
			Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write "<img src=""images/icon9.png"" align=""absmiddle"" /> <a href=""user_order.asp"">商城订单</a>"
			Response.Write "<span><a href=""user_order.asp?action=coupon"">优惠券</a></span></li>"
		 End if
		

		 If KSUser.CheckPower("s20")=true Then 
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"	<img src=""images/icon13.png"" align=""absmiddle"" /> <a href=""User_ItemSign.asp"">文档签收</a>"
			Response.Write "<span><a  href=""User_ItemSign.asp"">查看</a></span></li>"
		 End If
         if KSUser.CheckPower("s16")=true then
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"	<img src=""images/star.gif"" align=""absmiddle"" /> <a href=""User_favorite.asp"">收藏夹</a>"
			Response.Write "<span><a  href=""User_MyComment.asp"">评论</a></span></li>"
		 End If
         if KSUser.CheckPower("s17")=true then
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"	<img src=""images/icon15.png"" align=""absmiddle"" /> <a href=""Complaints.asp"">投诉建议</a>"
			Response.Write "<span><a  href=""Complaints.asp?Action=Add"">+发布</a></span></li>"
		 End If
		 %>
		  </ul>
		</div>
     <h3>内容发布</h3>
		<div class="left02">
		  <ul>
		    <%
  
  
  	 
'模型的投稿
if KSUser.CheckPower("s18")<>false Then 
			 Dim Node,Ico,ItemUrl,PubUrl,Itemname
			 If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig()
			 
			 For Each Node In Application(KS.SiteSN&"_ChannelConfig").DocumentElement.SelectNodes("channel[@ks21=1 and @ks36!=0 and @ks0!=6]")
				Ico=Node.SelectSingleNode("@ks51").text
				If KS.IsNul(Ico) Then Ico="images/icon7.png"
				Select Case KS.ChkClng(Node.SelectSingleNode("@ks6").text) 
				  Case 1 ItemUrl="User_ItemInfo.asp":PubUrl="user_post.asp"
				  Case 2 ItemUrl="User_ItemInfo.asp":PubUrl="user_post.asp"
				  Case 3 ItemUrl="User_ItemInfo.asp":PubUrl="user_post.asp"
				  Case 4 ItemUrl="User_ItemInfo.asp":PubUrl="User_Myflash.asp"
				  Case 5 ItemUrl="User_ItemInfo.asp":PubUrl="User_MyShop.asp"
				  Case 7 ItemUrl="User_ItemInfo.asp":PubUrl="User_MyMovie.asp"
				  Case 8 ItemUrl="User_ItemInfo.asp":PubUrl="User_MySupply.asp"
				  Case 9 ItemUrl="User_MyExam.asp":ItemUrl="User_MyExam.asp"
			   End Select
			        ItemName=Node.SelectSingleNode("@ks52").text
					If KS.IsNul(ItemName) Then ItemName=KS.C_S(Node.SelectSingleNode("@ks0").text,3)
			   		Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
					Response.Write "<img src=""" & Ico & """ align=""absmiddle"" /> <a href=""" & ItemUrl &"?channelid="& Node.SelectSingleNode("@ks0").text & """>" & ItemName & "</a>"
					If KS.ChkClng(Node.SelectSingleNode("@ks6").text) =9 Then
					Response.Write "<span><a href=""User_MyExam.asp?action=record"">+记录</a></span></li>"
					Else
					Response.Write "<span><a href=""" & PubUrl &"?channelid="& Node.SelectSingleNode("@ks0").text & "&Action=Add"">+发布</a></span></li>"
					End If
			 Next
	   End If
		 
		 '求职
		If KS.C_S(10,21)=1 Then
			If KSUser.GetUserInfo("UserType")=0 Then
				If KSUser.CheckPower("s14")=true Then 
					 If KS.C_S(10,21)="1" Then
						Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
						Response.Write "<img src=""images/icon10.png"" align=""absmiddle"" /> <a href=""User_JobResume.asp"">找工作</a>"
						Response.Write "<span><a href=""User_JobResume.asp"">+简历</a></span></li>"

					 End If
				End If
			Else			 
			if KSUser.CheckPower("s14")=true  then
						Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
						Response.Write "<img src=""images/icon10.png"" align=""absmiddle"" /> <a href=""user_Enterprise.asp?action=job"">找人才</a>"
						Response.Write "<span><a href=""User_JobCompanyZW.asp?Action=Add"">+发布</a></span></li>"
			 end if
            End If
		End If
		 If KSUser.CheckPower("s09")=true  and  KS.ASetting(0)="1" Then 
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"	<img src=""images/icon11.png"" align=""absmiddle"" /> <a href=""User_Askquestion.asp"">问答</a>"
			Response.Write "<span><a  href=""../ask/a.asp"" target=""_blank"">+提问</a></span></li>"
		 End If
		 If KSUser.CheckPower("s19")=true and KS.Setting(56)="1" Then 
		    Response.Write "<li onMouseOver=""this.className='hvr'"" onMouseOut=""this.className=''"">"
			Response.Write"	<img src=""images/icon11.png"" align=""absmiddle"" /> <a href=""User_mytopic.asp"">论坛</a>"
			Response.Write "<span><a  href=""User_mytopic.asp?action=fav"">收藏帖</a></span></li>"
		 End If
		 %>
		 </ul>
		</div>
		
		 <%
End Sub
 
'------扫描会员中心主体框架------

 
 
 
 '取得某个字段的默认值
 Function GetDiyFieldValue(FieldXML,FieldName) 
        Dim V,Xnode:Set Xnode=FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='" & FieldName &"']")
		if Not Xnode is nothing then
		  v=Xnode.selectsinglenode("defaultvalue").text
		End If
		If Instr(V,"|")<>0 Then
			 If Not KS.IsNul(KS.C("UserName")) Then
			 V=LFCls.GetSingleFieldValue("select top 1 " & Split(V,"|")(1) & " from " & Split(V,"|")(0) & " where username='" & KSUser.UserName & "'") 
			 Else
			 V=""
			 End If
		End If
		GetDiyFieldValue=v
 End Function

'参数 isTemplate true 后台生成表单模板调用,channelid 模型id, id 编辑时的文章ID
Function GetInputForm(IsTemplate,ChannelID,FieldXML,FieldNode,FieldDictionary,id,KSUser,RS)
  Dim FNode,ClassID,Title,KeyWords,Author,Origin,Content,Verific,PhotoUrl,BigPhoto,Intro,FullTitle,ReadPoint,Province,City,UserDefineFieldArr,I,SelButton,MapMarker,PicUrls,ShowStyle,PageNum,DownSize,SizeUnit,DownPT,YSDZ,ZCDZ,JYMM,DownUrls
if IsObject(RS) And IsTemplate=false Then
	If Not RS.Eof Then
		     If KS.C_S(ChannelID,42) =0 And RS("Verific")=1 and KS.ChkClng(KS.U_S(KSUser.GroupID,1))=0 Then
			   RS.Close():Set RS=Nothing
			   KS.ShowTips "error",server.urlencode("本频道设置已审核" & KS.C_S(ChannelID,3) & "不允许修改!")
			   KS.Die ""
			 End If
		     ClassID  = RS("Tid")
			 Title    = RS("Title")
			 KeyWords = RS("KeyWords")
			 Author   = RS("Author")
			 Origin   = RS("Origin")
			 Select Case KS.ChkClng(KS.C_S(ChannelID,6))
			  case 1 
			    Content  = RS("ArticleContent"):FullTitle= RS("FullTitle")
				Province = RS("Province"):  City  = RS("City")
                Intro    = RS("Intro")
			  case 2 
			    PicUrls  = RS("PicUrls"):Content  = RS("PictureContent")
				ShowStyle= RS("ShowStyle"):PageNum  = RS("PageNum")
			  case 3
			    DownSize = RS("DownSize") : DownPT = RS("DownPT") :DownUrls=Split(RS("DownUrls"),"|")(2)
				YSDZ = RS("YSDZ") : ZCDZ = RS("ZCDZ") : JYMM = RS("JYMM") : BigPhoto=RS("BigPhoto")
				SizeUnit = Right(DownSize, 2):DownSize = Replace(DownSize, SizeUnit, "") : Content=RS("DownContent")
			 End Select
			 
			 Verific  = RS("Verific")
			 If Verific=3 Then Verific=0
			 PhotoUrl   = RS("PhotoUrl")
			 
			 
			 ReadPoint= RS("ReadPoint")
			 if KS.ChkClng(KS.C_S(ChannelID,6))<=2 Then
			  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='map']/showonuserform").text="1" Then	MapMarker=RS("MapMarker")
			 End If
				'自定义字段
				Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype!=0]")
				If diynode.length>0 Then
					Set FieldDictionary=KS.InitialObject("Scripting.Dictionary")
					For Each FNode In DiyNode
					   FieldDictionary.add lcase(FNode.SelectSingleNode("@fieldname").text),RS(FNode.SelectSingleNode("@fieldname").text)
					   If FNode.SelectSingleNode("showunit").text="1" Then
					   FieldDictionary.add lcase(FNode.SelectSingleNode("@fieldname").text) &"_unit",RS(FNode.SelectSingleNode("@fieldname").text&"_Unit")
					   End If
					Next
				End If
		   End If
		   SelButton=KS.C_C(ClassID,1)
		Else
		 If IsTemplate=false Then
		     If Not KS.IsNul(KS.C("UserName")) Then
		     Call KSUser.CheckMoney(ChannelID)
			 Author=KSUser.GetUserInfo("RealName")
			 Origin=LFCls.GetSingleFieldValue("SELECT top 1 CompanyName From KS_EnterPrise Where UserName='" & KSUser.UserName & "'")
			 End If
			 ClassID=KS.S("ClassID")
			 If ClassID="" Then ClassID="0"
			 If ClassID="0" Then
			 SelButton="选择栏目..."
			 Else
			 SelButton=KS.C_C(ClassID,1)
			 End If
			 ReadPoint=0 : Verific=0 : ShowStyle=4 : PageNum=10
			 YSDZ="http://" : ZCDZ="http://"
		 Else
		    ShowStyle="[#ShowStyle]":PageNum="[#PageNum]":PicUrls="[#PicUrls]":Title="[#Title]":FullTitle="[#FullTitle]":KeyWords="[#KeyWords]":Author="[#Author]":Origin="[#Origin]":Province="[#Province]":City="[#City]":Author="[#Author]":Intro="[#Intro]":Content="[#Content]":PhotoUrl="[#PhotoUrl]":BigPhoto="[#BigPhoto]":ReadPoint="[#ReadPoint]":Verific="[#Verific]":MapMarker="[#MapMarker]":DownSize="[#DownSize]":DownPT="[#DownPT]":YSDZ="[#YSDZ]":ZCDZ="[#ZCDZ]":JYMM="[#JYMM]":DownUrls="[#DownUrls]"
			'自定义字段
			Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype!=0]")
			If diynode.length>0 Then
					Set FieldDictionary=KS.InitialObject("Scripting.Dictionary")
					For Each FNode In DiyNode
					   FieldDictionary.add lcase(FNode.SelectSingleNode("@fieldname").text),"[#" & FNode.SelectSingleNode("@fieldname").text & "]"
					Next
			End If
			
		 End If
		End If
		%><table  width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="border">
 <tr class="title">
  <td colspan=2 align=center><%
	      If IsTemplate Then
		  Response.Write "[#PubTips]"
		  ElseIF ID<>0 Then
			  response.write "修改" & KS.C_S(ChannelID,3)
		  Else
		      response.write "发布" & KS.C_S(ChannelID,3)
		 End iF%></td>
 </tr>
 <%
For Each FNode In FieldNode
	    If KS.ChkClng(FNode.SelectSingleNode("fieldtype").text)<>0 Then
		    Response.Write KSUser.GetDiyField(ChannelID,FieldXML,FNode,FieldDictionary)
		Else
		 Dim XTitle:XTitle=FNode.SelectSingleNode("title").text
	     Select Case lcase(FNode.SelectSingleNode("@fieldname").text)
	       case "tid"
 %>
 <tr class="tdbg">
  <td width="12%"  height="25" align="center"><span><%=XTitle%>：</span></td>
  <td width="88%"><%
				If IsTemplate Then
				  Response.Write "[#SelectClassID]"
				Else
				 If KS.C("UserName")="" Then  '游客投稿
					response.write "[" & KS.GetClassNP(KS.S("ClassID")) & "] <a href=""Contributor.asp""><<重新选择>></a>"
					response.write "<input type=""hidden"" name=""ClassID"" value=""" & KS.S("ClassID") & """>"
				 Else
				 Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton) 
				 End If
				End If
			 %></td>
 </tr>
<%case "title"%>
 <tr class="tdbg">
    <td  height="25" align="center"><span><%=XTitle%>：</span></td>
    <td><input class="textbox" name="Title" type="text" id="Title" style="width:250px; " value="<%=Title%>" maxlength="100" /><span style="color: #FF0000">*</span></td>
 </tr>
<%case "fulltitle"%>
 <tr class="tdbg">
    <td  height="25" align="center"><span><%=XTitle%>：</span></td>
    <td><input class="textbox" name="FullTitle" type="text" style="width:250px; " value="<%=FullTitle%>" maxlength="100" /><span class="msgtips"> 完整标题，可留空</span></td>
 </tr>
<%case "keywords"%>
 <tr class="tdbg">
    <td height="25" align="center"><span><%=XTitle%>：</span></td>
    <td><input name="KeyWords"  class="textbox" type="text" id="KeyWords" value="<%=KeyWords%>" style="width:220px; " /><a href="javascript:void(0)" onclick="GetKeyTags()" style="color:#ff6600">【自动获取】</a> <span class="msgtips">多个关键字请用英文逗号(&quot;<span style="color: #FF0000">,</span>&quot;)隔开</span></td>
  </tr>
<%case "author"%>
<tr class="tdbg">
    <td  height="25" align="center"><span><%=XTitle%>：</span></td>
    <td height="25"><input name="Author" class="textbox" type="text" id="Author" style="width:220px; " value="<%=Author%>" maxlength="30" /> <span class="msgtips"><%=KS.C_S(ChannelID,3)%>的作者<span></td>
  </tr>
<%case "origin"%>
<tr class="tdbg">
   <td  height="25" align="center"><span><%=XTitle%>：</span></td>
   <td><input class="textbox" name="Origin" type="text" id="Origin" style="width:220px; " value="<%=Origin%>" maxlength="100" /> <span class="msgtips"><%=KS.C_S(ChannelID,3)%>的来源<span></td>
  </tr>
<%case "nature"%>
 <tr class="tdbg">
       <td height="25" align="center"><%=XTitle%>：</td>
       <td>
	   <%If Istemplate Then%>
	   类别:<select name='DownLB'>[#DownLBList]</select> 语言:<select name='DownYY' size='1'>[#DownYYList]</select>授权:<select name='DownSQ' size='1'>[#DownSQList]</select>
	   <%Else%>
	   类别:<select name='DownLB'><%=DownLBList%></select> 语言:<select name='DownYY' size='1'><%=DownYYList%></select>授权:<select name='DownSQ' size='1'><%=DownSQList%></select>
	   <%End If%>
	   <%
		Response.Write "大小:<input type='text' class='textbox' style='text-align:center' size=4 id='DownSize' name='DownSize' value='" & DownSize & "'> "
If Istemplate Then
      Response.Write "[#SizeUnit]"
Else
		If SizeUnit = "KB" Then
			Response.Write "<input name=""SizeUnit"" type=""radio"" value=""KB"" checked id=""kb""><label for=""kb"">KB</label> " & vbCrLf
			Response.Write "<input type=""radio"" name=""SizeUnit"" value=""MB"" id=""mb""><label for=""mb"">MB</label> " & vbCrLf
		Else
			Response.Write "<input name=""SizeUnit"" type=""radio"" value=""KB""  id=""kb""><label for=""kb"">KB</label> " & vbCrLf
			Response.Write "<input type=""radio"" name=""SizeUnit"" value=""MB"" checked id=""mb""><label for=""mb"">MB</label> " & vbCrLf
		End If
	End If%>                      
	</td>
</tr>
<%case "platform"%>
<tr class="tdbg">
     <td height="25" align="center"><%=XTitle%>：</td>
     <td><input class='textbox' type='text' size=70 name='DownPT' value="<%=DownPT%>"><br>
		 <font color='#808080'>平台选择 <%If Istemplate Then%>[#DownPTList]<%Else%><%=DownPTList%><%End If%></font></td>
</tr>
<%case "ysdz"%>
<tr class="tdbg">
   <td height="25" align="center"><%=XTitle%>：</td>
   <td><input class="textbox" name="YSDZ" type="text" value="<%=YSDZ%>" id="YSDZ" style="width:250px; " maxlength="100" /></td>
</tr>
<%case "zcdz"%>
<tr class="tdbg">
   <td height="25" align="center"><%=XTitle%>：</td>
   <td><input class="textbox" name="ZCDZ" type="text" value="<%=ZCDZ%>" id="ZCDZ" style="width:250px; " maxlength="100" /></td>
</tr>
<%case "jymm"%>
<tr class="tdbg">
   <td height="25" align="center"><%=XTitle%>：</td>
   <td><input class="textbox" name="JYMM" type="text" value="<%=JYMM%>" id="JYMM" style="width:250px; " maxlength="100" /></td>
</tr>
<%case "area"%>
<tr class="tdbg">
    <td  height="25" align="center"><span><%=XTitle%>：</span></td>
    <td><script>try{setCookie("pid",'<%=province%>');}catch(e){}</script>
							<script src="../plus/area.asp" language="javascript"></script>
							<script language="javascript">
							  <%if Province<>"" then%>
							  $('#Province').val('<%=province%>');
							  <%end if%>
							  <%if City<>"" Then%>
							  $('#City').val('<%=City%>');
							  <%end if%>
							</script>
	</td>
 </tr>
<%case "map"%>
<tr class="tdbg">
    <td height="25" align="center"><span><%=XTitle%>：</span></td>
    <td>经纬度：<input class="textbox" value="<%=MapMarker%>" type='text' name='MapMark' id='MapMark' /> <a href='javascript:void(0)' onclick='addMap()'> <img src='images/edit_add.gif' align='absmiddle' border='0'>添加电子地图标志</a>
	</td>
  </tr>
<%case "intro"%>
<%if KS.ChkClng(KS.C_S(ChannelID,6))=1 then%>
 <tr class="tdbg">
   <td  height="25" align="center"><span><%=XTitle%>：</span><br><input name='AutoIntro' type='checkbox' checked value='1'><font color="#FF0000">自动截取内容的200个字作为导读</font></td>
   <td><textarea class='textbox' name="Intro" style='width:95%;height:95px'><%=intro%></textarea></td>
  </tr>
<%end if%>
<%case "content"%>
<%  select case KS.ChkClng(KS.C_S(ChannelID,6))
     case 1
%>
<tr class="tdbg">
   <td><%=XTitle%>:<br><img src="images/ico.gif" width="17" height="12" /><font color="#FF0000">如果<%=KS.C_S(ChannelID,3)%>较长可以使用分页标签：[NextPage]</font></td>
   <td><%
	If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attachment']/showonuserform").text="1" Then
	%>
		<table border='0' width='100%' cellspacing='0' cellpadding='0'>
		<tr><td height='35' width=70 nowrap="nowrap">&nbsp;<strong><%=FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attachment']/title").text%>:</strong></td><td><iframe id='upiframe' name='upiframe' src='BatchUploadForm.asp?ChannelID=<%=ChannelID%>' frameborder=0 scrolling=no width='560' height='24'></iframe></td></tr>
		</table>
	 <%end if%>
		<textarea name="Content" ID="Content" style="display:none"><%=Server.HTMLEncode(Content)%></textarea>
        <script type="text/javascript">CKEDITOR.replace('Content', {width:"98%",height:"320",toolbar:"Basic",filebrowserBrowseUrl :"../editor/ksplus/SelectUpFiles.asp",filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script>
	</td>
</tr>
<%case 2%>
	<tr class="tdbg">
		<td height="35" align="center"><span>显示样式：</span></td>
		<td><%if IsTemplate Then%>
		[#ShowStyle]
		<%Else%><table width='80%'><tr><td>
					<input type='radio' onClick="$('#pagenums').hide();" name='showstyle' value='4'<%If ShowStyle="4" Then response.Write " checked"%>><img src='../images/default/p4.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td><td>
					<input type='radio' onClick="$('#pagenums').hide();" name='showstyle' value='1'<%If ShowStyle="1" Then response.Write " checked"%>><img src='../images/default/p1.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td>
		   <td><input type='radio' onClick="$('#pagenums').show();" name='showstyle' value='2'<%If ShowStyle="2" Then Response.Write " checked"%>><img src='../images/default/p2.gif' title='当图片组只有一张图片时无效,设置此样式无效!'></td><td><input type='radio' onClick="$('#pagenums').show();" name='showstyle' value='3'<%If ShowStyle="3" Then Response.Write " checked"%>><img src='../images/default/p3.gif'>
		   </td><td><input type='radio' onClick="$('#pagenums').hide();" name='showstyle' value='5'<%If ShowStyle="5" Then Response.Write " checked"%>><img src='../images/default/p5.gif'>
		   </td></tr></table><div style="margin:5px" id="pagenums"
			<%If ShowStyle="1" or ShowStyle="4" Then Response.Write " style='display:none'"%>
			>每页显示<input type="text" name="pagenum" value="<%=PageNum%>" style="text-align:center;width:30px">张</div>
		<%End If%>
		</td>
  </tr>
 <tr class="tdbg">
       <td height="40" align="center" nowrap><span><%=XTitle%>：</span></td>
       <td>
	    <table>
		 <tr>
		  <td><div class="pn" style="margin: -6px 0px 0 0;">
			 <span id="spanButtonPlaceholder"></span>
			</div></td>
		 <td>
		 <button type="button"  class="pn" onClick="OnlineCollect()" style="margin: -6px 0px 0 0;"><strong>网上地址</strong></button><%if IsTemplate Then%>
		 [#ShowSelectTK]
	   <%ElseIf KS.C("UserName")<>"" Then%>
		 <button type="button"  class="pn" onClick="AddTJ();" style="margin: -6px 0px 0 0;"><strong>图片库...</strong></button>
	   <%End If%>
		 </td>
		 </tr>
		</table>

		<label><input type="checkbox" name="AddWaterFlag" value="1" onClick="SetAddWater(this)" checked="checked"/>图片添加水印</label>
		<div id="divFileProgressContainer"></div>
	    <div id="thumbnails"></div>
		<input type='hidden' name='PicUrls' id='PicUrls' value="<%=PicUrls%>">
	</td>
</tr>
<%case 3%>
<tr class="tdbg">
     <td align="center"><%=XTitle%>：</td>
         <td align="center"><textarea name="Content" style="display:none"><%=Server.HTMLEncode(Content)%></textarea>
         <script type="text/javascript">
			CKEDITOR.replace('Content', {width:"99%",height:"200px",toolbar:"Basic",filebrowserBrowseUrl :"../editor/ksplus/SelectUpFiles.asp",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
		</script> 
   </td>
</tr>
<%end select%>
<%case "address"%>
 <tr class="tdbg">
     <td height="25" align="center"><%=KS.C_S(ChannelID,3)%>地址：</td>
     <td valign="top">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td  width="275"><input type="text" class="textbox" name='DownUrlS' id='DownUrlS' value='<%=DownUrls%>' style="width:250px; "> <span style="color: #FF0000">*</span>
                 </td>
					<%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='uploadsoft']/showonuserform").text="1" Then%>
					<td><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_upfile.asp?type=UpByBar&channelid=<%=ChannelID%>' frameborder="0" scrolling="no" width='280' height='25'></iframe></td>
					<%end if%>
			</tr>
		</table>
   </td>
</tr>
<%
case "photourl"
%>
<% select case KS.ChkClng(KS.C_S(ChannelID,6))
 case 1%>
<tr class="tdbg">
    <td height="25" align="center"><%=XTitle%>：</td>
    <td height="25"> <input name='PhotoUrl' style="float:left;width:230px;margin-top:3px;margin-right:4px" type='text' id='PhotoUrl' value="<%=PhotoUrl%>" size='40'  class="textbox"/>
			<%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='uploadphoto']/showonuserform").text="1" Then%>	
              <iframe  id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?Type=Pic&ChannelID=<%=ChannelID%>' frameborder="0" scrolling="No"  width='340' height='30'></iframe>
			<%end if%>   
	</td>
</tr>
<%case 2%>
<tr class="tdbg">
     <td height="35" align="center"><span><%=XTitle%>：</span></td>
     <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
									 <tr>
									 <%if IsTemplate Then%>
									 [#ShowPhotoUrl]
									 <%
									 Else
									   If KS.C("UserName")="" Then%>
									  <td width="240"><input class="textbox" name='PhotoUrl' value="<%=PhotoUrl%>" type='text' style="width:230px;" id='PhotoUrl' maxlength="100" />
									  </td>
									 <%Else%>
									  <td width="340"><input class="textbox" name='PhotoUrl' value="<%=PhotoUrl%>" type='text' style="width:230px;" id='PhotoUrl' maxlength="100" />
                                          
                                          <input class="uploadbutton1" type='button' name='Submit3' value='选择图片' onClick="OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("选择" & KS.C_S(ChannelID,3))%>&ChannelID=<%=ChannelID%>',500,360,window,document.myform.PhotoUrl);" />
								      </td>
									 <%End If
									 End If%>
									  <td>
									  <iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_upfile.asp?channelid=<%=ChannelID%>&Type=Pic' frameborder=0 scrolling=no width='250' height='30'> </iframe>
									  </td>
									 </tr>
									 </table><%if IsTemplate Then%>
										[#ShowSetThumb]
										<%elseif action<>"Edit" Then%>
										 <label><input type='checkbox' name='autothumb' id='autothumb' value='1' checked>使用图集的第一幅图</label>
										<%end if%>
				 </td>
	</tr>
<%case 3%>
<tr class="tdbg">
      <td height="25" align="center"><%=XTitle%>：</td>
      <td><input class="textbox"  name="PhotoUrl" value="<%=PhotoUrl%>" type="text" id="PhotoUrl" style="width:250px; float:left;margin-top:3px;margin-right:3px;" maxlength="100" /><input type="hidden" name="BigPhoto" id="BigPhoto" value="<%=BigPhoto%>"/>
		<%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='uploadphoto']/showonuserform").text="1" Then%>
		<iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_upfile.asp?channelid=<%=ChannelID%>&Type=Pic' frameborder=0 scrolling=no width='250' height='30'> </iframe>
		<%end if%>
 </td>
</tr>
<%
end select
case "chargeoption"%>
<tr class="tdbg">
        <td height="25" align="center"><span>阅读<%=KSUser.GetModelCharge(channelid)%>：</span></td>
         <td height="25"><input type="text" style="text-align:center" name="ReadPoint" class="textbox" value="<%=ReadPoint%>" size="6"><%=KS.Setting(46)%> <span class="msgtips">如果免费阅读请输入“<font color=red>0</font>”</span></td>
   </tr>
<%case "picturecontent"%>
<tr class="tdbg">
          <td align="center"><%=XTitle%>：<br /></td>
          <td align="center">
           <textarea style="display:none;" name="Content" id="Content"><%=Server.HTMLEncode(Content)%></textarea>
			<script type="text/javascript">
			CKEDITOR.replace('Content', {width:"98%",height:"150px",toolbar:"Basic",filebrowserBrowseUrl :"../editor/ksplus/SelectUpFiles.asp",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
			</script>
		 </td>
</tr>
<%
   End Select
 End IF
Next


If IsTemplate Then
%>[#ShowQuestionAndVerify]
<%
ElseIf KS.C("UserName")="" Then
Call PubQuestion()
%>
    <tr class="tdbg">
			<td  height="25" align="center"><span>验证码：</span></td>
			<td>
			 <script type="text/javascript">writeVerifyCode('<%=KS.Setting(3)%>',1,'textbox')</script>
			</td>
	</tr>
<%
End If
%>


 <tr class="tdbg">
   <td height="40"></td><td><button class="pn" id="submit1" type="submit" onclick="return(CheckForm())"><strong>OK, 保 存</strong></button>&nbsp;<%if IsTemplate Then 
   Response.Write "[#Status]" 
   Elseif id<>0 Then
		     If Verific<>1 Then
			  if Verific=2 Then
		       response.write "<label style='color:#999999'><input type='checkbox' name='status' value='2' onclick='if(!this.checked){return(confirm(""确定立即投稿发布吗?""));}' checked>放入草稿</label>"
			  else
		       response.write "<label style='color:#999999'><input type='checkbox' name='status' onclick='if(!this.checked){return(confirm(""确定立即投稿发布吗?""));}' value='2'>放入草稿</label>"
			  end if
			 Else
			  response.write "<input type=""hidden"" name=""okverific"" value=""1""><input type=""hidden"" name=""verific"" value=""1"">"
			 End If
	ElseIf KS.C("UserName")<>"" Then
		    response.write "<label style='color:#999999'><input type='checkbox' name='status' value='2'>放入草稿</label>"
   End If%> </td>
 </tr>
</table>
<br/><%
End Function
%>