<%@ Language="VBSCRIPT" codepage="65001" %>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../KS_Cls/Kesion.KeyCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"
Dim SupplyPayPoint:SupplyPayPoint=1    '默认查看供求信息扣一个点券，设置为0不扣点！

Dim KS:Set KS=New PublicCls
If KS.IsNul(Request.ServerVariables("HTTP_REFERER")) Then KS.Die "error"

Dim Action
Action=KS.S("Action")
Select Case Action
 Case "paySupplyShow" paySupplyShow
 Case "DelPhoto"  DelPhoto
 Case "SetAttributeFields" SetAttributeFields
 Case "Ctoe" CtoE
 Case "GetTags" GetTags
 Case "GetRelativeItem" GetRelativeItem
 Case "GetClassOption" GetClassOption
 Case "GetFieldOption" GetFieldOption
 Case "GetModelAttr" GetModelAttr
 Case "SpecialSubList" SpecialSubList
 Case "GetArea" GetArea
 Case "GetFunc" GetFunc
 Case "GetSchool" GetSchool
 Case "AddFriend" AddFriend
 Case "MessageSave" MessageSave
 Case "CheckMyFriend" CheckMyFriend
 Case "SendMsg" SendMsg
 Case "SearchUser" SearchUser
 Case "CheckLogin" CheckLogin
 Case "relativeDoc" relativeDoc
 Case "getModelType" getModelType
 Case "addCart" addShoppingCart
 Case "GetPackagePro" GetPackagePro
 Case "GetSupplyContact" GetSupplyContact
 Case "HitsGuangGao" HitsGuangGao
 Case "GetClubBoardOption" GetClubBoardOption
 Case "getclubboard" GetClubboard
 Case "GetClubPushModel" GetClubPushModel
 Case "getclubboardcategory" getclubboardcategory
 Case "getonlinelist" getonlinelist
End Select
Set KS=Nothing
CloseConn()

Sub SetAttributeFields()
  Dim ChannelID:ChannelID=KS.ChkClng(KS.S("Channelid"))
  If ChannelID=0 Then KS.Die ""
  If KS.C_S(ChannelID,6)="1" Then
  %>
<tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eIsVideo' value='1'></td>
				<td class='clefttitle' align='right'><strong><%=escape("是否带视频")%>:</strong></td>
				<td><input name='IsVideo' type='radio' value='1'> 是  <input name='IsVideo' type='radio' value='0' checked> 否</td>
		  </tr>	  <%
  ElseIf KS.C_S(ChannelID,6)="3" Then
  '取得下载参数
		 Dim DownLBList, DownYYList, DownSQList, DownPTList, RSP, DownLBStr, LBArr, YYArr, SQArr, PTArr, DownYYStr, DownSQStr, DownPTStr
		  Set RSP = Server.CreateObject("Adodb.RecordSet")
		  RSP.Open "Select top 1 * From KS_DownParam Where ChannelID=" & ChannelID, conn, 1, 1
		  If Not RSP.Eof Then
		   DownLBStr = RSP("DownLB"):DownYYStr = RSP("DownYY"): DownSQStr = RSP("DownSQ"): DownPTStr = RSP("DownPT")
		  End If
		  RSP.Close:Set RSP = Nothing
		  '下载类别
		  LBArr = Split(DownLBStr, vbCrLf)
		  For I = 0 To UBound(LBArr)
			DownLBList = DownLBList & "<option value='" & escape(LBArr(I)) & "'>" & escape(LBArr(I)) & "</option>"
		  Next
		  '下载语言
		  YYArr = Split(DownYYStr, vbCrLf)
		  For I = 0 To UBound(YYArr)
			DownYYList = DownYYList & "<option value='" & escape(YYArr(I)) & "'>" & escape(YYArr(I)) & "</option>"
		  Next
		'下载授权
		  SQArr = Split(DownSQStr, vbCrLf)
		  For I = 0 To UBound(SQArr)
			DownSQList = DownSQList & "<option value='" & escape(SQArr(I)) & "'>" & escape(SQArr(I)) & "</option>"
		  Next
		'下载平台
		  PTArr = Split(DownPTStr, vbCrLf)
		  For I = 0 To UBound(PTArr)
			DownPTList = DownPTList & "<a href='javascript:SetDownPT(""" & PTArr(I) & """)'>" & PTArr(I) & "</a>/"
		  Next
		  %>
		  <tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eDownServer' value='1'></td>
				<td class='clefttitle' align='right'><strong><%=escape("设置服务器")%>:</strong></td>
				<td><select name="DownServer"><option value='0'>↓不使用下载服务器↓</option><%
				Dim rsobj
			Set rsobj = conn.Execute("SELECT downid,DownloadName,depth,rootid FROM KS_DownSer WHERE depth=0 And ChannelID="& ChannelID)
			Do While Not rsobj.EOF
				 response.write escape("<option value=""" & rsobj("downid") & """>" & rsobj(1) & "</option>")
				rsobj.movenext
			Loop
			rsobj.Close:Set rsobj = Nothing
				
				%></select></td>
		  </tr>
		  <tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eDownLB' value='1'></td>
				<td class='clefttitle' align='right'><strong><%=escape("下载类别")%>:</strong></td>
				<td><select name="DownLB"><%=DownLBList%></select></td>
		  </tr>	
		  <tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eDownYY' value='1'></td>
				<td class='clefttitle' align='right'><strong><%=escape("语言")%>:</strong></td>
				<td><select name="DownYY"><%=DownYYList%></select></td>
		  </tr>	
		  <tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eDownSQ' value='1'></td>
				<td class='clefttitle' align='right'><strong><%=escape("授权方式")%>:</strong></td>
				<td><select name="DownSQ"><%=DownSQList%></select></td>
		  </tr>	
		  <tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eDownPT' value='1'></td>
				<td class='clefttitle' align='right' nowrap><strong><%=escape("运行平台")%>:</strong></td>
				<td><input type="text" size='40' name='DownPT' id='DownPT' class='textbox'><br/><%=DownPTList%></td>
		  </tr>	
		  <tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eYSDZ' value='1'></td>
				<td class='clefttitle' align='right' nowrap><strong><%=escape("演示地址")%>:</strong></td>
				<td><input type="text" size='40' name='YSDZ' id='YSDZ' class='textbox'></td>
		  </tr>	
		  <tr class='tdbg'> 
				<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eZCDZ' value='1'></td>
				<td class='clefttitle' align='right' nowrap><strong><%=escape("注册地址")%>:</strong></td>
				<td><input type="text" size='40' name='ZCDZ' id='ZCDZ' class='textbox'></td>
		  </tr>	
		  <%
  End If
  Dim RS:Set RS=server.CreateObject("ADODB.RECORDSET")
  RS.Open "Select FieldName,Title,FieldType,Options,Width From KS_Field Where FieldType<>0 and ChannelID="& ChannelID &" Order By OrderID ,FieldID",conn,1,1
  If Not RS.EOf Then
      %>
		<tr><td colspan=3 align='center' class='clefttitle' style="font-weight:bold;color:blue;height:20px;text-align:center">========<%=escape("以下列出自定义字段")%>===========</td></tr>
	<%
	  Do While Not RS.Eof 
		%>
		<tr class='tdbg'> 
			<td class='clefttitle' height='25' align='center'><input type='checkbox' name='e<%=rs(0)%>' value='1'></td>
			<td class='clefttitle' nowrap align='right'><strong><%=escape(rs(1))%>:</strong></td>
			<td><%
			Dim O_Arr,O_Len,O_Value,O_Text,F_V,K,BrStr
			select case rs(2) 
			  case 3,11
			   KS.Echo "<select class=""upfile"" style=""width:" & rs(4) & "px"" name=""" & rs(0) & """>"
			   O_Arr=Split(RS(3),vbcrlf): O_Len=Ubound(O_Arr)
				 For K=0 To O_Len
					If O_Arr(K)<>"" Then
							 F_V=Split(O_Arr(K),"|")
							 If Ubound(F_V)=1 Then  O_Value=F_V(0):O_Text=F_V(1) Else  O_Value=F_V(0):O_Text=F_V(0)
							KS.Echo Escape("<option value=""" & O_Value& """>" & O_Text & "</option>")
					End If
				 Next
			   KS.Echo "</select>"
			 case 6,7
			   O_Arr=Split(RS(3),vbcrlf): O_Len=Ubound(O_Arr)
						   For K=0 To O_Len
							   F_V=Split(O_Arr(K),"|")
							   If O_Arr(K)<>"" Then
							    If Ubound(F_V)=1 Then O_Value=F_V(0):O_Text=F_V(1) Else	O_Value=F_V(0):O_Text=F_V(0)
								If rs(2) = 6 Then
							     KS.Echo escape("<input type=""radio"" name=""" & RS(0) & """ value=""" & O_Value& """>" & O_Text)&BrStr
								Else
							    KS.Echo escape("<input type=""checkbox"" name=""" & RS(0) & """ value=""" & O_Value& """>" & O_Text)&BrStr
								End If
							 End If
						   Next
			  case else
			%><input type="text" size='40' name='<%=rs(0)%>' id='<%=rs(0)%>' class='textbox'/>
			<%
			end select 
			%>
			</td>
		 </tr>
		<%
	  RS.MoveNext
	  Loop
  End If
  RS.Close :Set RS=Nothing
End Sub

Sub getModelType()
 Dim ChannelID:ChannelID=KS.ChkClng(Request("channelid"))
 If ChannelID<>0 Then KS.Echo KS.C_S(Channelid,6)
End Sub




'取中文首字母
Sub Ctoe()
 Dim FolderName:FolderName=KS.DelSQL(UnEscape(Request("FolderName")))
 Dim CE:Set CE=New CtoECls
 Response.Write Escape(CE.CTOE(FolderName))
 Set CE=Nothing
End Sub

'取关键词tags
Sub GetTags()
 Dim Text:Text=UnEscape(Request("Text"))
 If Text<>"" Then
     Dim MaxLen:MaxLen=KS.ChkClng(KS.S("MaxLen"))
	 Dim WS:Set WS=New Wordsegment_Cls
	 Response.Write Escape(WS.SplitKey(text,4,MaxLen))
	 Set WS=Nothing
 End If
End Sub


'相关信息
Sub GetRelativeItem()
 Dim Key:Key=KS.DelSql(UnEscape(request("Key")))
 Dim Rtitle:rtitle=lcase(KS.G("rtitle"))
 Dim RKey:Rkey=lcase(KS.G("Rkey"))
 Dim ChannelID:ChannelID=KS.ChkClng(KS.S("Channelid"))
 Dim ID:ID=KS.ChkClng(KS.G("ID"))
 Dim Param,RS,SQL,k,SqlStr
 If Key<>"" Then
   If (Rtitle="true" Or RKey="true") Then
	 If Rtitle="true" Then
	   param=Param & " title like '%" & key & "%'"
	 end if
	 If Rkey="true" Then
	   If Param="" Then
	     Param=Param & " keywords like '%" & key & "%'"
	   Else
	     Param=Param & " or keywords like '%" & key & "%'"
	   End If
	 End If
 Else
    Param=Param & " keywords like '%" & key & "%'"
 End If
End If

 
 If Param<>"" Then 
  	Param=" where verific=1 and InfoID<>" & id & " and (" & param & ")"
 else
    Param=" where verific=1 and  InfoID<>" & id
 end if
 
  If ChannelID<>0 Then Param=Param & " and ChannelID=" & ChannelID


 SqlStr="Select top 30 ChannelID,InfoID,Title From KS_ItemInfo " & Param & " order by id desc"
 Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open SqlStr,conn,1,1
 If Not RS.Eof Then
  SQL=RS.GetRows(-1)
 End If
 RS.Close
 Set RS=Nothing
 If IsArray(SQL) Then
	 For k=0 To Ubound(SQL,2)
	   Response.Write "<option value='" & SQL(0,K) & "|" & SQL(1,K) & "'>" & SQL(2,K) & "</option>" 
	 Next
 End If
End Sub


'检查是否登录
Sub CheckLogin()
  If KS.C("UserName")="" Then KS.Echo "false" Else  KS.Echo "true"
End Sub

'取栏目选项
Sub GetClassOption()
 Dim ChannelID:ChannelID=KS.ChkCLng(Request.Querystring("ChannelID"))
   Dim KSUser:Set KSUser=New UserCls
		Dim Node,K,SQL,NodeText,Pstr,TJ,SpaceStr,TreeStr,nbsp
		KS.LoadClassConfig()
		If ChannelID<>0 Then Pstr=" and @ks12=" & channelid & ""
		For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&"]")
		  SpaceStr="" 
		 If (Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,KSUser.GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3) ) Or Node.SelectSingleNode("@ks20").text="0"  Then
		 Else
			  TJ=Node.SelectSingleNode("@ks10").text
			  If TJ>1 Then
				 For k = 1 To TJ - 1
					SpaceStr = SpaceStr & "──" 
				 Next
			  End If
			  TreeStr = TreeStr & "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & SpaceStr & Node.SelectSingleNode("@ks1").text & " </option>"
		 End If 
		Next
		Set KSUser=Nothing
 KS.Die Escape(TreeStr)
End Sub


Sub SpecialSubList()
	  Dim ClassID, RS,SpecialXML,Node
	  ClassID=KS.ChkClng(Request.QueryString("ClassID"))
	  If ClassID=0 Then Exit Sub
	  Set RS=Conn.Execute("Select SpecialID,SpecialName from KS_Special Where ClassID=" & ClassID & " Order BY SpecialAddDate Desc")
	  If Not RS.Eof Then Set SpecialXML=KS.RsToXml(RS,"row","xmlroot")
	  RS.Close:Set RS=Nothing
	  If IsObject(SpecialXml) Then
	  	For Each node in SpecialXml.DocumentElement.SelectNodes("row")
		  KS.Echo Escape("<div><img src=""images/folder/Special.gif"" align='absmiddle'>")
          KS.Echo Escape("<a href=""#"">"  & Trim(Node.SelectSingleNode("@specialname").text) & "</a><input type='checkbox' onclick=""set(" & Node.SelectSingleNode("@specialid").text & ",'" & Node.SelectSingleNode("@specialname").text & "');"" value='" & Node.SelectSingleNode("@specialid").text & "'></div>")
	    Next
		 Set SpecialXml=Nothing
      End If
End Sub

Sub GetFieldOption()
    Dim Node,ChannelID
	ChannelID=KS.ChkClng(Request.QueryString("ChannelID"))
	If ChannelID=0 Then Exit Sub
	Dim FieldXML,FieldNode,KSUser
	Set KSUser=New UserCls
	Call KSUser.LoadModelField(ChannelID,FieldXML,FieldNode)
	If Not IsObject(FieldXML) Then Exit Sub
	if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				  Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype!=0&&fieldtype!=13]")
				  If DiyNode.Length>0 Then
						For Each Node In DiyNode
							KS.Echo "<li class='diyfield' title=""" & Node.SelectSingleNode("title").text &""" onclick=""FieldInsertCode('" & Node.SelectSingleNode("@fieldname").text & "'," & Node.SelectSingleNode("fieldtype").text & ")"">" & Node.SelectSingleNode("title").text & "</li>"
							if Node.SelectSingleNode("showunit").text="1" then
							KS.Echo "<li class='diyfield' style='color:#ff3300' title=""" & Node.SelectSingleNode("title").text &""" onclick=""FieldInsertCode('" & Node.SelectSingleNode("@fieldname").text & "_unit'," & Node.SelectSingleNode("fieldtype").text & ")"">“" & Node.SelectSingleNode("title").text & "”单位</li>"
							end if
						Next
				  End If
	End If
	Set KSUser=Nothing
End Sub

Sub GetModelAttr()
    Dim Node,ChannelID
	ChannelID=KS.ChkClng(Request.QueryString("ChannelID"))
	If ChannelID=0 Then Exit Sub
	Dim FieldXML,FieldNode,KSUser,Attr
	Attr=Request("Attr")
	Set KSUser=New UserCls
	Call KSUser.LoadModelField(ChannelID,FieldXML,FieldNode)
	If Not IsObject(FieldXML) Then Exit Sub
	if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				  Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype=13]")
				  If DiyNode.Length>0 Then
						For Each Node In DiyNode
						  If KS.FoundInArr(lcase(Attr&""),lcase(Node.SelectSingleNode("@fieldname").text),"|") Then
							KS.Echo "<label style='color:brown'><input type='checkbox' checked name='attr' value='" & Node.SelectSingleNode("@fieldname").text & "'>" & Node.SelectSingleNode("title").text & "</label>"
						  Else
							KS.Echo "<label style='color:brown'><input type='checkbox' name='attr' value='" & Node.SelectSingleNode("@fieldname").text & "'>" & Node.SelectSingleNode("title").text & "</label>"
						  End If
						Next
				  End If
	End If
	Set KSUser=Nothing
End Sub

'取得ajax选项
sub GetArea()
Dim Parentid:parentid=KS.ChkClng(Request("parentid"))
Dim Param:Param="where parentid=0"
if parentid<>0 then param=" where parentid=" & parentid
If Parentid<>0 Then
  response.write escape("<div><a href='javascript:void(0)' onclick='goBack()'>返回上一级</a></div>")
End If
Dim ors : set ors=Conn.Execute("select ID,City FROM KS_Province " & Param & " order by orderid")
 do while not ors.eof
  if parentid=0 then
  response.write escape("<label><input type='checkbox' name='province' onclick='loadSecond(" & ors(0) & ",""" & ors(1) & """)' value='" & ors(1) & "'>" & ors(1) &" </label>")
  else
  response.write escape("<label><input type='checkbox' name='province' onclick='addPreItem()' value='" & ors(1) & "'>" & ors(1) &" </label>")
  end if
 ors.movenext
 loop
 ors.close
 set ors=nothing

end sub

'取得院校
sub GetSchool()
Dim Parentid:parentid=KS.ChkClng(Request("parentid"))
Dim Param:Param="where parentid=0"
Dim sqlstr
if parentid<>0 then
  sqlstr="select id,schoolname from ks_job_school where provinceid=" & parentid & " order by orderid,id"
  response.write escape("<div><a href='javascript:void(0)' onclick='goBack()'>返回上一级</a></div>")
else
  sqlstr="select id,city from ks_province where parentid=0 order by orderid,id"
End If
Dim ors : set ors=Conn.Execute(sqlstr)
 do while not ors.eof
  if parentid=0 then
  response.write escape("<label><input type='checkbox' name='province' onclick='loadSecond(" & ors(0) & ",""" & ors(1) & """)' value='" & ors(1) & "'>" & ors(1) &" </label>")
  else
  response.write escape("<label><input type='checkbox' name='province' onclick='addPreItem()' value='" & ors(0) & "'>" & ors(1) &" </label>")
  end if
 ors.movenext
 loop
 ors.close
 set ors=nothing

end sub

'取得职能
sub GetFunc()
Dim Parentid:parentid=KS.ChkClng(Request("parentid"))
Dim Param:Param="where parentid=0"
if parentid<>0 then param=" where parentid=" & parentid
If Parentid<>0 Then
  response.write escape("<div><a href='javascript:void(0)' onclick='goBack()'>返回上一级</a></div>")
End If
Dim ors : set ors=Conn.Execute("select ID,hymc FROM KS_Job_hyzw " & Param & " order by orderid")
 do while not ors.eof
  if parentid=0 then
  response.write escape("<label><input type='checkbox' name='province' onclick='loadSecond(" & ors(0) & ",""" & ors(1) & """)' value='" & ors(1) & "'>" & ors(1) &" </label>")
  else
  response.write escape("<label><input type='checkbox' name='province' onclick='addPreItem()' value='" & ors(1) & "'>" & ors(1) &" </label>")
  end if
 ors.movenext
 loop
 ors.close
 set ors=nothing

end sub

'请求加为好友
Sub AddFriend()
 If KS.C("UserName")="" Then KS.Echo "nologin" : Response.End
 Dim UserName:UserName=KS.DelSQL(UnEscape(Request("UserName")))
 Dim Message:Message=KS.DelSQL(UnEscape(Request("Message")))
 If Len(Message)>255 Then 
   KS.Echo escape("附言字数太多,最多只能输入255个字符!")
   exit sub
 End If
 If UserName="" Then KS.Echo escape("没有输入好友名称!") : Exit Sub
 call saveFriend(username,message,0)
 KS.Echo "success"
End Sub
'检查是否好友
Sub CheckMyFriend()
 If KS.C("UserName")="" Then KS.Echo "nologin" : Response.End
 Dim UserName:UserName=KS.DelSQL(UnEscape(Request("UserName")))
 Dim RS:Set RS=Conn.Execute("Select Top 1 accepted from KS_Friend Where UserName='" & KS.C("UserName") & "' and friend='" & username & "'")
 If rs.eof then
  KS.Echo "false"
 Else
  If rs(0)="1" then
   KS.Echo "true"
  Else
   KS.Echo "verify"
  End If
 End If
 RS.Close:Set RS=Nothing
End Sub

sub saveFriend(username,message,accepted)
		dim incept,i,sql,rs
		incept=KS.R(username)
		incept=split(incept,",")
		set rs=server.createobject("adodb.recordset")
		for i=0 to ubound(incept)
			sql="select top 1 UserName from KS_User where UserName='"&incept(i)&"'"
			set rs=Conn.Execute(sql)
			if rs.eof and rs.bof then
				rs.close:set rs=nothing
				KS.Echo escape("系统没有（"&incept(i)&"）这个用户，操作未成功。")
				Set KS=Nothing
				Response.End
			end if
			set rs=Nothing
			
			if KS.C("UserName")=Trim(incept(i)) then
			   KS.Echo escape("不能把自已添加为好友。")
			   Set KS=Nothing
			   Response.End
			end if
			
			sql="select top 1 id,friend,accepted from KS_Friend where username='"&KS.C("UserName")&"' and  friend='"&incept(i)&"'"
			set rs=Conn.Execute(sql)
			if rs.eof and rs.bof then
				sql="insert into KS_Friend (username,friend,addtime,flag,message,accepted) values ('"&KS.C("UserName")&"','"&Trim(incept(i))&"',"&SqlNowString&",1,'" & replace(message,"'","") & "'," & accepted & ")"
				set rs=Conn.Execute(sql)
			else
			    if rs("accepted")=0 then
				  conn.execute("update ks_friend set message='" & replace(message,"'","") & "' where id=" & rs("id"))
				end if
			end if
		
		next
		set rs=nothing
end sub
'发送短消息
Sub SendMsg()
     If Request.ServerVariables("HTTP_REFERER")="" Then KS.Die "error!"
     If KS.C("UserName")="" Then Response.End
	 Dim UserName:UserName=KS.DelSQL(UnEscape(Request("UserName")))
	 Dim Message:Message=KS.DelSQL(UnEscape(Request("Message")))
	 If Len(Message)>255 Then 
	   KS.Echo escape("附言字数太多,最多只能输入255个字符!")
	   exit sub
	 End If

     Call KS.SendInfo(UserName,KS.C("UserName"),KS.Gottopic(Message,100),Message)
	 KS.Echo "success"
End Sub

'搜索好友
Sub SearchUser()
 Dim Page:Page=KS.ChkClng(Request("Page")) : If Page= 0 Then Page=1
 Dim Province:Province=KS.DelSQL(UnEscape(Request("Province")))
 Dim City:City=KS.DelSQL(UnEscape(Request("City")))
 Dim Birth_Y:Birth_Y=KS.ChkClng(Request("Birth_Y"))
 Dim Birth_M:Birth_M=KS.ChkClng(Request("Birth_M"))
 Dim Birth_D:Birth_D=KS.ChkClng(Request("Birth_D"))
 Dim RealName:RealName=KS.DelSQL(UnEscape(Request("RealName")))
 Dim Sex:Sex=KS.DelSQL(UnEscape(Request("Sex")))
 Dim RS:Set RS=Server.CreateObject("Adodb.recordset")
 Dim Param,SQLStr,XML,Node,totalPut,MaxPerPage,TotalPage,N
 MaxPerPage=10
 Param="Where locked=0"
 If Province<>"" Then Param=Param &" and Province='"& Province & "'"
 If City<>"" Then Param=Param & " and city='" & city & "'"
 If Sex<>"" Then Param=Param & " and sex='" & Sex & "'"
 If RealName<>"" Then Param=Param & " and realname like '%" & RealName & "%'"
 If Birth_Y<>0 Then Param=Param & " and year(birthday)=" & Birth_Y & ""
 If Birth_M<>0 Then Param=Param & " and month(birthday)=" & Birth_m & ""
 If Birth_D<>0 Then Param=Param & " and day(birthday)=" & Birth_d & ""

 
 SQLStr="Select userid,username,realname,sex,birthday,province,city,userface,isonline from ks_user " & param & " order by userid desc"
 'response.write sqlstr
 RS.Open SQLStr,conn,1,1
 If RS.Eof And RS.Bof Then
   RS.Close: Set RS=Nothing
    KS.Echo Escape("<div style='text-align:center'>对不起,找不到您要查找的用户!请更换查询条件,重新检索!</div>")
 Else
    totalPut = Conn.Execute("Select Count(*) From KS_User " & Param)(0)
	If Page < 1 Then	Page = 1
	If (totalPut Mod MaxPerPage) = 0 Then
		TotalPage = totalPut \ MaxPerPage
	Else
		TotalPage = totalPut \ MaxPerPage + 1
	End If
	
	If Page > 1  and (Page - 1) * MaxPerPage < totalPut Then
		RS.Move (Page - 1) * MaxPerPage
	Else
		Page = 1
	End If
	Set XML=KS.ArrayToXML(RS.GetRows(MaxPerPage),RS,"row","")
	RS.Close : Set RS=Nothing
	If IsObject(XML) Then
	  Dim user_face,UserName
	 For Each Node In XML.DocumentElement.SelectNodes("row")
	  user_face=node.selectsinglenode("@userface").text
	  If user_face="" then 
	    if node.selectSingleNode("@sex").text="男" then  user_face="images/face/0.gif" else user_face="images/face/girl.gif"
	  End If
	  If lcase(left(user_face,4))<>"http" then user_face=KS.Setting(2) & "/" & user_face
      username=Node.selectsinglenode("@username").text
	  KS.Echo "<li>"
	  KS.Echo "<table border='0' width='100%'>"
	  KS.Echo "<tr><td class='face'> <a href='" & KS.GetSpaceUrl(Node.SelectSingleNode("@userid").text) & "' target='_blank'><img src='" & user_face & "' alt='" & username & "' /></a></td>"
	  KS.Echo " <td align='left' class='realname'>"
	  KS.Echo   Escape(Username & "(" & Node.SelectSingleNode("@realname").text & ")")
	  if isdate(Node.SelectSingleNode("@birthday").text) then
	  KS.Echo Escape(" <br />性别：" & Node.SelectSingleNode("@sex").text & "　出生：" & formatdatetime(Node.SelectSingleNode("@birthday").text,2))
	  else
	  KS.Echo Escape(" <br />性别：" & Node.SelectSingleNode("@sex").text & "　出生：" & Node.SelectSingleNode("@birthday").text)
	  end if
	  KS.Echo Escape(" <br />来自：" & Node.SelectSingleNode("@province").text & Node.SelectSingleNode("@city").text)
	  KS.Echo Escape(" <br />状态：")
	  If Node.SelectSingleNode("@isonline").text="1" Then KS.Echo escape("<font color='red'>在线</font>") else KS.Echo Escape("离线")
	  KS.Echo Escape(" <br /><img src='" & KS.Setting(3) & "images/user/log/106.gif' border='0' align='absmiddle'> <a href='javascript:void(0)' onclick=""addF(event,'" & username & "')"">加为好友</a> <img src='" & KS.Setting(3) & "images/user/mail.gif' align='absmiddle'> <a href=""javascript:void(0)"" onClick=""sendMsg(event,'" & username & "')"">发送消息</a>")
	  KS.Echo " </td>"
	  KS.Echo "</tr>"
	  KS.Echo "</table>"
	  KS.Echo "</li>"
	 Next
	End If
 End If
 If TotalPut<>0 Then
	 KS.Echo "<div id=""pageNext"" style=""text-align:center;clear:both;"">"
	 KS.Echo "<table align=""center""><tr><td>"
	 If Page>=2 Then
	  KS.Echo Escape("<a class='prev' href='javascript:void(0)' onclick=""query.page(" & Page-1 & ")"">上一页</a>")
	 End If
	 
	 If Page>=10 Then
	  KS.Echo "<a class=""num"" href=""javascript:void(0)"" onclick=""query.page(1)"">1</a> <a class=""num"" href=""javascript:void(0)"" onclick=""query.page(2)"">2</a> <a href='#' class='dh'>...</a>"
	 End If
	 
	 Dim StartPage,EndPage
	 If TotalPage<10 Or Page<10 Then
	  StartPage=1
	  If Page<10 Then EndPage=10 Else  EndPage=TotalPage
	 ElseIf Page>=10 Then
	  StartPage=Page-4
	  EndPage=Page+4
	 ElseIf Page<TotalPage Then
	  StartPage=TotalPage-10
	  EndPage=TotalPage
	 End If
	 If EndPage>TotalPage Then EndPage=TotalPage : StartPage=TotalPage-10
	 If StartPage<0 Then StartPage=1
	 For N=StartPage To EndPage
	  If N=Page Then
	   KS.Echo "<a class=""curr"" href=""#""><span style=""color:red"">" & N & "</a> "
	  Else
	   KS.Echo "<a class=""num"" href=""javascript:void(0)"" onclick=""query.page(" & n &")"">" & N & "</a> "
	  End If
	 Next
	 
	 If TotalPage>10 And Page<TotalPage-4 Then
	  KS.Echo "<a href='#' class='dh'>...</a>"
	  KS.Echo "<a class=""num"" href=""javascript:void(0)"" onclick=""query.page(" & TotalPage-1 & ")"">" & TotalPage-1 & "</a> <a href=""javascript:void(0)"" class=""num"" onclick=""query.page(" & TotalPage & ")"">" & TotalPage & "</a>"
	 End If
	 If Page<>TotalPage Then
	  KS.Echo Escape("<a class='next' href='javascript:void(0)' onclick=""query.page(" & Page+1 & ")"">下一页</a>")
	 End If
	 KS.Echo "</td></tr></table>"
	 
	 KS.Echo "</div>"
	End If
End Sub

'保存空间留言
Sub MessageSave()
		 Dim Content:Content=Request("Content")
		 Dim AnounName:AnounName=KS.S("AnounName")
         Dim HomePage:HomePage=KS.S("HomePage")
         Dim Title:Title=KS.S("Title")
		if AnounName="" Then  KS.Die "请填写你的昵称!"
		if Title="" Then 
		 'Response.Write("请填写留言主题!")
		 'Response.End
		End if
		if Content="" Then KS.Die "请填写留言内容!"
		If Len(KS.LoseHtml(Content))>=500 Then KS.Die "留言内容不能超过500个字!"
		IF lcase(Trim(KS.S("Verifycode")))<>lcase(Trim(Session("Verifycode"))) Then KS.Die "你输入的认证码不正确!"
		
		Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select top 1 * From KS_BlogMessage where 1=0",Conn,1,3
		RS.AddNew
		 RS("AnounName")=AnounName
		 RS("Title")=Title
		 RS("UserName")=KS.S("UserName")
		 RS("HomePage")=HomePage
		 RS("Content")=Content
		 RS("UserIP")=KS.GetIP
		 If KS.SSetting(24)="1" Then
		 RS("Status")=0
		 Else
		 RS("Status")=1
		 End If
		 RS("AddDate")=Now
		RS.UpDate
		 RS.Close:Set RS=Nothing
		 If Not KS.IsNul(KS.C("UserName"))<>"" Then
		  Set KSUser=New UserCls
		  KSUser.UserLoginChecked
		  Set KSUser=Nothing
		 End If
		 KS.Die "<script>alert('恭喜，您的留言已提交!');top.location.reload();</script>"
End Sub 

'用户名
Function GetUserID()
		  If KS.IsNul(KS.C("UserName")) Then
			GetUserID=KS.C("CartID")
		  Else
		    GetUserID=KS.C("UserName")
		  End If
End Function
'加到购物车
Sub addShoppingCart()
   Dim RS,RealPrice,n,arrGroupID,KSUser,LoginTF,str
   Dim Prodid:Prodid=KS.ChkClng(request("id"))
   Dim AttrID:AttrID=KS.ChkClng(request("attrid"))
   Dim KBID:KBID=KS.FilterIds(KS.S("KBID"))
   if Prodid=0 then KS.Die ""
   Dim ProductList:ProductList=Session("ProductList")
   Dim Num:Num=KS.ChkClng(Request("Num"))
   Dim Attr:Attr=KS.DelSQL(UnEscape(Request("AttributeCart")))
   Set RS=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "select top 1 arrGroupID From KS_Product Where id=" & Prodid,conn,1,1
   If RS.Eof And RS.Bof Then
      RS.Close : Set RS=Nothing
	  ks.die "var data={'flag':'error','str':''}"
   End If
   arrGroupID=RS(0)
     Set KSUser=New UserCls
     LoginTF=KSUser.UserLoginChecked
   If Not KS.IsNul(arrGroupID) Then
     If KS.FoundInArr(arrGroupID,KSUser.GetUserInfo("GroupID"),",")=false Then
	  RS.Close:Set RS=Nothing
	  ks.die "var data={'flag':'error1','str':''}"
	 End If
   End If
   RS.Close
   
   
   Dim RSA:Set RSA=Server.CreateObject("ADODB.RECORDSET")
   rsA.open "select top 1 * from KS_ShoppingCart where flag=0 and attrid=" & attrid & " and username='" & GetUserID & "' And proid=" & Prodid,conn,1,3
   if rsa.eof and rsa.bof then
			   rsa.addnew
    end if
	  rsa("flag")=0
	  rsa("proid")=Prodid
	  rsa("username")=GetUserID
	  rsa("attr")=attr
	  rsa("adddate")=now
	  rsa("amount")=Num
	  rsa("attrid")=attrid
	  rsa.update
	rsa.close:set rsa=nothing

	
	if KBID<>"" then  '加捆绑商品
	      Dim K,Price,KBIDArr
		  KBIDArr=Split(KBID,",")
		  For K=0 To Ubound(KBIDArr)
		   If KS.ChkClng(KBIDArr(K))<>0 Then
			  RS.Open "Select top 1 KBPrice From KS_ShopBundleSale Where ProID=" & Prodid & " And KBProID=" &KBIDArr(K),conn,1,1
				 If Not RS.Eof Then
			       Set RSA=Server.CreateObject("ADODB.RECORDSET")
				   RSA.Open "Select top 1 * From KS_ShopBundleSelect where username='" & GetUserID & "' and pid=" & KBIDArr(K) & " and proid=" & Prodid,conn,1,3
				  If RSA.Eof Then
					RSA.AddNew
					RSA("UserName")=GetUserID
					RSA("Pid")=KBIDArr(K)
					RSA("ProID")=Prodid
					RSA("Amount")=1
					RSA("AddDate")=Now
					RSA("Price")=RS(0)
					RSA.Update
				  End If
				  RSA.Close:Set RSA=Nothing
				 End If
				 RS.Close
		 End If
		Next
	end if
	
   str=("<div style=""FONT-SIZE: 10pt;OVERFLOW-y: auto;overflow-x:hidden; WIDTH: 100%; LINE-HEIGHT: 20px; HEIGHT: 150px"">")
   RS.Open "Select i.id,i.title,i.Price_Member,i.Price,i.isdiscount,c.attr,c.amount,c.attrid from KS_Product i Inner Join KS_ShoppingCart c on i.id=c.proid where c.flag=0 and c.username='" & GetUserID & "' and i.verific=1 order by i.id desc",conn,1,1
   if not rs.eof then
      str=str & ("购物车里已有<font style=""color:red"">" & rs.recordcount & "</font>样商品。")
	  str=str & "<table border=""0"" width=""98%"" align=""center"" cellspacing=""0"" cellpadding=""0"" style=""margin-top:10px;"">"
	  n=1
	   Do While Not RS.Eof
	   
      If RS("AttrID")<>0 Then 
	  Dim RSAttr:Set RSAttr=Conn.Execute("Select top 1  * From KS_ShopSpecificationPrice Where ID=" & RS("AttrID"))
	  If Not RSAttr.Eof Then
		RealPrice=RSAttr("Price")
	  Else
		RealPrice=RS("Price_Member")
	  End If
	  RSAttr.CLose:Set RSAttr=Nothing
	 Else
	    RealPrice=RS("Price_Member")
	 End If	  
	 
	IF Cbool(LoginTF)=true Then
	   Dim Discount:Discount=KS.U_S(KSUser.GroupID,17)
	   If Not IsNumeric(Discount) Then Discount=0
	  If KS.ChkClng(RS("isdiscount"))<>0 and Discount<>0 Then
	   RealPrice=FormatNumber(RealPrice*discount/10,2,-1)
	  End If
    End If 
	   
		
	    Num=rs("amount")
		If Num=0 Then Num=1
	    str=str & ("<tr><td style=""line-height:24px; border-bottom:#f1f1f1 1px solid; font-size:12px;""><input type=""hidden"" name=""id"" value=""" & rs(0) & """>" & n & "、<font color=#555>" & ks.gottopic(rs(1),36) & "</font>&nbsp;&nbsp;<span style=""font-size:12px;font-weight:normal;color:#999999"">" & rs("attr") & "</span>")
		set rsa=conn.execute("Select I.ID,I.Title,i.weight,b.Price,b.amount,b.id as selid From KS_Product I inner Join KS_ShopBundleSelect b on i.id=b.pid Where B.ProID=" & RS("ID") & " and b.username='" & GetUserID & "' order by I.id")
		if not rsa.eof then
		    str=str & "<div style=""color:green;font-size:12px"">捆绑购买:</div>"
			do while not rsa.eof
			  str=str & "<div style=""line-height:20px; font-weight:normal;font-size:12px""><span style=""color:#999;float:right"">￥" &rsa("price") & "×1</span>" & rsa("title") & "</div>"
			rsa.movenext
			loop
		end if
		rsa.close

		str=str & ("</td><td width=""80"" style=""line-height:24px;  font-weight:bold; border-bottom:#f1f1f1 1px solid; color:#ff6600;"">￥" & RealPrice & "×" & Num & "</td></tr>")
		n=n+1
	   RS.MoveNext
	   Loop
	  str=str & "</table><br/>"
   end if
   str=str & "</div>"
   RS.Close: Set RS=Nothing
   ks.die "var data={'flag':'ok','str':'" & str & "'}"
End Sub

Sub GetClubBoardOption()
 Call KS.LoadClubBoard()
   Dim node,Xml,n
   Set Xml=Application(KS.SiteSN&"_ClubBoard")
        KS.Echo Escape("<select name=""boardid"">")
   for each node in xml.documentelement.selectnodes("row[@parentid=0]")
		KS.Echo Escape("<optgroup label='" & node.selectsinglenode("@boardname").text &"'>")
		for each n in xml.documentelement.selectnodes("row[@parentid=" & Node.SelectSingleNode("@id").text & "]")
		   KS.Echo Escape("<option value='" & N.SelectSingleNode("@id").text & "'>---" & n.selectsinglenode("@boardname").text &"</option>")
		next
	next
	KS.Echo Escape("</select>")
    Set Xml=Nothing
End Sub

Sub GetPackagePro()
    Dim RS,Key,pricetype,tid,minPrice,maxPrice,param,sqlstr,xml,node
	dim id:id=ks.chkclng(request("id"))
	dim proid:proid=ks.s("proid")
	Key=KS.DelSQL(unescape(Request("Key")))
	pricetype=KS.ChkClng(KS.S("pricetype"))
	tid=KS.S("tid")
	minPrice=KS.S("minPrice"):If Not Isnumeric(minPrice) Then minPrice=0
	maxPrice=KS.S("maxPrice"):If Not Isnumeric(maxPrice) Then maxPrice=0
	param=" where verific=1"
	if tid<>"" and tid<>"0" then param=param & " and tid in(" & KS.GetFolderTid(TID) &")"
	if proid<>"" then param=param & " and proid='"& proid & "'"
    if id<>0 then param=param & " and id<>" & id 

	If PriceType<>0 Then
	  Select Case PriceType
	   case 1 : param=param & " and price>=" & minPrice & " and price<=" & maxPrice
	   case 2 : param=param & " and Price_Original>=" & minPrice & " and Price_Original<=" & maxPrice
	   case 3 : param=param & " and Price_Member>=" & minPrice & " and Price_Member<=" & maxPrice
	  End Select
	End If
	if key<>"" Then
	  Param=Param & " and title like '%" & key & "%'"
	End If
	sqlstr="select top 500 id,title from ks_product" & param & " order by id desc"
	
	set rs=conn.execute(sqlstr)
	if not rs.eof then
	 set xml=KS.RstoXml(rs,"row","")
	end if
	rs.close:set rs=nothing
	if isobject(xml) then
	  for each node in xml.documentelement.selectnodes("row")
       ks.echo "<option value='" & node.selectsinglenode("@id").text & "'>" & node.selectsinglenode("@title").text & "</option>"
	  next
    end if
End Sub

'查看联系信息
Sub GetSupplyContact()
 Dim ID:ID=KS.ChkClng(Request("id"))
 Set RS=Server.CreateObject("Adodb.Recordset")
 RS.Open "Select top 1 b.classpurview,b.defaultarrgroupid,a.* From KS_GQ a inner join KS_Class b on a.Tid=b.ID where a.verific=1 and a.ID=" & ID,Conn,1,1
 if rs.eof and rs.bof then
   rs.close:set rs=nothing
   ks.echo escape("加载出错!")
 else
    if not conn.execute("select top 1 adminid from ks_admin where username='" & rs("inputer") & "'").eof then '判断是管理员发布的信息，则直接显示网站的联系方式
	 rs.close:set rs=nothing
	 KS.Die LFCls.GetConfigFromXML("supply","/labeltemplate/label","noencrypted")
	end if
   Dim KSUser:Set KSUser=New UserCls
   Dim UserLoginTF:UserLoginTF=KSUser.UserLoginChecked
    Dim ClassPurView:ClassPurview=rs("classpurview")
	Dim DefaultArrGroupID:DefaultArrGroupID=rs("defaultarrgroupid")
	' If ClassPurView="2" Then
	     If SupplyPayPoint=0 Then Call ShowSupplyContactInfo(rs):rs.close:set rs=nothing:exit sub
		 IF UserLoginTF=false Then
		        response.write ("<div style='padding:10px;border:1px dashed #cccccc;text-align:center'>对不起,您还没有登录，请<a href='" & KS.Setting(2) & "/user/login/' target='_blank'>登录</a>后再查看联系信息。</div>")
				rs.close:set rs=nothing
				response.end
		 ElseIf KS.FoundInArr(DefaultArrGroupID,KSUser.GroupID,",")=false Then
		        If SupplyPayPoint>0 Then
				   Dim ModelChargeType,ChargeTableName,DateField,ChargeStr,ChargeStrUnit,CurrPoint,IncomeOrPayOut  
				   ModelChargeType=KS.ChkClng(KS.C_S(8,34))
				   Select Case ModelChargeType
					case 1 ChargeStrUnit="元人民币": ChargeTableName="KS_LogMoney" : DateField="PayTime": IncomeOrPayOut="IncomeOrPayOut" : CurrPoint=KSUser.GetUserInfo("Money")
					case 2  ChargeStrUnit="分积分": ChargeTableName="KS_LogScore": DateField="AddDate":IncomeOrPayOut="InOrOutFlag": CurrPoint=KSUser.GetUserInfo("Score")
					case else   '按点券
					  ChargeStrUnit=KS.Setting(46)&KS.Setting(45) : ChargeTableName="KS_LogPoint" : DateField="AddDate" :IncomeOrPayOut="InOrOutFlag": CurrPoint=KSUser.GetUserInfo("Point")
					End Select
				  If Conn.Execute("Select top 1 Times From " & ChargeTableName & " Where ChannelID=8 And InfoID=" & ID & " And " & IncomeOrPayOut & "=2 and UserName='" & KSUser.UserName & "'").eof Then
		          response.write ("<div style='padding:10px;border:1px dashed #cccccc;text-align:center'>需要支付 <span style=""color:red"">" & SupplyPayPoint & " </span>" & ChargeStrUnit  & "才可以查看联系方式,您当前余额 <span style='color:green'>" & CurrPoint & " </span>" & ChargeStrUnit & ",确认支付吗？<br/><input type='button' class='btn' value='确认支付' onclick=""payShow(" & ID & ")""/></div>")
				  else
				   Call ShowSupplyContactInfo(rs)
				  end if
				Else
		          response.write ("<div style='padding:10px;border:1px dashed #cccccc;text-align:center'>对不起,您的级别不够,无法查看联系信息!得到更好服务,请联系本站管理员。</div>")
				End If
				rs.close:set rs=nothing
				response.end
		 End If
	 'End If
     Call ShowSupplyContactInfo(rs)
 end if
 rs.close:set rs=nothing
End Sub
Sub ShowSupplyContactInfo(rs)
   Dim template:template=LFCls.GetConfigFromXML("supply","/labeltemplate/label","contactinfo")
   template=replace(template,"{$GetContactMan}",LFCls.ReplaceDBNull(rs("contactman"),"---"))
   template=replace(template,"{$GetContactTel}",LFCls.ReplaceDBNull(rs("tel"),"---"))
   template=replace(template,"{$GetFax}",LFCls.ReplaceDBNull(rs("fax"),"---"))
   template=replace(template,"{$GetEmail}",LFCls.ReplaceDBNull(rs("email"),"---"))
   template=replace(template,"{$GetHomePage}",LFCls.ReplaceDBNull(rs("homepage"),"---"))
   template=replace(template,"{$GetAddress}",LFCls.ReplaceDBNull(rs("address"),"---"))
   ks.echo (template)   
End Sub
Sub paySupplyShow()
 Dim ID:ID=KS.ChkClng(KS.S("ID"))
 If ID=0 Then KS.Die escape("error:参数出错!")
 Dim KSUser:Set KSUser=New UserCls
 If Cbool(KSUser.UserLoginChecked)=false Then KS.Die escape("error:请先登录!")
 If SupplyPayPoint<=0 Then Exit Sub
 Dim RS:Set RS=Server.CreateObject("Adodb.Recordset")
 RS.Open "Select top 1 * From KS_GQ where verific=1 and ID=" & ID,Conn,1,1
 If RS.Eof And RS.Bof Then 
  RS.Close :Set RS=Nothing
  KS.Die escape("error:找不到记录了!")
 End If
 Descript="查看供求信息[" & RS("Title") & "]的联系方式"
  Select Case KS.ChkClng(KS.C_S(8,34))
		 case 1 
		   If round(KSUser.GetUserInfo("money"),2)<round(SupplyPayPoint,2) Then rs.close:set rs=nothing :KS.Die escape("error:对不起，您的可用余额不足，您当前余额为 " & KSUser.GetUserInfo("money") & " 元!")
		  Call KS.MoneyInOrOut(KSUser.UserName,KSUser.UserName,SupplyPayPoint,4,2,now,0,"系统",Descript,8,ID,1)
		 case 2 
		   If round(KSUser.GetUserInfo("score"),2)<round(SupplyPayPoint,2) Then rs.close:set rs=nothing :KS.Die escape("error:对不起，您的可用余额不足，您当前积分为 " & KSUser.GetUserInfo("score") & " 分!")
		   Session("ScoreHasUse")="+" '设置只累计消费积分
		  Call KS.ScoreInOrOut(KSUser.UserName,2,KS.ChkClng(SupplyPayPoint),"系统",Descript,8,ID)
		 case else
		 	If round(KSUser.GetUserInfo("point"),2)<round(SupplyPayPoint,2) Then rs.close:set rs=nothing :KS.Die escape("error:对不起，您的可用余额不足，您当前" & KS.Setting(45) & "为 " & KSUser.GetUserInfo("point") & " " & KS.Setting(46) & "!")
		   Call KS.PointInOrOut(8,ID,KSUser.UserName,2,SupplyPayPoint,"系统",Descript,0)
  End Select
  ShowSupplyContactInfo(rs)
  RS.Close:Set RS=Nothing
End Sub

'删除图片
Sub DelPhoto()
 Dim UserName,Pass,UserID,i,p,picarr,pic:pic=KS.S("Pic")
 Dim Flag:Flag=KS.ChkClng(Request("flag"))
 Dim PicID:PicID=KS.ChkClng(Request("picid"))
 If Not KS.IsNul(Pic) Then
    PicArr=Split(pic,"|")
	If flag=1 then
	  UserName=KS.C("AdminName")
	  Pass=KS.C("AdminPass")
	  if KS.IsNul(UserName) Or KS.IsNul(Pass) Then
	   ks.die "error"
	  End If
	  If Conn.Execute("Select top 1 * From KS_Admin Where UserName='" & UserName & "' and PassWord='" & Pass & "'").eof Then
	    KS.Die "error"
	  End If
	Else
	  Set KSUser=New UserCls
      LoginTF=KSUser.UserLoginChecked
	  If LoginTF=false Then KS.Die "error!"
	  UserID=KSUser.GetUserInfo("userid")
	end if
	for i=0 to ubound(PicArr)-1
	  p=PicArr(i)
	  If Not KS.IsNul(p) Then 
	     p=replace(p,KS.Setting(2),"")
		 if flag=1 then
		  Call KS.DeleteFile(p)
		 else
		   if instr(lcase(p),lcase("/" & KS.Setting(91) & "user/" & userid & "/"))<>0 then
		    Call KS.DeleteFile(p)
		   end if
		 end if
	  End If
	next
	if picid<>0 then conn.execute("delete from ks_proimages where id=" & picid)
 End If
End Sub


'记录点击广告
Sub HitsGuangGao()
dim Url,getid,getclick,geturl,adssql,RSObj,SqlStr,getip,sitename
		getid=KS.ChkClng(KS.S("id"))
		set RSObj=server.createobject("adodb.recordset")
		adssql="Select top 1 id,url,click,sitename from KS_Advertise where id="&getid
		RSObj.open adssql,Conn,1,3
		if (rsobj.eof and rsobj.bof) then
		 rsobj.close
		 set rsobj=nothing
		 exit sub
		end if
		getclick=RSObj(2)+1
		sitename=RSOBJ(3)
		RSObj(2)=getclick
		RSObj.Update
		Url=RSObj(1)
		RSObj.Close
		'暂且关闭记录IP功能
		SqlStr="select top 1 * from KS_Adiplist where 1=0"
		RSObj.open SqlStr,Conn,1,3
		RSObj.AddNew
		RSObj("adid") = getid
		RSObj("time") = now()
		RSObj("ip") = KS.GetIP
		RSObj("class") = 2
		RSObj.update
		RSObj.close
		set RSObj=nothing 
		
		'========点广告加积分==================
		 if KS.Setting(166)="1" And KS.ChkClng(KS.Setting(167))>0 Then
		   If KS.C("UserName")<>"" Then
		      getid=KS.ChkClng(right(year(now),2)& "" & month(now) & "" & day(now)) & "" & getid  '每天产生不同的ID号，以便第二天增加积分
			  If Conn.Execute("Select top 1 * From KS_LogScore Where UserName='" & KS.C("UserName") & "' and year(adddate)=year(" & SQLNowString  &") and month(adddate)=month(" & SQLNowString &") and day(adddate)=day(" & SQLNowString & ") and channelid=1000 and infoid=" & getid).Eof Then
			  
			   Call  KS.ScoreInOrOut(KS.C("UserName"),1,KS.ChkClng(KS.Setting(167)),"系统","点击广告[" & sitename & "(" & url & ")]所得!!",1000,getid)
			  

			  End If
			  
		   End If
		 End If
		'=====================================
End Sub

Sub GetClubboard()
   Dim Xml,Node,Pid
   Pid=KS.ChkClng(KS.G("pid"))
   KS.LoadClubBoard()
	
%>
<table border="0">
<form name="postform" method="get" action="<%=KS.Setting(3)&KS.Setting(66)%>/post.asp">
 <tr>
 <td><select name="pid" id="pid" size="10" style="width:220px;height:270px" onChange="loadBoard(this.value)">
 <%
 if isobject(Application(KS.SiteSN&"_ClubBoard")) then
	 Set Xml=Application(KS.SiteSN&"_ClubBoard")
	for each node in xml.documentelement.selectnodes("row[@parentid=0]")
	  If trim(Pid)=trim(Node.SelectSingleNode("@id").text) Then
		KS.Echo "<option value='" & Node.SelectSingleNode("@id").text & "' selected>" & node.selectsinglenode("@boardname").text &"</option>"
	  Else
		KS.Echo "<option value='" & Node.SelectSingleNode("@id").text & "'>" & node.selectsinglenode("@boardname").text &"</option>"
	  End If
	next
 end if
 %>
 </select></td>
 <td><select name="bid" id="bid" size="10" style="width:220px;height:270px" onChange=" $('#navlist2').html('->'+$('#bid>option:selected').text());">
 <%
 if isobject(Application(KS.SiteSN&"_ClubBoard")) and pid<>0 then
	for each node in xml.documentelement.selectnodes("row[@parentid=" & pid &"]")
		KS.Echo "<option value='" & Node.SelectSingleNode("@id").text & "'>" & node.selectsinglenode("@boardname").text &"</option>"
	next
 end if
 %>
 </select></td>
 <td id="btns">
 <input type="button" value=" 进 入 " style="margin-bottom:6px" class="btn" onClick="toBoard()"><br/>
 <input type="submit" value=" 发 帖 " style="margin-bottom:6px" class="btn" onClick="return(toPost())"><br/>
 <input type="button" value=" 关 闭 " style="margin-bottom:6px" class="btn" onClick="parent.box.close()">
 </td>
 </tr>
 </form>
</table>
<%		  
 xml=empty
 set node=nothing
End Sub

Sub GetClubPushModel()
  KS.Echo "<select name=""ModelID"" style=""width:130px"" onchange=""getpushclass(this.value)"" Id=""ModelID"" size=""5"">"
  Dim ModelXML,Node
  Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
  For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks6=1]")
	if Node.SelectSingleNode("@ks21").text="1" Then
	  KS.echo "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
	End If
  next
  KS.Echo "<select>"
End Sub

Sub getclubboardcategory()
Dim BoardID:BoardID=KS.ChkClng(Request("boardid"))
If BoardID<>0 Then
     Dim CategoryStr
	 KS.LoadClubBoardCategory
	 For Each CategoryNode In Application(KS.SiteSN&"_ClubBoardCategory").DocumentElement.SelectNodes("row[@boardid=" &BoardID &"]")
     	CategoryStr=CategoryStr & "<option value='" &CategoryNode.SelectSingleNode("@categoryid").text  & "'>" & CategoryNode.SelectSingleNode("@categoryname").text &"</option>"
	Next
	If Not KS.IsNul(CategoryStr) Then
		CategoryStr="<strong>主题分类:</strong><select name=""CategoryId"" id=""CategoryId""><option value='0'>==选择分类==</option>"  & CategoryStr &"</select>"
	End If
	KS.Die Escape(CategoryStr)
End If

End Sub

Sub getonlinelist()
 KS.Echo "<hr size='1' color='#cccccc'/>"
 Dim RS,UserName,Page,PageNum,MaxPerPage,TotalPut,n
 MaxPerPage=24
 Page=KS.ChkClng(Request("page")) : If Page=0 Then Page=1
 Set RS=Conn.Execute("select * from [KS_Online] order by startTime desc")
 If Not RS.Eof Then
            TotalPut=Conn.Execute("Select Count(1) From [KS_Online]")(0)
            If Page < 1 Then Page = 1
			If (totalPut Mod MaxPerPage) = 0 Then
				PageNum = totalPut \ MaxPerPage
			Else
				PageNum = totalPut \ MaxPerPage + 1
			End If

			If (Page - 1) * MaxPerPage < totalPut Then
				RS.Move (Page - 1) * MaxPerPage
			Else
				Page = 1
			End If
	 n=0
	 Do While NOt RS.Eof
	   n=n+1
	   userName=RS("UserName")
	   If UserName="匿名用户" Then
	   KS.Echo "<li><img src='" & KS.Setting(3) & KS.Setting(66) & "/images/guest.png' align='absmiddle'> <a title=""用 户 名:游客&#13;当前位置:" & RS("station") & "&#13;来访时间:" & rs("starttime") & """ href='#'>游客</a></li>"
	   Else
	   KS.Echo "<li><img src='" & KS.Setting(3) & KS.Setting(66) & "/images/" & GetOnlinePic(UserName) & "' align='absmiddle'> <a title=""用 户 名:" & username & "&#13;当前位置:" & RS("station") & "&#13;来访时间:" & rs("starttime") & """ href='" & KS.GetDomain & "space/?" & UserName & "' target='_blank'>" & KS.Gottopic(UserName,15) &"</a></li>"
	   End If
	   If N>=MaxPerPage Then Exit Do
	   RS.MoveNEXT
	 Loop
 End If
 RS.Close
 Set RS=Nothing
 KS.Echo "<div style='clear:both'></div>"
  KS.Echo "<hr size='1' color='#f1f1f1'/>"
 KS.Echo "<div style=""text-align:left"">总在线:<span color='green'>" & TotalPut & "</span> 人 共分为<font color=red>" & PageNum & "</font>页,当前第<font color=red>" & Page & "</font>页"
			  if page>1 then
			  KS.Echo " <a href=""javascript:onlineList(1);"">首页</a>"
			  KS.Echo " <a href=""javascript:onlineList(" & page-1 & ");"">上一页</a>"
			  end if
			  
			  If page<>PageNum Then
			  KS.Echo " <a href=""javascript:onlineList(" & page+1 & ");"">下一页</a>"
			  KS.Echo " <a href=""javascript:onlineList(" & pagenum & ");"">末页</a>"
			  End If
			  KS.Echo "</div>"
 
End Sub
Function GetOnlinePic(username)
 if not conn.execute("select top 1 username from ks_admin where username='" & username & "'").eof then
   GetOnlinePic="admin.png" 
 Elseif not conn.execute("select top 1 master from ks_guestboard where master+',' like'%" & username & "%,'").eof then
   GetOnlinePic="mod.png"
 Else
   GetOnlinePic="member.png"
 end if
End Function
%>