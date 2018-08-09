<%
Class ManageCls
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
	
       '从xml中加载模型字段
	   Sub LoadModelField(ChannelID,ByRef FieldXML,ByRef FieldNode)
	        set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			FieldXML.async = false
			FieldXML.setProperty "ServerHTTPRequest", true 
			FieldXML.load(Server.MapPath(KS.Setting(3)&"Config/fielditem/field_" & ChannelID&".xml"))
			if FieldXML.parseError.errorCode<>0 Then
				 Call CreateModelField(ChannelID)
				 FieldXML.load(Server.MapPath(KS.Setting(3)&"Config/fielditem/field_" & ChannelID&".xml"))
			End If
			if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set FieldNode=FieldXML.DocumentElement.SelectNodes("fielditem[showonform=1&&fieldtype!=13]")
			end if
	   End Sub
	   
	   '判断并生成系统默认字段	   
	   Sub CreateModelField(ChannelID)
	          Dim DefaultField,FieldArr,FieldItem,N,RS,itemname
			  Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "select top 1 * From KS_Field Where ChannelID=" &ChannelID &" And FieldType=0",conn,1,1
			  IF rs.EOF And RS.Bof Then
			    itemname=KS.C_S(ChannelID,3)
			    Select Case KS.ChkClng(KS.C_S(ChannelID,6))
			       Case 1
				    DefaultField="简短标题@title|标题属性@titleattribute|完整标题@fulltitle|归属栏目@tid|"&itemname&"属性@attribute|推送到论坛@pushtobbs|转向链接@turnto|关 键 字@keywords|"&ItemName&"作者@author|"&itemname&"来源@origin|地图标注@map|选择地区@area|"&itemname&"导读@intro|"&itemname&"内容@content|附件上传@attachment|分页标题@pagetitle|图片地址@photourl|上传图片@uploadphoto|添加日期@adddate|"&itemname&"等级@rank|点 击 数@hits|模板选择@template|自定义文件名@fname|归属专题@special|属性选项@attroption|SEO优化选项@seooption|收费选项@chargeoption|立即发布@pub|签收选项@signoption|相关选项@relativeoption"
				   Case 2
				    DefaultField=ItemName & "名称@title|归属栏目@tid|地图标注@map|" & itemname & "属性@attribute|关 键 字@keywords|" & itemname & "作者@author|" & itemname & "来源@origin|缩 略 图@photourl|" &itemname & "内容@content|" & itemname & "介绍@picturecontent|添加日期@adddate|" & itemname & "等级@rank|浏 览 数@hits|模板选择@template|自定义文件名@fname|归属专题@special|属性选项@attroption|SEO优化选项@seooption|收费选项@chargeoption|立即发布@pub|相关选项@relativeoption"
                  Case 3
				    DefaultField=ItemName & "名称@title|归属栏目@tid|下载地址@address|版本号@version|" &itemname & "属性@attribute|" &itemname & "性质@nature|系统平台@platform|" & itemname & "图片@photourl|上传图片@uploadphoto|关 键 字@keywords|作者开发商@author|" & itemname & "来源@origin|上传" & itemname &"@uploadsoft|" & itemname &"介绍@content|演示地址@ysdz|注册地址@zcdz|解压密码@jymm|添加日期@adddate|" & itemname & "等级@rank|浏 览 数@hits|模板选择@template|文 件 名@fname|所属专题@special|属性选项@attroption|SEO优化选项@seooption|收费选项@chargeoption|立即发布@pub|相关选项@relativeoption"
				  Case 4
				    DefaultField=ItemName & "名称@title|归属栏目@tid|"&itemname&"属性@attribute|" & itemname & "图片@photourl|上传图片@uploadphoto|关 键 字@keywords|" & itemname & "作者@author|" & itemname & "来源@origin|" & itemname &"地址@uploadflash|" & itemname &"介绍@content|添加日期@adddate|" & itemname & "等级@rank|浏 览 数@hits|模板选择@template|文 件 名@fname|所属专题@special|属性选项@attroption|SEO优化选项@seooption|收费选项@chargeoption|立即发布@pub|相关选项@relativeoption"
				 Case 5
				    DefaultField=ItemName & "名称@title|归属栏目@tid|归属品牌@brandid|"&itemname&"属性@attribute|列表图片@photourl|上传图片@uploadphoto|商品单位@unit|商品价格@price|允许折扣@isdiscount|库存设置@totalnum|购物车属性@cartpropert|下载地址@downurl|" & ItemName & "介绍@prointro|组图上传@uploadphotos|属性选项@attroption|SEO优化选项@seooption|捆绑销售@kboption|立即发布@pub|相关选项@relativeoption"
				  Case 7
				    DefaultField=ItemName & "名称@title|归属栏目@tid|"&itemname&"属性@attribute|"&itemname&"参数@parameter|上映时间@screentime|主要演员@movieact|" & itemname & "导演@moviedy|" &itemname & "图片@photourl|上传图片@uploadphoto|关 键 字@keywords|" & itemname &"地址@uploadmovie|" & itemname &"介绍@content|添加日期@adddate|" & itemname & "等级@rank|观看次数@hits|模板选择@template|文 件 名@fname|所属专题@special|属性选项@attroption|SEO优化选项@seooption|收费选项@chargeoption|立即发布@pub|相关选项@relativeoption"
				End Select
				 
				 FieldArr=Split(DefaultField,"|")
				 For N=0 To Ubound(FieldArr)
				    FieldItem=split(FieldArr(N),"@")
				    Conn.Execute("INSERT INTO KS_Field(ChannelID,FieldName,Title,FieldType,ShowOnForm,ShowOnUserForm,OrderID) VALUES(" & ChannelID &",'" & FieldItem(1) & "','" & FieldItem(0) &"',0,1,1," & N+1 &")")
				 Next
			End If
			Call CreateFieldXML(ChannelID,"") '生成xml缓存
			RS.Close:Set RS=Nothing
	   End Sub
	   
	   '生成模型字段xml
	   Sub CreateFieldXML(ChannelID,Param)
	      Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "select * from KS_Field Where ChannelID=" & ChannelID & " " & Param & " Order By OrderID,FieldID",Conn,1,1
	      Dim XMLStr:XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
		  XMLStr=XMLStr&" <field>" &vbcrlf
		  If Not RS.Eof Then
					Do While Not RS.Eof
					    XMLStr=XMLStr & "  <fielditem id=""" & RS("FieldID") &""" fieldname=""" & replace(rs("fieldname"),"&","") & """>"&vbcrlf
						XMLStr=XMLStr & "    <title>" & rs("title") & "</title>" &vbcrlf
						XMLStr=XMLStr & "    <tips><![CDATA[" & rs("tips") &"]]></tips>" &vbcrlf
						XMLStr=XMLStr & "    <fieldtype>" & rs("fieldtype") & "</fieldtype>" &vbcrlf
						XMLStr=XMLStr & "    <defaultvalue><![CDATA[" & rs("defaultvalue") &"]]></defaultvalue>" &vbcrlf
						If Not KS.IsNul(rs("options")) Then
						XMLStr=XMLStr & "    <options><![CDATA[" & replace(rs("options"),vbcrlf,"\n") &"]]></options>" &vbcrlf
						Else
						XMLStr=XMLStr & "    <options><![CDATA[" & rs("options") &"]]></options>" &vbcrlf
						End If
						XMLStr=XMLStr & "    <mustfilltf>" & rs("mustfilltf") & "</mustfilltf>" &vbcrlf
						XMLStr=XMLStr & "    <showonform>" & rs("showonform") & "</showonform>" &vbcrlf
						XMLStr=XMLStr & "    <showonuserform>" & rs("showonuserform") & "</showonuserform>" &vbcrlf
						XMLStr=XMLStr & "    <showonclubform>" & ks.chkclng(rs("showonclubform")) & "</showonclubform>" &vbcrlf
						XMLStr=XMLStr & "    <allowfileext>" & rs("AllowFileExt") & "</allowfileext>" &vbcrlf
						XMLStr=XMLStr & "    <width>" & rs("width") & "</width>" &vbcrlf
						XMLStr=XMLStr & "    <height>" & rs("height") & "</height>" &vbcrlf
						XMLStr=XMLStr & "    <maxfilesize>" & rs("maxfilesize") & "</maxfilesize>" &vbcrlf
						XMLStr=XMLStr & "    <editortype>" & rs("editortype") & "</editortype>" &vbcrlf
						XMLStr=XMLStr & "    <showunit>" & rs("showunit") & "</showunit>" &vbcrlf
						if not KS.IsNul(rs("unitoptions")) Then
						XMLStr=XMLStr & "    <unitoptions><![CDATA[" & replace(rs("unitoptions"),vbcrlf,"\n") &"]]></unitoptions>" &vbcrlf
						Else
						XMLStr=XMLStr & "    <unitoptions><![CDATA[" & rs("unitoptions") &"]]></unitoptions>" &vbcrlf
						End If
						XMLStr=XMLStr & "    <parentfieldname>" & rs("ParentFieldName") & "</parentfieldname>" &vbcrlf
						XMLStr=XMLStr & "    <maxlength>" & rs("maxlength") & "</maxlength>" &vbcrlf
					    XMLStr=XMLStr & "  </fielditem>"&vbcrlf
					 RS.MoveNext
					Loop
		  End If
		   XMLStr=XMLStr &" </field>" &vbcrlf
		   Call KS.WriteTOFile(KS.Setting(3) & "config/fielditem/field_" & ChannelID & ".xml",xmlstr)
		  RS.Close :Set RS=Nothing
	   End Sub
	   
	   '检查录入的自定义字段
		Sub CheckDiyField(FieldXML,byref ErrMsg)
		      Dim Node,FieldName,XTitle,FieldType
			  if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				  Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype!=0&&fieldtype!=13]")
				  If DiyNode.Length>0 Then
					For Each Node In DiyNode
					 FieldName = Node.SelectSingleNode("@fieldname").text
					 FieldType = Node.SelectSingleNode("fieldtype").text
					 XTitle    = Node.SelectSingleNode("title").text
					 If Node.SelectSingleNode("mustfilltf").text="1" And KS.IsNul(KS.G(FieldName)) Then ErrMsg = ErrMsg & XTitle & "必须填写!\n"
					 If (FieldType="4" or FieldType="12") And Not KS.IsNul(KS.G(FieldName)) And Not Isnumeric(KS.G(FieldName)) Then ErrMsg = ErrMsg& XTitle & "必须填写数字!\n"
					 If FieldType="5" And Not KS.IsNul(KS.G(FieldName)) And Not IsDate(KS.G(FieldName)) Then ErrMsg = ErrMsg& XTitle & "必须填写正确的日期!\n" 
					 If FieldType="8" And Not KS.IsValidEmail(KS.G(FieldName)) and Node.SelectSingleNode("mustfilltf").text="1" Then ErrMsg = ErrMsg& XTitle & "必须填写正确的Email格式!\n" 
					Next
				End If
			End If
		   End Sub
		   '更新自定义字段的值
		   Sub AddDiyFieldValue(ByRef RS,FieldXML)
		      Dim Node,FieldName,FieldType
			  if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				  Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype!=0 && showonform=1]")
				  If DiyNode.Length>0 Then
						For Each Node In DiyNode
							 FieldName = Node.SelectSingleNode("@fieldname").text
							 FieldType = Node.SelectSingleNode("fieldtype").text
							  If (Not KS.IsNul(KS.G(FieldName)) And (FieldType="4" Or FieldType="12")) or  (FieldType<>"4" and FieldType<>"12") Then
								If FieldType="10"  Then   '支持HTML时
								 RS("" & FieldName & "")=Request.Form(FieldName)
								elseIf FieldType="5" and not isdate(KS.G(FieldName)) Then
								ElseIf FieldType="13" Then
								 RS("" & FieldName & "")=KS.ChkClng(KS.G(FieldName))
								Else
								 RS("" & FieldName & "")=KS.G(FieldName)
								end if
								If Node.SelectSingleNode("showunit").text="1"  Then
								RS("" & FieldName & "_Unit")=KS.G(FieldName&"_Unit")
								End If
							 End If
						Next
				 End If
			 End If
		   End Sub
		   '更新相关信息
		   Sub UpdateRelative(ChannelID,InfoID,InfoList,Deltf)
		      If DelTF=1 Then Conn.Execute("Delete From KS_ItemInfoR Where InfoID=" & InfoID & " and channelid=" & ChannelID)
		      If InfoList<>"" Then
				Dim SelectInfoList,RelativeArr,I,HasInRelativeID
				SelectInfoList=Split(InfoList,",")
				For I=0 To Ubound(SelectInfoList)
					If Instr(SelectInfoList(i),"↓")=0 Then SelectInfoList(i)=SelectInfoList(i) &"↓"
					RelativeArr=split(SelectInfoList(i),"↓")
					If KS.FoundInArr(HasInRelativeID,SelectInfoList(i),",")=false Then
						   Conn.Execute("Insert Into KS_ItemInfoR(ChannelID,InfoID,RelativeChannelID,RelativeID,relativeText) values(" & ChannelID &"," & InfoID & "," & Split(RelativeArr(0),"|")(0) & "," & Split(RelativeArr(0),"|")(1) & ",'" & RelativeArr(1) &"')")
					   HasInRelativeID=HasInRelativeID & SelectInfoList(i) & ","
					 End If
				Next
			  End If
		   End Sub
		   
			'返回系统支持的生成类型(.htm,.html,.shtml.shtm等)参  数：ExtType 预定选中的类型
			Public Function GetFsoTypeStr(ExtType)
			  GetFsoTypeStr = "<select name='fnametype' id='fnametype'>"
			If ExtType = ".html" Then
			  GetFsoTypeStr = GetFsoTypeStr & "<option value='.html' selected>.html</option>"
			Else
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.html'>.html</option>"
			End If
			If ExtType = ".htm" Then
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.htm' selected>.htm</option>"
			Else
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.htm'>.htm</option>"
			End If
			If ExtType = ".shtm" Then
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.shtm' selected>.shtm</option>"
			Else
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.shtm'>.shtm</option>"
			End If
			If ExtType = ".shtml" Then
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.shtml' selected>.shtml</option>"
			Else
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.shtml'>.shtml</option>"
			End If
			If ExtType = ".asp" Then
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.asp' selected>.asp</option>"
			Else
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.asp'>.asp</option>"
			End If
			 GetFsoTypeStr = GetFsoTypeStr & "</select>"
			End Function
       '取得专题
		Sub Get_KS_Admin_Special(ChannelID,InfoID)
		   With KS
		     .echo "<script language='javascript'>" & vbcrlf
			 .echo "  SelectSpecial=function(){" &vbcrlf
			 .echo "		new KesionPopup().PopupCenterIframe('选择专题','KS.Special.asp?action=Select',350,400,'auto')" & vbcrlf
			 .echo "	}" &vbcrlf
			 .echo "  SelectSpecial1=function(){" &vbcrlf
			 .echo "		var strUrl = 'KS.SpecialSelect.asp'; "& vbcrlf
			 .echo "		var isMSIE= (navigator.appName == 'Microsoft Internet Explorer');" & vbcrlf
			 .echo "		var ReturnStr = null;" &vbcrlf
			 .echo "		if (isMSIE){ReturnStr= window.showModalDialog(strUrl,self,'width=250,height=400,resizable=yes,scrollbars=yes');}" &vbcrlf
			 .echo "		else{ var win=window.open(strUrl,'newWin','left=150,width=350,height=400,resizable=yes,scrollbars=yes'); }"&vbcrlf
			 .echo "		if (ReturnStr != null){" & vbcrlf
			 .echo "			UpdateSpecial(ReturnStr);}" & vbcrlf
			 .echo "	}" &vbcrlf
			 .echo "    function UpdateSpecial(arrstr){" &vbcrlf
			 .echo "	  if (arrstr!=''){" &vbcrlf
			 .echo "	  $('#SpecialList').show();" & vbcrlf
			 
			 .echo "     var finder=false;" & vbcrlf
			 .echo "	  var arr=arrstr.split('@@@');" & vbcrlf
			 .echo "     $('#SpecialID>option').each(function(){" & vbcrlf
			 .echo "     if (arr[0]==this.value){" & vbcrlf
			 .echo "       $('#SpecialID>option[value='+arr[0]+']').attr('selected',true);finder=true;return false;}" &vbcrlf
			 .echo "  });" & vbcrlf
			 .echo "  if (finder==false){" & vbcrlf
			 .echo "	$('#SpecialID').append(""<option value=""+arr[0]+"">""+arr[1]+""</option>"");" & vbcrlf
			 .echo "	$('#SpecialID >option[value='+arr[0]+']').attr('selected',true);" & vbcrlf
			 .echo " }" & vbcrlf
			 .echo "	 }" & vbcrlf
			 .echo "	}" & vbcrlf
			 .echo " </script>" & vbcrlf
			.echo "<table border=0 width='100%'><tr>"
			Dim ShowSpecialStr:ShowSpecialStr=" style='display:none'"
			If InfoID<>0 Then
			   Dim OptionStr,RSB
			   Set RSB=Conn.Execute("Select a.SpecialID,SpecialName From KS_Special A inner join KS_SpecialR b on a.specialid=b.specialid Where ChannelID=" & ChannelID & " and InfoID=" & InfoID)
				If Not RSB.Eof Then
				  ShowSpecialStr=""
				  Do While Not RSB.Eof
				   OptionStr=OptionStr & "<option value='" & RSB(0) & "' selected>" &RSb(1) & "</option>"
				  RSB.MoveNext
				  Loop
				End If
				RSB.Close:Set RSB=Nothing
			End If
			.echo "<td width='200' id='SpecialList'" & ShowSpecialStr &">"
			.echo "<select name='SpecialID' id='SpecialID' multiple style='height:100px;width:200px;'>" & OptionStr & "</select><div style='text-align:center'><font color=red>X</font> <a href='javascript:UnSelectAll()'><font color='#999999'>取消选定的专题</font></a></div></td>"
			.echo "              <td><input class='button'  type='button' name='Submit' value='选择专题...' onClick='SelectSpecial();'></td>"
			.echo "</table>"
		  End With
		End Sub
	  '从数据表添加数据到option选项 参数:表名,字段,查询条件
	  Function Get_O_F_D(Table,FieldStr,Param)
	       Dim KS_RS_Obj,Arr,I
		      If Instr(lcase(FieldStr),"distinct")<=0 and Instr(lcase(FieldStr),"top")<=0 Then FieldStr=" top 50 " &FieldStr
			  Set KS_RS_Obj = conn.Execute("Select " & FieldStr & " FROM "  & Table & " Where " & Param)
			  If Not KS_RS_Obj.Eof Then
			    Arr=KS_RS_Obj.GetRows(-1)
				KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
				For I=0 To Ubound(Arr,2)
					Get_O_F_D = Get_O_F_D & "<option value=""" & Arr(0,i) & """>" & Arr(0,i) & "</option>"
				Next
			   End If
	  End Function
	  '取得相应的模板  参数 obj对象
	  Function Get_KS_T_C(obj)
	    Dim CurrPath:CurrPath=KS.Setting(3)&KS.Setting(90)
		If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath,Len(CurrPath)-1)
        Get_KS_T_C= "<input type='button' name=""Submit"" class=""button"" value=""选择模板..."" onClick=""OpenThenSetValue('KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle="& server.URLEncode("导入模板")&"&CurrPath=" &server.urlencode(CurrPath) & "',450,350,window," & obj & ");"">"	 
	   End Function
       
	   '取得自定义属性录入
	   Sub GetDiyAttribute(FieldXML,FieldDictionary)
	        Dim ANode,AttrNode:Set AttrNode=FieldXML.DocumentElement.SelectNodes("fielditem[fieldtype=13]")
			If IsObject(AttrNode) Then
				 For Each ANode In AttrNode
					KS.Echo "<label><input name='" & ANode.SelectSingleNode("@fieldname").text & "' type='checkbox' value='1'"
				   If Isobject(FieldDictionary) Then
						 if FieldDictionary.item(lcase(ANode.SelectSingleNode("@fieldname").text))="1" then  KS.Echo " checked"
				   ElseIf ANode.SelectSingleNode("defaultvalue").text="1" then 
				          KS.Echo " checked"
				   End If
					KS.Echo ">" & ANode.SelectSingleNode("title").text & "</label>"
				 Next
			End If
	   End Sub
	   
	   '取得后台信息添加时的自定义字段表单
	    Function GetDiyField(ChannelID,FieldXML,Node,FieldDictionary,V_Tag)
		      Dim I,K,O_Arr,F_Value,fieldname,fieldtype,XTitle,XWidth,XHeight,XMaxlength
			  Dim O_Text,O_Value,BRStr,O_Len,F_V,UnitValue,V_Arr
			  		
				  If Node.SelectSingleNode("parentfieldname").text="0" Or KS.IsNul(Node.SelectSingleNode("parentfieldname").text) Then
				    fieldname = Node.SelectSingleNode("@fieldname").text
					fieldtype = Node.SelectSingleNode("fieldtype").text
				    XTitle    = Node.SelectSingleNode("title").text
					XWidth    = Node.SelectSingleNode("width").text
					XHeight   = Node.SelectSingleNode("height").text
					XMaxlength= Node.SelectSingleNode("maxlength").text
				    If (ChannelID=101 and Node.SelectSingleNode("showonuserform").text="0") Or (ChannelID=101 And Lcase(Node.SelectSingleNode("showonuserform").text)="mobile") Then
				    GetDiyField=GetDiyField & "<tr class='tdbg'{@NoDisplay(" & fieldname & ")}>" & vbcrlf 
					Else
				    GetDiyField=GetDiyField & "<tr class='tdbg'>" & vbcrlf 
					End If
					GetDiyField=GetDiyField & " <td width=""85"" nowrap align=""right"" class='clefttitle'><strong>" & XTitle & "：</strong></td>" & vbcrlf
					GetDiyField=GetDiyField & " <td style=""word-break:break-all"">"
					 If Isobject(FieldDictionary) Then
					    F_Value=FieldDictionary.item(lcase(fieldname))
					    If Node.SelectSingleNode("showunit").text="1" Then
					    UnitValue=FieldDictionary.item(lcase(fieldname) &"_unit")
						End If
					 Else
					   if lcase(Node.SelectSingleNode("defaultvalue").text)="now" then
					   F_Value=now
					   elseif lcase(Node.SelectSingleNode("defaultvalue").text)="date" then
					   F_Value=date
					   else
					   F_Value=Node.SelectSingleNode("defaultvalue").text
					   end if
					   If Instr(F_Value,"|")<>0 Then 
					   	F_Value=LFCls.GetSingleFieldValue("select top 1 " & Split(F_Value,"|")(1) & " from " & Split(F_Value,"|")(0) & " where username='" & KS.C("UserName") & "'") 
					   End If
					 End If
					 
				   If V_Tag=1 Then	 
				    GetDiyField=GetDiyField & "[@" & fieldname &"]"
                   ElseIf lcase(fieldname)="province&city" Then
				   	GetDiyField=GetDiyField & "<script language=""javascript"" src=""" & KS.Setting(2) & "/Plus/Area.asp""></script>"
				   Else
					   Select Case fieldtype
						 Case 2
						   GetDiyField=GetDiyField & "<textarea style=""width:" & XWidth & "px;height:" & XHeight & "px"" rows=""5"" class=""upfile"" name=""" & fieldname & """>" & F_Value & "</textarea>"
						 Case 3,11
							   If fieldtype=11 Then
								 GetDiyField=GetDiyField & "<select class=""upfile"" style=""width:" & XWidth & "px"" name=""" & fieldname & """ onchange=""fill" & fieldname &"(this.value)""><option value=''>---请选择---</option>"
	
							   Else
							  GetDiyField=GetDiyField & "<select class=""upfile"" style=""width:" & XWidth & "px"" name=""" & fieldname & """>"
							   End If
								   O_Arr=Split(Node.SelectSingleNode("options").text,"\n"): O_Len=Ubound(O_Arr)
								   For K=0 To O_Len
									If O_Arr(K)<>"" Then
									   F_V=Split(O_Arr(K),"|")
									   If Ubound(F_V)=1 Then
										O_Value=F_V(0):O_Text=F_V(1)
									   Else
										O_Value=F_V(0):O_Text=F_V(0)
									   End If						   
									 If F_Value=O_Value Then
									  GetDiyField=GetDiyField & "<option value=""" & O_Value& """ selected>" & O_Text & "</option>"
									 Else
									  GetDiyField=GetDiyField & "<option value=""" & O_Value& """>" & O_Text & "</option>"
									 End If
									End If
								   Next
							  GetDiyField=GetDiyField & "</select>"
							  '联动菜单
							  If fieldtype=11  Then
								Dim JSStr
								GetDiyField=GetDiyField &  GetLinkAgeMenuStr(ChannelID,FieldXML,FieldDictionary,fieldname,JSStr) & "<script type=""text/javascript"">" &vbcrlf & JSStr& vbcrlf &"</script>"
							  End If
						 Case 6
						   O_Arr=Split(Node.SelectSingleNode("options").text,"\n"): O_Len=Ubound(O_Arr)
						   If O_Len>1 And Len(Node.SelectSingleNode("options").text)>50 Then BrStr="<br>" Else BrStr=""
						   For K=0 To O_Len
							   F_V=Split(O_Arr(K),"|")
							   If O_Arr(K)<>"" Then
							   If Ubound(F_V)=1 Then
								O_Value=F_V(0):O_Text=F_V(1)
							   Else
								O_Value=F_V(0):O_Text=F_V(0)
							   End If						   
							 If F_Value=O_Value Then
							  GetDiyField=GetDiyField & "<input type=""radio"" name=""" & fieldname & """ value=""" & O_Value& """ checked>" & O_Text & BRStr
							 Else
							  GetDiyField=GetDiyField & "<input type=""radio"" name=""" & fieldname & """ value=""" & O_Value& """>" & O_Text & BRStr
							 End If
							End If
						   Next
						 Case 7
						   O_Arr=Split(Node.SelectSingleNode("options").text,"\n"): O_Len=Ubound(O_Arr)
						   For K=0 To O_Len
						     If O_Arr(K)<>"" Then
							   F_V=Split(O_Arr(K),"|")
							   If Ubound(F_V)=1 Then
								O_Value=F_V(0):O_Text=F_V(1)
							   Else
								O_Value=F_V(0):O_Text=F_V(0)
							   End If						   
							 If KS.FoundInArr(F_Value,O_Value,",")=true Then
							  GetDiyField=GetDiyField & "<input type=""checkbox"" name=""" & fieldname & """ value=""" & O_Value& """ checked>" & O_Text
							 Else
							  GetDiyField=GetDiyField & "<input type=""checkbox"" name=""" & fieldname & """ value=""" & O_Value& """>" & O_Text
							 End If
							End If
						   Next
						 case 9
						 Case 10
						    if KS.IsNUL(F_Value) Then F_Value=""
							GetDiyField=GetDiyField & "<textarea id=""" & fieldname &""" name=""" & fieldname &""">"& Server.HTMLEncode(F_Value) &"</textarea><script type=""text/javascript"">CKEDITOR.replace('" & fieldname &"', {width:""" & XWidth &""",height:""" & Xheight & """,toolbar:""" & Node.SelectSingleNode("editortype").text & """,filebrowserBrowseUrl :""Include/SelectPic.asp?from=ckeditor&Currpath="& KS.GetUpFilesDir() &""",filebrowserWindowWidth:650,filebrowserWindowHeight:290});</script>"
						 Case Else
						   Dim MaxLength:MaxLength=XMaxLength
						   If Not IsNumerIc(MaxLength) Or MaxLength="0" Then MaxLength=255
						   GetDiyField=GetDiyField & "<input maxlength=""" &MaxLength &""" type=""text"" class=""textbox"" style=""width:" & XWidth & "px"" name=""" & fieldname & """ id=""" & fieldname & """ value=""" & F_Value & """>"
					   End Select
				   End If
				   
				   If Node.SelectSingleNode("showunit").text="1" and channelid<>101 Then 
					  GetDiyField=GetDiyField & " <select name=""" & fieldname & "_Unit"" id=""" & fieldname & "_Unit"">"
					  If Not KS.IsNul(Node.SelectSingleNode("unitoptions").text) Then
				       Dim UnitOptionsArr:UnitOptionsArr=Split(Node.SelectSingleNode("unitoptions").text,"\n")
					   For K=0 To Ubound(UnitOptionsArr)
					       if trim(UnitValue)=trim(UnitOptionsArr(k)) then
					       GetDiyField=GetDiyField & "<option value='" & UnitOptionsArr(k) & "' selected>" & UnitOptionsArr(k) & "</option>"                 
						   else
					       GetDiyField=GetDiyField & "<option value='" & UnitOptionsArr(k) & "'>" & UnitOptionsArr(k) & "</option>"                 
						   end if
					   Next
					  End If
					  GetDiyField=GetDiyField & "</select>"
				   End If
				   
				   if fieldtype=9 and V_Tag<>1 Then GetDiyField=GetDiyField & "<table border=0 cellspaceing='0' cellpadding='0'><tr><td><input maxlength=""" &MaxLength &""" type=""text"" class=""textbox"" style=""width:" & XWidth & "px"" name=""" & fieldname & """ id=""" & fieldname & """ value=""" & F_Value & """></td><td align='left'><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='KS.UpFileForm.asp?UPType=Field&AllowFileExt=" & Node.SelectSingleNode("allowfileext").text & "&MaxFileSize=" & Node.SelectSingleNode("maxfilesize").text & "&FieldName=" & fieldname & "&FieldID=" & Node.SelectSingleNode("@id").text & "&ChannelID=" & ChannelID &"' frameborder=0 scrolling=no width='165' height='28'></iframe></td><td><button class=""button""  type='button' name='Submit' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID &"&CurrPath=" & KS.GetUpFilesDir() & "',550,290,window,$('#" & fieldname & "')[0]);"">文件库...</button></td></tr></table>"
				   If Node.SelectSingleNode("mustfilltf").text="1" Then GetDiyField=GetDiyField & "<font color=red> * </font>"
				   If  Node.SelectSingleNode("tips").text<>"" Then GetDiyField=GetDiyField & " <span style=""margin-top:5px"">" &  Node.SelectSingleNode("tips").text & "</span>"
				   GetDiyField=GetDiyField &" </td>" &vbcrlf
				   GetDiyField=GetDiyField & "</tr>" &vbcrlf
				 End If
		   End Function
		   '取得联动菜单
		   Function GetLinkAgeMenuStr(ChannelID,FieldXML,FieldDictionary,byVal ParentFieldName,JSStr)
		     Dim OptionS,OArr,I,VArr,V,F,Str,Node,FieldName
			 If ParentFieldName="0" Or ParentFieldName="" Then Exit Function
			 Dim PNode:Set PNode=FieldXML.DocumentElement.selectsinglenode("fielditem[parentfieldname='" & ParentFieldName &"']")
			 If not pnode is nothing Then 
			     FieldName=pnode.selectsinglenode("@fieldname").text
			     Str=Str & " <select name='" & FieldName & "' id='" & FieldName & "' onchange='fill" & FieldName & "(this.value)' style='width:" & pnode.selectsinglenode("width").text & "px'><option value=''>--请选择--</option>"
				 JSStr=JSStr & "var sub" &ParentFieldName & " = new Array();" &vbcrlf
				  Options=pnode.selectsinglenode("options").text
				  OArr=Split(Options,"\n")
				  For I=0 To Ubound(OArr)
				    Varr=Split(OArr(i),"|")
					If Ubound(Varr)=1 Then 
					 V=Varr(0):F=Varr(1)
					Else
					 V=Varr(0):F=Varr(0)
					End If
				    JSStr=JSStr & "sub" & ParentFieldName&"[" & I & "]=new Array('" & V & "','" & F & "')" &vbcrlf
				  Next
				 Str=Str & "</select>" &vbcrlf
				 JSStr=JSStr & "function fill"& ParentFieldName&"(v){" &vbcrlf &_
							   "$('#"& FieldName&"').empty();" &vbcrlf &_
							   "$('#"& FieldName&"').append('<option value="""">--请选择--</option>');" &vbcrlf &_
							   "for (i=0; i<sub" &ParentFieldName&".length; i++){" & vbcrlf &_
							   " if (v==sub" &ParentFieldName&"[i][0]){document.getElementById('" & FieldName & "').options[document.getElementById('" & FieldName & "').length] = new Option(sub" &ParentFieldName&"[i][1], sub" &ParentFieldName&"[i][1]);}}" & vbcrlf &_
							   "}" &vbcrlf
				 Dim DefaultVAL
				 If IsObject(FieldDictionary) Then DefaultVAL=FieldDictionary.item(lcase(fieldName))
				 If Not KS.IsNul(DefaultVAL) Then
				  str=str & "<script>$(document).ready(function(){fill"&ParentFieldName&"($('select[name=" &ParentFieldName&"] option:selected').val()); $('#"& FieldName&"').val('" & DefaultVAL & "');})</script>" &vbcrlf
				 End If
				 GetLinkAgeMenuStr=str & GetLinkAgeMenuStr(ChannelID,FieldXML,FieldDictionary,FieldName,JSStr)
			 Else
			     JSStr=JSStr & "function fill" & ParentFieldName &"(v){}"				 
			 End If
		   End Function

	   
	   
	   
	   
	   
	   '====================================================复制操作开始=================================
	    '粘贴
		Sub Paste(ChanneLID)
		 Dim DestFolderID, ContentID,Url
		  DestFolderID = KS.G("DestFolderID")
		  ContentID = KS.G("ContentID")
		  If DestFolderID = ""  Then Call KS.AlertHistory("参数传递出错!", 1):Exit Sub
		  Call PasteByCopy(ChannelID,DestFolderID, ContentID)
		  KS.Echo "<script>location.href='?ChannelID=" & ChannelID &"&ID=" & DestFolderID & "&Page=" & KS.S("Page") & "';</script>"
		End Sub
	   
	    '过程:PasteByCopy复制粘贴
		'参数:ChannelID--模型ID,NewClassID--目标目录,ContentID---被复制的文件
		Sub PasteByCopy(ChannelID,NewClassID, ContentID)
		 If ContentID <> "0" Then 
		   Dim IDS:IDS=KS.FilterIDs(ContentID)
		   Dim Flag:Flag=true '取"复制(n)"样式
		  Dim RS, IRS, NewID,OriTitle, SqlStr,I,Intro,PhotoUrl
		  Set RS = Server.CreateObject("Adodb.RecordSet")
		  SqlStr = "Select * From " & KS.C_S(ChannelID,2) &" Where ID In(" & IDS & ") And DelTF=0"
		  RS.Open SqlStr, conn, 1, 1
		  If Not RS.EOF Then
		     Do While Not RS.Eof
				If Flag = True Then OriTitle = GetNewTitle(KS.C_S(ChannelID,2),NewClassID, RS("Title"))
				If OriTitle="" Then OriTitle = RS("Title")
			   Set IRS = Server.CreateObject("Adodb.RecordSet")
			   IRS.Open "Select top 1 * From " & KS.C_S(ChannelID,2) &" Where 1=0", conn, 1, 3
				IRS.AddNew
				For I=2 To RS.Fields.Count-1
				 IRS(I)=RS(I)
				Next
				If ChannelID=5 Then
				 IRS("ProID")=KS.GetInfoID(5)
				End If
				IRS("Title") = OriTitle
				IRS("Tid")   = NewClassID
				IRS("DelTF") = 0
				IRS.Update
				IRS.MoveLast
				NewID=IRS("ID")
				IRS("Fname")=NewID & Mid(Trim(RS("Fname")), InStrRev(Trim(RS("Fname")), "."))
				IRS.Update
				
				select case Cint(KS.C_S(ChannelID,6))
				 case 1 Intro=RS("Intro")
				 case 2 Intro=RS("PictureContent")
				 case 3 Intro=RS("DownContent")
				 case 4 Intro=RS("FlashContent")
				 case 5 Intro=RS("ProIntro")
				 case 7 Intro=RS("MovieContent")
				 case 8 Intro=RS("GQContent")
				end select
				Call LFCls.AddItemInfo(ChannelID,NewID,OriTitle,NewClassID,Intro,RS("KeyWords"),RS("PhotoUrl"),Now,KS.C("AdminName"),RS("Hits"),RS("HitsByDay"),RS("HitsByWeek"),RS("HitsByMonth"),RS("Recommend"),RS("Rolls"),RS("Strip"),RS("Popular"),RS("Slide"),RS("IsTop"),RS("Comment"),RS("Verific"),IRS("Fname"))
				IRS.Close
			  RS.MoveNext
			Loop
		  End If
		  RS.Close:Set RS = Nothing:Set IRS = Nothing
		 End If
		End Sub
		
		'得到复制的名称
		Function GetNewTitle(TableName,NewClassID, OriTitle)
			Dim RSC, CheckRS
			On Error Resume Next
			Set CheckRS=Conn.Execute("Select Title From " & TableName & " Where TID='" & NewClassID & "' And Title='" & OriTitle & "' And DelTF=0")
			  If Not CheckRS.EOF Then
				 Set RSC=Server.Createobject("Adodb.recordset")
				 RSC.Open "Select Title From " & TableName & " Where TID='" & NewClassID & "' And Title Like '复制%" & OriTitle & "' And DelTF=0 Order By ID Desc",conn,1,1
				 If Not RSC.EOF Then
					RSC.MoveFirst
					If RSC.RecordCount = 1 Then
					   RSC.Close:Set RSC = Nothing:CheckRS.Close:Set CheckRS = Nothing
					  GetNewTitle = "复制(1) " & OriTitle
					  Exit Function
					Else
					  GetNewTitle = "复制(" & CInt(Left(Split(RSC("Title"), "(")(1), 1)) + 1 & ") " & OriTitle
					End If
					 CheckRS.Close:RSC.Close:Set RSC = Nothing: Set CheckRS = Nothing
				 Else
				  RSC.Close:Set RSC = Nothing:CheckRS.Close:Set CheckRS = Nothing
				  GetNewTitle = "复制 " & OriTitle
				  Exit Function
				 End If
				 RSC.Close:Set RSC = Nothing
			  Else
				CheckRS.Close:Set CheckRS = Nothing
				GetNewTitle = OriTitle
				Exit Function
			  End If
		End Function
		'====================================================复制操作结束==================================================

		'====================================================回收站及彻底删除处理===========================================
		 Sub RefreshHtml(ChannelID,Param,Flag,Tips)
		 	'===========生成列表页和首页====================================
			 If KS.C_S(ChannelID,7)=1 or Split(KS.Setting(5),".")(1)<>"asp" Then
			    response.write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'><meta http-equiv=Content-Type content='text/html; charset=utf-8'>"
				Response.Write "<Br><br><br><table align='center' width='95%' height='200' class='ctable' cellpadding='1' cellspacing='1'><tr class='sort'><td  height='36' colspan=2>系统操作提示信息</td></tr>    <tr class='tdbg'><td align='center'><img src='images/succeed.gif'></td><td  style='font-size:16px;color:red;font-weight:bold'><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;恭喜，" & Tips & KS.C_S(ChannelID,3) & "成功,正在执行相关页面的生成操作,5秒钟后自动返回！</b><br><div style='margin-top:15px;border: #E7E7E7;height:220; overflow: auto; width:100%'><div>"
				  If KS.C_S(ChannelID,7)=1 Then
					If Flag=0 Then
						 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
						 RS.Open "Select tid From " & KS.C_S(ChannelID,2) & " where " & Param,Conn,1,1
						 Do While Not RS.Eof 
						   Dim I,FolderIDArr:FolderIDArr=Split(left(KS.C_C(rs(0),8),Len(KS.C_C(rs(0),8))-1),",")
							For I=0 To Ubound(FolderIDArr)
								Response.Write "<iframe src=""Include/RefreshHtmlSave.Asp?ChannelID=" & ChannelID &"&Types=Folder&RefreshFlag=ID&FolderID=" & FolderIDArr(i) &""" width=""100%"" height=""90"" frameborder=""0"" allowtransparency='true'></iframe>"
							Next
						 RS.MoveNext
						 Loop
						 RS.CLose
						 Set RS=Nothing
					Else
					   Dim N,IDSArr
					   IDSArr=Split(Param,",")
					   For N=0 To Ubound(IDSArr)
					        FolderIDArr=Split(left(KS.C_C(IDSArr(n),8),Len(KS.C_C(IDSArr(n),8))-1),",")
							For I=0 To Ubound(FolderIDArr)
								Response.Write "<iframe src=""Include/RefreshHtmlSave.Asp?ChannelID=" & ChannelID &"&Types=Folder&RefreshFlag=ID&FolderID=" & FolderIDArr(i) &""" width=""100%"" height=""90"" frameborder=""0"" allowtransparency='true'></iframe>"
							Next
					   Next
					End If
				 End If
				 
				 
				 If Split(KS.Setting(5),".")(1)<>"asp" Then
				   If Not KS.ReturnPowerResult(0, "KMTL20000") Then
				    response.write "<div align=center>由于您没有发布首页的权限，所以网站首页没有生成！</div>"
				   Else
					response.Write "<div align=center><iframe src=""Include/RefreshIndex.asp?ChannelID=" & ChannelID &"&RefreshFlag=Info"" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
				   End If
				 End If
				 
				 response.write "</div></div></td></tr>	  <tr>		<td  class='tdbg' height='25' align='center' colspan=2><input type='button' value=' 返 回 ' onclick=""location.href='" &Request.ServerVariables("HTTP_REFERER") & "';"" class='button'/></td>	  </tr>	</table> "
			    response.write "<script>setTimeout(function(){location.href='" & Request.ServerVariables("HTTP_REFERER") & "';},5000);</script>"
				 
			 End If
			 
			 If KS.C_S(ChannelID,7)<>1 and  Split(KS.Setting(5),".")(1)="asp" Then
			   Response.Redirect Request.ServerVariables("HTTP_REFERER")
			 End If
			'===============================================================

		 End Sub
		 
		 '放入回收站
		 Sub Recely(ChannelID)
		    If KS.ChkClng(Split(KS.C_S(KS.G("ChannelID"),46)&"||||","|")(5))=1 Then  '关闭回收站
			  Call DelBySelect(ChannelID) 
			Else
				Conn.Execute("Update [KS_ItemInfo] Set DelTF=1 where ChannelID=" & ChannelID & " and Infoid in(" & KS.FilterIDs(KS.S("ID")) & ")")
				Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set DelTF=1 where id in(" & KS.FilterIDs(KS.S("ID")) & ")")
				Call RefreshHtml(ChannelID,"id in(" & KS.FilterIDs(KS.S("ID")) & ")",0,"放入回收站")
			End If
		 End Sub
		 '回收站还原
		 Sub RecelyBack(ChannelID)
			Conn.Execute("Update [KS_ItemInfo] Set DelTF=0 where ChannelID=" & ChannelID & " and Infoid in(" & KS.FilterIDs(KS.S("ID")) & ")")
			Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set DelTF=0 where id in(" & KS.FilterIDs(KS.S("ID")) & ")")
			Call RefreshHtml(ChannelID,"id in(" & KS.FilterIDs(KS.S("ID")) & ")",0,"还原")
		 End Sub
		 
		 '清空回收站
		 Sub DeleteAll()
		   If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Dim ModelXML,Node
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.documentElement.SelectNodes("channel")
			     If Node.SelectSingleNode("@ks21").text="1" and Node.SelectSingleNode("@ks0").text<>"6" and Node.SelectSingleNode("@ks0").text<>"9" and Node.SelectSingleNode("@ks0").text<>"10" Then
				 Call DelModelInfo(Node.SelectSingleNode("@ks0").text,"Select ID From " & Node.SelectSingleNode("@ks2").text & " Where Deltf=1")
				 End If
			    Next
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		 End Sub
		 '删除选中模型信息操作
		Sub DelBySelect(ChannelID)
			Dim Tids,RS:Set RS=Conn.Execute("Select Tid FROM " & KS.C_S(ChannelID,2) &" Where ID in(" & KS.FilterIDs(Request("ID")) & ")")
			If NOT RS.Eof Then
			  Do While Not RS.Eof
			   If Tids="" Then Tids=RS(0) Else Tids=Tids & "," & RS(0)
			   RS.MoveNext
			  Loop
			End If
			RS.CLose:Set RS=Nothing
			Call DelModelInfo(ChannelID,Request("ID"))
			
			Call RefreshHtml(ChannelID,Tids,1,"彻底删除")
		End Sub
		 
		 '删除信息
		 Sub DelModelInfo(ChannelID,NewsID)
			  Dim K, CurrPath,FolderID,N,ImgSrcArr,RS
			  Dim ContentPageArr, TotalPage, I, CurrPathAndName, FExt, Fname
			  conn.Execute ("Delete From KS_ItemInfo Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")")
			  conn.Execute ("Delete From KS_ItemInfoR Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")")
			  conn.Execute ("Delete From KS_Comment Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")")
			  conn.Execute ("Delete From KS_SpecialR Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")")
			  conn.Execute ("Delete From KS_Digg Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")")
			  conn.Execute ("Delete From KS_DiggList Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")")
			  conn.Execute ("Delete From KS_GuestBook Where ChannelID<>0 And InfoID<>0 and ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")")
			  
			  
			  '===============5-18 删除下载模型的附件===========================
			  If KS.C_S(ChannelID,6)=3 Then
			  Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "Select DownUrls From " & KS.C_S(ChannelID,2) & " Where ID in(" & NewsID & ")",conn,1,1
			  Do While NOt RS.Eof
			    Dim DownUrls:DownUrls=RS(0)
				Dim DownArr,DownItemArr,DownUrl
				If Not KS.IsNul(DownUrls) Then
				    DownArr=Split(DownUrls,"|||")
					For K=0 To Ubound(DownArr)
					  DownItemArr = Split(DownArr(k),"|")
					  DownUrl = Replace(DownItemArr(2),KS.Setting(2),"")
					  Call KS.DeleteFile(DownUrl)  '删除
					Next
				End If
				RS.MoveNext
			  Loop
			  RS.Close
			  End If
			  '=============================================================

			  
			  
			  Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "Select FileName From KS_UploadFiles Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")",Conn,1,1
			  Do While Not RS.Eof
			   if conn.execute("select top 1 filename From KS_UploadFiles Where InfoID not in(" & NewsID & ") and FileName like '%" & RS("FileName") & "%'").eof Then
			    Call KS.DeleteFile(RS(0))
			   end if
			   RS.MoveNext
			  Loop
			  RS.Close
			  conn.Execute ("Delete From KS_UploadFiles Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")")
			  
			  If ChannelID=5 Then  '商城删除订单
			     Conn.Execute("Delete From KS_OrderItem Where ProID in(" & NewsID & ")")
				 conn.execute("Delete From KS_ShopBundleSale Where ProID in(" &NewsID &")")
				 conn.execute("Delete From KS_ShopSpecificationPrice Where ProID in(" &NewsID &")")
				 On error resume next
				 Set RS=Conn.Execute("Select SmallPicUrl,BigPicUrl From KS_ProImages Where ProID in(" & NewsID & ")")
				 Do While Not RS.Eof
				  Call KS.DeleteFile(RS(0))
				  Call KS.DeleteFile(RS(1))
				 RS.MoveNext
				 Loop
				 RS.Close:Set RS=Nothing
				 Conn.Execute("Delete From KS_ProImages Where ProID in(" & NewsID & ")")
			  End IF
			  
			  Set RS=Server.CreateObject("ADODB.Recordset")
			  RS.Open "Select * FROM " & KS.C_S(ChannelID,2) &" Where ID in(" & NewsID & ")", conn, 1, 1
			  Do While Not RS.EOF 
				 FolderID = Trim(RS("Tid"))
				 
				 If KS.C_S(ChannelID,6)=1 Then
				  ContentPageArr = Split(RS("ArticleContent")&" ", "[NextPage]")
				 ElseIf KS.C_S(ChannelID,6)=2 Then
				  ContentPageArr = Split(RS("PicUrls"), "|||")
				 End If
				 Call DelInfoFile(ChannelID,FolderID,ContentPageArr,RS("Fname"),RS("ID"))
			 RS.MoveNext
			Loop
			  RS.Close
			Set RS = Nothing
			conn.execute("delete  FROM " & KS.C_S(ChannelID,2) &" Where ID in(" & NewsID & ")")
		End Sub
		
		'参数:ChannelID-模型id,FolderID-栏目ID,ContentPageArr-分页数组，FileName-文件名
		Sub DelInfoFile(ChannelID,FolderID,ContentPageArr,FileName,InfoID)
		        on error resume next
		 		Dim CurrPath,FExt,Fname,TotalPage,I,CurrPathAndName
				CurrPath = KS.LoadFsoContentRule(ChannelID,FolderID,InfoID)		 
				FExt = Mid(Trim(FileName), InStrRev(Trim(FileName), "."))    '分离出扩展名
				Fname = Replace(Trim(FileName), FExt, "")                    '分离出文件名 如 2005/9-10/1254ddd
				  		 
	    		  If IsArray(ContentPageArr) Then TotalPage = UBound(ContentPageArr) + 1 Else TotalPage=1
				  If TotalPage > 1 and  KS.C_S(ChannelID,6)<=2 Then
					For I = LBound(ContentPageArr) To UBound(ContentPageArr)
					 If I = 0 Then
					  CurrPathAndName = CurrPath & FileName
					 Else
					  CurrPathAndName = CurrPath & Fname & "_" & (I + 1) & FExt
					 End If
					 Call KS.DeleteFile(CurrPathAndName)
					Next
				  Else
				   CurrPathAndName = CurrPath & FileName
				   Call KS.DeleteFile(CurrPathAndName)
				  End If
		End Sub
		 '======================================================回收站/删除结束=========================================
		 
		 '======================================================审核投稿开始============================================
		  '批量审核
		 Sub VerificAll(ChannelID)
		  Dim InputerStr,Inputer,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		   InputerStr="Inputer"
		  RS.Open "Select " & InputerStr & ",Title,Verific,ID From " & KS.C_S(ChannelID,2) & " Where Verific<>2 And ID In(" & KS.FilterIDs(KS.G("ID")) & ")",Conn,1,3
		  Do While Not RS.Eof
			 Inputer=RS(0)
			 IF Inputer<>"" And Inputer<>KS.C("AdminName") Then Call KS.SignUserInfoOK(ChannelID,Inputer,RS(1),RS(3))
			 RS("Verific")=1
			 RS.Update
			 RS.MoveNext
		  Loop
		  RS.Close :Set RS=Nothing
		  Conn.Execute("Update [KS_ItemInfo] Set Verific=1 Where Verific<>2 and channelid=" & ChannelID & " And InfoID In(" & KS.FilterIDs(KS.G("ID")) & ")")
		  Response.Redirect Request.ServerVariables("HTTP_REFERER")
		 End Sub
		 '批量退稿
		 Sub Tuigao(ChannelID)
		  Dim RS,Content
		  Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select * From " & KS.C_S(ChannelID,2) & " Where Verific<>1 And ID In(" & KS.FilterIDs(KS.G("ID")) & ")",conn,1,3
		  Do While Not RS.Eof
		   RS("Verific")=3
		   RS.Update
		   If Request("Email")="1" Then
		   Content=Request("AnnounceContent")
		   Content=Replace(Content,"{$Title}",RS("Title"))
		   Content=Replace(Content,"{$UserName}",RS("Inputer"))
		   Call KS.SendInfo(RS("Inputer"),KS.Setting(0),"退稿通知",Content)
		   End If
		   RS.MoveNext
		  Loop
		  RS.Close
		  Set RS=Nothing
		  Conn.Execute("Update [KS_ItemInfo] Set Verific=3 Where Verific<>1 and channelid=" & ChannelID & " And InfoID In(" & KS.FilterIDs(KS.G("ID")) & ")")
		  Response.Redirect Request.ServerVariables("HTTP_REFERER")
		 End Sub
	 '======================================================审核投稿结束============================================
			
	Sub BatchSet(ChannelID)
		  Dim NID:NID=KS.FilterIDs(KS.G("ID"))
		  Select Case (KS.ChkClng(KS.S("SetAttributeBit")))
		    Case 1
				Conn.Execute("Update [KS_ItemInfo] Set Recommend=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
				Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Recommend=1 where id in(" & NID & ")")
			Case 2
				Conn.Execute("Update [KS_ItemInfo] Set Slide=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Slide=1 where id in(" & NID & ")")
			Case 3
			    Conn.Execute("Update [KS_ItemInfo] Set Popular=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Popular=1 where id in(" & NID & ")")
			Case 4
			    Conn.Execute("Update [KS_ItemInfo] Set Comment=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Comment=1 where id in(" & NID & ")")
			Case 5
			    Conn.Execute("Update [KS_ItemInfo] Set strip=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set strip=1 where id in(" & NID & ")")
			Case 6
			    Conn.Execute("Update [KS_ItemInfo] Set istop=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set istop=1 where id in(" & NID & ")")
			Case 7
			    Conn.Execute("Update [KS_ItemInfo] Set rolls=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set rolls=1 where id in(" & NID & ")")
		    Case 8
			    Conn.Execute("Update [KS_ItemInfo] Set Recommend=0 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Recommend=0 where id in(" & NID & ")")
			Case 9
			    Conn.Execute("Update [KS_ItemInfo] Set Slide=0 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Slide=0 where id in(" & NID & ")")
			Case 10
			    Conn.Execute("Update [KS_ItemInfo] Set Popular=0 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Popular=0 where id in(" & NID & ")")
			Case 11
			    Conn.Execute("Update [KS_ItemInfo] Set Comment=0 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Comment=0 where id in(" & NID & ")")
			Case 12
			    Conn.Execute("Update [KS_ItemInfo] Set strip=0 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
				Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set strip=0 where id in(" & NID & ")")
			Case 13
			    Conn.Execute("Update [KS_ItemInfo] Set istop=0 where id in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set istop=0 where id in(" & NID & ")")
			Case 14
			    Conn.Execute("Update [KS_ItemInfo] Set rolls=0 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set rolls=0 where id in(" & NID & ")")
			Case 15
			    Conn.Execute("Update [KS_ItemInfo] Set Verific=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Verific=1 where id in(" &NID& ")")
			Case 16
			    Conn.Execute("Update [KS_ItemInfo] Set Verific=0 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Verific=0 where id in(" &NID& ")")
		  End Select
		  Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub
		
		Public Sub AddToSpecial(ChannelID)
		Dim NewsID:NewsID = Trim(Request("NewsID"))
		With KS
		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv='Content-Type' content='text/html; chaRSet=utf-8'>"
		.echo "<title>加入到专题</title>"
		.echo "<link href='Include/Admin_Style.css' rel='stylesheet'>"
		.echo "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		.echo "</head>"
		.echo "<body style=""background: #EAF0F5;"" topmargin='0' leftmargin='0' scroll=auto>"
		.echo "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
		.echo "  <form name='specialform' action='?ChannelID=" & ChannelID&"&Action=Special' method='post'>"
		.echo "  <input type='hidden' value='Add' Name='Flag'>"
		.echo "  <input type='hidden' name='SpecialName'>"
		.echo "  <input type='hidden' value='" & NewsID & "' Name='NewsID'>"
		.echo "  <tr>"
		.echo "    <td height='18'>&nbsp;</td>"
		.echo "  </tr>"
		.echo "  <tr>"
		.echo "    <td height='30' align='center'> <strong>请选择一个或多个专题</strong><br>"
		.echo "      <select name='SpecialID'  multiple style='height:340px;width:260px;'>"
		.echo KS.ReturnSpecial("")
		.echo "      </select><br><font color=blue>提示：按住""CTRL""或""Shift""键可以进行多选</font>"
		.echo "    <br/><label><input onclick=""alert('提示：选中后原来文章所属的专题会被清掉!');"" type='checkbox' name='delzt' value='0'>删除原来归属的专题</label></td>"
		.echo "  </tr>"
		.echo "  <tr align='center'>"
		.echo "   <td height='30'> <input type='button' class='button' name='button1' value='加入专题' onclick='CheckForm()'>"
		.echo "      &nbsp; <input type='button' class='button' onclick='window.close();' name='button2' value=' 取消 '>"
		.echo "    </td>"
		.echo "  </tr>"
		.echo "  </form>"
		.echo "</table>"
		.echo "</body>"
		.echo "</html>"
		.echo "<Script>"
		.echo "function CheckForm()"
		.echo "{"
		'.echo " if (document.specialform.SpecialID.value=='0')"
		'.echo "  { alert('对不起,您没有选择专题名称!');"
		'.echo "     document.specialform.SpecialID.focus();"
		'.echo "     return false;"
		'.echo "  }"
		'.echo " document.specialform.SpecialName.value=document.specialform.SpecialID.options[document.specialform.SpecialID.selectedIndex].text;"
		.echo "  document.specialform.submit();"
		.echo "  return true"
		.echo "}"
		.echo "</Script>"
		
		If Request.Form("Flag") = "Add" Then
		   Dim SpecialID, NewsIDArr, K,I
		   SpecialID = Replace(Request.Form("SpecialID")," ","")
		   
		   NewsID=KS.FilterIDs(NewsID)
		  If NewsID<>"" Then 
		   Dim NArr:Narr=Split(NewsID,",")
		   SpecialID= Split(SpecialID,",")
		   For K=0 To Ubound(NArr)
		     If KS.ChkClng(Request("delzt"))=1 Then
		      Conn.Execute("Delete From KS_SpecialR Where InfoID=" & NArr(K) & " and channelid=" & ChannelID)
			 End If
			 For I=0 To Ubound(SpecialID)
			  If Conn.Execute("select top 1 * from KS_SpecialR Where SpecialID=" & SpecialID(I) &" And InfoID=" & NArr(K) & " And ChannelID=" & ChannelID).eof Then
			 Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & NArr(K) & "," & ChannelID & ")")
			  End If
			 Next
		   Next
		 End If  

		  .echo ("<script>alert('操作成功!');window.close();</script>")
         
		End If
		 End With
		End Sub
		
		'SEO优化选项
		Sub LoadSeoOption(ChannelID,OptionTitle,SEOTitle,SEOKeyWord,SEODescript)
		  With KS
		    .echo " <div class=tab-page id=seo-page>"
		    .echo "  <H2 class=tab>" & OptionTitle & "</H2>"
		    .echo "	<SCRIPT type=text/javascript>"
		    .echo "		tabPane1.addTabPage( document.getElementById( ""seo-page"" ) );"
		    .echo "	</SCRIPT>"

            .echo "<TABLE style='margin:1px' width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"
			.echo "           <tr class='tdbg'>"
			.echo "              <td class='clefttitle' style='height:30px;width:100px;text-align:right'><strong>" & KS.C_S(ChannelID,3) & "页面标题:</strong></td>"
			.echo "              <td><input type='text' maxlength='255' name='SEOTitle' size='40' value='" & SEOTitle & "' class='textbox' /> <span class='tips'>留空则默认显示" & KS.C_S(ChannelID,3) &"名称,模板里请用标签{$GetSEOTitle}调用。</span></td>"
			.echo "           </tr>"
			.echo "           <tr class='tdbg'>"
			.echo "              <td class='clefttitle' style='height:30px;width:100px;text-align:right'><strong>页面关键字:</strong>(meta_keywords)</td>"
			.echo "              <td><textarea name='SEOKeyWord' class='textbox' style='width:300px;height:60px'>" & SEOKeyWord & "</textarea><span class='tips'>留空则默认显示" & KS.C_S(ChannelID,3) &"里设置的KeyWords,模板里请用标签{$GetSEOKeyWords}调用。</span></td>"
			.echo "           </tr>"
			.echo "           <tr class='tdbg'>"
			.echo "              <td class='clefttitle' style='height:30px;width:100px;text-align:right'><strong>页面描述:</strong>(meta_description)</td>"
			.echo "              <td><textarea name='SEODescript' class='textbox' style='width:300px;height:60px'>" & SEODescript & "</textarea><span class='tips'>留空则默认显示" & KS.C_S(ChannelID,3) &"简介,模板里请用标签{$GetSEODescription}调用。</span></td>"
			.echo "           </tr>"
			.echo "</table>"
			.echo "</div>"
         End With
		End Sub
		
		'收费选项
		Sub LoadChargeOption(ChannelID,ChargeType,InfoPurview,arrGroupID,ReadPoint,PitchTime,ReadTimes,DividePercent)
		  With KS
		     Dim ModelChargeType:ModelChargeType=.ChkClng(.C_S(ChannelID,34))
		     Dim ChargeStr
			 If ModelChargeType=0 Then 
			   ChargeStr=.Setting(45)
			 ElseIf ModelChargeType=1 Then
			   ChargeStr="资金"
			 Else
			   ChargeStr="积分"
			 End If
		    .echo " <div class=tab-page id=poweroption-page>"
			.echo "  <H2 class=tab>权限选项</H2>"
			.echo "	<SCRIPT type=text/javascript>"
			.echo "				 tabPane1.addTabPage( document.getElementById( ""poweroption-page"" ) );"
			.echo "	</SCRIPT>"
				
			 .echo "<TABLE style='margin:1px' width='100%' align='center' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"	
			 .echo "             <tr  class='tdbg'>"
			 .echo "               <td align='right' width='100'  class='clefttitle' height=30><strong>阅读权限:</strong></td>"
			 .echo "                <td height='30' nowrap> "
			 .echo "                <label><input name='InfoPurview' onclick=""$('#sGroup').hide();"" type='radio' value='0'"
			 if InfoPurview=0 Then .echo " checked"
			 .echo ">继承栏目权限（当所属栏目为认证栏目时，建议选择此项）</label><br>"
			 .echo "            <label><input name='InfoPurview' onclick=""$('#sGroup').hide();"" type='radio' value='1'"
			 If InfoPurview=1 Then .echo " checked"
			 .echo ">所有会员（当所属栏目为开放栏目，想单独对某些" & KS.C_S(ChannelID,3) & "进行阅读权限设置，可以选择此项）</label><br/>"
			 .echo "            <label><input name='InfoPurview' onclick=""$('#sGroup').show();"" type='radio' value='2'" 
			 IF InfoPurview=2 Then .echo " Checked"
			 .echo ">指定会员组（当所属栏目为开放栏目，想单独对某些" & KS.C_S(ChannelID,3) & "进行阅读权限设置，可以选择此项）</label><br/>"
			 .echo "<table border='0' align=center width='90%'>"
			 .echo " <tr>"
			 IF InfoPurview=2 Then
			 .echo " <td id='sGroup'>"
			 Else
			 .echo "<td id='sGroup' style='display:none'>"
			 End If
			 .echo KS.GetUserGroup_CheckBox("GroupID",arrGroupID,5)
			 .echo " </td>"
			 .echo "  </tr></table>"
			 .echo "                </td>"
             .echo "               </tr>"
			 .echo "             <tr  class='tdbg'>"
			 .echo "               <td align='right'  class='clefttitle' height=30><strong>阅读" & ChargeStr & ": </strong></td>"
			 .echo "                <td height='30' nowrap> &nbsp;"
			 .echo "                <input style='text-align:center' name='ReadPoint' type='text' id='ReadPoint'  value='" & ReadPoint & "' size='6' class='textbox'> 　免费阅读请设为 ""<font color=red>0</font>""，否则有权限的会员阅读此" & KS.C_S(ChannelID,3) & "时将消耗相应" & ChargeStr & "，游客将无法阅读此" & KS.C_S(ChannelID,3) & ""
			 .echo "                 </td>"
             .echo "               </tr>"
			 .echo "             <tr  class='tdbg'>"
			 .echo "               <td align='right' class='clefttitle' height=30><strong>重复收费:</strong></td>"
			 .echo "                <td height='30' nowrap> "
			 .echo "                <input name='ChargeType' type='radio' value='0' "
			 IF ChargeType=0 Then .echo " checked"
			 .echo" >不重复收费(如果需扣" & ChargeStr & "" & KS.C_S(ChannelID,3) & "，建议使用)<br>"
			 .echo "<input name='ChargeType' type='radio' value='1'"
			 IF ChargeType=1 Then .echo " checked"
			 .echo ">距离上次收费时间 <input name='PitchTime' type='text' class='textbox' value='" & PitchTime & "' size='8' maxlength='8' style='text-align:center'> 小时后重新收费<br>            <input name='ChargeType' type='radio' value='2'"
			 IF ChargeType=2 Then .echo " checked"
			 .echo ">会员重复阅读此" & KS.C_S(ChannelID,3) & " <input name='ReadTimes' type='text' class='textbox' value='" & ReadTimes & "' size='8' maxlength='8' style='text-align:center'> 页次后重新收费<br>            <input name='ChargeType' type='radio' value='3'"
			 IF ChargeType=3 Then .echo " checked"
			 .echo ">上述两者都满足时重新收费<br>            <input name='ChargeType' type='radio' value='4'"
			 IF ChargeType=4 Then .echo " checked"
			 .echo ">上述两者任一个满足时就重新收费<br>            <input name='ChargeType' type='radio' value='5'"
			 IF ChargeType=5 Then .echo " checked"
			 .echo ">每阅读一页次就重复收费一次（建议不要使用,多页" & KS.C_S(ChannelID,3) & "将扣多次" & ChargeStr & "）"
			 .echo "                 </td>"
             .echo "               </tr>"
			 .echo "             <tr  class='tdbg' style=""display:none"">"
			 .echo "               <td align='right' width='80'  class='clefttitle' height=30><strong>分成比例: </strong></td>"
			 .echo "                <td height='30' nowrap> &nbsp;"
			 .echo "                <input name='DividePercent' type='text' id='DividePercent'  value='" & DividePercent & "' size='6' class='textbox'>% 　如果比例大于0，则将按比例把向阅读者收取的点数支付给投稿者 "
			 .echo "                 </td>"
             .echo "               </tr>"            
			 .echo "    </TABLE>"
			 .echo "  </div>"
		  End With
		End Sub
		
		'相关选项
		Sub LoadRelativeOption(ChannelID,ID)
		    %>
			<script language="javascript">
			$(document).ready(function(){
			 <!--- 相关信息---->
			  $('#relativeButton').click(function(){
			   GetRealtiveItem();
			  });
	          $('#RAddButton').click(function(){
			   var alloptions = $("#TempInfoList option");
			   var so = $("#TempInfoList option:selected");
			   var a = (so.get(so.length-1).index == alloptions.length-1)? so.prev().attr("selected",true):so.next().attr("selected",true);
                
				if (!$("#SelectInfoList option[value="+so.val()+"]").attr("selected")){
				
					var txt=$('#relativeText option:selected').text();
					if (txt!=''){
					so.each(function(){
					  $(this).text(txt+'↓'+$(this).text());
					  $(this).val($(this).val()+"↓"+txt);
					});}
					$("#SelectInfoList").append(so);
				 }else{
				 so.remove();}
			  });
			  
			  $('#RAddMoreButton').click(function(){
			     $("#TempInfoList option").each(function(){
				  if ($("#SelectInfoList option[value="+$(this).val()+"]").attr("selected")){ $(this).remove() }
				 });
				 var so=$("#TempInfoList option").attr("selected",true);
					var txt=$('#relativeText option:selected').text();
					if (txt!=''){
					so.each(function(){
					  $(this).text(txt+'↓'+$(this).text());
					  $(this).val($(this).val()+"↓"+txt);
					});}
				 
			    $("#SelectInfoList").append(so);
				
				
			  });
			  $('#RDelButton').click(function(){
			     var alloptions = $("#SelectInfoList option");
				 var so = $("#SelectInfoList option:selected");
				 var a = (so.get(so.length-1).index == alloptions.length-1)? so.prev().attr("selected",true):so.next().attr("selected",true);
			     so.each(function(){
				   var stext=$(this).text();
				   var sval=$(this).val();
				   if (stext.indexOf('↓')!=-1){
				     $(this).text(stext.split('↓')[1]);
					 $(this).val(sval.split('↓')[0]);
				   }
				 });
				  so.remove();
			   
				$("#TempInfoList").append(so);
			  });
			  $('#RDelMoreButton').click(function(){
			    var so=$("#SelectInfoList option");
				
			     so.each(function(){
				   var stext=$(this).text();
				   var sval=$(this).val();
				   if (stext.indexOf('↓')!=-1){
				     $(this).text(stext.split('↓')[1]);
					 $(this).val(sval.split('↓')[0]);
				   }
				 });
			  
			    $("#TempInfoList").append(so);
				
				
			  });
			  
			  });
			
			GetRealtiveItem=function(){
			 $(parent.document).find("#ajaxmsg").toggle("fast");
			 var key=escape($('input[name=RelativeKey]').val());
			 var Rtitle=$('#RelativeTypeTitle').attr("checked");
			 var Rkey=$('#RelativeTypeKey').attr("checked");
			 var ChannelID=$('#ChannelID').val();
			 $.get("../plus/ajaxs.asp", { action: "GetRelativeItem", channelid:ChannelID,key: key,rtitle:"'"+Rtitle+"'",rkey:"'"+Rkey+"'",id:"<%=KS.G("ID")%>"},
			 function(data){
					$(parent.document).find("#ajaxmsg").toggle("fast");
					$("#TempInfoList").empty();
					$("#TempInfoList").append(data);
			  });
			}
			function setrelcategory(str){
			 $("#relativeText").empty();
			 $("#relativeText").append(str);
			 closeWindow();
			 alert('恭喜,相关信息分类已更新!');
			}
			 </script>
			<%
			With KS
		    .echo " <div class=tab-page id=relation-page>"
			.echo "  <H2 class=tab>相关信息</H2>"
			.echo "	<SCRIPT type=text/javascript>"
			.echo "		 tabPane1.addTabPage( document.getElementById( ""relation-page"" ) );"
			.echo "	</SCRIPT>"
			.echo "    <TABLE style='margin:1px' width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"	
			.echo "     <tr>"
			.echo "      <td align='center'><strong>关键字:</strong><input type='text' class='textbox' value='' id='RelativeKey' name='RelativeKey'> <strong>条件:</strong> <label><input type='checkbox' id='RelativeTypeTitle' name='RelativeTypeTitle' value='1'>标题</label> <label><input type='checkbox' name='RelativeTypeKey' id='RelativeTypeKey' value='2' checked>关键字TGA "
			.echo "      <select name='ChannelID' id='ChannelID'>"
			.echo "       <option value='0' style='color:red'>--不指定模型--</option>"
			.LoadChannelOption ChannelID
			.echo "      </select>"
			.echo "  <input class='button' type='button' value=' 查找相似信息 ' id='relativeButton'></td>"
			.echo "     </tr>"
			.echo "     <tr>"
			.echo "      <td align='center'><table border='0' width='90%'>" & vbcrlf
			.echo "           <tr><td>待选信息<br /><select id='TempInfoList' name='TempInfoList' multiple style='width:240px;height:250px'></select></td>" & vbcrlf
			.echo "          <td>"
			.echo "类型:<select name='relativeText' id='relativeText'>"
			Dim Xml,Node:Set Xml=LfCls.GetXMLFromFile("relativetype")
			If IsObject(Xml) Then
				For Each Node In Xml.DocumentElement.SelectNodes("model[@channelid=" & channelid &"]/item")
				 .echo "<option>" & Node.text & "</option>"
				Next
			End If
			.echo "</select> <a href=""javascript:void(0)"" onclick=""new KesionPopup().PopupCenterIframe('设置相关链接分类','KS.RelativeCategory.asp?channelid=" & channelid & "',320,250,'auto')"" style=""color:green"">分类管理</a><br/>"
			.echo "<input type='button' value=' 添加选中 >  ' id='RAddButton' class='button'><br /><br /><input type='button' value=' 全部添加 >> ' id='RAddMoreButton' class='button'><br /><br /><input type='button' value=' < 删除选中  ' id='RDelButton' class='button'><br /><br /><input type='button' value=' << 全部删除 ' id='RDelMoreButton' class='button'></td>"
			.echo "          <td>选中信息<br /><select id='SelectInfoList' name='SelectInfoList' multiple style='width:240px;height:250px'>"
			If ID<>0 Then
				 Dim RArray,I,RSR,SQLStr
				 SQLStr="Select TOP 200 I.ChannelID,I.InfoID,I.Title,r.relativeText From KS_ItemInfo I Inner Join KS_ItemInfoR R On I.InfoID=R.RelativeID Where R.ChannelID=" & ChannelID &"  and R.InfoID=" & ID &" and R.RelativeChannelID=I.ChannelID"
				 
				 Set RSR=Conn.Execute(SQLStr)
				 If Not RSR.Eof Then
				  RArray=RSR.GetRows(-1)
				 End If
				 RSR.Close
				 Set RSR=Nothing
				 If IsArray(RArray) Then
				   For i=0 To Ubound(RArray,2)
				    If Not KS.IsNul(Rarray(3,i)) Then
					.echo "<option value='" & RArray(0,I) & "|" & RArray(1,i) & "↓"& RArray(3,i) & "' selected>"& RArray(3,i) & "↓" & RArray(2,i) & "</option>"
					Else
					.echo "<option value='" & RArray(0,I) & "|" & RArray(1,i) & "' selected>" & RArray(2,i) & "</option>"
					End If
				   Next
				 End If
            End If
			.echo "</select></td></tr>"
			.echo "     </tr>"
			
			.echo "    </TABLE>"
			.echo "  </div>"
		 End With
	End Sub

		
	Sub AddKeyTags(KeyWords)
		     dim i
			 dim trs:set trs=server.createobject("adodb.recordset")
			 dim karr:karr=split(KeyWords,",")
			 for i=0 to ubound(karr)
			 trs.open "select * from ks_keywords where keytext='" & left(karr(i),100) & "'",conn,1,3
			 if trs.eof then
			   trs.addnew
			   trs("keytext")=left(karr(i),100)
			   trs("adddate")=now
			  trs.update
		   end if
			  trs.close
		  next
		   set trs=nothing
	End Sub
		

	Sub ClassAction(ChannelID)
				'KS.Echo "<iframe src=""KS.ClassMenu.asp?action=Create"" frameborder=""0"" width=""0"" height=""0""></iframe>"
'exit sub
				
				 Dim KSR:Set KSR=New Refresh
                 Call KS.CreateListFolder(KS.Setting(3) & KS.Setting(93))
				 Dim SearchJS,FsoPath
				  FsoPath=KS.Setting(3) & KS.Setting(93) & "S_" & KS.C_S(ChannelID,10) & ".js"
				  SearchJS = "<table width=""98%"" border=""0"" align=""center"">" & vbCrLf
				  SearchJS = SearchJS & "<form id=""SearchForm"" name=""SearchForm"" method=""get"" action=""" & KS.Setting(3) & "item/index.asp"">" & vbCrLf
				  SearchJS = SearchJS & "  <tr>" & vbCrLf
				  SearchJS = SearchJS & "    <td align=""center""><select name=""t"">" & vbCrLf
				  
				  select case ks.c_s(channelid,6)
				   case 1
				  SearchJS = SearchJS & "        <option value=""1"">标 题</option>" & vbCrLf
				  SearchJS = SearchJS & "        <option value=""2"">内 容</option>" & vbCrLf
				  SearchJS = SearchJS & "        <option value=""3"">作 者</option>" & vbCrLf
				  SearchJS = SearchJS & "        <option value=""4"">录入者</option>" & vbCrLf
				  SearchJS = SearchJS & "        <option value=""5"">关键字</option>" & vbCrLf
				   case 2
				  SearchJS = SearchJS & "          <option value=""1"">名 称</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">简 介</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">作 者</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""4"">录入者</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">关键字</option>" & vbCrLf
				   case 3
				  SearchJS = SearchJS & "          <option value=""1"">名 称</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">简 介</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">开发商</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""4"">录入者</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">关键字</option>" & vbCrLf
				   case 4
				  SearchJS = SearchJS & "          <option value=""1"">名 称</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">简 介</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">作 者</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""4"">录入者</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">关键字</option>" & vbCrLf
				   case 5
				  SearchJS = SearchJS & "          <option value=""1"">商品名称</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">商品介绍</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">生 产 商</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""4"">录入者</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">商品Tags</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""6"">商品ID</option>" & vbCrLf
				   case 7
				  SearchJS = SearchJS & "          <option value=""1"">影片名称</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">影片介绍</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">影片主演</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""4"">影片上传者</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""6"">影片导演</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">影片Tags</option>" & vbCrLf
				   case 8
				  SearchJS = SearchJS & "          <option value=""1"">信息主题</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">信息介绍</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""4"">发布者</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">关键字</option>" & vbCrLf
				  end select
				  SearchJS = SearchJS & "      </select>" & vbCrLf
				  SearchJS = SearchJS & "        <select name=""tid"" style=""width:150px"">" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""0"" selected=""selected"">所有栏目</option>" & vbCrLf
				  SearchJS = SearchJS & KS.LoadClassOption(ChannelID,false)
				  SearchJS = SearchJS & "        </select>" & vbCrLf
				  SearchJS = SearchJS & "        <input name=""key"" type=""text"" class=""textbox""  value=""关键字"" onfocus=""this.select();""/>" & vbCrLf
				  SearchJS = SearchJS & "        <input name=""ChannelID"" value=""" & channelid & """ type=""hidden"" />" & vbCrLf
				  SearchJS = SearchJS & "        <input type=""submit"" class=""inputButton"" name=""sbtn"" value=""搜 索"" /></td>" & vbCrLf
				  SearchJS = SearchJS & "  </tr>" & vbCrLf
				  SearchJS = SearchJS & "</form>" & vbCrLf
				  SearchJS = SearchJS & "</table>"
				  
				  SearchJS = Replace(Replace(SearchJS,"'","\'"),"""","\""")
				  SearchJS = KSR.ReplaceJsBr(SearchJS)
				  
				  Call KSR.FsoSaveFile(SearchJS,FsoPath)
                  Set KSR=Nothing
			End Sub
End Class
%> 
