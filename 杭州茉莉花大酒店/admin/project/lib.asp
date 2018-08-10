<%
Function FatherName(ClassID,TableName,lianjie_zifu,lianjie,weblink,ziti)
'函数名程:ShowName
'功能描述:返回改分类名称以及父名称   ----没有链接
'传入参数  Cid :已选择ID
Set rs=Server.CreateObject("adodb.recordset")
  sql="Select ParPath,ClassName,ID from "&TableName&" Where ID="&ClassID
  rs.Open Sql,Conn,1,1
  if rs.eof then Exit Function
     If rs(0)=0 Then
	     'FatherID=ClassID
		 'response.Write(FatherID)
		 if cint(lianjie)=1 then
		    FatherNameS="<a href='"&weblink&"?ClassID="&rs(2)&"' title='"&rs(1)&"' class='"&ziti&"'>"&rs(1)&"</a>" 
		   else 
		    FatherNameS=rs(1)
		  end if
		 response.Write FatherNameS 
	 Else
	     FatherIDs=ClassID
	     Set rs_2=Server.CreateObject("Adodb.recordset")
		 'SQL="Select ID From "&TableName&" Where ID in ("&rs(0)&","&ClassID&")"
		 SQL="Select ID From "&TableName&" Where ID in ("&rs(0)&")"
		 'response.Write SQL
		 rs_2.Open Sql,Conn,1,1
		    Do While not rs_2.Eof
			  FatherIDs=rs_2(0)&","&FatherIDs
			rs_2.movenext
			loop
	     rs_2.close
		 set rs_2=Nothing
		 FatherID=FatherIDs
	     Set rs_N=Server.CreateObject("Adodb.recordset")
         SQL_N="Select ClassName,ID From "&TableName&" Where ID in ("&FatherID&") order by Sequence asc"
		 'response.End
		 rs_N.Open SQL_N,Conn,1,1
		    d=1
		    Do While not rs_N.Eof
			if d=1 then
		      if cint(lianjie)=1 then
		         FatherNameS="<a href='"&weblink&"?ClassID="&rs_N(1)&"' title='"&rs_N(0)&"' class='"&ziti&"'>"&rs_N(0)&"</a>" 
		       else 
			     FatherNameS=rs_N(0)
		      end if
			else
		      if cint(lianjie)=1 then
		         FatherNameS=lianjie_zifu&"<a href='"&weblink&"?ClassID="&rs_N(1)&"' title='"&rs_N(0)&"' class='"&ziti&"'>"&rs_N(0)&"</a>" 
		       else 
			  FatherNameS=lianjie_zifu&rs_N(0)
		      end if
			end if
			FatherNameSS=FatherNameSS+FatherNameS
			d=d+1
			rs_N.movenext
			loop
	     rs_N.close
		 set rs_N=Nothing
		 response.Write FatherNameSS
	  End If 
   rs.Close
   Set rs=Nothing
 End Function
Private Function FilterSQL(strValue)
'函数名称: FilterSQL
'功能描述: 过滤字符串中的单引号

'使用方法：FilterSQL(strValue)
	FilterSQL=Replace(strValue,"'","''")
End Function

Private Function IsSubmit()
'函数名称: IsSubmit
'功能描述: 判断页面是否提交
'使用方法:如果是提交则返回 True 否则返回 False
'		 If IsSubmit Then
'  		 ...
'		 else
'		 ...
'		 End if
	IsSubmit=Request.ServerVariables("request_method")="POST"
End Function

Function HTMLcode(fString)
'函数名称: HTMLcode
'功能描述: 转换字符为HTML格式
'使用方法：HTMLcode(fString)
	If Not isnull(fString) then
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")
		fString = Replace(fString, CHR(32), "&nbsp;")
		fString = Replace(fString, CHR(9), "&nbsp;")
		fString = Replace(fString, CHR(34), "&quot;")
		fString = Replace(fString, CHR(39), "&#39;")
		fString = Replace(fString, CHR(13), "")
		fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
		fString = Replace(fString, CHR(10), "<BR> ")
		HTMLcode = fString
	End if
End function

Function gotTopic(str,strlen)
'函数名称: gotTopic
'功能描述: 控制字符串显示的长度
'使用方法：gotTopic(str,strlen)
Dim l,t,c
	l=len(str)
	t=0
	If IsNull(str) Then Exit Function
	For i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		If c>255 Then
			t=t+2
		Else
			t=t+1
		End if
		If t >= strlen Then
			gotTopic=left(str,i)
			exit for
		Else
			gotTopic=str&""
		End if
	Next
End function

function cutStr(str,strlen)
'函数名称: cutStr
'功能描述: 限制标题长度
'传入参数:str 字符串;strlen 长度
'传出参数:cutStr
'使用方法：cutStr(str,strlen)
	dim l,t,c
	l=len(str)
    If l>strlen Then
	   cutStr=left(str,strlen-2)&"..."
	Else
	   cutStr=str
	End IF
end function

Private Function GetLongDate(Value)
'函数名称: GetLongDate
'功能描述: '把时间转换为长日期格式格式 与 FormatDateTime函数相似
'使用方法：GetLongDate(Value)
    Dim strYear, strMonth, strDate
    strYear = Year(Value)
    strMonth = Month(Value)
    strDate = Day(Value)
    GetLongDate = strYear & " 年 " & strMonth & " 月 " & strDate & "日"
End Function

Private Function GetFields(Value)
'函数名称: GetFields
'功能描述: 当数据库中字段为空时,返回空
'使用方法：GetFields(Value)
	If IsNull(Value) Then
		GetFields=""
	Else
		GetFields= Value 
	End If
End Function




private function OnlyWord(strng)
'函数名程:OnlyWord
'功能描述:只替换字符串中的图片
'传入参数:strng
'使用方法:OnlyWord(strng)
Set re=new RegExp 
re.IgnoreCase =True 
re.Global=True 

re.Pattern = "(<)(.[^<]*)(src=)('|"&CHR(34)&"| )?(.[^'|\s|"&CHR(34)&"]*)(\.)(jpg|gif|png|bmp|jpeg|swf)('|"&CHR(34)&"|\s|>)(.[^>]*)(>)" '设置模式。 
OnlyWord=re.Replace(strng,"") 
Set re= nothing 
end function 
 
Function RemoveHTML(strHTML)
'函数名程:RemoveHTML
'功能描述:去除字符串中的html代码,包括图片
'传入参数:strHTML
'使用方法:RemoveHTML(strHTML)
Dim objRegExp, Match, Matches 
Set objRegExp = New Regexp 

objRegExp.IgnoreCase = True 
objRegExp.Global = True 
'取闭合的<> 
objRegExp.Pattern = "<.+?>" 
'进行匹配 函数的建立

Set Matches = objRegExp.Execute(strHTML) 

' 遍历匹配集合，并替换掉匹配的项目 
For Each Match in Matches 
strHtml=Replace(strHTML,Match.Value,"") 
Next 
RemoveHTML=strHTML 
Set objRegExp = Nothing 
End Function 


sub ShowPage(Url,TotleNum,NumPerPage,page,ShowJump,pagestyle)
  '函数名：showpage(Url,TotleNum,NumPerPage,ShowJump)
    '作  用：显示分页代码
    '参  数：Url:传递查询参数
    '        TotleNum:总条数
	'        NumPerPage:每页条数
	'        ShowJump:是否显示跳转按钮 (true or false)
	'        pagestyle:  1:上一页下一页    2:多页
	if TotleNum<=NumperPage Then Exit Sub	
    Url=trim(Url)
	'arrurl=Url       '为跳转框实现GET功能
	if Url<>"" then Url=Url&"&"
	Dim strTemp
	if TotleNum mod NumPerPage=0 then
    	n= TotleNum\NumPerPage
  	else
    	n= TotleNum\NumPerPage+1
  	end if	
  	strTemp= "<script language=javascript>function chkUrl(){formx.action=""?"&Url&"page=""+formx.Page.value;return true;}</script><table align='center'  width=""100%""><form name=""formx"" method=""post"" action="""" onSubmit=""return chkUrl()""><tr><td align=""center"">"
	strTemp=strTemp & "页次：" & Page & "/" & n & "页 "
	strTemp=strTemp & NumPerPage & "条/页 "
	strTemp=strTemp & "共" & TotleNum & "条 &nbsp;&nbsp;&nbsp;&nbsp;"	
  select case pagestyle
    case 1
	if Page<2 then
    		strTemp=strTemp & "<font color=""#999999"">首页 上页</font> "
  	else
    		strTemp=strTemp & "<a href='?" & Url & "page=1'  class=""link"">首页</a> "
    		strTemp=strTemp & "<a href='?" & Url & "page=" & (Page-1) & "'  class=""link"">上页</a> "
  	end if
  	if n-Page<1 then
    		strTemp=strTemp & "<font color=""#999999"">下页 尾页</font> "
  	else
    		strTemp=strTemp & "<a href='?" & Url & "page=" & (Page+1) & "'  class=""link"">下页</a> "
    		strTemp=strTemp & "<a href='?" & Url & "page=" & n & "'  class=""link"">尾页</a>  "
  	end if 
  case 2
    if page-1 mod 10=0 then
		p=(page-1) \ 10
	else
		p=(page-1) \ 10
	end if
	if p*10>0 then strTemp=strTemp &"<a href='?" & Url & "page="&p*10&"' title=上十页 >[&lt;&lt;]</a>   "
    uming_i=1
	for ii=p*10+1 to P*10+10
		   if ii=page then  
	         strTemp=strTemp &"<strong><font color=#ff0000>["+Cstr(ii)+"]</font></strong> "
		   else
		     strTemp=strTemp &"<a href='?" & Url & "page="&ii&"'>["+Cstr(ii)+"]</a> "
		   end if
		if ii=n then exit for
		 uming_i=uming_i+1
	next
  	if ii<=n and uming_i=11 then strTemp=strTemp &"<a href='?" & Url & "page="&ii&"' title=下十页>[&gt;&gt;]</a>  "
   end select	 
  
   
	if ShowJump=True then strTemp=strTemp & "  &nbsp;跳至第&nbsp;<input type=text size=3 name=""Page"">页 <input type=""Submit"" name=""Submit"" value=""跳转""  class=""sbe_button""> "
	strTemp=strTemp & "</td></tr></form></table>"
	response.write strTemp
end sub


Function DeleteFile(delfile,filepath) 
'函数名：DeleteFile 
'作  用：删除文件。
'参  数：delfile(要删除的文件名) | filepath (删除路径)
'返回值：无
Set fso = Server.CreateObject("Scripting.FileSystemObject")
   if instr(delfile,"|")>0 then
    dim morefile
    morefile=split(delfile,"|")
    for tempnum=0 to ubound(morefile)
        delfilepath=server.MapPath(filepath&"/"&morefile(tempnum))
	if fso.FileExists(delfilepath) then
	    fso.DeleteFile(delfilepath)	
	end if 
    next
   else
        delfilepath=server.MapPath(filepath&"/"&delfile)
	if fso.FileExists(delfilepath) then
	   fso.DeleteFile(delfilepath)
        end if
   end if
 set fso=nothing
 End Function


function ReturnSel(str1,str2,seltype)
'函数名：ReturnSel
'作  用：下拉框,复选框选择
'参  数：str1 原有值;str2 数据库值;seltype:类型
'返回值：无
select case seltype
         case 1
            if str1=str2 then response.write("selected")
         case 2
            if str1=str2 then response.write("checked")
     end select
end function


Function Judgement(content) '函数的建立

'函数名：judgement 
'作  用：判断是否。
'参  数：content---判断内容
'返回值：√ or ×
if content=true then
   response.Write("<b><font color=#009900>√</font></b>")
  else 
   response.Write("<b><font color=#FF0000>×</font></b>")
  end if
end Function

Function Judgement1(content) '函数的建立
'函数名：judgement1 
'作  用：判断是否。
'参  数：content---判断内容
if content=true then
   response.Write("<b><font color=#009900>中</font></b>")
  else 
   response.Write("<b><font color=#FF0000>英</font></b>")
  end if
end Function

Function Judgement2(content) '函数的建立
'函数名：judgement2 
'作  用：判断是否。
'参  数：content---判断内容
'返回值：√ or ×
if content=true then
   response.Write("代理商")
  else 
   response.Write("专卖店")
  end if
end Function
Private Sub Del(Table_name,ItemID,intID)
'过程名:Del
'功能描述: 删除数据库中的记录
'Table_name数据表名
'     ItemID:字段名
'     intID:ID序号
sql="delete from "&Table_name&" where "&ItemID&" =" &clng(intID)
conn.execute(sql)
End Sub

Private Sub page_back(strValue)
'对数据库修改，删除，添加之后的返回信息
'调用方式 page_back("数据修改成功 返回继续修改")
	response.write("<script>alert('"& strValue &"');this.location.href='"& Request.ServerVariables("HTTP_REFERER") &"';</script>")
End Sub



Function WriteErr(Msg,ErrType)
'********************************************************
'函数名:WriteErr(Msg,ErrType)
'功能 ：显示错误对话框
'参数说明：
'       Msg ---  显示出错的内容
'       ErrType --- 显示类型，"back"：返回  ； "close":关闭
'********************************************************
   Select Case ErrType
       Case 1
	        Response.Write("<script language=""javascript"">alert("""&Msg&""");window.history.back(-1);</script>")
       Case 2
	        Response.Write("<script language=""javascript"">alert("""&Msg&""");window.close();</script>")
   End Select
   Response.End()
End Function


Function ShowClass(ClassTitle,ClassID)
'函数名程:ShowClass
'功能描述:显示分类下拉列表
'传入参数:ClassTitle：分类名 如：Sbe_Product  ;  ClassID :已选择ID
'使用方法:<select name="select">
'          <option>请选择...</option>
'		   <#Call ShowClass("sbe_product",0)#>  '如无已选项则Classid=0
'         </select> 
SClassID=ClassID
        If ClassID="" Then sClassID=0
		sClassID=Cint(sClassID)
	    Set Rs_ShowClass=Server.CreateObject("adodb.recordset")
	    Sql="Select Depth,ClassName,ID from "&ClassTitle&"_Class order by sequence"
		Rs_ShowClass.Open Sql,Conn,1,1
		StrShowClass=""
		  do while not Rs_ShowClass.eof
		  StrShowClass=StrShowClass&"<option value="""&rs_ShowClass("ID")&""""		  
		  if sClassID=rs_ShowClass("id") Then StrshowClass=StrshowClass&" selected"
		  StrShowClass=StrShowClass&">"
		  If Rs_ShowClass("Depth")=0 Then
		     StrShowClass=StrShowClass&"┣"
		  Else
		     For ShowClass_i=1 to Rs_ShowClass("Depth")
			    StrShowClass=StrShowClass&"&nbsp;│"
			 Next
			 StrShowClass=left(StrShowClass,len(StrShowClass)-1)&"├"
		  End If
		  StrShowClass=StrShowClass&Rs_ShowClass("ClassName")
		  
		  StrShowClass=StrShowClass&"</option>"
		  Rs_ShowClass.MoveNext
		  Loop
		  
		  Rs_ShowClass.Close
		  Set Rs_ShowClass=Nothing
		  Response.Write(StrShowClass)
End Function 


Function ShowClassName(ClassTitle,ClassID)
'函数名程:ShowClassName
'功能描述:返回分类名称
'传入参数:ClassTitle：分类名 如：Sbe_Product  ;  ClassID :已选择ID
'使用方法: Tname=ShowClassName("sbe_product",tid)  '如无已选项则Classid=0
If ClassID="" or ClassID="" Then
	    ShowClassName=""
	 ElSE
	      Set Rs_ShowClassName=Conn.execute("Select top 1 ClassName From "&ClassTitle&"_Class Where ID="&ClassID)		  
		  If Not Rs_ShowClassName.Eof Then
		       ShowClassName=Rs_ShowClassName(0)
		  Else
		       ShowClassName=""
		  End If
		  Set Rs_ShowClassName=Nothing
	 End If
End Function 

Function ChildrenID(ClassTitle,ClassID)
'函数名程:ChildrenID
'功能描述:返回改分类下所有子分类及子身ID
'传入参数:ClassTitle：分类名 如：Sbe_Product  ;  ClassID :已选择ID
Set Rs_ChildrenID=Server.CreateObject("adodb.recordset")
  sql="Select ChildNum,ParPath from "&ClassTitle&"_Class Where ID="&ClassID  
  Rs_ChildrenID.Open Sql,Conn,1,1
     If Rs_ChildrenID(0)=0 Then
	     ChildrenID=ClassID
	 Else
	     ChildrenIDs=ClassID
	     Set Rs_ChildrenIDS=Server.CreateObject("Adodb.recordset")
		 SQL="Select ID From "&ClassTitle&"_Class Where ParPath like '"&Rs_ChildrenID(1)&","&ClassID&"%'"
		 Rs_ChildrenIDS.Open Sql,Conn,1,1
		    Do While not Rs_ChildrenIDS.Eof
			  ChildrenIDs=ChildrenIDs&","&Rs_ChildrenIDS(0)
			Rs_ChildrenIDS.movenext
			loop
	     Rs_ChildrenIDS.close
		 set Rs_ChildrenIDS=Nothing
		 ChildrenID=ChildrenIDs
	  End If 
   Rs_ChildrenID.Close
   Set Rs_ChildrenID=Nothing
End Function 


Private Function str_id(parid,tablename)
'=====================================================
'函数名程:str_id
'功能描述:指定ID的所有儿子,孙子,重孙子的ID字段组
'传入参数:parid：分类id ;  tablename :要查询的表单名
'使用方法: response.write str_id(parid,tablename)
'======================================================

parid=parid
tablename=tablename
str=parid&","
Set oRs=Conn.Execute("select ID,parID from "& tablename &" where parID="& parid &" order by id asc")
If (oRs.eof and oRs.bof) Then
 str=parid
Else 
 do while not oRs.eof
   str=str&","&str_ID(oRs("id"),tablename) 
  oRs.Movenext
  Loop
End IF
 IF instr(str,",,")>0 Then  
  str=replace(str,",,",",")
 Else
  str=str
 End IF
str_id=str
oRS.Close:set oRs=Nothing
End Function



Function FormatDate(FormatStr, CurDateTime)
  Dim sTemp,YYYY,YY,MM,DD,HH,mmm,SS
  sTemp = FormatStr
  If IsDate(CurDateTime) Then
    YYYY = Year(CurDateTime)
    YY = Mid(Year(CurDateTime),3,2)
    MM = Month(CurDateTime)
    If CInt(MM) < 10 Then MM = "0"&MM
    DD  = Day(CurDateTime)
    If CInt(DD) < 10 Then DD = "0"&DD
    HH = Hour(CurDateTime)
    If CInt(HH) < 10 Then HH = "0"&DD
    mmm = Minute(CurDateTime)+1
    If CInt(mmm) < 10 Then mmm = "0"&mmm
    SS = Second(CurDateTime)
    If CInt(SS) < 10 Then SS = "0"&SS
    sTemp = Replace(Replace(Replace(Replace(Replace(Replace(Replace(sTemp,"YYYY",YYYY),"YY",YY),"MM",MM),"DD",DD),"HH",HH),"mm",mmm),"SS",SS)
  End If
  If IsDate(sTemp) Then 
    FormatDate = sTemp
  Else 
    FormatDate = CurDateTime
  End If
End Function
%> 

