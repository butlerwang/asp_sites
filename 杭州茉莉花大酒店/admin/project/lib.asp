<%
Function FatherName(ClassID,TableName,lianjie_zifu,lianjie,weblink,ziti)
'��������:ShowName
'��������:���ظķ��������Լ�������   ----û������
'�������  Cid :��ѡ��ID
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
'��������: FilterSQL
'��������: �����ַ����еĵ�����

'ʹ�÷�����FilterSQL(strValue)
	FilterSQL=Replace(strValue,"'","''")
End Function

Private Function IsSubmit()
'��������: IsSubmit
'��������: �ж�ҳ���Ƿ��ύ
'ʹ�÷���:������ύ�򷵻� True ���򷵻� False
'		 If IsSubmit Then
'  		 ...
'		 else
'		 ...
'		 End if
	IsSubmit=Request.ServerVariables("request_method")="POST"
End Function

Function HTMLcode(fString)
'��������: HTMLcode
'��������: ת���ַ�ΪHTML��ʽ
'ʹ�÷�����HTMLcode(fString)
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
'��������: gotTopic
'��������: �����ַ�����ʾ�ĳ���
'ʹ�÷�����gotTopic(str,strlen)
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
'��������: cutStr
'��������: ���Ʊ��ⳤ��
'�������:str �ַ���;strlen ����
'��������:cutStr
'ʹ�÷�����cutStr(str,strlen)
	dim l,t,c
	l=len(str)
    If l>strlen Then
	   cutStr=left(str,strlen-2)&"..."
	Else
	   cutStr=str
	End IF
end function

Private Function GetLongDate(Value)
'��������: GetLongDate
'��������: '��ʱ��ת��Ϊ�����ڸ�ʽ��ʽ �� FormatDateTime��������
'ʹ�÷�����GetLongDate(Value)
    Dim strYear, strMonth, strDate
    strYear = Year(Value)
    strMonth = Month(Value)
    strDate = Day(Value)
    GetLongDate = strYear & " �� " & strMonth & " �� " & strDate & "��"
End Function

Private Function GetFields(Value)
'��������: GetFields
'��������: �����ݿ����ֶ�Ϊ��ʱ,���ؿ�
'ʹ�÷�����GetFields(Value)
	If IsNull(Value) Then
		GetFields=""
	Else
		GetFields= Value 
	End If
End Function




private function OnlyWord(strng)
'��������:OnlyWord
'��������:ֻ�滻�ַ����е�ͼƬ
'�������:strng
'ʹ�÷���:OnlyWord(strng)
Set re=new RegExp 
re.IgnoreCase =True 
re.Global=True 

re.Pattern = "(<)(.[^<]*)(src=)('|"&CHR(34)&"| )?(.[^'|\s|"&CHR(34)&"]*)(\.)(jpg|gif|png|bmp|jpeg|swf)('|"&CHR(34)&"|\s|>)(.[^>]*)(>)" '����ģʽ�� 
OnlyWord=re.Replace(strng,"") 
Set re= nothing 
end function 
 
Function RemoveHTML(strHTML)
'��������:RemoveHTML
'��������:ȥ���ַ����е�html����,����ͼƬ
'�������:strHTML
'ʹ�÷���:RemoveHTML(strHTML)
Dim objRegExp, Match, Matches 
Set objRegExp = New Regexp 

objRegExp.IgnoreCase = True 
objRegExp.Global = True 
'ȡ�պϵ�<> 
objRegExp.Pattern = "<.+?>" 
'����ƥ�� �����Ľ���

Set Matches = objRegExp.Execute(strHTML) 

' ����ƥ�伯�ϣ����滻��ƥ�����Ŀ 
For Each Match in Matches 
strHtml=Replace(strHTML,Match.Value,"") 
Next 
RemoveHTML=strHTML 
Set objRegExp = Nothing 
End Function 


sub ShowPage(Url,TotleNum,NumPerPage,page,ShowJump,pagestyle)
  '��������showpage(Url,TotleNum,NumPerPage,ShowJump)
    '��  �ã���ʾ��ҳ����
    '��  ����Url:���ݲ�ѯ����
    '        TotleNum:������
	'        NumPerPage:ÿҳ����
	'        ShowJump:�Ƿ���ʾ��ת��ť (true or false)
	'        pagestyle:  1:��һҳ��һҳ    2:��ҳ
	if TotleNum<=NumperPage Then Exit Sub	
    Url=trim(Url)
	'arrurl=Url       'Ϊ��ת��ʵ��GET����
	if Url<>"" then Url=Url&"&"
	Dim strTemp
	if TotleNum mod NumPerPage=0 then
    	n= TotleNum\NumPerPage
  	else
    	n= TotleNum\NumPerPage+1
  	end if	
  	strTemp= "<script language=javascript>function chkUrl(){formx.action=""?"&Url&"page=""+formx.Page.value;return true;}</script><table align='center'  width=""100%""><form name=""formx"" method=""post"" action="""" onSubmit=""return chkUrl()""><tr><td align=""center"">"
	strTemp=strTemp & "ҳ�Σ�" & Page & "/" & n & "ҳ "
	strTemp=strTemp & NumPerPage & "��/ҳ "
	strTemp=strTemp & "��" & TotleNum & "�� &nbsp;&nbsp;&nbsp;&nbsp;"	
  select case pagestyle
    case 1
	if Page<2 then
    		strTemp=strTemp & "<font color=""#999999"">��ҳ ��ҳ</font> "
  	else
    		strTemp=strTemp & "<a href='?" & Url & "page=1'  class=""link"">��ҳ</a> "
    		strTemp=strTemp & "<a href='?" & Url & "page=" & (Page-1) & "'  class=""link"">��ҳ</a> "
  	end if
  	if n-Page<1 then
    		strTemp=strTemp & "<font color=""#999999"">��ҳ βҳ</font> "
  	else
    		strTemp=strTemp & "<a href='?" & Url & "page=" & (Page+1) & "'  class=""link"">��ҳ</a> "
    		strTemp=strTemp & "<a href='?" & Url & "page=" & n & "'  class=""link"">βҳ</a>  "
  	end if 
  case 2
    if page-1 mod 10=0 then
		p=(page-1) \ 10
	else
		p=(page-1) \ 10
	end if
	if p*10>0 then strTemp=strTemp &"<a href='?" & Url & "page="&p*10&"' title=��ʮҳ >[&lt;&lt;]</a>   "
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
  	if ii<=n and uming_i=11 then strTemp=strTemp &"<a href='?" & Url & "page="&ii&"' title=��ʮҳ>[&gt;&gt;]</a>  "
   end select	 
  
   
	if ShowJump=True then strTemp=strTemp & "  &nbsp;������&nbsp;<input type=text size=3 name=""Page"">ҳ <input type=""Submit"" name=""Submit"" value=""��ת""  class=""sbe_button""> "
	strTemp=strTemp & "</td></tr></form></table>"
	response.write strTemp
end sub


Function DeleteFile(delfile,filepath) 
'��������DeleteFile 
'��  �ã�ɾ���ļ���
'��  ����delfile(Ҫɾ�����ļ���) | filepath (ɾ��·��)
'����ֵ����
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
'��������ReturnSel
'��  �ã�������,��ѡ��ѡ��
'��  ����str1 ԭ��ֵ;str2 ���ݿ�ֵ;seltype:����
'����ֵ����
select case seltype
         case 1
            if str1=str2 then response.write("selected")
         case 2
            if str1=str2 then response.write("checked")
     end select
end function


Function Judgement(content) '�����Ľ���

'��������judgement 
'��  �ã��ж��Ƿ�
'��  ����content---�ж�����
'����ֵ���� or ��
if content=true then
   response.Write("<b><font color=#009900>��</font></b>")
  else 
   response.Write("<b><font color=#FF0000>��</font></b>")
  end if
end Function

Function Judgement1(content) '�����Ľ���
'��������judgement1 
'��  �ã��ж��Ƿ�
'��  ����content---�ж�����
if content=true then
   response.Write("<b><font color=#009900>��</font></b>")
  else 
   response.Write("<b><font color=#FF0000>Ӣ</font></b>")
  end if
end Function

Function Judgement2(content) '�����Ľ���
'��������judgement2 
'��  �ã��ж��Ƿ�
'��  ����content---�ж�����
'����ֵ���� or ��
if content=true then
   response.Write("������")
  else 
   response.Write("ר����")
  end if
end Function
Private Sub Del(Table_name,ItemID,intID)
'������:Del
'��������: ɾ�����ݿ��еļ�¼
'Table_name���ݱ���
'     ItemID:�ֶ���
'     intID:ID���
sql="delete from "&Table_name&" where "&ItemID&" =" &clng(intID)
conn.execute(sql)
End Sub

Private Sub page_back(strValue)
'�����ݿ��޸ģ�ɾ�������֮��ķ�����Ϣ
'���÷�ʽ page_back("�����޸ĳɹ� ���ؼ����޸�")
	response.write("<script>alert('"& strValue &"');this.location.href='"& Request.ServerVariables("HTTP_REFERER") &"';</script>")
End Sub



Function WriteErr(Msg,ErrType)
'********************************************************
'������:WriteErr(Msg,ErrType)
'���� ����ʾ����Ի���
'����˵����
'       Msg ---  ��ʾ���������
'       ErrType --- ��ʾ���ͣ�"back"������  �� "close":�ر�
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
'��������:ShowClass
'��������:��ʾ���������б�
'�������:ClassTitle�������� �磺Sbe_Product  ;  ClassID :��ѡ��ID
'ʹ�÷���:<select name="select">
'          <option>��ѡ��...</option>
'		   <#Call ShowClass("sbe_product",0)#>  '������ѡ����Classid=0
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
		     StrShowClass=StrShowClass&"��"
		  Else
		     For ShowClass_i=1 to Rs_ShowClass("Depth")
			    StrShowClass=StrShowClass&"&nbsp;��"
			 Next
			 StrShowClass=left(StrShowClass,len(StrShowClass)-1)&"��"
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
'��������:ShowClassName
'��������:���ط�������
'�������:ClassTitle�������� �磺Sbe_Product  ;  ClassID :��ѡ��ID
'ʹ�÷���: Tname=ShowClassName("sbe_product",tid)  '������ѡ����Classid=0
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
'��������:ChildrenID
'��������:���ظķ����������ӷ��༰����ID
'�������:ClassTitle�������� �磺Sbe_Product  ;  ClassID :��ѡ��ID
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
'��������:str_id
'��������:ָ��ID�����ж���,����,�����ӵ�ID�ֶ���
'�������:parid������id ;  tablename :Ҫ��ѯ�ı���
'ʹ�÷���: response.write str_id(parid,tablename)
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

