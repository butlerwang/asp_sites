<%
Function ShowPage(Url,nowpage,totalpage,totalrecord,pagestyle,leibie)            '��׼��ҳ����
'============================'============================'============================
'��������:ShowPage  ��׼��ҳ����
'��������:��׼��ҳ ��ҳ ��һҳ ��һҳ βҳ
'�������:Url:���ӵ�ַ,nowpage:��ǰҳ��,totalpage:��ҳ��,totalrecord:�ܼ�¼��,pagestyle:��ʽ,leibie:1-����,2-Ӣ��
'ʹ�÷���:ShowPage(Url,nowpage,totalpage,totalrecord,pagestyle,leibie)
Url=Url
nowpage=nowpage
totalpage=totalpage
totalrecord=totalrecord
pagestyle=pagestyle
leibie=leibie
if leibie="" then leibie=1
if leibie=1 then
   classname1="��ҳ"
   classname2="��һҳ"
   classname3="��һҳ"
   classname4="βҳ" 
 elseif leibie=2 then
   classname1="Home"
   classname2="Previous"
   classname3="Next"
   classname4="Mei page" 
 end if
response.Write("<div style='padding-left:0px;'>")
if nowpage>1 then
   response.Write("<a href='?page=1"&Url&"' class='"&pagestyle&"' title='"&classname1&"'>"&classname1&"</a>") 
 else
   response.Write(""&classname1&"")
end if
response.Write("&nbsp;")
if nowpage>1 then
   response.Write("<a href='?page="&nowpage-1&""&Url&"' class='"&pagestyle&"' title='"&classname2&"'>"&classname2&"</a>") 
 else
   response.Write(""&classname2&"")
end if
response.Write("&nbsp;")
if nowpage<totalpage then
   response.Write("<a href='?page="&nowpage+1&""&Url&"' class='"&pagestyle&"' title='"&classname3&"'>"&classname3&"</a>") 
 else
   response.Write(""&classname3&"")
end if
response.Write("&nbsp;")
if nowpage<totalpage then
   response.Write("<a href='?page="&totalpage&""&Url&"' class='"&pagestyle&"' title='"&classname4&"'>"&classname4&"</a>") 
 else
   response.Write(""&classname4&"")
end if
response.Write("&nbsp;")
if leibie=1 then
   response.Write("ҳ�Σ�"&nowpage&"/"&totalpage&"ҳ ��"&totalrecord&"����¼")
 elseif leibie=2 then
   response.Write("Page��"&nowpage&"/"&totalpage&" All "&totalrecord&" Record")
 end if  
response.Write("</div>")
end Function

Function ShowClassName(ClassID,TableName,lianjie_zifu,lianjie,weblink,ziti)
'��������:ShowClassName
'��������:���ط�������
'�������:ClassTitle�������� �磺Sbe_Product  ;  ClassID :��ѡ��ID
'ʹ�÷���: Tname=ShowClassName("sbe_product",tid)  '������ѡ����Classid=0
If ClassID="" or ClassID="" Then
	    ShowClassName=""
	 ElSE
	      Set Rs_ShowClassName=Conn.execute("Select top 1 ClassName,ID From "&TableName&" Where ID="&ClassID)	  
		  If Not Rs_ShowClassName.Eof Then
		    if cint(lianjie)=1 then
			   ShowClassName=lianjie_zifu&"<a href='"&weblink&"?ClassID="&Rs_ShowClassName(1)&"' title='"&Rs_ShowClassName(0)&"' class='"&ziti&"'>"&Rs_ShowClassName(0)&"</a>"  
			 else
 		       ShowClassName=lianjie_zifu&Rs_ShowClassName(0)
			 end if  
		  Else
		       ShowClassName=""
		  End If
		  Set Rs_ShowClassName=Nothing
	 End If
End Function
Function FatherName(ClassID,TableName,lianjie_zifu,lianjie,weblink,ziti)
'��������:ShowName
'��������:���ظķ��������Լ�������   ----û������
'�������  Cid :��ѡ��ID
Set rsf=Server.CreateObject("adodb.recordset")
  sql="Select ParPath,ClassName,ID from "&TableName&" Where ID="&ClassID
  rsf.Open Sql,Conn,1,1
  if rsf.eof then Exit Function
     If rsf(0)=0 Then
	     'FatherID=ClassID
		 'response.Write(FatherID)
		 if cint(lianjie)=1 then
		    FatherNameS="<a href='"&weblink&"?ClassID="&rsf(2)&"' title='"&rsf(1)&"' class='"&ziti&"'>"&rsf(1)&"</a>" 
		   else 
		    FatherNameS=rsf(1)
		  end if
		 response.Write FatherNameS 
	 Else
	     FatherIDs=ClassID
	     Set rs_2=Server.CreateObject("Adodb.recordset")
		 'SQL="Select ID From "&TableName&" Where ID in ("&rsf(0)&","&ClassID&")"
		 SQL="Select ID From "&TableName&" Where ID in ("&rsf(0)&")"
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
   rsf.Close
   Set rsf=Nothing
 End Function
Function ChildrenID(ClassTitle,ClassID)
'��������:ChildrenID
'��������:���ظķ����������ӷ��༰����ID
'�������:ClassTitle�������� �磺Sbe_Product  ;  ClassID :��ѡ��ID
Set Rs_ChildrenID=Server.CreateObject("adodb.recordset")
  sql="Select ChildNum,ParPath from "&ClassTitle&"_Class Where lock=0 and ID="&ClassID  
  Rs_ChildrenID.Open Sql,Conn,1,1
     If Rs_ChildrenID(0)=0 Then
	     ChildrenID=ClassID
	 Else
	     ChildrenIDs=ClassID
	     Set Rs_ChildrenIDS=Server.CreateObject("Adodb.recordset")
		 SQL="Select ID From "&ClassTitle&"_Class Where lock=0 and ParPath like '"&Rs_ChildrenID(1)&","&ClassID&"%'"
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

'//��̬��ȡͼƬ�ߴ����
Class imgInfo 
dim aso 
Private Sub Class_Initialize 
set aso=CreateObject("Adodb.Stream") 
aso.Mode=3 
aso.Type=1 
aso.Open 
End Sub 
Private Sub Class_Terminate
err.clear
set aso=nothing 
End Sub 

Private Function Bin2Str(Bin) 
Dim I, Str 
For I=1 to LenB(Bin) 
clow=MidB(Bin,I,1) 
if ASCB(clow)<128 then 
Str = Str & Chr(ASCB(clow)) 
else 
I=I+1 
if I <= LenB(Bin) then Str = Str & Chr(ASCW(MidB(Bin,I,1)&clow)) 
end if 
Next 
Bin2Str = Str 
End Function 

Private Function Num2Str(num,base,lens) 
dim ret 
ret = "" 
while(num>=base) 
ret = (num mod base) & ret 
num = (num - num mod base)/base 
wend 
Num2Str = right(string(lens,"0") & num & ret,lens) 
End Function 

Private Function Str2Num(str,base) 
dim ret 
ret = 0 
for i=1 to len(str) 
ret = ret *base + cint(mid(str,i,1)) 
next 
Str2Num=ret 
End Function 

Private Function BinVal(bin) 
dim ret 
ret = 0 
for i = lenb(bin) to 1 step -1 
ret = ret *256 + ascb(midb(bin,i,1)) 
next 
BinVal=ret 
End Function 

Private Function BinVal2(bin) 
dim ret 
ret = 0 
for i = 1 to lenb(bin) 
ret = ret *256 + ascb(midb(bin,i,1)) 
next 
BinVal2=ret 
End Function 

Private Function getImageSize(filespec) 
dim ret(3) 
aso.LoadFromFile(filespec) 
bFlag=aso.read(3) 
select case hex(binVal(bFlag)) 
case "4E5089": 
aso.read(15) 
ret(0)="PNG" 
ret(1)=BinVal2(aso.read(2)) 
aso.read(2) 
ret(2)=BinVal2(aso.read(2)) 
case "464947": 
aso.read(3) 
ret(0)="GIF" 
ret(1)=BinVal(aso.read(2)) 
ret(2)=BinVal(aso.read(2)) 
case "535746": 
aso.read(5) 
binData=aso.Read(1) 
sConv=Num2Str(ascb(binData),2 ,8) 
nBits=Str2Num(left(sConv,5),2) 
sConv=mid(sConv,6) 
while(len(sConv)<nBits*4) 
binData=aso.Read(1) 
sConv=sConv&Num2Str(ascb(binData),2 ,8) 
wend 
ret(0)="SWF" 
ret(1)=int(abs(Str2Num(mid(sConv,1*nBits+1,nBits),2)-Str2Num(mid(sConv,0*nBits+1,nBits),2))/20) 
ret(2)=int(abs(Str2Num(mid(sConv,3*nBits+1,nBits),2)-Str2Num(mid(sConv,2*nBits+1,nBits),2))/20) 
case "FFD8FF": 
do 
do: p1=binVal(aso.Read(1)): loop while p1=255 and not aso.EOS 
if p1>191 and p1<196 then exit do else aso.read(binval2(aso.Read(2))-2) 
do:p1=binVal(aso.Read(1)):loop while p1<255 and not aso.EOS 
loop while true 
aso.Read(3) 
ret(0)="JPG" 
ret(2)=binval2(aso.Read(2)) 
ret(1)=binval2(aso.Read(2)) 
case else: 
if left(Bin2Str(bFlag),2)="BM" then 
aso.Read(15) 
ret(0)="BMP" 
ret(1)=binval(aso.Read(4)) 
ret(2)=binval(aso.Read(4)) 
else 
ret(0)="" 
end if 
end select 
ret(3)="width=""" & ret(1) &""" height=""" & ret(2) &"""" 
getimagesize=ret 
End Function 

Public Function imgW(pic_path) 
Set fso1 = server.CreateObject("Scripting.FileSystemObject") 
If (fso1.FileExists(pic_path)) Then 
Set f1 = fso1.GetFile(pic_path) 
ext=fso1.GetExtensionName(pic_path) 
select case ext 
case "gif","bmp","jpg","png": 
arr=getImageSize(f1.path) 
imgW = arr(1) 
end select 
Set f1=nothing 
else
imgW = 0
End if 
Set fso1=nothing 
End Function 

Public Function imgH(pic_path) 
Set fso1 = server.CreateObject("Scripting.FileSystemObject") 
If (fso1.FileExists(pic_path)) Then 
Set f1 = fso1.GetFile(pic_path) 
ext=fso1.GetExtensionName(pic_path) 
select case ext 
case "gif","bmp","jpg","png": 
arr=getImageSize(f1.path) 
imgH = arr(2) 
end select 
Set f1=nothing 
else
imgH = 0 
End if 
Set fso1=nothing 
End Function 
'          //��ȡ����,д��
'imgpath="img.jpg"
'set pp=new imgInfo 
'w = pp.imgW(server.mappath(imgpath)) 
'h = pp.imgH(server.mappath(imgpath)) 
'set pp=nothing
End Class
'//���

Function Show_pic(address,width,height,weblink,pic_title,leibie,pic_target,table_width,table_height,beijing,bianhuang,nopic,class_pic,xianshi)
'�÷�: address -ͼƬ��ַ   width-��  height-��  weblink-���ӵ�ַ   pic_title-title����  leibie-�Ƿ������� 1-��  2-û��
pic_address=trim(address)
pic_width=cint(trim(width))
pic_height=cint(trim(height))
weblink=trim(weblink)
pic_title=trim(pic_title)
leibie=cint(trim(leibie))
pic_target=trim(pic_target)
table_width=cint(trim(table_width))
table_height=cint(trim(table_height))
beijing=trim(beijing)
bianhuang=trim(bianhuang)
class_pic=trim(class_pic)
xianshi=cint(xianshi)
if instr(trim(pic_address),".jpg")=0 and instr(trim(pic_address),".gif")=0 and instr(trim(pic_address),".jpeg")=0 and instr(trim(pic_address),"bmp")=0 then
response.Write("<table width="""&cint(table_width)&""" height="""&cint(table_height)&"""  cellspacing=""0"" cellpadding=""0"" border=""0""")
if bianhuang=1 then response.Write(" style=""border:1px solid #cccccc""")
response.Write(">")
response.write("<tr><td")
if beijing=1 then response.Write(" bgcolor=#ffffff")
response.Write(">")
if leibie=1 then 
response.Write("<a href="""&weblink&""" target="""&pic_target&""" ")
if xianshi=1 then response.Write("onClick=""javascript:Check_url('"&weblink&"');return false;"" ") 
response.Write(" >")
end if
response.Write ("<img src="""&nopic&""" width="""&cint(pic_width)&""" height="""&cint(pic_height)&""" border=""0"" title="""&pic_title&""" class="""&class_pic&""">")
if leibie=1 then response.Write("</a>")
response.Write("</td></tr></table>")
   else
   imgpath=pic_address
set pp=new imgInfo 
w = pp.imgW(server.mappath(imgpath))
h = pp.imgH(server.mappath(imgpath)) 
set pp=nothing
if w=0 or h=0 then
   w2=pic_width
   h2=pic_height
  else
if pic_width>0 and w>pic_width then
   w1=pic_width
  else
   w1=w
 end if
   h1=(h*w1)/w
if pic_height>0 and h1>pic_height then
   h2=pic_height
  else
   h2=h1
 end if
   w2=(w1*h2)/h1
  end if
response.Write("<table width="""&cint(table_width)&""" height="""&cint(table_height)&""" cellspacing=""0"" cellpadding=""0""")
if bianhuang=1 then response.Write(" style=""border:1px solid #cccccc""")
response.Write(">")
response.Write("<tr><td align=""center"" valign=""middle"">")
response.Write("<table width="""&cint(w2)&""" height="""&cint(h2)&"""  cellspacing=""0"" cellpadding=""0"" border=""0""><tr><td")
if beijing=1 then response.Write(" bgcolor=#ffffff")
response.Write(">")
if leibie=1 then 
response.Write("<a href="""&weblink&""" target="""&pic_target&""" ") 
if xianshi=1 then response.Write("onClick=""javascript:Check_url('"&weblink&"');return false;"" ")
response.Write(">")
end if
   response.Write ("<img src="""&imgpath&""" width="""&cint(w2)&""" height="""&cint(h2)&""" border=""0"" title="""&pic_title&"""  class="""&class_pic&""">")
if leibie=1 then response.Write("</a>")
response.Write("</td></tr></table>")
response.Write("</td></tr></table>")
end if
end function
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
			gotTopic=left(str,i)&"..."
			exit for
		Else
			gotTopic=str&""
		End if
	Next
End function

Function gotTopic1(str,strlen)
'��������: gotTopic1
'��������: �����ַ�����ʾ�ĳ���
'ʹ�÷�����gotTopic1(str,strlen)
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
			gotTopic1=left(str,i)&""
			exit for
		Else
			gotTopic1=str&""
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


  function picc(imgpath)
imgpath=imgpath 
set pp=new imgInfo 
w = pp.imgW(server.mappath(imgpath)) 
h = pp.imgH(server.mappath(imgpath)) 
set pp=nothing
if w>291 then 
   w1=291
   h1=(h*w1)/w
   if h1>560 then
     h2=560
     w2=(h2*w1)/h1
	 else
     h2=h1
     w2=w1
    end if
 end if
response.write("width="""&cint(w2)&""" height="""&cint(h2)&"""")
End Function 

Call OpenData()
 Sql="Select * From WebConfig"
 Set rs=conn.execute(Sql)
 if not rs.eof then
    NetUrl=replace(Rs("Web"),"http://","")     '��ַ
	NetName = Rs("WebName")                '��վ����
	Company = Rs("Company")               '��˾����
	WebName2 = Rs("WebName2")             '��վ�ؼ���
	WebName3 = Rs("WebName3")             '��վ����
    jishu_web=Rs("msn")                   '����֧��-��ַ
    jishu_name=Rs("WatermarkWord")        '����֧��-��վ����
	flag_web = Rs("flag_web")              '��վ״̬
	web_miaoshu = Rs("web_miaoshu")        '״̬����
	tel1 = Rs("tel1")                      '���쵥λ
	tel2 = Rs("tel2")                      '�绰
	tel3 = Rs("tel3")                      '����
	'email = Split(Rs("email"),",")
	youbian = Rs("email")                  'email
	address_company=Rs("qq")       '��˾��ַ
  end if
 rs.close
 set rs=nothing
   if session("over")="" then
    jsqtoday=1
    if application("dntime")<=cint(hour(time())) then
	conn.execute("update WebConfig set jsqtoday=jsqtoday+1")
	tmprs=conn.execute("Select jsqtoday from WebConfig")
	jsqtoday=tmprs(0)
    else
	conn.execute("update WebConfig set jsqtoday=1")
	tmprs=conn.execute("Select jsqtoday from WebConfig")
	jsqtoday=tmprs(0)
    end if
    application("dntime")=cint(hour(time()))
    set tmprs=nothing
    
	conn.execute("update WebConfig set jsq=jsq+1")
	tmprs=conn.execute("Select jsq from WebConfig")
	jsq=tmprs(0)
    set tmprs=nothing
    
    session("over")=true
else
    jsqtoday=1
	tmprs=conn.execute("Select jsqtoday from WebConfig")
	jsqtoday=tmprs(0)
	
	tmprs=conn.execute("Select jsq from WebConfig")
	jsq=tmprs(0)
    set tmprs=nothing
end if



Sub PageControl(iCount,pagecount,page)
	'������һҳ��һҳ����
    Dim query, a, x, temp
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")

    query = Split(Request.ServerVariables("QUERY_STRING"), "&")
    For Each x In query
        a = Split(x, "=")
        If StrComp(a(0), "page", vbTextCompare) <> 0 Then
            temp = temp & a(0) & "=" & a(1) & "&"
        End If
    Next

    Response.Write(" <table width=100% border=0>" & vbCrLf )        
    Response.Write("<form method=get style=margin:0px onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;""><TR height=15>" & vbCrLf )
    Response.Write("<TD align=left class=bai2>")
        
    if page<=1 then
        Response.Write ("��ҳ " & vbCrLf)        
        Response.Write ("��ҳ " & vbCrLf)
    else        
        Response.Write("<A HREF=" & action & "?" & temp & "Page=1>��ҳ</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page-1) & ">��ҳ</A> " & vbCrLf)
    end if

    if page>=pagecount then
        Response.Write ("��ҳ " & vbCrLf)
        Response.Write ("βҳ " & vbCrLf)            
    else
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page+1) & ">��ҳ</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & pagecount & ">βҳ</A> " & vbCrLf)            
    end if

    Response.Write(" ҳ�Σ�" & page & "/" & pageCount & "ҳ" &  vbCrLf)
    Response.Write(" ����" & iCount & "����¼" &  vbCrLf)
    'Response.Write(" ת��" & "<INPUT TYEP=TEXT NAME=page SIZE=1 Maxlength=5 VALUE=" & page & ">" & "ҳ"  & vbCrLf & "<INPUT type=submit value=GO>")
    Response.Write("</TD>" & vbCrLf )                
    Response.Write("</TR></form>" & vbCrLf )        
    Response.Write("</table>" & vbCrLf )        
End Sub
Call CloseDataBase()
news_id1     = 1    '�ۺ�����
news_id2     = 4    '��ҵ����
news_id3     = 5    '������
pro_id1      = 9    '������� 
pro_id2      = 10   '���ܼ��
pro_id3      = 11   '�������߼��
pro_id4      = 12   '���ع����ܼ��
pro_id5      = 13   '�������
%>  
<script language="javascript">
  <!--
  function Check_url(url)
  {
     var strURL=url;
window.open (strURL,"_blank","status=no,resizable=0,toolbar=no,menubar=no,scrollbars=yes,width=600,height=550,left=300,top=30,help:no,scroll:no");
  }
//function Check_url(url){
//  var strURL=url;
//window.showModalDialog(""+strURL,"","dialogwidth=350px;dialogheight=160px;status=no;help:no;scroll:no");
//}

  -->
  </script>
     <Script Language="JavaScript">
<!--
 	function Check_user(ID){
  var strURL="check_user.asp?ID="+ID;
window.open (strURL,"_blank","status=no,toolbar=no,menubar=no,scrollbars=no,width=300,height=50,left=300,top=200,scroll:no");
// window.showModalDialog(strURL,"","dialogwidth=350px;dialogheight=160px;status=no;help:no;scroll:no");
  }
   // -->
</Script>