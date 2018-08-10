<% if Session("id")="" then Response.Redirect("../error3.html") %>
<!--#INCLUDE FILE="data.asp" -->
<!--#include FILE="upload.asp"-->
<%

 	dim adOpenKeyset
	dim adLockOptimistic
	dim NumOfRecord
	dim sql

	adOpenKeyset = 1
	adLockOptimistic = 3

	set con2 = Server.CreateObject("ADODB.Connection") 
	ConnStr="DBQ=" & Server.Mappath("db/mails1.mdb") & ";DRIVER={Microsoft Access Driver (*.mdb)};"
	con2.Open(ConnStr)
%>

<%dim upload,file,formName,formPath,iCount
set upload=new upload_file ''建立上传对象
     '--------将日期转化成文件名--------
    function MakedownName()
        dim fname
          fname = now()
        fname = replace(fname,"-","")
         fname = replace(fname," ","") 
        fname = replace(fname,":","")
          fname = replace(fname,"PM","")
          fname = replace(fname,"AM","")
        fname = replace(fname,"上午","")
          fname = replace(fname,"下午","")
          fname = int(fname) + int((100)*Rnd + 1)
        MakedownName=fname & iext
    end function 

Dim iDateTime,iaddfile,ifrom,iinfo,ilevel,cent,iread,iadd,info,iext	
	
formPath="file/"
iCount=0
for each formName in upload.file 
 set file=upload.file(formName) 
if file.FileSize>0 then       
	iCount=iCount+1
	iext="."&mid(file.FileName,InStrRev(file.FileName, ".")+1)
else 
	iaddfile=""
end if
next


iDateTime=upload.form("DateTime")
ifrom=upload.form("from")

iinfo=replace(upload.form("info"),"<s","&lt;s")
iinfo=replace(iinfo,"<S","&lt;S")+vbcrlf+vbcrlf+vbcrlf+vbcrlf+vbcrlf+vbcrlf

iadd=replace(Session("iAdd"),"<S","&lt;S")
iadd=replace(iadd,"<s","&lt;s")
if iinfo="" then iinfo="无内容"
iinfo=replace(iinfo,"'","’")
info=iinfo+"<p></p>"+iadd

ilevel=upload.form("level")
cent=replace(upload.form("cent"),"<","&lt;")
if cent="" then cent="无标题"
iread="f"

	
'--------邮件抄报抄送函数--------
function SendAlso(MailBox,FT,NameList)
	if NameList="" then exit function
	dim x1,x2,TempName
	x1=0
	x2=0
	if FT="ito" then ifrom=upload.form("id")    '保存备份
	while not x1=len(NameList)
		x2=instr(x1+1,NameList,",")
		if x2-x1>=0 then TempName=mid(NameList,x1+1,x2-x1-1)
		if file.FileSize>0 then       
			newname=MakedownName()
			iaddfile=newname
			file.SaveAs Server.mappath(formPath&newname)   ''保存文件 
		else 
			iaddfile=""
		end if
		sql="INSERT INTO "+MailBox+TempName+"(iDateTime,iaddfile,"+FT+",iinfo,ilevel,cent,iread) VALUES ('"
		sql=sql+iDateTime+"','"+iaddfile+"','"+ifrom+"','"+info+"','"+ilevel+"','"+cent+"','"+iread+"')"
		con2.Execute(sql)
		x1=x2
	wend
	ifrom=upload.form("from")
end function 



%><head>



<link rel="stylesheet" type="text/css" href="IMAGES/vblife.css">

</head>






<p>　</p>

<table bgColor=#0066ff border=0 align="center" cellPadding=0 cellSpacing=0 height=55 class=css1 width=620>
  <tr>
    <td bgColor=#ebebeb height=1 vAlign=bottom width="100%"><img border=0 
      height=35 src="../images/system.gif" width=154></td></tr>
  <tr>
    <td align=middle height=22 width="100%">
      &nbsp;<font 
color=#ffffff><b>系统提示</b></font></td></tr>
  <tr>
    <td align=middle height=22 width="100%" bgcolor="#FFFFFF">
　</td></tr>
  <tr>
    <td align=middle height=22 width="100%" bgcolor="#FFFFFF">

<%
if upload.form("op")="0" then    '私人信件发送


 
	iid="" & upload.form("id")
	Set Record=Server.CreateObject("ADODB.Recordset") 
	sql="select * from user where 用户名='"+iid+"'"
	Record.open sql,Conn,1,3 
	if Record.eof and Record.bof then 
		Response.write("您所输入的收信人不存在！")
	else
		call SendAlso("recived","ifrom",upload.form("id")+",")
		call SendAlso("recived","ifrom",upload.form("sendalso"))  '抄送
		if upload.form("check1")="checkbox"  then	call SendAlso("sendout","ito",Session("id")+",")
		Response.write("您的邮件已经成功发送！")
	end if
	Record.close
	set Record=nothing
else
		if file.FileSize>0 then       
			newname=MakedownName()
			iaddfile=newname
			file.SaveAs Server.mappath(formPath&newname)   ''保存文件 
		else 
			iaddfile=""
		end if
			
	sql="INSERT INTO message(iDateTime,iaddfile,ifrom,iinfo,ilevel,cent,iread) VALUES ('"
	sql=sql+iDateTime+"','"+iaddfile+"','"+ifrom+"','"+info+"','"+ilevel+"','"+cent+"','"+iread+"')"
	con2.Execute(sql)
	
	call SendAlso("recived","ifrom",upload.form("sendalso"))   '抄送
		
	if  upload.form("check1")="checkbox" then
	
		if file.FileSize>0 then       
			newname=MakedownName()
			iaddfile=newname
			file.SaveAs Server.mappath(formPath&newname)   ''保存文件 
		else 
			iaddfile=""
		end if
		
		
		sql="INSERT INTO sendout"+Session("id")+"(iDateTime,iaddfile,ito,iinfo,ilevel,cent,iread) VALUES ('"
		sql=sql+iDateTime+"','"+iaddfile+"','公告栏','"+info+"','"+ilevel+"','"+cent+"','"+iread+"')"
		con2.Execute(sql)
	end if
	Response.write("您的邮件已经成功发送！")
end if

con2.close
set con2= nothing
con.close
set con=nothing
set upload=nothing  ''删除此对象

%>

</td></tr>
  <tr>
    <td align=middle height=59 width="100%" bgcolor="#FFFFFF">
         <p><input type="button" value="  返回  " name="B3" class=css0 onclick="javascript:history.back(-1)"></p>
      </td></tr></table>





