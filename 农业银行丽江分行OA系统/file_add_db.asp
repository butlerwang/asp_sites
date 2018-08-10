<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="mouse.js" -->

<%
set rs=server.createobject("ADODB.recordset") 
rs.Open "SELECT * FROM jhtdata Where ID is null",conn,1,3 
rs.addnew

application("downdir")="download/"
if Session("Ulogin")<>"yes" then
	Response.Redirect ("index.htm")
end if
ip= Request.ServerVariables("REMOTE_ADDR")
 nowtime=now()
sj=cstr(year(nowtime))+"-"+cstr(month(nowtime))+"-"+cstr(day(nowtime))+" "+cstr(hour(nowtime))+":"+right("0"+cstr(minute(nowtime)),2)+":"+right("0"+cstr(second(nowtime)),2)
%>
<!--#include FILE="upload_5xsoft.inc"-->
<%dim upload,file,formName,formPath,iCount
set upload=new upload_5xSoft ''建立上传对象
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
	  	fname = int(fname) + int((10-1+1)*Rnd + 1)
		MakedownName=fname
	end function 
formPath="download/"
iCount=0
for each formName in upload.file ''列出所有上传了的文件
 set file=upload.file(formName)  ''生成一个文件对象
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
newname=MakedownName()&"."&mid(file.FileName,InStrRev(file.FileName, ".")+1)

  file.SaveAs Server.mappath(formPath&newname)   ''保存文件
  iCount=iCount+1
 else 
  response.write "未找到文件 &nbsp;&nbsp;<A HREF=javascript:history.back(1)>返回</A>"
  response.end
 end if
 set file=nothing
next



	rs("type") ="1"
	rs("真实姓名") =upload.form("Rname")
	rs("部门") = upload.form("Upart")
	rs("标题") = upload.form("biaoti")
	rs("内容") = upload.form("Rname")&"的上传文件"
	rs("链接") =application("downdir")&newname
	rs("IP")=ip
	rs("时间")=sj
	rs.Update
	rs.close
	Set rs=nothing
	Conn.Close
	Set Conn=nothing
%>
<LINK href="oa.css" rel=stylesheet>
<BODY bgColor=#ffffff leftMargin=0 
style="BACKGROUND-ATTACHMENT: scroll; BACKGROUND-IMAGE: url(images/main_bg.gif); BACKGROUND-POSITION: left bottom; BACKGROUND-REPEAT: no-repeat" 
topMargin=0>
<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
  <TBODY>
  <TR>
    <TD bgColor=#4e5960 class=heading colSpan=2 height=3></TD></TR>
  <TR>
    <TD bgColor=#4e5960 class=heading>　<FONT 
color=#ffffff><B>上报文件</B></FONT></TD>
    <TD align=right bgColor=#4e5960 class=heading height=20></TD></TR>
  <TR>
    <TD align=middle vAlign=top width=109></TD>
    <TD align=middle>
        <TABLE bgColor=#666666 border=0 cellPadding=1 cellSpacing=1 
        width="100%">
	  <TR> 
            <TD bgColor=#efefef>文件已经成功上报，是否继续添加……<BR>
<P>上报人：<%=upload.form("Rname")%><BR>
单&nbsp;&nbsp;位：<%=upload.form("Upart")%><BR>
说&nbsp;&nbsp;明：<%=upload.form("biaoti")%><BR>
<P><P><A HREF="addfile.asp">继续添加</A>&nbsp;&nbsp;<A HREF="images/oa_menu.swf">返回主页</A></TD>
            
      </TR>
	  
      </table>
    </td></tr></table>

<%set upload=nothing  ''删除此对象
%>