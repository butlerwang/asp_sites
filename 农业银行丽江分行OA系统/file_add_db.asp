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
set upload=new upload_5xSoft ''�����ϴ�����
 	'--------������ת�����ļ���--------
	function MakedownName()
		dim fname
	  	fname = now()
		fname = replace(fname,"-","")
	 	fname = replace(fname," ","") 
		fname = replace(fname,":","")
	  	fname = replace(fname,"PM","")
	  	fname = replace(fname,"AM","")
		fname = replace(fname,"����","")
	  	fname = replace(fname,"����","")
	  	fname = int(fname) + int((10-1+1)*Rnd + 1)
		MakedownName=fname
	end function 
formPath="download/"
iCount=0
for each formName in upload.file ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
newname=MakedownName()&"."&mid(file.FileName,InStrRev(file.FileName, ".")+1)

  file.SaveAs Server.mappath(formPath&newname)   ''�����ļ�
  iCount=iCount+1
 else 
  response.write "δ�ҵ��ļ� &nbsp;&nbsp;<A HREF=javascript:history.back(1)>����</A>"
  response.end
 end if
 set file=nothing
next



	rs("type") ="1"
	rs("��ʵ����") =upload.form("Rname")
	rs("����") = upload.form("Upart")
	rs("����") = upload.form("biaoti")
	rs("����") = upload.form("Rname")&"���ϴ��ļ�"
	rs("����") =application("downdir")&newname
	rs("IP")=ip
	rs("ʱ��")=sj
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
    <TD bgColor=#4e5960 class=heading>��<FONT 
color=#ffffff><B>�ϱ��ļ�</B></FONT></TD>
    <TD align=right bgColor=#4e5960 class=heading height=20></TD></TR>
  <TR>
    <TD align=middle vAlign=top width=109></TD>
    <TD align=middle>
        <TABLE bgColor=#666666 border=0 cellPadding=1 cellSpacing=1 
        width="100%">
	  <TR> 
            <TD bgColor=#efefef>�ļ��Ѿ��ɹ��ϱ����Ƿ������ӡ���<BR>
<P>�ϱ��ˣ�<%=upload.form("Rname")%><BR>
��&nbsp;&nbsp;λ��<%=upload.form("Upart")%><BR>
˵&nbsp;&nbsp;����<%=upload.form("biaoti")%><BR>
<P><P><A HREF="addfile.asp">�������</A>&nbsp;&nbsp;<A HREF="images/oa_menu.swf">������ҳ</A></TD>
            
      </TR>
	  
      </table>
    </td></tr></table>

<%set upload=nothing  ''ɾ���˶���
%>