<%
if session("Urule")<>"a" then
response.redirect "error.asp?id=admin"
end if
%>
<!--#INCLUDE FILE="data.asp" -->
<!--#include file="html.asp"-->

<%
set rs=server.createobject("ADODB.recordset") 
rs.Open "SELECT * FROM soft Where ID is null",conn,1,3 
rs.addnew
 nowtime=now()
sj=cstr(year(nowtime))+"-"+right("0"+cstr(month(nowtime)),2)+"-"+right("0"+cstr(day(nowtime)),2)
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
formPath="soft/"
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
next

    rs("name") =htmlencode2(upload.form("softname"))
    rs("content") =htmlencode2(upload.form("content"))
    rs("url") ="soft/"&newname
    rs("size")=cint(file.FileSize/1024)&"K"
    rs("time")=sj
    rs("best")=upload.form("best")
    rs.Update
    rs.close
    Set rs=nothing
    Conn.Close
    Set Conn=nothing
%>
<script language=javascript>
opener.location=opener.location
</script>

<LINK href="oa.css" rel=stylesheet>
<BODY>
        <TABLE border=1 bordercolorlight='000000' bordercolordark=#ffffff cellspacing=0 cellpadding=0 align=center>
      <TR> 
            <TD>�ļ��Ѿ��ɹ��ϴ����Ƿ������ӡ���<BR>
<%
response.write file.FilePath&file.FileName&" ("&cint(file.FileSize/1024)&"K) �ϴ� �ɹ�!<br>"
%>
<P><P><A HREF="uploadfile.asp">�������</A>&nbsp;&nbsp;<A HREF="javascript:window.close()">�رմ���</A></TD>
            
      </TR>
      
      </table>

<%
set file=nothing
set upload=nothing  ''ɾ���˶���
%>
