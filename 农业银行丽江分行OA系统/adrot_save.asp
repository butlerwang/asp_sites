<%
if session("Urule")<>"a" then
response.redirect "error.asp?id=admin"
end if
%>
<!--#INCLUDE FILE="data.asp" -->

<%
set rs=server.createobject("ADODB.recordset") 
rs.Open "SELECT * FROM adrot Where ID is null",conn,1,3 
rs.addnew
 nowtime=now()
sj=cstr(year(nowtime))+"-"+right("0"+cstr(month(nowtime)),2)+"-"+right("0"+cstr(day(nowtime)),2)
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
formPath="adrot/"
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
next

    rs("url") =upload.form("url")
    rs("alt") =upload.form("content")
    rs("src") ="adrot/"&newname
    rs("width")=upload.form("width")
    rs("height")=upload.form("height")
    rs("type")=upload.form("type")
    rs.Update
    rs.close
    Set rs=nothing
    Conn.Close
    Set Conn=nothing
%>
<script language=javascript>
opener.location=opener.location;window.close();
</script>

<LINK href="oa.css" rel=stylesheet>
<BODY>
        <TABLE border=1 bordercolorlight='000000' bordercolordark=#ffffff cellspacing=0 cellpadding=0 align=center>
      <TR> 
            <TD>文件已经成功上传，是否继续添加……<BR>
<%
response.write file.FilePath&file.FileName&" ("&cint(file.FileSize/1024)&"K) 上传 成功!<br>"
%>
<P><P><A HREF="uploadfile.asp">继续添加</A>&nbsp;&nbsp;<A HREF="javascript:window.close()">关闭窗口</A></TD>
            
      </TR>
      
      </table>

<%
set file=nothing
set upload=nothing  ''删除此对象
%>
