<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<%
'        Allrights Qbtrade.com
'             Maozai
FormSize = Request.TotalBytes
FormData = Request.BinaryRead(FormSize)
Image=ImageUp (FormSize,Formdata)
if formsize>50000 then
response.write "图片太大,<A HREF=archives.asp>返回</A>"
response.end
end if
set rs=server.CreateObject("adodb.recordset")
sql="SELECT * FROM userinfo WHERE userid =" &session("Uid")
rs.Open sql,conn,1,3
rs("photo") = NULL
rs("photo").appendchunk Image
rs("havephoto")=true

rs.Update
id=rs("id")
rs.Close
set rs=nothing
  conn.close
  set conn=nothing


function ImageUp(formsize,formdata)
    bncrlf=chrb(13) & chrb(10)
    divider=leftb(formdata,instrb(formdata,bncrlf)-1)
    datastart=instrb(formdata,bncrlf&bncrlf)+4
    dataend=instrb(datastart+1,formdata,divider)-datastart
    imageup=midb(formdata,datastart,dataend)
end function

%>
<style type="text/css">
<!--
table {  font-size: 9pt}
select {  font-size: 9pt}
input {  font-size: 9pt; background-color: #CCCCFF; font-weight: bold; color: #FF6633; border-style: groove}
.smallbox {  font-size: 1pt}
a:link {  font-size: 9pt; text-decoration: none; }
a:hover {  font-size: 9pt;}
body {  font-size: 9pt}
-->
</style>
<title>完成</title>
<body bgcolor="#FFFFFF" text="#000000">
<table width="95%" border="0" align="center" cellpadding="5" height="167">
  <tr align="center" valign="middle"> 
    <td height="128"> 
      <p>以下是你刚才上传的图片，如果你不满意，可以返回重新修改一次。<%=formsize%>
      
	  <p><img src="showpic.asp?id=<%=id%>" width="80" height="100"></p>
      </td>
  </tr>
</table>
