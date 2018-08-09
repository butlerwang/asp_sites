<!-- #include file="../inc/access.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../Style.css" rel=stylesheet type=text/css>
<style type="text/css">
<!--
.STYLE1 {font-size: 12px}
-->
</style>
</head>
<body leftmargin="0" topmargin="0">
<%
filename=request.querystring("filename")
if filename<>"" then
response.write "<span class='forumRow STYLE1'>上传成功！</span>"
response.write "<script>parent.form1.web_image.value='"&filename&"'</script>"
else
%>
<table bg>
  <form name="form" method="post" action="upfile.asp" enctype="multipart/form-data" >
  <tr>
    <td width="362"><input type="hidden" name="filepath" value="uploadImages">
    <input type="hidden" name="act" value="upload">
    <input class=c type="file" name="file1" size=10 >
	<input type="hidden" name="juhaoyongUploadFileName" value="<%=trim(request("juhaoyongUploadFileName"))%>">
	<input type="hidden" name="juhaoyongUpLoadPath" value="<%=trim(request("juhaoyongUpLoadPath"))%>">
    <input type="submit" name="Submit" value="上传" class=c>    </td>
    <td width="142"><span class="forumRow STYLE1">格式:jpg,gif,bmp,png<1M </span></td>
  </tr></form>
</table>
<%end if%>
</body>
</html>
 