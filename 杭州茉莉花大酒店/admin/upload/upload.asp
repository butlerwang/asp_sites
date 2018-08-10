<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<HTML>
<HEAD>
<STYLE TYPE="text/css">
<!--
td {font-size: 9pt}
.bt {border-left:1px solid #C0C0C0; border-top:1px solid #C0C0C0; font-size: 9pt; border-right-width: 1; border-bottom-width: 1; height: 20px; width: 80px; background-color: #EEEEEE; cursor: hand; border-right-style:solid; border-bottom-style:solid}
.tx1 { width: 200 ;height: 20px; font-size: 9pt; border: 1px solid; border-color: black black #000000; color: #0000FF}
-->
</STYLE>
<script language="JavaScript">
function check(){
 if(form1.file1.value==""){
 alert("请选择上传文件！");
 form1.file1.focus();
 return false
 }
 
 document.form1.Submit.disabled=true;
 document.form1.Submit.value="正在上传,请稍候"
 return true;
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></HEAD>
<BODY bgcolor="#F5F8FA" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
  Call OpenData()
  Set Rs=Server.CreateObject("adodb.recordset")
  Sql="Select UpfileType,UpfileSize,UpmovieType,UpmovieSize from sbe_WebConfig"
  Rs.Open Sql,Conn,1,1
  UpfileType=Rs("UpfileType")
  UpfileSize=Rs("UpfileSize")
  UpmovieType=Rs("UpmovieType")
  UpmovieSize=Rs("UpmovieSize")
  Rs.close:set rs=nothing
  
  if Request.QueryString("UploadFile")="url" then
     link="upmovie.asp"
  else
     link="upfile.asp"
  end if
%>
<table width="500" height="20" border="0" cellpadding="0" cellspacing="0">
  <form name="form1" method="post" action="<%=link%>" onSubmit="return check()" enctype="multipart/form-data">
    <tr> 
      <td width="206" height="20"> 
        <input type='file' name='file1' size='20'  class='tx1'>
        <input name="Form_Name" type="hidden" id="Form_Name" value="<%=Request.QueryString("Form_Name")%>">
        <input name="UploadFile" type="hidden" id="UploadFile" value="<%=Request.QueryString("UploadFile")%>">
	  </td>
      <td width="93"> 
        <input type="submit" name="Submit" value="☉上传☉" class="bt">
      </td>
	  <td width="179"><%if Request.QueryString("UploadFile")="url" then%>允许类型:<%=UpmovieType%> &nbsp;<%=UpmovieSize%>M以内<%else%>允许类型:<%=UpfileType%>&nbsp; <%=UpfileSize%>K<%end if%></td>
    </tr>
  </form>
</table>
<%Call CloseDataBase()%>
</BODY>
</HTML>