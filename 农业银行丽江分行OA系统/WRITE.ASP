<% if Session("id")="" then Response.Redirect("../error3.html") %>
<html>
<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<% 
	Set Record= Server.CreateObject("ADODB.Recordset") 
	strSql="select * from user "
	Record.open strSql,Conn,1,3 
	NumRec=Record.RecordCount 

%>

<%
	Session("sub")="写邮件"
if Request.QueryString("op")="back" then
	Session("iBackid")=Session("Backid")
	Session("iBackcent")=Session("Backcent")
	Session("iBackinfo")=Session("Backinfo")
	Session("sub")="回复邮件"
end if
%>
<head>
<script language="vbscript" >
function MM_changeProp() 
if  document.body.clientWidth>400 and  document.body.clientHeight>300 then

document.all.layer1.style.left=document.body.clientWidth-350+document.body.scrollLeft
document.all.layer1.style.top=document.body.clientHeight-300+document.body.scrollTop
document.all.Form.info.rows=int((document.body.clientHeight-200)/20)+5
document.all.Form.info.cols=int((document.body.clientWidth)/10)
document.all.Form.cent.size=document.all.Form.info.cols+1
document.all.Form.sendalso.size=document.all.Form.info.cols-27
document.all.Form.file.size=document.all.Form.info.cols-9
end if
end function

</script>

<script language="JavaScript">
<!--
function MM_popupMsg(msg) { //v1.0
  alert(msg);
}
//-->
</script>
<link name="stsheet" id="stsheet" rel="stylesheet" href="images/vblife.css">
</head>

<script language="vbscript">
if screen.width>1000 then document.stylesheets("stsheet").href="images/vblife1024.css"
</script>

<script language="jscript" runat=server>
var today=new Date()
var nowdate=today.getFullYear()+"-"
if (today.getMonth()+1<10) nowdate+="0"
nowdate=nowdate+(today.getMonth()+1)+"-"
if (today.getDate()<10) nowdate+="0"
nowdate=nowdate+(today.getDate())+" "
if (today.getHours()<10) nowdate+="0"
nowdate=nowdate+(today.getHours())+":"
if (today.getMinutes()<10) nowdate+="0"
nowdate=nowdate+(today.getMinutes())
</script>

<body leftmargin="0" topmargin="0" onload=MM_changeProp() onResize="MM_changeProp()">
<div id="Layer88" style="position:absolute; width:100%; height:33px; z-index:0; left: 0; top: 0"><img src="IMAGES/line.jpg" width="100%" height="100%"></div>
<div id="Layer89" style="position:absolute; width:100%; height:33px; z-index:2; left: 0; top: 0" class="bigtitle"><img src="images/iwrite.gif" width="32" height="32" align="absmiddle"><%=Session("sub")%></div>



<div id="layer2" style="position:absolute; width:100%; height:115px; z-index:1; left: 1px; top: 50px"> 

   <form method="post" name="Form" action="send.asp"   enctype="multipart/form-data" >
    <table width="100%" border="0" cellspacing="1" cellpadding="0">
	  <tr> 
	    <td width="104" height="25" class="css1" align="right" bgcolor="#0066CC" nowrap><b><font color="#FFFFFF"> 
		  收件人：</font></b></td>
	    <td height="20" class="css1" nowrap width="897"> 
		  <select name="id" size="1" onchange="vbscript:document.all.sendalso.value=''">
		  <option value="">--------------请选择收信人---------------</option>
		  <% for i=1 to NumRec  %>
		  <option value="<%=Record("用户名")%>"<% if Session("iBackid")=Record("用户名") then Response.Write(" selected")%>><%=Record("用户名")%> 
		  </option>
		  <% Record.MoveNext
             next
           Record.MoveFirst%>
		</select>
	  </td>
	</tr>
	<tr> 
	    <td width="104" height="25" class="css1" align="right" bgcolor="#0066CC" nowrap><b><font color="#FFFFFF">抄报/抄送：</font></b></td>
	    <td height="20" class="css1" nowrap width="897"> 
		  <select name="D1" size="1" class="css0" >
		  <option value="" selected>抄报/抄送人</option>
		  <%
		  while not Record.eof
		  if Record("ilevel")>="2" then
		  	response.write("<option value="+Record("用户名")+">"+Record("用户名")+"</option>")
		  end if
		  Record.MoveNext
		  wend
		  %>
		</select>
		<input type="button" name="Submit" value="添加" class="css0" onClick="vbscript:if document.all.D1.value<>'' and document.all.D1.value<>document.all.id.value and instr(document.all.sendalso.value,document.all.D1.value)=0 then document.all.sendalso.value=document.all.sendalso.value+document.all.D1.value+','">
		<input type="button" name="Submit2" value="清空" class="css0" onClick="vbscript:document.all.sendalso.value=''">
		  <input type="text" name="sendalso" class=css0 size="42" onFocus="javascript:MM_popupMsg('请不要手动编辑此框内容,否则将造成不可估计的错误!');document.all.Form.cent.focus()">
	  </td>
	</tr>
	<tr> 
	    <td width="104" height="25" class="css1" align="right" bgcolor="#0066CC" nowrap><b><font color="#FFFFFF">主题：</font></b></td>
	    <td height="20" class="css1" nowrap width="897"> 
		  <input type="text" name="cent" value="<%=Session("iBackcent")%>" maxlength="100"  class="css0" size="70">
	  </td>
	</tr>
	<tr> 
	    <td width="104" height="25"  valign="top" class="css1" bgcolor="#0066CC" nowrap> 
		  <p align="right"><b><font color="#FFFFFF">内容：</font></b></p>
		</td>
	    <td height="100" class="css1" nowrap rowspan="2" width="897" > 
		  <textarea name="info" rows="24" cols="100"  class="css0"><%=Session("iBackinfo")%></textarea>
		<input type="hidden" name="DateTime" value="<%=nowdate%>">
		<input type="hidden" name="from" value="<%=Session("id")%>">
		<input type="hidden" name="level" value="<%=Session("level")%>">
		<input type="hidden" name="op">
	  </td>
	</tr>
	<tr>
	    <td width="104" height="100"  valign="top" class="css1" nowrap>&nbsp;</td>
	</tr>
	<tr> 
	    <td width="104" height="25" class="css1" align="right" bgcolor="#0066CC" nowrap><font color="#FFFFFF"><b>附件:</b></font></td>
	    <td height="20" class="css1" colspan="2" nowrap width="897"> 
		  <input type="file" name="file" class="css0" size="60">
	  </td>
	</tr>
	<tr> 
	    <td height="25"  width="104" nowrap>&nbsp; </td>
	    <td height="20" class="css1" nowrap width="897"> 
		
		<input type="Submit" name="Submit4" class="css0"  value="立即发送"  IsShowProcessBar="True"  onClick=op.value=0>
		<input type="reset" name="Submit22" class="css0"  value="重写"  >
		<input type="submit" name="Submit32" class="css0"  value="作为公告发布"  IsShowProcessBar="True"  onClick=op.value=1>
		<input type="checkbox"  name="check1" value="checkbox" checked>保存副本
		<input type="hidden" name="temp">
		<input type="hidden" name="sfsdfsdf" value="sdfsdfsfdsfsdf">
		</td>
	</tr>
  </table>                                            
</form>   
</div>
<div id="layer1" style="position:absolute; width:200px; height:115px; z-index:0; left: 1000; top: 1000; visibility: visible"><img src="IMAGES/bg.jpg" width="375" height="300" name="bgimg"></div>
                                     
</body>                                            
<%                                            
	Session("iBackid")=""                                            
	Session("iBackcent")=""                                            
	Session("iBackinfo")=""                                            
	Session("Backid")=""                                            
	Session("Backcent")=""                                            
	Session("Backinfo")=""                                            
                                            
%>                                             
</html>