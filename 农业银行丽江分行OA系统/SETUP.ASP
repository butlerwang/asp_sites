<% if Session("id")="" then Response.Redirect("../error3.html") %>
<html>
<head>
<title>个人设置</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link id="stsheet" name="stsheet" rel="stylesheet" href="images/vblife.css">
<script language="JavaScript">
<!--
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

</script>
<script language="VBScript">
function MM_changeProp() 
	if  document.body.clientWidth>400 and  document.body.clientHeight>300 then
		document.all.layer1.style.left=document.body.clientWidth-350
		document.all.layer1.style.top=document.body.clientHeight -300
		document.all.Form1.textarea.rows=int((document.body.clientHeight)/25)-10
		document.all.Form1.textarea.cols=int((document.body.clientWidth)/10)
	end if
end function
</script>

</head>

<script language="vbscript">
if screen.width>1000 then document.stylesheets("stsheet").href="images/vblife1024.css"
</script>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" class=css1 onload=MM_changeProp() onResize="MM_changeProp()">
<div id="Layer88" style="position:absolute; width:100%; height:35px; z-index:0; left: 0; top: 0"><img src="IMAGES/line.jpg" width="100%" height="100%"></div>
<div id="Layer89" style="position:absolute; width:100%; height:35px; z-index:1; left: 0; top: 0" class="bigtitle"><img src="images/setup.gif" width="32" height="32" align="absmiddle">个性设置</div>


<form name="form1" action="setup.asp" method="post" >
  <p>&nbsp;</p>
  <div id="layer1" style="position:absolute; width:200px; height:115px; z-index:0; left: 258px; top: 178px; visibility: visible"><img src="IMAGES/bg.jpg" width="375" height="300" name="bgimg"></div>
  <div id="layer2" style="position:absolute; width:594px; height:115px; z-index:1; left: 39px; top: 50px"> 
	<% if Request.Form("iPageSize").Count<>0 then
	 if Request.Form("pwd1")=Request.Form("pwd2") then
		'创建并打开 Connection 对象。
		set con = Server.CreateObject("ADODB.Connection") 
		ConnStr="DBQ=" & Server.Mappath("db\member.mdb") & ";DRIVER={Microsoft Access Driver (*.mdb)};"
		con.Open(ConnStr)

			sql="update member set pwd='"+Request.Form("pwd1")+"',iPageSize="+Request.Form("iPageSize")+",iAdd='"+Request.Form("iAdd")+"',iBegin='"+Request.Form("begin")+"' where id='"+Session("id")+"'" 
			response.write("<font color=red><b>您的更改已经保存！</b></font>")
			con.Execute(sql)
			Session("iPageSize")=Request.Form("iPageSize")
			Session("iAdd")=Request.Form("iAdd")
			Session("pwd")=Request.Form("pwd1")
			Session("iBegin")=Request.Form("begin")	
		con.close
		set con=nothing
      else
		response.write("<font color=red><b>您输入的密码和确认密码不相符！</b></font>")
	  end if
	end if
%>
	<p class="title"> <font color="#0000FF">每页显示邮件数:</font><br>
	  <% 	for i=1 to 4
			Response.Write("<input type=radio name=iPageSize value=" & i*5)
			if cint(Session("iPageSize"))=i*5 then Response.Write(" checked ")
			Response.Write( ">" & i*5 &"条&nbsp;&nbsp;&nbsp;&nbsp;") 
		next
	%>
	</p>
	<p class="title"> <font color="#0000FF">进入时转向信箱：</font><br>
	  <input type="radio" value="common" <%if Session("iBegin")="common" then response.write(" checked ")%> name="begin">
	  公共信息&nbsp;&nbsp; 
	  <input type="radio" value="recived" <%if Session("iBegin")="recived" then response.write(" checked ")%>  name="begin">
	  收件箱&nbsp;&nbsp; 
	  <input type="radio" value="sendout" <%if Session("iBegin")="sendout" then response.write(" checked ")%>  name="begin">
	  发件箱&nbsp;&nbsp; 
	  <input type="radio" value="del" <%if Session("iBegin")="del" then response.write(" checked ")%>  name="begin">
	  回收站 </p>
	<p class="title"><font color="#0000FF">签名档：</font><br>
	  <textarea name="textarea" cols="50" rows="7" class="css0"><%=Session("iAdd")%></textarea>
	</p>
	<font color="#0000FF" class="title">密码修改：</font><br>
	<table width="500" border="0" cellspacing="1" cellpadding="0" class=css1>
	  <tr> 
		<td width="99" bgcolor="#0066ff"> 
		  <div align="right"><b><font color="#FFFFFF">密码：</font></b></div>
		</td>
		<td width="401"> 
		  <input type="password" name="pwd1" size="20" maxlength="20" value="<%=Session("pwd")%>" class="css0">
		</td>
	  </tr>
	  <tr> 
		<td width="99" bgcolor="#0066ff"> 
		  <div align="right"><b><font color="#FFFFFF">确认密码：</font></b></div>
		</td>
		<td width="401"> 
		  <input type="password" name="pwd2" size="20" maxlength="20" value="<%=Session("pwd")%>" class="css0">
		</td>
	  </tr>
	</table>
	<p> 
	  <input type="submit" name="Submit2" value="保存修改" class=css0 >
	</p>
  </div>
</form>               
</body>               
</html>               
