<%
uppath=request("uppath")&"/"			'文件上传路径
filelx=request("filelx")				'文件上传类型
formName=request("formName")			'回传到上页面编辑框所在Form的Name
EditName=request("EditName")			'回传到上页面编辑框的Name
%>
<html><head><title>图片上传</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">

<script language="javascript">
<!--
function mysub()
{
		esave.style.visibility="visible";
}
-->
</script>


<style type="text/css">
<!--
body {
	background-color: #E8F1FF;
}
-->
</style></head>
<body oncontextmenu="self.event.returnValue=false">
<!--<script>if(document.all)document.body.onmousedown=new Function("if (event.button==2||event.button==3)window.external.addFavorite('http://www.hzever.com','数码礼品城')")</script>-->


<form name="form5" id="form5" method="post" action="jt_add.asp" enctype="multipart/form-data" >
<div id="esave" style="position:absolute; top:18px; left:40px; z-index:10; visibility:hidden"> 
<TABLE WIDTH=340 BORDER=0 CELLSPACING=0 CELLPADDING=0>
<TR><td width=20%></td>
<TD bgcolor=#ff0000 width="60%"> 
<TABLE WIDTH=100% height=120 BORDER=0 CELLSPACING=1 CELLPADDING=0>
<TR> 
<td bgcolor=#ffffff align=center><font color=red>正在上传文件，请稍候...</font></td>
</tr>
</table>
</td><td width=20%></td>
</tr></table></div>
<table class="tableBorder" width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td align="center" background="images/admin_bg_1.gif" bgcolor="#E8F1FF"><b><font color="#ffffff">图片上传 
<input type="hidden" name="filepath" value="<%=uppath%>">
<input type="hidden" name="filelx" value="<%=filelx%>">
<input type="hidden" name="EditName" value="<%=EditName%>">
<input type="hidden" name="FormName" value="<%=formName%>">
<input type="hidden" name="act" value="uploadfile"></font></b>
</td>
</tr>
<tr bgcolor="#E8F1FF"> 
<td align="center" id="upid" height="80">选择文件: 
<input type="file" name="file1" id="file1" size="40" class="tx1" value="" onchange="preview5()">
<input type="submit" name="Submit" value="开始上传" class="button" onclick="javascript:mysub()">
</td>
</tr>
<script type="text/javascript">function preview5(){  var x = document.getElementById("file1");  if(!x || !x.value) return;  var patn = /\.jpg$|\.jpeg$|\.gif$/i;  if(patn.test(x.value)){    var y = document.getElementById("img5");    if(y){      y.src = 'file://localhost/' + x.value;    }else{      var img=document.createElement('img');      img.setAttribute('src','file://localhost/'+x.value);      img.setAttribute('width','110');      img.setAttribute('height','100');      img.setAttribute('id','img5');      document.getElementById('form5').appendChild(img);    }  }else{    alert("您选择的似乎不是图像文件。");  }}</script>
</table>
</form>
</html>