<!--#INCLUDE FILE="data.asp" -->
<%
Set rs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from help where id="&request("id")
rs.open strSql,Conn,1,1 
%>

<script language="JavaScript">
<!--

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>


<html><head><title><%=rs("title")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
<script language="JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>

</head>
<body bgcolor="#efefef" topmargin="0" leftmargin="0">
<div align="center">
  <center>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr bgcolor="#4e5960"> 
    <td  class="heading" height="26">
      <p align="left"><b>¡¡<font color="#FFFFFF"><%=rs("time")%></font> 
</b></p>  
      </font>  
    </td>  
  </tr>  
</table>  
  </center>  
</div>  
<div align="center">  
  <center>  
<table width="100%" border="0" cellspacing="0" cellpadding="2" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" height="40">  
  <tr> 
    <td bgcolor="#bfbfbf" class="heading" width="451" height="20"  align=center>¡¡<b><%=rs("title")%></b></td>
  </center> 
    <td bgcolor="#bfbfbf" class="heading" width="274" height="20">
      <p align="right">
   
      ¡¡</p>
  </td>
  </tr>
  <center>  
  <tr>      
	<td class="show"  colspan="2" height="16" width="727">
<%=rs("content")%>
</td> 
  </tr> 
   
</table>  
  </center> 
</div> 
  
<table border="0" width="100%">
  <tr>
    <td width="33%">
      <p align="right"></td>
   
    <td width="16%">  
<a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image11','','images/close_2.gif',1);"  onclick="MM_callJS('window.close()')"><img name="Image11" border="0" src="images/close_1.gif" hspace="5" vspace="5"></a> 
    </td>
    <td width="34%"></td>
  </tr>
</table>
</body>  
</html>  
