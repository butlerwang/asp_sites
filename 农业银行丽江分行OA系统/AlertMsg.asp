<!--#INCLUDE FILE="data.asp" -->
    <%
    set rs=server.createobject("ADODB.recordset")
    rs.open "select * from chat where id="&request("id"),conn,3,3
    rs("zt")=true
    rs.update 
    %>

<script>
function OpenWindows(url,widthx,heighx)
{
  var 
 newwin=window.open(url,"_blank","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=20,left=60,width=600,height=500");
 return false;
 
}
</script>

<html>
<head>
<title>收到消息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
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

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}
//-->
</script>
</head>
    <tr>    

<body bgcolor="#276DB2" leftmargin="0" topmargin="5" marginwidth="0" onLoad="MM_preloadImages('images/history_on.gif','images/cancel_on.gif','images/reset_on.gif','images/submit_on.gif','images/more_on.gif','images/close_on.gif','images/reply_on.gif','images/cancel_m_on.gif')">
<table border="0" cellspacing="1" cellpadding="2" width="310">
  <tr> 
    <td colspan="2"><font color="#FFFFFF"><b>姓名 
      <input type="text" name="receiveuser" style="height:12pt; background-color:#A9C5E0" size="11" value="<%=rs("from")%>" size=15>
      时间：<%=rs("time")%> </b></font> </td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <textarea name="message" style="background-color:#A9C5E0" rows="6" cols="38"><%=replace(rs("message"),"("&rs("time")&") "&rs("from")&" 对 "&session("Rname")&chr(13),"")%></textarea>
    </td>
  </tr>
  <tr> 
    <td width="305" align="right"> <a href="show.asp?receiveuser=<%=rs("from")%>&id=<%=rs("send")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','images/reply_on.gif',1)"><img name="Image1" border="0" src="images/reply_off.gif" width="56" height="22" hspace="5"></a><a href="javascript:window.close();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image2','','images/cancel_m_on.gif',1)"><img name="Image2" border="0" src="images/cancel_m_off.gif" width="56" height="22" hspace="5"></a>
    </td>
    <td width="5" align="right">&nbsp;</td>
  </tr>
</table>

</body>
</html>
<%rs.close
  set rs=nothing
  conn.close
  set conn=nothing%>
