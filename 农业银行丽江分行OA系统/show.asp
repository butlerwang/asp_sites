<!--#INCLUDE FILE="data.asp" -->
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
<title>Ctrl+Enter 快捷发送消息</title>
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

function MM_findObj(n, d) { //v3.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
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
//-->
</script>
<script language=javascript>
function checkform()
{   
	if (document.form1.message.value=="")
	{
		alert("不能发送空内容");
  	    return  false;
	}
	else  
    return true;
}

function presskey(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.form1.submit();}}

</script>

</head>
<%
nowtime=now()
sj=cstr(year(nowtime))+"-"+cstr(month(nowtime))+"-"+cstr(day(nowtime))+" "+cstr(hour(nowtime))+":"+right("0"+cstr(minute(nowtime)),2)
session("receiveuser")=request("receiveuser")
session("receive")=request("id")

%>

<body bgcolor="#276DB2" leftmargin="0" topmargin="5" marginwidth="0" onLoad="MM_preloadImages('images/history_on.gif','images/cancel_on.gif','images/reset_on.gif','images/submit_on.gif','images/more1_on.gif','images/close0_on.gif','images/userinfo_on.gif')">
<form method="get" action="savemessage.asp" name="form1" >
  <table border="0" cellspacing="1" cellpadding="2" width="310">
    <tr> 
      <td colspan="2"><font color="#FFFFFF"><b>姓名 
        <input type="text" name="receiveuser" style="height:12pt; background-color:#A9C5E0" size="11" value="<%=session("receiveuser")%>">
         时间：<%=sj%>
        </b></font> </td>
    </tr>
    <tr> 
      <td colspan="2"> 
 <!--webbot bot="Validation" B-Value-Required="TRUE" I-Maximum-Length="10" --> 
 <textarea name="message" style="background-color:#A9C5E0" rows="6" cols="38" onkeydown=presskey()></textarea>
        <INPUT TYPE="hidden" name=id value="<%=request("id")%>">
      </td>
    </tr>
    <tr> 
      <td colspan="2"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image2','','images/history_on.gif',1)">
      <img name="Image2" border="0" src="images/history_off.gif" hspace="1" width="77" height="22" onClick="javascript:window.resizeTo(310,350);"></a> 
        <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image41','','images/reset_on.gif',1)"><img name="Image41" border="0" src="images/reset1_off.gif" hspace="1" onClick="MM_callJS('window.close();')" width="77" height="22"></a>
        
        <a href="Javascript:document.form1.submit();" onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_swapImage('Image5','','images/submit_on.gif',1)"> 
        <img name="Image5" border="0" src="images/submit1_off.gif" hspace="1" width="56" height="22" onclick="return checkform();"></a>  
        
        <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image6','','images/userinfo_on.gif',1)"><img name="Image6" border="0" src="images/userinfo_off.gif" onClick="MM_openBrWindow('userinfo.asp?id=<%=request("id")%>','job','scrollbars=yes,left=100,top=0,width=500,height=520')" hspace="1" width="56" height="22"></a> 
      </td>   
      
    </tr>   
    <%
	set rs=server.createobject("ADODB.recordset")
    rs.open "select * from chat where send='"&session("Uid")&"' and receive='"&request("id")&"'order by id desc",conn,1,1
	%>
    <tr>    
      <td colspan="2">    
        <textarea name="textarea"  rows="8" cols="38">
<%
if not rs.eof then
do while not (rs.eof or rs.bof)
%>
<%=rs("message")%>

<%
rs.movenext 
loop 
end if%>

</textarea>   
      </td>   
    </tr>   
    <tr>    
         
    <td colspan="2" align="right"> 
    <a href="clearMsg.asp?id=<%=request("id")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','images/cancel_on.gif',1)"><img name="Image8" border="0" src="images/cancel_off.gif" width="77" height="22"></a> 
    <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image7','','images/close0_on.gif',1)"><img name="Image7" border="0" src="images/close0_off.gif" onClick="javascript:window.resizeTo(310,190);" hspace="10"></a></td>                  
    </tr>                  
  </table>                  
</form>                  
</body>                  
</html>                  
