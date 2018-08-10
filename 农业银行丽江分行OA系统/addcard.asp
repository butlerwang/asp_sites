<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<!--#INCLUDE FILE="html.asp" -->

<%
name=request("cname")


if name="" then
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

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
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





<script>
function OpenWindows(url,widthx,heighx)
{
  var 
 newwin=window.open(url,"_blank","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=20,left=60,width=600,height=500");
 return false;
 
}
</script>

<title>新增个人名片</title>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
</head>

<body bgcolor="#efefef" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form action="addcard.asp" method="post" name="AddCard" onsubmit="return ChkCard();">
  <table width="100%" border="0" cellspacing="1" cellpadding="2">
    <tr bgcolor="#4E5960"> 
      <td class="heading" height="20"><font color="#FFFFFF"><b>新增个人名片</b></font></td>
    </tr>
  </table>
  <table width="100%" border="1" cellspacing="0" cellpadding="1" bordercolordark="#FFFFFF" bordercolor="#000000">
    <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">姓名：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="cname" maxlength="20" size="20">
    </td>
  </tr>
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">单位：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="company" size="60" maxlength="100">
    </td>
  </tr>
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">通讯地址：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="comaddress" size="60" maxlength="200">
    </td>
  </tr>
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">邮编：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="postcode" size="10" maxlength="6">
    </td>
  </tr> 
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">职务：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="duty" size="20" maxlength="50">
    </td>
  </tr>  
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">电话：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="phone" size="20" maxlength="50">
    </td>
  </tr>
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">传真：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="fax" size="20" maxlength="50">
    </td>
  </tr>
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">手机：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="handset" size="20" maxlength="50">
    </td>
  </tr>
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">信箱：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="email" size="40" maxlength="100">
    </td>
  </tr>
   
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">备注：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <textarea name="remark" cols="60" rows="4"></textarea>
    </td>
  </tr>
</table>
<div align="center"><br> 
  </div> 
  <div align="center"> &nbsp; 
  <input type="image" src="images/sendarticle_off.gif" WIDTH="60" HEIGHT="19">
  <a href="javascript:document.AddCard.reset();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image2','','images/rewrite_on.gif',1)"><img name="Image2" border="0" src="images/rewrite_off.gif" width="60" height="19" hspace="5"></a><a href="Javascript:window.close();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image3','','images/close_2.gif',1)"><img name="Image3" border="0" src="images/close_1.gif" width="85" height="19" hspace="5"></a> 
  </div> 
</form>
</body>
</html>
<script language="javascript">
function ChkCard()
{
    if (document.AddCard.cname.value=="" )
    {
        alert("注意！姓名不能为空哦！");
        return false;
    } 
       
}
</script>





<%
else
set rs=server.createobject("ADODB.recordset") 
rs.Open "SELECT * FROM card Where ID is null",conn,1,3 
rs.addnew

rs("cname")=request("name")
rs("company")=request("company")
rs("comaddress")=request("comaddress")
rs("postcode")=request("postcode")
rs("duty")=request("duty")
rs("phone")=request("phone")
rs("fax")=request("fax")
rs("handset")=request("handset")
rs("email")=request("email")
rs("remark")=request("remark")
rs("userid")=session("Uid")
rs.update 
id=rs("id")
rs.close
set rs=nothing
%>
<script language=javascript>
opener.location=opener.location;
</script>

<script>
function OpenWindows(url,widthx,heighx)
{
  var 
 newwin=window.open(url,"_blank","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=20,left=60,width=600,height=500");
 return false;
 
}
</script>

<title>修改个人名片</title>
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
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
</head>
<body bgcolor="#efefef" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/delete_on.gif','images/modify_on.gif','images/close_2.gif')">
<form action="editcard.asp" method=post name=AddCard onsubmit="return ChkCard();">
<table width="100%" border="0" cellspacing="1" cellpadding="2">
    <tr bgcolor="#4E5960"> 
      <td class="heading" height="20"><font color="#FFFFFF"><b>个人名片成功添加</b></font></td>
    </tr>
  </table>
  <table width="100%" border="1" cellspacing="0" cellpadding="1" bordercolordark="#FFFFFF" bordercolor="#000000">
    <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">姓名：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="cname" maxlength="20" size="20" value="<%=request("cname")%>">
      <input type="hidden" name="id"  value="<%=id%>">
    </td>
  </tr>

  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">单位：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="company" size="60" maxlength="100" value="<%=request("company")%>">
    </td>
  </tr>
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">通讯地址：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="comaddress" size="60" maxlength="200" value="<%=request("comaddress")%>">
    </td>
  </tr>


  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">邮编：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="postcode" size="10" maxlength="6" value="<%=request("postcode")%>">
    </td>
  </tr>
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">职务：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="duty" size="20" maxlength="50" value="<%=request("duty")%>">
    </td>
  </tr>  
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">电话：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="phone" size="20" maxlength="50" value="<%=request("phone")%>">
    </td>
  </tr>
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">传真：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="fax" size="20" maxlength="50" value="<%=request("fax")%>">
    </td>
  </tr>
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">手机：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="handset" size="20" maxlength="50" value="<%=request("handset")%>">
    </td>
  </tr>
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">信箱：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <input type="text" name="email" size="40" maxlength="100" value="<%=request("email")%>">
    </td>
  </tr>
   
  <tr> 
      <td width="15%" bgcolor="#bfbfbf"> 
        <div align="center">备注：</div>
    </td>
      <td width="85%" bgcolor="#efefef"> 
        <textarea name="remark" cols="60" rows="4"><%=request("remark")%></textarea>
    </td>
  </tr>
</table>
<div align="center"><br>
        
        <a href="Javascript:DelChk(70);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','images/delete_on.gif',1)"><img name="Image1" border="0" src="images/delete_off.gif" width="60" height="19" hspace="5"></a>&nbsp; 
        <a href="Javascript:document.AddCard.submit();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image2','','images/modify_on.gif',1)"><img name="Image2" border="0" src="images/modify_off.gif" width="60" height="19" hspace="5"></a>&nbsp; 
        
        <a href="Javascript:window.close();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image3','','images/close_2.gif',1)"><img name="Image3" border="0" src="images/close_1.gif" width="85" height="19" hspace="5"></a>
</div>
</form>

</body>
</html>
<script language="javascript">
function DelChk(cardid)
    {
        if(confirm("确认删除吗?"))
            document.location="delcard.asp?flag=0&cardid="+cardid ;
    }
function ChkCard()
{
    if (document.AddCard.cname.value=="" )
    {
        alert("注意！姓名不能为空哦！");
        return false;
    }      
}
</script>
<%end if%>
