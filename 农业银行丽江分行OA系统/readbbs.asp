<!--#include file="data.asp"-->
<!--#include file="check.asp"-->
<script>
function OpenWindows(url)
{
  var 
 newwin=window.open(url,"_blank","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=50,left=120,width=600,height=400");
 return false;
 
}
function OpenSmallWindows(url)
{
  var 
 newwin=window.open(url,"_blank","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=50,left=120,width=480,height=320");
 return false;
 
}
</script>

<script language=JavaScript>
function subchk()
{
if(document.form1.title.value=="")
{
alert("请输入你的文章标题!\n");
return  false;
}
if(document.form1.content.value=="")
{
alert("请输入你的文章内容!\n");
return  false;
}

}
</script>
<%
strSql="select * from bbs where Id="&request("SubjectId")&" ORDER BY time desc, id DESC"
set my_rs=server.createobject("adodb.recordset")
my_rs.open strsql,conn,3,3
my_rs("Knock")=my_rs("Knock")+1
my_rs.update

%>
<html>
<head>
<title><%=my_rs("Subject")%></title>
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
</head>
<body bgcolor="#efefef" topmargin="0" leftmargin="0" onLoad="MM_preloadImages('images/iwantanswer_on.gif','images/rarticle_on.gif','images/newarticle_on.gif','images/newarticle1_on.gif','images/close_on.gif','images/sendarticle_on.gif','images/rewrite_on.gif')">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr bgcolor="#4e5960"> 
      <td  class="heading"><b>　<font color="#FFFFFF">讨论中心</font></b></td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="1" cellpadding="2">
    <tr> 
      <td  class="heading" width="100%" colspan=3 bgcolor="#bfbfbf"><img src="images/<%=my_rs("pic")%>"><%=my_rs("Subject")%></td>
	  </tr>
	  
	  <tr>
      <td  class="heading" width="100%" colspan=3 bgcolor="#DFDFDF"> 
        <div align="center"><font  color=red><%=my_rs("name")%></font>  发表于
        <font size="2"><i><font style='font-size:9pt;color:gray'>【<%=my_rs("time")%>】</font></i></font></div> 
      </td>
    </tr>
	<tr><td align=center  bgcolor="#ffffff">
	<table border=0 width=95%>
    <tr> 
      <td class="show"><%=my_rs("content")%><br>
</td>
    </tr>
	</table>
	</td></tr>
    <tr> 
      <td colspan=3 class="show" align=right><a href="addbbs.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image12','','images/newarticle1_on.gif',1)"><img name="Image12" border="0" src="images/newarticle1_off.gif" width="69" height="19" hspace="5"></a><a href="addbbs.asp?id=<%=my_rs("id")%>" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image14','','images/rarticle_on.gif',1)"><img name="Image14" border="0" src="images/rarticle_off.gif" height="19" hspace="5"></a><a href="Javascript:window.close();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image13','','images/close1_on.gif',1)"><img name="Image13" border="0" src="images/close1_off.gif" width="69" height="19" hspace="5"></a></td>
    </tr>
 <%
	strSql="select * from bbs where SubjectId="&request("SubjectId")&" ORDER BY time desc, id DESC"
    set rs=server.createobject("adodb.recordset")
    rs.open strsql,conn,1,1
    if not rs.eof then


	%>

    <tr bgcolor="#4e5960"> 
      <td  class="heading" colspan=3 height="21"><b>　<font color="#FFFFFF">相关回帖</font></b></td>
    </tr>
       
   <%	do while not rs.eof
   %>
    <tr> 
      <td bgcolor="#bfbfbf" class="heading" width="100%"  colspan="3"><img src="images/note1.gif"><%=rs("subject")%></td>
	  </tr>
	  <tr>
      <td bgcolor="#dfdfdf" class="heading" width="100%"  colspan="3">
        <div align="center"><font size="2"><%=rs("name")%></font>
		 <font size="2"><i><font style='font-size:9pt;color:gray'><%=rs("time")%></font></i></font></div>
      </td>
    </tr>
	
	
<tr><td align=center  bgcolor="#ffffff">
	<table border=0 width=95%>
    <tr> 
      <td width=95% colspan=3 class="show" bgcolor="#ffffff"><%=rs("content")%></td>
    </tr>
	</table>
      </td>
    </tr>
	<%rs.movenext
i=i+1
loop
else rsponse.write " "
end if
%>

          </table>   
      </td>   
    </tr>   
  </table> 
  <%
  	my_rs.close
	set my_rs=nothing

  rs.close
  set rs=nothing
  %>
</body>   
</html>   
