<%
if session("Urule")<>"a" then
response.redirect "error.asp?id=admin"
end if
%>
<!--#include file="data.asp"-->
<html><head><title>url</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
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

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}
//-->
</script>
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
<script language="JavaScript">
    function  validate()
    {
       
        if  (document.myform.src.value=="")
        {
            alert("图片地址不能为空");
            document.myform.src.focus();
            return false ;
        }
        if  (document.myform.alt.value=="")
        {
            alert("说明不能为空");
            document.myform.alt.focus();
            return false ;
        }
     
}

function cform(){
 if(!confirm("你是否确认删除？"))
 return false
 else
 return true

}
</script>
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" onLoad="MM_preloadImages('images/add_on.gif','images/modify_on.gif','images/delete_on.gif','images/add_on.gif','images/modi_2.gif','images/dele_2.gif','images/showall_on.gif')">
<tr> <td class="heading" bgcolor="#4e5960" colspan="2" height="3"></td></tr> <tr> 
<td class="heading" bgcolor="#4e5960" colspan="2" >　<font color="#FFFFFF"><b>广告管理</b></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a HREF="#" onclick="window.open('adrot_add.asp','','width=400 height=250')"><img src="images/add_off.gif" align=middle border=0></a></td></tr> 
<table width="100%" border="0" cellspacing="1" cellpadding="2"> <tr bgcolor="#999999"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0"> <tr> <td width="2%" align="right"><img src="images/adorn.gif" width="10" height="18"></td><td align="left"> 
</td><td colspan="3" align="right" valign="middle"><%   
dim page
page=request("page")
PageSize = 3
dim rs,strSQL,news
strSQL ="SELECT * FROM adrot ORDER BY id DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSQL,Conn,1,1
    rs.PageSize = PageSize
    totalfilm=rs.recordcount
    pgnum=rs.Pagecount
    if page="" or clng(page)<1 then page=1
    if clng(page) > pgnum then page=pgnum
    if pgnum>0 then rs.AbsolutePage=page

if rs.eof then
response.write "<font color='#ffffff' class='3dfont'>还没有任何东东</font>"
else
%> <table border="0" cellspacing="0" cellpadding="0" width="100%"> <tr bgcolor="#303430"> <form method=Post action="adrot.asp">
<font COLOR="#ffffff"> [<b><%=rs.pagecount%></b>/<%=page%>页] [共<%=totalfilm%>个] 
<%if page=1 then%> [首 页] [上一页] <% else %> [<a href="?page=1">首 页</a>] [<a href="?page=<%=page-1%>">上一页</a>]<%end if%><%if rs.pagecount-page<1 then%> 
[下一页] [尾 页] <%else%> [<a href="?page=<%=page+1%>">下一页</a>] [<a href="?page=<%=rs.pagecount%>">尾 
页</a>]</font> <%end if%> <input type='text' name='page' size=2 maxlength=10 style="font-size:9pt;color:#FFFFFF;background-color:#666666;border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000" value="<%=page%>" align=center> 
<input style="border:1 solid black;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; FONT-WEIGHT: normal; HEIGHT: 18px; LINE-HEIGHT: normal" type='submit'  value=' Goto '   size=2> </td>
</tr></form> </table></td> <td width="3%"><img src="images/adorn.gif" width="10" height="18"></td></tr> 
</table> <tr> <td> <%
count=0 
do while not (rs.eof or rs.bof) and count<rs.PageSize 
%> <table border="1" cellspacing="0" cellpadding="0" width="100%" bordercolorlight=000000 bordercolordark=ffffff> 
<tr> <td colspan=2><%if rs("type")="GIF" then%><img src="<%=rs("src")%>"><%else%><%end if%></td></tr> 
<tr> <td>说明:<textarea NAME="alt" ROWS="2" COLS="20" style="overflow: auto"><%=rs("alt")%></textarea> 
</td><td> 图片:<input TYPE="text" value="<%=rs("src")%>" name="src" size=16> <input TYPE="text" NAME="width" value="<%=rs("width")%>" size=2>×<input TYPE="text" NAME="height" value="<%=rs("height")%>" size=1> 
<select NAME="type" style="height:18px;font-size:9pt"><option value="GIF" <%if rs("type")="GIF" then response.write " selected" end if%>>GIF</option><option value="SWF" <%if rs("type")="SWF" then response.write " selected" end if%>>SWF</option></select><br>链接:<input TYPE="text" value="<%=rs("url")%>" name="url"><input TYPE="submit" value="修改" name="edit"><input TYPE="submit" value="删除" name="del" onclick="return cform();"><input TYPE="hidden" name="id" value="<%=rs("id")%>"> 
</td></tr> </table><%rs.movenext 
count=count+1
loop 
end if%> </td></tr> </table>   
        </body>                       
              
