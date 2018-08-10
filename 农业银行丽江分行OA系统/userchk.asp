<!--#INCLUDE FILE="data.asp" -->
<%
if Session("Urule")<>"a" then
response.redirect "error.asp?id=admin"
end if
%><!--#INCLUDE FILE="check.asp" -->

<HTML><HEAD><TITLE>uesercheck</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="oa.css" rel=stylesheet>
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
function cform(){
 if(!confirm("您确认删除此用户！"))
 return false;

}
function pass(){
 if(!confirm("您确认该用户通过审核！"))
 return false;

}

</script>
</script> <script language="JavaScript">
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
</script> <script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
//-->
</script
></HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0>
<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%"> <TBODY> <TR> <TD bgColor=#4e5960 class=heading colSpan=2 height=3></TD></TR> 
<TR> <TD bgColor=#4e5960 class=heading>　<FONT 
    color=#ffffff><B>注册用户管理 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="#" onClick="MM_openBrWindow('register.asp','','width=510,height=430')">添加新用户</a></B></FONT></TD></TR> <TR> <FORM action="userchk.asp" method=post name=sele><TD vAlign=top><TABLE border=0 cellPadding=2 cellSpacing=1 width="100%"> 
<TBODY> <TR bgColor=#999999> <TD class=heading colspan=9><TABLE border=0 cellPadding=0 cellSpacing=0 width="100%"> 
<TBODY> <TR> <TD align=right width="2%"><IMG height=18 
                  src="images/adorn.gif" width=10></TD><TD align=right><%   


dim page
page=request("page")
PageSize = 15
dim rs,strSQL,news
strSQL ="select * from user ORDER BY id DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSQL,Conn,3,3
    rs.PageSize = PageSize
	totalfilm=rs.recordcount
    pgnum=rs.Pagecount
    if page="" or clng(page)<1 then page=1
    if clng(page) > pgnum then page=pgnum
    if pgnum>0 then rs.AbsolutePage=page

if rs.eof then
response.write "<font color='#ffffff' class='3dfont'>还没有任何东东</font>"
else
%> <table border="0" cellspacing="0" cellpadding="0" width="100%"> <tr> <form method=Post action="userchk.asp"><FONT COLOR="#ffffff"> 
[<b><%=rs.pagecount%></b>/<%=page%>页] [共<%=totalfilm%>个] <%if page=1 then%> [首 
页] [上一页] <% else %> [<a href="userchk.asp?page=1">首 页</a>] [<a href="userchk.asp?page=<%=page-1%>">上一页</a>]<%end if%><%if rs.pagecount-page<1 then%> 
[下一页] [尾 页] <%else%> [<a href="userchk.asp?page=<%=page+1%>">下一页</a>] [<a href="userchk.asp?page=<%=rs.pagecount%>">尾 
页</a>]</FONT> <%end if%> <input type='text' name='page' size=2 maxlength=10 style="font-size:9pt;color:#FFFFFF;background-color:#666666;border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000" value="<%=page%>" align=center> 
<input style="border:1 solid black;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; FONT-WEIGHT: normal; HEIGHT: 18px; LINE-HEIGHT: normal" type='submit'  value=' Goto '   size=2></td> </tr>
</table></TD> <TD width="3%"><IMG height=18 
                  src="images/adorn.gif" 
            width=10></TD></TR> </TBODY> </TABLE></TD> </TR> <TR bgColor=#bfbfbf align=center> 
<TD><b>姓名</b></TD><TD><b>用户名</b></TD><TD><b>注册时间</b></TD><TD><b>所属单位</b></TD><TD><b>级别</b></TD><TD><b>审核</b></TD><TD><B>修改</B></TD><TD><b>删除</b></TD><TD><b>权限</b></TD></TR> 
<%
count=0 
do while not (rs.eof or rs.bof) and count<rs.PageSize 
%> <TR> <td height="23" bgColor=#efefef> <p align="center"><%=rs("姓名")%> </td><td bgColor=#efefef> 
<p align="center"><%=rs("用户名")%> </td><td bgColor=#efefef> <p align="center"><%=rs("时间")%> 
</td><td bgColor=#efefef> <p align="center"><%=rs("部门")%> </td><td bgcolor=#efefef> 
<p align="center"><%=rs("ilevel")%></td><td align="center" bgColor=#efefef><% if rs("审核")=true then %>已审核<%else%><a href="shenghe.asp?id=<%=rs("id")%>" onclick="return pass();">待审核</a><%end if%></td><td bgColor=#efefef> 
<p align="center"><a href="#" onClick="MM_openBrWindow('edit.asp?id=<%=rs("id")%>','','width=500,height=300')">修改 
</a> </td><td  bgColor=#efefef> <p align="center"> <a href="dele.asp?id=<%=rs("id")%>" onclick="return cform();">删除</a> 
</td><td bgColor=#efefef><%
						  if rs("权限")="a" then 
						  response.write "超级用户"  
						  else if rs("权限")="b" then 
						  response.write "管理员" 
						  else 
						  response.write "普通用户" 
						  end if
						  end if%></td></TR> <%rs.movenext 
count=count+1
loop 
end if%> <tr> <td colspan=9 bgcolor=#9c9a9c style="color:red">（注:可以通过修改用户资料来更改用户权限和邮箱级别） 
</td></tr> </TBODY> </TABLE></TBODY></TABLE>
</BODY></HTML>
