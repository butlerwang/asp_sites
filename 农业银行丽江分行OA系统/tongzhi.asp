<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->

<%if Request("actt")="myreload" then%>
<SCRIPT language=Javascript>
<!-- hide
{
window.location.reload(true);
}
// -->
</SCRIPT>
<%end if%>
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
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

<html><head><title>info_note</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
</head>
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
function cform(){
 if(!confirm("您确认删除此通知！请注意删除后无法恢复"))
 return false;

}//-->
</script>
<script language="JavaScript">
<!--
function MM_goToURL() { //v3.0
if (confirm("您是否要删除此条信息？按确定删除。")==1) {
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2)  eval(args[i]+".location='"+args[i+1]+"'");
}
}
//-->
</script>

<script>
function js_openpage(url) {
  var 
newwin=window.open(url,"NewWin","toolbar=no,resizable=yes,location=no,directories=no,status=no,menubar=no,scrollbars=yes,top=220,left=220,width=500,height=230");
 // newwin.focus();
  return false;
}

function js_openpage_1(url) {
  var 
newwin=window.open(url,"NewWin","toolbar=yes,resizable=yes,location=yes,directories=no,status=yes,menubar=yes,scrollbars=yes,top=220,left=220,width=500,height=310");
 // newwin.focus();
  return false;
}
</script>

<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0"
style="background-image: url('images/main_bg.gif'); background-attachment: scroll; background-repeat: no-repeat; background-position: left bottom" onLoad="MM_preloadImages('images/manage_on.gif')" >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="heading" bgcolor="#4e5960" colspan="2" height="3"></td>
  </tr>
  <tr>
    <td bgcolor="#4e5960" class="heading" width="85%">　<font color="#FFFFFF"><b>内部通知</b></font></td>
    <td bgcolor="#4e5960" width="15%">&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="100%">
  <tr> 
    <td class="heading" colspan="2" bgcolor="#4e5960"></td>
  </tr>
  <tr> 
    <td width="109" valign="top">&nbsp;</td>
    <td valign="top"> 
    <%
if Session("Ulogin")<>"yes" then
    Response.Redirect ("login.asp")
end if
myUid=Session("Uid")
myUname=Session("Uname")
myUpass=Session("Upass")
myUrealname=Session("Urealname")
myUpart=Session("Upart")
myUrule=Session("Urule")
myUlogin=Session("Ulogin")
if myUrule="a" then my_yonghu_quanxian="管理员"
if myUrule="b" then my_yonghu_quanxian="高级用户"
if myUrule="c" then my_yonghu_quanxian="普通用户"
my_biaoti=my_yonghu_quanxian&"："&myUrealname
%>

<table border="0" width="100%" bgcolor="#BFBFBF" cellspacing="1">
        <tr>
          <td bgcolor="#BFBFBF">
          <TABLE width=100%>
          <TR align=center>
          <td width="86%" class="heading">　<font color="red"><b>通  知</b></font></td>
          <td width="14%"><%if myUrule="a" then %><a href="#" onclick="window.open('sendinf.asp','','width=400 height=400')"><img name="Image5" border="0" src="images/putout.gif" width="85" height="19"></a><%else%>&nbsp;<%end if%></td>
          </TR>
          </TABLE></td> 
        </tr>
        <tr>
          <td>
<%Set mrs= Server.CreateObject("ADODB.Recordset") 
strSql="select top 1 * from jhtdata where type=0 order by id desc"
mrs.open strSql,Conn,1,1 
if mrs.eof then
response.write "还没有任何通知"
else
%>

          <TABLE width=100% border=1  bgcolor="#efefef" bordercolorlight=#ffffff bordercolordark=#ffffff cellspacing=0 cellpadding=0>
          <TR align=center>
              <TD bgcolor="#efefef"><FONT style="font-size:11pt"><B><%=mrs("标题")%></B></FONT></TD>
          </TR>
          <TR>
              <TD align=right style="font-size:8pt"><%=mrs("部门")%>　<%=mrs("真实姓名")%> 发布于 <%=mrs("时间")%></TD>
          </TR>
          <TR>
              <TD bgcolor="#ffffff" height=80 valign=top><%=mrs("内容")%></TD>
          </TR>
          </TABLE>

<%
mrs.colse
set mrs=nothing
end if%>

          </td>
        </tr>
        <tr>
          <td width="100%" bgcolor="#FFFFFF" colspan=3>
            
            <table border="1" width="100%" cellspacing="0" cellpadding=0 bordercolordark=#ffffff bordercolorlight=#ffffff>
              <%   

dim keyword
keyword=request("selecttext")

dim page
page=request("page")
PageSize = 12
dim rs,strSQL,news
strSQL ="select * from jhtdata where type=0 ORDER BY 时间 desc,id DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSQL,Conn,3,3
    rs.PageSize = PageSize
    totalfilm=rs.recordcount
    pgnum=rs.Pagecount
    if page="" or clng(page)<1 then page=1
    if clng(page) > pgnum then page=pgnum
    if pgnum>0 then rs.AbsolutePage=page

if rs.eof then
%>
<tr>
                <td width="50%"><img border="0" src="images/icon_group.gif" width="15" height="15" align="absmiddle"> 
                  暂无通知</td> 
                <td width="37%">暂无通知</td>
                <td width="4%"><img border="0" src="images/dele_1.gif" alt="删除此通知！" align="absmiddle" width="15" height="15"></td>
              </tr>
<%else
%>   <tr>
               <td colspan=3> 
               <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr bgcolor="#303430"> <form method=Post action="tongzhi.asp"><FONT COLOR="#ffffff">
               [<b><%=rs.pagecount%></b>/<%=page%>页] [<%=totalfilm%>] <%if page=1 then%> [上一页] <% else %>  
               [<a href="tongzhi.asp?page=<%=page-1%>">上一页</a>]<%end if%><%if rs.pagecount-page<1 then%> [下一页] <%else%> [<a href="tongzhi.asp?page=<%=page+1%>">下一页</a>] </FONT> <%end if%> <input type='text' name='page' size=2 maxlength=10 style="font-size:9pt;color:#FFFFFF;background-color:#666666;border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000" value="<%=page%>" align=center> <input style="border:1 solid black;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; FONT-WEIGHT: normal; HEIGHT: 18px; LINE-HEIGHT: normal" type='submit'  value=' Goto '   size=2></td>
             </tr>
           </table>
               </td>
             </tr>
<%
count=0 
do while not (rs.eof or rs.bof) and count<rs.PageSize 
my_shuoming=rs("时间")
%>     
              <tr bgcolor="#efefef">
                <td><a href="#" onClick="MM_openBrWindow('view_inf.asp?view_id=<%=rs("id")%>','','width=400,height=240,scrollbars=yes')"><FONT COLOR="#000000"><%=rs("标题")%></FONT></a></td>
                <td><%=my_shuoming%></td>
                <%if myUrule="a" then%><td><a href="del.asp?delid=<%=rs("id")%>&delbz=My_public" onclick="return cform();"><img border="0" src="images/dele_1.gif" alt="删除此条信息！" align="absmiddle" width="15" height="15"></td><%end if%>
              </tr>
              
        <%rs.movenext 
count=count+1
loop 
end if%>            
  
        
            </table>
          </td>
        </tr>
      </table>　
     

    </td>                                                                                               </tr>                       
</table>                                                                                                

</body>                                                                                                                                           
</html>                                                                                                                                           
