<!--#INCLUDE FILE="data.asp" -->
<%
if Session("Urule")<>"a" then
response.redirect "error.asp?id=admin"
end if
%>
<!--#INCLUDE FILE="check.asp" -->

<HTML><HEAD><TITLE>addfile_manager</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="oa.css" rel=stylesheet>
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<SCRIPT>

function js_openpage(url) {
  var 
newwin=window.open(url,"NewWin","toolbar=no,resizable=yes,location=no,directories=no,status=no,menubar=no,scrollbars=yes,top=220,left=220,width=500,height=310");
 // newwin.focus();
  return false;
} 
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall' & e.name!='selected_id')
       e.checked = form.chkall.checked; }
  }

 function del(url) 
 {  
  if (confirm("是否要删除该单位报送的文件")) 
  {
     window.open(url,"NewWin","toolbar=no,resizable=yes,location=no,directories=no,status=no,menubar=no,scrollbars=yes,top=220,left=220,width=450,height=200");
  }
} 

</SCRIPT>
<script language="JavaScript">
<!--
function MM_goToURL() { //v3.0
if (confirm("是否要删除该单位报送的文件")==1) {
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2)  eval(args[i]+".location='"+args[i+1]+"'");
}
}
//-->
function cform(){
 if(!confirm("是否要删除该单位报送的文件！"))
 return false;

}
</script>
<%
myUid=Session("Uid")
myUname=Session("Uname")
myUpass=Session("Upass")
myUrealname=Session("Rname")
myUpart=Session("Upart")
myUrule=Session("Urule")
myUlogin=Session("Ulogin")
if myUrule="a" then my_yonghu_quanxian="管理员"
if myUrule="b" then my_yonghu_quanxian="高级用户"
if myUrule="c" then my_yonghu_quanxian="普通用户"
my_biaoti=my_yonghu_quanxian&"："&myUrealname
%>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0>
<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
  <TBODY>
  <TR>
    <TD bgColor=#4e5960 class=heading colSpan=2 height=3></TD></TR>
  <TR>
    <TD bgColor=#4e5960 class=heading>　<FONT 
    color=#ffffff><B>处机关收取文件列表</B></FONT></TD>
  </TR>
  <TR>
    <FORM action="shouqu.asp" method=post name=sele>
    <TD vAlign=top>
        <TABLE border=0 cellPadding=2 cellSpacing=1 width="100%">
          <TBODY> 
          <TR bgColor=#999999> 
            <TD class=heading colspan=6> 
              <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
                <TBODY> 
                <TR> 
                  <TD align=right width="2%"><IMG height=18 
                  src="images/adorn.gif" width=10></TD>
                  
                  <TD align=right><%   

dim keyword
keyword=request("key")

dim page
page=request("page")
PageSize = 15
dim rs,strSQL,news
strSQL ="select * from jhtdata where type=1 ORDER BY 时间 desc,id DESC"
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
%> <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr> <form method=Post action="shouqu.asp"><FONT COLOR="#ffffff">
               [<b><%=rs.pagecount%></b>/<%=page%>页] [共<%=totalfilm%>个] <%if page=1 then%> [首 页] [上一页] <% else %> [<a href="shouqu.asp?page=1">首 页</a>] 
               [<a href="shouqu.asp?page=<%=page-1%>">上一页</a>]<%end if%><%if rs.pagecount-page<1 then%> [下一页] [尾 页]  <%else%> [<a href="shouqu.asp?page=<%=page+1%>">下一页</a>]  [<a href="shouqu.asp?page=<%=rs.pagecount%>">尾 页</a>]</FONT> <%end if%> <input type='text' name='page' size=2 maxlength=10 style="font-size:9pt;color:#FFFFFF;background-color:#666666;border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000" value="<%=page%>" align=center> <input style="border:1 solid black;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; FONT-WEIGHT: normal; HEIGHT: 18px; LINE-HEIGHT: normal" type='submit'  value=' Goto '   size=2></td>
             </tr>
           </table></TD>
                  <TD width="3%"><IMG height=18 
                  src="images/adorn.gif" 
            width=10></TD>
                </TR>
                </TBODY> 
              </TABLE>
            </TD>
          </TR>
          <TR bgColor=#bfbfbf align=center> 
            <TD><b>单位名称</b></TD>
            <TD> 
              <div ><b>文件报送人</b></div>
            </TD>
            <TD><b>报送日期</b></TD>
			<TD><b>详细信息</b></TD>
            <TD align=middle><B>删除</B></TD>
          </TR>
<%
count=0 
do while not (rs.eof or rs.bof) and count<rs.PageSize 
%>   
          <TR> 
            <TD bgColor=#efefef><%=rs("部门")%></TD>
            <TD bgColor=#efefef><%=rs("真实姓名")%></TD>
            <TD bgColor=#efefef><%=rs("时间")%></TD>
			<TD bgColor=#efefef align=center><A HREF="#" onClick="MM_openBrWindow('soft.asp?id=<%=rs("id")%>','','width=500,height=300')"><img src="images/detail_off.gif" border=0></A></TD>
            <TD align=middle bgColor=#efefef><a href="del_from_db.asp?delid=<%=rs("id")%>&delbz=My_only" onclick="return cform();"><IMG 
            border=0 height=16 name=Image101 src="images/dele_1.gif" 
            width=14></A></TD>
          </TR>

		 <%rs.movenext 
count=count+1
loop 
end if%>    
          <tr>
		    <td colspan=6 bgcolor=#9c9a9c style="color:red">（注：建议用下载工具下载或在“下载”链接上点右键选择“目标另存为”）
			</td>
		  </tr>
          </TBODY> 
        </TABLE>
        <BR>
      </TD></FORM></TR></TBODY></TABLE></BODY></HTML>