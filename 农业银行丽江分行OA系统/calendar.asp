<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
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
 newwin=window.open(url,"_blank","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=50,left=120,width=550,height=270");
 return false;
 
}
</script>

<html>
<head>
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
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF" onLoad="MM_preloadImages('images/add_on.gif')">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="heading" bgcolor="#4e5960" colspan="2" height="3"></td>
  </tr>
  <tr> 
    <td class="heading" bgcolor="#4e5960" colspan="2">��<font color="#FFFFFF"><b>�ճ̰���</b></font></td>
  </tr>
	
    <td width="100%" bgcolor="#BFBFBF"> 
      <table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#FFFFFF">
        <tr bgcolor="#999999"> 
            <td colspan="5" align="right"> 
      
      <table border=0 width=100% bgcolor="#999999" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="2%" align="right"><img src="images/adorn.gif" width="10" height="18"></td>
          <td width=65><a href="addcalendar.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','images/add_on.gif',1)" onClick="return OpenSmallWindows(this.href);"><img name="Image1" border="0" src="images/add_off.gif"></a></td>
          <td><%   
dim page
page=request("page")
PageSize = 15
dim rs,strSQL,news
strSQL ="SELECT * FROM calendar where userid="&session("Uid")&" ORDER BY id DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSQL,Conn,3,3
    rs.PageSize = PageSize
	totalfilm=rs.recordcount
    pgnum=rs.Pagecount
    if page="" or clng(page)<1 then page=1
    if clng(page) > pgnum then page=pgnum
    if pgnum>0 then rs.AbsolutePage=page

if rs.eof then
response.write "<font class='3dfont'>��û���κζ���</font>"
else
%> <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr> <form method=Post action=""><FONT COLOR="#ffffff">
               [<b><%=rs.pagecount%></b>/<%=page%>ҳ] [��<%=totalfilm%>��] <%if page=1 then%> [�� ҳ] [��һҳ] <% else %> [<a href="?page=1">�� ҳ</a>] 
               [<a href="?page=<%=page-1%>">��һҳ</a>]<%end if%><%if rs.pagecount-page<1 then%> [��һҳ] [β ҳ]  <%else%> [<a href="?page=<%=page+1%>">��һҳ</a>]  [<a href="?page=<%=rs.pagecount%>">β ҳ</a>]</FONT> <%end if%> <input type='text' name='page' size=2 maxlength=10 style="font-size:9pt;color:#FFFFFF;background-color:#666666;border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000" value="<%=page%>" align=center> <input style="border:1 solid black;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; FONT-WEIGHT: normal; HEIGHT: 18px; LINE-HEIGHT: normal" type='submit'  value=' Goto '   size=2></td>
             </tr>
           </table></td>
		   
          <td width="3%"><img src="images/adorn.gif" width="10" height="18"></td>
        </tr>
      </table>
        <TR bgcolor="#bfbfbf" align=center> 
          <TD width=450 ><b>�ճ�����</b></TD>           
            
          <TD width=120> <b>�ʱ��</b> </TD>           
            
          <TD width=50> <b>״̬</b> </TD>                       
        </TR>                    
        <%
count=0 
do while not (rs.eof or rs.bof) and count<rs.PageSize 
%>   
          
        <TR bgcolor="#efefef"> 
          <TD><a href="modCalendar.asp?Id=<%=rs("id")%>" onclick="return OpenSmallWindows(this.href);"><%=rs("title")%></a></TD>           
            
          <TD><%=left(rs("time"),4)%>/<%=mid(rs("time"),5,2)%>/<%=mid(rs("time"),7,2)%>&nbsp;&nbsp;<%=mid(rs("time"),9,2)%>:<%=right(rs("time"),2)%></TD>           
            
          <TD><%if rs("state")=false then%><font color=red>δ����</font><%else%>������<%end if%></TD>                       
          </TR> 
		  <%rs.movenext 
count=count+1
loop 
end if%>     
             
        </TABLE>    
         </td>
  </tr>
</table>

</body>
</html>
