<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<script>
function OpenWindows(url,widthx,heighx)
{
  var 
 newwin=window.open(url,"_blank","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=20,left=60,width=600,height=500");
 return false;
 
}
</script>

<html><head><title>ecard_peson</title>
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



<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" onLoad="MM_preloadImages('images/detail_on.gif','images/add_on_2.gif','images/showall_on.gif','images/search_on.gif')">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="heading" bgcolor="#4e5960" colspan="2" height="3"></td>
  </tr>
  <tr> 
    <td class="heading" bgcolor="#4e5960" colspan="2" >　<font color="#FFFFFF"><b>个人通讯录</b></font></td>
  </tr>
    <tr>  
      
    <td width="100%" bgcolor="#BFBFBF"> 
      <table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#FFFFFF">
        <tr bgcolor="#999999"> 
            <td colspan="5" align="right"> 
              
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
			  <td width=15><img src="images/adorn.gif" height="18"></td><td align=center>
                
                  
                   <%   

dim keyword
keyword=request("key")

dim page
page=request("page")
PageSize = 15
dim rs,strSQL,news
if request("style")="cname" then
strSQL ="SELECT * FROM card where cname like '%"&keyword&"%' and userid="&session("Uid")&" ORDER BY id DESC"
else
strSQL ="SELECT * FROM card where company like '%"&keyword&"%' and userid="&session("Uid")&" ORDER BY id DESC"
end if
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
              <tr bgcolor="#303430"> <form method=Post action="txl.asp"><FONT COLOR="#ffffff">
               [<b><%=rs.pagecount%></b>/<%=page%>页] [共<%=totalfilm%>个] <%if page=1 then%> [首 页] [上一页] <% else %> [<a href="txl.asp?page=1">首 页</a>] 
               [<a href="txl.asp?page=<%=page-1%>">上一页</a>]<%end if%><%if rs.pagecount-page<1 then%> [下一页] [尾 页]  <%else%> [<a href="txl.asp?page=<%=page+1%>">下一页</a>]  [<a href="txl.asp?page=<%=rs.pagecount%>">尾 页</a>]</FONT> <%end if%> <input type='text' name='page' size=2 maxlength=10 style="font-size:9pt;color:#FFFFFF;background-color:#666666;border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000" value="<%=page%>" align=center> <input style="border:1 solid black;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; FONT-WEIGHT: normal; HEIGHT: 18px; LINE-HEIGHT: normal" type='submit'  value=' Goto '   size=2></td>
             </tr>
           </table>
              </td><td width=15>
			    <img src="images/adorn.gif" width="10" height="18"></td> 
              </tr>
            </table>
            </td>
          </tr>

      
          <tr bgcolor="#bfbfbf"> 
            <td width="60"><b>姓名</b></td>
            <td><b>单位</b></td>
            <td width="65" align="center"><b>手机</b></td>
            <td width="40" align="center"><b>详情</b></td>
          </tr>
         <%
count=0 
do while not (rs.eof or rs.bof) and count<rs.PageSize 
%>   
          <tr> 
            <td width="60" bgcolor="#efefef"><%=rs("cname")%></td>
            <td bgcolor="#efefef"><%=rs("company")%></td>
            <td width="65" bgcolor="#efefef" align="center"><%=rs("handset")%></td>
            <td bgcolor="#efefef" align="center" width="40"><a href="editcard1.asp?id=<%=rs("id")%>" onClick="return OpenWindows(this.href)"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image70','','images/detail_on.gif',1)"><img name="Image70" border="0" src="images/detail_off.gif" width="14" height="16" ></a></td>
          </tr>
           <%rs.movenext 
count=count+1
loop 
end if%>     
        </table>	
  <table width="100%" cellpadding="0" cellspacing="0" border="0">
    <tr>
          <form name="form1" action="txl.asp" method="post">     
      <td bgcolor="#BFBFBF" width="15%"> 查询：</td>  
            <td height=20 width="16%" bgcolor="#BFBFBF"> 
              <select name="style">
                  <option value="cname" selected>姓名</option>
                  <option value="company">单位</option>
                  
                </select>
            </td>
	  <td bgcolor="#BFBFBF" width="4%" >为</td>
       
      <td width="20%" bgcolor="#BFBFBF"> 
        <input size=30 name="key">
                <input value="0" name="flag" type="hidden">
              </td>
              
            
            <td bgcolor="#BFBFBF" width="45%"><INPUT TYPE="image" SRC="images/search_off.gif"> <a href="addcard.asp" onclick="return OpenWindows(this.href);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image4','','images/add_on_2.gif',1)"><img name="Image4" border="0" src="images/add_off_2.gif" height="19"></a>
            </td>
	          </form>
            </tr>
          </table>



	      </td>
    </tr>
  </table>
</body>
</html>
