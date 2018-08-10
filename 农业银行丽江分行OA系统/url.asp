<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->

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
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

<html><head><title>url</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="oa.css">
</head>
<script>
function js_openpage(url) {
  var 
   newwin=window.open(url,"NewWin","toolbar=no,resizable=yes,location=no,directories=no,status=no,menubar=no,scrollbars=yes,top=220,left=220,width=450,height=200");
 // newwin.focus();
  return false;
} 

function del(url) 
{  
  if (confirm("是否要删除此信息")) 
  {
     window.open(url,"NewWin","toolbar=no,resizable=yes,location=no,directories=no,status=no,menubar=no,scrollbars=yes,top=220,left=220,width=450,height=200");
  }
} 

</script>
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0"
 onLoad="MM_preloadImages('images/add_on.gif','images/modify_on.gif','images/delete_on.gif','images/add_on.gif','images/modi_2.gif','images/dele_2.gif','images/showall_on.gif')">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="heading" bgcolor="#4e5960" colspan="2" height="3"></td>
  </tr>
  <tr> 
    <td class="heading" bgcolor="#4e5960" colspan="2" >　<font color="#FFFFFF"><b>常用网址</b></font></td>
  </tr>
  <tr> 
    
    <form method="post" action=""   name="sele"  onsubmit="return ckse()">
      <td> 
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr bgcolor="#999999"> 
            <td colspan="4" class="heading"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="2%" align="right"><img src="images/adorn.gif" width="10" height="18"></td>
                  <td align="left">
                   
                     </td>
                  <td colspan="3" align="right" valign="middle"> 
                                   

 <%   

dim keyword
keyword=request("selecttext")

dim page
page=request("page")
PageSize = 17
dim rs,strSQL,news
strSQL ="SELECT * FROM url where 网站名称 like '%"&keyword&"%' ORDER BY id DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSQL,Conn,3,3
    rs.PageSize = PageSize
	totalfilm=rs.recordcount
    pgnum=rs.Pagecount
    if page="" or clng(page)<1 then page=1
    if clng(page) > pgnum then page=pgnum
    if pgnum>0 then rs.AbsolutePage=page

if rs.eof then
response.write "<font color='#ffffff' class='3dfont'>还没有任何网址</font>"
else
%> <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr bgcolor="#303430"> <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript"><!--
function FrontPage_Form2_Validator(theForm)
{

  if (theForm.selecttext.value == "")
  {
    alert("请在 selecttext 域中输入值。");
    theForm.selecttext.focus();
    return (false);
  }

  if (theForm.selecttext.value.length > 12)
  {
    alert("在 selecttext 域中，请最多输入 12 个字符。");
    theForm.selecttext.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method=Post action="url.asp" onsubmit="return FrontPage_Form2_Validator(this)" name="FrontPage_Form2"><FONT COLOR="#ffffff">
               [<b><%=rs.pagecount%></b>/<%=page%>页] [共<%=totalfilm%>个] <%if page=1 then%> [首 页] [上一页] <% else %> [<a href="url.asp?page=1">首 页</a>] 
               [<a href="url.asp?page=<%=page-1%>">上一页</a>]<%end if%><%if rs.pagecount-page<1 then%> [下一页] [尾 页]  <%else%> [<a href="url.asp?page=<%=page+1%>">下一页</a>]  [<a href="url.asp?page=<%=rs.pagecount%>">尾 页</a>]</FONT> <%end if%> <input type='text' name='page' size=2 maxlength=10 style="font-size:9pt;color:#FFFFFF;background-color:#666666;border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000" value="<%=page%>" align=center> <input style="border:1 solid black;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; FONT-WEIGHT: normal; HEIGHT: 18px; LINE-HEIGHT: normal" type='submit'  value=' Goto '   size=2></td>
             </tr>
           </table>
 </td>                 
                  <td width="3%"><img src="images/adorn.gif" width="10" height="18"></td>                 
                </tr>                 
              </table>                 
            </td>                 
          </tr>                 
                           
          <tr bgcolor="#bfbfbf">                  
            <td><b>网站名称</b></td>                 
            <td><b>网站地址</b></td> 
			<td><b>网站说明</b></td>                 

            <% if session("Urule")="a" then			
			%> 
			<td width="10%"><b>删 除</b></td>                 
            <%end if%>
          </tr>                 
          <%
count=0 
do while not (rs.eof or rs.bof) and count<rs.PageSize 
%>                 
          <tr>                  
            <td bgcolor="#efefef"><%=rs("网站名称")%></td>                 
            <td bgcolor="#efefef"><A HREF="http://<%=rs("网址")%>" target=_blank><%=rs("网址")%></A></td>                 
            <td bgcolor="#efefef"><%if len(rs("网站说明"))>15 then%><%=left(rs("网站说明"),15)%>……<%else%><%=rs("网站说明")%><%end if%></td> <script>  
			function cform(){
 if(!confirm("您确认删除该网址！"))
 return false;

}
</script>               
              <% if session("Urule")="a" then			
			%> 
			<td width="10%" bgcolor="#efefef"><A HREF="delurl.asp?id=<%=rs("id")%>" onclick="return cform();">删除</A></td>                 
            <%end if%>            
          </tr>                 
          <%rs.movenext 
count=count+1
loop 
end if%>            
                           
          <!--记录数为零时-->                 
                           
        </table>                 
                                
        <table width="100%" cellpadding="0" cellspacing="0" border="0">                 
          <tr>                 
            <td bgcolor="#BFBFBF" width="11%">                  
                               
              </a>                  
              <script language="Javascript">                                              function ckse()                                                                                                                   
                        {                                                                                                                            
                            if (sele.selecttext.value=="")                                                                                                                                            
                                 {   alert ("请输入查询内容！");                                                                                                              sele.selecttext.focus();                                                                                                                               
                                     return false;                                                                        }                                                                                                                                                   
                            }                                                                                                                                                                                                                      
                        </script>                 
              &nbsp;查询： </td>                 
			                   
            <td bgcolor="#BFBFBF" width="16%">&nbsp;                  
              <select size="1" name="seler">                 
                <option value="phonename">网站名称</option>
              </select>
            </td>
			  
            <td bgcolor="#BFBFBF" width="4%"> 为 </td>
			  
            <td bgcolor="#BFBFBF" width="16%"> 
              <!--webbot bot="Validation" B-Value-Required="TRUE"
              I-Maximum-Length="12" -->
              <input type="text" name="selecttext" size="10" maxlength="12">
            </td>
            <td bgcolor="#BFBFBF" width="53%"><a href="javascript:document.sele.submit();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','images/search_on.gif',1)"><img name="Image8" border="0" src="images/search_off.gif" align="middle"></a> 
            <%
			if Session("Urule")<>"c" then
			%>  
			<A HREF="#" onClick="MM_openBrWindow('addurl.asp','','width=400,height=250')"><img src="images/add_off.gif" align="middle" border=0></a>
			<%
			end if
			%>
            </td>
          </tr>
        </table>                                                                                        
    </td></form>                                                                                      
  </tr>                                                                                      
</table>                                                                
                  
                                   
             
</body>                       
