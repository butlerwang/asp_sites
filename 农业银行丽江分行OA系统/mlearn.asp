<%
if session("Urule")="c" then
response.redirect "error.asp?id=admin"
end if
%>
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
<html><head><title>manager_learn_art</title>
<link rel="stylesheet" href="oa.css">
</head>
<script>
function js_openpage(url) 
{
  var 
  newwin=window.open(url,"NewWin","toolbar=no,resizable=yes,location=no,directories=no,status=no,menubar=no,scrollbars=yes,top=0,left=100,width=550,height=460");
  }
  
</script>



<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0"
style="background-image: url('images/main_bg.gif'); background-attachment: scroll; background-repeat: no-repeat; background-position: left bottom">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="heading" bgcolor="#4e5960" height=20>　<font color="#FFFFFF"><b>学习文件管理</font></b></td>
    <td class="heading" bgcolor="#4e5960">
      <p align="right">
       
      </td>
  </tr>
  <tr> 
    <td width="110" valign="top"> <br>
      <table width="100%" border="0" cellspacing="0" cellpadding="2" align="center" >
        <tr > 
          <td>
          　<img src="images/open.gif" align="absmiddle"> 
		  <a href="javascript:js_openpage('freeadd.asp')" class="t1">添加文件</a><br> 
              <%if session("Urule")="a" then%>  
			  &nbsp;&nbsp;<img src="images/open.gif" align="absmiddle"> 
			  <a href="delearn.asp" class="t1">栏目管理
               </font></a><br> 
			  　<img src="images/open.gif" align="absmiddle"> 
			  <a href="elearn.asp" class="t1">删除修改</font></a><br> 
			  <%end if%>

          </td> 
     </tr> 
      </table> 
    </td> 
      <td valign="top" >  
      <table width="100%" border="0" cellspacing="0" cellpadding="0"> 
        <tr >  
          <td align="right" width="100%">  
 
<TABLE width="100%" border="0" cellspacing="0" cellpadding="0">
<TR>
	<TD><BR>
	 <P align=center>管理页面管理员可进行操作说明：

<P>1，通过Web添加文件。<BR>操作用户：普通管理员 

<P>2，对已经添加文件修改或删除，请点左边相关连接进行操作。<BR>操作用户：超级用户，普通管理员 

<P>3，对栏目进行添加修改删除，请点左边相关连接进行操作。<BR>操作用户：超级用户 

	</TD>
</TR>
</TABLE>


         </td>                                          
       </tr>                                         
      </table> </td> </tr>             
      </table>                                        
       </td>                                        
     </tr>                                        
 </table>                                        
                               
</body>                                        
</html>                            
                            
                            
                            
                       
                       
                       
                       
                       
        
        
