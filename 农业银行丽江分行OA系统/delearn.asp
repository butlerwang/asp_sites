<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<%
if session("Urule")<>"a" then
response.write "您没有足够的权限查看此页：P"
response.end
end if
Set rs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from type"
rs.open strSql,Conn,1,1 
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
function cform(){
 if(!confirm("您确认删除？在删除此类别前，请先将该类别文件删除或更改类别！"))
 return false;

}
</script>



<html><head><title>scroll menu</title>
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
              　<img src="images/open.gif" align="absmiddle"> 
			  <a href="delearn.asp" class="t1">栏目管理
               </font></a><br> 
              　<img src="images/open.gif" align="absmiddle"> 
			  <a href="elearn.asp" class="t1">删除修改</font></a><br> 
          </td> 
     </tr> 
      </table> 
    </td> 
      <td valign="top" >  
      <table width="100%" border="0" cellspacing="0" cellpadding="0"> 
        <tr >  
          <td align="right" width="100%">  
 
<TABLE width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor=#636563>
      <tr>                
      <td bgcolor="#F6F6F6" align=center>所 有 类 别 </td>                                   
      <td bgcolor="#F6F6F6" align=center colspan=2>管 理</td>                               
      </tr>
	                  
    
    <%if rs.eof and rs.bof then
response.write "<font color='red'>还没有任何类别</font>"
else

do while not (rs.eof or rs.bof)
%>     
<FORM METHOD=POST ACTION="editlearn.asp">
<tr>                                                    
      <td bgcolor="#F6F6F6" width="80%"><INPUT TYPE="text" value="<%=rs("type")%>" name=type style="border:1pt solid #636563;font-size:9pt"><INPUT TYPE="hidden" name=id value=<%=rs("id")%>>
      </td>                                   
      <td bgcolor="#F6F6F6" align=center><INPUT TYPE="submit" name="edit" value="修改" style="border:1pt solid #636563;font-size:9pt; LINE-HEIGHT: normal;HEIGHT: 18px;">
      </td>                               
      <td bgcolor="#F6F6F6" align="center"><INPUT TYPE="submit" value="删除" name="del" style="border:1pt solid #636563;font-size:9pt; LINE-HEIGHT: normal;HEIGHT: 18px;" onclick="return cform();">                  
      </td>                               
       </tr>
       </FORM>
 
		<%rs.movenext 
loop 
end if%>               

<FORM METHOD=POST ACTION="editlearn.asp">

<tr>
  <td colspan=3 bgcolor="#F6F6F6">新增加类别:<INPUT TYPE="text" NAME="type" style="border:1pt solid #636563;font-size:9pt">&nbsp;&nbsp;&nbsp;<INPUT TYPE="submit" name="add" value="增加" style="border:1pt solid #636563;font-size:9pt; LINE-HEIGHT: normal;HEIGHT: 18px;">
  </td>
</tr>
</TABLE>
</FORM>


         </td>                                          
       </tr>                                         
      </table> </td> </tr>             
      </table>                                        
       </td>                                        
     </tr>                                        
 </table>                                        
                               
</body>                                        
</html>                            
                            
                            
                            
                       
                       
                       
                       
                       
        
        
