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

<html><head><title>learn_art</title>
<link rel="stylesheet" href="oa.css">
</head>
<script>
function js_openpage(url) 
{
  var 
  newwin=window.open(url,"NewWin","toolbar=no,resizable=yes,location=no,directories=no,status=no,menubar=no,scrollbars=yes,top=100,left=120,width=600,height=360");
  // newwin.focus();
  return false;
  }
  
function js_openpage_1(url) 
{
  var 
  newwin=window.open(url,"NewWin","toolbar=no,resizable=yes,location=no,directories=no,status=no,menubar=no,scrollbars=yes,top=100,left=120,width=600,height=400");
  // newwin.focus();
  return false;
}

function  DelChk()                       
           {    flag=0;                         
                for (j=0;j<form.elements.length;j++)   {                        
                   if (form.elements[j].checked==true){                        
                   flag=flag+1;                        
                   break;                        
                }                        
           }                        
           if (flag !=0){                         
               if (confirm("此操作将删除所选择的文件，请您确定删除！"))  {                        
               var url="manage/articledel.asp?ownid=" ;                       
               form.action=url;                       
               form.submit();}                        
          }                        
         else  { alert("(没有选择文件)请在复选框内选择要删除的文件") }                        
 }                   
</script>



<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0"
style="background-image: url('images/main_bg.gif'); background-attachment: scroll; background-repeat: no-repeat; background-position: left bottom" onLoad="MM_preloadImages('images/more_on.gif')" >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="heading" bgcolor="#4e5960" height=20>　<font color="#FFFFFF"><b> 帮 助 文  件 </font></b></td>
    <td class="heading" bgcolor="#4e5960">
      <p align="right">
   </td>
  </tr>
  <tr> 
    <td width="110" valign="top"> <br>
      <table width="100%" border="0" cellspacing="0" cellpadding="2" align="center" >
        <tr > 
          <td>

		  <%
Set my_rs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from helptype"
my_rs.open strSql,Conn,1,1 
if my_rs.eof then
response.write "<font class='3dfont'>还没有任何类别</font>"
else
do while not (my_rs.eof or my_rs.bof)
%>

          　<img src="images/open.gif" align="absmiddle"> 
	<%if my_rs("id")=request("typeid") then%>
	<a href="?typeid=<%=my_rs("id")%>">
         <FONT COLOR="red"><B><%=my_rs("type")%></B></FONT></a>
	<%else%>	  <a href="?typeid=<%=my_rs("id")%>">
          <%=my_rs("type")%></a><br> 
     <%end if%>        
      <%my_rs.movenext 
loop 
end if
my_rs.close
set my_rs=nothing%>       
          </td> 
     </tr> 
      </table> 
    </td> 
      <td valign="top" >  
      <table width="100%" border="0" cellspacing="0" cellpadding="0"> 
        <tr >  
          <td align="right" width="100%">  
 
 
<table width="100%" border=0 align="center" cellspacing="1" > 
     <tr >    
        <td bgcolor="#C0C0C0" colspan="4">    
          <%   

dim keyword
keyword=request("selecttext")
typeid=request("typeid")
dim page
page=request("page")
PageSize = 12
dim rs,strSQL,news
if typeid="" then
strSQL ="SELECT * FROM help where title like '%"&keyword&"%'  ORDER BY id DESC"
else
strSQL ="SELECT * FROM help where title like '%"&keyword&"%' and type='"&typeid&"' ORDER BY id DESC"
end if
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSQL,Conn,1,1
    rs.PageSize = PageSize
	totalfilm=rs.recordcount
    pgnum=rs.Pagecount
    if page="" or clng(page)<1 then page=1
    if clng(page) > pgnum then page=pgnum
    if pgnum>0 then rs.AbsolutePage=page

if rs.eof and rs.bof then
response.write "<font color='#ffffff' class='3dfont'>该类别还没有任何文件</font>"
else
%> <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr bgcolor="#303430"> <form method=Post action=""><FONT COLOR="#ffffff">
               [<b><%=rs.pagecount%></b>/<%=page%>页] [共<%=totalfilm%>个] <%if page=1 then%> [首页] [上一页] <% else %> [<a href="?page=1&typeid=<%=typeid%>">首页</a>] 
               [<a href="?page=<%=page-1%>&typeid=<%=typeid%>">上一页</a>]<%end if%><%if rs.pagecount-page<1 then%> [下一页] [尾页]  <%else%> [<a href="?page=<%=page+1%>&typeid=<%=typeid%>">下一页</a>]  [<a href="?page=<%=rs.pagecount%>&typeid=<%=typeid%>">尾页</a>]</FONT> <%end if%> <input type='text' name='page' size=2 maxlength=10 style="font-size:9pt;color:#FFFFFF;background-color:#666666;border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000" value="<%=page%>" align=center> <INPUT TYPE="hidden" name=type value="<%=typeid%>"><input style="border:1 solid black;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; FONT-WEIGHT: normal; HEIGHT: 18px; LINE-HEIGHT: normal" type='submit'  value=' Goto '   size=2></td>
             </tr>
           </table>   
        </td>                                   
    </tr>                            
       
    <%
count=0 
do while not (rs.eof or rs.bof) and count<rs.PageSize 
%>                                                         
    <tr>                
                     
      <td bgcolor="#F6F6F6" width="60%" colspan="2">  <a href="show_help.asp?id=<%=rs("id")%>" onClick="return js_openpage(this.href);"><%=rs("title")%>                                                                                      </A>                                      
                　</td>                                   
      <td bgcolor="#F6F6F6" width="20%">                  
        <p align="center"><%=rs("time")%>                                    
        　</p>              
      </td> 
	  <td bgcolor="#F6F6F6" width="20%">                  
        <p align="center">
		<%Set mrs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from helptype where id="&rs("type")
mrs.open strSql,Conn,1,1
if mrs.eof then
response.write "<FONT COLOR='red'>已删除类</FONT>"
else
response.write mrs("type")
end if
mrs.close
set mrs=nothing%>                                    
        　</p>              
      </td>
       </tr>                               
        <%rs.movenext 
count=count+1
loop 
end if%>               
</form>                        
          
 <form  name="sele"  method="post"  action=""  onsubmit="return  ckse()">              
  <tr >                            
    <td bgcolor="#E4E4E4" colspan="4">                       
                
       <table width="100%" cellpadding="0" cellspacing="0" border="0">                        
          <tr>                        
            <td bgcolor="#BFBFBF" width="20%">                         

              </a>
              <script language="Javascript">
                        function ckse()
                        {
                            if (sele.selecttext.value=="")
                                 {   alert ("请输入查询内容！");
                                     sele.selecttext.focus();
                                     return false;
                                      }
                            }
                        </script>
              &nbsp;输入关键字： </td>
            <td bgcolor="#BFBFBF" width="16%">                        
              <input type="text" name="selecttext" size="10" maxlength="12">                      
            </td>                      
            <td bgcolor="#BFBFBF" width="53%"><a href="javascript:document.sele.submit();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','images/search_on.gif',1)"><img name="Image8" border="0" src="images/search_off.gif" align="middle"></a>   <%if session("Urule")="a" then%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                    
              <FONT COLOR="red"><A HREF="mhelp.asp">管理入口</A></FONT> 
			  <%end if%>
             </td>                      
            </tr>                      
           </table>               
         </td>                                          
       </tr>                                         
      </form>                                                                          
      </table> </td> </tr>             
      </table>                                        
       </td>                                        
     </tr>                                        
 </table>                                   
</body>                                        
</html>                            