<!--#INCLUDE FILE="data.asp" -->
<%
if session("Urule")<>"a" then
response.redirect "error.asp?id=admin"
end if
%>
<!--#INCLUDE FILE="check.asp" -->

<script language="JavaScript">
<!--
function  validate()
    {
       
        if  (document.myform.type.value=="")
        {
            alert("新加单位不能为空");
            document.myform.type.focus();
            return false ;
        }
		}
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
 if(!confirm("你是否确认删除该单位？"))
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
    <td class="heading" bgcolor="#4e5960" height=20>　<font color="#FFFFFF"><b>单位管理</font></b></td>
    <td class="heading" bgcolor="#4e5960">
      <p align="right">
       
      </td>
  </tr>
  <tr> 
    <td width="110" valign="top"> &nbsp;
    </td> 
      <td valign="top" >  
      <table width="100%" border="0" cellspacing="0" cellpadding="0"> 
        <tr >  
          <td align="right" width="100%"><%   

dim keyword
keyword=request("key")

dim page
page=request("page")
PageSize = 14
dim rs,strSQL,news
strSQL ="select * from bumen ORDER BY id DESC"
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
%> <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr> <form method=Post action="mm.asp">
               [<b><%=rs.pagecount%></b>/<%=page%>页] [共<%=totalfilm%>个] <%if page=1 then%> [首 页] [上一页] <% else %> [<a href="mm.asp?page=1">首 页</a>] 
               [<a href="mm.asp?page=<%=page-1%>">上一页</a>]<%end if%><%if rs.pagecount-page<1 then%> [下一页] [尾 页]  <%else%> [<a href="mm.asp?page=<%=page+1%>">下一页</a>]  [<a href="mm.asp?page=<%=rs.pagecount%>">尾 页</a>] <%end if%> <input type='text' name='page' size=2 maxlength=10 style="font-size:9pt;color:#FFFFFF;background-color:#666666;border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000" value="<%=page%>" align=center> <input style="border:1 solid black;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; FONT-WEIGHT: normal; HEIGHT: 18px; LINE-HEIGHT: normal" type='submit'  value=' Goto '   size=2></td>
             </tr>
           </table>  
 
<TABLE width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor=#636563>
      <tr>                
      <td bgcolor="#F6F6F6" align=center>单 位 名 称 </td>                                   
      <td bgcolor="#F6F6F6" align=center colspan=2>管 理</td>                               
      </tr>
	          </form>        
    
    <%
	count=0 
do while not (rs.eof or rs.bof) and count<rs.PageSize 

%>     
<FORM METHOD=POST ACTION="emm.asp">
<tr>                                                    
      <td bgcolor="#F6F6F6" width="80%"><INPUT TYPE="text" value="<%=rs("type")%>" name=type style="border:1pt solid #636563;font-size:9pt" size=30><INPUT TYPE="hidden" name=id value=<%=rs("id")%>>
      </td>                                   
      <td bgcolor="#F6F6F6" align=center><INPUT TYPE="submit" name="edit" value="修改" style="border:1pt solid #636563;font-size:9pt; LINE-HEIGHT: normal;HEIGHT: 18px;">
      </td>                               
      <td bgcolor="#F6F6F6" align="center"><INPUT TYPE="submit" value="删除" name="del" style="border:1pt solid #636563;font-size:9pt; LINE-HEIGHT: normal;HEIGHT: 18px;" onclick="return cform();">                  
      </td>                               
       </tr>
       </FORM>
 
		<%
		rs.movenext 
count=count+1
loop 
end if%>               

<FORM METHOD=POST ACTION="emm.asp" name="myform">

<tr>
  <td colspan=3 bgcolor="#F6F6F6">新增加部门:<INPUT TYPE="text" size=30 NAME="type" style="border:1pt solid #636563;font-size:9pt">&nbsp;&nbsp;&nbsp;<INPUT TYPE="submit" name="add" value="增加" style="border:1pt solid #636563;font-size:9pt; LINE-HEIGHT: normal;HEIGHT: 18px;" onclick="return  validate()">
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
                            
                            
                            
                       
                       
                       
                       
                       
        
        
