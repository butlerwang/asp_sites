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
               if (confirm("�˲�����ɾ����ѡ����ļ�������ȷ��ɾ����"))  {                        
               var url="manage/articledel.asp?ownid=" ;                       
               form.action=url;                       
               form.submit();}                        
          }                        
         else  { alert("(û��ѡ���ļ�)���ڸ�ѡ����ѡ��Ҫɾ�����ļ�") }                        
 }                   
</script>



<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0"
style="background-image: url('images/main_bg.gif'); background-attachment: scroll; background-repeat: no-repeat; background-position: left bottom" onLoad="MM_preloadImages('images/more_on.gif','images/add_on.gif','images/up_on.gif')" >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td class="heading" bgcolor="#4e5960" height=20>��<font color="#FFFFFF"><b>�������</font></b></td>
    <td class="heading" bgcolor="#4e5960">
      <p align="right">
   </td>
  </tr>
  <tr> 
    <td width="110" valign="top"> <br>
      <table width="100%" border="0" cellspacing="0" cellpadding="2" align="center" >
        <tr > 
          <td>

		   
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
        <td bgcolor="#C0C0C0">    
          <%   

dim keyword
keyword=request("selecttext")
dim page
page=request("page")
PageSize = 6
dim rs,strSQL,news
strSQL ="SELECT * FROM soft where name like '%"&keyword&"%'  ORDER BY id DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSQL,Conn,1,1
    rs.PageSize = PageSize
	totalfilm=rs.recordcount
    pgnum=rs.Pagecount
    if page="" or clng(page)<1 then page=1
    if clng(page) > pgnum then page=pgnum
    if pgnum>0 then rs.AbsolutePage=page

if rs.eof and rs.bof then
response.write "<font color='#ffffff' class='3dfont'>�����û���κ��ļ�</font>"
else
%> <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr bgcolor="#303430"> <form method=Post action="file.asp"><FONT COLOR="#ffffff">
               [<b><%=rs.pagecount%></b>/<%=page%>ҳ] [��<%=totalfilm%>��] <%if page=1 then%> [��ҳ] [��һҳ] <% else %> [<a href="file.asp?page=1&typeid=<%=typeid%>">��ҳ</a>] 
               [<a href="file.asp?page=<%=page-1%>&typeid=<%=typeid%>">��һҳ</a>]<%end if%><%if rs.pagecount-page<1 then%> [��һҳ] [βҳ]  <%else%> [<a href="file.asp?page=<%=page+1%>&typeid=<%=typeid%>">��һҳ</a>]  [<a href="file.asp?page=<%=rs.pagecount%>&typeid=<%=typeid%>">βҳ</a>]</FONT> <%end if%> <input type='text' name='page' size=2 maxlength=10 style="font-size:9pt;color:#FFFFFF;background-color:#666666;border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000" value="<%=page%>" align=center> <INPUT TYPE="hidden" name=type value="<%=typeid%>"><input style="border:1 solid black;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; FONT-WEIGHT: normal; HEIGHT: 18px; LINE-HEIGHT: normal" type='submit'  value=' Goto '   size=2></td>
             </tr>
           </table>   
        </td>                                   
    </tr>                            
    <tr>
	    <td>
    <%
count=0 
do while not (rs.eof or rs.bof) and count<rs.PageSize 
%>  <table border=1 cellspacing="0" cellpadding="0" width="100%" bordercolorlight=#000000 bordercolordark=#ffffff>                                                       
    <tr style="padding:1pt;background-color:cccccc">                
                     
      <td><a href="<%=rs("url")%>"><%=rs("name")%></A> <%if rs("best")=true then response.write " <FONT COLOR=red>*</FONT>" end if%><%if session("Urule")="a" then%>[<A HREF="#" onclick="window.open('esoft.asp?id=<%=rs("id")%>','','width=400 height=250')"><FONT COLOR="#000000">��</FONT></A>] [<A HREF="delsoft.asp?id=<%=rs("id")%>" onclick="return delsoft()"><FONT COLOR="#000000">ɾ</FONT></A>] <%end if%>                                     
                ��</td>                                   
      <td width="70" align=center>                  
        <%=rs("time")%>                                    
        ��        
      </td>
	  <td width="30" align=center><%=rs("size")%>
	  </td>
       </tr>
	   <tr>
	     <td colspan=3 style="padding:5pt"><%=rs("content")%>
		 </td>
	   </tr>
	   </table>
        <%rs.movenext 
count=count+1
loop 
end if
%>
</td>
</tr></form>                        

</table>
          
 <form  name="sele"  method="post"  action="file.asp"  onsubmit="return  ckse()">              
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
                                 {   alert ("�������ѯ���ݣ�");
                                     sele.selecttext.focus();
                                     return false;
                                      }
                            }
                        </script>
              &nbsp;����ؼ��֣� </td>
            <td bgcolor="#BFBFBF" width="16%">                        
              <input type="text" name="selecttext" size="10" maxlength="12">                      
            </td>                      
            <td width="53%"><a href="javascript:document.sele.submit();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','images/search_on.gif',1)"><img name="Image8" border="0" src="images/search_off.gif" align="middle"></a> <%if session("Urule")="a" then%>   <a href="#" onclick="window.open('addsoft.asp','','width=400 height=250')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image9','','images/add_on.gif',1)"><img name="Image9" border="0" src="images/add_off.gif" align="middle"></a> <a href="#" onclick="window.open('uploadfile.asp','','width=450 height=200')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','images/up_on.gif',1)"><img name="Image1" border="0" src="images/up_off.gif" align="middle"></a>  <%end if%>                     
                                    
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