<!--#INCLUDE FILE="data.asp" -->
<%
if Session("Urule")<>"a" then
response.redirect "error.asp?id=admin"
end if
%><!--#INCLUDE FILE="check.asp" -->

<HTML><HEAD><TITLE>uesercheck</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="oa.css" rel=stylesheet>
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
function cform(){
 if(!confirm("��ȷ��ɾ�����û���"))
 return false;

}
function pass(){
 if(!confirm("��ȷ�ϸ��û�ͨ����ˣ�"))
 return false;

}

</script>
</script> <script language="JavaScript">
<!--
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}
//-->
</script> <script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
//-->
</script
></HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0>
<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%"> <TBODY> <TR> <TD bgColor=#4e5960 class=heading colSpan=2 height=3></TD></TR> 
<TR> <TD bgColor=#4e5960 class=heading>��<FONT 
    color=#ffffff><B>ע���û����� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="#" onClick="MM_openBrWindow('register.asp','','width=510,height=430')">������û�</a></B></FONT></TD></TR> <TR> <FORM action="userchk.asp" method=post name=sele><TD vAlign=top><TABLE border=0 cellPadding=2 cellSpacing=1 width="100%"> 
<TBODY> <TR bgColor=#999999> <TD class=heading colspan=9><TABLE border=0 cellPadding=0 cellSpacing=0 width="100%"> 
<TBODY> <TR> <TD align=right width="2%"><IMG height=18 
                  src="images/adorn.gif" width=10></TD><TD align=right><%   


dim page
page=request("page")
PageSize = 15
dim rs,strSQL,news
strSQL ="select * from user ORDER BY id DESC"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSQL,Conn,3,3
    rs.PageSize = PageSize
	totalfilm=rs.recordcount
    pgnum=rs.Pagecount
    if page="" or clng(page)<1 then page=1
    if clng(page) > pgnum then page=pgnum
    if pgnum>0 then rs.AbsolutePage=page

if rs.eof then
response.write "<font color='#ffffff' class='3dfont'>��û���κζ���</font>"
else
%> <table border="0" cellspacing="0" cellpadding="0" width="100%"> <tr> <form method=Post action="userchk.asp"><FONT COLOR="#ffffff"> 
[<b><%=rs.pagecount%></b>/<%=page%>ҳ] [��<%=totalfilm%>��] <%if page=1 then%> [�� 
ҳ] [��һҳ] <% else %> [<a href="userchk.asp?page=1">�� ҳ</a>] [<a href="userchk.asp?page=<%=page-1%>">��һҳ</a>]<%end if%><%if rs.pagecount-page<1 then%> 
[��һҳ] [β ҳ] <%else%> [<a href="userchk.asp?page=<%=page+1%>">��һҳ</a>] [<a href="userchk.asp?page=<%=rs.pagecount%>">β 
ҳ</a>]</FONT> <%end if%> <input type='text' name='page' size=2 maxlength=10 style="font-size:9pt;color:#FFFFFF;background-color:#666666;border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000" value="<%=page%>" align=center> 
<input style="border:1 solid black;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-VARIANT: normal; FONT-WEIGHT: normal; HEIGHT: 18px; LINE-HEIGHT: normal" type='submit'  value=' Goto '   size=2></td> </tr>
</table></TD> <TD width="3%"><IMG height=18 
                  src="images/adorn.gif" 
            width=10></TD></TR> </TBODY> </TABLE></TD> </TR> <TR bgColor=#bfbfbf align=center> 
<TD><b>����</b></TD><TD><b>�û���</b></TD><TD><b>ע��ʱ��</b></TD><TD><b>������λ</b></TD><TD><b>����</b></TD><TD><b>���</b></TD><TD><B>�޸�</B></TD><TD><b>ɾ��</b></TD><TD><b>Ȩ��</b></TD></TR> 
<%
count=0 
do while not (rs.eof or rs.bof) and count<rs.PageSize 
%> <TR> <td height="23" bgColor=#efefef> <p align="center"><%=rs("����")%> </td><td bgColor=#efefef> 
<p align="center"><%=rs("�û���")%> </td><td bgColor=#efefef> <p align="center"><%=rs("ʱ��")%> 
</td><td bgColor=#efefef> <p align="center"><%=rs("����")%> </td><td bgcolor=#efefef> 
<p align="center"><%=rs("ilevel")%></td><td align="center" bgColor=#efefef><% if rs("���")=true then %>�����<%else%><a href="shenghe.asp?id=<%=rs("id")%>" onclick="return pass();">�����</a><%end if%></td><td bgColor=#efefef> 
<p align="center"><a href="#" onClick="MM_openBrWindow('edit.asp?id=<%=rs("id")%>','','width=500,height=300')">�޸� 
</a> </td><td  bgColor=#efefef> <p align="center"> <a href="dele.asp?id=<%=rs("id")%>" onclick="return cform();">ɾ��</a> 
</td><td bgColor=#efefef><%
						  if rs("Ȩ��")="a" then 
						  response.write "�����û�"  
						  else if rs("Ȩ��")="b" then 
						  response.write "����Ա" 
						  else 
						  response.write "��ͨ�û�" 
						  end if
						  end if%></td></TR> <%rs.movenext 
count=count+1
loop 
end if%> <tr> <td colspan=9 bgcolor=#9c9a9c style="color:red">��ע:����ͨ���޸��û������������û�Ȩ�޺����伶�� 
</td></tr> </TBODY> </TABLE></TBODY></TABLE>
</BODY></HTML>
