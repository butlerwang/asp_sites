<!--#include file="data.asp"-->
<!--#include file="check.asp"-->
<html><head><title></title>
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

function MM_findObj(n, d) { //v3.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<script Language="javaScript">
    function  validate()
    {
        if  (document.myform.name.value=="")
        {
            alert("��������Ϊ��");
            document.myform.name.focus();
            return false ;
        }
        if  (document.myform.Userid.value=="")
        {
            alert("��¼�ʺŲ���Ϊ��");
            document.myform.Userid.focus();
            return false ;
        }
		if  (document.myform.company.value=="")
        {
            alert("�������Ʋ���Ϊ��");
            document.myform.company.focus();
            return false ;
        }
		if  (document.myform.tel.value=="")
        {
            alert("�绰���벻��Ϊ��");
            document.myform.tel.focus();
            return false ;
        }
		if  (document.myform.email.value=="")
        {
            alert("�����ʼ�����Ϊ��");
            document.myform.email.focus();
            return false ;
        }
        if  (document.myform.password.value=="")
        {
            alert("���벻��Ϊ��");
            document.myform.password.focus();
            return false ;
        }
        return  true;
    }
</script>

</head>
<script>
function js_openpage(url) {
  var 
newwin=window.open(url,"NewWin","toolbar=no,resizable=yes,location=no,directories=no,status=no,menubar=no,scrollbars=yes,top=220,left=220,width=500,height=310");
 // newwin.focus();
  return false;
}</script>



<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" style="background-image: url('images/main_bg.gif'); background-attachment: scroll; background-repeat: no-repeat; background-position: left bottom">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="heading" bgcolor="#4e5960" colspan="2" height="3"></td>
  </tr>
  <tr> 
    <td class="heading" bgcolor="#4e5960" colspan="2">��<font color="#FFFFFF"><b>�޸�����</b></font></td>
  </tr>
  <tr> 
    <td width="109" valign="top">&nbsp;</td>
	<td valign="top"> 
	  <%
Set myrs= Server.CreateObject("ADODB.Recordset") 
strSql="select * from bumen"
myrs.open strSql,Conn,1,1 
dim sql
dim rs
 sql="select * from user where id="&session("uid")
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
                %>
	  <form method="post" action="savepd.asp?id=<%=request("id")%>" name="myform" onsubmit="return  validate()">
  
  		<table border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF" bordercolor="#FFFFFF" bordercolorlight="#000000" style="font-size:9pt" width="449">
		  <tr> 
			<td align=center colspan=2 bgcolor=#000000><font COLOR="#ffffff"><b>�޸ĸ�������</b></font> 
			</td>
		  </tr>
		  <tr> 
			<td width="24%" valign="top"> 
			  <p align="right">�������:</p>
			</td>
			<td width="76%"> 
			  <input type="text" name="name" class="form" value="<%=rs("����")%>" size="24">
			</td>
		  </tr>
		  <tr> 
			<td width="24%" valign="top" height="6"> 
			  <p align="right">��¼�ʺ�:</p>
			</td>
			<td width="76%" height="6"> 
			  <input type="hidden" name="Userid"  value="<%=rs("�û���")%>"  >
			  <input type="text" name="Userid2" class="form" value="<%=rs("�û���")%>" size="24" disabled>
			</td>
		  </tr>
		  <tr> 
			<td width="24%"  valign="top" height="16"> 
			  <p align="right">��¼����:</p>
			</td>
			<td width="76%" height="16"> 
			  <input type="password" name="password" class="form" size="24" value="<%=rs("����")%>">
			</td>
		  </tr>
		  <tr> 
			<td width="24%"  valign="top" height="16"> 
			  <p align="right">��������:</p>
			</td>
			<td width="76%" height="16"> 
			  <input type="text" name="question" class="form" size="24" value="<%=rs("����")%>">
			</td>
		  </tr>
		  <tr> 
			<td width="24%"  valign="top" height="16"> 
			  <p align="right">�����:</p>
			</td>
			<td width="76%" height="16"> 
			  <input type="text" name="answer" class="form" size="24" value="<%=rs("��")%>">
			</td>
		  </tr>
		  <tr> 
			<td width="24%"  valign="top"> 
			  <p align="right">��������: 
			</td>
			<td width="76%"> 
			  <select NAME="company">
				<%if myrs.eof and myrs.bof then
response.write "<font color='red'>��û���κ�����</font>"
else

do while not (myrs.eof or myrs.bof)
if myrs("type")=rs("����") then
sel="selected"
else 
sel=""
end if
%>
				<option value="<%=myrs("type")%>" <%=sel%>><%=myrs("type")%></option>
				<%myrs.movenext 
loop 
end if%>
			  </select>
			</td>
		  </tr>
		  <tr> 
			<td width="24%"  valign="top"> 
			  <p align="right">�绰����:</p>
			</td>
			<td width="76%"> 
			  <input type="text" name="tel" class="form" value="<%=rs("�绰")%>" size="24">
			</td>
		  </tr>
		  <tr> 
			<td width="24%"  valign="top"> 
			  <p align="right">�����ʼ�:</p>
			</td>
			<td width="76%"> 
			  <input type="text" name="email" class="form" value="<%=rs("����")%>" size="24">
			</td>
		  </tr>
		  <tr> 
			<td width="24%"  valign="top"> 
			  <p align="right">�ֻ�����:</p>
			</td>
			<td width="76%"> 
			  <input type="text" name="mobile" class="form" value="<%=rs("mobile")%>" size="24">
			</td>
		  </tr>
		  <tr> 
			<td align=center height="28"> 
			  <div align="right">ÿҳ��ʾ�ʼ���:</div>
			</td>
			<td align=left height="28"> 
			  <% 	for i=1 to 4
			Response.Write("<input type=radio name=iPageSize value=" & i*5)
			if cint(Session("iPageSize"))=i*5 then Response.Write(" checked ")
			Response.Write( ">" & i*5 &"��&nbsp;&nbsp;&nbsp;&nbsp;") 
		next
	%>
			  <div align="left"></div>
			</td>
		  </tr>
		  <tr> 
			<td align=center> 
			  <div align="right">�ʼ�ǩ����:</div>
			</td>
			<td align=left> 
			  <div align="left"> 
				<textarea name="iAdd" cols="50" rows="7" class="css0"><%=Session("iAdd")%></textarea>
			  </div>
			</td>
		  </tr>
		  <tr> 
			<td align=center colspan=2> 
			  <input type=image  src="images/modify_off.gif">
			</td>
		  </tr>
		</table>
		</form> 
</td>
  </tr>
</table>
</body>

</html>
