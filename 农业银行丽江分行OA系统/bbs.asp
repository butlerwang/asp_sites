<!--#include file="data.asp"-->
<!--#include file="check.asp"-->

<HTML><HEAD><title>bbs_center</title>

<SCRIPT>
function OpenWindows(url)
{
  var 
 newwin=window.open(url,"_blank","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=50,left=120,width=600,height=400");
 return false;
 
}
function OpenSmallWindows(url)
{
  var 
 newwin=window.open(url,"_blank","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=50,left=120,width=600,height=450");
 return false;
 
}
</SCRIPT>
<script language="JavaScript">
<!--
function MM_goToURL() { //v3.0
if (confirm("���Ƿ�Ҫɾ��������Ϣ����ȷ��ɾ����")==1) {
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2)  eval(args[i]+".location='"+args[i+1]+"'");
}
}
function cform(){
 if(!confirm("��ȷ��ɾ����֪ͨ��"))
 return false;

}//-->
</script>

<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="oa.css" rel=stylesheet>
<SCRIPT language=JavaScript>
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
</SCRIPT>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 
onload="MM_preloadImages('images/more_on.gif','images/newarticle_on.gif','images/manage_on.gif','images/delete_on.gif','images/search_on.gif')" 
style="BACKGROUND-ATTACHMENT: scroll; BACKGROUND-POSITION: left bottom; BACKGROUND-REPEAT: no-repeat" 
topMargin=0>
<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
  <TBODY>
  <TR>
    <TD bgColor=#4e5960 class=heading colSpan=2 height=3></TD></TR>
  <TR>
    <TD bgColor=#4e5960 class=heading height=20>��<FONT 
      color=#ffffff><B>��������</B></FONT></TD>
    <TD bgColor=#4e5960 class=heading height=20>
      &nbsp;</TD></TR>
  <TR>
    <TD colSpan=2 height=131 vAlign=top>
      <TABLE border=0 cellPadding=2 cellSpacing=1 width="100%">
        <TBODY>
        <TR bgColor=#999999>
          <TD class=heading colSpan=6>
            <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
              <TBODY>
              <TR>
                <TD align=right width="2%"><IMG height=18 name=Image3 
                  src="images/adorn.gif" width=10></TD>
                <TD align=left width="10%"> </TD>

                <TD align=left width="40%"><A 
                  href="addbbs.asp" 
                  onclick="return OpenSmallWindows(this.href);" 
                  onmouseout=MM_swapImgRestore() 
                  onmouseover="MM_swapImage('Image2','','images/newarticle_on.gif',1)"><IMG 
                  border=0 height=19 hspace=5 name=Image2 
                  src="images/newarticle_off.gif" width=85></A> </TD>
                <TD align=right>&nbsp;&nbsp;&nbsp;</TD>
                <TD width="3%"><IMG height=18 name=Image1 
                  src="images/adorn.gif" 
        width=10></TD></TR></TBODY></TABLE></TD></TR>
        <TR bgColor=#bfbfbf>
          <TD bgColor=#bfbfbf><B>����</B></TD>
          <TD width=110><B>����</B></TD>
          <TD align=middle><B>�����</B></TD>
          <TD align=middle><B>����ʱ��</B></TD>        
          <%if Session("Urule")<>"c" then%><TD width=45><B>ɾ��</B></TD><%end if%>
</TR>
        
<%'�����ҳ�ĺ�������totalnumber��maxperpage��filename��Ϊ��������ڡ�
function showpages(totalnumber,maxperpage,filename)
  dim n
  if totalnumber mod maxperpage=0 then
     n= totalnumber \ maxperpage
  else
     n= totalnumber \ maxperpage+1
  end if
  if CurrentPage<2 then
    response.write "<font color='999966'><img border='0' src='images/1-prev.gif' align='absmiddle'></font>&nbsp;"
  else
    response.write "<a href="&filename&"?page="&CurrentPage-1&"><img border='0' src='images/1-prev.gif' align='absmiddle'></a>&nbsp;"
  end if
  response.write "&nbsp;��ҳ�Σ�<strong><font color=red>"&CurrentPage&"</font>/"&n&"</strong>��&nbsp;"
  if n-currentpage<1 then
    response.write "<font color='999966'><img border='0' src='images/1-next.gif' align='absmiddle'></font>"
  else
    response.write "<a href="&filename&"?page="&(CurrentPage+1)
    response.write "><img border='0' src='images/1-next.gif' align='absmiddle'></a>"
  end if
end function

keyword=request("keyword")
if request("style")="title" then
strSql="select * from bbs where subject like '%"&keyword&"%' ORDER BY time desc, id DESC"
else if request ("style")="content" then
strSql="select * from bbs where content like '%"&keyword&"%' ORDER BY time desc, id DESC"
else if request ("style")="name" then
strSql="select * from bbs where name like '%"&keyword&"%' ORDER BY time desc, id DESC"
else 
strSql="select * from bbs ORDER BY time desc, id DESC"
end if
end if
end if
set my_rs=server.createobject("adodb.recordset")
my_rs.open strsql,conn,1,1
dim currentpage  '���嵱ǰҳ
dim filename     '�ļ���
Const MaxPerPage=10  'ÿҳ��ʾ�ļ�¼����
dim totalnumber  '��¼����
filename="bbs.asp"
if not isempty(request("page")) then
      currentPage=cint(request("page"))
   else
      currentPage=1
end if
if not my_rs.eof then
    totalnumber = my_rs.recordcount     '���ü�¼����
    my_rs.PageSize = MaxPerPage
    my_rs.AbsolutePage = currentpage   '��ָ���ƶ�����ǰҳ
    i=1
    do while not my_rs.eof and i<=MaxPerPage
%>



        <TR>
          <TD bgColor=#efefef><%if my_rs("SubjectId")="0" then%><IMG src="images/<%=my_rs("Pic")%>"><a href="readbbs.asp?SubjectId=<%=my_rs("ID")%>"  onClick='return OpenWindows(this.href);'><%=my_rs("Subject")%></A><%else%>&nbsp;&nbsp;<IMG src="images/<%=my_rs("Pic")%>"><a href="readbbs.asp?SubjectId=<%=my_rs("ID")%>"  onClick='return OpenWindows(this.href);'><%=my_rs("Subject")%></a><%end if%></TD>
          <TD bgColor=#efefef width=110><%=my_rs("name")%></TD>
          <TD align=middle bgColor=#efefef width=50><%=my_rs("Knock")%></TD>
          <TD align=middle bgColor=#efefef><%=my_rs("Time")%></TD>   
          <%if Session("Urule")<>"c" then%>
          <TD bgColor=#efefef><a href="delebbs.asp?id=<%=my_rs("id")%>&delbz=My_only" onclick="return cform();"><img border="0" src="images/dele_1.gif" alt="ɾ����"></A> 
          </TD><%end if%>
</TR>
<%my_rs.movenext
i=i+1
loop%>
<tr><td colspan="5">
    <p align="center"><%showpages totalnumber,MaxPerPage,filename '����ҳ����ʾ����%></td></tr>
<%else%>
<tr><td colspan="5">
    <p align="center">û���κ�����</td></tr><%end if
my_rs.close%>


      </TBODY></TABLE>
      <TABLE bgColor=#666666 border=0 cellPadding=2 cellSpacing=0 
        width="100%"><TBODY>
        <TR>
          <FORM action="bbs.asp" method=get name=form1>
          <TD bgColor=#bfbfbf class=show vAlign=center width=300><SELECT 
            name=style> <OPTION selected value=title>�������ѯ</OPTION> <OPTION 
              value=content>�����ݲ�ѯ</OPTION><OPTION value=name>�����߲�ѯ</OPTION></SELECT> <INPUT name=keyword size=11> 
             </TD>
          <TD bgColor=#bfbfbf class=show vAlign=center><A 
            href="javascript:document.form1.submit();" 
            onmouseout=MM_swapImgRestore() 
            onmouseover="MM_swapImage('Image8','','images/search_on.gif',1)"><IMG 
            border=0 height=19 name=Image8 src="images/search_off.gif" 
            width=60></A></TD></FORM></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<SCRIPT language=javascript>          
    function DelChk()          
    {          
        if(confirm("��ȷ��ɾ���ü�¼��"))          
            document.delform.submit();                    
    }          
</SCRIPT>
</BODY></HTML>
