<!--#include file="data.asp"-->
<!--#include file="char.asp"--> <%
    dim rs,sql
    dim title
    dim content
    dim articleid
    dim classid,Nclassid
    dim Nkey
    dim writer
    dim writefrom
    dim errmsg
    dim founderr
    founerr=false
    if trim(request.form("txttitle"))="" then
          founderr=true
          errmsg="<li>�ļ����ⲻ��Ϊ��</li>"
    end if
    if trim(request.form("typeid"))="" then
          founderr=true
          errmsg="<li>�ļ����Ͳ���Ϊ��</li>"
    end if
    if trim(request.form("txtcontent"))="" then
          founderr=true
          errmsg=errmsg+"<li>�ļ����ݲ���Ϊ��</li>"
    end if

    if founderr=false then
        title=htmlencode(request.form("txttitle"))
        typeid=request.form("typeid")

        if request("htmlable")="yes" then
        content=htmlencode2(request("txtcontent"))
        else
        content=ubbcode(request.form("txtcontent"))
        end if

    set rs=server.createobject("adodb.recordset")
    if request("action")="add" then
        call newsoft()
    elseif request("action")="edit" then
        call editsoft()
    else
        founderr=true
        errmsg=errmsg+"<li>û��ѡ������</li>"
    end if
sub newsoft()
    sql="select * from learn where (id is null)" 
    rs.open sql,conn,3,3
    rs.addnew
    rs("title")=title
    rs("content")=content
    rs("type")=typeid
    rs("time")=date()
    rs.update
    articleid=rs("id")
end sub
sub editsoft()
    sql="select * from learn where id="&request("id")
    rs.open sql,conn,1,3
    rs("title")=title
    rs("content")=content
    rs("type")=typeid
    rs.update
    articleid=rs("id")
end sub

    rs.close
    set rs=nothing
    conn.close
    set conn=nothing
%> <title></title> <link rel="stylesheet" type="text/css" href="style.css"> <div align="center"><center> 
<br><br> <table width="50%" border="1" cellpadding="0" cellspacing="0" bordercolor="#999999"> 
<tr> <td width="100%" bgcolor="#999999" height="20"><p align="center"><font color="#FFFFFF"><b> 
<%if request("action")="add" then%>���<%else%>�޸�<%end if%>�ļ��ɹ�</b></font></td></tr> 
<tr> <td width="100%"><p align="left"><br> ���±���Ϊ��<%response.write title%></p><A HREF="freeadd.asp">�������</A> 
&nbsp;&nbsp;&nbsp; <A HREF="javaScript:window.close()">�رմ���</A> </td></tr> </table></center></div><%
else
 response.write "�������µ�ԭ���ܳɹ�����ļ����ݣ�"
 response.write errmsg
 response.write "<BR><A HREF=javascript:history.back(1)>��������д</A>"
end if
%> 
</body>
</html>
