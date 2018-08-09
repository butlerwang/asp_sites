<!--#include file="../conn.asp"-->
<%
If Request.ServerVariables("HTTP_REFERER")="" then Response.End()
Dim flag:flag=Request("flag")   '当flag值为getid时，则返回ID号
%>
var subcity = new Array();
<%
set ors=Conn.Execute("select a.ParentID,a.City,b.City,a.id FROM KS_Province a inner join KS_Province b on b.id=a.parentid WHERE a.parentid<>0 order by a.orderid")
dim n:n=0
do while not ors.eof
 if flag="getid" then%>
subcity[<%=n%>] = new Array('<%=ors(0)%>','<%=trim(ors(1))%>','<%=trim(ors(3))%>')
<%else%>
subcity[<%=n%>] = new Array('<%=ors(2)%>','<%=trim(ors(1))%>')
<%end if
ors.movenext
n=n+1
loop
ors.close
set ors=nothing
%>
function changecity(selectValue)
{
 try{
 setCookie("pid",selectValue);
 }catch(e)
 {
 }
document.getElementById('City').length = 0;
document.getElementById('City').options[0] = new Option('请选择','');
for (i=0; i<subcity.length; i++)
{
<%If flag="getid" then%>
 if (subcity[i][0] == selectValue){document.getElementById('City').options[document.getElementById('City').length] = new Option(subcity[i][1], subcity[i][2]);}
<%else%>
 if (subcity[i][0] == selectValue){document.getElementById('City').options[document.getElementById('City').length] = new Option(subcity[i][1], subcity[i][1]);}
<%end if%>
}
}

<%
exec="select ID,City from KS_Province where parentid=0 order by orderid"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1
%>
document.write ("<select name='Province' id='Province' class='select' onChange='changecity(this.value)'>");
document.write ("<option value='' selected>选择省份...</option>");
<%
do while not rs.eof
 if flag="getid" then  
  response.write ("document.write (""<option value=" & rs(0) & ">" & rs(1) & "</option>"");")
 else
  response.write ("document.write (""<option value=" & rs(1) & ">" & rs(1) & "</option>"");")
 end if
rs.movenext
loop
rs.close
set rs=nothing
%>
document.write ("</select>")

document.write (" <select name='City' class='select' id='City'>");
document.write ("<option value='' selected>请选择..</option>");
document.write ("</select>")
<%
CloseConn
%>
try
{changecity(getCookie("pid"));
}catch(e)
{}
