<!--#include file="data.asp"-->
<%
if session("Urule")<>"a" then
response.redirect "error.asp?id=admin"
end if
    set rs=server.createobject("adodb.recordset")
    if request("del")<>"" then
        call deladrot()
    else
        call editadrot()
    end if

sub editadrot()
    sql="select * from adrot where id="&request("id")
    rs.open sql,conn,1,3
    rs("type")=request("type")
    rs("alt")=request("alt")
    rs("src")=request("src")
    rs("width")=request("width")
    rs("height")=request("height")
    rs("url")=request("url")
    rs.update
end sub
sub deladrot()
    sql="DELETE * from adrot where id="&request("id")
    rs.open sql,conn,1,3
    rs.update
end sub

    rs.close
    set rs=nothing
    conn.close
    set conn=nothing

    Response.Redirect("adrot.asp")

%>
