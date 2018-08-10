<!--#INCLUDE FILE="data.asp" -->
<!--#INCLUDE FILE="check.asp" -->
<%
set rs=server.createobject("ADODB.recordset") 
rs.Open "DELETE * FROM card Where ID="&request("id"),conn,1,3 
rs.update 
rs.close
set rs=nothing
conn.Close
set conn = nothing

%>
<script>
opener.location=opener.location;window.close()
</script>

