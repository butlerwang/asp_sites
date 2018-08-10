<%
Session.Abandon()
Response.Write("<script>this.top.location.href='login.asp';</script>")
%>