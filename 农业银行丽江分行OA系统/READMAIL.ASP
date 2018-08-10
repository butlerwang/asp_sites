<% if Session("id")="" then Response.Redirect("../error3.html") %>
<!--#INCLUDE FILE="check.asp" -->

<%

select case Session("send")
case "sendout" 
   ActFrom="发给"
   ActTime="发出时间"
   ActFromFieldName="ito"
   ActTitle="阅读邮件——发件箱"
   ActImg="sendout.gif"
case "del"
   ActFrom="发给/来自"
   ActTime="时  间"
   ActFromFieldName="ifrom"
   ActTitle="阅读邮件——回收站"
   ActImg="del.gif"
case "common"
   ActTime="发布时间"
   ActFrom="发布人"
   ActFromFieldName="ifrom"
   ActTitle="阅读邮件——公共信件"
   ActImg="common.gif"
case "recived"

   ActTime="收到时间"
   ActFrom="来自"
   ActFromFieldName="ifrom"
   ActTitle="阅读邮件——收件箱"
   ActImg="recived.gif"
end select
%>
<!--  #Include File=conn.asp -->


<head>


<link name="stsheet" id="stsheet" rel="stylesheet" href="images/vblife.css">
</head>
<script Language="vbScript">
if screen.width=1024 then document.stylesheets("stsheet").href="images/vblife1024.css"
</script>
<script Language="JavaScript">
<!-- on mouse double click, scrolldown


var currentpos, timer;
function initialize(){
 timer = setInterval("scrollwindow()", 1);
}
function sc(){
 clearInterval(timer);
}
function scrollwindow(){
 currentpos = document.body.scrollTop;
window.scroll(0, ++currentpos);

 if(currentpos != document.body.scrollTop){
  sc();
 }
}
document.onmousedown = sc;

document.ondblclick = initialize;

-->
</script>
<body leftmargin="0" topmargin="0" class=css1>

<div id="Layer88" style="position:absolute; width:100%; height:33px; z-index:0; left: 0; top: 0"><img src="IMAGES/line.jpg" width="100%" height="100%"></div>
<div id="Layer89" style="position:absolute; width:100%; height:33px; z-index:1; left: 0; top: 0" class="bigtitle"><img src="images/<%=ActImg%>" width="32" height="32"  align="absmiddle"><%=ActTitle%></div>

<br><br><br>
<font color=green>双击鼠标，屏幕自动滚动；单击鼠标，屏幕停止滚动.</font>

  <%
dim numm, RsCount
RsCount=Record.RecordCount
numm=(cint(Request.QueryString("Page"))-1)*Record.PageSize+cint(Request.QueryString("num"))-1
Record.Move(numm)
if Session("send")<>"common" then 
	if Record("iread")="f" then 
		Record("iread")="t"
		if Session("send")="recived" then Session("num")=Session("num")-1
	end if
	Record.update
end if

%> 
      <div align="right">
		   <form action="mailbox.asp?Page=<%=Request.QueryString("Page")%>&mailbox=<%Response.Write(Session("send"))%>" method="post">
		  
		  <input type="button"  class="css0"  value=" 上一封 " onClick="window.location.href='readmail.asp?num=<%=(cint(Request.QueryString("num"))-1)%>&Page=<%=Request.QueryString("Page")%>'"  <% if Request.QueryString("num")="1" then response.write("DISABLED")%>>
			<input type="button"  class="css0"  value=" 下一封 " onClick="window.location.href='readmail.asp?num=<%=(cint(Request.QueryString("num"))+1)%>&Page=<%=Request.QueryString("Page")%>'"  <% if Request.QueryString("num")=cstr(RsCount) then response.write("DISABLED")%>>
				
				<input type="submit" name="Submit2" class=css0  value="  回复  " onClick=op.value=1 <%if ActTime<>"收到时间" then %> disabled<%end if%>>
				
				<input type="submit" name="Submit32" class=css0  value="  删除  " onClick=op.value=2>
				<input type="submit" name="Submit4" class=css0 value="  返回  " onClick=op.value=3>
			<input type="hidden" name="num" value="<%=Request.QueryString("num")-1%>">
			<input type="hidden" name="op">
			<hr noshade size="1" color=#000099></form>
        </div>
<table width=100% border="0" cellspacing="1" cellpadding="1" class=css1>

    <tr align="center" bgcolor="#eeeeff"> 
      <td height="20"  width="15%"> <b><%=ActFrom%></b> </td>
      <td height="20"  width="20%"><%=Record(ActFromFieldName)%></td>
      <td height="20"  width="15%"><b><%=ActTime%></b></td>
      <td height="20"  width="30%"><%=Record("iDateTime")%></td>
      <td height="20"  width="10%"><b>级别</b></td>
      <td height="20"  width="10%"><%=Record("ilevel")%></td>
    </tr>
    <tr align="center" valign="top"> 
      <td width="15%" height="145" bgcolor="#eeeeff"><b>邮件内容</b> <br>
       <%if Record("iaddfile")<>"" and Record("iaddfile")<>"_" then %><a href="file/<%=replace(Record("iaddfile"),"_","")%>" ><img src=images/add.jpg width=12 height=24 border=0></a><%end if%>
      </td>
      <td align="left" height="145" colspan="5">
<% 
	response.write(replace(Record("iinfo"),chr(13),"<br>"))

%></td>
    </tr>
  </table>
      <div align="right">
		   <form action="mailbox.asp?Page=<%=Request.QueryString("Page")%>&mailbox=<%Response.Write(Session("send"))%>" method="post">
		  <hr noshade size="1" color=#000099>
		  <input type="button"  class="css0"  value=" 上一封 " onClick="window.location.href='readmail.asp?num=<%=(cint(Request.QueryString("num"))-1)%>&Page=<%=Request.QueryString("Page")%>'"  <% if Request.QueryString("num")="1" then response.write("DISABLED")%>>
			<input type="button"  class="css0"  value=" 下一封 " onClick="window.location.href='readmail.asp?num=<%=(cint(Request.QueryString("num"))+1)%>&Page=<%=Request.QueryString("Page")%>'"  <% if Request.QueryString("num")=cstr(RsCount) then response.write("DISABLED")%>>
				<input type="submit" name="Submit2" class=css0  value="  回复  " onClick=op.value=1 <%if ActTime<>"收到时间" then %> disabled<%end if%>>

				<input type="submit" name="Submit32" class=css0  value="  删除  " onClick=op.value=2>
				<input type="submit" name="Submit4" class=css0 value="  返回  " onClick=op.value=3>
			<input type="hidden" name="num" value="<%=Request.QueryString("num")-1%>">
			<input type="hidden" name="op">
			</form>
        </div>
<%
Session("Backid")=Record(ActFromFieldName)
Session("Backcent")="回复:"+Record("cent")
Session("Backinfo")=vbcrlf+vbcrlf+vbcrlf+vbcrlf+"原文内容如下："+vbcrlf+"-----------------------------"+vbcrlf+Record("iinfo")
%>
<% Record.close
set Record=nothing
con.close
set con=nothing
%>
