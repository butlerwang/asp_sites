<% if Session("id")="" then Response.Redirect("../error3.html") %>
<% if Request.Form("op")=1 then Response.Redirect("write.asp?op=back")%>
<!--#INCLUDE FILE="check.asp" --><%
'====设置要打开的信箱参数===========================================================================================
Session("send")=Request.QueryString("mailbox")
select case Session("send")
case "sendout" 
   ActFrom="发给"
   ActTime="发出时间"
   ActFromFieldName="ito"
   ActTitle="发件箱"
case "del"
   ActFrom="发给/来自"
   ActTime="时  间"
   ActFromFieldName="ifrom"
   ActTitle="回收站"
case "common"
   ActTime="发布时间"
   ActFrom="发布人"
   ActFromFieldName="ifrom"
   ActTitle="公共信件"
case "recived"
   ActTime="收到时间"
   ActFrom="来自"
   ActFromFieldName="ifrom"
   ActTitle="收件箱"
end select
'====设置要打开的信箱参数代码结束===========================================================================================
%>

<html>
<head>
<!--  #Include File=conn.asp -->
<link name="stsheet" id="stsheet" rel="stylesheet" href="images/vblife.css">

</head>
<script language="vbscript" >
function MM_changeProp() 
if  document.body.clientWidth>400 and  document.body.clientHeight>300 then

document.all.layer1.style.left=document.body.clientWidth-380+document.body.scrollLeft
document.all.layer1.style.top=document.body.clientHeight-300+document.body.scrollTop

end if
end function

</script>
<script language="vbscript">
if screen.width>1000 then document.stylesheets("stsheet").href="images/vblife1024.css"
</script>

<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') {
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (val<min || max<val) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  //document.MM_returnValue = (errors == '');
}

function MM_popupMsg(msg) { //v1.0
  document.MM_returnValue =alert(msg);
}
//-->
</script>


<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" onscroll="vbscript:MM_changeProp" onResize="vbscript:MM_changeProp" onload="vbscript:MM_changeProp">
<p class="al">&nbsp;</p>
<div id="Layer88" style="position:absolute; width:100%; height:33px; z-index:0; left: 0; top: 0"><img src="IMAGES/line.jpg" width="100%" height="100%"></div>
<div id="layer1" style="position:absolute; width:200px; height:115px; z-index:0; left: 1000; top: 1000; visibility: visible;" ><img src="IMAGES/bg.jpg" width="375" height="300" name="bgimg" class="al"></div>
<div id="layer2" style="position:absolute; z-index:2; left: 0px; top: 0px; visibility: visible"  >
<div style="position:absolute; z-index:5; left: 0px; top: 0px; visibility: visible" class="bigtitle"><img src="images/<% if Session("send")="sendout" then %>sendout.gif<% elseif Session("send")="del" then%>del.gif<%else %>recived.gif<%end if%>" width="32" height="32" align="absmiddle"><%=ActTitle%></div> 
<br><br><br>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="20" nbr>
  <tr> <form action="mailbox.asp?mailbox=<%=Session("send")%>" method="post">
    <td  align="left" class=css1 nowrap> 

<%
'=====删除文件函数======================================================================================	
				function FileDel(FileName)
				
	  					Dim fso, f2
  						Set fso = CreateObject("Scripting.FileSystemObject")
 						
  						If fso.FileExists(Server.Mappath("../file/"+FileName)) Then
   							Set f2 = fso.GetFile(Server.Mappath("../file/"+FileName))
   							f2.Delete
   							FileDel=1
   						else
   							FileDel=2
   						end if
   						set f2=nothing
   						set fso=nothing
	
				end function
'=====删除文件函数结束===========================================================================================
	Numm=1
'======清空当前信箱=================================================================================================================
	if Request.QueryString("op")="delall" then  '清空
		if Session("send")="del" then   '如果是回收站,则删除文件
			Record.close
			Record.open("select * from "+Session("send")+Session("id")+" where iaddfile<>''")
			while not Record.eof
				FileDel replace(Record("iaddfile"),"_","")
				Record.movenext
			wend
		else                    '如果不是回收站,则添加到回收站
			sql=""
			if Session("send")="sendout" then '添加到回收站时不包含附件地址
				sql="insert into del"+Session("id")+"(iDateTime,iaddfile,ifrom,iinfo,ilevel,cent,iread) select  iDateTime,'_'+iaddfile,"+ActFromFieldName+",iinfo,ilevel,cent,'f' from "+Session("send")+Session("id")
			elseif Session("send")<>"common" then
				sql="insert into del"+Session("id")+"(iDateTime,iaddfile,ifrom,iinfo,ilevel,cent,iread) select  iDateTime,iaddfile,"+ActFromFieldName+",iinfo,ilevel,cent,'f' from "+Session("send")+Session("id")
			elseif Session("level")="0" then '公共信箱特殊处理
				sql="insert into del"+Session("id")+"(iDateTime,iaddfile,ifrom,iinfo,ilevel,cent,iread) select  iDateTime,iaddfile,"+ActFromFieldName+",iinfo,ilevel,cent,'f' from message"
			else
				response.write("<font color=red><b>对不起,您没有清空此信箱的权限.</b></font><br>")	
			end if
			if sql<>"" then	con.Execute(sql)
		end if
		
		'以下代码清空当前信箱
		'---------------------------------------------------------------------
		if Session("send")<>"common" then 
			con.Execute("delete * from "+Session("send")+Session("id"))
			if Session("send")="recived" then  Session("num")=0
			Record.close
			Record.open("select cent from "+Session("send")+Session("id"))	
		elseif Session("level")=0 then
			con.Execute("delete * from message")
			Record.close
			Record.open("select cent from message")
		end if
		'----------------------------------------------------------------------
	end if
'======清空当前信箱结束=================================================================================================================================================================
%>


<%
'=============删除一条记录===================================================================================================
	if Request.Form("op")=2 and Request.Form("num").Count <>0 then '删除记录
    		Numm=(Request.QueryString("Page")-1)*Record.PageSize+Request.Form("num")

              Record.Move Numm 
              if Session("send")<>"common" or Session("level")=0 or Session("id")=Record(ActFromFieldName) then
              	dim sql
			  	sql=""
			  	if Session("send")<>"del" then
					dim adstr
					if  Session("send")<>"sendout" then adstr=Record("iaddfile") else adstr="_"+Record("iaddfile")
					sql="insert into del"+Session("id")+"(iDateTime,iaddfile,ifrom,iinfo,ilevel,cent,iread) VALUES ('" & Record("iDateTime")&"','"&adstr&"','"&Record(ActFromFieldName)&"','"&""&Record("iinfo")&"','"&Record("ilevel")&"','"&Record("cent")&"','f')"
					con.Execute(sql)
			  	else
			  		if Record("iaddfile")<>"" then        '删除文件
			  			FileDel Record("iaddfile") 
  					end if

			  	end if
			  	
			  	
			  	
              	Record.Delete
              	Record.Update
              else
              	response.write("<font color=red><b>对不起,您没有删除此条信息的权限.</b></font><br>")
              end if
              if not Record.eof and not Record.eof then Record.MoveFirst
              
    end if
'=============删除一条记录代码结束===================================================================================================
		

	NumberOfRecord=Record.RecordCount
	NumberOfPage=Record.PageCount
	Numm=1
	if Request.QueryString("Page").Count<>0 then iPage=clng(Request.QueryString("Page"))
	'response.write(Request.Form("go").Count)
	if Request.Form("go").Count<>0 then iPage=clng(Request.Form("go"))
	if iPage=0 then iPage=1
	if NumberOfRecord>0 then 
   		response.Write("共有" & NumberOfRecord & "封&nbsp;分" & NumberOfPage & "组&nbsp;第" & iPage & "组")

%>


<% 	Group=(iPage-(iPage) mod 10)/10

	if Group>0 then %> <a href="mailbox.asp?Page=<%=Group*10-9%>&mailbox=<%=Session("send")%>">上10组</a> <%end if%>    
<%   
	for i=Group*10+1 to Group*10+10    
        	if i<=NumberOfPage then %> <a href="mailbox.asp?Page=<%=i%>&mailbox=<%=Session("send")%>"><%=i%></a>    
<%			end if   
	Next    
   
	if NumberOfPage>Group*10+10 then %> <a href="mailbox.asp?Page=<%=Group*10+11%>&mailbox=<%=Session("send")%>">下10组</a>&nbsp;    
<%	end if%>   
   
   
   
        <input name=submit1 class="css0"  type="submit" value="转到组">   
        <input name=go type=text size=4  maxlength="6" value=<%=iPage%> class="css0">   
   
    </td></form>   
  </tr>   
</table>   
   
    
  <table border="0" cellspacing="1" cellpadding="0" class=css1 width="100%">
	<tr bgcolor="#0066CC" align="center"> 
	  <td height="28" width="35"> <b><font color="#FFFFFF">附件</font> </b> </td>   
      <td height="28" width="78" bgcolor="#0066CC"><b><font color="#FFFFFF"><%=ActFrom%></font></b></td>   
      <td height="28"> <b><font color="#FFFFFF">邮件主题</font></b></td>   
      <td height="28" width="140"> <b><font color="#FFFFFF"><%=ActTime%></font> 
		</b> </td>   
      <td height="28" width="35"> <b><font color="#FFFFFF">级别</font></b></td>   
  </tr>   
</table>   
         
<% if not Record.EOF then Record.Move (iPage-1)*Record.PageSize,0    
   
      	do while not Record.EOF and Numm<=Record.PageSize   %>   
  <table border="0" cellspacing="1" cellpadding="0" class=css1 width="100%" bordercolor="0"  >
	<tr align="center" bgcolor="#eeeeFF"> 
	  <td height="28" width="35" > 
		<%if Record("iaddfile")<>"" and Record("iaddfile")<>"_" then %>
		<a href="file/<%=replace(Record("iaddfile"),"_","")%>" ><img src=images/add.jpg width=12 height=24 border=0></a> 
		<%end if%>
	  </td>   
<%if Record("iread")="f" then %>   
      <td height="28" width="78" ><b><a href="readmail.asp?num=<%=Numm%>&Page=<%=iPage%>" ><%=Record(ActFromFieldName)%></a></b></td>   
      <td height="28" align="left"><b><a href="readmail.asp?num=<%=Numm%>&Page=<%=iPage%>"><%=Record("cent")%></a></b></td>   
      <td height="28" width="140" nowrap><%=Record("iDateTime")%></td>   
      <td height="28" colspan="2" width="35" nowrap><%=Record("ilevel")%></td>   
<%else%>   
      <td height="28" width="78" ><a href="readmail.asp?num=<%=Numm%>&Page=<%=iPage%>" ><%=Record(ActFromFieldName)%></a></td>   
      <td height="28" align="left" ><a href="readmail.asp?num=<%=Numm%>&Page=<%=iPage%>"><%=Record("cent")%></a></td>   
      <td height="28" width="140" nowrap><%=Record("iDateTime")%></td>   
      <td height="28" colspan="2" width="35" nowrap><%=Record("ilevel")%></td>   
<%end if%>   
  </tr>   
</table>   
<%    
	Numm=Numm+1   
	Record.MoveNext   
	loop   
%>  
<form name="delall" method="POST" action="mailbox.asp?op=delall&mailbox=<%=Session("send")%>"  >
 &nbsp;&nbsp;&nbsp;&nbsp; <input type="submit" value="清空<%=Acttitle%>" name="B3" class=css0 <%if session("level")<>"0" and session("send")="common" then response.write("disabled")%>>
</form>
<script language="vbscript">
function delall_ONSUBMIT()

delall_ONSUBMIT=(msgbox("该操作将无法恢复,你确认要全部删除?",52,"警告!"))=vbyes

end function
</script>

<hr size="1" noshade>   
<font color=red><span class="css1"> 注意：所有附件只保存60天，请收信人及时下载！</span></font>    
<%  
	else  
		response.Write("<font color=red>"+ActTitle+"中没有中没有任何信件。</font>")  
	end if  
  
%>
</div>
</body>
</html>