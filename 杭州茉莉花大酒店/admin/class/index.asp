<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('网络超时，或者你还没有登陆! ');this.location.href='../login.asp';</script>"
response.End
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理系统</title>

<link href="../include/style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="25"><font color="#6A859D">分类设置</font></td>
  </tr>
  <tr>
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>
<br>
<% OpenData()
   ClassTitle=Request("ClassTitle")
   if ClassTitle="sbe_product" then
      depth_num=cint(ProClass_num)
    elseif  ClassTitle="Sbe_news" then
	  depth_num=cint(NewsClass_num)
	elseif ClassTitle="Sbe_Company" then
	  depth_num=cint(qiyeClass_num)
	elseif ClassTitle="Sbe_Down" then
	  depth_num=cint(downClass_num)
   end if
   Parid=request("Parid")
   If Parid="" then Parid=0   
   Dim Act
   Act=Request("Act")   
   Select Case Act
     Case "up": Call Up()
     Case "save": Call SaveData()
	 Case "down": Call Down()
	 Case "modify": Call modify()
	 Case "savemodify": Call savemodify()
	 Case "lock": Call DoLock()
	 case "del": Call Del()
     Case Else: Call Main()	 
   End Select
   Call CloseDataBase()
 Sub Del()
   ID=Cint(Request.QueryString("ID"))
   Set Rs=Server.CreateObject("Adodb.recordset")
   Sql="Select * From "&classtitle&"_Class where id="&id
   Rs.Open Sql,Conn,1,3
    if classtitle="Sbe_Jishu" or classtitle="sbe_product" then
       IF Rs("ParID")=0 Then
          Sql="delete from "&classtitle&" where bigclass="&ID
          Conn.Execute Sql
	   else
	      Sql="delete from "&classtitle&" where Tid="&ID
	      Conn.Execute Sql
	   end if
	else
	      Sql="delete from "&classtitle&" where Tid="&ID
	      Conn.Execute Sql 
   end if
   if rs("childNum")=0 then   
      IF Rs("ParID")<>0 Then
      Sql="Update "&classtitle&"_Class Set ChildNum=ChildNum-1 Where ID="&Rs("ParID")
      Conn.Execute Sql
	  end if
	  sql="update "&classtitle&"_Class set Sequence=Sequence-1 where sequence>"&rs("sequence")
	  conn.execute sql
	  rs.delete
   else
      Call WriteErr("请先删除此分类下的子分类！",1)
   end if
   rs.close
   set rs=nothing
   response.redirect(request.ServerVariables("HTTP_REFERER"))
 End Sub
   
 Sub DoLock()
 ID=Cint(Request.QueryString("ID"))
 Set Rs=Server.CreateObject("adodb.recordset")
 sql="select * from "&ClassTitle&"_Class where ID="&id
   Rs.Open Sql,Conn,1,3
    if Rs("Lock")=false Then
	   Rs("Lock")=1
	 else
	   Rs("Lock")=0
	 end if
    Rs.Update
  Rs.Close
  Set Rs=Nothing
  Response.Redirect(request.ServerVariables("HTTP_REFERER")) 
 End Sub
 Sub SaveModify()
 id=Cint(request.Form("id"))
 classname=request.Form("classname")
 Spic=Trim(request.Form("Spic"))
 set rs=server.CreateObject("adodb.recordset")
 if ClassTitle="sbe_product" then
   sql="select *  from "&ClassTitle&"_Class where ID="&id&" "
   rs.open sql,conn,1,1
   if not (rs.eof and rs.bof) then
	 if rs("Depth")=0 then
		leixing=trim(request("leixing"))
		set rs1=server.CreateObject("adodb.recordset")
		sql1="select * from "&ClassTitle&"_Class where ID in ("&ChildrenID(ClassTitle,Cint(id))&") "
		rs1.open sql1,conn,1,3
'		response.Write sql1
'		response.End
		if not (rs1.eof and rs1.bof) then
		do while not rs1.eof
           Rs1("leixing")=leixing
        rs1.Update
        rs1.movenext
        loop
		end if
		rs1.close
	  else
		leixing=rs("leixing")		
      end if
	  end if		  
   rs.close
  end if
 if classname="" then Call WriteErr("请填写类名！",1)
 'if ClassTitle="sbe_product" then
details=request("details")
xl_name=request.Form("xl_name")
 'end if
 sql="select * from "&ClassTitle&"_Class where ID="&id
 Rs.Open Sql,Conn,1,3
   Rs("ClassName")=ClassName
    Rs("xl_name")=xl_name
 if ClassTitle="sbe_product"  then
    Rs("Spic")=Spic
'	 response.Write xl_name
' response.End
  if rs("Depth")<>0 then
       Rs("leixing")=leixing
	end if	   
 end if
  ' Rs("kroean_name")=kroean_name
  ' Rs("english_name")=english_name
   'Rs("Spic")=Spic   
   Rs.Update
 Rs.Close
 Set Rs=Nothing
 Response.Write("<script language=javascript>alert('修改成功！');window.location.href='"&request.Form("url")&"';</script>") 
 Response.End()
 End Sub
   
 Sub Up()
    MoveStep=Request.Form("MoveStep")
	MoveID=Request.Form("MoveID")
	If MoveStep="" Then MoveStep=1
    Call Upto(ClassTitle,MoveID,MoveStep)
    response.Redirect(request.ServerVariables("HTTP_REFERER"))
 End Sub
 
  Sub Down()
    MoveStep=Request.Form("MoveStep")
	MoveID=Request.Form("MoveID")
	If MoveStep="" Then MoveStep=1	
    Call DownTo(ClassTitle,MoveID,MoveStep)
    response.Redirect(request.ServerVariables("HTTP_REFERER"))
 End Sub
 
 
 '===========================
 '*  ClassTitle : 分类表名
 '*  MoveID : 要移动的分类ID
 '*  MoveStep ：移动位数
 '===========================
 
 
  Function DownTo(ClassTitle,MoveID,MoveStep)
    Set rs=server.CreateObject("adodb.recordset")	
	sql="select * from "&ClassTitle&"_Class where ID="&MoveID
	rs.open sql,conn,1,3
	   ParID=rs("ParID")
	   Sequence=Rs("Sequence")
	   ParPath=Rs("ParPath")&","&rs("ID")
	   MoveNum=1
	   StartSequence=Rs("Sequence")+1
	   ChildNum=Rs("ChildNum")
	   
	   '==== 获取需要移动数目 ===
	   If ChildNum>0 Then	   
	      Set rs2=Server.CreateObject("adodb.recordset")
	      sql="select * from "&ClassTitle&"_Class where ParPath like '"&ParPath&"%' order by sequence desc"
	      Rs2.Open sql,conn,1,3
	         MoveNum=MoveNum+Rs2.recordcount
			 StartSequence=rs2("sequence")+1
	   End If
	   '=====  End  ====
	   
	   
	   '====== 获取移动最后的Sequence值
	   Set Rs1=Server.CreateObject("adodb.recordset")
	   sql="Select * from "&ClassTitle&"_Class where Sequence>"&Sequence&" and ParID="&ParID&" order by sequence"
	   rs1.open sql,conn,1,1
	   if rs1.eof then
	      rs.close
		  set rs=nothing
	      rs1.close
		  set rs1=nothing
		  exit Function
	   end if
	   CountNum=rs1.recordcount
	   If Cint(MoveStep)>Cint(CountNum) Then MoveStep=CountNum
	   rs1.move MoveStep-1
	     EndSequence=Rs1("sequence")
		 EndChildNun=rs1("childnum")		 
	     TempPath=Rs1("ParPath")&","&rs1("ID")
	   rs1.Close
	   if EndChildNun>0 then
	     sql="select Max(sequence) as EndSequence from "&ClassTitle&"_Class where ParPath like '"&TempPath&"%'"
		 rs1.open sql,conn,1,1
		   EndSequence=rs1("EndSequence")
		 Rs1.close
	   End If 
	   Set rs1=Nothing	   
	   '=====  End  ====
	   
	   
	   sql="update "&ClassTitle&"_Class set Sequence=Sequence-"&MoveNum&" where sequence<="&EndSequence&" and sequence>="&StartSequence
	   conn.execute sql
	   
	   StartSequence=EndSequence-MoveNum+1
	   
	   rs("Sequence")=StartSequence
	   rs.update
	   rs.close
	   set rs=nothing
	   
	   If ChildNum>0 Then
	     ii=EndSequence
	     do while not rs2.eof
		   rs2("sequence")=ii
		   rs2.update
		   ii=ii-1
		   rs2.movenext
		 loop
		 rs2.close
		 set rs2=nothing
	   End If   	
 End Function
 Function Upto(ClassTitle,MoveID,MoveStep)
    MoveStep=Cint(MoveStep)
    Set rs=server.CreateObject("adodb.recordset")	
	sql="select * from "&ClassTitle&"_Class where ID="&MoveID
	rs.open sql,conn,1,3
	   ParID=rs("ParID")
	   Sequence=Rs("Sequence")
	   ParPath=Rs("ParPath")&","&rs("ID")
	   MoveNum=1
	   ChildNum=Rs("ChildNum")
	   
	   If ChildNum>0 Then
	   '==== 获取需要移动数目 ===
	      Set rs2=Server.CreateObject("adodb.recordset")
	      sql="select * from "&ClassTitle&"_Class where ParPath like '"&ParPath&"%' order by sequence"
	      Rs2.Open sql,conn,1,3
	         MoveNum=MoveNum+Rs2.recordcount	      	   
	   '=====  End  ====
	   End If
	   
	   Set Rs1=Server.CreateObject("adodb.recordset")
	   sql="Select * from "&ClassTitle&"_Class where Sequence<"&Sequence&" and ParID="&ParID&" order by sequence desc"
	
	   rs1.open sql,conn,1,1
	   if rs1.eof then
	      rs.close
		  set rs=nothing
	      rs1.close
		  set rs1=nothing
		  exit Function
	   end if
	   CountNum=rs1.recordcount
	   If Cint(MoveStep)>Cint(CountNum) Then MoveStep=CountNum	  	  
	   rs1.move MoveStep-1
	     StartSequence=Rs1("Sequence")
	   rs1.Close
	   Set rs1=Nothing 
	   
	   sql="update "&ClassTitle&"_Class set Sequence=Sequence+"&MoveNum&" where sequence>="&StartSequence&" and sequence<"&sequence
	   conn.execute sql
	   
	   rs("Sequence")=StartSequence
	   rs.update
	   rs.close
	   set rs=nothing
	   
	   If ChildNum>0 Then
	     ii=StartSequence+1
	     do while not rs2.eof
		   rs2("sequence")=ii
		   rs2.update
		   ii=ii+1
		   rs2.movenext
		 loop
		 rs2.close
		 set rs2=nothing
	   End If   	
 End Function

   Sub SaveData()
     ClassName=Trim(Request.Form("ClassName"))
	 If ClassName="" then Call WriteErr("请填写分类名！",1)
	 xl_name=request("xl_name")
	 Spic=request("Spic")
	 '== 检测参数 ==
	 set rs=server.CreateObject("adodb.recordset") 
	 
	 if ParID=0 Then '如果为第一层
	    sql="select Max(sequence) as maxid from "&ClassTitle&"_Class"
		rs.open sql,conn,1,1
		  if isnull(rs("maxid")) then
		     Sequence=1
		  else
		     Sequence=rs("maxid")+1			
		  end if		  
		 rs.close		
		 Depth=0		 
		 ParPath="0"
	     leixing=trim(request("leixing"))
	  else
	     sql="select * from "&ClassTitle&"_Class where ID="&ParID
		 rs.open sql,conn,1,3
		    ParDepth=rs("Depth")
			ParParPath=rs("ParPath")
			ParSequence=rs("sequence")
			rs("ChildNum")=rs("ChildNum")+1
	        if ClassTitle="sbe_product" then
		    leixing=rs("leixing")
			end if
			rs.update
		 rs.close   
         sql="Select Max(Sequence) as maxid from "&ClassTitle&"_Class where parid="&parid
	     Rs.open sql,conn,1,1
	     If isnull(rs("maxid")) Then
	        sql="update "&ClassTitle&"_Class set sequence=sequence+1 where Sequence>"&ParSequence
			conn.execute sql
			Sequence=ParSequence+1			
	     Else
	        sql="update "&ClassTitle&"_Class set sequence=sequence+1 where Sequence>"&rs("maxid")
			conn.execute sql
			Sequence=rs("MaxID")+1			
	     end if
		 rs.close
		 
		 ParPath=ParParPath&","&ParID		
		 Depth=ParDepth+1
	  end if  
	
	 Set rs=server.CreateObject("adodb.recordset")
	 sql="select * from "&ClassTitle&"_Class where id=0"
	 rs.open sql,conn,1,3
	   rs.addnew
	   rs("ParID")=ParID
	   rs("ClassName")=ClassName
	   rs("Sequence")=maxid
	   rs("Depth")=Depth
	   rs("ChildNum")=0
	   rs("ParPath")=ParPath
	   rs("sequence")=Sequence
	   rs("Lock")=0
       Rs("xl_name")=xl_name
 if ClassTitle="sbe_product" then
    Rs("leixing")=trim(leixing)
       Rs("Spic")=Spic
 end if
	   rs.update
	 rs.close
	 set rs=nothing

 	 response.Write("<script language=javascript>alert('分类添加成功！');window.location.href='index.asp?Depth="&Depth&"&Parid="&Parid&"&ClassTitle="&ClassTitle&"';</script>")
	 response.End()
   End Sub
   
   Sub Main()
%>

<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="24"> <a href="index.asp?parid=0&ClassTitle=<%=ClassTitle%>"><strong>一级分类</strong></a><strong> 
      <%ShowClass(parid)%>
      </strong></td>
  </tr>
</table> 
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
    <tr class="sbe_table_title"> 
      <td width="51%" height="25" class="sbe_table_title">分类名称(子分类数)</td>
      <td height="25" class="sbe_table_title">上移</td>
      <td height="25" class="sbe_table_title">下移</td>
      <td height="25" class="sbe_table_title">编辑</td>
      <td height="25" class="sbe_table_title">权限</td>
      <td height="25" class="sbe_table_title">删除</td>
    </tr> 
	<%
   Set rs=server.CreateObject("Adodb.recordset")
   Sql="select * from "&ClassTitle&"_Class Where ParID="&ParID&" order by Sequence"
   Rs.open Sql,conn,1,1
   if rs.eof then
   %>
    <tr> 
      <td height="25" colspan="6">此分类下暂且没有分类...</td>
    </tr>
    <%else
	AllRecordNum=rs.recordcount
	icount=1
	do while not rs.eof

	%>
    <tr>
      <td height="25"><font color="#0336699">
        <li type="circle"> <strong>
		<%if rs("Depth")<Depth_num-1 then%>
		<a href="index.asp?parid=<%=rs("id")%>&classtitle=<%=classtitle%>"><%=RS("ClassName")%></a>
		<%else%>
		<%=RS("ClassName")%>
		<%end if%></strong>(<%=Rs("ChildNum")%>)&nbsp;&nbsp;<%'if (ClassTitle ="sbe_product1" and rs("Depth")=0) then
		                                     'if rs("leixing")=1 then response.Write("产品")
		                                    ' if rs("leixing")=2 then response.Write("新闻")
										'end if%></li>
        </font></td>
	 <form name="move" method="post" action="index.asp">
	  <td width="13%" height="25" align="center" bgcolor="#E9EFF3"> 
        <%if icount>1 then%>
        <input type="hidden" name="moveid" value="<%=rs("id")%>">
        <input type="hidden" name="act" value="up"> 
		<input type="hidden" name="ClassTitle" value="<%=ClassTitle%>">
        <select name="movestep">
	      <%for j=1 to icount-1%>
          <option value="<%=j%>"><%=j%></option>
          <%next%>
        </select> <input type="submit" name="Submit2" value="上移" class="sbe_button">
	   <%end if%>      </td>
	  </form>

	   <form name="move" method="post" action="index.asp">
      <td width="15%" align="center"> 
        <%if icount<AllRecordNum then%>
        <input type="hidden" name="moveid" value="<%=rs("id")%>">
        <input type="hidden" name="act" value="down">
		<input type="hidden" name="ClassTitle" value="<%=ClassTitle%>">		
		<select name="movestep">
	      <%for j=1 to AllRecordNum-icount%>
          <option value="<%=j%>"><%=j%></option>
          <%next%>
        </select> <input type="submit" name="Submit22" value="下移" class="sbe_button"> 
		<%end if%>      </td>
	  </form>
      
    <td width="7%" align="center" bgcolor="#E9EFF3"><a href="index.asp?act=modify&classtitle=<%=classtitle%>&id=<%=rs("id")%>"><img src="../images/edit.gif" border="0"></a></td>
      
    <td width="7%" align="center">
<%if rs("lock")=true then
   response.Write("<a href='index.asp?id="&rs("id")&"&classtitle="&classtitle&"&act=lock' title='关闭'><b><font color=#FF0000>×</font></b></a>")
  else 
   response.Write("<a href='index.asp?id="&rs("id")&"&classtitle="&classtitle&"&act=lock' title='开放'><b><font color=#009900>√</font></b></a>")
  end if
	  %>    </td>
      
    <td width="7%" align="center">
	<%'if ((ClassTitle="Sbe_Company" and ParID=0) or (trim(rs("id"))=trim(jqzph))) then%>
	<%'if (ClassTitle="sbe_product" and ParID=0) then%>
	<!--<img src="../images/delete.gif" border="0">-->
	<%'else%>
<%if rs("childnum")=0 then%>
	<a href="index.asp?id=<%=rs("id")%>&classtitle=<%=classtitle%>&act=del" onClick="return confirm('确定删除吗？')">
	<%else%>
	<a href="#" onClick="javascript:alert('请先删除此分类下的所有子分类！');return false;">
	<%end if%>
	<img src="../images/delete.gif" border="0"></a>
	<%'end if%>	</td>
    </tr>
    <%
    Rs.movenext
	icount=icount+1
	loop
	end if
	RS.close
	set rs=nothing
	%> 
<%if (ClassTitle="sbe_product1" and ParID=0) then
else%>
 <form name="add" method="post" action="index.asp" OnSubmit="return CheckForm();">
    <tr> 
      <td height="25" colspan="6">新建同级分类： 
        <input name="ClassName" type="text" class="input" id="zt2" size="20"> 
        &nbsp;<%if (ClassTitle="sbe_product" and ParID=0) then%><input name="leixing" type="hidden" value="1" checked><!--产品<input name="leixing" type="radio" value="2">新闻--><%end if%>        <!--<%'if ClassTitle="sbe_product1" then%>
	  <div id="xie" style=" float:center;" width="110" hight="100"></div>
	  <%'end if%>--></td>
    </tr>
<tr <%=banben_display%>> 
      <td height="25" colspan="6">&nbsp;&nbsp;&nbsp;&nbsp;  英文名称&nbsp;： 
        <input name="xl_name" type="text" class="input" id="xl_name" value="" size="20"></td>
    </tr>
	<%if ClassTitle="sbe_product" then%>
<tr> 
      <td height="25" colspan="6" >&nbsp; 分类缩略图&nbsp;：
        <input name="Spic" type="text" class="input" id="Spic" value="nopic_c.jpg" size="30">  
                  
                    <iframe src="../upload/upload.asp?Form_Name=add&UploadFile=Spic" width="304" height="25" frameborder="0" scrolling="no"></iframe>
                    246*163</td>
    </tr>
	<%end if%>
	<tr> 
      <td height="25" colspan="6">&nbsp;
        <input name="Submit3" type="submit" value=" 添加 " class="sbe_button">
        &nbsp;&nbsp;&nbsp;&nbsp;
        <input name="Submit4" type="reset" value="重置" class="sbe_button">
        <input name="ClassTitle" type="hidden" id="ClassTitle" value="<%=ClassTitle%>"> 
        <input name="ParID" type="hidden" id="ParID" value="<%=ParID%>"> 
        <input name="act" type="hidden" id="act" value="save"></td>
    </tr>
  </form>
		<%end if%>
</table>

</table>
<br>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <form name="form1" method="post" action="index.asp">
    <tr> 
      <td width="56%" height="25"><br> <%
	  Set Rs=Server.CreateObject("adodb.recordset")
	  sql="select * from  "&ClassTitle&"_Class order by sequence"
	  rs.open sql,conn,1,1 
		do while not rs.eof
		str=""		
		if rs("parid")=0 Then
		  str=str&"<img src=""../images/Tplus.gif"" align=""middle"">"
		else		 
		  for i=0 to rs("depth")
		    str=str&"<img src=""../images/I.gif"" align=""middle"">&nbsp;"
		  next
		  str=left(str,len(str)-48)&"<img src=""../images/T.gif"" align=""middle"">"		  
		end if
		response.Write(str&rs("classname")&"<br>")
		rs.movenext
		loop
	  rs.close
	  set rs=nothing	    
	  %>
        <br>
      </td>
    </tr>
  </form>
</table>
 	 <%if ClassTitle="sbe_product" and ParID=0 then%>

	 
	 <%else%>
	 <script language="javascript">
document.add.ClassName.focus();
</script>
<%end if%>
<br>
<%End Sub
  Sub Modify()
  id=Cint(request.QueryString("id"))
  Set Rs=Server.CreateObject("adodb.recordset")
  sqlwhere="ClassName,details,Depth,xl_name,Spic"
  if ClassTitle="sbe_product" then sqlwhere=sqlwhere&",leixing"
  sql="select "&sqlwhere&" From "&ClassTitle&"_Class where id="&id
  rs.open sql,conn,1,1  
 %>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <form name="add" method="post" action="index.asp">
    <tr> 
      <td height="25" colspan="5"<%'if ClassTitle<>"Sbe_project" then response.Write("colspan='2'") end if%>>修改分类名称: 
        <input name="ClassName" type="text" class="input" id="ClassName" value="<%=rs("classname")%>" size="20"> 
        &nbsp;<%if (ClassTitle="sbe_product1" and rs(2)=0) then%><input name="leixing" type="radio" value="1" <%if rs(5)=1 then response.Write("checked") end if%>>产品<input name="leixing" type="radio" value="2" <%if rs(5)=2 then response.Write("checked") end if%>>新闻
        <%end if%><input name="leixing" type="hidden" value="1" checked>        <!--<%'if ClassTitle="sbe_product1" then%><div id="xie" style=" float:center;" width="110" hight="100"></div><%'end if%>--></td>
    </tr>

    <tr <%=banben_display%>> 
      <td height="25" colspan="3">&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;英文名称:      
        <input name="xl_name" type="text" class="input" id="xl_name" value="<%=rs(3)%>" size="20"></td>
    </tr>
	<%if ClassTitle="sbe_product" then%>
   <tr> 
      <td height="25" >&nbsp;&nbsp; 分类缩略图:
        <input name="Spic" type="text" class="input" id="Spic" value="<%if trim(rs(4))<>"" then response.Write rs(4) else response.Write("nopic_c.jpg") end if%>" size="30">
        <iframe src="../upload/upload.asp?Form_Name=add&UploadFile=Spic" width="304" height="25" frameborder="0" scrolling="no"></iframe>
        246*163</td>
		<td colspan="4"></td>
</tr>
<%end if%>
<tr> 
      <td width="55%" height="25" colspan="5"> &nbsp;
         <input name="Submit" type="submit" value=" 修改 " class="sbe_button">
         &nbsp;&nbsp;
         <input name="Submit56" type="reset" value="重置" class="sbe_button">
         <input name="ClassTitle" type="hidden" id="ClassTitle" value="<%=ClassTitle%>">
        <input name="id" type="hidden" id="id" value="<%=id%>">
        <input name="url" type="hidden" id="id" value="<%=request.ServerVariables("HTTP_REFERER")%>">
        <input name="act" type="hidden" id="act" value="savemodify"></td>
    </tr>
	<input name="name" type="hidden" value="<%=ClassTitle%>">
  </form>
</table>
 <%
 rs.close
 set rs=nothing
 End Sub
 Sub ShowClass(par_id)
   Set rs_Class=Conn.execute("Select ClassName,id,parid from "&ClassTitle&"_Class where id="&par_id)
   if not rs_class.eof then
      showClass(rs_class("Parid"))
      response.Write(" >> <a href='index.asp?ClassTitle="&ClassTitle&"&Parid="&rs_class("id")&"'>"&rs_class("classname")&"</a>")
   End If
   Set rs_class=nothing 
 End Sub
%>
</body>
</html>
<script language="javascript">
function CheckForm()
{
    if(document.add.ClassName.value==""){
   alert("系统提示\n类别不能为空")
   document.add.ClassName.focus();
   return false;
   } 
//if (eWebEditor1.getHTML()==""){    
//      alert("系统提示\n内容不能为空");    
//     return (false);
//    }
}	
</script>
<%if ClassTitle="sbe_product1" then%>
		       <script language="javascript">
				function dochangepic()
			{
					xie.innerHTML='<img src=../../UploadFile/'+add.Spic.value+' width=92 hight=66>';
				}
					xie.innerHTML='<img src=../../UploadFile/'+add.Spic.value+' width=92 hight=66>';
			   </script>
			   <%end if%>