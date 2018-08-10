<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<script language="JavaScript" src="../include/meizzDate.js"></script>
<%
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('网络超时，或者你还没有登陆! ');this.location.href='../login.asp';</script>"
end if
 	temp_check_rights = Split(session("manconfig"),",")
	for temp_rights_count=0 to ubound(temp_check_rights)
	    if trim(temp_check_rights(temp_rights_count)) = "2" then
			rights_check_passkey = trim(temp_check_rights(temp_rights_count))
		end if
	next
	if rights_check_passkey <> "2" then
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('您的权限不足，请不要非法调用其它管理模块，否则您的账号将被系统自动删除!');this.location.href='../login.asp';</Script>"
	Response.end
	end if%>
<%Dim Act
  Act=Request.Form("act")
  openData()
  If Act="save" Then
     Call SaveData()
  Else
     Call Main()
  End If
  Call CloseDataBase()
  
  Sub SaveData()
     Pname=Trim(Request.Form("Pname"))
	 Tid=Cint(Request.Form("Tid"))
	 Hot=Cint(Request.Form("Hot"))
	 Ptype=Trim(Request.Form("Ptype"))
	 Bpic=Trim(Request.Form("Bpic"))
	 spic=Trim(Request.Form("Spic"))
	 Price=Trim(Request.Form("Price"))
	 if price="" then price=0
	 leibie=Trim(Request.Form("leibie"))
	 Tuijian=Request.Form("Tuijian")
	 num=Request.Form("num")
	 Show=Request.Form("Show")
	 datet=Request.Form("datet")
	 detail =Request.Form("detail")
	 pic =Request.Form("pic")
	 shifou =trim(Request.Form("shifou"))	 
	 password =trim(Request.Form("password"))
  sqlsize ="select * from Sbe_Product_Class where ID ="&Tid
  set rssize=conn.execute(sqlsize)
  if not (rssize.eof and rssize.bof) then
    if  rssize("Depth") = 0 then
	   bigclass=rssize("ID")
	   else
       bigclass = rssize("ParID")
	end if
  end if 
  rssize.close
 set rssize=nothing
     Content = ""
     For i = 1 To Request.Form("content").Count
       Content = Content & Request.Form("content")(i)
     Next
     Uploadfile=request.Form("Uploadfile")	 
	 set rs_max=server.CreateObject("adodb.recordset")
     sql="select max(sequence) as maxid from Sbe_product"
     rs_max.open sql,conn,1,1
     if isnull(rs_max("maxid")) then
        sequence=1
     else
        sequence=rs_max("maxid")+1
     end if
     rs_max.close
     set rs_max=nothing	 
	 
	 
	 Set Rs=Server.CreateObject("adodb.recordset")
	 sql="select * From Sbe_Product Where ID=0"
	 Rs.Open Sql,Conn,1,3
	    Rs.AddNew
		Rs("Pname")=Pname
		Rs("Tid")=Tid
		Rs("Ptype")=Ptype
		Rs("Bpic")=Bpic
		Rs("spic")=spic
		Rs("Price")=Price
		Rs("Tuijian")=Tuijian
		Rs("Content")=request("Content")
		Rs("Uploadfile")=Uploadfile
		Rs("Show")=Show
		Rs("Sequence")=Sequence
		Rs("Hot")=Hot
		Rs("leibie")=leibie
		Rs("datet")=datet
		Rs("detail")=detail
		Rs("bigclass")=bigclass
		Rs("pic")=pic
		Rs("password")=password	
		Rs("shifou")=shifou
		rs("num")=num
'		Set Rs1=Server.CreateObject("adodb.recordset")
'		Sql="Select FieldTitle From Sbe_Product_Field Where Lock=0 "
'		Rs1.Open Sql,Conn,1,1
'		   do While Not Rs1.Eof
'		      Rs(CStr(rs1(0)))=request.Form(CStr(rs1(0)))
'		   Rs1.MoveNext
'		   Loop
'		Rs1.Close
'		Set Rs1=Nothing
		rs.update
	rs.Close
	Set Rs=Nothing 
    response.Write("<script language=javascript>alert('客房添加成功！');window.location.href='add.asp?tid="&tid&"';</script>")
	response.End()
  End Sub
  
  Sub Main()
  Tid=request("tid")
  if tid="" then tid=0
  tid=cint(tid)
  Set Rs=Server.CreateObject("adodb.recordset")
  Sql="select picAuto from Sbe_WebConfig"
  Rs.Open sql,Conn,1,1
     PicAuto=rs(0)
  Rs.Close
  Set Rs=Nothing
  
  Set Rs=Server.CreateObject("adodb.recordset")
  sql="Select FieldName,FieldShow,FieldShowLength,Show,FieldLength from Sbe_Product_Field Where Lock=1 order by Sequence "
  Rs.Open Sql,Conn,1,1
    if rs.recordcount<>10 then
	   response.Write("系统字段丢失，请检查SBE_PRODUCT_FIELD表！")
	   Response.End()
	end if  
%>
<html>                                                                               
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理系统</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function check(){
  if(form1.Tid.value==""){
     alert("请选择分类！");
	 form1.Tid.focus();
	 return false;
  }
  if(form1.Pname.value==""){
     alert("请填写客房名称！");
	 form1.Pname.focus();
	 return false;
  }
  if(document.form1.shifou[1].checked==true){
  if (form1.password.value==""){
     alert("请填写查看用户名！");
	 form1.password.focus();
	 return false;
	 } 
  } 
 document.form1.addbtn.disabled=true;
 document.form1.addbtn.value="请稍候..."
  return true;
} 
function show_user_rights_menu(menu_id)
{
if (menu_id==0)
{
eval("show_user_rights.style.display=\"none\";");
}
else
{
eval("show_user_rights.style.display=\"\";");
}
}
  function PasswordShow(flag){
   if (flag==1){
       Showpassword.style.display="";
	   }
   if (flag==0){
       Showpassword.style.display="none";
	   }
  }
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="25"><font color="#6A859D">客房中心 &gt;&gt; 添加客房</font></td>
  </tr>
  <tr>
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>
<br>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <form name="form1" method="post" action="add.asp" onSubmit="return check()">
      <tr <%Call OpenClose(rs(3))%>> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td colspan="2"><select name="Tid" class="sbe_button">
          <option>请选择...&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
          <%
		    Call ShowClass("sbe_product",tid)%>
        </select> </td>
    </tr>
    <%
	rs.movenext '移动到Ptype字段
	%>
    <tr <%Call OpenClose(rs(3))%>> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td colspan="2"> <%Call ShowField("Pname",rs(1),rs(2),"",rs(4))%> </td>
    </tr>
    <%
	rs.movenext '移动到Ptype字段
	%>
    <tr <%Call OpenClose(rs(3))%>  style="display:none"> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td colspan="2"><%Call ShowField("Ptype",rs(1),rs(2),"",rs(4))%></td>
    </tr>
    <%
	rs.movenext '移动到Bpic字段
	%>
    <tr <%Call OpenClose(rs(3))%> style="display:none"> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td width="23%"> <%Call ShowField("Bpic",rs(1),rs(2),"",rs(4))%></td>
      <td width="61%"><iframe src="../upload/upload.asp?Form_Name=form1&UploadFile=Bpic" width="64%" height="25" frameborder="0" scrolling="no"></iframe>
      (图片最佳尺寸:225*300)</td>
    </tr>
    <%
	rs.movenext '移动到Spic字段
	%>
    <tr <%Call OpenClose(rs(3))%>> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td><%Call ShowField("Spic",rs(1),rs(2),"",rs(4))%></td>
      <td>  <iframe src="../upload/upload.asp?Form_Name=form1&UploadFile=Spic" width="64%" height="25" frameborder="0" scrolling="no"></iframe> 
        
        (图片最佳尺寸:393*278)</td>
    </tr>
    <%
	rs.movenext '移动到Price字段
	%>
<!--    <tr <%'Call OpenClose(rs(3))%> > 
      <td height="25" align="center"><%'=rs(0)%></td>
      <td colspan="2"><%'Call ShowField("Price",rs(1),rs(2),"",rs(4))%></td>
    </tr>-->
    <%
	rs.movenext '移动到Tuijian字段
	%>
    <tr <%Call OpenClose(rs(3))%>  style="display:none"> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td colspan="2"><input type="radio" name="Tuijian" value="1">
        是 &nbsp;&nbsp; <input name="Tuijian" type="radio" value="0" checked>否</td>
    </tr>
<!--  <tr id="show_user_rights" style="display:none;">    onclick=show_user_rights_menu(1)
    <td align="center">上传图片</td> 
    <td width="23%"><input name="pic" type="text" class="input"size="25"></td>
    <td width="61%"><iframe src="../upload/upload.asp?Form_Name=form1&UploadFile=pic" width="304" height="25" frameborder="0" scrolling="no"></iframe> 图片尺寸：112*148</td>
  </tr>-->
    <%
	rs.movenext '移动到产品类型字段
	%>
    <tr <%if rs(3)=true then%><%=banben_display%><%else%><%Call OpenClose(rs(3))%><%end if%>> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td colspan="2"> <input type="radio" name="leibie" value="1" checked="checked">
        中 &nbsp;&nbsp; <input name="leibie" type="radio" value="2">
        英</td>
    </tr>
<tr  style="display:none"> 
      <td height="25" align="center">查看权限</td>
      <td colspan="2"> <input type="radio" name="shifou" value="0" checked="checked" onClick="PasswordShow(0)">
        不需用户名 &nbsp;&nbsp; <input name="shifou" type="radio" value="1" onClick="PasswordShow(1)">
        需要用户名<span id="Showpassword" style="display:none;">&nbsp;&nbsp;
        <input name="password" type="text" class="input" style="ime-mode:Disabled;" value="" size="25" maxlength="20">
        &nbsp;(<font color="#FF0000">请输入用户名</font>)</span></td>
    </tr>
   <%Set Rs1=Server.CreateObject("adodb.recordset")
     Sql="Select FieldName,FieldShow,FieldShowLength,FieldTitle,Show,FieldLength from Sbe_Product_Field Where Lock=0 order by Sequence"
	 rs1.open sql,conn,1,1 
	   do while not rs1.eof%>
    <tr <%if Rs1(4)=0 then response.Write("class=""display""") end if%>>
      <td height="25" align="center"><%=rs1(0)%></td>
      <td colspan="2"><%Call ShowField(rs1(3),rs1(1),rs1(2),"",rs1(5))%></td>
    </tr>
	<% rs1.movenext
	   Loop
	  Rs1.Close
	  Set Rs1=Nothing
	%>
    <%
	rs.movenext '移动到Content字段
	%>	
  <!--  <tr <%Call OpenClose(rs(3))%>  style="display:none"> 
      <td height="12" align="center"><%=rs(0)%> <input name="content" type="hidden" id="content"> <input name="uploadfile" type="hidden" id="uploadfile"></td>
      <td colspan="2"><iframe ID="eWebEditor1" src="../eWebEditor/ewebeditor.asp?id=content&style=sbe&savefilename=uploadfile" frameborder="0" scrolling="no" width="100%" HEIGHT="350"></iframe></td>
    </tr>-->
    <tr <%Call OpenClose(rs(3))%>>
      <td height="6" align="center">数量</td>
      <td colspan="2"><input name="num" type="text" id="num"></td>
    </tr>
    <tr <%Call OpenClose(rs(3))%>>
      <td height="7" align="center">描述</td>
      <td colspan="2"><textarea name="content" cols="50" rows="10" id="content"></textarea></td>
    </tr>
    <tr <%Call OpenClose(rs(3))%>>
      <td height="25" align="center">价格</td>
      <td colspan="2"><textarea name="price" cols="50" rows="8" id="price"></textarea></td>
    </tr>
    <%
	rs.movenext '移动到Tuijian字段
	%>
    <tr <%Call OpenClose(rs(3))%>> 
      <td height="25" align="center"><%=rs(0)%></td>
      <td colspan="2"> <%Call ShowField("Show",rs(1),rs(2),false,rs(4))%></td>
    </tr>
    <%rs.close
	  set rs=Nothing
	  '关闭
	  %>
    <tr> 
      <td height="40" align="center">&nbsp;</td>
      <td colspan="2"> <input name="addbtn" type="submit" value=" 增加 " class="sbe_button"> 
        &nbsp; <input type="reset" name="Submit2" value=" 清空 " class="sbe_button">
        <input name="act" type="hidden" id="act" value="save"> </td>
    </tr>
  </form>
</table>
<br>
</body>
</html>
<% End Sub%>
<%
 Sub ShowField(FieldName,FieldType,FieldLength,Fieldvalue,chandu)
   If FieldType=5 Then
   if FieldName="Show" then
      Fieldvalue1=true
	else
      Fieldvalue1=false
   end if
	 %>
<input type="radio" name="<%=FieldName%>" value="1" <%Call ReturnSel(Fieldvalue1,true,2)%>>
        是 &nbsp;&nbsp; <input name="<%=FieldName%>" type="radio" value="0"  <%Call ReturnSel(Fieldvalue1,false,2)%>>
        否	  
<%Elseif FieldType=2 Then 
      Response.Write("<textarea name="""&FieldName&""" cols="""&FieldLength&""" rows=""3"" class=""input"">"&FieldValue&"</textarea>")
   elseIf FieldType=3 Then
      Response.Write("<input type=""password"" name="""&FieldName&""" size="""&FieldLength&""" value="""&FieldValue&""" class=""input"" maxlength="""&chandu&""">")
   elseIf FieldType=4 Then
      Response.Write("<input type=""hidden"" name="""&FieldName&""" value="""&FieldValue&""" class=""input""><iframe ID=""eWebEditor1"" src=""../eWebEditor/ewebeditor.asp?id=content&style=sbe&savefilename=uploadfile"" frameborder=""0"" scrolling=""no"" width=""100%"" HEIGHT=""350""></iframe>") 
   else
   'If FieldType=1 Then
   if FieldName="datet" then
      Response.Write("<input type=""text"" name="""&FieldName&""" onFocus=""setday(this)"" size="""&FieldLength&""" class=""input"" value="""&date()&""" maxlength="""&chandu&""" readonly>")
	  else
      Response.Write("<input type=""text"" name="""&FieldName&""" size="""&FieldLength&""" class=""input"" value="""&FieldValue&""" maxlength="""&chandu&""">")
	 end if
   End If
 End Sub
 
 Sub OpenClose(Flag)
   If Flag=false Then Response.Write("style=""display:none""")
 End Sub

%>