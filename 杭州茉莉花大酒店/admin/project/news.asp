<!--#include file="../check.asp"-->
<!--#include file="../../inc/conn.asp"-->
<!--#include file="lib.asp"-->
 <%
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('网络超时，或者你还没有登陆! ');this.location.href='../login.asp';</script>"
end if
 	temp_check_rights = Split(session("manconfig"),",")
	for temp_rights_count=0 to ubound(temp_check_rights)
	    if trim(temp_check_rights(temp_rights_count)) = "6" then
			rights_check_passkey = trim(temp_check_rights(temp_rights_count))
		end if
	next
	if rights_check_passkey <> "6" then
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('您的权限不足，请不要非法调用其它管理模块，否则您的账号将被系统自动删除!');this.location.href='../login.asp';</Script>"
	Response.end
	end if%>
<%
Call OpenData()
 CompanyID = Trim(Request("ID"))
' tid=trim(request("tid"))
If IsSubmit then  
  Dim msg  
  Set rs=server.createobject("adodb.recordset")
  If len(CompanyID)<=0 Then
   set rs_max=server.CreateObject("adodb.recordset")
     sql="select max(sequence) as maxid from Sbe_project"
     rs_max.open sql,conn,1,1
     if isnull(rs_max("maxid")) then
        sequence=1
     else
        sequence=rs_max("maxid")+1
     end if
     rs_max.close
     set rs_max=nothing	 
	'msg = "资讯添加成功!"
	Rs.open "Select * from Sbe_project where id Is null",conn,1,3	
	Rs.addnew
	Rs("Sequence")= sequence   
  Else
	'msg = "资讯修改成功！"
	Rs.open "Select * from Sbe_project where ID=" & clng(CompanyID) ,conn,1,3	
  End if
  Rs("tid")= Request.Form("select")
  Rs("title")=Request.Form("title")
  Rs("code") = Request.Form("code")
  Rs("Produced") = Request.Form("Produced")
  Rs("Quality") = Request.Form("Quality")
  Rs("Ulnarcode") = Request.Form("Ulnarcode")
 ' Rs("writer")= Request.Form("writer")
  Rs("keyword")= Request.Form("keyword")
  if  Request.Form("newsdate")<>"" then
  Rs("newsdate")= Request.Form("newsdate")
  end if
  Rs("content")= Request.Form("content")
  Rs("pic")= Request.Form("pic")
  if Request.form("tuijian")="" then
     Rs("tuijian")=0
   else
     Rs("tuijian")= Request.Form("tuijian")
  end if
  if Request.form("Newproducts")="" then
     Rs("Newproducts")=0
   else
     Rs("Newproducts")= Request.Form("Newproducts")
  end if
  Rs("PhotoNew")= Request.Form("PhotoNew")    
 Rs("price")=Request.Form("price")
 Rs("detail")=Request.Form("detail")
		Rs("Bpic")=Request.Form("Bpic")
		Rs("Bpic2")=Request.Form("Bpic2")
		Rs("spic")=Request.Form("spic")
		Rs("Show")=request("Show")
		Rs("leibie")=request("leibie")
		Rs("writer")=request("writer")		
  rs.update
  rs.close
  Set rs=nothing	
   If len(CompanyID)<=0 Then
	Response.Write"<script>alert('客户增加成功');this.location.href='news.asp';</script>"
   Else
    Response.Write"<script>alert('客户修改成功');this.location.href='main.asp';</script>"
   End IF

ElseIF Len(CompanyID)>0 Then	
	Dim StrSQL
	Dim objRec
	StrSQL = "Select * from Sbe_project Where ID=" & CompanyID	
	Set objRec=server.createobject("adodb.recordset")
	 objRec.open StrSQL,conn,1,1
	With ObjRec
		If .Eof And .Bof Then
			Response.Write "<Script>alert('操作失败');history.back();</script>" 
			Response.End
		Else
		    code = objRec("code")
		    Produced = objRec("Produced")
			Quality = objRec("Quality")
			Ulnarcode= objRec("Ulnarcode")
			title = objRec("title")
			tid= objRec("tid")  
            writer= objRec("writer")
            newsdate= objRec("newsdate")
            content= objRec("content")
			tuijian= objRec("tuijian")          
            PhotoNew= objRec("PhotoNew")
			keyword=objRec("keyword")
			pic=objRec("pic")
			detail=objRec("detail")
			Show=objRec("Show")
			Spic=objRec("Spic")
			Bpic=objRec("Bpic")
			leibie=objRec("leibie")
			Newproducts=objRec("Newproducts")
			Bpic2=objRec("Bpic2")
		End If
	End With
	objRec.Close:set objRec=Nothing
elseif Len(CompanyID)=0 then
PhotoNew=true	
End if
'Private Sub MessageBoxOK(strValue,tid)
	'With Response
		'.Write "<script>" & vbcrlf
		'.Write "alert('"+strValue+"');" & vbcrlf
		'.Write "this.location.href='"& request.Cookies("refer_page") &"?tid="& tid &"';" & vbcrlf
		'.Write "</'script>" & vbcrlf
	'End With
'End Sub
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>添加客户</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<script language="javascript">
  function foreColor()
   {
    var arr = showModalDialog("../eWebEditor/Dialog/selcolor.htm", "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0");
    if (arr != null) document.add.title.value='<font color='+arr+'>'+document.add.title.value+'</font>'
    else document.add.title.focus();
}

function clk(value){
 add.writer.value=value;
}
</script>
<script language="JavaScript">
function check(){
  if(add.select.value==""){
     alert("请选择所属区域！");
	 add.select.focus();
	 return false;
  }
  if(add.title.value==""){
     alert("请填写代理商名称！");
	 add.title.focus();
	 return false;
  }
//  if(add.code.value==""){
//     alert("请填写客户编码！");
//	 add.code.focus();
//	 return false;
//  }
 document.add.Submit.disabled=true;
 document.add.value="请稍候..."
  return true;
}
function show_spic_menu(menu_id)
{
if (menu_id==1)
{
eval("show_spic.style.display=\"\";");
}
else
{
eval("show_spic.style.display=\"none\";");
}
}
</script>
<script language="JavaScript" src="../include/meizzDate.js"></script>
</head>

<body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" >
  <tr> 
    <td height="25"><font color="#6A859D">合作客户&gt;&gt; 客户管理 </font></td>
  </tr>
  <tr>
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>


<br>
<form name="add" method="post" action="" onSubmit="return check()">
<table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0" id="sbe_table">
  <tr align="center">
    <td colspan="3" class="sbe_table_title">客户管理</td>
  </tr>
  <tr>
    <td width="16%" align="right">所属省市:</td>
    <td colspan="2">
	     <select name="select" class="sbe_button" style="width:132px;">
           <option value="">请选择地区</option>
          <%
		    Call ShowClass("Sbe_project",tid)%>
        </select></td>
  </tr>
  <tr>
    <td align="right">代理商名称:</td>
    <td colspan="2"><input name="title" type="text" class="input" id="title" style="width:180px;" value="<%=title%>" maxlength="200">
    <!--<img class="Ico" src="../eWebEditor/ButtonImage/standard/forecolor.gif" onClick="foreColor();">--></td>
  </tr>
    <tr>
    <td align="right">联系人:</td>
    <td colspan="2"><input name="code" type="text" class="input" id="code" value="<%=code%>" maxlength="50"></td>
  </tr>
  <tr>
    <td align="right">联系电话:</td>
    <td colspan="2"><input name="Produced" type="text" class="input" id="Produced" value="<%=Produced%>" maxlength="50"></td>
  </tr>
  <tr>
    <td align="right">QQ:</td>
    <td colspan="2"><input name="Quality" type="text" class="input" id="Quality" value="<%=Quality%>" maxlength="50"></td>
  </tr>
  <tr>
    <td align="right">Email:</td>
    <td colspan="2"><input name="keyword" type="text" class="input" id="keyword" value="<%=keyword%>" maxlength="50"></td>
  </tr>
  <tr>
    <td align="right">联系地址:</td>
    <td colspan="2"><input name="Ulnarcode" type="text" class="input" id="Ulnarcode" style="width:300px;" value="<%=Ulnarcode%>" maxlength="200"></td>
  </tr>
  <tr style="display:none;">
    <td align="right">上传大图:</td> 
    <td width="23%"><input name="Bpic2" type="text" class="input" value="<%=Bpic2%>" size="25"></td>
    <td width="61%"><iframe src="../upload/upload.asp?Form_Name=add&UploadFile=Bpic2" width="65%" height="25" frameborder="0" scrolling="no"></iframe> 
    图片尺寸比例：252*176</td>
  </tr>
  <tr style="display:none;">
    <td align="right">资讯来源:</td>
    <td colspan="2"><input name="writer" type="text" class="input" id="writer" value="<%=writer%>">
      选择:<%Call news_come_Class()%> ---<a href="news_come_class.asp" onClick="window.open(this.href,'', 'height=350,width=400,toolbar=no,location=no,status=no,menubar=no');return false">资讯来源设置</a></td>
  </tr>
  <tr style="display:none;">
    <td align="right">是否推荐:</td>
    <td colspan="2"> <input type="radio" name="Tuijian" value="1" <%Call ReturnSel(tuijian,true,2)%> onClick="show_spic_menu(1)">推荐 <input name="Tuijian" type="radio" value="0" <%call ReturnSel(Tuijian,false,2)%> onClick="show_spic_menu(0)"> 不推荐</td>
  </tr>
<tr id="show_spic" <%if tuijian=false then response.Write("style='display:none;'") end if%>>
    <td align="right">缩&nbsp;&nbsp;略&nbsp;图:</td> 
    <td width="23%"><input name="Spic" type="text" class="input" value="<%=Spic%>" size="25"></td>
    <td width="61%"><iframe src="../upload/upload.asp?Form_Name=add&UploadFile=Spic" width="65%" height="25" frameborder="0" scrolling="no"></iframe> 
    图片尺寸比例：126*88</td>
  </tr>
<!--  <tr style="display:none;">
    <td align="right">图片新闻:</td>
    <td colspan="2"><input type="radio" name="PhotoNew" value="1" <%'Call ReturnSel(PhotoNew,true,2)%>>
        是 &nbsp;&nbsp; <input name="PhotoNew" type="radio" value="0"  <%'Call ReturnSel(PhotoNew,false,2)%>>否</td>
  </tr>-->
<tr>
    <td align="right">所属类别:</td>
    <td colspan="2"> <input type="radio" name="PhotoNew" value="1" <%Call ReturnSel(PhotoNew,true,2)%>>
      代理商
      <input name="PhotoNew" type="radio" value="0" <%call ReturnSel(PhotoNew,false,2)%>>
      专卖店</td>
  </tr>
  <tr style="display:none;">
    <td align="right">简要说明:</td>
    <td colspan="2"><textarea name="detail" cols="50" rows="3" class="input" id="detail"><%=detail%></textarea></td>
  </tr> 
  <tr style="display:none;">
    <td align="right">详细内容:</td>
    <td colspan="2"><textarea name="content" id="textarea" style="display:none"><%=content%></textarea><iframe ID="eWebEditor1" src="../eWebEditor/ewebeditor.asp?id=content&style=sbe&savefilename=uploadfile" frameborder="0" scrolling="no" width="100%" HEIGHT="350"></iframe></td>
  </tr>
  <tr style="display:none;">
    <td align="right">上传图片:</td> 
    <td width="23%"><input name="pic" type="text" class="input" value="<%=newspic%>" size="25"></td>
    <td width="61%"><iframe src="../upload/upload.asp?Form_Name=add&UploadFile=pic" width="100%" height="25" frameborder="0" scrolling="no"></iframe></td>
  </tr>
  <tr style="display:none;">
    <td align="right">类&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;别:</td>
    <td colspan="2"> <input type="radio" name="leibie" value="1" <%if CompanyID="" then%>checked <%else%> <%Call ReturnSel(leibie,true,2)%> <%end if%>>
        中 &nbsp;&nbsp; <input name="leibie" type="radio" value="0" <%if Company<>"" then%> <%Call ReturnSel(leibie,false,2)%> <%end if%>>
        英</td>
  </tr>
  <tr>
    <td align="right">添加时间:</td>
    <td colspan="2"><input name="newsdate" type="text" class="input" id="newsdate" onFocus="setday(this)" <%if newsdate="" then response.Write ("value='"&date()&"'") else response.Write ("value='"&newsdate&"'") end if%>></td>
  </tr>
  <tr>
    <td align="right">是否显示:</td>
    <td colspan="2"> <input type="radio" name="Show" value="1" <%if CompanyID="" then%>checked <%else%> <%Call ReturnSel(Show,true,2)%> <%end if%>>
        是 &nbsp;&nbsp; <input name="Show" type="radio" value="0" <%if Company<>"" then%> <%Call ReturnSel(Show,false,2)%> <%end if%>>
        否</td>
  </tr>
  <tr align="center">
    <td colspan="3"><input type="hidden" name="ID" value="<%=CompanyID%>"><input name="Submit" type="submit" class="sbe_button" value="提交">
    <input name="Submit2" type="reset" class="sbe_button" value="重置"></td>
  </tr>
</table>
</form>
<%Call CloseDataBase()%>
</body>
</html>
<%
Private Sub news_come_Class()
'读取新闻来源
 Set oRs=Conn.Execute("select * from news_come_class order by id asc")
 IF oRs.Eof and oRs.bof Then Exit Sub
 Do While not oRs.eof 
  response.write "<a href=""javascript:clk('"& oRs("title") &"');"" >"& oRs("title") &"</a>/"& vbCrLf
 oRs.Movenext
 Loop
End Sub
%>