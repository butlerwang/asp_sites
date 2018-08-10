<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%openData()
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('网络超时，或者你还没有登陆! ');this.location.href='../login.asp';</script>"
end if
  IF instr(webConfig,", 6")=0 or instr(session("manconfig"),", 6")=0 Then'网站功能配置
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('您的权限不足，请不要非法调用其它管理模块，否则您的账号将被系统自动删除!');this.location.href='../login.asp';</Script>"
Response.end
end if%>
<%Dim Act
  Act=Request.Form("act")
  Select Case Act
    Case "save" : Call SaveData()
    Case else : Call Main()
  End Select
  Call CloseDataBase()
    
  Sub SaveData()
    Department = Request.Form("Department")
	Job = Request.Form("Job")
	Sex = Request.Form("Sex")
	Age = Request.Form("Age")
	Education = Request.Form("Education")
	Years = Request.Form("Years")
	Money = Request.Form("Money")
	Num = Request.Form("Num")
	EffectTime = Request.Form("EffectTime")
	Contact = Request.Form("Contact")
	Tel = Request.Form("Tel")
	Content = Request.Form("Content")
	leibie = Request.Form("leibie")
	AddDate= Request.Form("AddDate")
	address= Request.Form("address")
	workingway= Request.Form("workingway")
	yingjie= Request.Form("yingjie")	
	Show=trim(Request.Form("Show"))
	Other= Request.Form("Other")
	If Job="" Then Call WriteErr("请填写招聘职位！",1)	
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="Select * From Sbe_Job Where ID=0"
	Rs.Open Sql,Conn,1,3
	   Rs.AddNew
	   Rs("Department")=Department
	   Rs("workingway")=workingway
	   Rs("yingjie")=yingjie   
	   Rs("Job")=Job
	   Rs("Sex")=Sex
	   Rs("Age")=Age
	   Rs("Education")=Education
	   Rs("Years")=Years
	   Rs("Money")=Money
	   Rs("Num")=Num
	   Rs("EffectTime")=EffectTime
	   Rs("Contact")=Contact
	   Rs("Tel")=Tel
	   Rs("Content")=Content
	   Rs("show")=1
	   Rs("AddDate")=AddDate
	   Rs("leibie")=leibie
	   Rs("address")=address
	   Rs("Show")=Show
	   Rs("Other")=Other	   	   	   
	   Rs.Update
	Rs.Close
	Set Rs=Nothing
	Response.Write("<script language=javascript>alert('招聘信息发布成功！');window.location.href='add.asp';</script>")
	Response.End()	
  End Sub  
  
  Sub Main()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理系统</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../include/meizzDate.js"></script>
<!--<script language="javascript" src="../admin.js"></script>-->
<script language="JavaScript">
function check(){
  if(form1.Job.value==""){
     alert("请选填写岗位名称！");
	 form1.Job.focus();
	 return false;
  }
  return true;  
}
</script></head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="25"><font color="#6A859D"> 在线招聘 &gt;&gt; 发布招聘信息</font></td>
  </tr>
  <tr> 
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>
  
<br>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" id="sbe_table">
  <form name="form1" method="post" action="" onSubmit="return check()">
    <tr > 
      <td height="25" colspan="2"  class="sbe_table_title">发布<strong>招聘信息</strong></td>
    </tr>
    <tr > 
      <td height="25" align="center"><strong>所属部门</strong></td>
      <td height="21"> <input name="Department" type="text" class="input" id="Department"></td>
    </tr>
    <tr > 
      <td width="13%" height="25" align="center"><strong>岗位名称</strong></td>
      <td width="87%" height="21"><input name="Job" type="text" class="input" id="Job"></td>
    </tr>
    <tr> 
      <td height="25" align="center"><strong>招聘人数</strong></td>
      <td height="21"><input name="Num" type="text" class="input" id="Num"></td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>性别要求</strong></td>
      <td height="21">
	    <select name="Sex" class="sbe_button" id="Sex">
          <option value="不限" selected>男女不限</option>
          <option value="男性">男性</option>
          <option value="女性">女性</option>
        </select></td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>年龄要求</strong></td>
      <td height="21"><input name="Age" type="text" class="input" id="Age"> </td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>学历要求</strong></td>
      <td height="21">
	   <select name="Education" class="sbe_button" id="Education">
          <option value="学历不限" selected>学历不限</option>
          <option value="博士以上">博士以上</option>
          <option value="硕士以上">硕士以上</option>
          <option value="本科以上">本科以上</option>
          <option value="大专以上">大专以上</option>
          <option value="中专以上">中专以上</option>
          <option value="职高/技校以上">职高/技校以上</option>
          <option value="高中以上">高中</option>
          <option value="初中以上">初中</option>
        </select></td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>是否应届</strong></td>
      <td height="21"><input name="yingjie" type="radio" value="应届">
      应届 &nbsp;<input name="yingjie" type="radio" value="已工作">已工作 &nbsp;<input name="yingjie" type="radio" value="应届、已工作均可" checked>
      应届、已工作均可</td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>工作年限</strong></td>
      <td height="21"><input name="Years" type="text" class="input" id="Years"></td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>薪水范围</strong></td>
      <td height="21"><input name="Money" type="text" class="input" id="Money"></td>
    </tr>
    <tr class="display">
      <td height="25" align="center"><strong>联 系 人</strong></td>
      <td height="21"><input name="Contact" type="text" class="input" id="Contact"></td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>联系电话</strong></td>
      <td height="21"><input name="Tel" type="text" class="input" id="Tel"></td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>工作方式</strong></td>
      <td height="21"><input name="workingway" type="text" class="input" id="workingway"></td>
    </tr>
    <tr class="display"> 
      <td height="25" align="center"><strong>工作地点</strong></td>
      <td height="21"><input name="address" type="text" class="input" id="address"></td>
    </tr>
    <tr> 
      <td height="25" align="center"><strong>职位要求</strong></td>
      <td height="21"><textarea name="Content" cols="80" rows="8" class="input" id="Content"></textarea></td>
    </tr>
    <tr> 
      <td height="25" align="center"><strong>待遇</strong></td>
      <td height="21"><textarea name="Other" cols="80" rows="5" class="input" id="Content"></textarea></td>
    </tr>
	    <tr> 
      <td height="25" align="center"><strong>发布时间</strong></td>
      <td height="21"><input name="AddDate" type="text" class="input" id="AddDate" onFocus="setday(this)" value="<%=date()%>" readonly=""> 
      发布时间一般不需修改,默认即可.需修改的话,请注意时间的格式.</td>
    </tr>
	<tr> 
      <td height="25" align="center"><strong>截止日期</strong></td>
      <td height="21"><input name="EffectTime" type="text" class="input" id="EffectTime" onFocus="setday(this)" value="<%=date()+30%>" readonly="">
	  </td>
    </tr>
<tr <%=banben_display%>> 
      <td height="25" align="center"><strong>类&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;别</strong></td>
      <td height="21"><input name="leibie" type="radio" id="leibie" value="1" checked="checked">中 <input name="leibie" type="radio" id="leibie" value="2">英 招聘信息的类别.</td>
    </tr>
 <tr> 
      <td height="25" align="center"><strong>显&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;示</strong></td>
      <td height="21"><input name="Show" type="radio" id="Show" value="1" checked="checked">是 <input name="Show" type="radio" id="Show" value="0">
      否</td>
    </tr>
    <tr> 
      <td height="25" align="center">&nbsp;</td>
      <td height="21">
	    <input type="submit" name="Submit" value="发布信息" class="sbe_button"> 
        <input name="act" type="hidden" id="act" value="save">
        <input type="reset" name="Submit2" value="重置" class="sbe_button"></td>
    </tr>

  </form>
</table>
<br>
</body>
</html>
<%
  End Sub%>
