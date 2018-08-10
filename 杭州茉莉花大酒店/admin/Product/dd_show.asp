<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%
Call OpenData()%>
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
<%
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('网络超时，或者你还没有登陆! ');this.location.href='../login.asp';<'/script>"
end if
 	temp_check_rights = Split(session("manconfig"),",")
	for temp_rights_count=0 to ubound(temp_check_rights)
	    if trim(temp_check_rights(temp_rights_count)) = "8" then
			rights_check_passkey = trim(temp_check_rights(temp_rights_count))
		end if
	next
	if rights_check_passkey <> "8" then
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('您的权限不足，请不要非法调用其它管理模块，否则您的账号将被系统自动删除!');this.location.href='../login.asp';<'/Script>"
	Response.end
	end if%>
<%if request("act")="add" then
'     response.Write request("flag")
'     response.End
     Set Rs= Server.CreateObject("ADODB.RecordSet") 
     Rs.open "Select * from Sbe_order where ID=" & clng(request("id")),conn,1,3
     Rs("showtime")= date()
     Rs("status")= request("flag")
     rs.update  
     rs.close
     Set rs=nothing
     response.Redirect("dingdan.asp")
	 response.End	
	'Response.Write("<script>alert(""回复成功"");location.href=""dingdan.asp"";</'script>") 	
else
  if request("id") ="" then
     response.Write "<script LANGUAGE=javascript>alert('参数错误! ');history.go(-1);</script>"
     response.End
  else
   id=request("id")
   Sql = "Select * from Sbe_order where ID = "&id
   set rs=conn.execute(Sql)
   if rs.eof and rs.bof then
      response.Write "<script LANGUAGE=javascript>alert('参数错误! ');history.go(-1);</script>"
      response.End
    else
	  huiyuan=rs("huiyuan")  
	  username=rs("username")         '
	  usertel=rs("usertel")                 '           '
	  useremail=rs("useremail")             '
	  useraddress=rs("useraddress")
	  remarks=rs("remarks")
	  productname=rs("productname")
	  category=rs("category")
	  xinghao=rs("xinghao")           '
	  jiage=rs("jiage")           '
	  content=rs("content")          '
	  status1=rs("status")          '	
	  timing=rs("timing")
	  showtime=rs("showtime")
	  detail=rs("detail")
	  if rs("status")=1 then
	     a="disabled"
	   else
         a=""
       end if
     end if
   rs.close
   set rs=nothing
  end if
end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>友情链接</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript>
// 检测浏览器
NS4 = document.layers && true;
IE4 = document.all && parseInt(navigator.appVersion) >= 4;
</script>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="19%" height="25"><font color="#6A859D">订单管理 &gt;&gt;订单处理</font></td>   
      <td width="81%">&nbsp;       </td>   
  </tr>
  <tr> 
    <td height="1" colspan="2" background="../images/dot.gif"></td>
  </tr>
</table>
<table width="60%" border="0" align="center" cellpadding="0" cellspacing="0"  id="sbe_table">
                <form name=form method=post onSubmit="return checked();" action="dd_show.asp?act=add">
				 <tr align="center"> 
                    <td height="30" colspan="2" bgcolor="#EFEFEF" class="sbe_table_title">订单管理 >> 订单处理</td>
                  </tr>
				  	<tr> 
                    <td class=M align="right" bgcolor="#EFEFEF"><strong>产品信息</strong>：</td>
                    <td>               </td>
                  </tr>
				  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">产品类别：</td>
                    <td>&nbsp;<%=category%>                 </td>
                  </tr>
                 <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">产品名称：</td>
                    <td>&nbsp;<%=productname%></td>
                  </tr>

                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">产品编号：</td>
                    <td>&nbsp;<%=xinghao%></td>
                  </tr>                 
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF" width="120">产品规格：</td>
                    <td>&nbsp;<%=jiage%>
					</td>
                  </tr>
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF" width="120">产品包装：</td>
                    <td>&nbsp;<%=detail%>
					</td>
                  </tr>
<!--                  <tr style="display:none"> 
                    <td class=M bgcolor="#EFEFEF" align="right">QQ/MSN：</td>
                    <td><input name="URL2" type="text" id="URL2" value="<%=lyqq%>" size="40" readonly=""></td>
                  </tr>-->
<tr > 
                    <td class=M bgcolor="#EFEFEF" align="right">产品内容：</td>
                    <td class=M>&nbsp;<%=content%></td>
                  </tr>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right"><strong>客户资料</strong>：</td>
                    <td></td>
                  </tr>
				  <%if trim(huiyuan)<>"" then%>
				  <tr style="display:none;"> 
                    <td class=M bgcolor="#EFEFEF" align="right">会员帐号：</td>
                    <td> &nbsp;<%=huiyuan%>
                    </td>
                  </tr>
				  <%end if%>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">客户姓名：</td>
                    <td> &nbsp;<%=username%>
                    </td>
                  </tr>
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">客户电话：</td>
                    <td>&nbsp;<%=usertel%>
                    </td>
                  </tr>
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">客户Email：</td>
                    <td>&nbsp;<%=useremail%>
                    </td>
                  </tr>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">客户地址：</td>
                    <td>&nbsp;<%=useraddress%>
                    </td>
                  </tr>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">订单内容：</td>
                    <td> 
                      &nbsp;<%=remarks%>
                    </td>
                  </tr>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">是否处理：</td>
                    <td>&nbsp;<input name="flag"  type="radio" class="input" value="1" <%if trim(status1)=1 then response.Write("checked") end if%>> 
                    处理&nbsp;&nbsp;
                    <input name="flag" type="radio" class="input" value="0" <%if trim(status1)=0 then response.write("checked") end if%> <%=a%>>
                    暂不处理
                    </td>
                  </tr>
                  <tr> 
                    <td align="right" bgcolor="#EFEFEF" class=M style="height:30px;">　</td>
                    <td> 
                      <input name="submit" type="submit" class="sbe_button" value=" 确 定 ">
                      &nbsp;
                      <input name="submit2" type="hidden" class="sbe_button" value=" 重 置 ">
                      &nbsp;
                      <input type="hidden" name="id" value=<%=id%>>
                    </td>
                  </tr>
                </form>
</table>
<table width="100" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="20">&nbsp;</td>
  </tr>
</table>

<Script Language="JavaScript">
	<!--
 function checked(){
//  if(document.form.hftheme.value == ""){
//   alert("回复主题不能为空!");
//   document.form.hftheme.focus();
//   return false;
//  }
 // if(document.form.hfremark.value == ""){
//   alert("回复内容不能为空!");
//   document.form.hfremark.focus();
//   return false;
//  }
     //if(confirm('Do you add this order,If you add,You will unEdit?')){
  // return true;}
   //{
   //return false;
   //}
//return true;
}
   // -->
	</Script>
<%Call CloseDataBase()%>
</body>
</html>
