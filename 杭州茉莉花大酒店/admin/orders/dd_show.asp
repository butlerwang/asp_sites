<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%
Call OpenData()%>
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
     'Rs("showtime")= date()
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
kssj=rs("kssj")
lksj=rs("lksj")
chengren=rs("chengren")
ertong=rs("ertong")
roomtype=rs("roomtype")
roomnum=rs("roomnum")
peoplenum=rs("peoplenum")
tianshu=rs("tianshu")
lasttime=rs("lasttime")
other=rs("other")
username=rs("username")
zhengjian=rs("zhengjian")
zhengjiannum=rs("zhengjiannum")
tel=rs("tel")
handphone=rs("handphone")
status1=rs("status")
	  set rs2=server.CreateObject("adodb.recordset")
	  sql="select * from sbe_product_class where id="&rs("roomtype")&""
	  rs2.open sql,conn,1,1
	  fang=rs2("classname")
	  rs2.close
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
                    <td class=M align="right" bgcolor="#EFEFEF">房间类型：</td>
                    <td>&nbsp;<%=fang%></td>
                  </tr>
                 <tr>
                   <td class=M align="right" bgcolor="#EFEFEF">入住时间：</td>
                   <td>&nbsp;<%=kssj%></td>
                 </tr>
                 <tr>
                   <td class=M align="right" bgcolor="#EFEFEF">离开时间：</td>
                   <td>&nbsp;<%=lksj%></td>
                 </tr>
                 <tr>
                   <td class=M align="right" bgcolor="#EFEFEF">最晚到达时间：</td>
                   <td>&nbsp;<%=lasttime%></td>
                 </tr>
                 <tr>
                   <td class=M align="right" bgcolor="#EFEFEF">需要房间数</td>
                   <td>&nbsp;<%=roomnum%></td>
                 </tr>
                 <tr>
                   <td class=M align="right" bgcolor="#EFEFEF">成人数：</td>
                   <td>&nbsp;<%=chengren%></td>
                 </tr>
                 <tr>
                   <td class=M align="right" bgcolor="#EFEFEF">儿童数：</td>
                   <td>&nbsp;<%=ertong%></td>
                 </tr>
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF" width="120">入住人数：</td>
                    <td>&nbsp;<%=peoplenum%>					</td>
                  </tr>


                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">入住天数：</td>
                    <td> &nbsp;<%=tianshu%>                    </td>
                  </tr>
                  <tr>
                    <td class=M bgcolor="#EFEFEF" align="right">姓名：</td>
                    <td><%=username%></td>
                  </tr>
                  <tr>
                    <td class=M bgcolor="#EFEFEF" align="right">电话：</td>
                    <td><%=tel%></td>
                  </tr>
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">手机：</td>
                    <td>&nbsp;<%=handphone%>                    </td>
                  </tr>
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">证件类型：</td>
                    <td>&nbsp;<%=zhengjian%>                    </td>
                  </tr>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">证件号码：</td>
                    <td>&nbsp;<%=zhengjiannum%>                    </td>
                  </tr>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">是否处理：</td>
                    <td>&nbsp;<input name="flag"  type="radio" class="input" value="1" <%if trim(status1)=1 then response.Write("checked") end if%>> 
                    处理&nbsp;&nbsp;
                    <input name="flag" type="radio" class="input" value="0" <%if trim(status1)=0 then response.write("checked") end if%> >
                    暂不处理                    </td>
                  </tr>
                  <tr> 
                    <td align="right" bgcolor="#EFEFEF" class=M style="height:30px;">　</td>
                    <td> 
                      <input name="submit" type="submit" class="sbe_button" value=" 确 定 ">
                      &nbsp;
                      <input name="submit2" type="hidden" class="sbe_button" value=" 重 置 ">
                      &nbsp;
                      <input type="hidden" name="id" value=<%=id%>>                    </td>
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
