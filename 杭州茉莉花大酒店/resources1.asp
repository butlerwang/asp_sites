<!-- #include file="inc/conn.asp"-->
<!-- #include file="Check_Sql.asp"-->
<!-- #include file="inc/lib.asp"-->
<%OpenData()%>
<%set rs=server.CreateObject("adodb.recordset")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>������Ƹ</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<link href="css/css.css" rel="stylesheet" type="text/css" />
</head>
			<%if request("action")="add" then
			sql="select * from sbe_resume"
			rs.open sql,conn,1,3
			rs.addnew
			rs("job")=request("jobid")
			rs("RealName")=trim(request("username"))
			rs("Birthday")=trim(request("csrq"))
			rs("Education")=trim(request("xueli"))
			rs("Profession")=trim(request("zhuanye"))
			rs("bysj")=trim(request("bysj"))
			rs("hj")=trim(request("hj"))
			rs("aihao")=trim(request("aihao"))
			rs("jk")=trim(request("jk"))
			rs("fuqin")=trim(request("fuqin"))
			rs("muqin")=trim(request("muqin"))
			rs("content")=trim(request("content"))
			rs.update
			rs.close%>
			<form name=reDirectURL action=resources1.asp?id=<%=request("id")%> method=post></form>
                <script language="javascript">
	                alert("ӦƸ��Ϣ��ӳɹ�");
                    document.reDirectURL.submit();
           </script>
		   <%end if%>
		   	<script language="javascript">
function CheckForm()
{

	if (document.myform.username.value=="") {
		alert("����û����д.");
		document.myform.username.focus();
		return false;
	}
		if (document.myform.csrq.value=="") {
		alert("��������û����д");
		document.myform.csrq.focus();
		return false;
	}
		if (document.myform.zhuanye.value=="") {
		alert("רҵû����д");
		document.myform.zhuanye.focus();
		return false;
	}
		if (document.myform.bysj.value=="") {
		alert("��ҵʱ��û����д");
		document.myform.bysj.focus();
		return false;
	}
		if (document.myform.xueli.value=="") {
		alert("ѧλû����д");
		document.myform.xueli.focus();
		return false;
	}
		if (document.myform.jk.value=="") {
		alert("����״��û����д");
		document.myform.jk.focus();
		return false;
	}
		if (document.myform.content.value=="") {
		alert("����û����д");
		document.myform.content.focus();
		return false;
	}
				//    if (document.fqorm.email.value.length !=0){
	//  if 
    //     ((document.fqorm.email.value).indexOf("@")<1||(document.fqorm.email.value).indexOf(".")<1)
    //    {alert("������������!");
		//document.fqorm.email.focus();
		//return false;}

	//} 

	return true;
}
</script>
<body>
<table width="810" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="9"><img  src="images/new1_01.jpg" width="9" height="152"></td>
        <td width="631" background="images/new1_02.jpg">&nbsp;</td>
        <td width="170"><img  src="images/new1_04.jpg" width="170" height="152" border="0" usemap="#Map" ></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFF9D7"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr valign="top">
        <td width="640" height="200" align="left" valign="top">
		<%if request("action")="yp" then%>
		
		<form name="myform"  onSubmit="return CheckForm();" method="post" action="?action=add&id=<%=request("id")%>" style="margin:0px">
            <table border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="88" rowspan="14">&nbsp;</td>
              <td height="18">&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
			<tr>
              <td width="99" height="25" align="right" class="szise3">ӦƸ��λ��</td>
              <td width="392" style="padding-left:4px">
			 
			  <%sql="select id,job from sbe_job where id="&request("id")&""
			  rs.open sql,conn,1,1
			  if not rs.eof then
			 %>
			 <%=rs(1)%>
			 <input type="hidden" name="jobid" value="<%=rs(0)%>">
			  <%
			  end if
			  rs.close%>
                </select>
                <span class="szise3">*</span></td>
            </tr>
            <tr>
              <td height="25" align="right" class="szise3">������</td>
              <td><input name="username" type="text" class="biaodan" id="username" size="32">                
                <span class="szise3">*</span></td>
            </tr>
            
            <tr>
              <td height="25" align="right" class="szise3">�������£�</td>
              <td><input name="csrq" type="text" class="biaodan" id="csrq" size="32">
                <span class="szise3">*</span></td>
            </tr>
            <tr>
              <td height="25" align="right" class="szise3">��ѧרҵ��</td>
              <td><input name="zhuanye" type="text" class="biaodan" size="32">
                <span class="szise3">*</span></td>
            </tr>
            <tr>
              <td height="25" align="right" class="szise3">��ҵʱ�䣺</td>
              <td><input name="bysj" type="text" class="biaodan" id="bysj" size="32">
                <span class="szise3">*</span></td>
              </tr>
            <tr>
              <td height="25" align="right" class="szise3">ѧλ��</td>
              <td><input name="xueli" type="text" class="biaodan" id="xueli" size="32">
                <span class="szise3">*</span></td>
            </tr>
            <tr>
              <td height="25" align="right" class="szise3">�������</td>
              <td><input name="hj" type="text" class="biaodan" id="hj" size="52"></td>
            </tr>
            
            <tr>
              <td height="25" align="right" class="szise3">��Ȥ���ã�</td>
              <td><input name="aihao" type="text" class="biaodan" id="aihao" size="32"></td>
            </tr>
            
            <tr>
              <td height="25" align="right" class="szise3">����״����</td>
              <td><input name="jk" type="text" class="biaodan" id="jk" size="32">
                  <span class="szise3">*</span></td>
              </tr>
            <tr>
              <td height="25" align="right" class="szise3">����ְҵ��</td>
              <td><input name="fuqin" type="text" class="biaodan" id="fuqin" size="32"></td>
            </tr>
            <tr>
              <td height="25" align="right" class="szise3">ĸ��ְҵ��</td>
              <td><input name="muqin" type="text" class="biaodan" id="muqin" size="32"></td>
            </tr>
            
            
            <tr>
              <td height="76" align="right" class="szise3">����������</td>
              <td valign="middle"><textarea name="content" cols="50" rows="5" class="biaodan" id="jinli"></textarea>
                <span class="szise3">*</span></td>
            </tr>
            <tr>
              <td height="35">&nbsp;</td>
              <td>&nbsp;
                <input type="submit" name="Submit" value="�ύ">
                &nbsp;
                <input type="reset" name="Submit2" value="����">                
                &nbsp;<span class="szise3">��*Ϊ������</span></td>
            </tr>
          </table>
		  </form>
		<%else%>
			<%
		           
			sql="select Department,job,Num,AddDate,EffectTime,Content,Other,id from sbe_job where id="&request("id")&" order by id desc"
			rs.open sql,conn,1,1
			if not rs.eof then%>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="3%" height="100" rowspan="2" valign="bottom">&nbsp;</td>
            <td width="73%" height="50" valign="bottom"><strong class="ziti7">��Ƹ<%=rs(1)%>�� </strong></td>
            <td width="24%" rowspan="2" valign="bottom" class="ziti2">����ʱ�䣺<%=rs(3)%></td>
          </tr>
          <tr>
            <td height="5" valign="bottom"></td>
          </tr>
          <tr>
            <td colspan="3" align="center" valign="middle">
			<table width="90%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5F4E07">
              <tr>
                <td width="102" align="center" valign="middle" bgcolor="#FFF9D7" class="ziti5"> ��Ƹ���ţ�</td>
                <td width="210" bgcolor="#FFF9D7" class="box" style="padding-left:7px"><%=rs(0)%></td>
                <td width="99" align="center" valign="middle" bgcolor="#FFF9D7" class="box1"><span class="ziti5">��Ƹ��λ��</span></td>
                <td width="160" bgcolor="#FFF9D7" class="box" style="padding-left:7px"><%=rs(1)%></td>
              </tr>
              <tr>
                <td align="center" valign="middle" bgcolor="#FFF9D7" class="ziti5">��Ƹ������</td>
                <td bgcolor="#FFF9D7" class="box" style="padding-left:7px"><%=rs(2)%></td>
                <td align="center" valign="middle" bgcolor="#FFF9D7" class="ziti5">��Чʱ�䣺</td>
                <td bgcolor="#FFF9D7" class="box" style="padding-left:7px"><%=rs(3)%>--<%=rs(4)%></td>
              </tr>
              <tr>
                <td align="center" valign="middle" bgcolor="#FFF9D7" class="box1"><span class="ziti5">��ƸҪ��</span></td>
                <td height="70" colspan="3" bgcolor="#FFF9D7" style="padding-left:7px"><%=HTMLcode(rs(5))%></td>
                </tr>
              <tr>
                <td align="center" valign="middle" bgcolor="#FFF9D7" class="ziti5">��ش�����</td>
                <td height="70" colspan="3" bgcolor="#FFF9D7" style="padding-left:7px"><%=HTMLcode(rs(6))%></td>
              </tr>
            </table>
			<table width="90%"><tr><td align="right" style="padding-right:12px"><a href="?action=yp&id=<%=rs(7)%>">��ҪӦƸ ��</a></td></tr></table>
						<%
						end if
						rs.close%>
			</td>
          </tr>
        </table>
		<%end if%>
		</td>
        <td width="69"><img id="new1_06" src="images/new1_06.jpg" width="69" height="115" alt="" /></td>
        <td width="101">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><img id="new1_09" src="images/new1_09.jpg" width="810" height="7" alt="" /></td>
  </tr>
  <tr>
    <td height="10" valign="top" bgcolor="#F5E9C3" class="ziti7">&nbsp;</td>
  </tr>
</table>

<map name="Map"><area shape="rect" coords="145,3,169,25" href="#" onClick="javascript:window.close();"></map></body>
</html>
