<!-- #include file="inc/conn.asp"-->
<!-- #include file="Check_Sql.asp"-->
<!-- #include file="inc/lib.asp"-->
<%OpenData()%>
<%set rs=server.CreateObject("adodb.recordset")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>杭州茉莉花大酒店</title>
<%if request("action")="add" then
sql="select * from sbe_order"
rs.open sql,conn,2,2
rs.addnew
rs("kssj")=trim(request("kssj"))
rs("lksj")=trim(request("lksj"))
rs("chengren")=trim(request("chengren"))
rs("ertong")=trim(request("ertong"))
rs("roomtype")=trim(request("roomtype"))
rs("roomnum")=trim(request("roomnum"))
rs("peoplenum")=trim(request("peoplenum"))
rs("tianshu")=trim(request("tianshu"))
rs("lasttime")=trim(request("lasttime"))
rs("other")=trim(request("other"))
rs("username")=trim(request("username"))
rs("zhengjian")=trim(request("zhengjian"))
rs("zhengjiannum")=trim(request("zhengjiannum"))
rs("tel")=trim(request("tel"))
rs("handphone")=trim(request("handphone"))
rs.update%>
<form name=reDirectURL action=online.asp method=post></form>
<script language="javascript">
	 alert("您的预定添加成功，我们管理员将和您取得联系");
     document.reDirectURL.submit();
</script>
<%end if
rs.close%>
	<script language="javascript">
function CheckForm()
{

	if (document.myform.username.value=="") {
		alert("联系人没有填写.");
		document.myform.username.focus();
		return false;
	}
		if (document.myform.zhengjian.value=="") {
		alert("证件没有填写");
		document.myform.zhengjian.focus();
		return false;
	}
		if (document.myform.zhengjiannum.value=="") {
		alert("证件号码没有填写");
		document.myform.zhengjiannum.focus();
		return false;
	}
		if (document.myform.tel.value=="") {
		alert("联系电话没有填写");
		document.myform.tel.focus();
		return false;
	}
	return true;
}
</script>

<link href="css/css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="1003" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF7E6">
  <tr>
    <td width="24%" valign="top"><table width="77%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img id="online_01" src="images/online_01.jpg" width="239" height="14" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_03" src="images/online_03.jpg" width="239" height="15" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_04" src="images/online_04.jpg" width="239" height="14" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_05" src="images/online_05.jpg" width="239" height="15" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_06" src="images/online_06.jpg" width="239" height="14" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_10" src="images/online_10.jpg" width="239" height="14" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_11" src="images/online_11.jpg" width="239" height="15" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_12" src="images/online_12.jpg" width="239" height="14" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_15" src="images/online_15.jpg" width="239" height="15" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_16" src="images/online_16.jpg" width="239" height="14" alt="" /></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="4%" background="images/online_21.jpg"><img id="online_18" src="images/online_18.jpg" width="5" height="58" alt="" /></td>
            <td width="93%" align="center" valign="middle" background="images/online_21.jpg"><table width="105%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><div align="center"><img id="online_20" src="images/online_20.jpg" width="172" height="27" alt="" /></div></td>
              </tr>
              <tr>
                <td><div align="center"><img id="online_24" src="images/online_24.jpg" width="172" height="31" alt="" /></div></td>
              </tr>
            </table></td>
            <td width="3%"><img id="online_23" src="images/online_23.jpg" width="7" height="58" alt="" /></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><img id="online_25" src="images/online_25.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_26" src="images/online_26.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_27" src="images/online_27.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_28" src="images/online_28.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_29" src="images/online_29.jpg" width="239" height="19" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_30" src="images/online_30.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_31" src="images/online_31.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_32" src="images/online_32.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_33" src="images/online_33.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_34" src="images/online_34.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_35" src="images/online_35.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_38" src="images/online_38.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_39" src="images/online_39.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_40" src="images/online_40.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_41" src="images/online_41.jpg" width="239" height="19" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_42" src="images/online_42.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_47" src="images/online_47.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_49" src="images/online_49.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_50" src="images/online_50.jpg" width="239" height="20" alt="" /></td>
      </tr>
      <tr>
        <td><img id="online_51" src="images/online_51.jpg" width="239" height="20" alt="" /></td>
      </tr>
    </table></td>
    <td width="76%" valign="top" bgcolor="#F5E9C3">
	<form name="myform" action="?action=add" method="post"  onSubmit="return CheckForm();">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="58" bgcolor="#000000">&nbsp;</td>
          </tr>
          <tr>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="54" bgcolor="#000000">&nbsp;</td>
                <td width="417" bgcolor="#000000"><img id="online_08" src="images/online_08.jpg" width="417" height="43" alt="" /></td>
                <td width="292" bgcolor="#000000">&nbsp;</td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td height="30" background="images/online_13.jpg">&nbsp;</td>
          </tr>
          <tr>
            <td height="283" valign="top" bgcolor="#FFF9D7">
              <table width="80%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    
                  </tr>
                  <tr>
                    <td colspan="6" bgcolor="#FFF9D7" class="notice"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;茉莉花大酒店客房预订</td>
                  </tr>
                  <tr>
                    <td width="40%" class="notice1"> <div align="right">入店时间：                  </div></td>
                    <td colspan="5" valign="middle" class="notice1"><table width="88%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="59%" align="left"><input name="kssj" type="text" class="f2" id="kssj" /></td>
                        <td width="41%"><span class="STYLE1">(格式：YYYY-MM-DD)</span></td>
                      </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td class="notice1"><div align="right">离店时间：</div></td>
                    <td colspan="5" class="notice1"><table width="88%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="59%" align="left"><input name="lksj" type="text" class="f2" id="lksj" /></td>
                        <td width="41%"><span class="STYLE1">(格式：YYYY-MM-DD)</span></td>
                      </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td class="notice1"><div align="right">客人类型：</div></td>
                    <td width="10%" class="STYLE1">成人</td>
                    <td width="20%" align="left" class="notice1"><select name="chengren">
					<option value="1">1</option>
					<option value="2">2</option>
					<option value="3">3</option>
					<option value="4">4</option>
					<option value="5">5</option>
					<option value="6">6</option>
					<option value="7">7</option>
					<option value="8">8</option>
					<option value="9">9</option>
					<option value="10">10</option>
					<option value="十人以上">十人以上</option>
					<option value="二十人以上">二十人以上</option>
					<option value="三十人以上">三十人以上</option>
					<option value="四十人以上">四十人以上</option>
					<option value="五十人以上">五十人以上</option>
					<option value="更多">更多</option>
					
                    </select>
                    </td>
                    <td width="13%" class="STYLE1">儿童</td>
                    <td width="18%" class="STYLE1"><span class="notice1">
                     <select name="ertong">
					<option value="1">1</option>
					<option value="2">2</option>
					<option value="3">3</option>
					<option value="4">4</option>
					<option value="5">5</option>
					<option value="6">6</option>
					<option value="7">7</option>
					<option value="8">8</option>
					<option value="9">9</option>
					<option value="10">10</option>
					<option value="十人以上">十人以上</option>
					<option value="二十人以上">二十人以上</option>
					<option value="三十人以上">三十人以上</option>
					<option value="四十人以上">四十人以上</option>
					<option value="五十人以上">五十人以上</option>
					<option value="更多">更多</option>
                    </select>
                    </span></td>
                    <td width="25%" class="STYLE1">&nbsp;</td>
                  </tr>
                  <tr align="left">
                    <td class="notice1"><div align="right">房间类型：</div></td>
                    <td colspan="5" valign="middle" class="notice1">
					<select name="roomtype">
					<%sql="select id,classname from sbe_product_class order by sequence desc"
					rs.open sql,conn,1,1
					if not rs.eof then
					do while not rs.eof%>
					<option value="<%=rs(0)%>" <%if rs(0)=cint(request("idd")) then%> selected="selected"<%end if%>><%=rs(1)%></option>
					<%rs.movenext
					loop
					end if
					rs.close%>
					</select>
					</td>
                  </tr>
                  <tr>
                    <td class="notice1"><div align="right">房间数量：</div></td>
                    <td colspan="5" class="notice1"><table width="80%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="44%"><input name="roomnum" type="text" class="f2" id="roomnum" /></td>
                        <td width="56%"> (单位：间)</td>
                      </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td class="notice1"><div align="right">入住人数：</div></td>
                    <td colspan="5" class="notice1"><table width="80%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="44%"><input name="peoplenum" type="text" class="f2" id="peoplenum" /></td>
                        <td width="56%"> (单位：人)</td>
                      </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td class="notice1"><div align="right">入住天数：</div></td>
                    <td colspan="5" class="notice1"><table width="80%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="44%"><input name="tianshu" type="text" class="f2" id="tianshu" /></td>
                        <td width="56%"> (单位：天)</td>
                      </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td class="notice1"><div align="right">最晚到达时间：</div></td>
                    <td colspan="5" class="notice1"><table width="90%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="59%"><input name="lasttime" type="text" class="f2" id="lasttime" /></td>
                        <td width="41%">(格式：YYYY-MM-DD)</td>
                      </tr>
                    </table></td>
                  </tr>
              </table>
              <table width="80%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="35%" class="notice1"><div align="right">其它要求：</div></td>
                  <td width="82%" class="notice1"><input name="other" type="text" class="f2" id="other" /></td>
                </tr>
              </table>
           </td>
          </tr>
          <tr>
            <td><img id="online_36" src="images/online_36.jpg" width="764" height="4" alt="" /></td>
          </tr>
      </table>
      <table width="80%" border="0" cellspacing="0" cellpadding="0">
          <tr>
           
          </tr>
          <tr>
            <td width="34%" class="notice1"><div align="right">入住人代表姓名：</div></td>
            <td width="80%"><table width="80%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="59%"><input name="username" type="text" class="f" id="username" /></td>
                <td width="41%"><span class="notice1"> (*) </span></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td class="notice1"><div align="right">证件类型：</div></td>
            <td><table width="89%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="29%"><input name="zhengjian" type="text" class="f1" id="zhengjian" /></td>
                <td width="23%"><strong class="notice1">证件号码：</strong></td>
                <td width="48%"><input name="zhengjiannum" type="text" class="f1" id="zhengjiannum" /></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td class="notice1"><div align="right">联系电话：</div></td>
            <td><table width="89%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="59%"><input name="tel" type="text" class="f" id="tel" /></td>
                <td width="41%" class="STYLE1">  (格式：区号-电话号码)  </td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td class="notice1"><div align="right">手   机：</div></td>
            <td><table width="96%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="51%"><input name="handphone" type="text" class="f" id="handphone" /></td>
                <td width="49%"> (注：外地手机请在前面加零) </td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><input type="image" id="online_44" src="images/online_44.jpg" width="110" height="23" alt="" /> <input type="image" id="online_45" src="images/online_45.jpg" width="114" height="23"  name="imgBtn" onClick="return resetBtn(this.form);"></td>
          </tr>
      </table></form>
   </td>
  </tr>
</table>
<Script Language="JavaScript">
<!--
function resetBtn(fm){
   　　 fm.reset();
    　　return false;
　　}
   // -->
</Script>
</body>
</html>
