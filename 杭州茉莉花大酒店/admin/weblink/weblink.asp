<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%Call OpenData()
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('网络超时，或者你还没有登陆! ');this.location.href='../login.asp';</script>"
end if
  IF instr(webConfig,", 9")=0 or instr(session("manconfig"),", 9")=0 Then'网站功能配置
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('您的权限不足，请不要非法调用其它管理模块，否则您的账号将被系统自动删除!');this.location.href='../login.asp';</Script>"
Response.end
end if
act=Request("act")
linkid=Request("id")
IF act="add" Then
'增加新的友情链接 
  set rs=server.createobject("adodb.recordset")  
  IF linkid="" Then
     set rs_max=server.CreateObject("adodb.recordset")
     sql="select max(orderid) as maxid from Sbe_Weblink"
     rs_max.open sql,conn,1,1
     if isnull(rs_max("maxid")) then
        sequence=1
     else
        sequence=rs_max("maxid")+1
     end if
     rs_max.close
     set rs_max=nothing	
     sql="select * from Sbe_Weblink order by id desc"
     rs.open sql,conn,1,3	   
     rs.addnew
     rs("orderid")=sequence
     msg="新增楼盘标志成功"   
  Else
    msg="链接楼盘标志成功"
    sql="select * from Sbe_Weblink where id =" &linkid
	rs.open sql,conn,1,3
  End IF
    companyname=request.form("companyName") 
    url=request.form("url")
    url=replace(url,"http://","")
    linktype=request.form("linktype")
    picurl=request.form("realpicname")
    linkman=request.form("linkman")
    PhoneNumber=request.form("PhoneNumber")
    FaxNumber=request.form("FaxNumber")
    email=request.form("email")
    remark=request.form("remark") 
	leibie=trim(request.form("leibie"))
    rs("companyname")=companyname
    rs("URL")=url
    rs("linktype")=linktype
    rs("picurl")=picurl
    rs("linkman")=linkman
    rs("phone")=phoneNumber
    rs("fax")=faxNumber & " "
    rs("email")=email & " "
    rs("remark")=remark & " "
    rs("status")=1
    rs("posttime")=now
    rs("leibie")=leibie	 
    rs.update
    rs.close
   set rs=nothing
   Call MessageBoxOK(msg)
ElseIF len(linkid)>0 and act="modify" Then
    Dim strSQL,objRec
	Set objRec=Server.Createobject("adodb.recordset")
	strSQL="select * from Sbe_Weblink where id=" &linkid
	objRec.Open strSQL,conn,1,1
	With objRec
	 IF .Eof And .Bof Then
	   companyName=""
       URL=""
       linktype=false
       realpicname=""
       linkman=""
       phoneNumber=""
       faxNumber=""
       email=""
       remark=""
	 Else
	   companyName=objRec("companyname")
       URL=objRec("URL")
       linktype=objRec("linktype")
       realpicname=objRec("picurl")
       linkman=objRec("linkman")
       phoneNumber=objRec("phone")
       faxNumber=objRec("fax")
       email=objRec("email")
       remark=objRec("remark")
	   leibie=objRec("leibie")
	 End IF	 
	End With
	objRec.Close:Set objRec=Nothing
webname="修改楼盘标志"
if weblink_leibie=1 then linktype=true
else
webname="增加楼盘标志"
if weblink_leibie=1 then linktype=true
   leibie=1
  end if
Private Sub MessageBoxOK(strValue)

	With Response
		.Write "<script>" & vbcrlf
		.Write "alert('"+strValue+"');" & vbcrlf
		.Write "this.location.href='list.asp'" & vbcrlf
		.Write "</script>" & vbcrlf
	End With
End Sub
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=webname%></title>
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
function checkform(theform){
  if(theform.companyName.value==""){
    alert("请填写网站名称!");
	theform.companyName.focus();
	return false;
  }
 
  if(theform.URL.value==""){
    alert("请填写网址!");
	theform.URL.focus();
	return false;
  }
  if(theform.linkman.value==""){
    alert("请填写联系人!");
	theform.linkman.focus();
	return false;
  }
  if(theform.linkman.value==""){
    alert("请填写联系人!");
	theform.linkman.focus();
	return false;
  }
   if(theform.phoneNumber.value==""){
    alert("请填写联系电话!");
	theform.phoneNumber.focus();
	return false;
  }

}
// 检测浏览器
NS4 = document.layers && true;
IE4 = document.all && parseInt(navigator.appVersion) >= 4;

// 选择指定的tab.
function selectTab(tab) {
    var form   = document.tabform;
    var TabLayer1 = getLayerStyle("TabLayer1");
    var TabLayer2 = getLayerStyle("TabLayer2");

    if (tab == "TabLayer2") {
        _showLayer(TabLayer1, false);
        _showLayer(TabLayer2, true);


    } else {
        _showLayer(TabLayer2, false);
        _showLayer(TabLayer1, true);

    }
    return true;
}

function _showLayer(layer, display) {
    if (layer) {
        if (display) {
            layer.display = "block";
        } else {
            layer.display = "none";
        }
    }
}

// 取得指定id的layer
function getLayerStyle(id) {
    if (IE4 && document.all(id)) {
        return document.all(id).style;
    } else if (NS4 && document.layers[id]) {
        return document.layers[id];
    } else {
        return null;
    }
}
</SCRIPT>
<script language="JavaScript" src="../include/meizzDate.js"></script>
</head>

<body><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="19%" height="25"><font color="#6A859D">楼盘标志 &gt;&gt; <%'=webname%>楼盘标志 </font></td>   
      <td width="81%">&nbsp;       </td>   
  </tr>
  <tr> 
    <td height="1" colspan="2" background="../images/dot.gif"></td>
  </tr>
</table>
<table width="85%" border="0" align="center" cellpadding="0" cellspacing="0"  id="sbe_table">
                <form name=form method=post onSubmit="return checkform(this)" action="weblink.asp?act=add">
				 <tr align="center"> 
                    <td height="30" colspan="2" bgcolor="#EFEFEF" class="sbe_table_title">楼盘标志　                    </td>
                  </tr>
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF"><a name="1"></a>名&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;称：</td>
                    <td> 
                      <input name="companyName" type="text" id="companyName" value="<%=companyName%>" size="40"><font color="#FF6600">*</font>
                    </td>
                  </tr>                 
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">链接地址：</td>
                    <td> 
                      <input name="URL" type="text" id="URL" value="<%=URL%>" size="40"><font color="#FF6600">*</font>
                    </td>
                  </tr>
                  <tr <%=banben_display%>> 
                    <td class=M bgcolor="#EFEFEF" align="right">所属类别：</td>
                    <td> 
                      <input type="radio" name="leibie" id="leibie" value="1" <%=ReturnSel(1,leibie,2)%>>
                      中文链接 
                      <input type="radio" name="leibie" id="leibie" value="2" <%=ReturnSel(2,leibie,2)%>>
                      英文链接</td>
                  </tr>
                  <tr <%=weblink_display1%>>
                    <td class=M bgcolor="#EFEFEF" align="right">链接类型：</td>
                    <td> 
                      <input type="radio" name="linktype" id="linktype" value="0" <%=ReturnSel(False,linktype,2)%> onClick="selectTab('TabLayer2');">
                      文字链接 
                      <input type="radio" name="linktype" id="linktype" value="1" <%=ReturnSel(true,linktype,2)%>  onClick="selectTab('TabLayer1');">
                      图片链接</td>
                  </tr>
                  <tr align="center" bgcolor="#FFFFCC" id="TabLayer1" <%IF linktype=True Then%>style="display:block;"<%Else%>style="display:none;"<%End IF%>> 
                    <td colspan="2" bgcolor="#D8E4F1" class=M> 
                      <div>
					    <table width="100%"  border="0" cellspacing="0" cellpadding="0" style="border:none; ">
                          <tr>
                            <td width="15%" align="right" style="border:none; ">图片：                            </td>
                            <td width="24%" align="right" style="border:none; "><input name="realpicname" type="text" value="<%=realpicname%>" size="13" readonly></td>
                            <td width="61%" align="left" style="border:none; "><iframe src="../upload/upload.asp?Form_Name=form&UploadFile=realpicname" width="100%" height="25" frameborder="0" scrolling="no"></iframe></td>
                          </tr>
                        </table>
                        <table width="58%" border="0" cellpadding="0" cellspacing="0" style="border:none; ">
                          <tr> 
                            <td style="border:none; "><img src="picture/weblink.jpg" alt="Preview" name="imagePreview" width=88 height="31" border=0 align=middle></td>
                            <td style="border:none; ">56x49像素<br>
                            图片大小不能超过<font color="#FF6600">200</font>K!</td>
                          </tr>
                        </table>
                      </div>
                    </td>
                  </tr>
                  <!--<tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">联&nbsp;&nbsp;系&nbsp;人：</td>
                    <td class=M> 
                      <input name="linkman" type="text" id="linkman" value="<%=linkman%>" size="40"><font color="#FF6600">*</font>
                    </td>
                  </tr>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">联系电话：</td>
                    <td> 
                      <input name="phoneNumber" type="text" id="PhoneNumber" value="<%=phoneNumber%>" size="20" maxlength="32"><font color="#FF6600">*</font>
                    </td>
                  </tr>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">传&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;真：</td>
                    <td> 
                      <input name="faxNumber" type="text" id="FaxNumber" value="<%=faxNumber%>" size="20" maxlength="32">
                    </td>
                  </tr>
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">电子信箱：</td>
                    <td> 
                      <input name="email" type="text" id="email" value="<%=email%>" size="40">
                    </td>
                  </tr>
                  <tr> 
                    <td class=M bgcolor="#EFEFEF" align="right">备&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;注：</td>
                    <td> 
                      <textarea name="remark" id="remark" cols="46" rows="6"><%=remark%></textarea>
                    </td>
                  </tr>-->
                  <tr> 
                    <td class=M align="right" bgcolor="#EFEFEF">　</td>
                    <td> 
                      <input name="submit" type="submit" class="sbe_button" value=" 确 定 ">
                      <input type="hidden" name="id" value=<%=linkid%>>
                    </td>
                  </tr>
                </form>
</table>
</body>
</html>
<%
Private Sub news_come_Class()
'读取资讯来源
 Set oRs=Conn.Execute("select * from news_come_class order by id asc")
 IF oRs.Eof and oRs.bof Then Exit Sub
 Do While not oRs.eof 
  response.write "<a href=""javascript:clk('"& oRs("title") &"');"" >"& oRs("title") &"</a>/"& vbCrLf
 oRs.Movenext
 Loop
End Sub
%>