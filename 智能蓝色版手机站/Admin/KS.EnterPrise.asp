<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_EnterPrise
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_EnterPrise
        Private KS,Param
		Private Action,i,strClass,sFileName,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		 With Response
					If Not KS.ReturnPowerResult(0, "KSMS10008") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
			  .Write "<html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""location.href='KS.Enterprise.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>企业管理</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.SpaceSkin.asp?flag=4';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>模板管理</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.EnterPrisePro.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/s.gif' border='0' align='absmiddle'>企业新闻</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.EnterPrisePro.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>企业产品</span></li>"
			  .Write "</ul>"
		End With
		
		maxperpage = 30 '###每页显示数
		If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
			Response.Write ("错误的系统参数!请输入整数")
			Response.End
		End If
		If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
			CurrentPage = CInt(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CInt(CurrentPage) = 0 Then CurrentPage = 1
		
		Param=" where 1=1"
		If KS.G("KeyWord")<>"" Then
		  If KS.G("condition")=1 Then
		   Param= Param & " and Companyname like '%" & KS.G("KeyWord") & "%'"
		  Else
		   Param= Param & " and a.username like '%" & KS.G("KeyWord") & "%'"
		  End If
		End If

		totalPut = Conn.Execute("Select Count(id) From KS_EnterPrise a " & Param)(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Select Case KS.G("action")
		 Case "Add","Edit" Call EnterPriseEdit()
		 Case "EditSave" Call DoSave()
		 Case "Del"  Call BlogDel()
		 Case "lock"  Call BlogLock()
		 Case "unlock"  Call BlogUnLock()
		 Case "verific"	  Call Blogverific()
		 Case "recommend"  Call Blogrecommend()
		 Case "Cancelrecommend" Call BlogCancelrecommend()
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td nowrap>公司名称</th>
	<td nowrap>创建者</th>
	<td nowrap>创建时间</th>
	<td nowrap>站点状态</th>
	<td nowrap>管理操作</th>
</tr>
<%
	sFileName = "KS.Enterprise.asp?"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select a.*,b.userid,b.blogname from KS_Enterprise a left join ks_blog b on a.username=b.username " & Param & " order by a.id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>没有用户申请开通企业空间！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action=?action=Del>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td class="splittd"><a href="../space/?<%=rs("username")%>" target="_blank"><%
	if ks.isnul(Rs("companyname")) then
	  response.write rs("blogname")
	else
	  response.write rs("companyname")
	end if%></a>
	<%if rs("recommend")="1" then response.write " <font color=red>荐</font>"%>
	</td>
	<td class="splittd" align="center"><%=Rs("username")%></td>
	<td class="splittd" align="center">&nbsp;<%=Rs("adddate")%>&nbsp;</td>
	<td class="splittd" align="center"><%
	select case rs("status")
	 case 0
	  response.write "未审"
	 case 1
	  response.write "<font color=red>已审</font>"
	 case 2
	  response.write "<font color=blue>锁定</font>"
	end select
	%></td>
	<td class="splittd" align="center"><a href="../space/?<%=rs("username")%>" target="_blank">浏览</a> <a href="?action=Edit&username=<%=RS("username")%>"  onclick="window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('企业&产品库 >> <font color=red>修改企业信息</font>')+'&ButtonSymbol=GOSave';">修改</a> <a href="?Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('确定删除该企业吗？'));">删除</a> <%IF rs("recommend")="1" then %><a href="?Action=Cancelrecommend&id=<%=rs("id")%>"><font color=red>取消推荐</font></a><%else%><a href="?Action=recommend&id=<%=rs("id")%>">设为推荐</a><%end if%>&nbsp;<%if rs("status")=0 then%><a href="?Action=verific&id=<%=rs("id")%>">审核</a> <%elseif rs("status")=1 then%><a href="?Action=lock&id=<%=rs("id")%>">锁定</a><%elseif rs("status")=2 then%><a href="?Action=unlock&id=<%=rs("id")%>">解锁</a><%end if%></td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class=Button type="submit" name="Submit2" value=" 删除选中的企业 " onclick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){this.document.selform.submit();return true;}return false;}">
	<input type="button" class="button" value="开通企业空间" onclick="window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('企业&产品库 >> <font color=red>开通企业空间</font>')+'&ButtonSymbol=GO';location.href='?action=Add';"/>
	</td>
</tr>
</form>
<tr>
	<td colspan=7>
	<%
	  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>
<div>
<form action="KS.EnterPrise.asp" name="myform" method="get">
   <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
      &nbsp;<strong>快速搜索=></strong>
	 &nbsp;关键字:<input type="text" class='textbox' name="keyword">&nbsp;条件:
	 <select name="condition">
	  <option value=1>公司名称</option>
	  <option value=2>用 户 名</option>
	 </select>
	  &nbsp;<input type="submit" value="开始搜索" class="button" name="s1">
	  </div>
</form>
</div>
<%
End Sub

Sub EnterPriseEdit()
Dim companyname,ClassID,BusinessLicense,LegalPeople,CompanyScale,RegisteredCapital,Province,City,ContactMan,Address
dim zipcode,telphone,Fax,weburl,BankAccount,AccountNumber,intro,SmallClassID,UserName,id,templateid,MapMarker
Dim Flag:Flag=KS.G("Flag")

username=KS.G("UserName")
	Dim RS:Set RS=server.createobject("adodb.recordset")
	RS.Open "Select top 1 * From KS_EnterPrise Where UserName='" & username &"'",conn,1,1
	 If RS.Eof And RS.Bof Then
	  RS.Close:Set RS=Nothing
      ClassID=0:SmallClassID=0:intro=" ":UserName=KS.G("UserName"):id=0:templateid=0
	 Else
	     id=rs("id")
		 username=rs("username")
		 companyname=rs("companyname")
		 ClassID=rs("ClassID") : SmallClassID=rs("SmallClassID")
		 BusinessLicense=rs("BusinessLicense")
		 LegalPeople=rs("LegalPeople")
		 CompanyScale=rs("CompanyScale")
		 RegisteredCapital=rs("RegisteredCapital")
		 Province=rs("Province")
		 City=rs("City") 
		 MapMarker=rs("MapMarker")
		 ContactMan=rs("ContactMan")
		 Address=rs("Address")
		 zipcode=rs("zipcode")
		 telphone=rs("telphone")
		 fax=rs("fax")
		 weburl=rs("weburl")
		 BankAccount=rs("BankAccount")
		 AccountNumber=rs("AccountNumber")
		 intro=rs("intro")
		 RS.Close 
		 set rs=conn.execute("select top 1 templateid from ks_blog where username='" & username & "'")
		 if not rs.eof then
		 	templateid=rs(0)
         end if
		 rs.close
		 Set RS=Nothing
    End If
%>
<script>
function CheckForm()
{
<%if request("action")="Add" then%>
if (document.myform.username.value=='')
{
 alert('请输入会员名称!');
 document.myform.username.focus();
 return false;
}
<%end if%>
if (document.myform.companyname.value=='')
{
 alert('请输入公司名称!');
 document.myform.companyname.focus();
 return false;
}
if ($("#templateid option:selected").val()=='0'){
 alert('请选择空间模板!');
 return false;
}
document.myform.submit();
}
</script>
<%If Flag<>"" Then%>
<div style='padding:10px;text-align:center;font-weight:bold'>升级用户<font color=red><%=username%></font>为企业空间</div>
<%end if
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
 <form name="myform" action="?action=EditSave" method="post">
   <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl">
      <%if request("action")="Add" then%>
	       <input type="hidden" value="0" name="id">
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>会员名称：</strong></td>
           <td height='28'>&nbsp;<input type='text' id='username' name='username' size="20"> <font color=red>*</font> 输入您要为其开通的会员名称</td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>空间LOGO：</strong></td>
           <td height='28'>&nbsp;<input type='text' id='logo' name='logo' size="40"> 支持jpg gif 或png格式 </td>
          </tr>  
	  <%else%>
   <input type="hidden" value="<%=id%>" name="id">
   <input type="hidden" value="<%=flag%>" name="flag">
   <input type="hidden" value="<%=username%>" name="username">
	  <%end if%>
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>公司名称：</strong></td>
           <td height='28'>&nbsp;<input type='text' id='companyname' name='companyname' value='<%=companyname%>' size="40"> <font color=red>*</font></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>绑定模板：</strong></td>
           <td height='28'>&nbsp;<select name="templateid" id='templateid'>
		   <option value='0'>--选择模板--</option>
		   <%set rs=conn.execute("select * from KS_BlogTemplate where flag=4 order by id desc")
		   do while not rs.eof
		      if trim(templateid)=trim(rs("id")) then
		      response.write "<option value='" & rs("id") & "' selected>" & rs("templatename") &"</option>"
			  else
		      response.write "<option value='" & rs("id") & "'>" & rs("templatename") &"</option>"
			  end if
		   rs.movenext
		   loop
		   rs.close
		   set rs=nothing
		   %>
		   </select></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>所属行业：</strong></td>
            <td height='28'>&nbsp;<%
		dim rss,sqls,count
		set rss=server.createobject("adodb.recordset")
		sqls = "select * from KS_enterpriseClass Where parentid<>0 order by orderid"
		rss.open sqls,conn,1,1
		%>
          <script language = "JavaScript">
		var onecount;
		subcat = new Array();
				<%
				count = 0
				do while not rss.eof 
				%>
		subcat[<%=count%>] = new Array("<%= trim(rss("id"))%>","<%=trim(rss("parentid"))%>","<%= trim(rss("classname"))%>");
				<%
				count = count + 1
				rss.movenext
				loop
				rss.close
				%>
		onecount=<%=count%>;
		function changelocation(locationid)
			{
			document.myform.SmallClassID.length = 0; 
			var locationid=locationid;
			var i;
			for (i=0;i < onecount; i++)
				{
					if (subcat[i][1] == locationid)
					{ 
						document.myform.SmallClassID.options[document.myform.SmallClassID.length] = new Option(subcat[i][2], subcat[i][0]);
					}        
				}
			}    
		
		</script>
		 <select class="face" name="ClassID" onChange="changelocation(document.myform.ClassID.options[document.myform.ClassID.selectedIndex].value)" size="1">
		<option value='0'>--请选择大类--</option>
		<% 
		dim rsb,sqlb
		set rsb=server.createobject("adodb.recordset")
		sqlb = "select * from ks_enterpriseclass where parentid=0 order by orderid"
        rsb.open sqlb,conn,1,1
		if rsb.eof and rsb.bof then
		else
		    do while not rsb.eof
					  If ClassID=rsb("id") then
					  %>
                    <option value="<%=trim(rsb("id"))%>" selected><%=trim(rsb("ClassName"))%></option>
                    <%else%>
                    <option value="<%=trim(rsb("id"))%>"><%=trim(rsb("ClassName"))%></option>
                    <%end if
		        rsb.movenext
    	    loop
		end if
        rsb.close
			%>
                  </select>
                  <font color=#ff6600>&nbsp;*</font>
                  <select class="face" name="SmallClassID">
				   <option value='0'>--请选择-</option>
                    <%dim rsss,sqlss
						set rsss=server.createobject("adodb.recordset")
						sqlss="select * from ks_enterpriseclass where parentid="& ks.chkclng(ClassID)&" order by orderid"
						rsss.open sqlss,conn,1,1
						if not(rsss.eof and rsss.bof) then
						do while not rsss.eof
							  if SmallClassID=rsss("id") then%>
							<option value="<%=rsss("id")%>" selected><%=rsss("ClassName")%></option>
							<%else%>
							<option value="<%=rsss("id")%>"><%=rsss("ClassName")%></option>
							<%end if
							rsss.movenext
						loop
					end if
					rsss.close
					%>
                </select></td>
          </tr> 
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>营业执照：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='BusinessLicense' value='<%=BusinessLicense%>' size="40"> <font color=red>*</font></td>
          </tr>  
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>企业法人：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='LegalPeople' value='<%=LegalPeople%>' size="40"> <font color=red>*</font></td>
          </tr>  
		  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
             <td  width='160' height='30' align='right' class='clefttitle'><span style="font-weight: bold">公司规模：</span></td>
              <td>&nbsp;
                              <select name="CompanyScale" id="CompanyScale">
							  <option value="1-20人"<%if CompanyScale="1-20人" then response.write " selected"%>>1-20人</option>
                      <option value="21-50人"<%if CompanyScale="21-50人" then response.write " selected"%>>21-50人</option>
                      <option value="51-100人"<%if CompanyScale="51-100人" then response.write " selected"%>>51-100人</option>
                      <option value="101-200人"<%if CompanyScale="101-200人" then response.write " selected"%>>101-200人</option>
                      <option value="201-500人"<%if CompanyScale="201-500人" then response.write " selected"%>>201-500人</option>
                      <option value="501-1000人"<%if CompanyScale="501-1000人" then response.write " selected"%>>501-1000人</option>
                      <option value="1000人以上"<%if CompanyScale="1000人以上" then response.write " selected"%>>1000人以上</option>
						    </select></td>
                          </tr>
                 <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
             <td  width='160' height='30' align='right' class='clefttitle'><span style="font-weight: bold">注册资金：</span></td>
                            <td>&nbsp;
							<select name="RegisteredCapital" id="RegisteredCapital">
							<option value="10万以下"<%if RegisteredCapital="10万以下" then response.write " selected"%>>10万以下</option>
                      <option value="10万-19万"<%if RegisteredCapital="10万-19万" then response.write " selected"%>>10万-19万</option>
                      <option value="20万-49万"<%if RegisteredCapital="20万-49万" then response.write " selected"%>>20万-49万</option>
                      <option value="50万-99万"<%if RegisteredCapital="50万-99万" then response.write " selected"%>>50万-99万</option>
                      <option value="100万-199万"<%if RegisteredCapital="100万-199万" then response.write " selected"%>>100万-199万</option>
                      <option value="200万-499万"<%if RegisteredCapital="200万-499万" then response.write " selected"%>>200万-499万</option>
                      <option value="500万-999万"<%if RegisteredCapital="500万-999万" then response.write " selected"%>>500万-999万</option>
                      <option value="1000万以上"<%if RegisteredCapital="1000万以上" then response.write " selected"%>>1000万以上</option>
					   </select></td>
                          </tr>
						  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                         <td  width='160' height='30' align='right' class='clefttitle'><span style="font-weight: bold">所在地区：</span></td>
                            <td>&nbsp;
                              <script src="../plus/area.asp" language="javascript"></script>
							  
							<script language="javascript">
							  <%if Province<>"" then%>
							  $('#Province').val('<%=Province%>');
							 <%end if%>
							  <%if City<>"" Then%>
							  $('#City')[0].options[1]=new Option('<%=City%>','<%=City%>');
							  $('#City')[0].options(1).selected=true;
							  <%end if%>
							</script>

					
							 </td>
                          </tr>
                     <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>电子地图：</strong></td>
                            <td>经纬度：<input value="<%=MapMarker%>" class="textbox" maxlength="255" type='text' name='MapMark' id='MapMark' /> <a href='javascript:void(0)' onclick='addMap()'> <img src='../user/images/edit_add.gif' align='absmiddle' border='0'>添加电子地图标志</a>
     <script src="../ks_inc/kesion.box.js"></script>
	 <script type="text/javascript">
		  function addMap(){
		  new KesionPopup().PopupCenterIframe('电子地图标注','../plus/baidumap.asp?MapMark='+escape($("#MapMark").val()),760,430,'auto');
		  }
		  </script><span class="msgtips">请选择贵公司所在的位置。</span>
							  </td>
                          </tr>
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>联 系 人：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='ContactMan' value='<%=ContactMan%>' size="40"> </td>
          </tr>  
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>公司地址：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='Address' value='<%=Address%>' size="40"></td>
          </tr>  
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>邮政编码：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='zipcode' value='<%=zipcode%>' size="40"></td>
          </tr>  
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>联系电话：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='Telphone' value='<%=telphone%>' size="40"></td>
          </tr>  
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>传真号码：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='Fax' value='<%=fax%>' size="40"></td>
          </tr>  
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>公司网站：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='weburl' value='<%=weburl%>' size="40"></td>
          </tr>  
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>开户银行：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='BankAccount' value='<%=BankAccount%>' size="40"></td>
          </tr>  
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>银行账号：</strong></td>
           <td height='28'>&nbsp;<input type='text' name='AccountNumber' value='<%=AccountNumber%>' size="40"></td>
          </tr>  
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>公司介绍：</strong></td>
           <td height='28'>
		   <textarea id="Intro" name="Intro"><%=KS.HTMLCode(Intro)%></textarea>
		   <script type="text/javascript" src="../editor/ckeditor.js" mce_src="../editor/ckeditor.js"></script>
		   <script type="text/javascript">
                CKEDITOR.replace('Intro', {width:"695",height:"300px",toolbar:"NewsTool",filebrowserBrowseUrl :"Include/SelectPic.asp?from=ckeditor&Currpath=<%=KS.GetUpFilesDir()%>",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
			</script>	
		   </td>
          </tr>  
      </table>
<%
End Sub

Sub DoSave()
       Dim ID:ID=KS.ChkClng(KS.G("id"))
	   Dim CompanyName:CompanyName=KS.LoseHtml(KS.S("CompanyName"))
	   Dim Province:Province=KS.S("Province")
	   Dim City:City=KS.S("City")
	   Dim Address:Address=KS.LoseHtml(KS.S("Address"))
	   Dim ZipCode:ZipCode=KS.LoseHtml(KS.S("ZipCode"))
	   Dim ContactMan:ContactMan=KS.LoseHtml(KS.S("ContactMan"))
	   Dim Telphone:TelPhone=KS.LoseHtml(KS.S("TelPhone"))
	   Dim Fax:Fax=KS.LoseHtml(KS.S("Fax"))
	   Dim WebUrl:WebUrl=KS.LoseHtml(KS.S("WebUrl"))
	   Dim Profession:Profession=KS.LoseHtml(KS.S("Profession"))
	   Dim CompanyScale:CompanyScale=KS.LoseHtml(KS.S("CompanyScale"))
	   Dim RegisteredCapital:RegisteredCapital=KS.LoseHtml(KS.S("RegisteredCapital"))
	   Dim LegalPeople:LegalPeople=KS.LoseHtml(KS.S("LegalPeople"))
	   Dim BankAccount:BankAccount=KS.LoseHtml(KS.S("BankAccount"))
	   Dim AccountNumber:AccountNumber=KS.LoseHtml(KS.S("AccountNumber"))
	   Dim BusinessLicense:BusinessLicense=KS.LoseHtml(KS.S("BusinessLicense"))
	   Dim ClassID:ClassID=KS.ChkClng(KS.G("ClassID"))
	   Dim SmallClassID:SmallClassID=KS.ChkClng(KS.G("SmallClassID"))
	   Dim Intro:Intro=Request.Form("Intro")
	   Dim UserName:UserName=KS.G("UserName")
	   
	   Dim RS,UserID: Set RS=Server.CreateObject("Adodb.RecordSet")
	   If ID=0 Then
	     If KS.IsNul(UserName) Then
		   KS.Die "<script>alert('请输入要开通企业空间的用户名!');history.back();</script>"
		 End If
	     RS.Open "Select top 1 UserID From KS_User Where UserName='" & UserName & "'",conn,1,1
		 If RS.Eof And RS.Bof Then
		   RS.Close : Set RS=Nothing
		   KS.Die "<script>alert('对不起，您输入的用户名不存在!');history.back();</script>"
		 End If
		 UserID=RS(0)
		 RS.Close
		 If Not Conn.Execute("Select top 1 UserName From KS_EnterPrise Where UserName='" & UserName & "'").eof Then
		   KS.Die "<script>alert('对不起，您输入的用户已开通过企业空间!');history.back();</script>"
		 End If
	   End If
	   
	   If KS.IsNul(CompanyName) Then Response.Write "<script>alert('公司名称必须输入');history.back();</script>":response.end
	   
	   If ID=0 Then
	     RS.Open "Select top 1 * From KS_Blog Where UserName='" & UserName & "'",conn,1,3
		 If RS.EOf And RS.Bof Then
		  RS.AddNew
		  RS("UserID")    = UserID
		  RS("UserName")  = UserName
		  RS("BlogName")  = CompanyName
		  RS("Logo")  = KS.G("Logo")
		  RS("TemplateID")= KS.ChkClng(KS.G("TemplateID"))
		  RS("ClassID") = KS.ChkClng(Conn.Execute("Select Top 1 ClassID From KS_BlogClass")(0))
		  RS("Announce")="暂无公告!"
		  RS("ContentLen")=500
		  RS("Recommend")=0
		  RS("Status")=1
		  RS.Update
		 End If
		 RS.Close
	   End If
            
		RS.Open "Select top 1 * From KS_Enterprise Where ID=" & ID,Conn,1,3
			  If RS.Eof Then
			    RS.AddNew
				 RS("AddDate")=NOW
				 RS("UserName")=UserName
				 RS("status")=1
			  End If
			     if KS.G("Flag")<>"" then
				 RS("status")=1
				 rs("username")=ks.g("username")
				 end if
				 RS("MapMarker")=KS.G("MapMark")
			     RS("CompanyName")=CompanyName
				 RS("Province")=Province
				 RS("City")=City
				 RS("Address")=Address
				 RS("ZipCode")=ZipCode
				 RS("ContactMan")=ContactMan
				 RS("Telphone")=Telphone
				 RS("Fax")=Fax
				 RS("WebUrl")=WebUrl
				 RS("ClassID")=ClassID
				 RS("SmallClassID")=SmallClassID
				 RS("CompanyScale")=CompanyScale
				 RS("RegisteredCapital")=RegisteredCapital
				 RS("LegalPeople")=LegalPeople
				 RS("BankAccount")=BankAccount
				 RS("AccountNumber")=AccountNumber
				 RS("BusinessLicense")=BusinessLicense
				 RS("Intro")=Intro
		 		 RS.Update
				
				 Dim FieldsXml:Set FieldsXml=LFCls.GetXMLFromFile("SpaceFields")
				If IsObject(FieldsXml) Then
					 Dim objNode,i,j,objAtr
					 Set objNode=FieldsXml.documentElement 
					 If objNode.Attributes.item(0).Text="2" Then
						 For i=0 to objNode.ChildNodes.length-1 
								set objAtr=objNode.ChildNodes.item(i) 
								on error resume next
								Conn.Execute("UPDATE KS_User Set " & objAtr.Attributes.item(1).Text & "='" & RS(objAtr.Attributes.item(0).Text) & "' Where UserName='" & RS("UserName") & "'")
						 Next
					 End If
			
				   End If
			     RS.Close
				 Set RS=Nothing
				 Conn.Execute("Update KS_EnterpriseNews Set BigClassID=" & ClassID &",SmallClassID=" & SmallClassID & " Where UserName='" & KS.G("UserName") & "'")
				 Conn.Execute("Update KS_Blog set TemplateID=" & KS.ChkClng(KS.G("TemplateID")) & " Where UserName='" & KS.G("UserName") &"'")
				 Conn.Execute("Update KS_User Set UserType=1 Where UserName='" & KS.G("UserName") & "'")
				 if KS.G("Flag")<>"" then
				 Response.Write "<script>alert('空间升级成功！');$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr='+escape('空间门户管理 >> <font color=red>企业空间管理</font>');location.href='KS.EnterPrise.asp';</script>"
				 ElseIf ID=0 Then
				 Response.Write "<script>alert('恭喜，企业空间创建成功！');$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr='+escape('空间门户管理 >> <font color=red>企业空间管理</font>');location.href='KS.EnterPrise.asp';</script>"
				 Else
				 Response.Write "<script>alert('企业基本信息资料更新成功！');$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr='+escape('空间门户管理 >> <font color=red>企业空间管理</font>');location.href='"& Request.Form("ComeUrl") & "';</script>"
				 End If

EnD Sub

'删除日志
Sub BlogDel()
 Dim ID:ID=KS.G("ID")
 Dim UserName,DefaultTemplateID
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 DefaultTemplateID=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=2 and IsDefault='true'")(0))
 Dim RS:Set RS=Server.CreateOBject("ADODB.RECORDSET")
 RS.Open "Select * From KS_EnterPrise Where id in(" & id & ")",conn,1,1
 do while not rs.eof
  username=rs("username")
  Conn.execute("Delete From KS_EnterPriseNews Where username='" & username & "'")
  Conn.Execute("UpDate KS_User Set UserType=0 Where UserName='" & username & "'")
  Conn.Execute("update ks_blog set templateid=" & DefaultTemplateID &",BlogName='" & rs("username") & "个人空间' where username='" & rs("username") & "'")
  rs.movenext
 loop
 rs.close:set rs=nothing
 Conn.execute("Delete From KS_UploadFiles Where channelid=1033 and infoid In("& id & ")")
 Conn.execute("Delete From KS_EnterPrise Where id In("& id & ")")
 Response.Write "<script>alert('删除成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'设为精华
Sub Blogrecommend()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_Enterprise Set recommend=1 Where id In("& id & ")")
 Conn.execute("Update KS_Blog Set recommend=1 Where username In(select username from ks_enterprise where id in("& id & "))")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'取消精华
Sub BlogCancelrecommend()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_Enterprise Set recommend=0 Where id In("& id & ")")
 Conn.execute("Update KS_Blog Set recommend=0 Where username In(select username from ks_enterprise where id in("& id & "))")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'锁定
Sub BlogLock()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_Enterprise Set status=2 Where id In("& id & ")")
 conn.execute("update ks_blog set status=2 where username in(select username from ks_enterprise where id in("&id&"))")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'解锁
Sub BlogUnLock()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_Enterprise Set status=1 Where id In("& id & ")")
 conn.execute("update ks_blog set status=1 where username in(select username from ks_enterprise where id in("&id&"))")
 
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'审核
Sub Blogverific
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_Enterprise Set status=1 Where id In("& id & ")")
 conn.execute("update ks_blog set status=1 where username in(select username from ks_enterprise where id in("&id&"))")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
End Class
%> 
