<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_UserGroup
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_UserGroup
        Private KS
		Private MaxPerPage
		Private RS,Sql
		Private ComeUrl
		Private ValidDays,tmpDays,BeginID,EndID,FoundErr,ErrMsg,PowerList
		Private iCount,Action,sPowerType,sDescript,sUserType,ValidType,ValidEmail

		Private Sub Class_Initialize()
		  MaxPerPage=20
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
       Sub Kesion()
            Response.Write"<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">"
			Response.Write "<html xmlns=""http://www.w3.org/1999/xhtml"">"
			Response.Write"<head>"
			Response.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			Response.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			Response.Write"<script src=""../KS_Inc/common.js""></script>"
			Response.Write"<script src=""../KS_Inc/jquery.js""></script>"
			Response.Write "<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>" & vbCrlf
			Response.Write "<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & vbCrlf
			%>
			<script>
			function AddGroup()
		 { 
		 location.href='KS.UserGroup.asp?Action=Add';
		$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape("用户组管理 >> <font color=red>添加用户组</font>")+'&ButtonSymbol=Go';
		}
		function EditGroup(ID)
		{
		 location.href='?Action=Modify&ID='+ID;
		 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape("用户组管理 >> <font color=red>修改用户组</font>")+'&ButtonSymbol=GoSave';
		}</script>
			<%
			Response.Write"</head>"
			Response.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Response.Write"	<ul id='mt'> "
			Response.Write "<div id='mtl'>用户组管理导航：</div><li><a href=""KS.UserGroup.asp"">管理首页</a>&nbsp;|&nbsp;<a href=""#"" onclick=""AddGroup()"">新增用户组</a>"
			Response.Write	" </ul>"
            If Not KS.ReturnPowerResult(0, "KMUA10004") Then
			  response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
			  Call KS.ReturnErr(1, "")
			End If

		Action=Trim(request("Action"))
			Select Case Action
			Case "Add", "Modify"
				call InfoPurview()
			Case "SaveAdd"
				call SaveAdd()
			Case "SaveModify"
				call SaveModify()
			Case "Del"
				call Del()
			Case else
				call main()
			End Select
			
			if FoundErr=True then
				KS.ShowError(ErrMsg)
			end if
			response.Write ""
		End Sub
		
		sub main()
			Set rs=Server.CreateObject("Adodb.RecordSet")
			sql="select * from KS_UserGroup WHERE [TYPE]<2 order by ID"
			OpenConn : rs.Open sql,Conn,1,1
		%>
        <script>
		 
		try{$(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
				$(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
		}catch(e){
		}	
		</script>
		<table border="0" align="center" width="100%" cellpadding="0" cellspacing="0">
		  <tr align="center" class="sort">
			<td  width="45">ID号</td>
			<td width="168">用户组名称</td>
			<td width="390">用户组简介</td>
			<td width="80">类 型</td>
			<td width="80">允许注册</td>
			<td width="120">会员数量</td>
			<td  width="150"> 操 作</td>
		  </tr>
		  <%do while not rs.EOF
			%>
		  <tr height="40" align="center" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
			<td class="splittd" width="45"><%=rs("ID")%></td>
			<td class="splittd"><%=rs("GroupName")%></td>
			<td class="splittd" width="390"><%=rs("Descript")%> </td>
			<td class="splittd" width="91"><%
			if rs("Type")<>0 then
				Response.Write "<font color=blue>自定义</font>"
			else
				Response.Write "<font color=#ff0033>系统</font>"
			end if
			%> </td>
			<td class="splittd" width="91"><%
			if rs("ShowOnReg")=1 then
				Response.Write "<font color=#ff0033>允许注册</font>"
			else
				Response.Write "<font color=green>不允许</font>"
			end if
			%> </td>
			<td class="splittd" width="120"><%=Conn.Execute("Select Count(UserID) From KS_User Where GroupID=" & RS("ID"))(0)%> 位</td>
			<td class="splittd" width="150"><%
			Response.Write "<a href='#' onclick=""EditGroup(" & RS("ID") & ")"">修改</a>&nbsp;&nbsp;"
			if rs("Type")<>0 then Response.Write "<a href='KS.UserGroup.asp?Action=Del&ID=" & rs("ID") & "' onClick=""return confirm('确定要删除此用户组吗？');"">删除</a>"
			%>
			<a href="KS.User.asp?UserSearch=10&GroupID=<%=RS("ID")%>">列出会员</a></td>
		  </tr>
		  <%
			rs.MoveNext
		loop
		  %>
		</table>  
		<%
			rs.Close:set rs=Nothing
		end sub
		
		sub InfoPurview()

		Dim frmAction,sSubmit,GroupSetting,GroupSetArr
		Dim sGroupName,sGroupImg,sFormID,sShowOnReg
		Dim sChargeType,sValidDays,sGroupPoint,sTemplateFile,SpaceSize
		%>
		<SCRIPT language=javascript>
		$(document).ready(function(){
		 setmail($("input[name=ValidType]:checked").val());
		});
		function setmail(n)
		 { 
		   if (n==1){
			  document.getElementById('mailarea').style.display='';
		   }else
			  document.getElementById('mailarea').style.display='none';
		}
		function CheckForm()
		{
		  if(document.myform.GroupName.value=="")
			{
			  alert("用户组名不能为空！");
			  document.myform.GroupName.focus();
			  return false;
			}
		 $("#myform").submit();
		}
		</script>
		  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="ctable" >
				<form method="post" id="myform" action="KS.UserGroup.asp" name="myform" onSubmit="return CheckForm();">
<%
		if Action="Modify" then
			dim GroupID
			GroupID=KS.ChkClng(Trim(Request("ID")))
			if GroupID=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>请指定要修改的用户组ID</li>"
				Exit Sub
			end if
			Set rs=Conn.Execute("Select * from KS_UserGroup where ID=" & GroupID)
			if rs.Bof and rs.EOF then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>不存在此用户组！</li>"
				Exit Sub
			end if
			sGroupName		= rs("GroupName")
			sDescript       = rs("Descript")
			sChargeType		= rs("ChargeType")
			sUserType       = rs("UserType")
			sValidDays		= rs("ValidDays")
			sGroupPoint		= rs("GroupPoint")
			sPowerType      = rs("PowerType")
			PowerList		= rs("PowerList")
			sShowOnReg      = rs("ShowOnReg")
			sTemplateFile   = rs("TemplateFile")
			sFormID         = rs("FormID")
			SpaceSize       = rs("SpaceSize")
			ValidType       = trim(rs("ValidType"))
			ValidEmail      = rs("ValidEmail")
			GroupSetting    = rs("GroupSetting")
			frmAction		= "Modify"
			sSubmit			= "修改"
			rs.close
		else
			sGroupName		= ""
			sChargeType		= 1
			sValidDays		= 0
			sGroupPoint		= 0
			sShowOnReg      = 0
			sDescript       = ""
			frmAction		= "Add"
			sSubmit			= "新增"
			sUserType       = 0
			sTemplateFile   = KS.Setting(116)
			SpaceSize       =1024
			ValidType       =0
			ValidEmail      ="欢迎您注册成为[" & KS.Setting(1) & "]网站会员！" & chr(13) & " 您的验证码：{$CheckNum}" & chr(13) & "请点击下面的地址，输入上面的验证码进行邮件验证。验证通过后，您就可以正式成为我们的会员，享受有关服务了！" & chr(13) & "<a href=""{$CheckUrl}"" target=""_blank"">{$CheckUrl}</a>"
		end if
		GroupSetting=GroupSetting & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
		GroupSetArr =Split(GroupSetting,",")
		Dim CurrPath:CurrPath=KS.Setting(3)&KS.Setting(90)
		If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath,Len(CurrPath)-1)
		%>
			<tr class="sort"> 
			  <td height="25" colspan="2" align="center"><font size="2"><strong><%=sSubmit%>用户组</strong></font></td>
			</tr>
	 </table>
		<div class="tab-page" id="usergroupPanel">
		 <SCRIPT type=text/javascript>
		   var tabPane1 = new WebFXTabPane( document.getElementById( "usergroupPanel" ), 1 )
		 </SCRIPT>
				 
		<div class=tab-page id=basic-page>
		 <H2 class=tab>基本信息</H2>
			<SCRIPT type=text/javascript>
				tabPane1.addTabPage( document.getElementById("basic-page") );
		    </SCRIPT>
		<table class='ctable' width="100%" height=273 border=0 align="center" cellpadding=1 cellspacing=1  style='margin:1px'>
			<tr class="tdbg"> 
			  <td style="width:150px" height="30" align="right" class="clefttitle"><div align="right"><strong>用户组名称：</strong></div></td>
			  <td><input name="GroupName" class="textbox" type="text" size=30 value="<%=sGroupName%>">		      </td>
			</tr>
			<tr class="tdbg">
			  <td  height="30" align="right" class="clefttitle"><div align="right"><strong>用户组简介：</strong></div></td>
			  <td><textarea name="Descript" cols="50" rows="5" id="Descript"><%=sDescript%></textarea></TD>
		    </tr>
			<tr class="tdbg">
			  <td  height="30" align="right" class="clefttitle"><div align="right"><strong>用户组类别：</strong></div></td>
			  <td><input name="UserType" type="radio" value="0" <%if sUserType=0 then Response.Write " checked"%>>
			    个人会员 
		        <input name="UserType" type="radio" value="1" <%if sUserType=1 then Response.Write " checked"%>>		        企业会员</TD>
		    </tr>
			<tr class="tdbg"> 
			  <td  height="30" align="right" class="clefttitle"><div align="right"><strong>用户组计费方式：</strong></div></td>
			  <td>
			    <label><input name="ChargeType" onclick="$('#ds').show();$('#yxq').hide();" type="radio" value="1" <%if sChargeType=1 then Response.Write " checked"%> >
				扣点数</label>
				<span id='ds'<%if sChargeType<>1 then Response.Write " style='display:none'"%>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;默认点数：
<input name="GroupPoint" type="text" id="GroupPoint" style="text-align:center" class="textbox"  value="<%=sGroupPoint%>" size="5" maxlength="5"> 
点</span><br>
				<label><input type="radio" onclick="$('#ds').hide();$('#yxq').show();" name="ChargeType" value="2" <%if sChargeType=2 then Response.Write " checked"%> >
				有效期</label><span id='yxq'<%if sChargeType<>2 then Response.Write " style='display:none'"%>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;默认有效期：
<input name="ValidDays" type="text" id="ValidDays" style="text-align:center" class="textbox" l value="<%=sValidDays%>" size="5" maxlength="5"> 
天</span><br />
 <label>
<input type="radio" name="ChargeType" value="3" <%if sChargeType=3 then Response.Write " checked"%>> 
无限期(永不过期)</label></TD>
			</tr>
			<tr class="tdbg"> 
			  <td height="30" align="right" class="clefttitle"><div align="right"><strong>会员中心模板：<br>
		      </strong></div></td>
			  <td>&nbsp;
			  <input type="text" name="TemplateFile" id="TemplateFile" size="30" value="<%=sTemplateFile%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle='+escape('选择模板')+'&CurrPath=<%=CurrPath%>',450,350,window,$('#TemplateFile')[0]);">		  </td>
			</tr>
			
			<tr class="tdbg"> 
			  <td height="30" align="right" class="clefttitle"><div align="right"><strong>选择录入表单：<br>
		      </strong></div></td>
			  <td height="30">&nbsp;
			  <select name="formid">
			   <%
			    Set RS=Conn.Execute("select id,formname from ks_userform")
				do while not rs.eof
				 If sFormID=rs(0) Then
				 response.write "<option value='" & rs(0) & "' selected>" & rs(1) & "</option>"
				 Else
				 response.write "<option value='" & rs(0) & "'>" & rs(1) & "</option>"
				 End If
				 rs.movenext
				loop
				rs.close
			   %>
			  </select>			  </td>
			</tr>
			<tr class="tdbg">
			  <td height="30" align="right" class="clefttitle"><div align="right"><strong>是否允许前台注册选择：</strong></div></td>
			  <td>
			  <input type='radio' name='ShowOnReg' value='1'<%if sShowOnReg="1" Then Response.Write " Checked"%>> 允许 <input type='radio' name='ShowOnReg' value='0'<%if sShowOnReg="0" Then Response.Write " Checked"%>>不允许			  </td>
		    </tr>
			<tr class="tdbg">
			      <td height="21" class='clefttitle' align="right"><div><strong>新会员注册验证方式：</strong></div></td>
			      <td>
				  <input id='a1' onClick="setmail(this.value)" name="ValidType" type="radio"  value="0"<%If ValidType="0" Then Response.Write " checked"%>><label for='a1'>无需验证</label><br>
			     <input id='a2' onClick="setmail(this.value)" name="ValidType" type="radio" value="1"<%If ValidType="1" Then Response.Write" Checked"%>><label for='a2'>邮件验证</label>(<font class='tips'>会员注册后系统会发一封带有验证码的邮件给此会员，会员必须在通过邮件验证后才能真正成为正式注册会员</font>)<br>
			     <input id='a3' onClick="setmail(this.value)" name="ValidType" type="radio" value="2"<%If ValidType="2" Then   Response.Write " Checked"%> /><label for='a3'>后台人工验证</label>
			 </td>
			</tr>
			<tr valign="middle" id="mailarea"  class="tdbg">
			      <td align="right" class='clefttitle'><strong>会员注册发送邮件内容：</strong></td>
			      <td ><textarea name="ValidEmail" id="ValidEmail" cols="70" rows="5"><%=ValidEmail%></textarea>
			<div style="margin:3px"><b>标签说明：</b><br/>{$CheckNum}：验证码<br/>{$CheckUrl}：会员注册验证地址<br/>{$UserName}：用户名<br/>{$PassWord}：密码</div></td>
			</tr>
		</table>
	   </div>
	   
		<div class=tab-page id=basic-page>
		 <H2 class=tab>权限分配</H2>
			<SCRIPT type=text/javascript>
				tabPane1.addTabPage( document.getElementById("basic-page") );
		    </SCRIPT>
		<table class='ctable' width="100%" height=273 border=0 align="center" cellpadding=1 cellspacing=1  style='margin:1px'>
			<tr class="tdbg">
			  <td height="30" style="width:150px" align="right" class="clefttitle"><div align="right"><strong>分配空间大小：</strong></div></td>
			  <td>&nbsp;
<input type="text" name="SpaceSize" size="10" class="textbox" style="text-align:center" value="<%=SpaceSize%>" />KB <font color="#FF0000">提示：1 KB = 1024 Byte，1 MB = 1024 KB</font> </td>
		    </tr>
			
			
			<tr class="tdbg">
			  <td height="30" align="right" class="clefttitle"><div align="right"><strong>会员功能分配：</strong></div></td>
			  <td height="30">
			    
				  <table border="0" width="100%">
				   <tr>
				    <td><input name="PowerList" type="checkbox" value="s01"<%if InStr(1, PowerList,"s01" ,1)<>0 then Response.Write( "checked") %>>个人/企业空间
					</td>
				    <td><input name="PowerList" type="checkbox" value="s02"<%if InStr(1, PowerList,"s02" ,1)<>0 then Response.Write( "checked") %>>日志功能
					</td>
				    <td><input name="PowerList" type="checkbox" value="s03"<%if InStr(1, PowerList,"s03" ,1)<>0 then Response.Write( "checked") %>>好友功能
					</td>
				    <td><input name="PowerList" type="checkbox" value="s04"<%if InStr(1, PowerList,"s04" ,1)<>0 then Response.Write( "checked") %>>音乐功能
					</td>
				    <td><input name="PowerList" type="checkbox" value="s05"<%if InStr(1, PowerList,"s05" ,1)<>0 then Response.Write( "checked") %>>相册功能
					</td>
				   </tr>
				   <tr>
				    <td><input name="PowerList" type="checkbox" value="s06"<%if InStr(1, PowerList,"s06" ,1)<>0 then Response.Write( "checked") %>>圈子功能
					</td>
				    <td><input name="PowerList" type="checkbox" value="s07"<%if InStr(1, PowerList,"s07" ,1)<>0 then Response.Write( "checked") %>>自定义分类
					</td>
				    <td><input name="PowerList" type="checkbox" value="s08"<%if InStr(1, PowerList,"s08" ,1)<>0 then Response.Write( "checked") %>>实名认证
					</td>
				    <td><input name="PowerList" type="checkbox" value="s09"<%if InStr(1, PowerList,"s09" ,1)<>0 then Response.Write( "checked") %>>显示问答
					</td>
				    <td>
					<input name="PowerList" type="checkbox" value="s19"<%if InStr(1, PowerList,"s19" ,1)<>0 then Response.Write( "checked") %>>显示论坛
					</td>
					</tr>
				   <tr>
				    <td><input name="PowerList" type="checkbox" value="s10"<%if InStr(1, PowerList,"s10" ,1)<>0 then Response.Write( "checked") %>>企业产品(宝贝)
					</td>
				    <td><input name="PowerList" type="checkbox" value="s11"<%if InStr(1, PowerList,"s11" ,1)<>0 then Response.Write( "checked") %>>企业新闻
					</td>
				    <td><input name="PowerList" type="checkbox" value="s12"<%if InStr(1, PowerList,"s12" ,1)<>0 then Response.Write( "checked") %>>关键词广告
					</td>
				    <td><input name="PowerList" type="checkbox" value="s13"<%if InStr(1, PowerList,"s13" ,1)<>0 then Response.Write( "checked") %>>荣誉证书
					</td>
				    <td><input name="PowerList" type="checkbox" value="s14"<%if InStr(1, PowerList,"s14" ,1)<>0 then Response.Write( "checked") %>>求职招聘
					</td>
				   </tr>
				  
				  <tr>
				    <td><input name="PowerList" type="checkbox" value="s15"<%if InStr(1, PowerList,"s15" ,1)<>0 then Response.Write( "checked") %>>积分兑换
					</td>
				    <td><input name="PowerList" type="checkbox" value="s16"<%if InStr(1, PowerList,"s16" ,1)<>0 then Response.Write( "checked") %>>收藏夹
					</td>
				    <td><input name="PowerList" type="checkbox" value="s17"<%if InStr(1, PowerList,"s17" ,1)<>0 then Response.Write( "checked") %>>投诉建议
					</td>
				    <td><input name="PowerList" type="checkbox" value="s18"<%if InStr(1, PowerList,"s18" ,1)<>0 then Response.Write( "checked") %>>内容发布(投稿)
					</td>
				    <td>
					<input name="PowerList" type="checkbox" value="s20"<%if InStr(1, PowerList,"s20" ,1)<>0 then Response.Write( "checked") %>>显示签收

					</td>
				   </tr>

				   
				   </table>
				   
			   </td>
		    </tr>
			<tr><td colspan=2><hr color="green" size="1"></td></tr>
			<tr class="tdbg">
			  <td height="30" align="right" class="clefttitle"><div align="right"><strong>特殊功能选项：</strong></div></td>
			  <td height="30">
			    <input type='checkbox' name='groupsetting0'<%if GroupSetArr(0)="1" then response.write " checked"%> value='1'>在发布信息需要审核的模型，此会员组发布信息不需要审核<br/>
			    <input type='checkbox' name='groupsetting1'<%if GroupSetArr(1)="1" then response.write " checked"%> value='1'>可以修改和删除已审核的（自己的）文档<br/>
				一天内最多只能在同一个模型发布<input type='text' class='textbox' name='GroupSetting2' value="<%=GroupSetArr(2)%>" style='text-align:center;width:30px' />篇文档  <font color='blue'>不限制请输入"-1"</font><br/>
			    <font color=#ff6600>一天内最多只能在同一类文档使用<input class='textbox' type='text' name='GroupSetting11' value="<%=GroupSetArr(11)%>" style='text-align:center;width:30px' />次  <font color='blue'>不限制请输入"0"</font>,此功能一般用于控制按有效期计费的会员权限,防止一次性恶意下载,查看全部收费信息，不是按有效期计费的用户组，建议设置成“0”。</font><br/>
				发布信息时获取积分为模型设置的<input type='text' class='textbox' name='GroupSetting3' value="<%=GroupSetArr(3)%>" style='text-align:center;width:30px' />倍 <font color='blue'>输入"0" 表示此会员组不得分</font><br/>
				发布信息时获取点券为模型设置的<input type='text' class='textbox' name='GroupSetting4' value="<%=GroupSetArr(4)%>" style='text-align:center;width:30px' />倍 <font color='blue'>输入"0" 表示此会员组不得点券</font><br/>
				发布信息时获取资金为模型设置的<input type='text' class='textbox' name='GroupSetting5' value="<%=GroupSetArr(5)%>" style='text-align:center;width:30px' />倍 <font color='blue'>输入"0" 表示此会员组不得资金</font><br/>
				此会员组发表评论可得：<input type="text" class='textbox' name="GroupSetting6" size="5" value="<%=GroupSetArr(6)%>" style="text-align:center"/>分积分作为奖励
               1个月内评论被删除将扣除<input type="text" class='textbox' name="GroupSetting7" size="5" value="<%=GroupSetArr(7)%>" style="text-align:center"/>分积分
			   <font color=blue>可输入"0"表示不增加也不减少积分</font><br/>
			   
			   此会员组发表信息被审核后是否发站内短消息通知<input type="radio" name="GroupSetting10" value="1" <%if GroupSetArr(10)="1" then response.write " checked"%>>是 <input type="radio" name="GroupSetting10" value="0" <%if GroupSetArr(10)="0" then response.write " checked"%>>否 <br/>
			   
			   此会员组会员每隔<input type="text" class='textbox' name="GroupSetting8" size="5" value="<%=GroupSetArr(8)%>" style="text-align:center"/>分钟后,重新登录奖励<input class='textbox' type="text" name="GroupSetting9" size="5" value="<%=GroupSetArr(9)%>" style="text-align:center"/>分积分 <font color=blue>不想奖励请输入"0"</font>
			   <br/>
			   
			   <div style="color:blue">
			   此会员组在允许刷新添加时间的模型里允许在<input class='textbox' type="text" name="GroupSetting12" size="5" value="<%=GroupSetArr(12)%>" style="text-align:center"/>分钟后重新刷新发布。不允许请输入"0"
			   </div>
			   
			   	短消息设置：最大容量为<input type="text" name="GroupSetting13" size="5" value="<%=GroupSetArr(13)%>" style="text-align:center"/>条,短信内容最多字符数<input type="text" name="GroupSetting14" size="5" value="<%=GroupSetArr(14)%>" style="text-align:center"/>个字符 群发限制人数<input type="text" name="GroupSetting15" size="5" value="<%=GroupSetArr(15)%>" style="text-align:center"/>人 <span style="color:#999">不限制，请输入"0"</span>
				<br/>
				允许上传附件:
				 
				 <%
				 Response.Write "<input type='radio' onclick=""$('#fj').show();"" name=""GroupSetting22"" value=""1"" " 
				If GroupSetArr(22) = "1" Then Response.Write (" checked")
				 response.write "> 允许"
				 Response.Write "    <input type=""radio"" onclick=""$('#fj').hide();"" name=""GroupSetting22"" value=""0"" "
				If GroupSetArr(22) = "0" Then Response.Write (" checked")
				 Response.Write "> 不允许"
				If GroupSetArr(22) = "1" Then
                 Response.Write "<div id='fj' style='color:blue'>"
				Else
                 Response.Write "<div id='fj' style='display:none;color:blue'>"
				End If
				 Response.Write "允许上传的附件扩展名:<input class='textbox' type='text' value='" & GroupSetArr(23) & "' name='GroupSetting23' /> 多个扩展名用 |隔开,如gif|jpg|rar等<Br/>允许上传的文件大小：<input class='textbox' name=""GroupSetting24"" type=""text"" value=""" & GroupSetArr(24) &""" style=""text-align:center"" size='8'>KB<br/>每天上传文件个数：<input class='textbox' name=""GroupSetting25"" type=""text"" value=""" & GroupSetArr(25) &""" style=""text-align:center"" size='8'>个,不限制请填0</font>"
				 %>
				 </div>
				 签名字数限制：最大<input type="text" name="GroupSetting26" size="5" value="<%=GroupSetArr(26)%>" style="text-align:center"/>个字符 <span style="color:#999">不限制，请输入"0"</span>
			  </td>
			 </tr>
 			<tr><td colspan=2><hr color="green" size="1"></td></tr>
			<tr class="tdbg">
			 <td height="30" align="right" class="clefttitle"><div align="right"><strong>商城权限：</strong></div></td>
<td>

				 <table border="0" width="100%" cellpadding="0" cellspacing="0">
				   <tr>
				    <td style="text-align:right;width:100px"><strong>自动升级权限：</strong></td>
					<td>累计在商城消费<input type="text" class='textbox' name="GroupSetting16" value="<%=GroupSetArr(16)%>" style="text-align:center;width:40px" />元，可以自动升级到此会员组。<span class='tips'>如果不想自动升级请输入“0”</span></td>
				   </tr>
				   <tr>
				    <td style="text-align:right;width:100px"><strong>商城优惠措施：</strong></td>
					<td>享受正价产品<input type="text"  class='textbox' name="GroupSetting17" value="<%=GroupSetArr(17)%>" style="text-align:center;width:40px" />折优惠 <span class='tips'>该用户组没有任何优惠请输入“0”</span><br/>
					享受正价产品<input type="text" class='textbox' name="GroupSetting18" value="<%=GroupSetArr(18)%>" style="text-align:center;width:40px" />倍积分  <span class='tips'>该用户组购物不奖积分请输入“0”</span></td>
				   </tr>
				   <tr>
				    <td style="text-align:right;width:100px"><strong>永久享受推荐奖励积分：</strong></td>
					<td><input type="radio" onclick="$('#jf').show();" name="GroupSetting19" value="1"<%if ks.chkclng(GroupSetArr(19))=1 then response.write " checked"%> />是 <input onclick="$('#jf').hide();" type="radio" name="GroupSetting19" value="0"<%if ks.chkclng(GroupSetArr(19))=0 then response.write " checked"%> />否  <span class='tips'>(选择“是”将享受推荐奖励积分制度)</span>
					 <div id='jf' <%if ks.chkclng(GroupSetArr(19))=0 then response.write " style='display:none'"%>>
					  享受奖励积分百分比<input type="text" class='textbox' name="GroupSetting20" value="<%=GroupSetArr(20)%>" style="text-align:center;width:40px" />% 
					 </div>
					 
					</td>
				   </tr>
				   <tr>
				    <td style="text-align:right;width:100px"><strong>独享VIP会员专用客服通道：</strong></td>
					<td><input type="radio"  name="GroupSetting21" value="1"<%if ks.chkclng(GroupSetArr(21))=1 then response.write " checked"%> />是 <input type="radio" name="GroupSetting21" value="0"<%if ks.chkclng(GroupSetArr(21))=0 then response.write " checked"%> />否
					
					</td>
				   </tr>
				 </table>

              </td>
			</tr>			
			
				<input name="ID" type="hidden" value="<%=GroupID%>">
				<input name="Action" type="hidden" id="Action" value="Save<%=frmAction%>">
		  </table>
		</form>
	  </div>
		<% 
			Set rs=Nothing
		end sub
		
		sub SaveAdd()
			Dim GroupName,ChargeType,ValidDays,GroupPoint,PowerType,PowerList,Descript,FormID,ShowOnReg,UserType,TemplateFile,SpaceSize,GroupSetting,i
			
			GroupName		= Trim(Request("GroupName"))
			ChargeType		= KS.ChkClng(Request("ChargeType"))
			PowerType       = KS.ChkClng(Request("PowerType"))
			PowerList       = Request("PowerList")
			ValidDays		= KS.ChkClng(Request("ValidDays"))
			GroupPoint		= KS.ChkClng(Request("GroupPoint"))
			FormID          = KS.ChkClng(Request("FormID"))
			ShowOnReg       = KS.ChkClng(Request("ShowOnReg"))
			Descript        = KS.G("Descript")
			UserType        = KS.ChkClng(Request("UserType"))
			TemplateFile    = Request("TemplateFile")
			SpaceSize       = KS.ChkClng(Request("SpaceSize"))
			ValidType       = KS.ChkClng(Request("ValidType"))
			ValidEmail      = Request.Form("ValidEmail")
			GroupSetting=""
			For i=0 to 30
			   If GroupSetting="" Then
			     GroupSetting=KS.ChkClng(Request("GroupSetting"&i))
			   Else
			     if i=16 or i=17 or i=18 then
			     GroupSetting=GroupSetting &"," & Request("GroupSetting"&i)
				 else
			     GroupSetting=GroupSetting &"," & KS.ChkClng(Request("GroupSetting"&i))
				 end if
			   End If
			Next


			if GroupName="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>用户组名称不能为空！</li>"
			end if
			if Not IsNumeric(Replace(Replace(PowerType,",","")," ","")) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>用户权限错误！</li>"
			end if
			if FoundErr=True then Exit Sub
			if ChargeType<>1 and ChargeType<>2 and ChargeType<>3 then ChargeType=1
			
			
			sql="Select * from KS_UserGroup where GroupName='"&GroupName&"'"
			Set rs=Server.CreateObject("Adodb.RecordSet")
			rs.Open sql,Conn,1,3
			if not (rs.bof and rs.EOF) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>数据库中已经存在此用户组名称！</li>"
				rs.close:set rs=Nothing
				exit sub
			end if
			rs.addnew
			rs("GroupName")		= GroupName
			rs("ChargeType")	= ChargeType
			rs("ValidDays")		= ValidDays
			rs("GroupPoint")	= GroupPoint
			rs("PowerList")		= PowerList
			rs("PowerType")     = PowerType
			rs("FormID")        = FormID
			rs("ShowOnReg")     = ShowOnReg
			rs("Descript")	    = Descript
			rs("UserType")      = UserType
			rs("TemplateFile")  = TemplateFile
			rs("SpaceSize")     = SpaceSize
			rs("ValidType")     = ValidType
			rs("ValidEmail")    = ValidEmail
			rs("GroupSetting")  = GroupSetting
			Rs("Type")          = 1
			rs.update
			rs.Close:set rs=Nothing
			Application(KS.SiteSN&"_UserGroup")=empty
			Call KS.Confirm("添加新用户组成功,继续添加吗?","KS.UserGroup.asp?Action=Add","KS.UserGroup.asp")
		end sub
		
		sub SaveModify()
			Dim GroupName,GroupID,GroupSetting,I
			Dim ChargeType,ValidDays,GroupPoint,PowerType,PowerList,Descript,FormID,ShowOnReg,UserType,TemplateFile,SpaceSize
			GroupID		= Trim(Request("ID"))
			GroupName		= Trim(Request("GroupName"))
			ChargeType		= KS.ChkClng(Request("ChargeType"))
			UserType        = KS.ChkClng(Request("UserType"))
			PowerType       = KS.ChkClng(Request("PowerType"))
			PowerList       = Request("PowerList")
			ValidDays		= KS.ChkClng(Request("ValidDays"))
			GroupPoint		= KS.ChkClng(Request("GroupPoint"))
			FormID          = KS.ChkClng(Request("FormID"))
			ShowOnReg       = KS.ChkClng(Request("ShowOnReg"))
			SpaceSize       = KS.ChkClng(Request("SpaceSize"))
			TemplateFile    = Request("TemplateFile")
			Descript        =KS.G("Descript")
			ValidType       = KS.ChkClng(Request("ValidType"))
			ValidEmail      = Request.Form("ValidEmail")
			
			GroupSetting=""
			For i=0 to 30
			   If GroupSetting="" Then
			     GroupSetting=KS.ChkClng(Request("GroupSetting"&i))
			   Else
			     if i=16 or i=17 or i=18 or i=23 then
			     GroupSetting=GroupSetting &"," & Request("GroupSetting"&i)
				 else
			     GroupSetting=GroupSetting &"," & KS.ChkClng(Request("GroupSetting"&i))
				 end if
			   End If
			Next
			
			if GroupName="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>用户组名称不能为空！</li>"
			end if
			if Not IsNumeric(Replace(Replace(PowerType,",","")," ","")) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>用户权限错误！</li>"
			end if
			if FoundErr=True then Exit Sub
			if ChargeType<>1 and ChargeType<>2 and ChargeType<>3 then ChargeType=1
			
			
			sql="Select * from KS_UserGroup where ID="&GroupID
			Set rs=Server.CreateObject("Adodb.RecordSet")
			rs.Open sql,Conn,1,3
			if not (rs.bof and rs.EOF) then
			rs("GroupName")		= GroupName
			rs("ChargeType")	= ChargeType
			rs("ValidDays")		= ValidDays
			rs("GroupPoint")	= GroupPoint
			rs("PowerList")		= PowerList
			rs("PowerType")     = PowerType
			rs("FormID")        = FormID
			rs("ShowOnReg")     = ShowOnReg
			rs("Descript")	    = Descript
			rs("UserType")      = UserType
			rs("TemplateFile")  = TemplateFile
			rs("SpaceSize")     = SpaceSize
			rs("ValidType")     = ValidType
			rs("ValidEmail")    = ValidEmail
			rs("GroupSetting")  = GroupSetting
			rs.update
		   else
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>无此用户组，参数传递出错！</li>"
		   end if
			rs.Close:set rs=Nothing
			conn.execute("update ks_user set usertype=" & UserType &" where groupid=" & groupid)
			IF FoundErr=true Then 
			 Exit Sub
			else
			Application(KS.SiteSN&"_UserGroup")=empty
			Response.Write ("<script>alert('用户组权限修改成功！');location.href='KS.UserGroup.asp';</script>")
			end if
		end sub
		sub Del()
		dim GroupID
		GroupID=Trim(Request("ID"))
		if GroupID="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请指定要删除的管理员ID</li>"
			exit sub
		end if
		GroupID=Clng(GroupID)
		'更改前台用户组所属组权限
		if GroupID=0 then KS.ShowError("<br><li>您不能删除系统用户组！</li>")
		Conn.Execute("Update KS_User Set GroupID=2 where GroupID=" & GroupID)
		Conn.Execute("delete from KS_UserGroup where ID=" & GroupID)
		Application(KS.SiteSN&"_UserGroup")=empty
		Call KS.Alert("删除用户组成功！","KS.UserGroup.asp")
end sub
End Class
		%>
 
