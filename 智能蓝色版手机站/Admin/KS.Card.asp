<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Card
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Card
        Private KS,CardType
		Private MaxPerPage,RS,TotalPut,TotalPages,I,CurrentPage,SQL,ComeUrl
		Private Sub Class_Initialize()
		  MaxPerPage=20
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
       Sub Kesion()
	        If Not KS.ReturnPowerResult(0, "KMUA10008") Then
			  Call KS.ReturnErr(1, "")
			End If
			CardType=KS.ChkClng(KS.G("CardType"))
          Response.Write "<html>"
			Response.Write"<head>"
			Response.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			Response.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css""><script src='../KS_Inc/common.js'></script><script src='../KS_Inc/jQuery.js'></script>"
			Response.Write"</head>"
			Response.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Response.Write"	<ul id='mt'> "
			Response.Write " <div id='mtl'>操作导航:</div><li>&nbsp;<a href=""?cardtype=" & cardtype & """>所有充值卡</a> | "
			If CardType=1 Then
			Response.Write "<a href=""?action=AddMore&cardtype=1"">生成在线充值卡</a></li>"
			Else
			Response.Write "<a href=""?status=1&cardtype=0"">未使用充值卡</a> | <a href=""?status=2&cardtype=0"">已使用充值卡</a> | <a href=""?status=3&cardtype=0"">已失效充值卡</a> | <a href=""?status=4&cardtype=0"">未失效充值卡</a> | <a href=""?action=Add&cardtype=0"">添加充值卡</a> | <a href=""?action=AddMore&cardtype=0"">批量生成充值卡</a></li>"
			End If
			Response.Write	" </ul>"

		ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))
		
		Select Case KS.G("Action")
		 Case "Add","Edit"
		  Call Add()
		 Case  "DoAdd"
		  Call DoAdd()
		 Case "AddMore"
		  Call AddMore()
		 Case "DoAddMore"
		  Call DoAddMore()
		 Case "Del"
		  Call Del()
		 Case Else
		  Call CardList()
		End Select
	End Sub
	
	'点卡列表
	Sub CardList()
		%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
  <tr class="sort">
   <%if cardtype="1" then%>
    <td width="38" align="center"><strong>选中</strong></td>
    <td align="center"><strong>充值卡名称</strong></td>
    <td width="75" align="center"><strong>面值</strong></td>
    <td width="79" align="center" nowrap="nowrap"><strong>点数/天数</strong></td>
    <td align="center"><strong>过期时间</strong></td>
    <td align="center"><strong>操作</strong></td>
   <%else%>
    <td width="38" align="center"><strong>选中</strong></td>
    <td width="116" align="center"><strong>名称</strong></td>
    <td width="116" align="center"><strong>充值卡号</strong></td>
    <td width="88" align="center"><strong>密码</strong></td>
    <td width="75" align="center"><strong>面值</strong></td>
    <td width="79" align="center" nowrap="nowrap"><strong>点数/天数</strong></td>
    <td width="100" align="center"><strong>过期时间</strong></td>
    <td width="60" align="center"><strong>出售</strong></td>
    <td width="60" align="center"><strong>使用</strong></td>
    <td width="100" align="center"><strong>使用者</strong></td>
    <td width="100" align="center"><strong>充值时间</strong></td>
    <td width="80" align="center"><strong>操作</strong></td>
  <%end if%>
  </tr>
  <%
  CurrentPage	= KS.ChkClng(request("page"))
  Dim Param:Param=" where cardtype=" &cardtype
  if KS.G("groupname")<>"" Then Param=Param & " and groupname='" & KS.G("groupname") & "'"
  if KS.G("KeyWord")<>"" Then Param=Param & " and cardnum='" & KS.G("KeyWord") & "'"
  Select Case  KS.ChkClng(KS.G("Status"))
   Case 1
     Param=Param & " And IsUsed=0"
   Case 2
     Param=Param & " And IsUsed=1"
   Case 3
     Param=Param & " And datediff(" & DataPart_D & ",EndDate,"&SqlNowString&")>0"
   Case 4
     Param=Param & " And datediff(" & DataPart_D & ",EndDate,"&SqlNowString&")<0"
  End Select
  
  Dim SqlStr:SqlStr="Select ID,CardNum,CardPass,Money,ValidNum,ValidUnit,AddDate,EndDate,UseDate,UserName,IsUsed,IsSale,groupname From KS_UserCard " & Param & " order by ID desc"
  Set RS=Server.CreateObject("ADODB.RecordSet")
    RS.Open SqlStr,conn,1,1
	If RS.Eof And RS.Bof Then
	 Response.Write "<tr><td colspan=11 align=center height=25>没有充值卡！</td></tr><tr><td colspan=13 background='images/line.gif'></td></tr>"
	Else
                    TotalPut=rs.recordcount
					if (TotalPut mod MaxPerPage)=0 then
						TotalPages = TotalPut \ MaxPerPage
					else
						TotalPages = TotalPut \ MaxPerPage + 1
					end if
					if CurrentPage > TotalPages then CurrentPage=TotalPages
					if CurrentPage < 1 then CurrentPage=1
					rs.move (CurrentPage-1)*MaxPerPage
					SQL = rs.GetRows(MaxPerPage)
					rs.Close:set rs=Nothing
					ShowContent
   End If
%>		
</table>
<%if cardtype=0 then%>
<table border="0" style="margin-top:20px" width="100%" align=center>
<form action="KS.Card.asp" name="myform" method="post">
<tr><td>
<div style="border:1px dashed #cccccc;margin:3px;padding:4px"><b>快速搜索=></b>:
<%
	           Response.Write " &nbsp;<select name='groupname'>"
			   Response.Write "<option value=''>====选择分类====</option>"
				 Dim ZRS:Set ZRS=Server.CreateObject("ADODB.RECORDSET")
				 ZRS.Open "select Distinct groupname from ks_usercard where groupname<>'' and groupname<>null",conn,1,1
				 If Not ZRS.Eof Then
				  Do While Not ZRS.Eof 
				   if ks.g("groupname")=zrs(0) then
				   Response.Write "<option value='" & ZRS(0) & "' selected>" & ZRS(0) & "</option>"
				   else
				   Response.Write "<option value='" & ZRS(0) & "'>" & ZRS(0) & "</option>"
				   end if
				   ZRS.MoveNext
				  Loop
				 End If
				 ZRS.Close:Set ZRS=Nothing
			    Response.Write "</select>"
	  %>
卡号
<input type="text" name="keyword" value="" class='textbox' size="20">&nbsp;<input type="submit" value="开始搜索" class="button">
</div>
</td></tr>
</form>
<tr><td><br><Font color=red><strong>提示：</strong>
已售出或已使用的充值卡，不允许删除，修改等操作。</font>
</td></tr>
</table>
<%
 end if
End Sub
Sub ShowContent
 Dim InPoint,OutPoint
 %>
 <form name=selform method=post action=?action=Del&cardtype=<%=cardtype%>>
 <%
For i=0 To Ubound(SQL,2)
	%>
  <tr height="22" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
    <td align="center" class="splittd"><input type="checkbox" name="id" value="<%=SQL(0,i)%>"></td>
    <td align="center" class="splittd"><%=SQL(12,i)%></td>
	<%if cardtype="1" then%>
		<td align="center" class="splittd"><%=formatnumber(SQL(3,i),2,-1)%>元</td>
		<td align="center" class="splittd"><%Response.Write SQL(4,I)
		if SQL(5,I)=1 Then 
		 Response.Write "点" 
		ELSEIf SQL(5,I)=2 Then 
		 Response.Write "天" 
		elseif SQL(5,I)=3 Then
		 response.write "元"
		end if%></td>
    <td align="center" class="splittd"><%Response.Write formatdatetime(SQL(7,I),2)%></td>
	<%Else%>
    <td align="center" class="splittd"><%=SQL(1,i)%></td>
    <td align="center" class="splittd"><%=KS.Decrypt(SQL(2,i))%></td>
    <td align="center" class="splittd"><%=formatnumber(SQL(3,i),2,-1)%>元</td>
    <td align="center" class="splittd"><%Response.Write SQL(4,I)
	if SQL(5,I)=1 Then 
	 Response.Write "点" 
	ELSEIf SQL(5,I)=2 Then 
	 Response.Write "天" 
	elseif SQL(5,I)=3 Then
	 response.write "元"
	end if%></td>
    <td align="center" class="splittd"><%Response.Write formatdatetime(SQL(7,I),2)%></td>
    <td align="center" class="splittd">
	<%
	IF SQL(11,I)=1 Then
	 Response.Write "已售出"
	Else
	 Response.Write "<font color=red>未出售</font>" 
	End If
	%></td>
    <td align="center" class="splittd">
	<%
	IF SQL(10,I)=1 Then
	 Response.Write "<font color='#a7a7a7'>已使用</font>"
	Else
	 Response.Write "<font color=red>未使用</font>" 
	End If
	%></td>
    <td align="center" class="splittd">&nbsp;<%Response.Write SQL(9,I)%>&nbsp;</td>
    <td align="center" class="splittd">
	<%if Isdate(Sql(8,i)) then
	   response.write formatdatetime(SQL(8,i),2)
	  end if%>&nbsp;</td>
	  
	<%end if%>
	<td align="center" class="splittd">
	<%if SQL(11,I)<>1 and SQL(10,I)<>1 then%>
	<a href="?action=Edit&ID=<%=SQL(0,i)%>&cardtype=<%=cardtype%>">修改</a> <a href="?action=Del&cardtype=<%=cardtype%>&ID=<%=SQL(0,i)%>">删除</a>
	<%end if%>&nbsp;
	</td>
  </tr>
  <%Next
  
  Response.Write "<tr onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'""><td height='30' colspan=4>"
  Response.Write "&nbsp;&nbsp;<input id=""chkAll"" onClick=""CheckAll(this.form)"" type=""checkbox"" value=""checkbox""  name=""chkAll"">全选&nbsp;&nbsp;<input class=Button type=""submit"" name=""Submit2"" value="" 删除充值卡 "" onclick=""{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){this.document.selform.submit();return true;}return false;}""> <input type='button' value=' 打 印 ' onclick='window.print()' class='button'></td><td colspan=13 align=right><br>"
  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
  Response.Write "</td></tr>  <tr><td colspan=13 background='images/line.gif'></td></tr></form>"
End Sub

 '删除充值卡
 Sub Del()
  Dim ID:ID=Replace(KS.G("ID")," ","")
  ID=KS.FilterIDs(ID)
  If ID="" Then Response.Write "<script>alert('请选择充值卡!');history.back();</script>"
  Conn.Execute("Delete From KS_UserCard Where ID In(" & ID &") and IsSale=0 and IsUsed=0")
  Response.Write "<script>alert('删除成功！');location.href='" & Request.Servervariables("http_referer") & "';</script>"
 End Sub
		
		'批量添加充值卡
  Sub AddMore()
		%>
  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>
		<form method='post' action='KS.Card.asp' name='myform'>
    <tr class='sort'> 
      <td height='22' colspan='2'> <div align='center'><strong>批 量 生 成 充 值 卡</strong></div></td>
    </tr>
    <tr class='tdbg'> 
      <td width='40%' class="clefttitle"><strong>充值卡名称：</strong></td>
      <td width='60%'><input name='GroupName' type='text' size='20' maxlength='100'> 如:“10元购100点充值卡”等</td>
    </tr>
    <tr class='tdbg'> 
      <td width='40%' class="clefttitle"><strong>充值方式：</strong></td>
      <td width='60%'>
	  <input type="radio" name="cardtype" value="1"  onclick="$('#showother').hide()"<%if cardtype=0 then response.write " disabled" else response.write " checked"%>>在线充值购买(会员中心由会员自己购买充值）<br/>
	  <input type="radio" name="cardtype" value="0" onclick="$('#showother').show()"<%if cardtype=1 then response.write " disabled" else response.write " checked"%>>线下销售
	  </td>
    </tr>
   <tbody id="showother" style="display:<%if cardtype=0 then response.write "" else response.write "none"%>">
    <tr class='tdbg'> 
      <td width='40%' class="clefttitle"><strong>充值卡数量：</strong></td>
      <td width='60%'><input name='Nums' type='text' value='100' size='10' maxlength='10'>
        张</td>
    </tr>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>充值卡号码前缀：</strong><br>
        例如：2006,KS2006等固定不变的字母或数字</td>
      <td width='60%'><input name='CardNumPrefix' type='text' id='CardNumPrefix' value='KS2007' size='10' maxlength='10'></td>
    </tr>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>充值卡号码位数：</strong><br>请输入包含前缀字符在内的总位数</td>
      <td width='60%'><input name='CardNumLen' type='text' id='CardNumLen' value='12' size='10' maxlength='10'>
        <font color='#0000FF'>建议设为10--15位</font></td>
    </tr>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>充值卡密码位数：</strong></td>
      <td width='60%'><input name='PasswordLen' type='text' id='PasswordLen' value='6' size='10' maxlength='10'>
        <font color='#0000FF'>建议设为6--10位</font></td>
    </tr>
    <tr class='tdbg'>
      <td class="clefttitle"><strong>卡密码构成方式：</strong><br>你可以选择数据或字母的组合</td>
      <td><input type="radio" name="zhtype" value="1" checked>纯数字 <input type="radio" name="zhtype" value="2">数字与字母随机组合 </td>
    </tr>
  </tbody>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>充值卡面值：</strong><br>
      即购买人需要花费的实际金额</td>
      <td width='60%'><input name='Money' type='text' id='Money' value='50' size='10'>
      元</td>
    </tr>
    <tr class='tdbg'> 
      <td width='40%' class="clefttitle"><strong>充值卡点数、资金或有效期：</strong><br>
        购买人可以得到的点数、资金、有效期和积分      </td>
      <td width='60%'><input name='ValidNum' type='text' id='ValidNum' value='50' size='10' maxlength='10'>
        <select name='ValidUnit' id='ValidUnit'>
          <option value='1' selected>点</option>
          <option value='2'>天</option>
          <option value='3'>元</option>
          <option value='4'>积分</option>
        </select></td>
    </tr>
	
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>允许使用此充值卡的用户组：</strong><br>
	  不限制请留空或全部选中。
     </td>
      <td width='60%'><%=KS.GetUserGroup_CheckBox("AllowGroupID","",5)%></td>
    </tr>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>充值后自动归入的用户组：</strong><br>
     </td>
      <td width='60%'><select name="GroupID" id="GroupID">
	  <option value='0'>---保持原有用户组---</option>
	<%=KS.GetUserGroup_Option(0)%>
	 </select></td>
    </tr>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>到期后自动归入的用户组：</strong><br>
	  <span style='color:blue'>指用户选择充值卡为账户充值后,当账户里的点券,有效天数或资金用完后(具体根据该卡是点券卡,有数天数卡或资金卡而定)。将过期的用户自动归入低一级的用户级别。</span>
     </td>
      <td width='60%'><select name="ExpireGroupID" id="ExpireGroupID">
	  <option value='0'>---保持原有用户组---</option>
	<%=KS.GetUserGroup_Option(0)%>
	 </select></td>
    </tr>
	
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>充值截止期限：</strong><br>
      购买人必须在此日期前进行充值，否则自动失效</td>
      <td width='60%' class='tdbg'><input name='EndDate' type='text' id='EndDate' value='<%=dateadd("yyyy",2,now)%>' size='20'></td>
    </tr>
    <tr class='tdbg'> 
      <td height='40' colspan='2' style="text-align:center"><input name='Action' type='hidden' id='Action' value='DoAddMore'> 
        <input  type='submit' class='button' name='Submit' value=' 开始生成 ' style='cursor:pointer;'> 
        &nbsp; <input class='button' name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="window.location.href='KS.Card.asp'" style='cursor:pointer;'></td>
    </tr>
  </table>
</form>
		<%
		End Sub	
		'添加充值卡
		Sub Add()
		  Dim CardNum,PassWord,IsSale,IsUsed,Money,ValidNum,ValidUnit,EndDate,action1,GroupName,AllowGroupID,GroupID,ExpireGroupID,cardtype
		  Dim ID:ID=KS.ChkClng(KS.G("ID"))
		  if KS.g("action")="Edit" then
		    Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			rs.open "select top 1 * from ks_usercard where ID=" & ID,conn,1,1
			if rs.bof and rs.eof then
			  rs.close:set rs=nothing
			  Call KS.AlertHistory("参数传递出错！",-1)
			  Exit sub
			end if
			CardNum=rs("CardNum")
			PassWord=KS.Decrypt(rs("CardPass"))
			Money=rs("money")
			ValidNum=rs("ValidNum")
			ValidUnit=rs("ValidUnit")
			EndDate=rs("EndDate")
			IsSale=rs("IsSale")
			IsUsed=rs("IsUsed")
			cardtype=rs("cardtype")
			GroupName=rs("GroupName")
			AllowGroupID=rs("allowgroupid")
			GroupID=rs("groupid")
			ExpireGroupID=rs("expiregroupid")
			action1="Edit"
		  else
		   IsSale=0:IsUsed=0:Money=50:ValidNum=50:ValidUnit=1:EndDate=Now+365
		  end if
		%>
  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>
		<form method='post' action='KS.Card.asp?action1=<%=action1%>&id=<%=ID%>' name='myform'>
    <tr class='sort'> 
      <td height='22' colspan='2'> <div align='center'><strong>
	  <%IF KS.g("Action")="Edit" then
	   response.write "修 改 充 值 卡"
	    Else
		Response.Write "添 加 充 值 卡"
	    End If
		%></strong></div></td>
    </tr>
    <tr class='tdbg'> 
      <td width='40%' class="clefttitle"><strong>充值卡名称：</strong></td>
      <td width='60%'><input name='GroupName' type='text' size='20' value="<%=GroupName%>" maxlength='100'> 如:“10元购100点充值卡”等</td>
    </tr>
    <tr class='tdbg'> 
      <td width='40%' class="clefttitle"><strong>充值方式：</strong></td>
      <td width='60%'>
	  <input type="radio" name="cardtype" value="1"  onclick="$('#showother').hide()"<%if cardtype=0 then response.write " disabled" else response.write " checked"%>>在线充值购买(会员中心由会员自己购买充值）<br/>
	  <input type="radio" name="cardtype" value="0" onclick="$('#showother').show()"<%if cardtype=1 then response.write " disabled" else response.write " checked"%>>线下销售
	  </td>
    </tr>
   <tbody id="showother" style="display:<%if cardtype=0 then response.write "" else response.write "none"%>">

    <tr class='tdbg'<%if KS.g("action")="Edit" Then response.write " style='display:none'"%>> 
      <td width='40%' class="clefttitle"><strong>添加方式：</strong></td>
      <td width='60%'><input name='AddType' type='radio' value='0' checked onclick="trSingle1.style.display='';trSingle2.style.display='';trBatch.style.display='none';"> 单张充值卡&nbsp;&nbsp;&nbsp;&nbsp;<input name='AddType' type='radio' value='1' onclick="trSingle1.style.display='none';trSingle2.style.display='none';trBatch.style.display='';">批量添加充值卡</td>
    </tr>
    <tr class='tdbg' id='trSingle1'>
      <td width='40%' class="clefttitle"><b>充值卡卡号：</b></td>
      <td><input name='CardNum' type='text' id='CardNum' size='20' value="<%=CardNum%>" maxlength='30'>
        <font color='#0000FF'>建议设为10--15位</font></td>
    </tr>
    <tr class='tdbg' id='trSingle2'>
      <td width='40%' class="clefttitle"><b>充值卡密码：</b></td>
      <td><input name='Password' type='text' id='Password' size='20' value="<%=PassWord%>" maxlength='30'>
        <font color='#0000FF'>建议设为6--10位 </font></td>
    </tr>
    <tr class='tdbg' id='trBatch' style='display:none'>
      <td width='40%' class="clefttitle"><b>格式文本：</b><br><font color='red'>请按照每行一张卡，每张卡按“卡号＋分隔符＋密码”的格式录入</font><br>例：734534759|kSo94Sf4Xs（以“|”作为分隔符）</td>
      <td><textarea name='CardList' rows='10' cols='50'></textarea></td>
    </tr>
	</tbody>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>充值卡面值：</strong><br>
      即购买人需要花费的实际金额</td>
      <td width='60%'><input name='Money' type='text' id='Money' value='<%=formatnumber(Money,2,-1)%>' size='10'>
      元</td>
    </tr>
    <tr class='tdbg'> 
      <td width='40%' class="clefttitle"><strong>充值卡点数、资金或有效期：</strong><br>
        购买人可以得到的点数、资金、有效期和积分      </td>
      <td width='60%'><input name='ValidNum' value="<%=ValidNum%>" type='text' id='ValidNum' size='10' maxlength='10'>
        <select name='ValidUnit' id='ValidUnit'>
          <option value='1'<%if ValidUnit="1" then response.write " selected"%>>点</option>
          <option value='2'<%if ValidUnit="2" then response.write " selected"%>>天</option>
          <option value='3'<%if ValidUnit="3" then response.write " selected"%>>元</option>
          <option value='4'<%if ValidUnit="4" then response.write " selected"%>>积分</option>
        </select></td>
    </tr>
	
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>允许使用此充值卡的用户组：</strong><br>
	  不限制请留空或全部选中。
     </td>
      <td width='60%'><%=KS.GetUserGroup_CheckBox("AllowGroupID",allowgroupid,5)%></td>
    </tr>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>充值后自动归入的用户组：</strong><br>
     </td>
      <td width='60%'><select name="GroupID" id="GroupID">
	  <option value='0'>---保持原有用户组---</option>
	<%=KS.GetUserGroup_Option(groupid)%>
	 </select></td>
    </tr>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>到期后自动归入的用户组：</strong><br>
	  <span style='color:blue'>指用户选择充值卡为账户充值后,当账户里的点券,有效天数或资金用完后(具体根据该卡是点券卡,有数天数卡或资金卡而定)。将过期的用户自动归入低一级的用户级别。</span>
     </td>
      <td width='60%'><select name="ExpireGroupID" id="ExpireGroupID">
	  <option value='0'>---保持原有用户组---</option>
	<%=KS.GetUserGroup_Option(expiregroupid)%>
	 </select></td>
    </tr>
	
	
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>充值截止期限：</strong><br>
      购买人必须在此日期前进行充值，否则自动失效</td>
      <td width='60%' class='tdbg'><input name='EndDate' type='text' id='EndDate' value='<%=EndDate%>' size='20'></td>
    </tr>
	<tr class='tdbg'<%if cardtype=1 then response.write " style='display:none'"%>>
      <td width='40%' class="clefttitle"><strong>是否出售：</strong><br>
      添加新充值卡，请选项未出售</td>
      <td width='60%' class='tdbg'><input name='issale' type='radio' id='issale' value='0'<%if issale=0 then response.write " checked"%>>未出售 <input name='issale' type='radio' id='issale' value='1'<%if issale=1 then response.write " checked"%>>已出售</td>
    </tr>
	<tr class='tdbg'<%if cardtype=1 then response.write " style='display:none'"%>>
      <td width='40%' class="clefttitle"><strong>是否使用：</strong><br>
      添加新充值卡，请选项未使用</td>
      <td width='60%' class='tdbg'><input name='isused' type='radio' id='isused' value='0'<%if isused=0 then response.write " checked"%>>未使用 <input name='isused' type='radio' id='isused' value='1'<%if isused=1 then response.write " checked"%>>已使用</td>
    </tr>
    <tr class='tdbg'> 
      <td height='40' colspan='2' style="text-align:center"><input name='Action' type='hidden' id='Action' value='DoAdd'> 
        <input  type='submit' name='Submit' class='button' value=' <% if KS.g("action")="Edit" then response.write "确定修改" Else Response.write "开始生成" %> ' style='cursor:pointer;'> 
        &nbsp; <input name='Cancel' type='button' class='button' id='Cancel' value=' 取 消 ' onClick="window.location.href='KS.Card.asp?cardtype=<%=cardtype%>'" style='cursor:pointer;'></td>
    </tr>
	</form>

  </table>
		<%
		End Sub
		
		'开始生成充值卡
		Sub DoAdd()
		 Dim AddType:AddType=KS.G("AddType")
		 Dim CardNum:CardNum=KS.G("CardNum")
		 Dim Password:Password=KS.G("Password")
		 Dim CardList:CardList=KS.G("CardList")
		 Dim Money:Money=KS.G("Money")
		 Dim ValidNum:ValidNum=KS.ChkClng(KS.G("ValidNum"))
		 Dim ValidUnit:ValidUnit=KS.G("ValidUnit")
		 Dim EndDate:EndDate=KS.G("EndDate")
		 Dim IsUsed:IsUsed=KS.G("IsUsed")
		 Dim ISSale:IsSale=KS.G("IsSale")
		 Dim GroupName:GroupName=KS.G("GroupName")
		 Dim CardType:CardType=KS.ChkClng(KS.G("CardType"))
		 Dim AllowGroupID:AllowGroupID=KS.G("AllowGroupID")
		 Dim GroupID:GroupID=KS.ChkClng(KS.G("GroupID"))
		 Dim ExpireGroupID:ExpireGroupID=KS.ChkClng(KS.G("expiregroupid"))
		 If GroupName="" Then Call KS.AlertHistory("请输入充值卡名称！",-1):exit sub
		 IF Not IsNumeric(Money) Or money="0" Then Call KS.AlertHistory("充值卡面值，必须大于0",-1):exit sub
		 IF ValidNum=0 Then Call KS.AlertHistory("充值卡点数，必须大于0",-1):exit sub
		 If Not IsDate(EndDate) Then Call KS.AlertHistory("充值截止期限格式不正确!",-1):exit sub
          If AddType=0 or KS.g("action1")="Edit" then
		    if cardtype=0 Then
				if CardNum="" then call KS.AlertHistory("你没有输入充值卡号!",-1):exit sub
				if PassWord=" "then call KS.AlertHistory("你没有输入充值卡密码",-1):exit sub
			end if
			
			   Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			    if KS.g("action1")="Edit" then
				 rs.open "select top 1 * from ks_usercard where id=" & KS.chkclng(KS.g("id")),conn,1,3
				else
					if not conn.execute("select cardnum from ks_usercard where cardnum='" & cardnum & "'").eof then
					  call KS.AlertHistory("你输入的充值卡号已存在，请重输!",-1):exit sub
					end if
				   rs.open "select top 1 * from ks_usercard",conn,1,3
				   rs.addnew
				   rs("AddDate")=now
				   rs("cardtype")=cardtype
			   end if
				 rs("cardnum")=CardNum
				 rs("cardpass")=KS.Encrypt(PassWord)
				 rs("money")=money
				 rs("ValidNum")=ValidNum
				 rs("ValidUnit")=ValidUnit
				 rs("enddate")=EndDate
				 rs("isused")=isused
				 rs("isSale")=issale
				 rs("groupname")=groupname
				 rs("allowgroupid")=allowgroupid
				 rs("groupid")=groupid
				 rs("expiregroupid")=expiregroupid
			   rs.update
			   rs.close:set rs=nothing
		  else 
		    if CardList="" then call KS.AlertHistory("你没有输入充值卡号!",-1):exit sub
			Dim i,j,CardAndPass,CardArr:CardArr=Split(CardList,vbcrlf)
			For I=0 to Ubound(CardArr)
			   CardAndPass=Split(CardArr(I),"|")
			   if not conn.execute("select cardnum from ks_usercard where cardnum='" & CardAndPass(0) & "'").eof then
					 ' call KS.AlertHistory("你输入的充值卡号已存在，请重输!",-1):exit sub
			   else
				   Set RS=Server.CreateObject("adodb.recordset")
				   rs.open "select top 1 * from ks_usercard",conn,1,3
				   rs.addnew
					 rs("cardnum")=CardAndPass(0)
					 rs("cardpass")=KS.Encrypt(CardAndPass(1))
					 rs("money")=money
					 rs("ValidNum")=ValidNum
					 rs("ValidUnit")=ValidUnit
					 rs("AddDate")=now
					 rs("enddate")=EndDate
					 rs("isused")=isused
					 rs("isSale")=issale
					 rs("groupname")=groupname
					 rs("cardtype")=cardtype
					 rs("allowgroupid")=allowgroupid
					 rs("groupid")=groupid
					 rs("expiregroupid")=expiregroupid
				   rs.update
				   rs.close:set rs=nothing
			  end if
			Next
		  end if
		  if KS.g("action1")="Edit" then
			   response.write "<script>alert('修改充值卡成功！');location.href='KS.Card.asp?cardtype=" & cardtype & "';</script>"
		  else
			   response.write "<script>alert('添加充值卡成功！');location.href='KS.Card.asp?cardtype=" & cardtype & "';</script>"
		  end if
		End Sub
		'批量生成充值卡操作
		Sub DoAddMore()
		 Dim Nums:Nums=KS.ChkClng(KS.G("Nums"))
		 Dim CardNumPrefix:CardNumPrefix=KS.G("CardNumPrefix")
		 Dim CardNumLen:CardNumLen=KS.ChkClng(KS.G("CardNumLen"))
		 Dim PasswordLen:PasswordLen=KS.ChkClng(KS.G("PasswordLen"))
		 Dim zhtype:zhtype=KS.G("zhtype")
		 Dim Money:Money=KS.ChkClng(KS.g("money"))
		 Dim ValidNum:ValidNum=KS.ChkClng(KS.G("ValidNum"))
		 Dim ValidUnit:ValidUnit=KS.G("ValidUnit")
		 Dim EndDate:EndDate=KS.G("EndDate")
		 Dim GroupName:GroupName=KS.G("GroupName")
		 Dim CardType:CardType=KS.ChkClng(KS.G("CardType"))
		 Dim AllowGroupID:AllowGroupID=KS.G("AllowGroupID")
		 Dim GroupID:GroupID=KS.ChkClng(KS.G("GroupID"))
		 Dim ExpireGroupID:ExpireGroupID=KS.ChkClng(KS.G("ExpireGroupID"))
		 
		 If GroupName="" Then Call KS.AlertHistory("请给充值卡取个名称!",-1):exit sub
		 If CardType=0 Then
			 IF Nums=0 Then Call KS.AlertHistory("生成充值卡数量，必须大于0",-1):exit sub
			 IF CardNumLen=0 Then Call KS.AlertHistory("充值卡号码长度，必须大于0",-1):exit sub
			 IF PasswordLen=0 Then Call KS.AlertHistory("充值卡密码长度，必须大于0",-1):exit sub
		 End If
		 IF Not IsNumeric(KS.G("money")) Or KS.G("money")=0 Then Call KS.AlertHistory("充值卡面值，必须大于0",-1):exit sub
		 IF ValidNum=0 Then Call KS.AlertHistory("充值卡点数，必须大于0",-1):exit sub
		 If Not IsDate(EndDate) Then Call KS.AlertHistory("充值截止期限格式不正确!",-1):exit sub
		 
		 If CardType=1 Then
		  	   Dim RSObj:Set RSObj=Server.CreateObject("adodb.recordset")
			   rsobj.open "select top 1 * from ks_usercard",conn,1,3
			   rsobj.addnew
				 rsobj("cardnum")=""
				 rsobj("cardpass")=""
				 rsobj("money")=KS.G("money")
				 rsobj("ValidNum")=ValidNum
				 rsobj("ValidUnit")=ValidUnit
				 rsobj("AddDate")=now
				 rsobj("enddate")=EndDate
				 rsobj("isused")=0
				 rsobj("isSale")=0
				 rsobj("groupname")=groupname
				 rsobj("groupid")=groupid
				 rsobj("expiregroupid")=expiregroupid
				 rsobj("allowgroupid")=allowgroupid
				 rsobj("cardtype")=1
			   rsobj.update
			   rsobj.close:set rsobj=nothing
			   Response.Write "<script>alert('恭喜，在线充值卡已生成！');location.href='KS.Card.asp?cardtype=1';</script>"

		 Else
		 %>
					   <br>
				  <table width='300'  border='0' align='center' cellpadding='2' cellspacing='1' class='ctable'>
					<tr class='sort'>
					  <td colspan='2' align='center'><strong>本次生成的点卡信息如下：</strong></td>
					</tr>
					<tr class='tdbg'>
					  <td width='100'>充值卡名称：</td>
					  <td><%=GroupName%></td>
					</tr>
					<tr class='tdbg'>
					  <td width='100'>充值卡数量：</td>
					  <td><%=nums%> 张</td>
					</tr>
					<tr class='tdbg'>
					  <td width='100'>充值卡面值：</td>
					  <td><%=money%> 元</td>
					</tr>
					<tr class='tdbg'>
					  <td width='100'>
					  <% select case ValidUnit
						case 1:response.write "充值卡点数："
						case 2:response.write "充值卡有效天数："
						case 3:response.write "充值卡金额："
						end select
						%></td>
					  <td>
					  <% response.write ValidNum
					  select case validunit
					   case 1:response.write " 点"
					   case 2:response.write " 天"
					   case 3:response.write " 元"
					  end select
					  %>
					  </td>
					</tr>
					<tr class='tdbg'>
					  <td width='100'>充值截止日期：</td>
					  <td><%=enddate%></td>
					</tr>
					
				</table>
				<br>
				<table width='300' border='0' align='center' cellpadding='2' cellspacing='1' class='ctable'>
			  <tr align='center' class='sort'>
				<td  width=150 height='22'><strong> 卡 号 </strong></td>
				<td  width=150 height='22'><strong> 密 码 </strong></td>
			  </tr>
			 <%
			 Dim n,currcard,CurrCardPass
			 For N=1 To Nums
			   CurrCard=KS.MakeRandom(CardNumLen-len(CardNumPrefix))
			   CurrCard=CardNumPrefix & CurrCard
			   If ZhType=2 then
				 CurrCardPass=KS.GetRndPassword(PasswordLen)
			   Else
				 CurrCardPass=KS.MakeRandom(PasswordLen)
			   End If
			   Do While not Conn.execute("select CardNum From KS_UserCard Where CardNum='" & CurrCard & "'").eof 
				   CurrCard=KS.MakeRandom(CardNumLen-len(CardNumPrefix))
				   CurrCard=CardNumPrefix & CurrCard
			   loop
			   
			   response.write "<tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbcrlf
			   response.write "<td height='22'>" & CurrCard & "</td>" & vbcrlf
			   response.write "<td>" & CurrCardPass & "</td>" & vbcrlf
			   response.write "</tr>" & vbcrlf
			   Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			   rs.open "select top 1 * from ks_usercard",conn,1,3
			   rs.addnew
				 rs("cardnum")=CurrCard
				 rs("cardpass")=KS.Encrypt(CurrCardPass)
				 rs("money")=money
				 rs("ValidNum")=ValidNum
				 rs("ValidUnit")=ValidUnit
				 rs("AddDate")=now
				 rs("enddate")=EndDate
				 rs("isused")=0
				 rs("isSale")=0
				 rs("groupname")=groupname
				 rs("groupid")=groupid
				 rs("expiregroupid")=expiregroupid
				 rs("allowgroupid")=allowgroupid
				 rs("cardtype")=0
			   rs.update
			   rs.close:set rs=nothing
			 Next
			  response.write "</table>"
         End If
	End SUb	
End Class
%> 
