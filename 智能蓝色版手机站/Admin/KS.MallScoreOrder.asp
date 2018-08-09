<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.UpFileCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_MallScoreOrder
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_MallScoreOrder
        Private KS,Param,KSCls
		Private Action,i,strClass,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		 With Response
					If Not KS.ReturnPowerResult(0, "KSMS20010") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
			  .Write "<html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../KS_Inc/jQuery.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('积分兑换系统 >> <font color=red>添加商品</font>')+'&ButtonSymbol=GOSave';location.href='KS.MallScore.asp?action=Add';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加商品</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.MallScoreOrder.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>处理兑换订单</span></li>"
			  .Write "<li style='margin-left:30px;margin-top:10px'><strong>查看方式:</strong><a href=""KS.MallScoreOrder.asp"">所有订单</a> <a href=""KS.MallScoreOrder.asp?flag=1"">已审核</a>  <a href=""KS.MallScoreOrder.asp?flag=-1"">未审核</a> <a href=""KS.MallScoreOrder.asp?flag=2"">配货中</a> <a href=""KS.MallScoreOrder.asp?flag=3"">已发货</a> <a href=""KS.MallScoreOrder.asp?flag=4"">已完成</a></li>"

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
		   Param= Param & " and b.ProductName like '%" & KS.G("KeyWord") & "%'"
		  ElseIf KS.G("condition")=2 Then
		   Param= Param & " and a.OrderID like '%" & KS.G("KeyWord") & "%'"
		  Else
		   Param= Param & " and a.RealName like '%" & KS.G("KeyWord") & "%'"
		  End If
		End If
		If KS.G("Flag")<>"" Then
		  If KS.G("Flag")="-1" Then 
		    Param=Param & " and a.Status=0"
		  Else
		   Param=Param & " and a.Status=" & KS.ChkClng(KS.G("Flag"))
		  End If
		End If
		If KS.S("ProductID")<>"" Then Param=Param & " and a.productid=" & KS.ChkClng(KS.G("ProductID"))

		totalPut = Conn.Execute("Select Count(id) From KS_MallScoreOrder a " & Param)(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Select Case KS.G("action")
		 Case "Add","Edit" Call ProductNameManage()
		 Case "EditSave" Call DoSave()
		 Case "Del"  Call OrderDel()
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
If KS.S("ProductID")<>"" Then
 %>
  <div style="height:45px;font-size:14px;line-height:45px;text-align:center;font-weight:bold">查看商品 <font color=red>[<%=LFCls.GetSingleFieldValue("Select ProductName From KS_MallScore Where ID=" & KS.ChkClng(KS.G("ProductID")))%>]</font> 的兑换记录</div>
 <%end If%>
<table width="100%" border="0" align="center" style="border-top:1px solid #cccccc" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td nowrap>订单号</th>
	<td nowrap>商品名称</th>
	<td nowrap>兑换人</th>
	<td nowrap>兑换时间</th>
	<td nowrap>兑换数量</th>
	<td nowrap>送货方式</th>
	<td nowrap>订单状态</th>
	<td nowrap>管理操作</th>
</tr>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select A.*,B.ProductName from KS_MallScoreOrder A left join KS_MallScore b on a.productid=b.id " & Param & " order by a.id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=10>对不起,找不到订单！</td></tr>"
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
	<td align="center" class="splittd"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td  class="splittd"><font color=green><%=Rs("orderid")%></font></td>
	<td  class="splittd"><%if Rs("ProductName")="" or isnull(Rs("ProductName")) Then Response.write "<font color=#999999>已删除</font>" Else Response.Write RS("ProductName")%></td>
	<td align="center" class="splittd"><%=Rs("username")%></td>
	<td align="center" class="splittd"><%=Rs("AddDate")%></td>
	<td align="center" class="splittd"><font color=#cccccc><%=RS("amount")%>  件</font></td>
	<td align="center" class="splittd">
	<%
	 if rs("DeliveryType")=1 then
	  response.write "快递到付"
	 else
	  response.write "自取"
	 end if
	%>
	</td>
	
	<td align="center" class="splittd"><%
		select case  rs("status")
		 case 1
		  response.write "已审"
		 case 2
		  response.write "<font color=blue>配货中</font>"
		 case 3
		  response.write "<font color=#ff6600>已发货</font>"
		 case 4
		  response.write "<font color=#999999>交易完成</font>"
		 case 5
		  response.write "<font color=green>无效订单(积分退回)</font>"
		 case else
		  response.write " <font color=red>未审</font>"
		end select
	%></td>
	<td align="center" class="splittd"><a href="?action=Edit&ID=<%=RS("ID")%>"  onclick="window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('积分兑换系统 >> <font color=red>修改团购信息</font>')+'&ButtonSymbol=GOSave';">查看/修改</a> <a href="?Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('此操作不可逆,确定删除该订单吗？'));">删除</a> 
		

	</td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class=Button type="submit" name="Submit2" value=" 删除选中的订单 " onclick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){this.document.selform.submit();return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  colspan=7 align=right>
	<%
	Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>
<div>
<form action="KS.MallScoreOrder.asp" name="myform" method="get">
   <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
      &nbsp;<strong>快速搜索=></strong>
	 &nbsp;关键字:<input type="text" class='textbox' name="keyword">&nbsp;条件:
	 <select name="condition">
	  <option value=1>按商品名称</option>
	  <option value=2>按订单号</option>
	  <option value=3>按收货人</option>
	 </select>
	  &nbsp;<input type="submit" value="开始搜索" class="button" name="s1">
    </div>
</form>
</div>
<%
End Sub

Sub ProductNameManage()
Dim ProductName,ActiveDate,AddDate,DeliveryType,Amount,Score,Telphone,RealName,ZipCode,Protection,BuyFlow,Notes,Tel,Status,Address,Email,Remark,UserName
Dim ID:ID=KS.ChkClng(KS.G("ID"))
Dim RS:Set RS=server.createobject("adodb.recordset")
If KS.G("Action")="Edit" Then
	RS.Open "Select a.*,b.productname,score From KS_MallScoreOrder a Left Join KS_MallScore b on a.productid=b.id Where a.ID=" & ID,conn,1,1
	 If RS.Eof And RS.Bof Then
	  RS.Close:Set RS=Nothing
	  Response.Write "<script>alert('参数传递出错！');history.back();</script>"
	  Response.End
	 Else
	   ProductName=RS("ProductName")
	   If KS.IsNul(ProductName) Then ProductName="已删除"
	   AddDate=RS("AddDate")
	   DeliveryType=RS("DeliveryType")
	   Amount=RS("Amount")
	   Score=RS("Score")
	   RealName=RS("RealName")
	   Address=RS("Address")
	   ZipCode=RS("ZipCode")
	   Tel=RS("Tel")
	   Email=RS("Email")
	   Remark=RS("Remark")
	   UserName=RS("UserName")
	   Status=RS("Status")
	 End If
Else
  AddDate=Now
  DeliveryType=Now+30
  ZipCode=0:Score=10
  Tel=0:Status=1
  Amount=100
  RealName=" "
  Address="../images/nopic.gif"
 End If
%>
<script>
function CheckForm()
{
	if ($('#RealName').val()=='')
	{
	 alert('请输入收货人!');
	 $("#RealName").focus();
	 return false;
	}

document.myform.submit();
}
</script>
<br>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
  <form name="myform" action="?action=EditSave" method="post">
    <input type="hidden" value="<%=Score%>" name="Score"/>
    <input type="hidden" value="<%=ID%>" name="id" />
    <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl" />
       <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>修改订单状态：</strong></td>
      <td height='30'>&nbsp;
	  <%if Status<>5 then%>
          <input type="radio" name="Status" value="0"<%if Status=0 then response.write " checked"%> />
        未通过审核
        <input type="radio" name="Status" value="1"<%if Status=1 then response.write " checked"%> />
        通过审核
        <input type="radio" name="Status" value="2"<%if Status=2 then response.write " checked"%> />
        配货中
        <input type="radio" name="Status" value="3"<%if Status=3 then response.write " checked"%> />
        已发货
        <input type="radio" name="Status" value="4"<%if Status=4 then response.write " checked"%> />
        交易完成
        <label style="color:green"><input type="radio" name="Status" onclick="alert('注意:一旦确定设置成此状态后,将不能再设置成其它状态!')" value="5"<%if Status=5 then response.write " checked"%> />
        无效订单并退回积分</label>
	<%else%>
	   <input type="hidden" name="status" value="-1">
	   <label style="color:green">无效订单并退回积分</label>
	 <%end if%>

		
				</td>
    </tr>
 <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>兑换商品：</strong></td>
      <td width="781" height='30'>&nbsp;
          <%=ProductName%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td height='30' align='right' class='clefttitle'><strong>兑换用户：</strong></td>
      <td height='30'>&nbsp;<%=username%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>兑换时间：</strong></td>
      <td height='30'>&nbsp;<%=adddate%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>兑换数量：</strong></td>
      <td height='30'>&nbsp;<%=amount%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td height='30' align='right' class='clefttitle'><strong>配送方式：</strong></td>
      <td height='30'>&nbsp;
<input type="radio" name="DeliveryType" value="1"<%if DeliveryType=1 then response.write " checked"%> />
快递到付
  <input type="radio" name="DeliveryType" value="2"<%if Status=2 then response.write " checked"%> />
自取 </td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>收 货 人：</strong></td>
      <td height='30'>&nbsp;
          <input type='text' name='RealName' id='RealName' value='<%=RealName%>' size="20" />
          <font color=red>*</font></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>联系电话：</strong></td>
      <td height='30'>&nbsp;
        <input type='text' name='Tel' value='<%=Tel%>' size="20" />
        <font color="red">*</font></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>收货地址：</strong></td>
      <td height='30'>&nbsp;
          <input type='text' name='Address' value='<%=Address%>' size="35" />        
        <font color=red>*</font></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>邮政编码：</strong></td>
      <td height='30'>&nbsp;
          <input name='ZipCode' value="<%=ZipCode%>" size="10" /></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>电子邮箱：</strong></td>
      <td height='30'>&nbsp;
        <input type='text' name='Email' value='<%=Email%>' size="25" /></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td height='30' align='right' class='clefttitle'><strong>备注说明：</strong></td>
      <td height='30'>&nbsp;
        <textarea name='Remark' style="width:400px;height:80px"><%=Remark%></textarea></td>
    </tr>
  </form>
</table>
<%
End Sub

Sub DoSave()
       Dim ID:ID=KS.ChkClng(KS.G("id"))
	   Dim Address:Address=KS.G("Address")
	   Dim RealName:RealName=KS.G("RealName")
	   Dim ZipCode:ZipCode=KS.G("ZipCode")
	   Dim Tel:Tel=KS.G("Tel")
	   Dim Status:Status=KS.ChkClng(KS.G("Status"))
	   Dim ComeUrl:ComeUrl=KS.G("ComeUrl")
	   Dim Remark:Remark=KS.G("Remark")
	   Dim Email:Email=KS.G("Email")
	   Dim DeliveryType:DeliveryType=KS.ChkClng(KS.G("DeliveryType"))
	   
	   If RealName="" Then Response.Write "<script>alert('收货人必须输入');history.back();</script>":response.end

            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select top 1 * From KS_MallScoreOrder Where ID=" & ID,Conn,1,3
				 RS("DeliveryType")=DeliveryType
				 RS("RealName")=RealName
				 RS("Address")=Address
				 RS("ZipCode")=ZipCode
				 RS("Tel")=Tel
				 RS("Remark")=Remark
				 RS("Email")=Email
				 IF (Status<>-1) then
				 RS("Status")=Status
				 end if
		 		 RS.Update
				 RS.MoveLast
				if Status=5 then
				   '更新用户积分
				   Session("ScoreHasUse")="-" '设置只累计消费积分
				   Call KS.ScoreInOrOut(RS("UserName"),1,KS.ChkClng(KS.G("Score"))*RS("Amount"),"系统","礼品兑换失败，返回兑换订单号<font color=red>" & RS("OrderID") & "</font>的礼品积分!",0,0)
				end if
				 
				 
			     RS.Close
				 Set RS=Nothing
				 
  Response.Write "<script>alert('兑换订单修改成功！');parent.frames['BottomFrame'].location.href='KS.Split.asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("积分兑换系统 >> <font color=red>订单管理</font>") & "';location.href='"& ComeUrl & "';</script>"

EnD Sub

'删除
Sub OrderDel()
 Dim ID:ID=KS.FilterIds(KS.G("ID"))
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "SELECT a.*,b.score FROM KS_MallScoreOrder a inner join KS_MallScore b On a.ProductID=b.id Where a.id In("& id & ")",conn,1,1
 Do While Not RS.Eof 
  If rs("Status")=0 Then
	Session("ScoreHasUse")="-" '设置只累计消费积分
	Call KS.ScoreInOrOut(RS("UserName"),1,KS.ChkClng(rs("Score"))*RS("Amount"),"系统","礼品兑换失败，返回订单号<font color=red>" & RS("OrderID") & "</font>的积分!",0,0)
  End If
  RS.MoveNext
 Loop
 RS.Close
 Set RS=Nothing
 Conn.execute("Delete From KS_MallScoreOrder Where id In("& id & ")")
 Response.Write "<script>alert('删除成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub



End Class
%> 
