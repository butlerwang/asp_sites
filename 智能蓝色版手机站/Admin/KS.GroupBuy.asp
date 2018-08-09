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
Set KSCls = New Admin_GroupBuy
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_GroupBuy
        Private KS,Param,KSCls
		Private Action,i,strClass,sFileName,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
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
		 %>
		 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

		 <%
			  '.Write "<html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../KS_Inc/kesion.box.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<div id='menu_top'>"
			  .Write "<li class='parent' onclick=""window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('团购系统 >> <font color=red>添加团购商品</font>')+'&ButtonSymbol=GOSave';location.href='?action=Add';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加团购</span></li>"
			  .Write "<li class='parent' onclick=""window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('团购系统 >> <font color=red>团购分类管理</font>')+'&ButtonSymbol=Disabled';location.href='?action=ClassManage';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>分类管理</span></li>"
			  .Write "<li class='parent' onclick=""window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('团购系统 >> <font color=red>添加团购分类</font>')+'&ButtonSymbol=Disabled';location.href='?action=AddClass';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addfolder.gif' border='0' align='absmiddle'>添加分类</span></li>"
			 
			  .Write "<li style='margin-left:30px;line-height:32px;'>&nbsp;&nbsp;<strong>查看方式:</strong><a href=""KS.GroupBuy.asp"">所有团购</a> <a href=""KS.GroupBuy.asp?flag=1"">进行中的团购</a>  <a href=""KS.GroupBuy.asp?flag=2"">已结束的团购</a> <a href=""KS.GroupBuy.asp?flag=3"">已锁定的团购</a></li>"

			  .Write "</div>"
		End With
		
		   	 If Not KS.ReturnPowerResult(5, "M530001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
			 End If

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
		   Param= Param & " and Subject like '%" & KS.G("KeyWord") & "%'"
		  Else
		   Param= Param & " and Intro like '%" & KS.G("KeyWord") & "%'"
		  End If
		End If
		If KS.G("Flag")<>"" Then
		  If KS.G("Flag")="1" Then Param=Param & " and locked=0 and endtf=0"
		  If KS.G("Flag")="2" Then Param=Param & " and endtf=1"
		  If KS.G("Flag")="3" Then Param=Param & " and locked=1"
		  
		End If

		totalPut = Conn.Execute("Select Count(id) From KS_GroupBuy " & Param)(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Select Case KS.G("action")
		 Case "Add","Edit" Call SubjectManage()
		 Case "EditSave" Call DoSave()
		 Case "Del"  Call GroupBuyDel()
		 Case "Recommend" Call Recommend()
		 Case "UnRecommend" Call UnRecommend()
		 Case "lock"  Call GroupBuyLock()
		 Case "unlock" Call GroupBuyUnLock()
		 Case "endtf"  Call GroupBuyendtf()
		 Case "Cancelendtf" Call GroupBuyCancelendtf()
		 Case "AddClass" Call AddClass()
		 Case "AddClassSave" Call AddClassSave()
		 Case "ClassManage" Call ClassManage()
		 Case "DelClass" Call DelClass()
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
%>
<script type="text/javascript">
function ShowSale(id,title)
 { new parent.KesionPopup().PopupCenterIframe("查看商品销售详情","KS.ShopProSale.asp?proid="+id+"&title="+escape(title),760,450,"auto")}
</script>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td nowrap>团购主题</th>
	<td nowrap>开始/结束时间</th>
	<td nowrap>原价</th>
	<td nowrap>现价</th>
	<td nowrap>团购状态</th>
	<td nowrap>管理操作</th>
</tr>
<%
	sFileName = "KS.GroupBuy.asp?"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_GroupBuy " & Param & " order by id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>还没有添加任何团购信息！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action="KS.GroupBuy.asp">
<input type="hidden" name="action" id="action" value=""/>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="30"  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td align="center" class="splittd"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td class="splittd"><a href='../shop/groupbuyshow.asp?id=<%=rs("id")%>' target='_blank'><%=KS.Gottopic(Rs("Subject"),35)%></a>
	<%If rs("recommend")="1" then response.write " <font color=green>荐</font>"%><br/>
	<span style='color:#999'>[总销量：<font color=blue><%=KS.ChkClng(conn.execute("select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.ordertype=1 and i.proid=" & rs("id"))(0))%></font> 件，已付：<font color=green><%=ks.chkclng(conn.execute("select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.ordertype=1 and o.MoneyReceipt>0 and i.proid=" & rs("id"))(0))%> </font>件，未付：<font color=red><%=ks.chkclng(conn.execute("select sum(amount) from ks_orderitem i inner join ks_order o on i.orderid=o.orderid where o.ordertype=1 and o.MoneyReceipt<=0 and i.proid=" & rs("id"))(0))%></font> 件]</span>
		</td>
	<td align="center" width="120" class="splittd">
	<%=Rs("adddate")%><br/> 至<br/> <%=Rs("ActiveDate")%>
	</td>
	
	<td align="center" class="splittd">
	<span style='color:#999999;text-decoration:line-through;'><%=rs("price_original")%> 元</span>
	</td>
	<td align="center" class="splittd">
	<span style='color:#ff6600'><%=rs("price")%> 元</span>
	</td>
	<td align="center" class="splittd">
	<%
	if DateDiff("s",now,RS("AddDate"))>0 Then
		response.write " <font color=green>未开始</font>"
	ElseIf DateDiff("s",now,RS("ActiveDate"))<0 Then
		response.write " <font color=#cccccc>已结束</font>"
	elseif rs("locked")=0 and rs("endtf")=0 then
	 response.write "<font color=red>进行中</font>"
	else
		if rs("locked")=1 then
		  response.write "<font color=blue>锁定</font>"
		end if
		if rs("endtf")=1 then
		  response.write " <font color=#cccccc>已结束</font>"
		end if
	end if
	%></td>
	<td align="center" class="splittd"><a href="?action=Edit&ID=<%=RS("ID")%>"  onclick="window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('团购系统 >> <font color=red>修改团购信息</font>')+'&ButtonSymbol=GOSave';">修改</a> <a href="?Action=Del&ID=<%=rs("id")%>" onClick="return(confirm('确定删除该团购吗？'));">删除</a> 
	
	&nbsp;<%if rs("locked")=0 then%><a href="?Action=lock&id=<%=rs("id")%>">锁定</a><%else%><a href="?Action=unlock&id=<%=rs("id")%>">解锁</a><%end if%>
		
		&nbsp;<%IF rs("endtf")="1" then %><a href="?Action=Cancelendtf&id=<%=rs("id")%>"><font color=red>打开</font></a><%else%><a href="?Action=endtf&id=<%=rs("id")%>">结束</a><%end if%>
		
		<a href="javascript:ShowSale(<%=rs("id")%>,'<%=rs("subject")%>');">销售</a>
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
	<input class="button" type="submit" name="Submit2" value=" 删除选中的团购 " onClick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){$('#action').val('Del');this.document.selform.submit();return true;}return false;}">
	<input class="button" type="submit" name="Submit2" value=" 批量推荐 " onClick="$('#action').val('Recommend');this.document.selform.submit();return true;">
	<input class="button" type="submit" name="Submit2" value=" 批量取消推荐 " onClick="$('#action').val('UnRecommend');this.document.selform.submit();return true;">
	
	
	
	</td>
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
<form action="KS.GroupBuy.asp" name="myform" method="get">
   <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
      &nbsp;<strong>快速搜索=></strong>
	 &nbsp;关键字:<input type="text" class='textbox' name="keyword">&nbsp;条件:
	 <select name="condition">
	  <option value=1>按商品名称</option>
	  <option value=2>按商品介绍</option>
	 </select>
	  &nbsp;<input type="submit" value="开始搜索" class="button" name="s1">
    </div>
</form>
</div>
<%
End Sub

Sub SubjectManage()
Dim Subject,ActiveDate,AddDate,Intro,Highlights,Protection,Notes,Locked,EndTF,PhotoUrl,BigPhoto,ClassID,AllowBMFlag,AllowArrGroupID,minnum,Comment
Dim Price_Original,Price,Discount,limitbuynum,weight,recommend,ProvinceID,CityID,HasBuyNum,MustPayOnline,CleanCart,showdelivery
Dim ID:ID=KS.ChkClng(KS.G("ID"))
Dim RS:Set RS=server.createobject("adodb.recordset")
If KS.G("Action")="Edit" Then
	RS.Open "Select top 1 * From KS_GroupBuy Where ID=" & ID,conn,1,1
	 If RS.Eof And RS.Bof Then
	  RS.Close:Set RS=Nothing
	  Response.Write "<script>alert('参数传递出错！');history.back();</script>"
	  Response.End
	 Else
	   Subject=RS("Subject")
	   Price_Original=RS("Price_Original")
	   Price=RS("Price")
	   Discount=RS("Discount")
	   ActiveDate=RS("ActiveDate")
	   AddDate=RS("AddDate")
	   Intro=RS("Intro")
	   PhotoUrl=RS("PhotoUrl")
	   BigPhoto=RS("BigPhoto")
	   Highlights=RS("Highlights")
	   Protection=RS("Protection")
	   ClassID=RS("ClassID")
	   Notes=RS("Notes")
	   Locked=RS("Locked")
	   EndTF=RS("EndTF")
	   Comment=RS("Comment")
	   AllowArrGroupID=RS("AllowArrGroupID")
	   AllowBMFlag=RS("AllowBMFlag")
	   minnum=RS("minnum")
	   limitbuynum=RS("limitbuynum")
	   Weight=RS("Weight")
	   recommend=RS("recommend")
	   ProvinceID=RS("ProvinceID")
	   CityID=RS("CityID")
	   HasBuyNum=RS("HasBuyNum")
	   MustPayOnline=RS("MustPayOnline")
	   CleanCart=RS("CleanCart")
	   showdelivery=RS("showdelivery")
	 End If
Else
  AllowBMFlag=0:Comment=1
  AddDate=Now: MustPayOnline=1 : CleanCart=1 : showdelivery=0
  ActiveDate=Now+10
  Locked=0:EndTF=0 :minnum=0:recommend=0:HasBuyNum=0
  Intro=" ":ProvinceID=0:CityID=0
 End If
%>
<script>
function CheckForm()
{
	if ($('#Subject').val()=='')
	{
	 alert('请输入团购主题!');
	 $("#Subject").focus();
	 return false;
	}
	if ($('#ClassID').val()=='')
	{
	 alert('请选择团购分类!');
	 $("#ClassID").focus();
	 return false;
	}

	if (CKEDITOR.instances.Intro.getData()=="")
	{
	 alert('请输入本单详情!');
	 CKEDITOR.instances.Intro.focus();
	 return false;
	}
	if ($("#Price_Original").val()=='')
	{
	 alert('请输入原价!');
	 $("#Price_Original").focus();
	 return false;
	}
	if ($("#Discount").val()=='')
	{
	 alert('请输入折扣！');
	 $("#Discount").focus();
	 return false;
	}
	if (parseFloat($("#Discount").val())>10){
	 alert('折扣不能大于10！');
	 $("#Discount").focus();
	 return false;
	}
	if ($("#Price").val()=='')
	{
	 alert('请输入团购价！');
	 $("#Price").focus();
	 return false;
	}
document.myform.submit();
}
function regInput(obj, reg, inputStr)
{
		var docSel = document.selection.createRange()
		if (docSel.parentElement().tagName != "INPUT")    return false
		oSel = docSel.duplicate()
		oSel.text = ""
		var srcRange = obj.createTextRange()
		oSel.setEndPoint("StartToStart", srcRange)
		var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
		return reg.test(str)
}
function getprice(discount){
     if (parseFloat(discount)>10){
	 alert('折扣不能大于10！');
	 $("#Discount").val(10);
	 return false;
	 }
     var Price_Original=$("#Price_Original").val();
	 if(Price_Original==''|| isNaN(Price_Original)){Price_Original=0;}
	 document.myform.Price.value=Math.round(Price_Original*(discount/10));
  }
</script>
<br>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
 <form name="myform" action="?action=EditSave" method="post">
   <input type="hidden" value="<%=ID%>" name="id">
   <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl">
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>团购主题：</strong></td>
            <td width="435" height='30'>&nbsp;<input class='textbox' type='text' name='Subject' id='Subject' value='<%=Subject%>' size="40"> <font color=red>*</font></td>
           <td width="217" rowspan="7" style="text-align:center"><div  style="margin:0 auto;filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(sizingMethod=scale);height:100px;width:95px;border:1px solid #777777">
						<img src="<%=PhotoUrl%>" onerror="this.src='../images/logo.png';" id="pic" style="height:100px;width:95px;">
		    </div></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td height='30' align='right' class='clefttitle'><strong>团购分类：</strong></td>
            <td height='30'>&nbsp;<select name="ClassID" class="ClassID">
			<option value='0'>---选择分类---</option>
			<%Dim RSC:Set RSC=Conn.Execute("select * From KS_GroupBuyClass Order By OrderID,ID")
			Do While Not RSC.Eof
			  If KS.ChkClng(ClassID)=RSC("ID") Then
			   Response.Write "<option value='" & RSC("ID") & "' selected>" & RSC("CategoryName") & "</option>"
			  Else
			   Response.Write "<option value='" & RSC("ID") & "'>" & RSC("CategoryName") & "</option>"
			  End If
			  RSC.MoveNext
			Loop
			RSC.Close
			Set RSC=Nothing
			%>
			</select></td>
          </tr>
		  
          <tr style="display:none" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>地区：</strong></td>
            <td height='30'>&nbsp;<script src="../plus/area.asp?flag=getid"></script> <span style='color:red'>tips:地区不选择的话该团购切换所有地区都会显示</span>
			<script type="text/javascript">
			<%if KS.ChkClng(ProvinceID)<>0 then%>
				  $('#Province').val('<%=provinceid%>');
			<%end if%>
			 <%if KS.ChkClng(CityID)<>0 Then%>
				$('#City').val(<%=CityID%>);
			<%end if%>
			</script>
			</td>
          </tr> 


          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>时间设置：</strong></td>
            <td height='30'>&nbsp;开始<input type='text' class='textbox' name='AddDate' value='<%=AddDate%>' size="20" /> 
			结束：<input type='text' class='textbox' name='ActiveDate' value='<%=ActiveDate%>' size="20" /> <br/>
            &nbsp;<span class='tips'>如2012-10-1 10:10</span> </td>
          </tr> 

		  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td height='30' align='right' class='clefttitle'><strong>购物车设置：</strong></td>
            <td height='30'>
			&nbsp;需要在线支付订单才生效：<label><input type="radio" name="MustPayOnline" value="0"<%if MustPayOnline="0" then response.write " checked"%>/>不需要</label>
			<label><input type="radio" name="MustPayOnline" value="1"<%if MustPayOnline="1" then response.write " checked"%>/>需要</label><br/> &nbsp;<span class="tips">如凭订单号享受打折的团购，建议选择不需要在线支付。</span>

			<br/>&nbsp;当购物车里有商品时先清空：<label><input type="radio" onClick="$('#delivery').show();" name="cleancart" value="1"<%if cleancart="1" then response.write " checked"%>/>是</label>
			<label><input type="radio" name="cleancart" onClick="$('#delivery').hide();" value="0"<%if cleancart="0" then response.write " checked"%>/>否</label>
			<br/>&nbsp;<span class="tips">当选择购物车里有商品时先清空，则订单里只能有这件商品。</span>
			<%if cleancart="1" then%>
			<div style="" id="delivery">
			<%else%>
			<div style="display:none" id="delivery">
			<%end if%>
			&nbsp;显示送货方式：<label><input type="radio" name="showdelivery" value="1"<%if showdelivery="1" then response.write " checked"%>/>显示</label><label><input type="radio" name="showdelivery" value="0"<%if showdelivery="0" then response.write " checked"%>/>不显示</label>
			<br/>
			 &nbsp;<span class="tips">如本地商家打折等团购建议选择不显示</span>
			</div>
			</td>
          </tr>
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		     <td height='30' align='right' class='clefttitle'><strong>商品图片：</strong></td>
		     <td height='30'>小图：<input class="textbox"  type="text" name="PhotoUrl" id="PhotoUrl" size="30" value="<%=photourl%>" /> <input class="button" type='button' name='Submit' value='选择小图...' onClick="OpenThenSetValue('Include/SelectPic.asp?ChannelID=5&CurrPath=<%=KS.GetUpFilesDir()%>',550,290,window,document.myform.PhotoUrl,'pic');">
			 <br/>
			 大图：<input value="<%=bigphoto%>" class="textbox" type="text" name='BigPhoto' type='text' id='BigPhoto' size="30" /> <input class="button" type='button' name='Submit' value='选择大图...' onClick="OpenThenSetValue('Include/SelectPic.asp?ChannelID=5&CurrPath=<%=KS.GetUpFilesDir()%>',550,290,window,document.myform.BigPhoto,'pic');">
  </tr>
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		     <td height='30' align='right' class='clefttitle'><strong>上传图片：</strong></td>
			 <td colspan="2">
 <iframe id='UpPhotoFrame' name='UpPhotoFrame' src='KS.UpFileForm.asp?showpic=pic&ChannelID=5&UpType=Pic' frameborder=0 scrolling=no width='100%' height='30'></iframe>			 </td>
		</tr>

 
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		     <td height='30' align='right' class='clefttitle'><strong>价格设置：</strong></td>
		     <td height='30' colspan="2">原价<input type="text" class="textbox" onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))" name="Price_Original" id="Price_Original" size="6" value="<%=Price_Original%>" style="text-align:center" />元 折扣<input class="textbox" onChange="getprice(this.value);" type="text" name="Discount" id="Discount" onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))" size="6" value="<%=Discount%>" style="text-align:center" />折  团购价<input type="text" name="Price" id="Price" size="6" value="<%=Price%>" class="textbox" onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))" style="text-align:center" />元
			 
			 重量：<input class="textbox" type='text' name='Weight' style="text-align:center" id='Weight' value='<%=Weight%>' size="6">KG
			 <span style='color:#999999'>计算运费用的,包邮请输入0。</span></td>
  </tr>
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>最低人数：</strong></td>
           <td height='30' colspan="2">&nbsp;<input class="textbox" type='text' name='minnum' style="text-align:center" id='minnum' value='<%=minnum%>' size="6"> 人 &nbsp;每人限制购买<input class="textbox" type='text' name='limitbuynum' style="text-align:center" id='limitbuynum' value='<%=limitbuynum%>' size="6"> 件 <font color=red>*</font> 不限制输入0  初始已销售<input type='text' name='hasbuynum' style="text-align:center" class="textbox" id='hasbuynum' value='<%=hasbuynum%>' size="6"> 件 <span style='color:#999999'>(作弊用的)</span> </td>
          </tr>  
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>本单详情：</strong></td>
            <td height='30' colspan="2"><%			
		    Response.Write "<textarea id=""Intro"" name=""Intro"" style=""display:none"">" &  server.HTMLEncode(Intro&"") &"</textarea>"
			%>	
			<script type="text/javascript" src="../editor/ckeditor.js" mce_src="../editor/ckeditor.js"></script>
			<script type="text/javascript">
             CKEDITOR.replace('Intro', {width:"98%",height:"160px",toolbar:"Basic",filebrowserBrowseUrl :"Include/SelectPic.asp?from=ckeditor&Currpath=<%=KS.GetUpFilesDir()%>",filebrowserWindowWidth:650,filebrowserWindowHeight:290});
			 </script>			</td>
          </tr>
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>精彩卖点：</strong></td>
           <td height='30' colspan="2">&nbsp;<textarea name='Highlights' cols="60" rows="4"><%=Highlights%></textarea></td>
          </tr>  
		   <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>团购保障：</strong></td>
           <td height='30' colspan="2">&nbsp;<textarea name='Protection' cols="60" rows="4"><%=Protection%></textarea></td>
          </tr>  
		  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>温馨提示：</strong></td>
            <td height='30' colspan="2">&nbsp;<textarea name='Notes' cols="60" rows="4"><%=Notes%></textarea></td>
          </tr>  
		  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>允许参加团购的权限：</strong></td>
            <td height='30' colspan="2">	    
		    <label><input type="radio" name="AllowBMFlag" value="0"<%if AllowBMFlag=0 then response.write " checked"%>>允许所有人报名参加,包括游客</label><br/>
			<label><input type="radio" name="AllowBMFlag" value="1"<%if AllowBMFlag=1 then response.write " checked"%>>只允许会员报名参加</label>
			<br/><label><input type="radio" name="AllowBMFlag" value="2"<%if AllowBMFlag=2 then response.write " checked"%>>只允许指定的会员组报名参加</label>			</td>
          </tr>  
		  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>允许参加团购的会员组：</strong><br/>
            <font color=blue>当上面选择只允许指定的会员组参加时，请在此指定会员组</font></td>
           <td height='30' colspan="2">&nbsp;<%=KS.GetUserGroup_CheckBox("AllowArrGroupID",AllowArrGroupID,5)%>			</td>
          </tr> 
		  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>是否推荐：</strong></td>
           <td height='30' colspan="2">&nbsp;
		    <input type="radio" name="recommend" value="0"<%if recommend=0 then response.write " checked"%>>否
			<input type="radio" name="recommend" value="1"<%if recommend=1 then response.write " checked"%>>是		   </td>
          </tr>  
		  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>是否允许评论：</strong></td>
           <td height='30' colspan="2">&nbsp;
		    <input type="radio" name="comment" value="0"<%if comment=0 then response.write " checked"%>>不允许（关闭）<Br/>
			&nbsp;&nbsp;<input type="radio" name="comment" value="1"<%if comment=1 then response.write " checked"%>>允许，评论内容需要审核<Br/>	
			&nbsp;&nbsp;<input type="radio" name="comment" value="2"<%if comment=2 then response.write " checked"%>>允许，评论不需要审核		   </td>
          </tr>  
		   
		  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>是否锁定：</strong></td>
           <td height='30' colspan="2">&nbsp;
		    <input type="radio" name="locked" value="0"<%if locked=0 then response.write " checked"%>>否
			<input type="radio" name="locked" value="1"<%if locked=1 then response.write " checked"%>>是		   </td>
          </tr>  
		  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>是否结束：</strong></td>
           <td height='30' colspan="2">&nbsp;
		    <input type="radio" name="endtf" value="0"<%if endtf=0 then response.write " checked"%>>否
			<input type="radio" name="endtf" value="1"<%if endtf=1 then response.write " checked"%>>是		   </td>
          </tr>  
		 </table>    
 
<%
End Sub

Sub DoSave()
       Dim ID:ID=KS.ChkClng(KS.G("id"))
	   Dim Subject:Subject=KS.LoseHtml(KS.G("Subject"))
       Dim ActiveDate:ActiveDate=KS.G("ActiveDate")
			if not isdate(ActiveDate) then
			 Response.Write "<script>alert('本单载止日期格式不正确！');history.back();</script>"
			 Exit Sub
			End If	  
       Dim AddDate:AddDate=KS.G("AddDate")
			if not isdate(AddDate) then
			 Response.Write "<script>alert('发布时间格式不正确！');history.back();</script>"
			 Exit Sub
		End If	 
			


	   Dim PhotoUrl:PhotoUrl=KS.G("PhotoUrl")
	   Dim BigPhoto:BigPhoto=KS.G("BigPhoto")


			 
	   Dim Intro:Intro=Request.Form("Intro")
	   Dim Fax:Fax=KS.LoseHtml(KS.G("Fax"))
	   Dim Highlights:Highlights=KS.LoseHtml(KS.G("Highlights"))
	   Dim Protection:Protection=KS.LoseHtml(KS.G("Protection"))
	   Dim Notes:Notes=KS.LoseHtml(KS.G("Notes"))
	   Dim Locked:Locked=KS.ChkClng(KS.G("Locked"))
	   Dim EndTF:EndTF=KS.ChkClng(KS.G("EndTf"))
	   Dim ComeUrl:ComeUrl=KS.G("ComeUrl")
	   Dim ClassID:ClassID=KS.ChkClng(KS.G("ClassID"))
	   Dim AllowBMFlag:AllowBMFlag=KS.ChkClng(KS.G("AllowBMFlag"))
	   Dim minnum:minnum=KS.ChkClng(KS.G("minnum"))
	   Dim AllowArrGroupID:AllowArrGroupID=KS.G("AllowArrGroupID")
	   Dim Price_Original:Price_Original=KS.G("Price_Original")
	   Dim Discount:Discount=KS.G("Discount")
	   Dim Price:Price=KS.G("Price")
	   Dim Weight:Weight=KS.G("Weight")
	   If Not IsNumeric(Weight) Then Weight=0
	   Dim recommend:recommend=KS.ChkClng(KS.G("recommend"))
	   Dim LimitBuyNum:LimitBuyNum=KS.ChkCLng(KS.G("LimitBuyNum"))
	   Dim ProvinceID:ProvinceID=KS.ChkClng(KS.G("province"))
	   Dim CityID:CityID=KS.ChkClng(KS.G("city"))
	   Dim HasBuyNum:HasBuyNum=KS.ChkClng(KS.G("hasbuynum"))
	   Dim MustPayOnline:MustPayOnline=KS.ChkClng(KS.G("MustPayOnline"))
	   Dim CleanCart:CleanCart=KS.ChkClng(KS.G("CleanCart"))
	   Dim Comment:Comment=KS.ChkClng(KS.G("Comment"))
	   Dim showdelivery:showdelivery=KS.ChkClng(KS.G("showdelivery"))
	   
		
	   If KS.IsNul(Subject) Then KS.Die "<script>alert('团购主题必须输入!');history.back();</script>"
	   If not isnumeric(Price_Original) Then KS.Die "<script>alert('原价必须输入正确的数字!');history.back();</script>"
	   If not isnumeric(Discount) Then KS.Die "<script>alert('折扣必须输入正确的数字!');history.back();</script>"
	   If not isnumeric(Price) Then KS.Die "<script>alert('团购价必须输入正确的数字!');history.back();</script>"

            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select top 1 * From KS_GroupBuy Where ID=" & ID,Conn,1,3
			  If RS.Eof And RS.Bof Then
			     RS.AddNEW
				 RS("IsSuccess")=0
			  End If
				 RS("AddDate")=AddDate
			     RS("Subject")=Subject
				 RS("ActiveDate")=ActiveDate
				 RS("Intro")=Intro
				 RS("PhotoUrl")=PhotoUrl
				 RS("BigPhoto")=BigPhoto
				 RS("Highlights")=Highlights
				 RS("Protection")=Protection
				 RS("ClassID")=ClassID
				 RS("Notes")=Notes
				 RS("Locked")=Locked
				 RS("EndTF")=EndTF
				 RS("minnum")=minnum
				 RS("LimitBuyNum")=LimitBuyNum
				 RS("Weight")=Weight
				 RS("AllowBMFlag")=AllowBMFlag
				 RS("AllowArrGroupID")=AllowArrGroupID
				 RS("Price_Original")=Price_Original
				 RS("Discount")=Discount
				 RS("Price")=Price
				 RS("recommend")=recommend
				 RS("HasBuyNum")=HasBuyNum
				 RS("MustPayOnline")=MustPayOnline
				 RS("CleanCart")=CleanCart
				 RS("Comment")=Comment
				 RS("showdelivery")=showdelivery
				 RS("ProvinceID")=ProvinceID
				 RS("CityID")=CityID
		 		 RS.Update
				 If ID=0 Then
				   RS.MoveLast
                   Call KS.FileAssociation(1005,RS("ID"),Intro&RS("PhotoUrl"),0)
				 Else
                   Call KS.FileAssociation(1005,ID,Intro&RS("PhotoUrl"),1)
				 End If
				 
			     RS.Close
				 Set RS=Nothing
				 If ID=0 Then
				  Response.Write "<script>if (confirm('团购信息发布成功!')){location.href='?action=Add';}else{parent.frames['BottomFrame'].location.href='KS.Split.asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("团购系统 >> <font color=red>管理首页</font>") & "';location.href='KS.GroupBuy.asp';}</script>"
				 Else
				  Response.Write "<script>alert('团购信息修改成功！');parent.frames['BottomFrame'].location.href='KS.Split.asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("团购系统 >> <font color=red>管理首页</font>") & "';location.href='"& ComeUrl & "';</script>"
				 End If

EnD Sub

	'删除
	Sub GroupBuyDel()
	 Dim ID:ID=KS.G("ID")
	 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
	 Conn.execute("Delete From KS_UploadFiles Where ChannelID=1005 and InfoID In("& id & ")")
	 Conn.execute("Delete From KS_GroupBuy Where id In("& id & ")")
	 Response.Write "<script>alert('删除成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	
	Sub Recommend()
	 Dim ID:ID=KS.G("ID")
	 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
	 Conn.execute("Update KS_GroupBuy  set recommend=1 Where id In("& id & ")")
	 Response.Write "<script>alert('恭喜，批量设置推荐成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	Sub UnRecommend()
	 Dim ID:ID=KS.G("ID")
	 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
	 Conn.execute("Update KS_GroupBuy  set recommend=0 Where id In("& id & ")")
	 Response.Write "<script>alert('恭喜，批量取消推荐成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	
	Sub GroupBuyendtf()
	 Dim ID:ID=KS.G("ID")
	 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
	 Conn.execute("Update KS_GroupBuy Set endtf=1 Where id In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	Sub GroupBuyCancelendtf()
	 Dim ID:ID=KS.G("ID")
	 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
	 Conn.execute("Update KS_GroupBuy Set endtf=0 Where id In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	'锁定
	Sub GroupBuyLock()
	 Dim ID:ID=KS.G("ID")
	 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
	 Conn.execute("Update KS_GroupBuy Set locked=1 Where id In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	'解锁
	Sub GroupBuyUnLock()
	 Dim ID:ID=KS.G("ID")
	 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
	 Conn.execute("Update KS_GroupBuy Set locked=0 Where id In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
    
	'添加团购分类
    Sub AddClass()
	 Dim CategoryName,OrderID,ID,RS
	 ID=KS.ChkClng(KS.G("ID"))
	 If ID=0 Then
	   OrderID=Conn.Execute("select Max(OrderID) From KS_GroupBuyClass")(0)
	   OrderID=KS.ChkClng(OrderID)+1
	 Else
	   Set RS=Conn.Execute("Select top 1 * From KS_GroupBuyClass Where ID=" & ID)
	   If Not RS.Eof Then
	    CategoryName=RS("CategoryName")
		OrderID=RS("OrderID")
	   End If
	   RS.Close
	   Set RS=Nothing
	 End If
	 
	%>
	<div style="text-align:center;margin:15px;font-weight:bold">
	<%If ID=0 Then%>
	添加团购分类
	<%else%>
	修改团购分类
	<%end if%>
	</div>
	<table width="80%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
 <form name="myform" action="?action=AddClassSave" method="post">
   <input type="hidden" name="id" value="<%=id%>"/>
   <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl">
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>分类名称：</strong></td>
            <td width="435" height='30'>&nbsp;<input type='text'  class="textbox" name='CategoryName' id='CategoryName' value='<%=CategoryName%>' size="40"> <font color=red>*</font></td>
          </tr> 
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='200' height='30' align='right' class='clefttitle'><strong>排列序号：</strong></td>
            <td width="435" height='30'>&nbsp;<input type='text'  class="textbox" name='OrderID' id='OrderID' value='<%=OrderID%>' size="5" style="text-align:center"> <font color=red>*</font></td>
          </tr> 
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td colspan=2 style="text-align:center"><input type="submit" value="确定保存" class="button"/></td>
          </tr> 
	</form>
	</table> 
	<%
	End Sub
	
	Sub AddClassSave()
	 Dim CategoryName,OrderID,ID
	 CategoryName=KS.G("CategoryName")
	 OrderID=KS.ChkClng(KS.G("OrderID"))
	 ID=KS.ChkClng(KS.G("ID"))
	 If KS.IsNul(CategoryName) Then KS.Die "<script>alert('请输入团购分类名称!');history.back();</script>"
	 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	 If ID=0 Then
	   RS.open "select top 1 * from KS_GroupBuyClass Where CategoryName='"& CategoryName & "'",CONN,1,1
	 Else
	   RS.open "select top 1 * from KS_GroupBuyClass Where ID<>" & ID & " and CategoryName='"& CategoryName & "'",CONN,1,1
	 End If
	 If Not RS.Eof Then
	 KS.Die "<script>alert('对不起，您输入的团购分类名称已存在!');history.back();</script>"
	 End If
	 RS.Close
	 
	 RS.Open "select top 1 * From KS_GroupBuyClass Where ID=" & ID,conn,1,3
	 If RS.Eof And RS.Bof Then
	 RS.AddNew
	 End If
	  RS("CategoryName")=CategoryName
	  RS("OrderID")=OrderID
	 RS.Update
	 RS.Close
	 Set RS=Nothing
	 
	 If ID=0 Then
	 KS.Die "<script>if (confirm('恭喜，团购分类添加成功,继续添加吗？')){location.href='?action=AddClass'}else{location.href='?action=ClassManage';}</script>"
	 Else
	 KS.Die "<script>alert('恭喜，团购分类修改成功!');location.href='?action=ClassManage';</script>"
	 End If
	End Sub

Private Sub ClassManage()
%>

<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td nowrap>分类名称</th>
	<td nowrap>序号</th>
	<td nowrap>管理操作</th>
</tr>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_GroupBuyClass order by orderid,id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>还没有添加任何团购分类！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action=?action=DelClass>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="30"  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td align="center" class="splittd"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td align="center" width="120" class="splittd">
	<%=Rs("CategoryName")%>
	</td>
	<td align="center" class="splittd">
	<%=Rs("OrderID")%>
	</td>

	<td align="center" class="splittd">
		
		<a href="?Action=AddClass&id=<%=RS("ID")%>">修改</a>
		<a href="?Action=DelClass&id=<%=RS("ID")%>" onClick="return(confirm('确定删除该分类吗?'))">删除</a>
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
	<input class="button" type="submit" name="Submit2" value=" 删除选中的团购分类 " onClick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){this.document.selform.submit();return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  colspan=7 align=right>
	<%
	Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>

<%
End Sub

 Sub DelClass()
   Dim ID:ID=KS.FilterIds(KS.G("ID"))
   If ID="" Then KS.Die "<script>alert('没有选择分类ID!');history.back();</script>"
   Conn.Execute("Delete From KS_GroupBuyClass Where  ID In (" & ID & ")")
   KS.Die "<script>alert('恭喜，删除成功!');location.href='KS.GroupBuy.asp?action=ClassManage';</script>"
 End Sub

End Class
%> 
