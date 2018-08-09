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
Set KSCls = New Admin_EnterPrise
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_EnterPrise
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
		 With KS
					If Not KS.ReturnPowerResult(0, "KSMS20010") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
              .echo "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">" & vbcrlf
			  .echo "<html xmlns=""http://www.w3.org/1999/xhtml"">"
			  .echo"<head>"
			  .echo"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .echo"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .echo "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .echo "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			  .echo "<script type=""text/javascript"" src=""../editor/ckeditor.js"" mce_src=""../editor/ckeditor.js""></script>"
			  .echo"</head>"
			  .echo"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .echo "<ul id='menu_top'>"
			  .echo "<li class='parent' onclick=""window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('积分兑换系统 >> <font color=red>添加商品</font>')+'&ButtonSymbol=GOSave';location.href='?action=Add';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加商品</span></li>"
			  .echo "<li class='parent' onclick=""location.href='KS.MallScoreOrder.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>处理兑换订单</span></li>"
			  .echo "<li class='view'><strong>查看方式：</strong><a href=""KS.MallScore.asp"">所有商品</a> <a href=""KS.MallScore.asp?flag=1"">已审</a>  <a href=""KS.MallScore.asp?flag=2"">未审</a> <a href=""KS.MallScore.asp?flag=3"">已结束</a></li>"

			  .echo "</ul>"
		
		
		maxperpage = 30 '###每页显示数
		If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
			.echo ("错误的系统参数!请输入整数")
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
		   Param= Param & " and ProductName like '%" & KS.G("KeyWord") & "%'"
		  Else
		   Param= Param & " and Intro like '%" & KS.G("KeyWord") & "%'"
		  End If
		End If
		If KS.G("Flag")<>"" Then
		  If KS.G("Flag")="1" Then Param=Param & " and Status=1"
		  If KS.G("Flag")="2" Then Param=Param & " and Status=0"
		  If KS.G("Flag")="3" Then Param=Param & " and datediff(day,enddate," & SqlNowString & ")>0"
		  
		End If

		totalPut = Conn.Execute("Select Count(id) From KS_MallScore " & Param)(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Select Case KS.G("action")
		 Case "Add","Edit" Call ProductNameManage()
		 Case "EditSave" Call DoSave()
		 Case "Del"  Call BlogDel()
		 Case "lock"  Call BlogLock()
		 Case "unlock"  Call BlogUnLock()
		 Case "recommend"  Call Blogrecommend()
		 Case "Cancelrecommend" Call BlogCancelrecommend()
		 Case Else
		  Call showmain
		End Select
	End With	
End Sub

Private Sub showmain()
%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</th>
	<td nowrap>商品名称</th>
	<td nowrap>所需积分</th>
	<td nowrap>库存数量</th>
	<td nowrap>到期时间</th>
	<td nowrap>浏览次数</th>
	<td nowrap>兑换次数</th>
	<td nowrap>状态</th>
	<td nowrap>管理操作</th>
</tr>
<%
	sFileName = "KS.MallScore.asp?"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_MallScore " & Param & " order by id desc"
		Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>还没有添加任何商品！</td></tr>"
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
	<td  class="splittd"><%=Rs("ProductName")%>
	</td>
	<td align="center" class="splittd"><font color=red><%=Rs("Score")%></font> 分</td>
	<td align="center" class="splittd"><font color=red><%=Rs("Quantity")%></font></td>
	<td align="center" class="splittd"><font color=#cccccc><%=RS("EndDate")%></font></td>
	<td align="center" class="splittd"><%=RS("Hits")%> 次</td>
	<td align="center" class="splittd">
	 <span style="color:red;font-weight:bold"><%=LFCls.GetSingleFieldValue("Select Count(*) From KS_MallScoreOrder Where ProductID=" & RS("ID"))%></span>
	(<a href="#" onclick="javascript:window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('积分兑换系统 >> <font color=red>查看订单</font>')+'&ButtonSymbol=Disabled';location.href='KS.MallScoreOrder.asp?productid=<%=RS("ID")%>'">查看</a>)
	
	</td>
	<td align="center" class="splittd"><%
		if rs("status")=1 then
		  response.write "<font color=#cccccc>已审</font>"
		else
		  response.write " <font color=red>未审</font>"
		end if
	%></td>
	<td align="center" class="splittd"><a href="?action=Edit&ID=<%=RS("ID")%>"  onclick="window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('积分兑换系统 >> <font color=red>修改团购信息</font>')+'&ButtonSymbol=GOSave';">修改</a> <a href="?Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('确定删除该商品吗？'));">删除</a> 
		
		&nbsp;<%IF rs("Status")="1" then %><a href="?Action=Cancelrecommend&id=<%=rs("id")%>"><font color=red>取审</font></a><%else%><a href="?Action=recommend&id=<%=rs("id")%>">审核</a><%end if%>

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
	<input class="button" type="submit" name="Submit2" value=" 删除选中的商品 " onclick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){this.document.selform.submit();return true;}return false;}"></td>
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
<form action="KS.MallScore.asp" name="myform" method="get">
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

Sub ProductNameManage()
Dim ProductName,ActiveDate,AddDate,EndDate,Quantity,Score,Telphone,Intro,Hits,Protection,BuyFlow,Notes,recommend,Status,PhotoUrl
Dim ID:ID=KS.ChkClng(KS.G("ID"))
Dim RS:Set RS=server.createobject("adodb.recordset")
If KS.G("Action")="Edit" Then
	RS.Open "Select * From KS_MallScore Where ID=" & ID,conn,1,1
	 If RS.Eof And RS.Bof Then
	  RS.Close:Set RS=Nothing
	  KS.AlertHintScript "参数传递出错！"
	  Response.End
	 Else
	   ProductName=RS("ProductName")
	   AddDate=RS("AddDate")
	   EndDate=RS("EndDate")
	   Quantity=RS("Quantity")
	   Score=RS("Score")
	   Intro=RS("Intro")
	   PhotoUrl=RS("PhotoUrl")
	   If PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="../images/nopic.gif"
	   Hits=RS("Hits")
	   recommend=RS("recommend")
	   Status=RS("Status")
	 End If
Else
  AddDate=Now
  EndDate=Now+30
  Hits=0:Score=10
  recommend=0:Status=1
  Quantity=100
  Intro=" "
  PhotoUrl="../images/nopic.gif"
 End If
%>
<script>
function CheckForm()
{
	if ($('input[name=ProductName]').val()=='')
	{
	 alert('请输入商品名称!');
	 $("input[name=ProductName]").focus();
	 return false;
	}
	if ($('input[name=Quantity]').val()=='')
	{
	 alert('请输入库存数量!');
	 $("input[name=Quantity]").focus();
	 return false;
	}
	if ($('input[name=Score]').val()=='')
	{
	 alert('请输入所需积分!');
	 $("input[name=Score]").focus();
	 return false;
	}
	if (CKEDITOR.instances.Intro.getData()=="")
	{
	 alert('请输入商品介绍!');
	 return false;
	}
document.myform.submit();
}
</script>
<br>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
  <form name="myform" action="?action=EditSave" method="post" enctype="multipart/form-data">
    <input type="hidden" value="<%=ID%>" name="id" />
    <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl" />
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>商品名称：</strong></td>
      <td width="435" height='30'>&nbsp;
          <input type='text' name='ProductName' value='<%=ProductName%>' size="40" />
          <font color=red>*</font></td>
      <td width="217" rowspan="4" align="center"><div id="pic" style="filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(sizingMethod=scale);height:100px;width:95px;border:1px solid #777777"> <img src="<%=PhotoUrl%>" style="height:100px;width:95px;" /> </div></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>发布时间：</strong></td>
      <td height='30'>&nbsp;
          <select name="AddDate1">
            <%on error resume next
					  for i=year(now) to year(now)+1
					   if trim(split(AddDate,"-")(0))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "年</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "年</option>"
					   end if
					  next
					  %>
          </select>
          <select name="AddDate2">
            <%
					  for i=1 to 12
					   if trim(split(AddDate,"-")(1))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "月</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "月</option>"
					   end if
					  next
					  %>
          </select>
          <select name="AddDate3">
            <%
					  for i=1 to 31
					   if trim(split(split(AddDate,"-")(2)," ")(0))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "日</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "日</option>"
					   end if
					  next
					  %>
          </select>
          <select name="AddDate4">
            <%
					  for i=0 to 23
					   if trim(split(split(AddDate," ")(1),":")(0))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "时</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "时</option>"
					   end if
					  next
					  %>
          </select>
          <select name="AddDate5">
            <%
					  for i=0 to 59
					   if trim(split(split(AddDate," ")(1),":")(1))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "分</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "分</option>"
					   end if
					  next
					  %>
          </select>
      </td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>到期时间：</strong></td>
      <td height='30'>&nbsp;
          <select name="EndDate1">
            <%
					  for i=year(now) to year(now)+1
					   if trim(split(EndDate,"-")(0))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "年</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "年</option>"
					   end if
					  next
					  %>
          </select>
          <select name="EndDate2">
            <%
					  for i=1 to 12
					   if trim(split(EndDate,"-")(1))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "月</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "月</option>"
					   end if
					  next
					  %>
          </select>
          <select name="EndDate3">
            <%
					  for i=1 to 31
					   if trim(split(split(EndDate,"-")(2)," ")(0))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "日</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "日</option>"
					   end if
					  next
					  %>
          </select>
          <select name="EndDate4">
            <%
					  for i=0 to 23
					   if trim(split(split(EndDate," ")(1),":")(0))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "时</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "时</option>"
					   end if
					  next
					  %>
          </select>
          <select name="EndDate5">
            <%
					  for i=0 to 59
					   if trim(split(split(EndDate," ")(1),":")(1))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "分</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "分</option>"
					   end if
					  next
					  %>
          </select>
        过了这个时间将不能兑换 </td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td height='30' align='right' class='clefttitle'><strong>商品图片：</strong></td>
      <td height='30'>&nbsp;
  <input class="textbox" type="file" name="photo" size="40" onchange='document.getElementById(&quot;pic&quot;).innerHTML=&quot;&quot;;document.getElementById(&quot;pic&quot;).filters.item(&quot;DXImageTransform.Microsoft.AlphaImageLoader&quot;).src=this.value;' />
        <font color="red">*</font> <br />
        &nbsp;&nbsp;<font color="blue">请上传少于200K的图片,支持jpg,gif,png格式</font></td></tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>库存数量：</strong></td>
      <td height='30' colspan="2">&nbsp;
          <input type='text' name='Quantity' value='<%=Quantity%>' size="10" />
          <font color=red>*</font></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>所需积分：</strong></td>
      <td height='30' colspan="2">&nbsp;
          <input type='text' name='Score' value='<%=Score%>' size="10" />
        分 <font color=red>*</font></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>商品简介：</strong></td>
      <td height='30' colspan="2">&nbsp;
          <%
		    Response.Write "<textarea id=""Intro"" name=""Intro"" style=""display:none"">" &  KS.HTMLCode(Intro) &"</textarea>"
			Response.Write "<script type=""text/javascript"">"
            Response.Write "CKEDITOR.replace('Intro', {width:""700"",height:""220px"",toolbar:""NewsTool"",filebrowserBrowseUrl :""Include/SelectPic.asp?from=ckeditor&Currpath="& KS.GetUpFilesDir() &""",filebrowserWindowWidth:650,filebrowserWindowHeight:290});"
			Response.Write "</script>"

								%>
      </td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>浏览次数：</strong></td>
      <td height='30' colspan="2">&nbsp;
          <input name='Hits' value="<%=Hits%>" size="10" /></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>是否推荐：</strong></td>
      <td height='30' colspan="2">&nbsp;
          <input type="radio" name="recommend" value="0"<%if recommend=0 then response.write " checked"%> />
        否
        <input type="radio" name="recommend" value="1"<%if recommend=1 then response.write " checked"%> />
        是 </td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>是否审核：</strong></td>
      <td height='30' colspan="2">&nbsp;
          <input type="radio" name="Status" value="0"<%if Status=0 then response.write " checked"%> />
        否
        <input type="radio" name="Status" value="1"<%if Status=1 then response.write " checked"%> />
        是 </td>
    </tr>
  </form>
</table>
<%
End Sub

Sub DoSave()
		  	Dim Fobj:Set FObj = New UpFileClass
			on error resume next
			FObj.GetData
			if err.number<>0 then
			 call ks.alerthistory("出错了,文件超出大小",-1)
			 response.End()
			end if

       Dim ID:ID=KS.ChkClng(Fobj.Form("id"))
	   Dim ProductName:ProductName=KS.LoseHtml(Fobj.Form("ProductName"))
       Dim AddDate:AddDate=Fobj.Form("AddDate1") & "-" & Fobj.Form("AddDate2") & "-" & Fobj.Form("AddDate3") & " " & Fobj.Form("AddDate4") & ":" & Fobj.Form("AddDate5")
			if not isdate(AddDate) then
			 Response.Write "<script>alert('发布时间格式不正确！');history.back();</script>"
			 Exit Sub
			End If	 
       Dim EndDate:EndDate=Fobj.Form("EndDate1") & "-" & Fobj.Form("EndDate2") & "-" & Fobj.Form("EndDate3") & " " & Fobj.Form("EndDate4") & ":" & Fobj.Form("EndDate5")
			if not isdate(EndDate) then
			 Response.Write "<script>alert('发布时间格式不正确！');history.back();</script>"
			 Exit Sub
			End If	 
			
			Dim MaxFileSize:MaxFileSize = 200   '设定文件上传最大字节数
			Dim AllowFileExtStr:AllowFileExtStr = "gif|jpg|png"
			Dim FormPath:FormPath = KS.GetUpFilesDir() & "/MallScore/"
			Call KS.CreateListFolder(FormPath) 
			Dim ReturnValue:ReturnValue = FObj.UpSave(FormPath,MaxFileSize,AllowFileExtStr,"fm" & right(Year(Now),2) & right("0" & Month(Now),2) & right("0" & Day(Now),2) & right("0"&Hour(Now),2) & right("0"&Minute(Now),2) & right("0"&Second(Now),2))
			Select Case ReturnValue
			  Case "errext" Call KS.AlertHistory("文件上传失败,文件类型不允许\n允许的类型有" + AllowFileExtStr + "\n",-1):exit sub
	          Case "errsize"  Call KS.AlertHistory("文件上传失败,文件超过允许上传的大小\n允许上传 " & MaxFileSize & " KB的文件\n",-1):exit sub
			End Select

	   Dim PhotoUrl:PhotoUrl=ReturnValue


			 
	   Dim Quantity:Quantity=KS.ChkClng(Fobj.Form("Quantity"))
	   Dim Score:Score=KS.ChkClng(Fobj.Form("Score"))
	   Dim Intro:Intro=Fobj.Form("Intro")
	   Dim Hits:Hits=KS.LoseHtml(Fobj.Form("Hits"))
	   Dim recommend:recommend=KS.ChkClng(Fobj.Form("recommend"))
	   Dim Status:Status=KS.ChkClng(Fobj.Form("Status"))
	   Dim ComeUrl:ComeUrl=Fobj.Form("ComeUrl")
	   Set Fobj=Nothing
	   
		
	   If ProductName="" Then Response.Write "<script>alert('商品名称必须输入');history.back();</script>":response.end

            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select * From KS_MallScore Where ID=" & ID,Conn,1,3
			  If RS.Eof And RS.Bof Then
			     RS.AddNEW
				 RS("Inputer")=KS.C("AdminName")
			  End If
				 RS("AddDate")=AddDate
				 RS("EndDate")=EndDate
			     RS("ProductName")=ProductName
				 RS("Quantity")=Quantity
				 RS("Score")=Score
				 RS("Intro")=Intro
				 IF PhotoUrl<>"" Then RS("PhotoUrl")=PhotoUrl
				 RS("Hits")=Hits
				 RS("recommend")=recommend
				 RS("Status")=Status
		 		 RS.Update
				 If ID=0 Then
				   RS.MoveLast
                   Call KS.FileAssociation(1004,RS("ID"),Intro&RS("PhotoUrl"),0)
				 Else
                   Call KS.FileAssociation(1004,ID,Intro&RS("PhotoUrl"),1)
				 End If
				 
			     RS.Close
				 Set RS=Nothing
				 If ID=0 Then
				  Response.Write "<script>if (confirm('商品信息发布成功!')){location.href='?action=Add';}else{parent.frames['BottomFrame'].location.href='KS.Split.asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("积分兑换系统 >> <font color=red>管理首页</font>") & "';location.href='KS.MallScore.asp';}</script>"
				 Else
				  Response.Write "<script>alert('商品信息修改成功！');parent.frames['BottomFrame'].location.href='KS.Split.asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("积分兑换系统 >> <font color=red>管理首页</font>") & "';location.href='"& ComeUrl & "';</script>"
				 End If

EnD Sub

'删除日志
Sub BlogDel()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Delete From KS_MallScore Where id In("& id & ")")
 Conn.execute("Delete From KS_UploadFiles Where ChannelID=1004 and InfoID In("& id & ")")
 Response.Write "<script>alert('删除成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

Sub Blogrecommend()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_MallScore Set Status=1 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
Sub BlogCancelrecommend()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_MallScore Set Status=0 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

End Class
%> 
