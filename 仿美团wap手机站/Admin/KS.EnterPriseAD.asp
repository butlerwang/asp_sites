<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.UpFileCls.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_EnterpriseAD
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_EnterpriseAD
        Private KS,typeflag
		Private Action,i,strClass,sFileName,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
		Private ComeUrl,Selbutton,LoginTF,Verific,PhotoUrl,bigclassid,smallclassid,flag
		Private ClassID,Title,ADWZ,URL,datatimed,Adtype,status,begindate,username

        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		 With Response
					If Not KS.ReturnPowerResult(0, "KSMS10013") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
		typeflag=ks.chkclng(ks.g("type"))
			  .Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  If KS.G("Action")<>"View" then
			  .Write "<div class='topdashed sort'>行业关键词广告管理管理  <a href='?type=" & typeflag & "&flag=1'>未审核广告</a>  <a href='?type=" & typeflag & "&flag=2'>已过期广告</a></div>"
			 End If
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
		totalPut = Conn.Execute("Select Count(id) From KS_EnterpriseAD")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Select Case KS.G("action")
		 Case "add","Edit" Call AddAd()
		 Case "DoSave" Call DoSave()
		 Case "Del" Call DelRecord()
		 Case "verific"  Call Verify()
		 Case "unverific"  Call UnVerify()
		 Case "View" Call ShowNews()
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
%>
<script src="../ks_inc/kesion.box.js"></script>
<script>
function ShowIframe(id)
{
  new KesionPopup().PopupCenterIframe('<b>查看行业关键词广告管理</b>',"KS.EnterpriseAD.asp?action=View&ProID="+id,550,300,'auto')
}
</script>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>选择</td>
	<td nowrap>广告名称</td>
	<td nowrap>申请者</td>
	<td nowrap>播放位置</td>
	<td nowrap>生效日期</td>
	<td nowrap>播放天数</td>
	<td nowrap>状态</td>
	<td nowrap>管理操作</td>
</tr>
<%
	sFileName = "KS.EnterpriseAD.asp?"
	Dim Param
	If KS.ChkCLng(KS.G("Flag"))=1 Then 
	  Param=" where status=0"
	ElseIf KS.ChkClng(KS.G("Flag"))=2 Then
	  Param=" where datediff(" & DataPart_D & ",BeginDate," &SqlNowString & ")>datatimed"
	End If
	
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_EnterpriseAD  " & Param & " order by id desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>找不到行业关键词广告！</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action="?">
<input type="hidden" name="type" value="<%=typeflag%>">
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td class="splittd"><a href="#" onclick="ShowIframe(<%=rs("id")%>)"><%=Rs("Title")%></a>
	<%
	 if  datediff("d",RS("Begindate"),now)> Rs("datatimed") then
	  response.write "<font color=red>已过期</font>"
	 end if
	%>
	
	</td>
	<td class="splittd" align="center"><a href='../space/?<%=rs("username")%>' target='_blank'><%=Rs("username")%></a></td>
	<td class="splittd" align="center"> 企业库</td>
	<td class="splittd" align="center"><%=Rs("begindate")%></td>
	<td class="splittd" align="center"><%=Rs("datatimed")%> 天</td>
	<td class="splittd" align="center"><%
	select case rs("status")
	 case 0
	  response.write "<font color=red>未审</font>"
	 case 1
	  response.write "<font color=#999999>已审</font>"
	 case 2
	  response.write "<font color=blue>锁定</font>"
	end select
	%></td>
	<td class="splittd" align="center"><a href="#" onclick="ShowIframe(<%=rs("id")%>)">浏览</a> 
	<a href="?type=<%=typeflag%>&Action=Edit&ID=<%=rs("id")%>">修改</a>
	<a href="?type=<%=typeflag%>&Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('确定删除吗？'));">删除</a> <a href="?type=<%=typeflag%>&Action=verific&id=<%=rs("id")%>">审核</a></td>
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
	<td  class="splittd" height='25' colspan=8>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class="button" type="submit" name="Submit2" value=" 删除选中的广告" onclick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){this.form.Action.value='Del';this.form.submit();return true;}return false;}">
	<input type="button" value="批量审核" class="button" onclick="this.form.Action.value='verific';this.form.submit();">
	<input type="button" value="批量取消审核" class="button" onclick="this.form.Action.value='unverific';this.form.submit();">
	<input type="hidden" value="Del" name="Action">
	<input type="button" class="button" value="添加广告" onclick="location.href='?type=<%=typeflag%>&action=add'">
	</td>
</tr>
</form>
<tr>
	<td colspan=10>
	<%
	Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>

<%
End Sub

Sub AddAd()
     if KS.S("Action")="Edit" Then
		  Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		   RSObj.Open "Select * From KS_EnterPriseAD Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
		   If Not RSObj.Eof Then
			 Title    = RSObj("Title")
			 ADType = RSObj("ADType")
			 BigClassID=RSObj("BigClassID")
			 SmallClassID=RSObj("SmallClassID")
			 URL   = RSObj("URL")
			 ADWZ  = RSObj("ADWZ")
			 datatimed=RSObj("datatimed")
			 PhotoUrl  = RSObj("PhotoUrl")
			 status=trim(rsobj("status"))
			 BeginDate=rsobj("Begindate")
			 If PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="/Images/nopic.gif"
			 flag=true
			 UserName=rsobj("username")
		   End If
		   RSObj.Close:Set RSObj=Nothing
		Else
		 PhotoUrl="/images/nopic.gif"
		 ADWZ="1"
		 URL="http://"
		 flag=false
		 status=1
		 BeginDate=Now
		End If
		%>
		<script language="javascript" src="../ks_inc/popcalendar.js"></script>

		<script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.Title.value=="")
				  {
					alert("请输入广告名称！");
					document.myform.Title.focus();
					return false;
				  }	
				
				if (document.myform.URL.value=="")
				  {
					alert("请输入广告地址！");
					document.myform.URL.focus();
					return false;
				  }	
				
				 return true;  
				}
				</script>
				
				
				<table  width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form  action="?Action=DoSave" method="post" name="myform" id="myform" onSubmit="return CheckForm();" enctype="multipart/form-data">
				  <input type="hidden" value="<%=typeflag%>" name="type">
				   <input type="hidden" value="<%=KS.S("ID")%>" name="id">
				    <tr>
					  <td colspan=3 align=center class="Title">
					       <%IF KS.S("Action")="Edit" Then
							   response.write "修改关键词广告"
							   Else
							    response.write "关键词广告提交"
							   End iF
							  %>                         </td>
					</tr>
                    
                      <tr class="tdbg">
                        <td  height="25" align="center">投放类型：</td>
                        <td>　
                          <input name="Adtype" type="radio" value="1" onClick="document.all.SmallClassID.disabled=true;">                                 
                          大类
                          <input <%if trim(adtype)="2" then response.write " checked"%> name="AdType" type="radio" onClick="document.all.SmallClassID.disabled=false;" value="2">        
                          小类</td><td width="36%" rowspan="10" align="center">
                          <img src="<%=photourl%>" width="250" height="120">							  </td>
                      </tr>
                      <tr class="tdbg">
                        <td  height="25" align="center">行业类别：</td>
                        <td>　
                          <%
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
		   <option value="">--请选择行业大类--</option>
		<% 
		dim rsb,sqlb
		set rsb=server.createobject("adodb.recordset")
		if typeflag=1 then
        sqlb = "select * from ks_enterpriseClass_zs where parentid=0 order by orderid"
		else
        sqlb = "select * from ks_enterpriseClass where parentid=0 order by orderid"
		end if
        rsb.open sqlb,conn,1,1
		if rsb.eof and rsb.bof then
		else
		    Dim N
		    do while not rsb.eof
			          N=N+1
					  If N=1 and flag=false Then BigClassID=rsb("id")
					  If BigClassID=rsb("id") then
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
                  <select class="face" name="SmallClassID"<%if adtype="1" then response.write " disabled"%>>
				  <option value="" selected>--请选择行业子类--</option>
                    <%dim rsss,sqlss
						set rsss=server.createobject("adodb.recordset")
						if typeflag=1 then
						sqlss="select * from ks_enterpriseclass_zs where parentid="& KS.ChkClng(BigClassID)&" order by orderid"
						else
						sqlss="select * from ks_enterpriseclass where parentid="&KS.ChkClng(BigClassID)&" order by orderid"
						end if
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
					 
                      <tr class="tdbg" style="display:none">
                                      <td  height="25" align="center"><span>投放位置：</span></td>
                                      <td height="25">　
									  <!--
                                        <input name="ADWZ" type="radio" value="1"<%if trim(ADWZ)="1" then response.write " checked"%>/>企业大全
                                        <input name="ADWZ" type="radio" value="2"<%if trim(ADWZ)="2" then response.write " checked"%>/>产品库      -->
										<input name="ADWZ" type="hidden" value="1" />
										                                 </td>
                              </tr>
                              <tr class="tdbg">
                                <td height="25" align="center">投放时间：</td>
                                <td height="25">　
                            <select name="datatimed" id="datatimed">
                                   <option value="" selected>请选择...</option>
                                   <option value="7"<%if datatimed="7" then response.write " selected"%>>一周</option>
                                   <option value="15"<%if datatimed="15" then response.write " selected"%>>半个月</option>
                                   <option value="30"<%if datatimed="30" then response.write " selected"%>>一个月</option>
                                   <option value="60"<%if datatimed="60" then response.write " selected"%>>二个月</option>
                                   <option value="90"<%if datatimed="90" then response.write " selected"%>>三个月</option>
                                   <option value="180"<%if datatimed="180" then response.write " selected"%>>半年</option>
                                   <option value="365"<%if datatimed="365" then response.write " selected"%>>一年</option>
                                   <option value="730"<%if datatimed="730" then response.write " selected"%>>二年</option>
                               </select></td>
                              </tr>
                              <tr class="tdbg">
                                <td height="25" align="center">生效日期：</td>
                                <td height="25">　
                                <input name="BeginDate" type="text" class="textbox" id="BeginDate" style="width:150px; " value="<%=BeginDate%>" maxlength="40" />
                                <span class="tips">格式：0000-00-00</span></td>
                              </tr>
							  <tr class="tdbg">
								   <td width="12%"  height="25" align="center"><span>广告名称：</span></td>
									  <td width="52%"> 　
												<input class="textbox" name="Title" type="text" style="width:250px; " value="<%=Title%>" maxlength="100" />
												  <span style="color: #FF0000">*</span></td>
							  </tr>
                              <tr class="tdbg">
                                      <td height="25" align="center"><span>链接地址：</span></td>
                                      <td height="25">　
                                        <input name="URL" class="textbox" type="text" id="URL" style="width:250px; " value="<%=URL%>" maxlength="30" />
                                        <span style="color: #FF0000">*</span></td>
                              </tr>
                      <tr class="tdbg">
                           <td  height="25" align="center"><span>图片地址：</span></td>
                        <td> 　
                               <input type="file" class="textbox" name="photourl" size="40">
                          <span style="color: #FF0000">*</span> <br>
                          　 <font color=red>说明：只支持JPG、GIF、PNG格式图片，不超过300K,大小650*90</font></td>
                      </tr>
                      <tr class="tdbg">
                        <td  height="25" align="center">用户名：</td>
                        <td>　
                           <input name="UserName" class="textbox" type="text" style="width:100px; " value="<%=username%>" maxlength="30" /></td>
                      </tr>
                      <tr class="tdbg">
                        <td  height="25" align="center">状  态：</td>
                        <td>　
						
						  <input type="radio" name="status" value="1"<%if trim(status)="1" then response.write " checked"%>> 已审
						  <input type="radio" name="status" value="0"<%if trim(status)="0" then response.write " checked"%>> 未审核						</td>
                      </tr>
                        
                             
			  
                    <tr class="tdbg">
                      <td height="30" style="text-align:center" colspan=3>
					 <input class="button" type="submit" name="Submit" value="OK, 保 存 " />
                            　
                            <input class="button" type="reset" name="Submit2" value=" 重 来 " />						</td>
                    </tr>
                  </form>
			    </table>
		        <br>
		        <table  width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <TR class="title">
                    <TD  height="24"><STRONG>注意事项：</STRONG></TD>
                  </TR>
                  <TR>
                    <TD bgColor="#ffffff" height="26"><TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
                        <TBODY>
                          
                          <TR>
                            <TD height="21"><IMG height="8" src="images/expand.gif" width="8">请确保您的广告健康，不含黄色信息。确定真实性，合法性，否则后果自负，<%=KS.Setting(1)%>不承担任何责任。</TD>
                          </TR>
                          <TR>
                            <TD height="21"><IMG height="8" src="images/expand.gif" width="8">提交的行业广告必须经过管理员审核后才能生效。生效时间以审核时间为准。</TD>
                          </TR>
                        </TBODY>
                    </TABLE></TD>
                  </TR>
            </table>
<%
End Sub


Sub DoSave()
  
            Dim fobj:Set FObj = New UpFileClass
			FObj.GetData
            Dim MaxFileSize:MaxFileSize = 300   '设定文件上传最大字节数
			Dim AllowFileExtStr:AllowFileExtStr = "gif|jpg|png"
			
			UserName=Fobj.Form("UserName")
			Dim UserID,RS:Set RS=Conn.Execute("Select top 1 userid From KS_User Where UserName='" & UserName & "'")
			If RS.eof Then
			    RS.Close:Set RS=Nothing
				Response.Write "<script>alert('你输入的用户名不存在!');history.back();</script>"
				Exit Sub
			Else
			 UserID=RS(0)
			End If	 

			Dim FormPath:FormPath =KS.ReturnChannelUserUpFilesDir(9994,UserID)
			Call KS.CreateListFolder(FormPath) 
			
           
				 Title=KS.LoseHtml(Fobj.Form("Title"))
				  If Title="" Then
				    Response.Write "<script>alert('你没有输入广告名称!');history.back();</script>"
				    Exit Sub
				  End IF
				 
				 Adtype=KS.ChkClng(Fobj.Form("Adtype"))
				 If AdType=0 Then AdType=1
				 BigClassID=KS.ChkCLng(Fobj.Form("ClassID"))
				 SmallClassID=KS.ChkCLng(Fobj.Form("SmallClassID"))
				 
				 URL=KS.DelSql(Fobj.Form("URL"))
				 ADWZ=KS.ChkClng(Fobj.Form("ADWZ"))
				 datatimed=KS.ChkClng(Fobj.Form("datatimed"))
				 status=KS.ChkClng(Fobj.Form("status"))
				 BeginDate=Fobj.Form("Begindate")
				 
				 If Not IsDate(BeginDate) Then
				 Call KS.AlertHistory("开始日期格式不正确!",-1)
				 Response.End()
				 End If
			
			Dim ReturnValue:ReturnValue = FObj.UpSave(FormPath,MaxFileSize,AllowFileExtStr,year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now))
			Select Case ReturnValue
			  Case "errext" Call KS.AlertHistory("文件上传失败,文件类型不允许\n允许的类型有" + AllowFileExtStr + "\n",-1):response.end
	          Case "errsize"  Call KS.AlertHistory("文件上传失败,文件超过允许上传的大小\n允许上传 " & MaxFileSize & " KB的文件\n",-1):response.End()
			End Select
			
			If ReturnValue="" and KS.ChkClng(Fobj.Form("ID"))=0 then
			 Call KS.AlertHistory("广告图片必须上传!",-1)
			 Response.End()
			End If

				  
				Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select * From KS_EnterPriseAD Where ID=" & KS.ChkClng(Fobj.Form("ID")),Conn,1,3
				If RSObj.Eof Then
				  RSObj.AddNew
				 End If
				  RSObj("UserName")=UserName
				  RSObj("Title")=Title
				  RSObj("ADType")=ADType
				  RSObj("URL")=URL
				  RSObj("ADWZ")=ADWZ
				  RSObj("BigClassID")=BigClassID
				  RSObj("SmallClassID")=SmallClassID
				  RSObj("datatimed")=datatimed
				  If ReturnValue<>"" then
				  RSObj("PhotoUrl")=ReturnValue
				  end if
  				  RSObj("Status")=status
				  RSObj("BeginDate")=BeginDate
				 RSObj.Update
				 If KS.ChkClng(Fobj.Form("ID"))=0 Then
				  Call KS.FileAssociation(1014,rsobj("id"),RSObj("PhotoUrl"),0)
				 Else
				  Call KS.FileAssociation(1014,rsobj("id"),RSObj("PhotoUrl"),1)
				 End If
				 
				 RSObj.Close:Set RSObj=Nothing
				 
               If KS.ChkClng(Fobj.Form("ID"))=0 Then
			     Set Fobj=Nothing
				 Response.Write "<script>if (confirm('关键词广告提交成功，继续提交吗?')){location.href='?type=" & typeflag &"&Action=add';}else{location.href='KS.EnterPriseAD.asp';}</script>"
			   Else
			     Set Fobj=Nothing
				 Response.Write "<script>alert('关键词广告修改成功!');location.href='KS.EnterPriseAD.asp?type=" & typeflag &"';</script>"
			   End If
  End Sub

'删除日志
Sub DelRecord()
 Dim I,ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 ID=Split(ID,",")
 For I=0 To Ubound(ID)
  KS.DeleteFile(conn.execute("select photourl from ks_EnterpriseAD where id=" & ID(I))(0))
  Conn.Execute("Delete From KS_UploadFiles Where ChannelID=1014 and infoid=" & ID(I))
  Conn.execute("Delete From KS_EnterpriseAD Where id="& id(I))
 Next 
 Response.Write "<script>alert('删除成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

'审核
Sub ShowNews()
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_EnterpriseAD where id=" &KS.ChkClng(KS.S("ProID")),conn,1,1
		If Not RS.Eof Then
		   Response.Write "<div><strong>投放类型：</strong>" 
		    If RS("AdType")=1 Then
			 Response.Write "大类"
			Else
			 Response.Write "小类"
			End If
		   Response.Write "</div>"
		   Response.WRITE "<div><strong>广告名称：</strong>" & rs("Title") & "</div>"
		   Response.Write "<div style=""text-align:left""><strong>链接地址：</strong>" & RS("url") & "</div>"
		   Response.Write "<div style=""text-align:left""><strong>播放位置：</strong>" 
		   If RS("ADWZ")="1" Then
		    response.write "产品库"
		   Else
		    response.write "企业大全"
		   End If
		   Response.Write "</div>"
		   Response.Write "<div style=""text-align:left""><strong>开始日期：</strong>" & RS("begindate") & "</div>"
		   Response.Write "<div style=""text-align:left""><strong>开始天数：</strong>" & RS("datatimed") & " 天</div>"
		   Dim PhotoUrl:PhotoUrl=RS("PhotoUrl")
		   If PhotoUrl<>"" And Not IsNull(PhotoURL) Then
		   Response.Write "<div style=""text-align:left""><strong>广告图片：</strong><img src='" & RS("photourl") & "'></div>"
		   End If
		End If
		RS.Close:Set RS=Nothing
End Sub
'审核
Sub Verify
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_EnterpriseAD Set status=1,begindate=" & SqlNowString & " Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'取消审核
Sub UnVerify
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_EnterpriseAD Set status=0 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

End Class
%> 
