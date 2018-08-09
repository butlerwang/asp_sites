<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Main
KSCls.Kesion()
Set KSCls = Nothing

Class Main
        Private KS,Action
		Private I, totalPut, CurrentPage, SqlStr, RSObj
        Private MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub


		Public Sub Kesion()
			If Not KS.ReturnPowerResult(0, "KSMS20014") Then
			  Call KS.ReturnErr(1, "")
			  exit sub
			End If
			Action=KS.G("Action")
			Select Case Action
			 Case "Add","Edit"
				  Call MailDepartAddOrEdit()
			 Case "Save"
			      Call DoSave()
			 Case "Del"
			      Call PKDelete()
			 Case Else
			   Call MainList()
			End Select
	    End Sub
		
		Sub MainList()
			If Request("page") <> "" Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
			With Response
			.Write "<html>"
			.Write "<head>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script language=""JavaScript"">" & vbCrLf
			.Write "var Page='" & CurrentPage & "';" & vbCrLf
			.Write "</script>" & vbCrLf
			.Write "<script language=""JavaScript"" src=""../ks_inc/jquery.js""></script>"
			.Write "<script language=""JavaScript"" src=""Include/ContextMenu1.js""></script>"
			.Write "<script language=""JavaScript"" src=""Include/SelectElement.js""></script>"
			%>
			<script language="JavaScript">
			var DocElementArrInitialFlag=false;
			var DocElementArr = new Array();
			var DocMenuArr=new Array();
			var SelectedFile='',SelectedFolder='';
			$(document).ready(function()
			{   if (DocElementArrInitialFlag) return;
				InitialDocElementArr('FolderID','PKID');
				InitialContextMenu();
				DocElementArrInitialFlag=true;
			});
			function InitialContextMenu()
			{	DocMenuArr[DocMenuArr.length]=new ContextMenuItem("window.parent.MailDepartAdd();",'添 加(N)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'全 选(A)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.MailDepartControl(1);",'编 辑(E)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.MailDepartControl(2);",'删 除(D)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.location.reload();",'刷 新(Z)','disabled');
			}
			function DocDisabledContextMenu()
			{
				DisabledContextMenu('FolderID','PKID','编 辑(E),删 除(D)','编 辑(E)','','','','')
			}
			function MailDepartAdd()
			{
				location.href='KS.PKZT.asp?Action=Add';
				window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr=活动管理中心 >> <font color=red>添加新活动</font>&ButtonSymbol=GO';
			}
			function EditMailDepart(id)
			{
				location="KS.PKZT.asp?Action=Edit&Page="+Page+"&Flag=Edit&PKID="+id;
				window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr=征集意见管理 >> <font color=red>编辑活动</font>&ButtonSymbol=GoSave';
			}
			function DelMailDepart(id)
			{
			if (confirm('真的要删除选中的活动吗?'))
			 location="KS.PKZT.asp?Action=Del&Page="+Page+"&PKID="+id;
			   SelectedFile='';
			}
			function MailDepartControl(op)
			{  var alertmsg='';
				GetSelectStatus('FolderID','PKID');
				if (SelectedFile!='')
				 {  if (op==1)
					{
					if (SelectedFile.indexOf(',')==-1) 
						EditMailDepart(SelectedFile)
					  else alert('一次只能编辑一条活动!')	
					SelectedFile='';
					}
				  else if (op==2)    
					 DelMailDepart(SelectedFile);
				 }
				else 
				 {
				 if (op==1)
				  alertmsg="编辑";
				 else if(op==2)
				  alertmsg="删除"; 
				 else
				  {
				  WindowReload();
				  alertmsg="操作" 
				  }
				 alert('请选择要'+alertmsg+'的活动');
				  }
			}
			function GetKeyDown()
			{ 
			if (event.ctrlKey)
			  switch  (event.keyCode)
			  {  case 90 : location.reload(); break;
				 case 65 : SelectAllElement();break;
				 case 78 : event.keyCode=0;event.returnValue=false; MailDepartAdd();break;
				 case 69 : event.keyCode=0;event.returnValue=false;MailDepartControl(1);break;
				 case 68 : MailDepartControl(2);break;
			   }	
			else	
			 if (event.keyCode==46)MailDepartControl(2);
			}
			</script>
			<%
			.Write "</head>"
			.Write "<body scroll=no topmargin=""0"" leftmargin=""0"" onclick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""MailDepartAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加PK主题</span></li>"
			  .Write "<li class='parent' onclick=""MailDepartControl(1);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>编辑PK主题</span></li>"
			  .Write "<li class='parent' onclick=""MailDepartControl(2);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>删除PK主题</span></li>"
			  .Write "</ul>"
			

			.Write "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
			.Write "  <tr>"			
			.Write "          <td width=""44%"" height=""25"" class=""sort"" align=""center"">PK主题名称</td>"
			.Write "          <td class=""sort"" align=""center"">栏目</td>"
			.Write "          <td class=""sort"" align=""center"">结束时间</td>"
			.Write "          <td align=""center"" class=""sort"">得票情况</td>"
			.Write "          <td align=""center"" class=""sort"">状态</td>"
			.Write "          <td align=""center"" class=""sort"">管理操作</td>"
			.Write "  </tr>"
			 
			 Set RSObj = Server.CreateObject("ADODB.RecordSet")
					   SqlStr = "SELECT * FROM KS_PKZT order by ID DESC"
					   RSObj.Open SqlStr, Conn, 1, 1
					 If RSObj.EOF And RSObj.BOF Then
					 Else
						totalPut = RSObj.RecordCount
			
								If CurrentPage < 1 Then CurrentPage = 1
								
			
								   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
									Else
										CurrentPage = 1
									End If
										Call showContent
				End If
				
			.Write "    </td>"
			.Write "  </tr>"
			.Write "</table>"
			.Write "</body>"
			.Write "</html>"
			End With
			End Sub
			Sub showContent()
			   on error resume next
			  With Response
					Do While Not RSObj.EOF
					  .Write "<tr>"
					  .Write "  <td class='splittd' width='44%' height='20'> &nbsp;&nbsp; <span PKID='" & RSObj("ID") & "' ondblclick=""EditMailDepart(this.PKID)""><img src='Images/Field.gif' align='absmiddle'>"
					  .Write "    <span style='cursor:default;'>" & KS.GotTopic(RSObj("Title"), 45) & "</span></span> "
					  .Write "  </td>"
					  .Write "  <td class='splittd' align='center'>" 
					  .Write  KS.C_C(RSObj("ClassID"),1)
					  .Write "  </td>"
					  
					  .Write "  <td class='splittd' align='center'>" 
					  if rsobj("timelimit")=1 then
					  .write rsobj("enddate")
					  else
					  .write "<font color=#cccccc>不限时间</font>"
					  end if
					  .Write " </td>"
					  .Write "  <td class='splittd' align='center'><a href='KS.PKGD.asp?pkid=" & rsobj("id") & "'>正:<font Color=red>" & rsobj("zfvotes") & "</font>票 反:<font Color=red>" & rsobj("ffvotes") & "</font>票 三:<font Color=red>" & rsobj("sfvotes") & "</font>票</a></td>"
					  .Write "  <td class='splittd' align='center'>"
					   if rsobj("status")=1 then
					    .write "<Font color=green>正常</font>"
					   else
					    .write "<Font color=red>锁定</font>"
					   end if
					  .Write "</td>"
					  .Write "  <td class='splittd' align='center'><a href='../plus/pk/pk.asp?id=" & rsobj("id") &"' target='_Blank'>查看</a> <a href='javascript:EditMailDepart(" & rsobj("id") &")'>编辑</a> <a href='javascript:DelMailDepart(" & rsobj("id") & ")'>删除</a></td>"
					  .Write "</tr>"
					 I = I + 1
					  If I >= MaxPerPage Then Exit Do
						   RSObj.MoveNext
					Loop
					  RSObj.Close
					  .Write "<tr><td height='26' colspan='5' align='right'>"
					  Call KS.ShowPageParamter(totalPut, MaxPerPage, "", True, "个", CurrentPage, "Action=" & Action)
				End With
			End Sub
			
			'添加修改活动
		  Sub MailDepartAddOrEdit()
		  		Dim PKID, RSObj,ClassID, TimeLimit,SqlStr, NewsLink,Title,enddate, ZFTips,FFTips, CategoryID, AddDate,Flag, Page,Status,ZFVotes,FFVotes,SFVotes,LoginTf,VerifyTF,OnceTF
				Flag = KS.G("Flag")
				Page = KS.G("Page")
				If Page = "" Then Page = 1
				If Flag = "Edit" Then
					PKID = KS.G("PKID")
					Set RSObj = Server.CreateObject("Adodb.Recordset")
					SqlStr = "SELECT * FROM KS_PKZT Where ID=" & PKID
					RSObj.Open SqlStr, Conn, 1, 1
					  Title     = RSObj("Title")
					  ZFTips    = RSObj("ZFTips")
					  FFTips    = RSObj("FFTips")
					  enddate  = RSObj("enddate")
					  NewsLink = RSObj("NewsLink")
					  Status = RSObj("Status")
					  LoginTf= RSObj("LoginTf")
					  TimeLimit=RSObj("TimeLimit")
					  enddate=RSObj("EndDate")
					  ZFVotes=RSObj("ZFVotes")
					  FFVotes=RSObj("FFVotes")
					  SFVotes=RSObj("SFVotes")
					  ClassID=RSObj("ClassID")
					  VerifyTF=RSObj("verifytf")
					  OnceTF=RSObj("oncetf")
					RSObj.Close:Set RSObj = Nothing
				Else
				  Flag = "Add"
				  status=1
				  TimeLimit=0
				  enddate=now
				  ZFVotes=0
				  FFVotes=0
				  SFVotes=0
				  LoginTf=1
				  VerifyTF=1
				  OnceTF=1
				End If
				With Response
				.Write "<html>"
				.Write "<head>"
				.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
				.Write "<title>新建PK主题</title>"
				.Write "</head>"
				.Write "<script src=""Include/Common.js"" language=""JavaScript""></script>"
				.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write "<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
				.Write " <div class='topdashed sort'>"
				If Flag = "Edit" Then
				 .Write "修改PK主题"
				Else
				 .Write "添加PK主题"
				End If
	            .Write "</div>"
				.Write "<br>"
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"
				.Write "  <form name=myform method=post action=""?Action=Save"">"
				.Write "   <input type=""hidden"" name=""Flag"" value=""" & Flag & """>"
				.Write "   <input type=""hidden"" name=""PKID"" value=""" & PKID & """>"
				.Write "   <input type=""hidden"" name=""Page"" value=""" & Page & """>"
				.Write "    <tr>"
				.Write "      <td height=""25"" align='right' width='85' class='clefttitle'><strong>PK主题名称:</strong></td>"
				.Write "      <td><input name=""Title"" type=""text"" id=""Title"" value=""" & Title & """ class=""textbox"" style=""width:60%""> 如:美军机赴台救灾，是否起到作用？</td>"
				 .Write "  </tr>"
				.Write "    <tr>"
				.Write "      <td height=""25"" align='right' width='85' class='clefttitle'><strong>指定频道:</strong></td>"
				.Write "      <td><select name=""ClassID"" style=""width:200;border-style: solid; border-width: 1"">"
		  If not IsObject(Application(KS.SiteSN&"_class")) Then KS.LoadClassConfig
			Dim ClassXML,Node
			Set ClassXML=Application(KS.SiteSN&"_class")
			For Each Node In ClassXML.documentElement.SelectNodes("class[@ks10=1 and @ks14=1]")
			  If Node.SelectSingleNode("@ks0").text = ClassID Then
			    .write "<option value='" &Node.SelectSingleNode("@ks0").text &"' selected>" & Node.SelectSingleNode("@ks1").text & "</option>"
			  else
			    .write "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
			  end if
			next
				.Write "      </select> 主要是起到按频道分类管理调用作用"
				.Write "      </td></tr>"
				 
				 
				 .Write "<tr>"
				.Write "  <td height=""25"" align='right' width='85' class='clefttitle'><strong>正方观点:</strong></td>"
				.Write "  <td><textarea ID='ZFTips' name='ZFTips' style='width:90%;height:60px'>" & ZFTips &"</textarea><br/></td></tr>"
				.Write "<tr><td height=""25"" align='right' width='85' class='clefttitle'><strong>反方观点:</strong></td>"
				.Write "<td><textarea ID='FFTips' name='FFTips' style='width:90%;height:60px'>" & FFTips &"</textarea></td></tr>"
				.Write "<tr><td height=""25"" align='right' width='85' class='clefttitle'><strong>背景新闻链接:</strong></td>"
				.Write "<td><input type='text' name='NewsLink' id='NewsLink' size='25' value='" & NewsLink & "'> </td></tr>"
				.Write "<tr><td height=""25"" align='right' width='85' class='clefttitle'><strong>得票情况:</strong></td>"
				.Write "<td>正方:<input type='text' style='text-align:center' name='ZFVotes' value='" & ZFVotes & "' size='4'> 反方:<input type='text' name='FFVotes' value='" & FFVotes & "' size='4' style='text-align:center'> 第三方:<input type='text' name='SFVotes' value='" & SFVotes & "' size='4' style='text-align:center'></td></tr>"
				.Write "          <tr>"
				.Write "            <td height=""25"" align='right' width='85' class='clefttitle'><strong>是否允许游客PK:</strong></td>"
				.Write "            <td>"
				.write "  <Input type='radio' name='LoginTf' value='0'"
				if LoginTf="0" then .write " checked"
				.Write ">允许"
				.write "  <Input type='radio' name='LoginTf' value='1'"
				if LoginTf="1" then .write " checked"
				.Write ">不允许"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr>"
				.Write "            <td height=""25"" align='right' width='85' class='clefttitle'><strong>每个用户只能PK一次:</strong></td>"
				.Write "            <td>"
				.write "  <Input type='radio' name='OnceTF' value='0'"
				if OnceTF="0" then .write " checked"
				.Write ">不是"
				.write "  <Input type='radio' name='OnceTF' value='1'"
				if OnceTF="1" then .write " checked"
				.Write ">是"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr>"
				.Write "            <td height=""25"" align='right' width='85' class='clefttitle'><strong>用户观点需要审核:</strong></td>"
				.Write "            <td>"
				.write "  <Input type='radio' name='VerifyTF' value='0'"
				if VerifyTF="0" then .write " checked"
				.Write ">不需要"
				.write "  <Input type='radio' name='VerifyTF' value='1'"
				if VerifyTF="1" then .write " checked"
				.Write ">需要"
				.Write "              </td>"
				.Write "          </tr>"
				
				
				.Write "          <tr>"
				.Write "            <td height=""25"" align='right' width='85' class='clefttitle'><strong>状态:</strong></td>"
				.Write "            <td>"
				.write "  <Input type='radio' name='status' value='0'"
				if status="0" then .write " checked"
				.Write ">锁定"
				.write "  <Input type='radio' name='status' value='1'"
				if status="1" then .write " checked"
				.Write ">正常"
				.Write "              </td>"
				.Write "          </tr>"
				
				
				.Write "          <tr>"
				.Write "            <td height=""25"" align='right' width='85' class='clefttitle'><strong>是否限定时间:</strong></td>"
				.Write "            <td>"
				
				.write "  <Input type='radio' onclick=""document.getElementById('timea').style.display='none';"" name='TimeLimit' value='0'"
				if TimeLimit="0" then .write " checked"
				.Write ">不限制"
				.write "  <Input type='radio'  onclick=""document.getElementById('timea').style.display='';"" name='TimeLimit' value='1'"
				if TimeLimit="1" then .write " checked"
				.Write ">限制时间"


               if TimeLimit="0" then
				.Write " <div id='timea' style='display:none'>"
			  Else
				.Write " <div id='timea'>"
			  End If
				.Write "<input type='text' name='enddate' value='" & enddate& "' size='30' class='textbox'> 格式:YYYY-MM-DD hh:mm:ss"
				.Write "</div>"
				.Write "              </td>"
				.Write "          </tr>"

				.Write "  </form>"
				.Write "</table>"
				.Write "</body>"
				.Write "</html>"
				.Write "<script language=""JavaScript"">" & vbCrLf
				.Write "<!--" & vbCrLf
				.Write "function CheckForm()" & vbCrLf
				.Write "{ var form=document.myform;" & vbCrLf
				.Write "  if (form.Title.value=='')" & vbCrLf
				.Write "   {" & vbCrLf
				.Write "    alert('请输入PK主题名称!');" & vbCrLf
				.Write "    form.Title.focus();" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf
				.Write "   if (form.ZFTips.value=='')" & vbCrLf
				.Write "   {" & vbCrLf
				'.Write "    alert('请输入活动介绍!');" & vbCrLf
				'.Write "    form.ZFTips.focus();" & vbCrLf
				'.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf

				.Write "   form.submit();"
				.Write "   return true;"
				.Write "}"
				.Write "//-->"
				.Write "</script>"
			 End With
		  End Sub
		  
		  '保存
		  Sub DoSave()
			Dim PKID, RSObj, SqlStr,ClassID,Title, AddDate, ZFTips, FFTips,TimeLimit,Flag, Page, RSCheck,Status,enddate,NewsLink,ZFVotes,FFVotes,SFVotes,LoginTf,VerifyTF,OnceTF
			Set RSObj = Server.CreateObject("Adodb.RecordSet")
			Flag = Request.Form("Flag")
			PKID = Request("PKID")
			Title = KS.G("Title")
			ZFTips = Request.Form("ZFTips")
			FFTips = Request.Form("FFTips")
			NewsLink=KS.G("NewsLink")
			Status = KS.ChkClng(KS.G("Status"))
			TimeLimit=KS.ChkClng(KS.G("TimeLimit"))
			ClassID=KS.G("ClassID")
			ZFVotes=KS.ChkClng(KS.G("ZFVotes"))
			FFVotes=KS.ChkClng(KS.G("FFVotes"))
			SFVotes=KS.ChkClng(KS.G("SFVotes"))
			LoginTf=KS.ChkClng(KS.G("LoginTf"))
			VerifyTF=KS.ChkClng(KS.G("VerifyTF"))
			OnceTF=KS.ChkClng(KS.G("OnceTF"))
			enddate=request("enddate")
			if not isdate(enddate) then enddate=now
			
			If Title = "" Then Call KS.AlertHistory("PK主题不能为空!", -1)
			If ZFTips = "" Then Call KS.AlertHistory("PK主题背景资料不能为空!", -1)
			
			Set RSObj = Server.CreateObject("Adodb.Recordset")
			If Flag = "Add" Then
			   RSObj.Open "Select ID From KS_PKZT Where Title='" & Title & "'", Conn, 1, 1
			   If Not RSObj.EOF Then
				  RSObj.Close
				  Set RSObj = Nothing
				  Response.Write ("<script>alert('对不起,PK主题名称已存在!');history.back(-1);</script>")
				  Exit Sub
			   Else
				RSObj.Close
				RSObj.Open "SELECT * FROM KS_PKZT Where 1=0", Conn, 1, 3
				RSObj.AddNew
				  RSObj("Title") = Title
				  RSObj("ClassID")=ClassID
				  RSObj("ZFTips") = ZFTips
				  RSObj("FFTips") = FFTips
				  RSObj("NewsLink")=NewsLink
				  RSObj("AddDate")=Now
				  RSObj("TimeLimit")=TimeLimit
				  RSObj("enddate") = enddate
				  RSObj("ZFVotes") = ZFVotes
				  RSObj("FFVotes") = FFVotes
				  RSObj("SFVotes") = SFVotes
				  RSObj("LoginTf") = LoginTf
				  RSObj("VerifyTf") = VerifyTf
				  RSObj("OnceTf") = OnceTf
				  RSObj("Status") =Status
				RSObj.Update
				 RSObj.Close
			  End If
			   Set RSObj = Nothing
			   Response.Write ("<script> if (confirm('PK主题添加成功!继续添加吗?')) {location.href='KS.PKZT.asp?Action=Add';}else{location.href='KS.PKZT.asp';parent.frames['BottomFrame'].location.href='KS.Split.asp?ButtonSymbol=Disabled&OpStr=PK系统管理 >> <font color=red>PK主题管理</font>';}</script>")
			ElseIf Flag = "Edit" Then
			  Page = Request.Form("Page")
			  RSObj.Open "Select ID FROM KS_PKZT Where Title='" & Title & "' And ID<>" & PKID, Conn, 1, 1
			  If Not RSObj.EOF Then
				 RSObj.Close
				 Set RSObj = Nothing
				 Response.Write ("<script>alert('对不起,PK主题名称已存在!');history.back(-1);</script>")
				 Exit Sub
			  Else
			   RSObj.Close
			   SqlStr = "SELECT * FROM KS_PKZT Where ID=" & PKID
			   RSObj.Open SqlStr, Conn, 1, 3
				  RSObj("Title") = Title
				  RSObj("ClassID")=ClassID
				  RSObj("ZFTips") = ZFTips
				  RSObj("FFTips") = FFTips
				  RSObj("NewsLink")=NewsLink
				  RSObj("TimeLimit")=TimeLimit
				  RSObj("enddate") = enddate
				  RSObj("ZFVotes") = ZFVotes
				  RSObj("FFVotes") = FFVotes
				  RSObj("SFVotes") = SFVotes
				  RSObj("LoginTf") = LoginTf
				  RSObj("VerifyTf") = VerifyTf
				  RSObj("OnceTf") = OnceTf
				  RSObj("Status") =Status
			   RSObj.Update
			   RSObj.Close
			   Set RSObj = Nothing
			  End If
			  Response.Write ("<script>alert('PK主题修改成功!');location.href='KS.PKZT.asp?Page=" & Page & "';parent.frames['BottomFrame'].location.href='KS.Split.asp?ButtonSymbol=Disabled&OpStr=PK系统管理 >> <font color=red>PK主题管理</font>';</script>")
			End If
		  End Sub
		  
		  '删除
		  Sub PKDelete()
		  		 Dim K, PKID, Page
				 Page = KS.G("Page")
				 PKID = Trim(KS.G("PKID"))
				 PKID = Split(PKID, ",")
				 For k = LBound(PKID) To UBound(PKID)
					Conn.Execute ("Delete From KS_PKZT Where ID =" & PKID(k))
				 Next
				 KS.Echo "<script>alert('恭喜,PK主题删除成功!');location.href='KS.PKZT.Asp';</script>"
		  End Sub
		  
	

End Class
%>
 
