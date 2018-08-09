<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New FriendLink
KSCls.Kesion()
Set KSCls = Nothing

Class FriendLink
        Private KS,KSCls
		Private MaxPerPage, Row
		Private CurrentPage
		Private totalPut,Action
		Private FolderID, I, SqlStr , RSObj
		Private Title , CreateDate, TempStr , GRS 
		Private KeyWord , SearchType, SearchParam 
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
   
		   If Not KS.ReturnPowerResult(0, "KSMS20001") Then                '检查友情链接操作(增和改)的权限检查
			  Call KS.ReturnErr(1, "")
			  Exit Sub
		   End If
		
		'收集搜索参数
		KeyWord = Request("KeyWord")
		SearchType = Request("SearchType")
		'搜索参数集合
		SearchParam = "KeyWord=" & KeyWord & "&SearchType=" & SearchType
		
		    Action=KS.G("Action")
			Select Case Action
			 Case "Verific"  '审核前台的申请
			   Call Link_VerificLink()
			 Case "VerificOK"
			   Call Link_VerificLinkOK()
			 Case "AddFolder","EditFolder" '添加修改类别
			   Call Link_AddFolder()
			 Case "ClassManage"
			   Call ClassManage()
			 Case "FolderSave"
			   Call Link_FolderSave()
			 Case "DelFolder"
			   Call Link_FolderDel()
			 Case "AddLink","EditLink"
			   Call Link_LinkAdd()
			 Case "SaveLink"
			   Call Link_SaveLink()
			 Case "DelLink"
			   Call Link_DelLink()
			 Case "Orders"
			   Call Orders()
			 Case Else
			   Call Link_MainList()
			 End Select
		End Sub
		
		Sub Head()
		 With KS
		 .echo "<html>"
		 .echo "<head>"
		 .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		 .echo "<title>友情链接管理</title>"
		 .echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		 .echo "<script language=""JavaScript"">" & vbCrLf
		 .echo "var FolderID='" & FolderID & "';         //友情链接类别ID" & vbCrLf
		 .echo "var KeyWord='" & KeyWord & "';           //搜索关键字" & vbCrLf
		 .echo "var Page='" & CurrentPage & "';         //当前页码" & vbCrLf
		 .echo "var SearchParam='" & SearchParam & "';   //搜索参数集合" & vbCrLf
		 .echo "</script>" & vbCrLf
		 .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
		 .echo "<script language=""JavaScript"" src=""../KS_Inc/jQuery.js""></script>"
		 .echo "<script language=""JavaScript"" src=""../KS_Inc/Kesion.Box.js""></script>"
		%>
		<script language="javascript">
		$(document).ready(function(){
		    $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
			$(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
		})
		function CreateFolder()
		{
		 new KesionPopup().PopupCenterIframe('添加友情链接类别','KS.FriendLink.asp?Action=AddFolder',600,300,'no')
		}
		function EditFolder(FolderID)
		{ 
		  new KesionPopup().PopupCenterIframe('编辑友情链接类别','KS.FriendLink.asp?Action=EditFolder&FolderID='+FolderID,600,300,'no')
		}
		function CreateLink()
		{
		 new KesionPopup().PopupCenterIframe('添加友情链接','KS.FriendLink.asp?Action=AddLink',650,430,'no')
		}
		function EditLink(linkID)
		{var ids=get_Ids(document.myform);
		 if (linkID==''){linkID=ids;}
		 if (linkID!=''){
			 if (linkID.indexOf(',')==-1){ 
			 new KesionPopup().PopupCenterIframe('编辑友情链接','KS.FriendLink.asp?Action=EditLink&Flag=EditLink&LinkID='+linkID,650,430,'no')
			 }else{alert('一次只能够编辑一个友情链接站点!'); }
		 }else{
		  alert('请选择要编辑的链接站点!');
		 }
		}
		
		function Delete(linkID)
		{   
		  if (linkID=='') linkID=get_Ids(document.myform);
		  if (linkID==''){
		   alert('请选择要删除的友情链接站点');
		  }else{
		   if (confirm('确定删除选中友情链接站点吗?')){
		     $("#Action").val("DelLink");
		     $("#myform")[0].action='KS.FriendLink.asp?page='+Page+'&id='+linkID+'&FolderID='+FolderID;
			 $("#myform").submit();
		    }
		  
		   }
		}
		function Verific()
		{
		 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr=常规管理 >> 友情链接管理 >> <font color=red>审核友情链接</font>&ButtonSymbol=Disabled';
		 location.href='KS.FriendLink.asp?Action=Verific';
		}
		
		function ClassToggle(f)
		{
			  setCookie("linkclassExtStatus",f)
			  $('#classNav').toggle('slow');
			  $('#classOpen').toggle('show');
		}
		</script>
		<%
		 .echo "</head>"
		 .echo "<body topmargin=""0"" leftmargin=""0"">"
		 .echo "<ul id='menu_top'>"
			  If KeyWord = "" Then
			   .echo "<li class='parent' onclick=""location.href='?action=ClassManage';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>分类管理</span></li>"
			   .echo "<li class='parent' onclick=""CreateFolder();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>添加类别</span></li>"
			    .echo "<li class='parent' onclick=""CreateLink()""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加链接</span></li>"
			   .echo "<li class='parent' onclick=""EditLink('')""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>编辑链接</span></li>"
			   .echo "<li class='parent' onclick=""Delete('');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>删除链接</span></li>"
			   .echo "<li class='parent' onclick=""Verific();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/verify.gif' border='0' align='absmiddle'>审核用户申请</span></li>"
			   .echo "<li class='parent' onclick=""parent.initializeSearch('友情链接站点')""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/s.gif' border='0' align='absmiddle'>搜索助理</span></li>"
		
			  Else
				 .echo ("<img src='Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""location.href='KS.FriendLink.asp';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr=常规管理 >> <font color=red>友情链接管理首页</font>';"">链接首页</span>")
			    .echo (">>> 搜索结果: ")
				Select Case SearchType
				 Case 0
				   .echo ("名称含有 <font color=red>" & KeyWord & "</font> 的站点")
				 Case 1
				   .echo ("简介含有 <font color=red>" & KeyWord & "</font> 的站点")
				 End Select
			   End If
			   
			 .echo "</ul>"
		  End With
		End Sub
		
		Sub Link_MainList()
		FolderID = Trim(Request.QueryString("FolderID"))
		If FolderID = "" Then FolderID = 0
		MaxPerPage = 16 '每页显示数量
		If Request("page") <> "" Then
			  CurrentPage = CInt(Request("page"))
		Else
			  CurrentPage = 1
		End If
		With KS
      
	     Call  Head()

		
		 '============分类显示,带记忆功能=======================================
			 Dim ExtStatus,CloseDisplayStr,ShowDisplayStr,classExtStatus
			 classExtStatus=request.cookies("linkclassExtStatus")
			 if classExtStatus="" Then classExtStatus=1
			 If classExtStatus=1 Then 
			  ExtStatus=2 :CloseDisplayStr="display:none;":ShowDisplayStr=""
			 Else 
			  ExtStatus=1 :CloseDisplayStr="":ShowDisplayStr="display:none;"
			 End If

			Dim RS,ClassXML,Node
			Set RS=Conn.Execute("Select FolderID,FolderName From KS_LinkFolder Order by OrderID")
			If Not RS.Eof Then Set ClassXML=KS.RsToXml(RS,"row","classxml")
			RS.Close:Set RS=Nothing
			If IsObject(ClassXML) Then
			.echo "<div id='classOpen' onclick=""ClassToggle("& ExtStatus& ")"" style='" & CloseDisplayStr &"cursor:pointer;text-align:center;position:absolute; z-index:2; left: 0px; top: 38px;' ><img src='images/kszk.gif' align='absmiddle'></div>"
		    .echo "<div id='classNav' style='" & ShowDisplayStr &"position:relative;height:auto;_height:30px;top:4px;line-height:30px;margin:8px 1px;border:1px solid #DEEFFA;background:#F7FBFE'>"
		    .echo "<div style='padding-top:2px;cursor:pointer;text-align:center;position:absolute; z-index:1; right: 0px; top: 2px;'  onclick=""ClassToggle(" & ExtStatus &")""> <img src='images/close.gif' align='absmiddle'></div>"
			 For Each Node In ClassXML.DocumentElement.SelectNodes("row")
			   .echo "<li style='margin:5px;float:left;width:100px'><img src='images/folder/folderopen.gif' align='absmiddle'><a href='?folderid=" & Node.SelectSingleNode("@folderid").text & "' title='" & Node.SelectSingleNode("@foldername").text & "'>" & KS.Gottopic(Node.SelectSingleNode("@foldername").text,10) & "</a></li>"
			 Next
			 .echo "</div>"
			End If
			 '=============================================================
			 
		
			
			  Dim Param
			  Param = " Where 1=1"
			  If KeyWord <> "" Then
				Select Case SearchType
				  Case 0
				   Param = Param & " And a.SiteName like '%" & KeyWord & "%'"
				  Case 1
				   Param = Param & " And a.Description like '%" & KeyWord & "%'"
				End Select
			  End If
			  If FolderID<>"" and FolderID<>"0" Then	 Param = Param & " And a.FolderID=" & FolderID
			  Param = Param & " Order BY a.AddDate desc"
			  SqlStr = "Select a.*,b.foldername From KS_Link a inner join KS_LinkFolder b ON A.FolderID=B.FolderID" & Param
		
			  .echo (" <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">")
			  .echo ("<form name=""myform"" id=""myform"" action=""KS.FriendLink.asp"" method=""post"">")
			  .echo ("<input name=""Action"" type=""hidden"" id=""Action"" value=""DelLink"">")
			  .echo ("       <tr align=""center"" width=""35"" class=""sort""><td>选择</td><td height=23>网站名称</td><td>所属类别</td><td>类 型</td><td>点击数</td><td>添加日期</td><td>推 荐</td><td>状 态</td><td>序号</td><td>管理操作</td></tr>")
		
		     Set RSObj = Server.CreateObject("AdoDb.RecordSet")
		     RSObj.Open SqlStr, Conn, 1, 1
			 If RSObj.EOF Then
			  .echo ("<tr><td colspan=10 align='center' class='splittd'>没有添加任何站点!</td></tr>")
			 Else
					        totalPut = RSObj.RecordCount
		
							If CurrentPage < 1 Then	CurrentPage = 1
		
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
		
							If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RSObj.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
							Call showContent
			End If
			 .echo " <tr>"
			 .echo " <td colspan='3'><div style='margin:5px'><b>选择：</b><a href='javascript:void(0)' onclick='Select(0)'>全选</a> -  <a href='javascript:void(0)' onclick='Select(1)'>反选</a> - <a href='javascript:void(0)' onclick='Select(2)'>不选</a> <input type='submit' class='button' value='删 除' onclick=""return(confirm('确定移除选中的站点吗?'))"">  <input type='submit' value='批量排序' class='button' onclick=""$('#Action').val('Orders');""></td></form>"
			 .echo "   <td align=""right"" colspan=8>"
				 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			.echo "   </td>"
			.echo "  </tr>"
		 .echo "</table>"
		 .echo "</div>"
		 .echo "</body>"
		 .echo "</html>"
		End With
		End Sub
		
		Sub ClassManage()
		  Dim RS,Node,ClassXml,FolderID
		  With KS
	       Call  Head()
		   	  .echo (" <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">")
			  .echo ("       <tr align=""center""><td height=23 width=200  class=""sort"">分类名称</td><td width=100 class=""sort"">链接数</td><td width=90  class=""sort"">排序号</td><td class=""sort"">管理操作</td></tr>")
			  Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "Select a.*,num From KS_LinkFolder a left join (select folderid,count(*) as num from ks_link group by folderid)b on a.folderid=b.folderid Order By a.OrderID",Conn,1,1
			  If Not RS.Eof Then
			   Set ClassXml=.RsToXml(RS,"row","root")
			  End If
			  RS.Close:Set RS=Nothing
			  If IsObject(ClassXml) Then
			    For Each Node In ClassXml.DocumentElement.SelectNodes("row")
				  FolderID=Node.SelectSingleNode("@folderid").text
				  .echo ("<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">")
				  .echo ("<td class='splittd'><img src='images/folder/folder.gif' align='absmiddle'><a href='?folderid=" & FolderID & "'>" & Node.SelectSingleNode("@foldername").text & "</a></td>")
				  .echo ("<td class='splittd' align=center>" & KS.ChkClng(Node.SelectSingleNode("@num").text) & "</td>")
				  .echo ("<td class='splittd' align=center>" & Node.SelectSingleNode("@orderid").text & "</td>")
				  .echo ("<td class='splittd' align=center><a href='javascript:EditFolder("& FolderID & ")'>修改</a> | <a href='KS.FriendLink.asp?Action=DelFolder&FolderID=" & FolderID & "' onclick=""return(confirm('删除分类将同时删除该分类下的链接,确定删除分类吗?'))"">删除 </td>")
				  .echo ("</tr>")
				Next
			  End If
              .echo ("</table>")
          End With
		End Sub
		
		Sub showContent()
		 With KS
		 Dim T, TitleStr, LockStr, ShortName, RecommendStr,LinkID
		 Do While Not RSObj.EOF
				If RSObj("Locked") = 1 Then
					LockStr = " <font color=red>锁</font>"
				Else
					LockStr = ""
				End If
				If RSObj("Recommend") = 1 Then
				 RecommendStr = "<font color=red>√</font>"
				Else
				 RecommendStr = "×"
				End If
				TitleStr = " TITLE='网站名称:" & RSObj("SiteName") & "&#13;&#10;网 址:" & RSObj("Url") & "&#13;&#10;添加时间:" & RSObj("AddDate") & "&#13;&#10;简要描述:" & RSObj("Description") & "'"
				LinkID=RSObj("LinkID")
				  .echo ("<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" &LinkID & "' onclick=""chk_iddiv('" & LinkID & "')""" & TitleStr & ">")
				  .echo ("<td class='splittd' align=center><input type='hidden' value='" & LinkID & "' name='LinkID'><input name='id'  onclick=""chk_iddiv('" & LinkID & "')"" type='checkbox' id='c"& LinkID & "' value='" & LinkID & "'></td>")				 
				  .echo ("<td class='splittd' height=25>&nbsp;<span ondblclick=""EditLink('" & LinkID & "');"" LinkID=""" & LinkID & """><span style=""cursor:default""><img src=""Images/Folder/Link.gif"" align=""absmiddle"">" & RSObj("SiteName") &LockStr & "</span></span></td>")
				 .echo ("<td class='splittd' align=""center"">" & RSObj("FolderName") & "</td>")
			  If RSObj("LinkType") = 0 Then
			   .echo ("<td class='splittd' align=""center"">文字链接</td>")
			  Else
			   .echo ("<td class='splittd' align=""center"">LOGO链接</td>")
			  End If
			  
			   .echo ("<td class='splittd' align=""center"">" & RSObj("Hits") & "</td>")
			   .echo ("<td class='splittd' align=""center"">" & FormatDateTime(RSObj("AddDate"),2) & "</td>")
			   .echo ("<td class='splittd' align=""center"">" & RecommendStr & "</td>")
			   .echo ("<td class='splittd' align=""center"">")
			    if RSObj("Verific")="1" Then
				 .echo ("已审核")
				else
				 .echo ("<span style='color:red'>未审核</span>")
				end if
			   .echo ("</td>")
			   .echo ("<td class='splittd' align=""center""><input type='text' name='OrderID' value='" & RSObj("OrderID") & "' size='6' class='textbox' style='text-align:center'></td>")
			   .echo ("<td class='splittd' align=""center""><a href=""javascript:EditLink('" & LinkID & "');"">修改</a> <a href=""javascript:Delete(" & LinkID & ")"">删除</a></td>")
			   .echo ("</tr>")
			  RSObj.MoveNext
			 I = I + 1
			 If I >= MaxPerPage Then Exit Do
			Loop
			RSObj.Close
			Set RSObj = Nothing
		  End With
		End Sub
		
		'批量排序
		Sub Orders()
		 Dim LinkID:LinkID=Replace(KS.G("LinkID")," ","")
		 Dim OrderID:OrderID=KS.G("OrderID")
		 Dim LinkIDArr,OrderIDArr,I
		 LinkIDArr=Split(LinkID,",")
		 OrderIDArr=Split(OrderID,",")
		 For I=0 To Ubound(LinkIDArr)
		  Conn.Execute("Update KS_Link Set OrderID=" & OrderIDArr(I) & " Where LinkID=" & LinkIDArr(i))
		 Next
		 KS.AlertHintScript "恭喜,批量排序设置成功!"
		End Sub
		
		'添加类别
		Sub Link_AddFolder()
			 With KS
			 .echo "<html>"
			 .echo "<head>"
			 .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			 .echo "<link href=""Include/Admin_Style.css"" rel=""stylesheet"">"
			 .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
			 .echo "<title>添加友情链接类别</title>"
			 .echo "</head>"
			
			Dim FolderID, FolderName, Descript, Flag,OrderID
			
			Flag = KS.G("Flag")
			FolderID = KS.G("FolderID")
			If Flag = "" Then Flag = "AddLink"
			If FolderID <> "" Then
			   
			   Dim RSObj:Set RSObj = Conn.Execute("Select FolderName,Description,OrderID From KS_LinkFolder Where FolderID=" & FolderID)
			  If Not RSObj.EOF Then
				 FolderName = RSObj(0)
				 Descript = RSObj(1)
				 OrderID=RSObj(2)
			  End If
			   RSObj.Close:Set RSObj = Nothing
			Else
				FolderName = "":Descript = "":OrderID=1
			End If
			
			 .echo "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			 .echo "  <form action=""KS.FriendLink.asp?Action=FolderSave"" name=""LinkForm"" method=""post"">"
			 .echo "   <input name=""Flag"" type=""hidden"" id=""Flag"" value=""" & Flag & """>"
			 .echo "   <input name=""FolderID"" type=""hidden"" value=""" & FolderID & """>"
			 .echo "  <br>"
			 .echo "        <table width=""100%"" class=""ctable"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			 .echo "          <tr class=""tdbg"">"
			 .echo "            <td width=""179"" class=""clefttitle"" height=""35""> <div align=""center"">类别名称</div></td>"
			 .echo "            <td> <input name=""FolderName"" value=""" & FolderName & """ type=""text"" id=""FolderName"" size=""30"" class=""textbox"">概况性说明文字 </td>"
			 .echo "          </tr>"
			 .echo "          <tr class=""tdbg"">"
			 .echo "            <td height=""35"" class=""clefttitle""> <div align=""center"">简要说明</div></td>"
			 .echo "            <td><textarea name=""Description"" rows=""8"" id=""Description"" style=""width:80%;height:100px;"" class=""textbox"">" & Descript & "</textarea></td>"
			 .echo "          </tr>"
			 .echo "          <tr class=""tdbg"">"
			 .echo "            <td width=""179"" height=""35"" class=""clefttitle""> <div align=""center"">排列序号</div></td>"
			 .echo "            <td width=""542""> <input name=""OrderID"" value=""" & OrderID & """ type=""text""  size=""8"" class=""textbox"">数字越小排在越前面 </td>"
			 .echo "          </tr>"
			 .echo "        </table>"
			 .echo "  </form>"
			 
			 .echo "<div id='save'>"
		.echo "<li class='parent' onclick=""return(CheckForm())""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/save.gif' border='0' align='absmiddle'>确定保存</span></li>"
		.echo "<li class='parent' onclick=""parent.closeWindow();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>关闭取消</span></li>"
		.echo "</div>"
			 
			 .echo "</body>"
			 .echo "</html>"
			 .echo "<Script Language=""javascript"">" & vbCrLf
			 .echo "<!--" & vbCrLf
			 .echo "function CheckForm()" & vbCrLf
			 .echo "{ var form=document.LinkForm;" & vbCrLf
			 .echo "   if (form.FolderName.value=='')" & vbCrLf
			 .echo "    {"
			 .echo "     alert(""请输入类别名称!"");" & vbCrLf
			 .echo "     form.FolderName.focus();" & vbCrLf
			 .echo "    return false;"
			 .echo "    }"
			 .echo "    form.submit();" & vbCrLf
			 .echo "    return true;" & vbCrLf
			
			 .echo "}" & vbCrLf
			 .echo "//-->" & vbCrLf
			 .echo "</Script>" & vbCrLf
			End With
		End Sub
		
		'保存类别
		Sub Link_FolderSave()
		With KS
		
		Dim FolderID, FolderName, Descript,TempObj, LinkRS, LinkSql,OrderID
		FolderID   = KS.ChkClng(KS.G("FolderID"))
		FolderName = KS.G("FolderName")
		Descript   = KS.G("Description")
		OrderID    = KS.ChkClng(KS.G("OrderID"))
			 If FolderName <> "" Then
				If Len(FolderName) >= 100 Then
					Call KS.AlertHistory("类别名称不能超过50个字符!", -1)
					Set KS = Nothing
					Exit Sub
				End If
			 Else
				Call KS.AlertHistory("请输入类别名称!", -1)
				Set KS = Nothing
				Exit Sub
			 End If
		   
			 Set TempObj = Conn.Execute("Select FolderName from [KS_LinkFolder] where FolderID<>" & FolderID & " and FolderName='" & FolderName & "'")
				If Not TempObj.EOF Then
					Call KS.AlertHintScript("数据库中已存在该类别名称!")
					Set KS = Nothing
					Exit Sub
				End If
			  Set LinkRS = Server.CreateObject("adodb.recordset")
			  LinkSql = "select * from [KS_LinkFolder] Where FolderID=" & FolderID
			  LinkRS.Open LinkSql, Conn, 1, 3
			  If LinkRS.Eof Then
			  LinkRS.AddNew
			  LinkRS("AddDate") = Now
			  End If
			  LinkRS("FolderName") = FolderName
			  LinkRS("Description") = Descript
			  LinkRS("OrderID")=OrderID
			  LinkRS.Update
			  LinkRS.Close
			  Set LinkRS = Nothing
			  If FolderID=0 Then
			   KS.echo ("<script>if (confirm('添加友情链接类别成功,继续添加吗?')) {location.href='KS.FriendLink.asp?Action=AddFolder';} else { top.frames[""MainFrame""].location.reload();}</script>")
		      Else
			  KS.echo ("<script>alert('修改友情链接类别成功!');top.frames[""MainFrame""].location.reload();</script>")
		      End If
		End With
		End Sub
		
		'删除类别
		Sub Link_FolderDel()
			Dim k, FolderID
			FolderID = Trim(KS.G("FolderID"))
			FolderID = Split(FolderID, ",")
			For k = LBound(FolderID) To UBound(FolderID)
				Conn.Execute ("Delete From KS_LinkFolder Where FolderID=" & FolderID(k))
				Conn.Execute ("Delete From KS_Link Where FolderID =" & FolderID(k))
			Next
			KS.AlertHintScript ("恭喜,分类删除成功!")
		End Sub
		
		'添加友情链接
		Sub Link_LinkAdd()
		With KS
		 .echo "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd"">"
         .echo "<html xmlns=""http://www.w3.org/1999/xhtml"">"
		 .echo "<head>"
		 .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		 .echo "<link href=""Include/admin_style.css"" rel=""stylesheet"">"
		 .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
		 .echo "<title>友情链接添加</title>"
		 .echo "</head>"
		
		Dim LinkID, SiteName, WebMaster, Email, PassWord, Locked, Url, LinkType, Hits, Recommend, Logo, AddDate, Descript, Flag, FolderID,OrderID,verific
		Dim CurrPath, InstallDir
		CurrPath = KS.GetCommonUpFilesDir()
		
		Flag = KS.G("Flag")
		LinkID = KS.G("LinkID")
		FolderID = KS.G("FolderID")
		
		If Flag = "" Then Flag = "AddLink"
		If LinkID <> "" Then
		   Dim RSObj
		  Set RSObj = Conn.Execute("Select top 1 * From KS_Link Where LinkID=" & LinkID)
		  If Not RSObj.EOF Then
			 SiteName = Trim(RSObj("SiteName"))
			 WebMaster = Trim(RSObj("WebMaster"))
			 Email = Trim(RSObj("Email"))
			 Locked = Trim(CStr(RSObj("Locked")))
			 Url = Trim(RSObj("Url"))
			 Logo = Trim(RSObj("Logo"))
			 LinkType = Trim(RSObj("LinkType"))
			 Hits = Trim(RSObj("Hits"))
			 Recommend = RSObj("Recommend")
			 AddDate = Trim(RSObj("AddDate"))
			 Descript = Trim(RSObj("Description"))
			 FolderID = RSObj("FolderID")
			 OrderID = RSObj("OrderID")
			 verific = RSObj("verific")
		  End If
		   RSObj.Close
		   Set RSObj = Nothing
		End If
		  If WebMaster = "" Then WebMaster = "保密"
		  If Email = "" Then Email = "@"
		  If Url = "" Then Url = "http://"
		  If Logo = "" Then Logo = "http://"
		  If AddDate = "" Then AddDate = Now
		  If Hits = "" Then Hits = 0
		  If LinkType = "" Then LinkType = 0
		  If Recommend = "" Then Recommend = 0
		  If OrderID="" Then OrderID=1
		  If verific="" Then verific=1
		
		 .echo "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		 .echo "  <form action=""KS.FriendLink.asp?Action=SaveLink"" name=""LinkForm"" method=""post"">"
		 .echo "   <input name=""Flag"" type=""hidden"" id=""Flag"" value=""" & Flag & """>"
		 .echo "   <input name=""LinkID"" type=""hidden"" value=""" & LinkID & """>"
		 .echo "        <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""ctable"">"
		 .echo "          <tr class=""tdbg"">"
		 .echo "            <td width=""179"" class=""clefttitle"" height=""25"" align=""center"">网站名称</td>"
		 .echo "            <td width=""542"" height=""25"">"
					
					If Flag = "EditLink" Then
						  .echo ("<input name=""SiteName""  value=""" & SiteName & """ type=""text"" id=""SiteName"" size=""38"" class=""textbox"">")
					  Else
						  .echo ("<input name=""SiteName""  type=""text"" id=""SiteName"" size=""38"" class=""textbox"">")
					  End If
					 
		 .echo "             </td>"
		 .echo "          </tr>"
		 .echo "          <tr class=""tdbg"">"
		 .echo "            <td height=""25"" class=""clefttitle"" align=""center"">所属类别</td>"
		 .echo "            <td height=""25"">"
		 .echo "              <select Name=""FolderID"" class=""textbox"">"
					   
					Dim GRS
					Set GRS = Conn.Execute("Select FolderID,FolderName From KS_LinkFolder")
					 Do While Not GRS.EOF
					   If CStr(FolderID) = CStr(GRS(0)) Then
						 .echo ("<Option value=" & GRS(0) & " selected>" & GRS(1) & "</OPTION>")
					   Else
						 .echo ("<Option value=" & GRS(0) & ">" & GRS(1) & "</OPTION>")
					   End If
					   GRS.MoveNext
					 Loop
					 GRS.Close
					 Set GRS = Nothing
				   
		  .echo "             </Select> </td>"
		  .echo "         </tr>"
		  .echo "          <tr class=""tdbg"">"
		 .echo "            <td height=""25"" class=""clefttitle"" align=""center"">网站站长</td>"
		 .echo "            <td height=""25"">"
		 .echo "              <input name=""WebMaster"" type=""text"" size=""38"" value=""" & WebMaster & """ class=""textbox""></td>"
		 .echo "          </tr>"
		 .echo "          <tr class=""tdbg"">"
		 .echo "            <td height=""25"" class=""clefttitle"" align=""center"">站长信箱</td>"
		 .echo "            <td height=""25"">"
		 .echo "              <input name=""Email"" type=""text"" size=""38"" value=""" & Email & """ class=""textbox""></td>"
		 .echo "          </tr>"
				  
		 .echo "          <tr class=""tdbg"">"
		 .echo "            <td height=""25"" class=""clefttitle"" align=""center"">网站密码</td>"
		 .echo "            <td height=""25"">"
		 .echo "              <input name=""PassWord"" type=""password"" size=""42"" class=""textbox"">"
		If Flag <> "EditLink" Then  .echo "不少于6位, 用于修改信息时用" Else  .echo "不修改请留空"
		 .echo " </td>"
		 .echo "          </tr>"
		 .echo "          <tr class=""tdbg"">"
		 .echo "            <td height=""25"" class=""clefttitle"" align=""center"">是否审核</td>"
		 .echo "            <td height=""25"">"
					 If verific = "1" Then
					  .echo ("<input type=""radio"" name=""verific"" value=""0""> 未审 ")
					  .echo ("<input type=""radio"" name=""verific"" value=""1"" checked> 已审 ")
					 Else
					   .echo ("<input type=""radio"" name=""verific"" value=""0"" checked> 未审 ")
					   .echo ("<input type=""radio"" name=""verific"" value=""1""> 已审 ")
					 End If
		 .echo "           </td>"
		 .echo "          </tr>"
				 
		 .echo "          <tr class=""tdbg"">"
		 .echo "            <td height=""25"" class=""clefttitle"" align=""center"">是否锁定</td>"
		 .echo "            <td height=""25"">"
					
					 If Locked = "1" Then
					  .echo ("<input type=""radio"" name=""Locked"" value=""0""> 正常 ")
					  .echo ("<input type=""radio"" name=""Locked"" value=""1"" checked> 锁定 ")
					 Else
					   .echo ("<input type=""radio"" name=""Locked"" value=""0"" checked> 正常 ")
					   .echo ("<input type=""radio"" name=""Locked"" value=""1""> 锁定 ")
					 End If
					  
		 .echo "              　　<font color=""#FF0000"">锁定的网站不能在前台显示和管理</font></td>"
		 .echo "          </tr>"
		
		 .echo "          <tr class=""tdbg"">"
		 .echo "            <td height=""25"" class=""clefttitle"" align=""center"">网站地址</td>"
		 .echo "            <td height=""25""><input name=""Url"" type=""text"" class=""textbox"" value=""" & Url & """ id=""Url"" size=""38""></td>"
		 .echo "          </tr>"
		 .echo "          <tr class=""tdbg"">"
		 .echo "            <td height=""25"" class=""clefttitle"" align=""center"">链接类型</td>"
		 .echo "            <td height=""25"">"
					 
					 If Trim(LinkType) = "1" Then
						   .echo ("<input type=""radio"" name=""LinkType"" onclick=""SetLogoArea('none')"" value=""0""> 文字链接 ")
						   .echo ("<input type=""radio"" name=""LinkType"" onclick=""SetLogoArea('')"" value=""1"" checked>  LOGO链接 ")
					  Else
						   .echo ("<input type=""radio"" name=""LinkType"" onclick=""SetLogoArea('none')"" value=""0"" checked> 文字链接 ")
						   .echo ("<input type=""radio"" name=""LinkType"" onclick=""SetLogoArea('')"" value=""1"">  LOGO链接 ")
					  End If
				   
		 .echo "             </td>"
		 .echo "          </tr>"
		If Trim(LinkType) = "1" Then
		 .echo "          <tr ID=""LinkArea"" class=""tdbg"">"
		Else
		 .echo ("         <tr Style=""display:none"" ID=""LinkArea"" class=""tdbg"">")
		End If
		 .echo "            <td height=""25"" class=""clefttitle"" align=""center"">Logo 地址</td>"
		 .echo "            <td height=""25""><input name=""Logo"" type=""text"" class=""textbox"" value=""" & Logo & """ id=""Logo"" size=""38""><input name=""SelectPic"" onClick=""OpenThenSetValue('Include/SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.LinkForm.Logo);"" class=""button"" type=""button"" id=""SelectPic"" value=""选择图片""></td>"
		 .echo "          </tr>"
		 .echo "          <tr class=""tdbg"">"
		 .echo "            <td height=""25"" class=""clefttitle"" align=""center"">点 击 数</td>"
		 .echo "            <td height=""25""><input name=""Hits"" type=""text"" class=""textbox"" value=""" & Hits & """ id=""Hits"" size=""10""></td>"
		 .echo "          </tr>"
		 .echo "          <tr class=""tdbg"">"
		 .echo "            <td height=""25"" class=""clefttitle"" align=""center"">推荐站点</td>"
		 .echo "            <td height=""25"">"
					 
					 If Trim(Recommend) = "1" Then
						   .echo ("<input type=""radio"" name=""Recommend""  value=""1"" checked> 是 ")
						   .echo ("<input type=""radio"" name=""Recommend"" value=""0"">  否 ")
					  Else
						   .echo ("<input type=""radio"" name=""Recommend""  value=""1""> 是 ")
						   .echo ("<input type=""radio"" name=""Recommend""  value=""0"" checked>  否 ")
					  End If
				   
		 .echo "             </td>"
		 .echo "          </tr>"
		 .echo "          <tr class=""tdbg"">"
		 .echo "            <td height=""25"" class=""clefttitle"" align=""center"">排列序号</td>"
		 .echo "            <td height=""25""><input name=""OrderID"" type=""text"" class=""textbox"" value=""" & OrderID & """ size=""10""> 数字越小越前面</td>"
		 .echo "          </tr>"
		 .echo "          <tr class=""tdbg"">"
		 .echo "            <td height=""25"" class=""clefttitle"" align=""center"">网站简介</td>"
		 .echo "            <td height=""25"">"
		 .echo "              <textarea name=""Description"" rows=""3"" id=""Description"" style=""width:80%;border-style: solid; border-width: 1"">" & Descript & "</textarea></td>"
		 .echo "          </tr>"
		 .echo "        </table>"
		.echo "<div id='save'>"
		.echo "<li class='parent' onclick=""return(CheckForm())""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/save.gif' border='0' align='absmiddle'>确定保存</span></li>"
		.echo "<li class='parent' onclick=""parent.closeWindow();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>关闭取消</span></li>"
		.echo "</div>"
		 .echo "  </form>"
		 .echo "</body>"
		 .echo "</html>"
		 .echo "<Script Language=""javascript"">" & vbCrLf
		 .echo "<!--" & vbCrLf
		 .echo "function SetLogoArea(Value)" & vbCrLf
		 .echo "{"
		 .echo "   document.all.LinkArea.style.display=Value;"
		 .echo "}" & vbCrLf
		 .echo "function CheckForm()" & vbCrLf
		 .echo "{ var form=document.LinkForm;" & vbCrLf
		 .echo "   if (form.SiteName.value=='')" & vbCrLf
		 .echo "    {"
		 .echo "     alert(""请输入网站名称!"");"
		 .echo "     form.SiteName.focus();"
		 .echo "     return false;" & vbCrLf
		 .echo "    }" & vbCrLf
		 .echo "   if (form.WebMaster.value=='')" & vbCrLf
		 .echo "    {"
		 .echo "     alert(""请输入网站站长!"");"
		 .echo "     form.WebMaster.focus();"
		 .echo "     return false;" & vbCrLf
		 .echo "    }" & vbCrLf
		 .echo "    if ((form.Email.value!='')&&(form.Email.value!='@')&&(is_email(form.Email.value)==false))" & vbCrLf
		 .echo "    {"
		 .echo "    alert('非法电子邮箱!');" & vbCrLf
		 .echo "     form.Email.focus();" & vbCrLf
		 .echo "     return false;" & vbCrLf
		 .echo "    }" & vbCrLf
		
			If Flag <> "EditLink" Then

		 .echo "    if (form.PassWord.value!='' && form.PassWord.value.length<6)"
		 .echo "    {"
		 .echo "      alert(""网站密码不能少于6位!"");"
		 .echo "     form.PassWord.focus();"
		 .echo "     return false;"
		 .echo "    }"
			End If
		 .echo "   if (form.Url.value=='')" & vbCrLf
		 .echo "    {" & vbCrLf
		 .echo "     alert(""请输入网站地址"");" & vbCrLf
		 .echo "     form.Url.focus();" & vbCrLf
		 .echo "     return false;" & vbCrLf
		 .echo "    }" & vbCrLf
		 .echo "    form.submit();" & vbCrLf
		 .echo "    return true;" & vbCrLf
		 .echo "}" & vbCrLf
		 .echo "//-->" & vbCrLf
		 .echo "</Script>"
		End With
		End Sub
		
		'保存链接
		Sub Link_SaveLink()
		With KS
				 .echo "<html>"
				 .echo "<head>"
				 .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
				 .echo "<link href=""Include/admin_style.css"" rel=""stylesheet"">"
				 .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
				 .echo "<title>保存新建友情链接</title>"
				 .echo "</head>"
				
				Dim LinkID, FolderID, SiteName, WebMaster, Email, PassWord, ConPassWord, Locked, Url, LinkType, Logo, Hits, Recommend, Descript, TrueIP, AddDate,TempObj, LinkRS, LinkSql,OrderID,verific
				LinkID = KS.ChkClng(KS.G("LinkID"))
				
				SiteName = KS.R(KS.G("SiteName"))
				WebMaster = KS.R(KS.G("Webmaster"))
				Email = KS.G("Email")
				FolderID = KS.G("FolderID")
				PassWord = KS.G("PassWord")
				ConPassWord = KS.G("ConPassWord")
				OrderID=KS.ChkClng(KS.G("OrderID"))
				verific=KS.ChkClng(KS.G("verific"))
				
				PassWord = MD5(KS.R(PassWord),16)
				Locked = KS.G("Locked")
				Url = KS.G("Url")
				LinkType = KS.G("LinkType")
				Logo = KS.G("Logo")
				Hits = KS.R(KS.G("Hits"))
				Recommend = KS.G("Recommend")
				AddDate = KS.G("AddDate")
				Descript = KS.R(KS.G("Description"))
				
				If SiteName <> "" Then
						If Len(SiteName) >= 200 Then
							Call KS.AlertHistory("网站名称不能超过100个字符!", -1)
							Set KS = Nothing
							Exit Sub
						End If
				 Else
						Call KS.AlertHistory("请输入网站名称!", -1)
						Set KS = Nothing
						Exit Sub
				 End If
				   
		
					    'Set TempObj = Conn.Execute("Select SiteName from [KS_Link] where LinkID<>" & LinkID & " and SiteName='" & SiteName & "'")
						'If Not TempObj.EOF Then
						'     TempObj.Close:Set TempObj=Nothing
						'	 Call KS.AlertHistory("数据库中已存在该友情链接的站点名称!",-1)
						'	 Set KS = Nothing
						'	 Exit Sub
						'End If
					  Set LinkRS = Server.CreateObject("adodb.recordset")
					  LinkSql = "select * from [KS_Link] Where LinkID=" & LinkID
					  LinkRS.Open LinkSql, Conn, 1, 3
					  If LinkRS.Eof Then 
					   LinkRS.AddNew
					   LinkRS("AddDate") = Now
					  End If
					  LinkRS("SiteName") = SiteName
					  LinkRS("WebMaster") = WebMaster
					  LinkRS("Email") = Email
					  LinkRS("FolderID") = FolderID
					  If KS.G("PassWord")<>"" Then
					  LinkRS("PassWord") = PassWord
					  End If
					  LinkRS("Locked") = Locked
					  LinkRS("Url") = Url
					  LinkRS("LinkType") = LinkType
					  LinkRS("Logo") = Logo
					  LinkRS("Hits") = Hits
					  LinkRS("Recommend") = Recommend
					  LinkRS("Description") = Descript
					  LinkRS("OrderID") = OrderID
					  LinkRS("Verific") = verific
					  LinkRS.Update
					  If LinkID=0 Then
					   LinkRS.MoveLast
					   Call KS.FileAssociation(1018,LinkRS("LinkID"),Logo,0)
					  Else
					   Call KS.FileAssociation(1018,LinkID,Logo,1)
					  End If
					  LinkRS.Close
					  Set LinkRS = Nothing
					 If LinkID=0 Then
					   .echo ("<script>if (confirm('添加友情链接成功,继续添加吗?')) {location.href='KS.FriendLink.asp?Action=AddLink&FolderID=" & FolderID & "';} else {top.frames[""MainFrame""].location.reload();}</script>")
					Else
					  .echo ("<script>alert('修改友情链接成功!');top.frames[""MainFrame""].location.reload();</script>")
					End If
				
			End With
		End Sub
		
		'删除友情链接站点
		Sub Link_DelLink()
			Dim Verific,k, LinkID,RSObj,FolderID, Page
			LinkID = Trim(KS.G("ID"))
			If LinkID="" Then KS.AlertHintScript "没有选择要删除的站点!" : Exit Sub
			Conn.Execute ("Delete From KS_UploadFiles Where ChannelID=1018 and infoid in(" & LinkID & ")")
			Conn.Execute ("Delete From KS_Link Where LinkID in (" & LinkID & ")")
			If KS.G("comefrom")="Verify" Then
			 KS.AlertHintScript "恭喜,删除成功!"
			Else
			KS.echo "<script>alert('恭喜,删除成功!');location.href='?page=" & KS.G("Page") & "&FolderID=" & KS.G("FolderID") & "';</script>"
			End If
		End Sub
		
		
		
		'审核前台的申请
		Sub Link_VerificLink()
		
		Row = 8         '每行显示数
		MaxPerPage = 20 '每页显示数量
		If KS.G("page") <> "" Then
			  CurrentPage = CInt(KS.G("page"))
		Else
			  CurrentPage = 1
		End If
		With KS
		 .echo "<html>"
		 .echo "<head>"
		 .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">" & vbCrLf
		 .echo "<title>管理员管理</title>" & vbCrLf
		 .echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">" & vbCrLf
		 .echo "<script language=""JavaScript"">" & vbCrLf
		 .echo "var Page='" & CurrentPage & "';         //当前页码" & vbCrLf
		 .echo "</script>" & vbCrLf
		 .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
		 .echo "<script language=""JavaScript"" src=""../KS_Inc/jQuery.js""></script>" & vbCrLf
		 .echo "<script language=""JavaScript"" src=""../KS_Inc/Kesion.Box.js""></script>" & vbCrLf
%>
<script language="javascript">
function ViewLink(linkID)
{
var ids=get_Ids(document.myform);
		 if (linkID==''){linkID=ids;}
		 if (linkID!=''){
			 if (linkID.indexOf(',')==-1){ 
			 new KesionPopup().PopupCenterIframe('编辑友情链接','KS.FriendLink.asp?Action=EditLink&Flag=EditLink&LinkID='+linkID,650,430,'no')
			 }else{alert('一次只能够编辑一个友情链接站点!'); }
		 }else{
		  alert('请选择要编辑的链接站点!');
		 }
}
function VerificLink()
{
var ids=get_Ids(document.myform);
 if (ids!='')
  { 
    $("#Action").val("VerificOK");
	$("#myform").submit();
  }
  else
   alert('请选择要审核的友情链接站点!');
}
	function Delete(linkID)
		{   
		  linkID=get_Ids(document.myform);
		  if (linkID==''){
		   alert('请选择要删除的友情链接站点');
		  }else{
		   if (confirm('确定删除选中友情链接站点吗?')){
		     $("#Action").val("DelLink");
		     $("#myform")[0].action='KS.FriendLink.asp?comefram=Verify&id='+linkID;
			 $("#myform").submit();
		    }
		  
		   }
		}

</script>
<%
		 .echo "</head>" & vbCrLf
		 .echo "<body  topmargin=""0"" leftmargin=""0"">" & vbCrLf
		 .echo "<ul id='menu_top'>"
		 .echo "<li class='parent' onclick=""VerificLink();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/verify.gif' border='0' align='absmiddle'>批量审核</span></li>"
		 .echo "<li class='parent' onclick=""Delete('');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>删除站点</span></li>"
		 .echo "<li class='parent' onclick=""location.href='KS.FriendLink.asp';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr=常规管理 >> <font color=red>友情链接管理</font>&ButtonSymbol=Disabled'""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>返回首页</span></li>"
		  .echo "</ul>"
		  .echo ("<div style=""height:94%; overflow: auto; width:100%"">")
		
		  SqlStr = "Select a.*,b.foldername From KS_Link A Left Join KS_LinkFolder B On A.FolderID=B.FolderID Where a.Verific=0 Order By a.AddDate Desc"
		  .echo (" <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">")
		  .echo (" <form name=""myform"" id=""myform"" action=""KS.FriendLink.asp?comefrom=Verify"" method=""post"">")
		  .echo (" <Input type=""hidden"" id=""Action"" name=""Action"" value=""DelLink"">")
		  .echo ("<tr align=""center""><td class=""sort"" width=""35"">选择</td><td height=23 width=220  class=""sort"">网站名称</td><td width=90  class=""sort"">所属类别</td><td width=100 class=""sort"">申请类型</td><td width=60 class=""sort"">站长</td><td width=120  class=""sort"">申请日期</td><td width=80 class=""sort"">状态</td></tr>")
		
		  Set RSObj = Server.CreateObject("AdoDb.RecordSet")
		  RSObj.Open SqlStr, Conn, 1, 1
			If RSObj.EOF Then
			 .echo ("<tr><td colspan=10 class='splittd' align='center'>没有未审核的链接!</td></tr>")
			Else
					totalPut = RSObj.RecordCount
		
							If CurrentPage < 1 Then	CurrentPage = 1
		
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
		
							If CurrentPage = 1 Then
								Call showLinkContent
							Else
								If (CurrentPage - 1) * MaxPerPage < totalPut Then
									RSObj.Move (CurrentPage - 1) * MaxPerPage
									Call showLinkContent
								Else
									CurrentPage = 1
									Call showLinkContent
								End If
							End If
			End If
			
			 .echo " <tr>"
			 .echo " <td colspan='13' height='35'><b>选择：</b><a href='javascript:void(0)' onclick='Select(0)'>全选</a> -  <a href='javascript:void(0)' onclick='Select(1)'>反选</a> - <a href='javascript:void(0)' onclick='Select(2)'>不选</a> <input type='submit' class='button' value='批量删除' onclick=""return(confirm('确定移除选中的站点吗?'))"">  <input type='submit' value='批量审核' class='button' onclick=""$('#Action').val('VerificOK');""></td></tr></form>"
			 .echo " <tr>"
			 .echo "   <td align=""center"" colspan=15>"
				 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			.echo "   </td>"
			.echo "  </tr>"
		 .echo "</table>"
		 .echo "</div>"
		 .echo "</body>"
		 .echo "</html>"
		End With
		End Sub
		Sub showLinkContent()
		 With KS
		 Dim T, TitleStr, ShortName, RecommendStr,LinkID
		 Do While Not RSObj.EOF
				If RSObj("Recommend") = 1 Then
				 RecommendStr = "<font color=red>√</font>"
				Else
				 RecommendStr = "×"
				End If
				LinkID=RSObj("LinkID")
				TitleStr = " TITLE='网站名称:" & RSObj("SiteName") & "&#13;&#10;网 址:" & RSObj("Url") & "&#13;&#10;添加时间:" & RSObj("AddDate") & "&#13;&#10;简要描述:" & RSObj("Description") & "'"
				.echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & LinkID & "' onclick=""chk_iddiv('" & LinkID & "')""" & TitleStr & ">"
			.echo "<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" & LinkID & "')"" type='checkbox' id='c"& LinkID & "' value='" & LinkID & "'></td>"
				 .echo ("<td class='splittd' height=25>&nbsp;<span ondblclick=""ViewLink(this.LinkID);"" LinkID=""" & RSObj("LinkID") & """><img src=Images/Folder/Link.gif border=0 align=absmiddle><span style=""cursor:default"">" & RSObj("SiteName") & "</span></span></td>")
				 .echo ("<td  class='splittd' align=""center"">" & RSObj("FolderName") & "</td>")
			  If RSObj("LinkType") = 0 Then
			   .echo ("<td class='splittd' align=""center"">文字链接</td>")
			  Else
			   .echo ("<td class='splittd' align=""center"">LOGO链接</td>")
			  End If
			  
			   .echo ("<td class='splittd' align=""center"">" & RSObj("WebMaster") & "</td>")
			   .echo ("<td class='splittd' align=""center"">" & RSObj("AddDate") & "</td>")
			   .echo ("<td class='splittd' align=""center""><font color=""red"">未审核</font></td>")
			   .echo ("</tr>")
			  RSObj.MoveNext
			   I = I + 1
			 If RSObj.EOF Or I>=MaxPerPage Then Exit Do
			Loop
			RSObj.Close:Set RSObj = Nothing
		 End With
		End Sub
		
		'审核
		Sub Link_VerificLinkOK()
			Dim LinkID, Page, I, RS
		   LinkID = Trim(KS.G("ID"))
		   LinkID = Split(LinkID, ",")
		   Page = KS.G("Page")
		   Set RS = Server.CreateObject("Adodb.Recordset")
		   For I = LBound(LinkID) To UBound(LinkID)
			 RS.Open "Select Verific From KS_Link Where LinkID=" & LinkID(I), Conn, 1, 3
			  RS("Verific") = 1
			  RS.Update
			  RS.Close
		   Next
		   Set RS = Nothing
		   KS.echo ("<script>location.href='KS.FriendLink.asp?Action=Verific&Page=" & Page & "';</script>")
		End Sub
End Class
%> 
