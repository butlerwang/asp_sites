<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%

Dim KSCls
Set KSCls = New JS_Main
KSCls.Kesion()
Set KSCls = Nothing

Class JS_Main
        Private KS
		'========================================================================
		Private JSSql, JSRS, FolderID, JSID, ChannelID, Channel, Action
		Private i, totalPut, CurrentPage, JSType
		Private KeyWord, SearchType, StartDate, EndDate
		'搜索参数集合
		Private SearchParam
		Private MaxPerPage
		Private Row 
		'========================================================================
		Private Sub Class_Initialize()
		  MaxPerPage = 96
		  Row = 8
		  Set KS=New PublicCls
		   Call KS.DelCahe(KS.SiteSn & "_labellist")
		   Call KS.DelCahe(KS.SiteSn & "_ReplaceFreeLabel")
		   Call KS.DelCahe(KS.SiteSn & "_jslist")
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		'采集搜索信息
		KeyWord = KS.G("KeyWord")
		SearchType = KS.G("SearchType")
		StartDate = KS.G("StartDate")
		EndDate = KS.G("EndDate")
		SearchParam = "KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate
		JSType = KS.G("JSType"):If JSType = "" Then JSType = 0
		
		Select Case KS.G("JsAction")
		 Case "JSDel"
		   Call JSDel()
		 Case "JSFolderDel"
		   Call JSFolderDel()
		 Case "JSView"
		   Call JSView()
		 Case Else
		   Call JSMainList()
		End Select
		End Sub
		
		Sub JSMainList()
		   With Response
			If JSType = 0 Then
				If Not KS.ReturnPowerResult(0, "KMTL10004") Then                '系统JS管理的权限检查
				  Call KS.ReturnErr(1, "")
				  .End
				End If
			ElseIf JSType = 1 Then
				If Not KS.ReturnPowerResult(0, "KMTL10005") Then                '自由JS管理的权限检查
				  Call KS.ReturnErr(1, "")
				  .End
				End If
			End If
			
			If Not IsEmpty(KS.G("page")) And KS.G("page") <> "" Then
				  CurrentPage = CInt(KS.G("page"))
			Else
				  CurrentPage = 1
			End If
			Action = KS.G("Action")
			FolderID = Trim(KS.G("FolderID"))
			If FolderID = "" Then FolderID = "0"
			Dim UPFolderRS, ParentID
			Set UPFolderRS = Conn.Execute("select * from [KS_LabelFolder] where  ID ='" & FolderID & "'")
			If Not UPFolderRS.EOF Then
			 ParentID = UPFolderRS("ParentID")
			End If
			UPFolderRS.Close:Set UPFolderRS = Nothing
		    .Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd"">"
			.Write "<head>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write "<title>JS列表</title>"
			.Write "<link href=""Admin_Style.CSS"" rel=""stylesheet"">"
			.Write "<script language=""JavaScript"">"
			.Write "var FolderID='" & FolderID & "';         //目录ID" & vbCrLf
			.Write "var ParentID='" & ParentID & "'; //父栏目ID" & vbCrLf
			.Write "var Page='" & CurrentPage & "';   //当前页码" & vbCrLf
			.Write "var KeyWord='" & KeyWord & "';    //关键字" & vbCrLf
			.Write "var SearchParam='" & SearchParam & "';  //搜索参数集合" & vbCrLf
			.Write "var Action='" & Action & "';" & vbCrLf
			.Write "var JSID='" & JSID & "';" & vbCrLf
			.Write "var JSType=" & JSType & ";" & vbCrLf
			.Write "</script>" & vbCrLf
		    .Write "<script language=""JavaScript"" src=""../../ks_inc/Common.js""></script>"
		    .Write "<script language=""JavaScript"" src=""../../ks_inc/jQuery.js""></script>"
		    .Write "<script language=""JavaScript"" src=""../../ks_inc/lhgdialog.js""></script>"
			.Write "<script language=""JavaScript"" src=""ContextMenu.js""></script>"
			.Write "<script language=""JavaScript"" src=""SelectElement.js""></script>"
			%>
			<script language="javascript">
			var DocElementArrInitialFlag=false;
			var DocElementArr = new Array();
			var DocMenuArr=new Array();
			var SelectedFile='',SelectedFolder='';
			$(document).ready(function(){
				if (DocElementArrInitialFlag) return;
				InitialDocElementArr('FolderID','JSID');
				InitialDocMenuArr();
				 DocElementArrInitialFlag=true;
			});
			function InitialDocMenuArr()
			{  
				if (KeyWord=='')
				{
				 DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.CreateFolder();",'新建目录(N)','disabled');
				 DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.AddJS('');",'新建JS(M)','disabled');
				 DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				}
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.JSView();",'预 览(V)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'全 选(A)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Edit('');",'编 辑(E)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Delete('');",'删 除(D)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('parent.ChangeUp();','后 退(B)','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('parent.Reload();','刷 新(Z)','');
			}
			function DocDisabledContextMenu()
			{
			   var TempDisabledStr=''; 
			   if (FolderID=='0') TempDisabledStr='后 退(B),';
				DisabledContextMenu('FolderID','JSID',TempDisabledStr+'预 览(V),编 辑(E),删 除(D)','预览(V),编 辑(E)','','预 览(V),编 辑(E)','','')
			}
			function ChangeUp()
			{
			 if (FolderID=='0') return;
			 location.href='JS_Main.asp?JSType='+JSType+'&FolderID='+ParentID;
			   if (JSType==0)
				  $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS 管理 >> 系统 JS&ButtonSymbol=SysJSList&LabelFolderID='+ParentID;
			   else
				 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS 管理 >> 自由 JS&ButtonSymbol=FreeJSList&LabelFolderID='+ParentID;
			 }
			function OpenFolder(FolderID)
			{
			 location.href='JS_Main.asp?JSType='+JSType+'&FolderID='+FolderID;
			   if (JSType==0)
				 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS 管理 >> 系统 JS&ButtonSymbol=SysJSList&LabelFolderID='+FolderID;
			   else
				$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS 管理 >> 自由 JS&ButtonSymbol=FreeJSList&LabelFolderID='+FolderID;
			}
			var box='';
			function CreateFolder()
			{ 
			
			  if (JSType==0)
			     box=$.dialog({title:"新建系统JS目录",content:"url:include/LabelFolder.asp?LabelType=2&FolderID="+FolderID,width:650,height:360});
			  else
			   box=$.dialog({title:"新建自由JS目录",content:"url:include/LabelFolder.asp?LabelType=3&FolderID="+FolderID,width:650,height:360});
			}
			function AddJS(TempUrl)
			{
			  if (JSType==0)
				{
				 location.href=TempUrl+'JS/AddSysJS.asp?FolderID='+FolderID+'&JSType="'+JSType+'&Action=AddNew';
				  $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS 管理 >> <font color=red>添加系统 JS</font>&ButtonSymbol=JSAdd';
				 }
			  else
				{location.href=TempUrl+'JS/AddFreeJS.asp?FolderID='+FolderID+'&Action='+Action+'&JSID='+JSID+'&JSType='+JSType
				 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS 管理 >> <font color=red>添加自由 JS</font>&ButtonSymbol=JSAdd';
				 }
			}
			function EditJS(TempUrl,ID)
			{  if (KeyWord=='')
				{   if (JSType==0)
					  {
					   location.href=TempUrl+'EditJS.asp?Page='+Page+'&JSType='+JSType+'&Action=Edit&JSID='+ID;
					   $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS管理 >> <font color=red>修改系统JS</font>&ButtonSymbol=JSEdit';
					  }
				   else
				   {
					 location.href=TempUrl+'EditJS.asp?Page='+Page+'&JSType=1&Action=Edit&JSID='+ID;
					 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS管理 >> <font color=red>修改自由JS</font>&ButtonSymbol=JSEdit';
					}
				}
			   else
				 {  if (JSType==0)
					 {
					  location.href=TempUrl+'EditJS.asp?'+SearchParam+'&Page='+Page+'&JSType='+JSType+'&Action=Edit&JSID='+ID;
					  $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS管理 >> 搜索系统JS结果 >><font color=red>修改系统JS</font>&ButtonSymbol=JSEdit';
					  }
				   else
					{
					 location.href=TempUrl+'EditJS.asp?'+SearchParam+'&Page='+Page+'&JSType=1&Action=Edit&JSID='+ID;
					 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS管理 >> 搜索自由JS结果 >> <font color=red>修改自由JS</font>&ButtonSymbol=JSEdit';
					 }
			  }
			}
			function EditFolder(ID)
			{
			 box=$.dialog({title:"编辑标签目录",content:"url:include/LabelFolder.asp?Action=EditFolder&FolderID="+ID,width:650,height:360});
			}
			function Edit(TempUrl)
			{   GetSelectStatus('FolderID','JSID');
				if (!((SelectedFile=='')&&(SelectedFolder=='')))
					{
						if (SelectedFolder!='')
						{ 
						if (TempUrl=='Folder'||TempUrl=='')
						 if (SelectedFolder.indexOf(',')==-1) 
						  {
						   EditFolder(SelectedFolder);
						 }
						else alert('一次只能够编辑一个标签目录');
					   }
					   if (SelectedFile!='')
						 {
						 if (TempUrl!='Folder'||TempUrl=='')
						 {	if (SelectedFile.indexOf(',')==-1) 
							 EditJS(TempUrl,SelectedFile);
						 else alert('一次只能够编辑一个JS');
			
						 }
						}
					}
				else 
				{
				alert('请选择要编辑的标签或目录');
				}
			}
			function Delete(TempUrl)
			{   GetSelectStatus('FolderID','JSID');
				if (!((SelectedFile=='')&&(SelectedFolder=='')))
					{  
						if (confirm('删除确认:\n\n真的要执行删除操作吗?'))
						  { if (SelectedFolder!='')
						   if (TempUrl=='Folder'||TempUrl=='')
							location='JS_Main.asp?JSAction=JSFolderDel&ID='+SelectedFolder;
						  if (SelectedFile!='')  
							if (TempUrl!='Folder'||TempUrl=='')
						location=TempUrl+'JS_Main.asp?JsAction=JSDel&Page='+Page+'&JSID='+SelectedFile;
						}	
					}
				else alert('请选择要删除的标签目录或标签');
			   SelectedFile='';
			   SelectedFolder='';
			}
			function DelFolder(ID){
		    if (confirm('删除确认:\n\n真的要执行删除JS目录操作吗?')){
			location='JS_Main.asp?JSAction=JSFolderDel&ID='+ID+'&FolderID='+FolderID;
			}
		}
		function DelJS(ID){
		    if (confirm('删除确认:\n\n真的要执行删除JS操作吗?')){
			location='JS_Main.asp?JsAction=JSDel&Page='+Page+'&JSID='+ID;
			}
		}
			
			function GetKeyDown()
			{
			if (event.ctrlKey)
			  switch  (event.keyCode)
			  {  case 90 :  Reload(); break;
				 case 78 : event.keyCode=0;event.returnValue=false; CreateFolder();break;
				 case 77 : event.keyCode=0;event.returnValue=false; AddJS('');break;
				 case 65 : SelectAllElement();break;
				 case 66 : event.keyCode=0;event.returnValue=false;ChangeUp();break;
				 case 69 : event.keyCode=0;event.returnValue=false;Edit('');break;
				 case 68 : Delete('');break;
				 case 86 : JSView();break;
				 case 70 : event.keyCode=0;event.returnValue=false;
				   if (JSType==0)
					parent.initializeSearch('SysJS')
				   else
					parent.initializeSearch('FreeJS')
			 }	
			else if (event.keyCode==46)
			Delete('');
			}
			function Reload()
			{
			 location.href='js_Main.asp?FolderID='+FolderID+'&JSType='+JSType+'&'+SearchParam
			}
			function JSView()
			{   GetSelectStatus('FolderID','JSID');
				if (SelectedFile!='')
				{
				 window.open('LabelFrame.asp?Url=JS_Main.asp&JSAction=JSView&JSID='+SelectedFile+'&PageTitle='+escape('预览JS显示效果'),'new','width=620,height=450');
				 SelectedFile='';
				 }
				else
				 alert('请选择您要预览的JS!')
			}

			</script>
			<%
			.Write "</head>"
			.Write "<body scroll=no topmargin=""0"" leftmargin=""0"" onclick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
		    .Write "<ul id='menu_top'>"
				 If KeyWord = "" Then
			.Write "<li class='parent' onclick=""AddJS('');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>添加JS</span></li>"
			.Write "<li class='parent' onclick=""CreateFolder();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>添加目录</span></li>"
			.Write "<li class='parent' onclick=""Edit('Folder');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/as.gif' border='0' align='absmiddle'>编辑目录</span></li>"
			.Write "<li class='parent' onclick=""Delete('Folder');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/del.gif' border='0' align='absmiddle'>删除目录</span></li>"

			 .Write "<li class='parent' onclick=""parent.initializeSearch("
			 If JSType = 0 Then .Write ("'系统 JS',0,'SysJS'") Else .Write ("'自由 JS',0,'FreeJS'")
			 .Write ");""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/s.gif' border='0' align='absmiddle'>搜索助理</span></li>"
			 .Write "<li class='parent' onclick=""ChangeUp();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/verify.gif' border='0' align='absmiddle'>回上一级</span></li>"
				 
				 Else
					  If JSType = 0 Then
					   .Write ("<img src='../Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('JS_Main.asp?JSType=0','Template_Left.asp','../KS.Split.asp?ButtonSymbol=SysJSList&OpStr=JS管理 >> <font color=red>系统JS管理</font>')"">系统JS首页</span>")
					Else
					   .Write ("<img src='../Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('JS_Main.asp?JSType=1','Template_Left.asp','../KS.Split.asp?ButtonSymbol=FreeJSList&OpStr=JS管理 >> <font color=red>自由JS管理</font>')"">自由JS首页</span>")
					End If
				   .Write (">>> 搜索结果: ")
					 If StartDate <> "" And EndDate <> "" Then
						.Write ("JS更新日期在 <font color=red>" & StartDate & "</font> 至 <font color=red> " & EndDate & "</font>&nbsp;&nbsp;&nbsp;&nbsp;")
					 End If
					Select Case SearchType
					 Case 0
					  .Write ("名称含有 <font color=red>" & KeyWord & "</font> 的JS")
					 Case 1
					  .Write ("描述中含有 <font color=red>" & KeyWord & "</font> 的JS")
					 Case 2
					  .Write ("文件名中含有 <font color=red>" & KeyWord & "</font> 的JS")
					 End Select
			End If
			
			.Write "    </ul>"

			.Write "<div style="" height:98%; overflow: auto; width:100%"" align=""center"">"
			.Write "<table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
			.Write "  <tr>"
			.Write "    <td  valign=""top"">"
			.Write "      <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
			.Write "        <tr>"
			.Write "          <td height=""8"" align=""center""></td>"
			.Write "        </tr>"
				   
					Dim FolderSql, Param
					 Param = " Where JsType=" & JSType
					If KeyWord <> "" Then
					   FolderSql = "SELECT ID,FolderName,Description,OrderID FROM [KS_LabelFolder] Where 1=0"
					  Select Case SearchType
						Case 0
						  Param = Param & " AND JSName like '%" & KeyWord & "%'"
						Case 1
						 Param = Param & " AND Description like '%" & KeyWord & "%'"
						Case 2
						 Param = Param & " AND JSFileName like '%" & KeyWord & "%'"
					  End Select
					  If StartDate <> "" And EndDate <> "" Then
						   Param = Param & " And (AddDate>=#" & StartDate & "# And AddDate<=#" & DateAdd("d", 1, EndDate) & "#)"
					  End If
					Else
					   Param = Param & " AND FolderID='" & FolderID & "'"
					   FolderSql = "SELECT ID,FolderName,Description,OrderID FROM [KS_LabelFolder] Where FolderType=" & JSType + 2 & " And ParentID='" & FolderID & "'"
					End If
					Param = Param & " ORDER BY OrderID"
			Set JSRS = Server.CreateObject("ADODB.recordset")
			JSRS.Open FolderSql & " UNION  Select JSID,JSName,Description,OrderID From KS_JSFile " & Param, Conn, 1, 1
			If JSRS.EOF And JSRS.BOF Then
					 Else
						totalPut = JSRS.RecordCount
								If CurrentPage < 1 Then
									CurrentPage = 1
								End If
			
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage = 1 Then
									Call showContent
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										JSRS.Move (CurrentPage - 1) * MaxPerPage
										
										Call showContent
									Else
										CurrentPage = 1
										Call showContent
									End If
								End If
							   
				End If
			 .Write "  </table>"
			 .Write "     </td>"
			 .Write "   </tr>"
			.Write "  </table>"
			.Write "  </di>"
			.Write "  </body>"
			.Write "  </html>"
			
			Set JSRS = Nothing
			Set Conn = Nothing
			End With
			End Sub
			
			   Sub showContent() 
			   %>
		 <style>
		 .labellist{}
		 .labellist td{position:relative;}
		 .labellist .m{display:none;position:absolute;right:5px;top:5px}
		 .labellist td.td{border:1px solid #fff;}
		 .labellist td.td:hover{border:1px solid #E4E4E4;background:#FBFDFF;}
		 .labellist td.td:hover .m{display:block;position:absolute;right:5px;top:5px}
		 </style>
		 <%
				Do While Not JSRS.EOF
			Response.Write "      <tr>"
			Response.Write "    <td>"
			Response.Write "    <table class=""labellist"" width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
			Response.Write "        <tr>"
					 
					  Dim T, TitleStr, JSName, ShortName, LabelTypeStr
						   For T = 1 To Row
						  If Not JSRS.EOF Then
								  JSName = JSRS(1)
								  ShortName = KS.ListTitle(Replace(Replace(JSName, "{JS_", ""), "}", ""), 24)
								  If JSType = 1 Then
									 LabelTypeStr = "自由JS"
								   Else
									 LabelTypeStr = "系统JS"
								  End If
								  TitleStr = " TITLE='名 称:" & JSName & "&#13;&#10;类 型:" & LabelTypeStr & "&#13;&#10;描 述:" & JSRS("Description") & "'"
							   Response.Write ("<td class=""td"" width=""" & CInt(100 / Row) & "%"" Style=""cursor:default"" align=""center""" & TitleStr & ">")
						  response.write "<span class=""m"">"
						   If JSRS(3) = 0 Then
						   response.write "<a href=""javascript:;"" onclick=""EditFolder('" &  JSRS(0) & "')"" style=""color:green;"" title=""修改JS目录"">E</a> <a href=""javascript:;"" onclick=""DelFolder('" & JSRS(0) & "');"" style=""color:red;""　title=""删除JS目录"">X</a>"
						   Else
						   response.write "<a href=""javascript:;"" onclick=""EditJS('','" & JSRS(0) & "')"" style=""color:green;"" title=""修改JS"">E</a> <a href=""javascript:;"" onclick=""DelJS('" & JSRS(0) & "');"" style=""color:red;"" title=""删除JS"">X</a>"
						   End If
						   response.write "</span>"
							   
							   
							  If JSRS(3) = 0 Then
							   Response.Write ("<span onmousedown=""mousedown(this);""  FolderID=""" & JSRS(0) & """ style=""POSITION:relative;"" onDblClick=""OpenFolder(this.FolderID);""> ")
							 Else
							   Response.Write ("<span onmousedown=""mousedown(this);""  JSID=""" & JSRS(0) & """ style=""POSITION:relative;""  onDblClick=""Edit('');""> ")
							 End If
							 If JSRS(3) = 0 Then
								 Response.Write ("<img src=""../Images/Folder/folder.gif""> ")
							 Else
								 Response.Write ("<img src=""../Images/Label/JS" & JSType & ".gif"">")
							 End If
						   Response.Write ("<span style=""display:block;height:16;padding:0px 0px 0px 0px;margin:1px;width:80%;cursor:default"">" & ShortName & "</span>")
						   Response.Write ("</span>")
						   Response.Write ("</td>")
						i = i + 1
						  If JSRS.EOF Or i >= MaxPerPage Then Exit For
						   JSRS.MoveNext
						 Else
						  Exit For
						 End If
					Next
					'不到7个单元格,则进行补空
					Do While T <= Row
					 Response.Write ("<td width=70>&nbsp;</td>")
					 T = T + 1
					 Loop
					  
			Response.Write "        </tr>"
			Response.Write "        <tr><td colspan=" & Row & " height=10></td></tr>"
			Response.Write "      </table></td>"
			Response.Write "  </tr>"
			 
					  If i >= MaxPerPage Then Exit Do
					  If JSRS.EOF Then Exit Do
					Loop
					  JSRS.Close
						 Conn.Close
			
			Response.Write "        <td   align=""right"">"
				   
					 Call KS.ShowPageParamter(totalPut, MaxPerPage, "JS_Main.asp", True, "个", CurrentPage, "JSType=" & JSType & "&" & SearchParam)
			Response.Write ("</td>")
			Response.Write "</tr>"
		End Sub
		
		'删除JS
		Sub JSDel()
		 Dim K, JSID, Page,RS,ArticleRS,JSType, CurrPath, JSFileName, JSDir, FolderID
		Set RS=Server.CreateObject("ADODB.Recordset")
		Set ArticleRS=Server.CreateObject("ADODB.Recordset")
		Page = Trim(KS.G("Page"))
		JSID = Split(KS.G("JSID"), ",") '获得要删除标签的ID集合
		For K = LBound(JSID) To UBound(JSID)
		  RS.Open "SELECT * FROM [KS_JSFile] WHERE JSID='" & JSID(K) & "'", Conn, 1, 3
		  If Not RS.EOF Then
			JSType = RS("JSType")
			FolderID = RS("FolderID")
			  '删除物理JS文件
			  JSFileName = Trim(RS("JSFileName"))
			  JSDir = Trim(KS.Setting(93))
			  If Left(JSDir, 1) = "/" Or Left(JSDir, 1) = "\" Then JSDir = Right(JSDir, Len(JSDir) - 1)
			  CurrPath = KS.Setting(3) & JSDir
			  Call KS.DeleteFile(CurrPath & JSFileName)
			  '从文章中删除此JSID
			  ArticleRS.Open "Select  JSID From KS_Article Where JSID like '%" & JSID(K) & "%'", Conn, 1, 3
			  If Not ArticleRS.EOF Then
				 While Not ArticleRS.EOF
					ArticleRS(0) = Replace(ArticleRS(0), JSID(K) & ",", "")
					ArticleRS.Update
					ArticleRS.MoveNext
				 Wend
			  End If
		  End If
		 RS.Delete:RS.Close:ArticleRS.Close
		Next
		Set RS = Nothing:Set ArticleRS = Nothing
		Response.Redirect "JS_Main.asp?Page=" & Page & "&JSType=" & JSType & "&FolderID=" & FolderID
		End Sub
		
		'删除JS目录
		Sub JSFolderDel()
		   Dim RS,K, ID, ParentID, FolderSql,LabelFolderID,LabelType
		   Set RS=Server.CreateObject("ADODB.Recordset")
		   ID = Split(Request("ID"), ",")     '获得要删除目录的ID集合
			For K = LBound(ID) To UBound(ID)
			  FolderSql = "select ID,ParentID,FolderType from [KS_LabelFolder] where ID='" & ID(K) & "'"
			  RS.Open FolderSql, Conn, 1, 1
			  If Not RS.EOF Then
				LabelFolderID = Trim(RS(0))
				ParentID = Trim(RS(1))
				LabelType = RS(2)
						  Dim RSJS,JSDir
						  Set RSJS=Server.CreateObject("ADODB.Recordset")
						  '删除JS物理文件
						  RSJS.Open "Select JSFileName From KS_JSFile Where FolderID='" & LabelFolderID & "'", Conn, 1, 1
								 JSDir = Trim(KS.Setting(93))
								If Left(JSDir, 1) = "/" Or Left(JSDir, 1) = "\" Then JSDir = Right(JSDir, Len(JSDir) - 1)
						  Do While Not RSJS.EOF
								Call KS.DeleteFile(KS.Setting(3) & JSDir & RSJS(0))
								RSJS.MoveNext
						  Loop
						  RSJS.Close
						  Set RSJS = Nothing
						  Conn.Execute ("DELETE  FROM KS_JSFILE WHERE FolderID='" & LabelFolderID & "'")
						  Conn.Execute ("DELETE  FROM KS_LabelFolder WHERE ID='" & LabelFolderID & "' OR TS like '%" & LabelFolderID & "%'")
			   End If
			  RS.Close
			Next
		 Set RS = Nothing
			Response.Write "<script>location.href='JS_Main.asp?JSType=" & (LabelType - 2) & "&Folderid=" & ParentID & "'</script>"
		End Sub
		
		'预览JS
		Sub JSView()
			Dim JSObj,JSID, JSdir,JSUrlStr
			JSID=Trim(Request.QueryString("JSID"))
			JSDir = KS.Setting(93)
			If Left(JSDir, 1) = "/" Or Left(JSDir, 1) = "\" Then JSDir = Right(JSDir, Len(JSDir) - 1)
			Set JSObj=Server.CreateObject("Adodb.Recordset")
			JSObj.OPEN "Select JSConfig,JSType,JSFileName From KS_JSFile Where JSID='" & JSID & "'",Conn,1,1
			IF JSObj.EOf AND JSObj.BOF THEN
			  Response.Write("参数传递出错!")
			  JSObj.Close
			  Set JSObj=Nothing
			  Response.End
			ELSE
			  IF (trim(Split(JSObj("JSConfig"),",")(0))="GetExtJS" Or JSObj("JSType")=0) or (Request.QueryString("CanView")="1") Then
			  JSUrlStr="<script language=""javascript"" src=""" & KS.GetDomain & JSDir & Trim(JSObj("JSFileName")) & """></script>"
			  Else
				JSObj.Close:Set JSObj=Nothing
				Response.Redirect "JSFreeView.asp?JSID=" &JSID
			  End IF
			END IF
			JSObj.close:Set JSObj=Nothing
			%>
			<html>
			<head>
			<meta http-equiv="Expires" CONTENT="0">        
			<meta http-equiv="Cache-Control" CONTENT="no-cache">        
			<meta http-equiv="Pragma" CONTENT="no-cache">      
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
			<link href="Admin_Style.CSS" rel="stylesheet">
			<title>JS预览</title>
			<script language="JavaScript" src="Common.js"></script>
			</head>
			<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  bgcolor="#F1FAFA">
			<br>
			<table width="100%" height="70%" border="0" cellpadding="0" cellspacing="0">
			  <tr> 
				<td align="center"  valign="top"><table width="90%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
					<tr>
					  <td align="center" valign="top"><%=JSUrlStr%></td>
					</tr>
				  </table></td>
			  </tr>
			</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			  <tr> 
				<td height="25"><strong> 　　　　　　　说明:</strong></td>
			  </tr>
			  <tr>
				<td height="25">　　　　　　　　　　１．如果JS中有设置样式,那么这里的预览效果可能会与实际有点差距</td>
			  </tr>
			  <tr> 
				<td height="25">　　　　　　　　　　２．如果看不到效果，请单击刷新按钮 <input type="button" value="刷新" onClick="window.location.reload()"><input type="button" value="关闭" onClick="window.parent.close()">
				<%if Request.QueryString("CanView")="1" then
				  Response.Write("<INPUT TYPE=BUTTON value=""返回"" onclick=""history.back();"">")
				  End IF
				  %></td>
			  </tr>
			</table>
			</body>
			</html>
      <%
		End Sub
End Class
%> 
