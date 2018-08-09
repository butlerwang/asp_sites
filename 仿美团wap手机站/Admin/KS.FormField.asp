<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Field
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Field
        Private KS,KSCls,Action,ItemID,Page,ItemName,TableName
		Private I, totalPut, CurrentPage, FieldSql, RS,MaxPerPage
		Private FieldName,ID,Contact, Title, Tips, FieldType, DefaultValue, MustFillTF, ShowOnForm,ShowOnManage, ShowOnUserForm,Options,OrderID,FolderID,MaxFileSize,Width,Height,AllowFileExt,Step,ParentFieldName,ShowUnit,UnitOptions,MaxLength

		Private Sub Class_Initialize()
		  MaxPerPage = 18
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
		With Response
		   If Not KS.ReturnPowerResult(0, "KSMS10006") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
		   End If
		.Write "<html>"
		.Write "<head>"
		.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
		.Write "<title>字段管理</title>"
		.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		.Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		.Write "<script language='JavaScript' src='../KS_Inc/jquery.js'></script>"
             Action=KS.G("Action")
			 ItemID=KS.ChkClng(KS.G("ItemID"))
			 Page=KS.G("Page")
			 Select Case Action
			  Case "Add"  Call FieldManage("Add")
			  Case "Edit" Call FieldManage("Edit")
			  Case "Del"  Call FieldDel()
			  Case "order" Call FieldOrder()
			  Case "AddSave" Call DoSave()
			  Case "EditSave"  Call FieldEditSave()
			  Case Else Call FieldList()
			 End Select
			.Write "</body>"
			.Write "</html>"
		 End With
		End Sub
		
		Sub FieldList()
		 On Error Resume Next
		If Not IsEmpty(KS.G("page")) Then
			  CurrentPage = KS.G("page")
		Else
			  CurrentPage = 1
		End If
		With Response
		.Write "<script language='JavaScript'>"
		.Write "var Page='" & CurrentPage & "';"
		.Write "var ItemName='" & ItemName & "';"
		.Write "var ItemID=" & ItemID & ";"
		.Write "</script>"
		.Write "<script language='JavaScript' src='../KS_Inc/jquery.js'></script>"
		.Write "<script language='JavaScript' src='Include/ContextMenu1.js'></script>"
		.Write "<script language='JavaScript' src='Include/SelectElement.js'></script>"
		%>
		 <script language="javascript">
		 var DocElementArrInitialFlag=false;
		var DocElementArr = new Array();
		var DocMenuArr=new Array();
		var SelectedFile='',SelectedFolder='';
		$(document).ready(function()
		{     if (DocElementArrInitialFlag) return;
			InitialDocElementArr('FolderID','FieldID');
			InitialContextMenu();
			DocElementArrInitialFlag=true;
		});
		function InitialContextMenu()
		{	DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.FieldAdd();",'添 加(N)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'全 选(A)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.FieldControl(1);",'编 辑(E)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.FieldControl(2);",'删 除(D)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.location.reload();",'刷 新(Z)','disabled');
		}
		function DocDisabledContextMenu()
		{
			DisabledContextMenu('FolderID','FieldID','编 辑(E),删 除(D)','编 辑(E)','','','','')
		}
		function FieldAdd()
		{
		   location.href='KS.FormField.asp?ItemID='+ItemID+'&Action=Add';
		   window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=自定义表单 >> 表单项管理 >> <font color=red>新增'+ItemName+'表单项</font>&ButtonSymbol=Go';
		}
		function EditField(id)
		{
		  location="KS.FormField.asp?ItemID="+ItemID+"&Page="+Page+"&Action=Edit&ID="+id;
		  window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=自定义表单 >> 表单项管理 >> <font color=red>编辑'+ItemName+'表单项</font>&ButtonSymbol=GoSave';
		}
		function DelField(id)
		{
		if (confirm('真的要删除该表单项吗?'))
		 location="KS.FormField.asp?ItemID="+ItemID+"&Action=Del&Page="+Page+"&id="+id;
		  SelectedFile='';
		}
		function FieldControl(op)
		{   var alertmsg='';
			GetSelectStatus('FolderID','FieldID');
			if (SelectedFile!='')
			 {
			   if (op==1)
				{
				if (SelectedFile.indexOf(',')==-1) 
					EditField(SelectedFile)
				  else alert('一次只能编辑一个表单项!')	
				SelectedFile='';
				}	
			  else if (op==2)    
			   DelField(SelectedFile);
			 }
			else 
			 {
			 if (op==1)
			  alertmsg="编辑";
			 else if(op==2)
			  alertmsg="删除"; 
			 else
			  {
			  alertmsg="操作" 
			  }
			 alert('请选择要'+alertmsg+'的表单项');
			  }
		}
		function GetKeyDown()
		{ 
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 : location.reload(); break;
			 case 65 : SelectAllElement();break;
			 case 78 : event.keyCode=0;event.returnValue=false;FieldAdd();break;
			 case 69 : event.keyCode=0;event.returnValue=false;FieldControl(1);break;
			 case 68 : FieldControl(2);break;
		   }	
		else	
		 if (event.keyCode==46)FieldControl(2);
		}
		 </script>
		<%
		.Write "</head>"
		.Write "<body scroll=no topmargin='0' leftmargin='0' onclick='SelectElement();' onkeydown='GetKeyDown();' onselectstart='return false;'>"
		.Write "<ul id='menu_top'>"
		.Write "<li class='parent' onclick='FieldAdd();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>新增表单项</span></li>"
		.Write "<li class='parent' onclick='FieldControl(1);'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>修改表单项</span></li>"
		.Write "<li class='parent' onclick='FieldControl(2)'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>删除表单项</span></li>"
		.Write "<li class='parent' onclick='location.href=""KS.Form.asp"";'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>回上一级</span></li>"
		.Write "</ul>"
        
		.Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
		.Write "<form action='KS.FormField.asp?action=order&ItemID=" & ItemID & "&page=" & Page & "' name='form1' method='post'>"
		.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		.Write "        <tr class='sort'>"
		.Write "         <td width='80' align='center'>排序</td>"
		.Write "         <td align='center'>表单项名称</td>"	
		.Write "         <td align='center'>表单项类型</td>"
		.Write "         <td align='center'>默认值</td>"
		.Write "         <td align='center'>是否显示</td>"
		.Write "         <td align='center'>管理显示</td>"
		.Write "         <td align='center'>↓管理操作</td>"
		.Write "        </tr>"
			 Set RS = Server.CreateObject("ADODB.RecordSet")
				   FieldSql = "SELECT * FROM KS_FormField Where ItemID=" & ItemID & " order by OrderID desc"
				   RS.Open FieldSql, conn, 1, 1
				 If RS.EOF And RS.BOF Then
				 Else
					totalPut = RS.RecordCount
		
							If CurrentPage < 1 Then	CurrentPage = 1
		
							If CurrentPage <> 1 Then
								If (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
								Else
									CurrentPage = 1
								End If
							End If
							Call showContent

			End If
		 .Write " <tr>"
		 .Write "   <td colspan='3'>&nbsp;&nbsp;<input type='button' value='表单预览' onclick=""SelectObjItem1(this,'自定义表单 >> <font color=red>表单预览</font>','gosave','KS.Form.asp?ItemID=" & ItemID & "&action=view');"" class='button'>&nbsp;<input type='submit' class='button' value='批量保存设置'> <font color=blue>越小排在越前面</font></td></form>"
		 .Write "   <td height='35' colspan='4' align='right'>"
		 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		.Write "    </td>"
		.Write " </tr>"
		.Write "</table>"
		.Write "</div>"
		End With
		End Sub
		Sub showContent()
		With Response
		Do While Not RS.EOF
		.Write "  <tr>"
		 .Write "<td class='splittd'>&nbsp;&nbsp;<input type='text' name='OrderID' style='width:45px;text-align:center' value='" & RS("OrderID") &"'><input type='hidden' name='FieldID' value='" & RS("FieldID") & "'></td>"
		.Write "    <td class='splittd'>"
		.Write "    <span FieldID='" & RS("FieldID") & "' onDblClick=""EditField(this.FieldID)"">"
		 .Write "     <img src='Images/Field.gif' align='absmiddle'>"
		 .Write "     <span style='cursor:default;'>" & RS("Title") & "</span>"
		 .Write "   </span>"
		 .Write "   </td>"
		 .Write "   <td align='center' class='splittd'>"
		 Select Case RS("FieldType")
				  Case 1:.Write "单行文本(text)"
				  Case 2:.Write "文本(不支持HTML)"
				  Case 10:.Write "多行文本(支持HTML)"
				  Case 3:.Write "下拉列表(select)"
				  Case 4:.Write "数字(text)"
				  Case 5:.Write "日期(text)"
				  Case 6:.Write "单选框(radio)"
				  Case 7:.Write "复选框(checkbox)"
				  Case 8:.Write "电子邮箱(text)"
				  Case 9:.Write "文件(text)"
				  Case 11:.Write "联动菜单(text)"
		 End Select
		 .Write "</td>"
		 .Write "   <td align='center' class='splittd'>" & RS("DefaultValue") & "&nbsp;</td>"
		 .Write "   <td align='center' class='splittd'>" 
		  
		  If RS("ShowOnForm")="1" Then
		   .Write "<input type='checkbox' name='ShowOnForm" & RS("FieldID") &"' value='1' checked>"
		  Else
		   .Write "<input type='checkbox' name='ShowOnForm" & RS("FieldID") &"' value='1'>"
		  End If
		  
		 .Write " </td>"
		 .Write "   <td align='center' class='splittd'>" 
		  If RS("ShowOnManage")="1" Then
		   .Write "<input type='checkbox' name='showonmanage" & RS("FieldID") &"' value='1' checked>"
		  Else
		   .Write "<input type='checkbox' name='showonmanage" & RS("FieldID") &"' value='1'>"
		  End If
		 .Write " </td>"
		 .Write " <td align='center' class='splittd'><a href='javascript:EditField(" & RS("FieldID") &");'>修改</a> | "
		 .Write "<a href='javascript:DelField(" & RS("FieldID") &");'>删除</a>"
		 .Write " </td></tr>"
			I = I + 1
			If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		   Loop
		  RS.Close
         End With
		 End Sub
		 
		 Sub FieldManage(OpType)
		 With Response
		  Dim RS, FieldSql,OpAction,OpTempStr,FormName,PostByStep,StepNum,Step,K
		 ID = KS.G("ID")
		 Set RS=Server.CreateObject("ADODB.Recordset")
		 RS.Open "Select top 1 FormName,PostByStep,StepNum From KS_Form Where ID=" & ItemID,conn,1,1
		 If RS.EOF And RS.Bof Then
		  Response.Write "<script>alert('error!');history.back();</script>"
		  Exit Sub
		 Else
		   FormName=RS(0):PostByStep=RS(1):StepNum=RS(2)
		 End If
		 RS.Close
		.Write "<html>"
		.Write "<head>"
		.Write "<meta http-equiv='Content-Type' content='text/html; chaRSet=utf-8'>"
		.Write "<title>字段管理</title>"
		.Write "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		.Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		.Write "</head>"
		.Write "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
		
		 If Optype = "Edit" Then
		     OpAction="EditSave":OpTempStr="编辑"
			 FieldSql = "Select top 1 * From [KS_FormField] Where FieldID=" & ID
			 RS.Open FieldSql, conn, 1, 1
			 If Not RS.EOF Then
				 ItemID    = RS("ItemID")
				 Title     = Trim(RS("Title"))
				 FieldName = RS("FieldName")
				 Tips      = Trim(RS("Tips"))
				 FieldType = Trim(RS("FieldType"))
				 DefaultValue = Trim(RS("DefaultValue"))
				 MustFillTF   = RS("MustFillTF")
				 ShowOnForm   = RS("ShowOnForm")
				 ShowOnManage = RS("ShowOnManage")
				 Options      = Trim(RS("Options"))
				 OrderID      = RS("OrderID")
				 Width        = RS("Width")
				 Height       = RS("Height")
				 AllowFileExt = RS("AllowFileExt")
				 MaxFileSize  = RS("MaxFileSize")
				 Step         = RS("Step")
				 ParentFieldName=RS("ParentFieldName")
				 ShowUnit       = RS("ShowUnit")
				 UnitOptions    = RS("UnitOptions")
				 MaxLength      = RS("MaxLength")
			 End If
	  Else
	     FieldName="KS_":FieldType=1:MaxLength=0:MustFillTF=0:ShowOnForm=1:ShowOnManage=1:ShowOnUserForm=1:Width="200":Height="100":AllowFileExt="jpg|gif|doc":MaxFileSize=1024 : ShowUnit=0:OrderID=KS.ChkClng(Conn.Execute("Select max(orderid) From KS_FormField Where ItemID=" & itemID)(0))+1
		 OpAction="AddSave":OpTempStr="添加"
	  End If
		 
		.Write "<div class='topdashed sort'>" & OpTempStr &"自定义表单项</div>"
		.Write "<br>"
        .Write " <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1""  class='ctable'>" & vbCrLf
		.Write "  <form  action='KS.FormField.asp?Action=" & OpAction &"' method='post' name='OrigArticlerm' onsubmit='return(CheckForm())'>"
		
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' width='180' align='right' class='clefttitle'><strong>表单项目：</strong></td>"
		.Write "      <td nowrap> &nbsp;&nbsp;<font color=#ff0000>" & FormName & "</font><input type='hidden' value='" & ItemID & "' name='ItemID'></td>"
		.Write "    </tr>"


		
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='45' align='right' class='clefttitle'><strong>表单项别名：</strong></td>"
		.Write "      <td height='45' nowrap>&nbsp;&nbsp;<input name='Title' type='text' size='30' class='textbox' value='" & Title & "'> *<font color=red>如，你的姓名等</font></td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='45' align='right' class='clefttitle'><strong>字段名称：</strong></td>"
		If Optype = "Edit" Then
		.Write "      <td height='45' nowrap>&nbsp;&nbsp;<input disabled name='FieldName' type='text'size='30' class='textbox' value='" & FieldName & "'> <font color=red>*必须以KS_开头，字段名由字母、数字、下划线组成,且不可修改</font> </td>"
		Else
		.Write "      <td height='45' nowrap>&nbsp;&nbsp;<input name='FieldName' type='text'size='30' class='textbox' value='" & FieldName & "'> <font color=red>*必须以KS_开头，字段名由字母、数字、下划线组成,且不可修改</font> </td>"
		End If
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right'' class='clefttitle'><strong>附加提示：</strong><br><font color=blue>在名称旁的提示信息</font></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<textarea name='Tips'  id='Tips' class='textbox' cols='30' rows='3' style='height:50px'>" & Tips & "</textarea></td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>字段类型：</strong></td>"
		.Write "      <td nowrap>&nbsp;"
		If Optype = "Edit" Then
		.Write "      <input type='hidden' value=" & FieldType & " name='FieldType'><select name=""FieldTypes"" disabled>"
		else
		.Write "      <select name=""FieldType"" onchange=""Setdisplay(this.value)"">"
		end if
		.Write " <option value=""1"""
		If FieldType=1 Then .Write " Selected"
		.Write ">单行文本(text)</option>"
     	.Write " <option value=""2"""
		If FieldType=2 Then .Write " Selected"
		.Write ">多行文本(不支持HTML)</option>"
     	.Write " <option value=""10"""
		If FieldType=10 Then .Write " Selected"
		.Write ">多行文本(支持HTML)</option>"
		.Write " <option value=""3"""
		If FieldType=3 Then .Write " Selected"
		.Write ">下拉列表(select)</option>"
		.Write " <option value=""11"""
		If FieldType=11 Then .Write " selected"
		.Write " style='color:blue'>联动下拉列表</option>"
        .Write " <option value=""4"""
		If FieldType=4 Then .Write " Selected"
		.Write ">数字(text)</option>"
		.Write " <option value=""5"""
		If FieldType=5 Then .Write " Selected"
		.Write ">日期(text)</option>"
		.Write " <option value=""6"""
		If FieldType=6 Then .Write " Selected"
		.Write ">单选框(radio)</option>"
		.Write " <option value=""7"""
		If FieldType=7 Then .Write " Selected"
		.Write ">复选框(checkbox)</option>"
		.Write " <option value=""8"""
		If FieldType=8 Then .Write " Selected"
		.Write ">电子邮箱(text)</option>"
		.Write " <option value=""9"""
		If FieldType=9 Then .Write " Selected"
		.Write ">文件(text)</option>"
		
		.Write " </select>"
		.Write " </td>"
		.Write "    </tr>"
		.Write "<tbody id=""extarea"">"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>允许上传的扩展名：</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='AllowFileExt' type='text' id='AllowFileExt' class='textbox' value='" & AllowFileExt & "' size='40'>&nbsp;<span style='color:#ff0000'>多个扩展展名，请用逗号“|”隔开</span>"
		.Write "       </td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>允许上传的文件大小：</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='MaxFileSize' type='text' id='MaxFileSize' class='textbox' value='" & MaxFileSize & "' size='8' style='width:50px'>&nbsp;KB <span style='color:#ff0000'>*</span>  <span style='color:blue'>提示：1 KB = 1024 Byte，1 MB = 1024 KB<span>  "
		.Write "       </td>"
		.Write "    </tr>"
		.Write " </tbody>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>默认值：</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='DefaultValue' type='text' id='DefaultValue' class='textbox' value='" & DefaultValue & "' size='40'>&nbsp;<span id='darea' style='color:#ff0000'>多个默认选项，请用逗号“,”隔开</span>"
		.Write "       </td>"
		.Write "    </tr>"
		
.Write "    <tr id=""ldArea"" style='display:none' class='tdbg'>"
		.Write "      <td align='right' class='clefttitle'><strong>所属父级字段：</strong><br><font color=blue>不选择表示一级联动字段<br/>不能指定为下级联动字段</font></td>"
		.Write "      <td>&nbsp;&nbsp;"
		  Dim PRS
		  If KS.ChkClng(ID)<>0 Then
		  Set PRS=Conn.Execute("Select FieldName,Title From KS_FormField Where  ItemID=" & ItemID&" and FieldType=11 And FieldID<>" & ID & " Order BY FieldID")
		  .Write "<select name='ParentFieldName' disabled>"
		  Else
		  Set PRS=Conn.Execute("Select FieldName,Title From KS_FormField Where ItemID=" & ItemID&" and FieldType=11 Order BY FieldID")
		  .Write "<select name='ParentFieldName'>"
		  End If
		  .Write "<option value='0'>--作为一级联动--</option>"
		  Do While Not PRS.Eof
		      If PRS(0)=ParentFieldName Then
		      .Write "<option value='" & PRS(0) & "' selected>" & Prs(1) & "(" & PRS(0) & ")</option>"
			  Else
		      .Write "<option value='" & PRS(0) & "'>" & Prs(1) & "(" & PRS(0) & ")</option>"
			  End If
		  PRS.MoveNext
		  Loop
		  PRS.Close: Set PRS=Nothing
		.Write "      </select> <font color=red>说明：一旦设定不能修改</font></td>"
		.Write "    </tr>"
		
		.Write "    <tr id=""OptionsArea"" style=""display:none"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td align='right' ' class='clefttitle'><strong>列表选项：</strong><br><font color=blue>每一行为一个列表选项</font>如果值和显示项不同可以用<font color=red>|</font>隔开<br/>正确格式如：男 或 0<font color=red>|</font>男</td>"
		.Write "      <td height='45' nowrap>&nbsp;&nbsp;<textarea name='Options' cols='50' rows='6' id='Options' style='height:60px' class='textbox'>" & Options & "</textarea>"
		.Write "      </td>"
		.Write "    </tr>"
		
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>是否显示下拉单位：</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;"
		 If Optype = "Edit" Then
		    If ShowUnit="1" Then .Write "是" Else .Write "否"
			.Write "<input type='hidden' name='ShowUnit' value='1'>"
		 Else
			.Write  "<input onclick=""$('#unitArea').show()"" name='ShowUnit' type='radio' id='ShowUnit' value='1'"
			If ShowUnit="1" Then .Write " Checked"
			.Write ">是"
			.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input onclick=""$('#unitArea').hide()"" name='ShowUnit' type='radio' id='ShowUnit' value='0'"
			If ShowUnit="0" Then .Write " Checked"
			.WRite ">否"
		 End If
		 .Write "&nbsp;&nbsp;<font color=red>说明：一旦设定不能修改</font>"
		.Write "       </td>"
		.Write "    </tr>"
		If ShowUnit="1" Then
		.Write "    <tr class=""tdbg"" id=""unitArea"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
	   else
		.Write "    <tr class=""tdbg"" id=""unitArea"" style=""display:none"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
	   end if
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>下拉单位选项：</strong><br/><font color=blue>每一行为一个列表选项<br/>如:件 个等</font></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<textarea name='UnitOptions' style='height:70px' cols='20' rows='6' id='UnitOptions' class='textbox'>" & UnitOptions & "</textarea> "
		.Write "       </td>"
		.Write "    </tr>"		
		
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>是否必填：</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='MustFillTF' type='radio' id='MustFillTF' value='1'"
		If MustFillTF=1 Then .Write " Checked"
		.Write ">是"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='MustFillTF' type='radio' id='MustFillTF' value='0'"
		If MustFillTF=0 Then .Write " Checked"
		.WRite ">否"
		.Write "       </td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>是否启用：</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='ShowOnForm' type='radio' id='ShowOnForm' value='1'"
		If ShowOnForm=1 Then .Write " Checked"
		.Write ">是"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnForm' type='radio' id='ShowOnForm' value='0'"
		If ShowOnForm=0 Then .Write " Checked"
		.WRite ">否"
		.Write "       </td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>列表管理显示该字段：</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='ShowOnManage' type='radio' id='ShowOnManage' value='1'"
		If ShowOnManage=1 Then .Write " Checked"
		.Write ">是"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnManage' type='radio' id='ShowOnManage' value='0'"
		If ShowOnManage=0 Then .Write " Checked"
		.WRite ">否"
		.Write "       </td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>大小设置：</strong> </td>"
		.Write "      <td nowrap>&nbsp;&nbsp;宽度<input name='Width' type='text' site='10' class='textbox' style='width:40px' value='" & Width & "'>px &nbsp;高度<input name='Height' type='text' site='10' class='textbox' style='width:40px' value='" & Height & "'>px</font>&nbsp;&nbsp;长度<input class='textbox' type='text' name='MaxLength' id='MaxLength' value='" & MaxLength &"' style='width:40px'>个字符,不限制请输入0"
		.Write "       </td>"
		.Write "    </tr>"				
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>排序序号：</strong><br><font color=blue>序号越小，排在越前面</font></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='OrderID' type='text' site='35' class='textbox' id='OrderID' value='" & OrderID & "'>"
		.Write "       </td>"
		.Write "    </tr>"
		If PostByStep="1" and StepNum>1 Then
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>分步提交设置：</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;该表单项在第<select name='step'>"
		  For K=1 To StepNum
		   If K=Step Then
		   .Write "<option selected>" & K & "</option>"
		   Else
		   .Write "<option>" & K & "</option>"
		   End If
		  Next
		.Write "       </select>步出现</td>"
		.Write "    </tr>"
		End If
		
		.Write "   <input type='hidden' value='" & ID & "' name='id'>"
		.Write "   <input type='hidden' value='" & Page & "' name='page'>"
		.Write "  </form>"
		.Write "</table>"
		
		
		 
		.Write "<Script Language='javascript'>"
		.Write "Setdisplay(" & FieldType & ");"
		.Write "function Setdisplay(s)"
		.Write  "{if (s==3||s==6||s==7||s==11){ document.all.OptionsArea.style.display='';} else document.all.OptionsArea.style.display='none';if (s==7)document.getElementById('darea').style.display='';else document.getElementById('darea').style.display='none';if(s==9)document.getElementById('extarea').style.display='';else document.getElementById('extarea').style.display='none'; if(s==11) document.getElementById('ldArea').style.display=''; else document.getElementById('ldArea').style.display='none';}"
		.Write "function CheckForm()"
		.Write "{ var form=document.OrigArticlerm;"
		.Write "   if (form.Title.value=='')"
		.Write "    {"
		.Write "     alert('请输入表单项名称!');"
		.Write "     form.Title.focus();"
		.Write "     return false;"
		.Write "    }"
		.Write "    form.submit();"
		.Write "    return true;"
		.Write "}"
		.Write "</Script>"
		End With
		End Sub
		 
		 Sub DoSave()
		 Dim RS,ColumnType,ItemID
		 ItemID      = KS.ChkClng(KS.G("ItemID"))
		 Title       = KS.G("Title")
		 FieldName   = KS.G("FieldName")
		 Tips        = Request.Form("Tips")
		 FieldType   = KS.G("FieldType")
		 DefaultValue= KS.G("DefaultValue")
		 MustFillTF  = KS.G("MustFillTF")
		 FieldType   = KS.G("FieldType")
		 ShowOnForm  = KS.G("ShowOnForm")
		 Options     = KS.G("Options")
		 FieldType   = KS.G("FieldType")
		 OrderID     = KS.G("OrderID")
		 Width       = KS.ChkClng(KS.G("Width"))
		 Height      = KS.ChkClng(KS.G("Height"))
		 AllowFileExt= KS.G("AllowFileExt")
		 MaxFileSize = KS.ChkClng(KS.G("MaxFileSize"))
		 Step        = KS.ChkClng(KS.G("Step"))
		 ParentFieldName = KS.G("ParentFieldName")
		 If KS.IsNul(ParentFieldName) Then ParentFieldName="0"
		 ShowUnit    = KS.ChkClng(KS.G("ShowUnit"))
		 UnitOptions = KS.G("UnitOptions")
		 ShowOnManage= KS.ChkClng(KS.G("ShowOnManage"))
		 MaxLength   = KS.ChkClng(KS.G("MaxLength"))
		 
		 If FieldName = "" Then Call KS.AlertHistory("请输入字段名称!", -1): Exit Sub
		 If Len(FieldName)<=3 Then Call KS.AlertHistory("字段名称长度必须大于3!", -1): Exit Sub
		 If Ucase(Left(FieldName,3))<>"KS_" Then Call KS.AlertHistory("字段名称格式有误，必须以""KS_开头""!", -1): Exit Sub
		 If Title="" Then Call KS.AlertHistory("字段标题必须输入!", -1): Exit Sub
		 If FieldType=4 And Not Isnumeric(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub
		 If FieldType=5 And Not IsDate(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub
		 if FieldType=8 And Not KS.IsValidEmail(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认格式不正确，请输入正确的Email!",-1):Exit Sub
	     on error resume next
		 Conn.Begintrans
		 Set RS = Server.CreateObject("ADODB.RECORDSET")
		 FieldSql = "Select top 1 * From [KS_FormField] Where FieldName='" & FieldName & "' And ItemID=" & ItemID
		 RS.Open FieldSql, conn, 3, 3
		 If RS.EOF And RS.BOF Then
		  RS.AddNew
		  RS("ItemID") = KS.G("ItemID")
		  RS("Title") = Title
		  RS("FieldName") = FieldName
		  RS("Tips") = Tips
		  RS("FieldType") = FieldType
		  RS("DefaultValue") = DefaultValue
		  RS("MustFillTF") = MustFillTF
		  RS("FieldType") = FieldType
		  RS("ShowOnForm") = ShowOnForm
		  RS("Options") = Options
		  RS("OrderID")=OrderID
		  RS("Width")  = Width
		  RS("Height") = Height
		  RS("AllowFileExt")= AllowFileExt
		  RS("MaxFileSize") = MaxFileSize
		  RS("Step") = Step
		  RS("ParentFieldName")=ParentFieldName
		  RS("ShowUnit")=ShowUnit
		  RS("UnitOptions")=UnitOptions
		  RS("ShowOnManage")=ShowOnManage
		  RS("MaxLength")   =MaxLength
		  RS.Update
		  
		  Select Case FieldType
		   Case 1,3,6,7,8,9,11
		     If MaxLength=0 Then
		     ColumnType="nvarchar(255)"
			 Else
		     ColumnType="nvarchar(" &MaxLength&")"
			 End If
		   Case 2,10
		     ColumnType="ntext"
		   Case 5
		     ColumnType="datetime"
		   Case 4
		     ColumnType="int"
		   Case else
		     Exit Sub
		 End Select
		 Dim TableName:TableName=Conn.Execute("Select TableName From KS_Form  Where ID=" & ItemID)(0)
		  Conn.Execute("Alter Table "&TableName&" Add "&FieldName&" "&ColumnType&"")
		  If ShowUnit=1 Then  '增加单位字段
		  Conn.Execute("Alter Table "&TableName&" Add "&FieldName&"_Unit nvarchar(200)")
		  End If
		 If err<>0 then
			Conn.RollBackTrans
			Call KS.AlertHistory("出错！出错描述：" & replace(err.description,"'","\'"),-1):response.end
		 Else
			Conn.CommitTrans
		 End IF
		 Response.Write ("<Script> if (confirm('表单项增加成功,继续添加吗?')) { location.href='KS.FormField.asp?ItemID=" & ItemID& "&Action=Add';} else{location.href='KS.FormField.asp?ItemID=" & ItemID&"&Page='"&Page &";$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=自定义表单管理 >> <font color=#ff0000>表单项管理</font>&ButtonSymbol=Disabled';}</script>")
		 Else
		   Call KS.AlertHistory("数据库中已存在该字段名称!", -1)
		   Exit Sub
		 End If
		 RS.Close
		 End Sub
		 
		 Sub FieldEditSave()
		 With Response
		 ID = KS.G("ID")
		 Title = KS.G("Title")
		 Tips = Request.Form("Tips")
		 DefaultValue = KS.G("DefaultValue")
		 MustFillTF = KS.G("MustFillTF")
		 FieldType = KS.G("FieldType")
		 ShowOnForm = KS.G("ShowOnForm")
		 Options = KS.G("Options")
		 FieldType =KS.G("FieldType")
		 OrderID   =KS.G("OrderID")
		 Width     = KS.ChkClng(KS.G("Width"))
		 Height    = KS.ChkClng(KS.G("Height"))
		 AllowFileExt = KS.G("AllowFileExt")
		 MaxFileSize  = KS.ChkClng(KS.G("MaxFileSize"))
		 Step         = KS.ChkClng(KS.G("Step"))
		 UnitOptions=KS.G("UnitOptions")
		 ShowOnManage=KS.ChkClng(KS.G("ShowOnManage"))
		 MaxLength   =KS.ChkClng(KS.G("MaxLength"))

		 '修改字段长度
		 if (FieldType=1 or FieldType=3 or FieldType=6 or FieldType=7 or FieldType=8 or FieldType=9 or FieldType=11) then
		     Dim TableName:TableName=Conn.Execute("Select TableName From KS_Form  Where ID=" & KS.ChkClng(KS.G("ItemID")))(0)
		     If TableName<>"" Then
				 Dim ColumnType
				 If MaxLength=0 Then
				 ColumnType="nvarchar(255)"
				 Else
				 ColumnType="nvarchar(" &MaxLength&")"
				 End If
				 on error resume next
				 Conn.Execute("Alter Table "&TableName&" Alter Column "&FieldName&" "&ColumnType&"")
				 if err then err.clear
			End If
		 end if


		 If Title="" Then Call KS.AlertHistory("表单项名称必须输入!", -1): Exit Sub
		 If FieldType=4 And Not Isnumeric(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub
		 If FieldType=5 And Not IsDate(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub
		 if FieldType=8 And Not KS.IsValidEmail(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认格式不正确，请输入正确的Email!",-1):Exit Sub

		 If Not IsNumeric(OrderID) Then OrderID=0

		 Set RS = Server.CreateObject("ADODB.RECORDSET")
		  FieldSql = "Select top 1 * From [KS_FormField] Where FieldID=" & ID 
		  RS.Open FieldSql, conn, 1, 3
		  RS("ItemID") = KS.ChkClng(KS.G("ItemID"))
		  RS("Title") = Title
		  RS("Tips") = Tips
		  RS("DefaultValue") = DefaultValue
		  RS("MustFillTF") = MustFillTF
		  RS("ShowOnForm") = ShowOnForm
		  RS("Options") = Options
		  RS("OrderID") = OrderID
		  RS("Width")   = Width
		  RS("Height")  = Height
		  RS("AllowFileExt")= AllowFileExt
		  RS("MaxFileSize") = MaxFileSize
		  RS("Step") = Step
		  RS("UnitOptions")=UnitOptions
		  RS("ShowOnManage")=ShowOnManage
		  RS("MaxLength")   =MaxLength
		  RS.Update
		  RS.Close
		 .Write ("<form name=""split"" action=""KS.Split.asp"" method=""GET"" target=""BottomFrame"">")
		 .Write ("<input type=""hidden"" name=""OpStr"" value=""自定义表单管理 >> <font color=red>自定义表单项管理</font>"">")
		 .Write ("<input type=""hidden"" name=""ButtonSymbol"" value=""Disabled""></form>")
		 .Write ("<script language=""JavaScript"">document.split.submit();</script>")
		 Call KS.Alert("表单项修改成功!", "KS.FormField.asp?ItemID=" & ItemID&"&Page=" & Page)
		 End With
		 End Sub
		 
		 Sub FieldDel()
			Dim TableName:TableName=Conn.Execute("Select TableName From KS_Form  Where ID=" & ItemID)(0)
			Dim ID:ID = KS.G("ID")
			Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			RSObj.Open "Select FieldName,FieldType From KS_FormField Where FieldID IN(" & ID & ")",Conn,1,1
			Do While Not RSObj.Eof 
			  If left(Lcase(RSObj(0)),3)<>"ks_" Then
			   RSObj.Close:Set RSObj=Nothing
			   Response.Write "<script>alert('对不起，系统字段不能删除!');history.back(-1);</script>"
			   Response.End()
			  Else
			   Conn.Execute("Alter Table "& TableName &" Drop column "& RSObj(0) &"")
			  End If
			  RSObj.MoveNext
			Loop
			RSObj.Close:Set RSObj=Nothing
			Conn.Execute("Delete From KS_FormField Where FieldID IN(" & ID & ")")
			Response.Redirect "KS.FormField.asp?ItemID=" & ItemID &"&Page=" & Page
		 End Sub
		 
		 Sub FieldOrder()
			  Dim FieldID:FieldID=KS.G("FieldID")
			  Dim OrderID:OrderID=KS.G("OrderID")
			  Dim I,FieldIDArr,OrderIDArr,ShowOnFormArr,ShowOnManageArr
			  FieldIDArr=Split(FieldID,",")
			  OrderIDArr=Split(OrderID,",")
			  For I=0 To Ubound(FieldIDArr)
			   Conn.Execute("update KS_FormField Set ShowOnForm=" & KS.ChkClng(KS.G("ShowOnForm" & trim(FieldIDArr(I)))) &",ShowOnManage=" & KS.ChkClng(KS.G("ShowOnManage" & trim(FieldIDArr(I)))) &",OrderID=" & OrderIDArr(i) &" where FieldID=" & trim(FieldIDArr(I)))
			  Next
			  Response.Write "<script>alert('批量保存字段设置成功！');location.href='KS.FormField.asp?ItemID=" & ItemID&"&Page=" & Page & "';</script>"
		 End Sub
End Class
%> 
