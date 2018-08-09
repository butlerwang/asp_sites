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
        Private KS,Action,ChannelID,Page,ItemName,TableName,KSCls
		Private I, totalPut, FieldSql, FieldRS,MaxPerPage
		Private FieldName,ID,Contact, Title, Tips, FieldType, DefaultValue, MustFillTF, ShowOnForm, ShowOnUserForm,ShowOnClubForm,Options,OrderID,AllowFileExt,MaxFileSize,Width,Height,EditorType,ShowUnit,UnitOptions,ParentFieldName,MaxLength

		Private Sub Class_Initialize()
		  MaxPerPage =50
		  Set KSCls=New ManageCls
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
             Action=KS.G("Action")
		With Response
		 If Action<>"" Then
		    .Write"<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">"
			.Write "<html xmlns=""http://www.w3.org/1999/xhtml"">"
		 Else
		    .Write "<html>"
		End If
		.Write "<head>"
		.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
		.Write "<title>字段管理</title>"
		.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
			 ChannelID=KS.ChkClng(KS.G("ChannelID"))
			 
			 TableName=KS.C_S(ChannelID,2)
			 If ChannelID=101 Then
			  TableName="KS_User"   : ItemName= "会员" '会员表
			 Else
			  ItemName=KS.C_S(ChannelID,3)
			 End If

			 if ChannelID=101 Then
		       If Not KS.ReturnPowerResult(0, "KMUA10012")  Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
			   End If
			 Else
		       If Not KS.ReturnPowerResult(0, "KSMM10003") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
			   End If
			 End If
			 
			 Select Case Action
			  Case "Add","Edit" Call FieldAddOrEdit(Action)
			  Case "Del"	    Call FieldDel()
			  Case "order"	    Call FieldOrder()
			  Case "AddSave"    Call FieldAddSave()
			  Case "EditSave"   Call FieldEditSave()
			  Case "setshowonform" Call setshowonform()
			  Case "setshowonuserform" Call setshowonuserform()
			  Case Else 	    Call FieldList()
			 End Select
			.Write "</body>"
			.Write "</html>"
		 End With
		End Sub
		
		Sub FieldList()
		
		With Response
		.Write "<script language='JavaScript'>"
		.Write "var Page='" & CurrentPage & "';"
		.Write "var ItemName='" & ItemName & "';"
		.Write "var ChannelID=" & ChannelID & ";"
		.Write "</script>"
		.Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		.Write "<script language='JavaScript' src='../KS_Inc/jquery.js'></script>"
		.Write "<script language='JavaScript' src='Include/ContextMenu1.js'></script>"
		.Write "<script language='JavaScript' src='Include/SelectElement.js'></script>"
		.Write "<script>"
		.Write "$(document).ready(function(){"
		.Write "parent.frames['BottomFrame'].Button1.disabled=true;"
		.Write "parent.frames['BottomFrame'].Button2.disabled=true;"
		.Write "})</script>"
		%>
		 <script language="javascript">
		 var DocElementArrInitialFlag=false;
		var DocElementArr = new Array();
		var DocMenuArr=new Array();
		var SelectedFile='',SelectedFolder='';
		$(document).ready(function()
		{  
		    if (DocElementArrInitialFlag) return;
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
		   location.href='KS.Field.asp?ChannelID='+ChannelID+'&Action=Add';
		   window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('模型管理 >> 模型字段管理 >> <font color=red>新增'+ItemName+'自定义字段</font>')+'&ButtonSymbol=Go';
		}
		function EditField(id)
		{
		  location="KS.Field.asp?ChannelID="+ChannelID+"&Page="+Page+"&Action=Edit&ID="+id;
		  window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('模型管理 >> 模型字段管理 >> <font color=red>编辑'+ItemName+'自定义字段</font>')+'&ButtonSymbol=GoSave';
		}
		function DelField(id)
		{
		if (confirm('真的要删除该自定义字段吗?'))
		 location="KS.Field.asp?ChannelID="+ChannelID+"&Action=Del&Page="+Page+"&id="+id;
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
				  else alert('一次只能编辑一个自定义字段!')	
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
			 alert('请选择要'+alertmsg+'的自定义字段');
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
		{
		 //if (event.keyCode==46)FieldControl(2);
		 }
		}
		 </script>
		<%
		.Write "</head>"
		.Write "<body scroll=no topmargin='0' leftmargin='0' onclick='SelectElement();' onkeydown='GetKeyDown();' onselectstart='return false;'>"
		.Write "<ul id='menu_top'>"
		.Write "<li class='parent' onclick=""FieldAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>新增字段</span></li>"
		.Write "<li class='parent' onclick=""FieldControl(1);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>修改字段</span></li>"
		.Write "<li class='parent' onclick=""FieldControl(2);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>删除字段</span></li>"
		.Write "<li class='parent' onclick=""location.href='KS.Model.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>回上一级</span></li>"
		.Write "</ul>"
        
		.Write ("<div style=""height:94%; overflow: scroll; width:100%"" align=""center"">")
		.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		.Write "<form action='KS.Field.asp?action=order&channelid=" & ChannelID&"&page="&CurrentPage &"'' name='form1' method='post'>"
		'if CurrentPage<=1  and channelid<>9 then
		'.Write " <tr><td style='border:1px solid #f9c943;background:#FFFFF6;padding:2px;padding-left:8px' colspan='10'><input class='button' type='button' value='显示系统字段' onclick=""if (this.value=='显示系统字段'){this.value='隐藏系统字段'}else{this.value='显示系统字段'};$('table').find('[name=sysfield]').toggle();""/><span class='tips'>&nbsp;Tips:系统字段已默认不显示，您可以点击左边的按钮显示。</span></td></tr>" 
	    'end if
		
		.Write " <tr class='sort'>"
		.Write "   <td width='80' align='center'>排序</td>"
		.Write "   <td width='100' align='center'>字段名称</td>"
		.Write "   <td align='center'>字段别名</td>"		
		.Write "   <td align='center'>归属模型</td>"
		.Write "   <td align='center'>字段类型</td>"
		.Write "   <td align='center'>后台显示</td>"
		.Write "   <td align='center'>前台显示</td>"
		.Write "   <td align='center'>↓管理操作</td>"
		.Write " </tr>"
			 Set FieldRS = Server.CreateObject("ADODB.RecordSet")
				   FieldSql = "SELECT * FROM KS_Field Where ChannelID=" & ChannelID & " order by orderid asc"
				   FieldRS.Open FieldSql, conn, 1, 1
				 If FieldRS.EOF And FieldRS.BOF Then
				 Else
					        totalPut = FieldRS.RecordCount
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									FieldRS.Move (CurrentPage - 1) * MaxPerPage
							End If
							Call showContent
			End If
		 .Write " <tr>"
		 .Write "   <td colspan='3'>&nbsp;&nbsp;<input type='submit' class='button' value='批量保存字段排序'> <font color=blue>值越小排在越前面</font></td></form>"
		 .Write "   <td height='35' colspan='5' align='right'>"
		 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		.Write "    </td>"
		.Write " </tr>"
		.Write "</table>"
		.Write "<br/><br/><br/></div>"
		End With
		End Sub
		Sub showContent()
		With Response
		Do While Not FieldRS.EOF
		 if KS.ChkClng(FieldRS("FieldType"))=0 or Left(Lcase(FieldRS("FieldName")),3)<>"ks_" Then
		 .Write "<tr name='sysfield' style='display:'>"
		 Else
		 .Write "<tr>"
		 End If
		 .Write "<td class='splittd'>&nbsp;&nbsp;<input type='text' name='OrderID' style='width:50px;text-align:center' value='" & FieldRS("OrderID") &"'><input type='hidden' name='FieldID' value='" & FieldRS("FieldID") & "'></td>"
		 .Write "  <td class='splittd'><span FieldID='" & FieldRS("FieldID") & "' onDblClick=""EditField(this.FieldID)""><img src='Images/Field.gif' align='absmiddle'><span  style='cursor:default;'>" & FieldRS("FieldName") & "</span></span></td>"
		 .Write "   <td align='center' class='splittd'>" & FieldRS("Title") & " </td>"
		 .Write "   <td align='center' class='splittd'><font color=red>"
		 If ChannelID=101 Then
		 .Write "会员系统"
		 Else
		  .Write KS.C_S(ChannelID,1) 
		 End If
		  .Write "</font>"
		 .Write "</td>"
		 .Write "   <td align='center' class='splittd'>"
				 Select Case FieldRS("FieldType")
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
				  Case 12: .Write "小数(text)"
				  Case 13: .Write "文档属性(checkbox)"
				 End Select
		  If Left(Lcase(FieldRS("FieldName")),3)<>"ks_" Then .Write "<font color=#cccccc>[系统]</font>"
		 .Write "</td>"
		 .Write "   <td align='center' class='splittd'>" 
		  If FieldRS("ShowOnForm")=1 Then
		   .Write "<a title='设置为后台不显示' href='?channelid=" & channelid & "&action=setshowonform&id=" & FieldRS("FieldID") &"&v=0'><font color=red>是</font></a>"
		  Else
		   .Write "<a title='设置为后台显示' href='?channelid=" & channelid & "&action=setshowonform&id=" & FieldRS("FieldID") &"&v=1'><font color=green>否</font></a>"
		  End If
		 .Write " </td>"
		 .Write "   <td align='center' class='splittd'>" 
		  If FieldRS("ShowOnUserForm")=1 Then
		   .Write "<a title='设置为前台不显示' href='?channelid=" & channelid & "&action=setshowonuserform&id=" & FieldRS("FieldID") &"&v=0'><font color=red>是</font></a>"
		  Else
		   .Write "<a title='设置为前台显示' href='?channelid=" & channelid & "&action=setshowonuserform&id=" & FieldRS("FieldID") &"&v=1'><font color=green>否</font></a>"
		  End If
		 .Write " </td>"
		 .Write " <td align='center' class='splittd'><a href='javascript:EditField(" & FieldRS("FieldID") &");'>修改</a> | "
		 If Left(Lcase(FieldRS("FieldName")),3)<>"ks_" Then
		 .Write "<font color=#cccccc title='系统字段不允许删除'>删除</font>"
		 Else
		 .Write "<a href='javascript:DelField(" & FieldRS("FieldID") &");'>删除</a>"
		 End If
		 .Write " </td></tr>"
								I = I + 1
								If I >= MaxPerPage Then Exit Do
							   FieldRS.MoveNext
							   Loop
								FieldRS.Close
						 
         End With
		 End Sub
		 
		 Sub setshowonform()
		    dim id:id=KS.ChkClng(request("id"))
			conn.execute("update KS_Field Set ShowOnForm=" & KS.ChkClng(Request("v")) & " Where FieldID=" & ID)
			Call KSCls.CreateFieldXML(ChannelID,"") '生成xml缓存
			response.Redirect request.ServerVariables("HTTP_REFERER")
		 End Sub
		 Sub setshowonuserform()
		    dim id:id=KS.ChkClng(request("id"))
			conn.execute("update KS_Field Set ShowOnUserForm=" & KS.ChkClng(Request("v")) & " Where FieldID=" & ID)
			Call KSCls.CreateFieldXML(ChannelID,"") '生成xml缓存
			response.Redirect request.ServerVariables("HTTP_REFERER")
		 End Sub
		 
		 Sub FieldAddOrEdit(OpType)
		 With Response
		  Dim FieldRS, FieldSql,OpAction,OpTempStr
		 ID = KS.G("ID")
		.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
        .Write "<html xmlns=""http://www.w3.org/1999/xhtml"">"
		.Write "<head>"
		.Write "<meta http-equiv='Content-Type' content='text/html; chaRSet=utf-8'>"
		.Write "<title>字段管理</title>"
		.Write "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		.Write "<script language='JavaScript' src='../KS_Inc/jquery.js'></script>"
		.Write "</head>"
		.Write "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
		
		 If Optype = "Edit" Then
		     OpAction="EditSave":OpTempStr="编辑"
			 Set FieldRS = Server.CreateObject("ADODB.RECORDSET")
			 FieldSql = "Select TOP 1 * From [KS_Field] Where FieldID=" & ID
			 FieldRS.Open FieldSql, conn, 1, 1
			 If Not FieldRS.EOF Then
				 FieldName = Trim(FieldRS("FieldName"))
				 ChannelID = FieldRS("ChannelID")
				 Title = Trim(FieldRS("Title"))
				 Tips = Trim(FieldRS("Tips"))
				 FieldType = Trim(FieldRS("FieldType"))
				 DefaultValue = Trim(FieldRS("DefaultValue"))
				 MustFillTF = FieldRS("MustFillTF")
				 ShowOnForm = FieldRS("ShowOnForm")
				 ShowOnUserForm=FieldRS("ShowOnUserForm")
				 ShowOnClubForm=FieldRS("ShowOnClubForm")
				 Options = Trim(FieldRS("Options"))
				 OrderID= FieldRS("OrderID")
				 AllowFileExt=FieldRS("AllowFileExt")
				 MaxFileSize=FieldRS("MaxFileSize")
				 Width=FieldRS("Width")
				 Height=FieldRS("Height")
				 MaxLength=FieldRS("MaxLength")
				 EditorType=FieldRS("EditorType")
				 ShowUnit=FieldRS("ShowUnit")
				 UnitOptions=FieldRS("UnitOptions")
				 ParentFieldName=FieldRS("ParentFieldName")
			 End If
	  Else
	     FieldName="KS_":FieldType=1:MustFillTF=0:ShowOnForm=1:ShowOnUserForm=1:ShowOnClubForm=0:AllowFileExt="jpg|gif|png":MaxFileSize=1024:Width=200:Height=80:EditorType="Basic":ShowUnit=0:MaxLength=255
		 OpAction="AddSave":OpTempStr="添加"
		 OrderID=KS.ChkClng(Conn.Execute("Select Max(OrderID) From KS_Field Where ChannelID=" & ChannelID)(0))+1
	  End If
		 
		.Write "<div class='topdashed sort'>" & OpTempStr &"自定义字段</div>"
		.Write "<br>"
        .Write "        <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1""  class='ctable'>" & vbCrLf
		.Write "  <form  action='KS.Field.asp?Action=" & OpAction &"' method='post' name='OrigArticlerm' onsubmit='return(CheckForm())'>"
		
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' class='clefttitle'><strong>所属系统：</strong></td>"
		.Write "      <td nowrap> &nbsp;&nbsp;<font color=#ff0000>"
		If ChannelID=101 Then
		.Write "会员系统"
		Else
		.Write KS.C_S(ChannelID,1)
		End If
		.Write "</font><input type='hidden' value='" & ChannelID & "' name='ChannelID'></td>"
		.Write "    </tr>"

   
    If FieldType="0" Then  '系统内置字段
		.Write "   <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td width='230' height='30' align='right' class='clefttitle'><strong>字段名称：</strong><br><font color=blue></font></td>"
		.Write "      <td height='45' nowrap>&nbsp;<input class='textbox' name='FieldName' type='text' readonly id='FieldName' value='" & FieldName & "' size='30'> </td></tr>" &vbcrlf
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='45' align='right' class='clefttitle'><strong>字段别名：</strong><br><font color=blue>便于在管理项目中显示</font></td>"
		.Write "      <td height='45' nowrap>&nbsp;<input name='Title' type='text' id='Title' size='30' class='textbox' value='" & Title & "'> *</td></tr>"
.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>字段类型：</strong></td>"
		.Write "      <td nowrap>&nbsp;<input type='hidden' value=" & FieldType & " name='FieldType'><select name=""FieldType"" disabled><option>系统内置字段</option></select></td>"
		.Write "    </tr>"		
.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>后台是否启用：</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='ShowOnForm' type='radio' id='ShowOnForm' value='1'"
		If ShowOnForm="1" Then .Write " Checked"
		.Write ">是"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnForm' type='radio' id='ShowOnForm' value='0'"
		If ShowOnForm="0" Then .Write " Checked"
		.WRite ">否"
		.Write "       </td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>会员中心是否启用：</strong><br><font color=blue>必须是启用，前台的会员中心才会显示</font></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='ShowOnUserForm' type='radio' id='ShowOnUserForm' value='1'"
		If ShowOnUserForm="1" Then .Write " Checked"
		.Write ">是"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnUserForm' type='radio' id='ShowOnUserForm' value='0'"
		If ShowOnUserForm="0" Then .Write " Checked"
		.WRite ">否"
		.Write "       </td>"
		.Write "    </tr>"	
		
	Else
		.Write "   <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td width='230' height='30' align='right' class='clefttitle'><strong>字段名称：</strong><br><font color=blue>为了和系统字段区分，必须以“KS_”开头,在模板中可以通过“{$KS_字段名称}”进行调用</font></td>"
		.Write "      <td height='45' nowrap>&nbsp;"
		.Write "        <input name='FieldName' type='text' id='FieldName' value='" & FieldName & "' size='30'"
		If Optype = "Edit" Then .Write " readonly"
		.Write " class='textbox'>* </td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='45' align='right' class='clefttitle'><strong>字段别名：</strong><br><font color=blue>便于在管理项目中显示</font></td>"
		.Write "      <td height='45' nowrap>&nbsp;&nbsp;<input name='Title' type='text' id='Title' size='30' class='textbox' value='" & Title & "'> *</td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right'' class='clefttitle'><strong>附加提示：</strong><br><font color=blue>在输入框旁边的提示信息</font></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<textarea name='Tips'  id='Tips' class='textbox' style='width:300px;height:60px'>" & Tips & "</textarea><font color=green>可以加入一些javascript事件</font></td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>字段类型：</strong></td>"
		If Optype = "Edit" Then
		.Write "      <td nowrap>&nbsp;&nbsp;<input type='hidden' value=" & FieldType & " name='FieldType'><select name=""FieldType"" disabled>"
		else
		.Write "      <td nowrap>&nbsp;&nbsp;<select name=""FieldType"" onchange=""Setdisplay(this.value)"">"
		end if
		If ChannelID<>101 and channelid<>9 Then
		.Write " <option value=""13"""
		If FieldType=13 Then .Write " Selected"
		.Write ">文档属性(标签调用属性)</option>"
		End If
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
        .Write " <option value=""12"""
		If FieldType=12 Then .Write " Selected"
		.Write ">小数(text)</option>"
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
		.Write "<font color=red>说明：一旦设定不能修改</font>"
		.Write " </td>"
		.Write "    </tr>"
		
		.Write "  <tbody id='editorarea'>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>编辑器类型：</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='EditorType' type='text' id='EditorType' class='textbox' value='" & EditorType & "' size='10'>&nbsp;<select onchange=""$('#EditorType').val(this.value)"" name='selecteditor'><option value='Default'>Default</option><option value='NewsTool'>NewsTool</option><option value='Simple'>Simple</option><option value='Basic'>Basic</option></select><span style='color:green'>您可以打开/Editor/config.js自定义编辑器类型</span>"
		.Write "       </td>"
		.Write "    </tr>"
		.Write " </tbody>"
		
		
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
		If ChannelID<>101 Then
		 .Write "<div id='dtips1'>&nbsp;&nbsp;<font color=green>为便于会员获取默认值，可绑定表KS_User或KS_Enterprise的字段值<br>&nbsp;&nbsp;格式：表名|字段名 如：<font color=red>KS_User|RealName</font></font><br/>&nbsp;&nbsp;<font color=blue>也可以将默认值设置为now或date取得当前时间</font></div><div id='dtips2'>&nbsp;&nbsp;输入<font color=red>“1”</font>，则添加文档时默认为该属性为选中状态</div>"
		End If
		.Write "       </td>"
		.Write "    </tr>"
		
		.Write "<tbody id='showattrarea'>"
		
		.Write "    <tr id=""ldArea"" style='display:none' class='tdbg'>"
		.Write "      <td align='right' class='clefttitle'><strong>所属父级字段：</strong><br><font color=blue>不选择表示一级联动字段<br/>不能指定为下级联动字段</font></td>"
		.Write "      <td>&nbsp;&nbsp;"
		  Dim PRS
		  If KS.ChkClng(ID)<>0 Then
		  Set PRS=Conn.Execute("Select FieldName,Title From KS_Field Where ChannelID=" & ChannelID& " and FieldType=11 And FieldID<>" & ID & " Order BY FieldID")
		  .Write "<select name='ParentFieldName' disabled>"
		  Else
		  Set PRS=Conn.Execute("Select FieldName,Title From KS_Field Where ChannelID=" & ChannelID& " and FieldType=11 Order BY FieldID")
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
		.Write "      <td align='right'  class='clefttitle'><strong>列表选项：</strong><br><font color=blue>每一行为一个列表选项</font><br>如果值和显示项不同可以用<font color=red>|</font>隔开<br>正确格式如：<font color=red>男</font> 或 <font color=red>0|男</font><br></td>"
		.Write "      <td height='45' nowrap>&nbsp;&nbsp;<textarea name='Options' style='height:70px' cols='50' rows='6' id='Options' class='textbox'>" & Options & "</textarea>"
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
		If MustFillTF="1" Then .Write " Checked"
		.Write ">是"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='MustFillTF' type='radio' id='MustFillTF' value='0'"
		If MustFillTF="0" Then .Write " Checked"
		.WRite ">否"
		.Write "       </td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>后台是否启用：</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='ShowOnForm' type='radio' id='ShowOnForm' value='1'"
		If ShowOnForm="1" Then .Write " Checked"
		.Write ">是"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnForm' type='radio' id='ShowOnForm' value='0'"
		If ShowOnForm="0" Then .Write " Checked"
		.WRite ">否"
		.Write "       </td>"
		.Write "    </tr>"
		'If ChannelID=101 Then 
		'.Write "    <tr style='display:none' "
		'Else
		.Write "    <tr "
		'End If
		.Write "class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>会员中心是否启用：</strong><br><font color=blue>必须是启用，前台的会员中心才会显示</font></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='ShowOnUserForm' type='radio' id='ShowOnUserForm' value='1'"
		If ShowOnUserForm="1" Then .Write " Checked"
		.Write ">是"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnUserForm' type='radio' id='ShowOnUserForm' value='0'"
		If ShowOnUserForm="0" Then .Write " Checked"
		.WRite ">否"
		.Write "       </td>"
		.Write "    </tr>"
		
		If  KS.ChkClng(KS.C_S(ChannelID,6))=1 Then
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>推送到论坛时显示：</strong><br><font color=blue>指当文章被推送到论坛时是否显示该字段内容</font></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='ShowOnClubForm' type='radio' value='1'"
		If ShowOnClubForm="1" Then .Write " Checked"
		.Write ">是"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnClubForm' type='radio' value='0'"
		If ShowOnClubForm="0" Then .Write " Checked"
		.WRite ">否"
		.Write "       </td>"
		.Write "    </tr>"
		End If
		
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>显示设置：</strong> </td>"
		.Write "      <td nowrap>&nbsp;&nbsp;宽度<input name='Width' type='text' site='10' class='textbox' style='width:40px' id='Width' value='" & Width & "'>px <font color=red>例如：200px</font>  长度<input name='MaxLength' type='text' site='10' class='textbox' style='width:40px' id='MaxLength' value='" & MaxLength & "'>个字符 不限制请输入0<br><span style='display:none' id='heightarea'>&nbsp;&nbsp;高度<input name='Height' type='text' site='10' class='textbox' style='width:40px' id='Height' value='" & Height & "'>px <font color=red>例如：100px</font></span>"
		.Write "       </td>"
		.Write "    </tr>"	
		.Write "</tbody>"
    End If		
			
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>排序序号：</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='OrderID' type='text' style='text-align:center' size='8' class='textbox' id='OrderID' value='" & OrderID & "'> <font color=blue>序号越小，排在越前面</font>"
		.Write "       </td>"
		.Write "    </tr>"
	
		.Write "   <input type='hidden' value='" & ID & "' name='id'>"
		.Write "    <input type='hidden' value='" & Page & "' name='page'>"
		.Write "  </form>"
		.Write "</table>"
		
		
		 
		.Write "<Script Language='javascript'>"
		If FieldType<>"0" Then
		.Write "Setdisplay(" & FieldType & ");"
		.Write "function Setdisplay(s)"
		.Write  "{if (s==3||s==6||s==7||s==11){ $('#OptionsArea').show();} else $('#OptionsArea').hide();if (s==7)$('#darea').show();else $('#darea').hide();if(s==9)$('#extarea').show();else $('#extarea').hide(); if(s==10)$('#editorarea').show();else $('#editorarea').hide();if (s==2||s==10) $('#heightarea').show();else $('#heightarea').hide();if(s==11) $('#ldArea').show(); else $('#ldArea').hide();if(s==13){$('#showattrarea').hide();$('#dtips1').hide();$('#dtips2').show();}else{$('#showattrarea').show();$('#dtips2').hide();$('#dtips1').show();}"
		.Write "}"
		End If
		.Write "function CheckForm()"
		.Write "{ var form=document.OrigArticlerm;"
		.Write "   if (form.FieldName.value==''||form.FieldName.value.length<=1)"
		.Write "    {"
		.Write "     alert('请输入字段名称!');"
		.Write "     form.FieldName.focus();"
		.Write "     return false;"
		.Write "    }"
		.Write "   if (form.Title.value=='')"
		.Write "    {"
		.Write "     alert('请输入字段标题!');"
		.Write "     form.Title.focus();"
		.Write "     return false;"
		.Write "    }"
		.Write "    form.submit();"
		.Write "    return true;"
		.Write "}"
		.Write "</Script>"
		End With
		End Sub
		 
		 Sub FieldAddSave()
		 Dim FieldRS,ColumnType
		 FieldName = Trim(KS.G("FieldName"))
		 Title = KS.G("Title")
		 Tips = Request.Form("Tips")
		 FieldType = KS.G("FieldType")
		 DefaultValue = KS.G("DefaultValue")
		 MustFillTF = KS.G("MustFillTF")
		 FieldType = KS.G("FieldType")
		 ShowOnForm = KS.ChkClng(KS.G("ShowOnForm"))
		 ShowOnUserForm=KS.ChkClng(KS.G("ShowOnUserForm"))
		 ShowOnClubForm=KS.ChkClng(KS.G("ShowOnClubForm"))
		 Options = KS.G("Options")
		 FieldType =KS.G("FieldType")
		 Width=KS.G("Width")
		 MaxLength=KS.ChkClng(KS.G("MaxLength"))
		 AllowFileExt=KS.G("AllowFileExt")
		 EditorType=KS.G("EditorType")
		 ShowUnit  =KS.ChkClng(KS.G("ShowUnit"))
		 UnitOptions=KS.G("UnitOptions")
		 ParentFieldName=KS.G("ParentFieldName")
		 If KS.IsNul(ParentFieldName) Then ParentFieldName="0"

		 If FieldName = "" Then Call KS.AlertHistory("请输入字段名称!", -1): Exit Sub
		 If Len(FieldName)<=3 Then Call KS.AlertHistory("字段名称长度必须大于3!", -1): Exit Sub
		 If Ucase(Left(FieldName,3))<>"KS_" Then Call KS.AlertHistory("字段名称格式有误，必须以""KS_开头""!", -1): Exit Sub
		 If Title="" Then Call KS.AlertHistory("字段标题必须输入!", -1): Exit Sub
		 If FieldType=4 And Not Isnumeric(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub
		 If FieldType=5 And Not IsDate(DefaultValue) And DefaultValue<>"" and lcase(DefaultValue)<>"now" and lcase(DefaultValue)<>"date" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub
		 if FieldType=8 And Not KS.IsValidEmail(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认格式不正确，请输入正确的Email!",-1):Exit Sub
		 Select Case FieldType
		   Case 1,3,6,7,8,9,11
		     If MaxLength=0 Then
		     ColumnType="nvarchar(255)"
			 Else
		     ColumnType="nvarchar(" &MaxLength&")"
			 End If
		   Case 13
		    ColumnType="tinyint default 0"
		   Case 2,10
		     ColumnType="ntext"
		   Case 5
		     ColumnType="datetime"
		   Case 4
		     ColumnType="int"
		   Case 12
		     ColumnType="float"
		   Case else
		     Exit Sub
		 End Select
		 Set FieldRS = Server.CreateObject("ADODB.RECORDSET")
		 FieldSql = "Select top 1 * From [KS_Field] Where FieldName='" & FieldName & "' And ChannelID=" & KS.G("ChannelID")
		 FieldRS.Open FieldSql, conn, 3, 3
		 If FieldRS.EOF And FieldRS.BOF Then
		  FieldRS.AddNew
		  FieldRS("FieldName") = FieldName
		  FieldRS("ChannelID") = KS.G("ChannelID")
		  FieldRS("Title") = Title
		  FieldRS("Tips") = Tips
		  FieldRS("FieldType") = FieldType
		  FieldRS("DefaultValue") = DefaultValue
		  FieldRS("MustFillTF") = MustFillTF
		  FieldRS("FieldType") = FieldType
		  FieldRS("ShowUnit")=ShowUnit
		  FieldRS("UnitOptions")=UnitOptions
		  FieldRS("ShowOnForm") = ShowOnForm
		  FieldRS("ShowOnUserForm")=ShowOnUserForm
		  FieldRS("ShowOnClubForm")=ShowOnClubForm
		  FieldRS("Options") = Options
		  FieldRS("OrderID")=KS.ChkClng(KS.G("OrderID"))
		  FieldRS("AllowFileExt")=KS.G("AllowFileExt")
		  FieldRS("MaxFileSize")=KS.ChkClng(KS.G("MaxFileSize"))
		  FieldRS("Width")=KS.ChkClng(KS.G("Width"))
		  FieldRS("Height")=KS.ChkClng(KS.G("Height"))
		  FieldRS("MaxLength")=MaxLength
		  FieldRS("EditorType")=EditorType
		  FieldRS("ParentFieldName")=ParentFieldName
		  FieldRS.Update
		  Conn.Execute("Alter Table "&TableName&" Add "&FieldName&" "&ColumnType&"")
		  If ShowUnit=1 Then  '增加单位字段
		  Conn.Execute("Alter Table "&TableName&" Add "&FieldName&"_Unit nvarchar(200)")
		  End If
		  
		  on error resume next
		  If KS.C_S(KS.G("ChannelID"),6)=1 or KS.C_S(KS.G("ChannelID"),6)=2 or KS.C_S(KS.G("ChannelID"),6)=5 Then
		  KS.ConnItem.Execute("Alter Table "&TableName&" Add "&FieldName&" "&ColumnType&"")
		  If ShowUnit=1 Then  '增加单位字段
		  KS.ConnItem.Execute("Alter Table "&TableName&" Add "&FieldName&"_Unit nvarchar(200)")
		  End If
		  if err then err.clear
		  
		  End If
		   Call KSCls.CreateFieldXML(ChannelID,"") '生成xml缓存
		 Response.Write ("<Script> if (confirm('字段增加成功,继续添加吗?')) { location.href='KS.Field.asp?ChannelID=" & ChannelID& "&Action=Add';} else{location.href='KS.Field.asp?ChannelID=" & ChannelID&"&Page='"&Page &";$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=模型管理 >> <font color=#ff0000>模型字段管理</font>&ButtonSymbol=Disabled';}</script>")
		 Else
		   Call KS.AlertHistory("数据库中已存在该字段名称!", -1)
		   Exit Sub
		 End If
		 FieldRS.Close
		 End Sub
		 
		 Sub FieldEditSave()
		 With Response
		 ID = KS.G("ID")
		 FieldName = Trim(KS.G("FieldName"))
		 Title = KS.G("Title")
		 Tips = Request.Form("Tips")
		 DefaultValue = KS.G("DefaultValue")
		 MustFillTF = KS.ChkClng(KS.G("MustFillTF"))
		 FieldType = KS.G("FieldType")
		 ShowOnForm = KS.ChkClng(KS.G("ShowOnForm"))
		 ShowOnUserForm=KS.ChkClng(KS.G("ShowOnUserForm"))
		 ShowOnClubForm=KS.ChkClng(KS.G("ShowOnClubForm"))
		 Options = KS.G("Options")
		 FieldType =KS.G("FieldType")
		 OrderID   =KS.G("OrderID")
		 EditorType=KS.G("EditorType")
		 ShowUnit  =KS.ChkClng(KS.G("ShowUnit"))
		 UnitOptions=KS.G("UnitOptions")
		 ParentFieldName=KS.G("ParentFieldName")
		 If KS.IsNul(ParentFieldName) Then ParentFieldName="0"
		 
		 If Title="" Then Call KS.AlertHistory("字段标题必须输入!", -1): Exit Sub
		' If FieldType=4 And Not Isnumeric(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub
		
		 If FieldType=5 And Not IsDate(DefaultValue) And DefaultValue<>"" and lcase(DefaultValue)<>"now" and lcase(DefaultValue)<>"date" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub
		
		
		 If Not IsNumeric(OrderID) Then OrderID=0
		 
		 '修改字段长度
		 if (FieldType=1 or FieldType=3 or FieldType=6 or FieldType=7 or FieldType=8 or FieldType=9 or FieldType=11) then
		     Dim ColumnType
			 If KS.ChkClng(KS.G("MaxLength"))=0 Then
		     ColumnType="nvarchar(255)"
			 Else
		     ColumnType="nvarchar(" &KS.ChkClng(KS.G("MaxLength"))&")"
			 End If
			 on error resume next
			 Conn.Execute("Alter Table "&TableName&" Alter Column "&FieldName&" "&ColumnType&"")
			 if err then err.clear
		 end if

		 Set FieldRS = Server.CreateObject("ADODB.RECORDSET")
		  FieldSql = "Select top 1 * From [KS_Field] Where FieldID=" & ID 
		  FieldRS.Open FieldSql, conn, 1, 3
		  FieldRS("ChannelID") = KS.G("ChannelID")
		  FieldRS("Title") = Title
		  FieldRS("Tips") = Tips
		  FieldRS("DefaultValue") = DefaultValue
		  FieldRS("MustFillTF") = MustFillTF
		  If FieldRS("FieldType")=4 And Not Isnumeric(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub
		'  If FieldRS("FieldType")=5 And Not IsDate(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("默认值格式不正确!", -1): Exit Sub

		 ' FieldRS("FieldType") = FieldType
		 ' FieldRS("ShowUnit")=ShowUnit
		  'FieldRS("ParentFieldName")=ParentFieldName
		  FieldRS("UnitOptions")=UnitOptions
		  FieldRS("ShowOnForm") = ShowOnForm
		  FieldRS("ShowOnUserForm")=ShowOnUserForm
		  FieldRS("ShowOnClubForm")=ShowOnClubForm
		  FieldRS("Options") = Options
		  FieldRS("OrderID")=OrderID
		  FieldRS("AllowFileExt")=KS.G("AllowFileExt")
		  FieldRS("MaxFileSize")=KS.ChkClng(KS.G("MaxFileSize"))
		  FieldRS("Width")=KS.ChkClng(KS.G("Width"))
		  FieldRS("Height")=KS.ChkClng(KS.G("Height"))
		  FieldRS("MaxLength")=KS.ChkClng(KS.G("MaxLength"))
		  FieldRS("EditorType")=EditorType
		  FieldRS.Update
		  FieldRS.Close
		  on error resume next
	   	  If KS.C_S(KS.G("ChannelID"),6)=1 or KS.C_S(KS.G("ChannelID"),6)=2 or KS.C_S(KS.G("ChannelID"),6)=5 Then
			KS.ConnItem.Execute("Update KS_FieldItem Set FieldTitle='" & Title & "',OrderID=" & OrderID &" Where FieldID=" & ID)
          End If
		 .Write ("<form name=""split"" action=""KS.Split.asp"" method=""GET"" target=""BottomFrame"">")
		 .Write ("<input type=""hidden"" name=""OpStr"" value=""模型管理 >> <font color=red>模型字段管理</font>"">")
		 .Write ("<input type=""hidden"" name=""ButtonSymbol"" value=""Disabled""></form>")
		 .Write ("<script language=""JavaScript"">document.split.submit();</script>")
		  Call KSCls.CreateFieldXML(ChannelID,"") '生成xml缓存
		 Call KS.Alert("字段修改成功!", "KS.Field.asp?ChannelID=" & ChannelID&"&Page=" & Page)
		 End With
		 End Sub
		 
		 Sub FieldDel()
		    on error resume next
			Dim ID:ID = KS.G("ID")
			Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			RSObj.Open "Select FieldName,FieldType,ShowUnit From KS_Field Where FieldID IN(" & ID & ")",Conn,1,1
			Do While Not RSObj.Eof 
			  If left(Lcase(RSObj(0)),3)<>"ks_" Then
			   RSObj.Close:Set RSObj=Nothing
			   Response.Write "<script>alert('对不起，系统字段不能删除!');history.back(-1);</script>"
			   Response.End()
			  Else
			      Conn.Execute("Alter Table "& TableName &" Drop column "& RSObj(0) &"")
				  If RSObj("ShowUnit")="1" Then
			      Conn.Execute("Alter Table "& TableName &" Drop column "& RSObj(0) &"_Unit")
				  End if
			   	  If KS.C_S(KS.G("ChannelID"),6)=1 or KS.C_S(KS.G("ChannelID"),6)=2 or KS.C_S(KS.G("ChannelID"),6)=5  Then
					  KS.ConnItem.Execute("Delete From KS_FieldItem Where FieldID IN(" & ID & ")")
					  KS.ConnItem.Execute("Delete From KS_FieldRules Where FieldID IN(" & ID & ")")
					  KS.ConnItem.Execute("Alter Table "& TableName &" Drop column "& RSObj(0) &"")
					  If RSObj("ShowUnit")="1" Then
					  KS.ConnItem.Execute("Alter Table "& TableName &" Drop column "& RSObj(0) &"_Unit")
					  End if
				  End If
			  End If
			  RSObj.MoveNext
			Loop
			RSObj.Close:Set RSObj=Nothing
			Conn.Execute("Delete From KS_Field Where FieldID IN(" & ID & ")")
			Call KSCls.CreateFieldXML(ChannelID,"") '生成xml缓存
			Response.Redirect "KS.Field.asp?ChannelID=" & ChannelID &"&Page=" & Page
		 End Sub
		 
		 Sub FieldOrder()
			  Dim FieldID:FieldID=KS.G("FieldID")
			  Dim OrderID:OrderID=KS.G("OrderID")
			  Dim I,FieldIDArr,OrderIDArr
			  FieldIDArr=Split(FieldID,",")
			  OrderIDArr=Split(OrderID,",")
			  For I=0 To Ubound(FieldIDArr)
			   Conn.Execute("update KS_Field Set OrderID=" & OrderIDArr(i) &" where FieldID=" & FieldIDArr(I))
			   on error resume next
			   If KS.C_S(ChannelID,6)=1 Then
				KS.ConnItem.Execute("Update KS_FieldItem Set OrderID=" & OrderIDArr(i) &" Where FieldID=" & FieldIDArr(I))
			   End If
			  Next
			  Call KSCls.CreateFieldXML(ChannelID,"") '生成xml缓存
			  Response.Write "<script>alert('批量保存字段排序成功！');location.href='?ChannelID=" & ChannelID & "&Page=" & Page&"';</script>"
		 End Sub
End Class
%> 
