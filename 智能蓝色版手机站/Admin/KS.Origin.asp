<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Origin
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Origin
        Private KS,Action,Page,KSCls
		Private I, totalPut, CurrentPage, OriginSql, RS,MaxPerPage
		Private OriginName,ID,Contact, Telphone, UnitName, UnitAddress, Zip, Email, QQ, HomePage, Note, OriginType

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
		With KS
		 CurrentPage = KS.ChkClng(KS.G("page"))
		 If CurrentPage=0 Then CurrentPage=1

		 .echo "<html>"
		 .echo "<head>"
		 .echo "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
		 .echo "<title>来源管理</title>"
		 .echo "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		 .echo "<script language='JavaScript'>"
		 .echo "var Page='" & CurrentPage & "';"
		 .echo "</script>"
		 .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
	     .echo "<script language=""JavaScript"" src=""../KS_Inc/jQuery.js""></script>" & vbCrLf
             Action=KS.G("Action")
			 Page=KS.G("Page")
			 
			 If Not KS.ReturnPowerResult(0, "KMST10015") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
			 End iF
			 
			 Select Case Action
			  Case "Add"
			    Call OriginAddOrEdit("Add")
			  Case "Edit"
			    Call OriginAddOrEdit("Edit")
			  Case "Del"
			    Call OriginDel()
			  Case "AddSave"
			    Call OriginAddSave()
			  Case "EditSave"
			    Call OriginEditSave()
			  Case Else
			   Call OriginList()
			 End Select
			 .echo "</body>"
			 .echo "</html>"
		 End With
		End Sub
		
		Sub OriginList()
		 On Error Resume Next
		With KS
		%>
		<script language="javascript">
		   function set(v)
			{
				 if (v==1)
				 KeyWordControl(1);
				 else if (v==2)
				 KeyWordControl(2);
			}
		function OriginAdd()
		{
		  new parent.KesionPopup().PopupCenterIframe('新增来源','KS.Origin.asp?Action=Add',630,410,'no')
		}
		function EditOrigin(id)
		{ 
		new parent.KesionPopup().PopupCenterIframe('编辑来源',"KS.Origin.asp?action=Edit&ID="+id,630,410,'no')
		}
		function DelOrigin(id)
		{
		if (confirm('真的要删除该来源吗?'))
		 location="KS.Origin.asp?Action=Del&Page="+Page+"&id="+id;
		  SelectedFile='';
		}
		function OriginControl(op)
		{  
		    var alertmsg='';
	        var ids=get_Ids(document.myform);
			if (ids!='')
			 {
			   if (op==1)
				{
				if (ids.indexOf(',')==-1) 
					EditOrigin(ids)
				  else alert('一次只能编辑一个来源!')	
				}	
			  else if (op==2)    
			   DelOrigin(ids);
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
			 alert('请选择要'+alertmsg+'的来源');
			  }
		}
		function GetKeyDown()
		{ 
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 : location.reload(); break;
			 case 65 : Select(0);break;
			 case 78 : event.keyCode=0;event.returnValue=false;OriginAdd();break;
			 case 69 : event.keyCode=0;event.returnValue=false;OriginControl(1);break;
			 case 68 : OriginControl(2);break;
		   }	
		else	
		 if (event.keyCode==46)OriginControl(2);
		}
		</script>
		<%
		 .echo "</head>"
		 .echo "<body scroll=no topmargin='0' leftmargin='0' onkeydown='GetKeyDown();' onselectstart='return false;'>"
		 .echo "<ul id='menu_top'>"
		 .echo "<li class='parent' onClick=""OriginAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>新增来源</span></li>"
		 .echo "<li class='parent' onclick='OriginControl(1);'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>修改来源</span></li>"
		 .echo "<li class='parent' onclick='OriginControl(2);'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>删除来源</span></li>"
		 .echo "</ul>"
		 .echo "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
		 .echo ("<form name='myform' method='Post' action='?'>")
	     .echo ("<input type='hidden' name='action' id='action' value='" & Action & "'>")
		 .echo "        <tr>"
		 .echo "          <td class=""sort"" width='35' align='center'>选择</td>"
		 .echo "          <td height='25' class='sort' align='center'>来源名称</td>"
		 .echo "          <td class='sort' align='center'>单位名称</td>"
		 .echo "          <td class='sort' align='center'>连接地址</td>"
		 .echo "          <td class='sort' align='center'>添加时间</td>"
		 .echo "        </tr>"
			 Set RS = Server.CreateObject("ADODB.RecordSet")
				   OriginSql = "SELECT * FROM [KS_Origin] Where OriginType=0 order by AddDate desc"
				   RS.Open OriginSql, conn, 1, 1
				 If RS.EOF And RS.BOF Then
				 Else
					       totalPut = RS.RecordCount
		
							If CurrentPage < 1 Then	CurrentPage = 1
		
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
							Call showContent

			End If
		     .echo "</table>"
			 .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
	         .echo ("<tr><td width='180'><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> </div>")
	         .echo ("</td>")
	         .echo ("<td><select style='height:18px' onchange='set(this.value)' name='setattribute'><option value=0>快速选项...</option><option value='1'>执行编辑</option><option value='2'>执行删除</option></select></td>")
	         .echo ("</form><td align='right'>")
	         Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	         .echo ("</td></tr></table>")
		End With
		End Sub
		Sub showContent()
		With KS
		Do While Not RS.EOF
		  .echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & RS("ID") & "' onclick=""chk_iddiv('" & RS("ID") & "')"">"
		  .echo "    <td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" &RS("ID") & "')"" type='checkbox' id='c"& RS("ID") & "' value='" &RS("ID") & "'></td>"
		  .echo "    <td class='splittd' height='19'><span OriginID='" & RS("ID") & "' onDblClick=""EditOrigin(this.OriginID)"">"
		  .echo "     <span  style='cursor:default;'>" & RS("OriginName") & "</span></span>"
		  .echo "   </td>"
		  .echo "   <td class='splittd' align='center'>&nbsp;" & RS("UnitName") & " </td>"
		  .echo "   <td class='splittd' align='center'>" & RS("HomePage") & "</td>"
		  .echo "   <td class='splittd' align='center'>" & RS("AddDate") & " </td>"
		  .echo " </tr>"
				I = I + 1
				If I >= MaxPerPage Then Exit Do
				RS.MoveNext
			 Loop
				RS.Close
         End With
		 End Sub
		 
		 Sub OriginAddOrEdit(OpType)
		 With KS
		  Dim RS, OriginSql
		 ID = KS.G("ID")
		 .echo "<html>"
		 .echo "<head>"
		 .echo "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
		 .echo "<title>来源管理</title>"
		 .echo "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		 .echo "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		 .echo "</head>"
		 .echo "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
		
		 Action="AddSave"
		 HomePage="http://"
		 If Optype = "Edit" Then
			 Set RS = Server.CreateObject("ADODB.RECORDSET")
			 OriginSql = "Select * From [KS_Origin] Where ID='" & ID & "'"
			 RS.Open OriginSql, conn, 1, 1
		 If Not RS.EOF Then
			 OriginName = Trim(RS("OriginName"))
			 Contact = Trim(RS("Contact"))
			 Telphone = Trim(RS("Telphone"))
			 UnitName = Trim(RS("UnitName"))
			 UnitAddress = Trim(RS("UnitAddress"))
			 Zip = Trim(RS("Zip"))
			 Email = Trim(RS("Email"))
			 QQ = Trim(RS("QQ"))
			 HomePage = Trim(RS("HomePage"))
			 Note = Trim(RS("Note"))
			 Action="EditSave"
		 End If
		 
   End If
		
		 .echo "<table width='100%' border='0' align='center' cellpadding='1' cellspacing='1' class='ctable' style='margin-top:5px;border-collapse: collapse'>"
		 .echo "  <form  action='KS.Origin.asp?ID=" & ID &"&page=" & Page &"' method='post' name='OrigArticlerm' onsubmit='return(CheckForm())'>"
		 .echo "    <input type='hidden' value='" & Action &"' name='action'>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td width='200' height='25' align='right' class='clefttitle'>来源名称：</td>"
		 .echo "       <td><input name='OriginName' value='" & OriginName &"' type='text' id='OriginName' class='textbox'></td>"
		 .echo "    </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td height='25' class='clefttitle' align='right' nowrap>联 系 人：</td>"
		 .echo "      <td height='25' nowrap><input name='Contact' value='" & Contact &"' type='text' id='Contact' class='textbox'></td>"
		 .echo "   </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td height='25' align='right' class='clefttitle' nowrap>联系电话：</td>"
		 .echo "      <td height='25' nowrap> <input name='Telphone' value='" & Telphone &"' type='text' id='Telphone' class='textbox'></td>"
		 .echo "    </tr>"
		 .echo "   <tr class='tdbg'>"
		 .echo "      <td height='25' align='right' nowrap class='clefttitle'>单位名称：</td>"
		 .echo "      <td height='25'><input name='UnitName' type='text' value='" & UnitName &"' id='UnitName' class='textbox'></td>"
		 .echo "    </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td height='25' align='right' nowrap class='clefttitle'>单位地址：</td>"
		 .echo "      <td height='25'> <input name='UnitAddress' type='text' value='" & UnitAddress &"' class='textbox'></td>"
		 .echo "    </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td height='25' align='right' nowrap class='clefttitle'>邮政编码：</td>"
		 .echo "      <td height='25' nowrap> <input name='Zip' type='text' value='" & zip &"' id='Zip' class='textbox'></td>"
		 .echo "    </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td align='right' class='clefttitle'>电子邮箱：</td>"
		 .echo "      <td><input name='Email' type='text' id='Email' value='" & Email &"' class='textbox'></td>"
		 .echo "    </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td height='25' align='right' nowrap class='clefttitle'>联系 QQ：</td>"
		 .echo "      <td height='25' nowrap> <input name='QQ' type='text' id='QQ' value='" & QQ &"' class='textbox'></td>"
		 .echo "    </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "       <td height='25' align='right' class='clefttitle'>主页地址：</td>"
		 .echo "       <td><input name='HomePage' type='text' id='HomePage' class='textbox' value='" & HomePage &"'></td>"
		 .echo "    </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td align='right' class='clefttitle'>备注说明：</td>"
		 .echo "      <td height='25' nowrap> <textarea name='Note' cols='50' style='width:250px;height:80px' rows='6' id='Note' class='textbox'>" & Note &"</textarea>"
		 .echo "     </td>"
		 .echo "    </tr>"
		 .echo "    <input type='hidden' name='OriginType' value='0'>"
		 .echo "  </form>"
		 .echo "</table>"
		 
		 .echo "<div id='save'>"
		 .echo "<li class='parent' onclick=""return(CheckForm())""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/save.gif' border='0' align='absmiddle'>确定保存</span></li>"
		 .echo "<li class='parent' onclick=""parent.closeWindow()""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>关闭取消</span></li>"
		 .echo "</div>"

		 .echo "<Script Language='javascript'>"
		 .echo "function CheckForm()"
		 .echo "{ var form=document.OrigArticlerm;"
		 .echo "   if (form.OriginName.value=='')"
		 .echo "    {"
		 .echo "     alert('请输入来源名称!');"
		 .echo "     form.OriginName.focus();"
		 .echo "     return false;"
		 .echo "    }"
		 .echo "   if ((form.Zip.value!='')&&((form.Zip.value.length>6)||(!is_number(form.Zip.value))))"
		 .echo "    {"
		 .echo "     alert('非法邮政编码!');"
		 .echo "     form.Zip.focus();"
		 .echo "     return false;"
		 .echo "    }"
		 .echo "    if ((form.Email.value!='')&&(is_email(form.Email.value)==false))"
		 .echo "    {"
		 .echo "    alert('非法电子邮箱!');"
		 .echo "     form.Email.focus();"
		 .echo "     return false;"
		 .echo "    }"
		 .echo "    form.submit();"
		 .echo "    return true;"
		 .echo "}"
		 .echo "</Script>"
		End With
		End Sub
		 
		 Sub OriginAddSave()
		 Dim RS
		 OriginName = KS.G("OriginName")
		 Contact = KS.G("Contact")
		 Telphone = KS.G("Telphone")
		 UnitName = KS.G("UnitName")
		 UnitAddress = KS.G("UnitAddress")
		 Zip = KS.G("Zip")
		 Email = KS.G("Email")
		 QQ = KS.G("QQ")
		 HomePage = KS.G("HomePage")
		 Note = KS.G("Note")
		 OriginType =KS.G("OriginType")

		 If OriginName = "" Then Call KS.AlertHistory("请输入来源名称!", -1): Exit Sub
		 Set RS = Server.CreateObject("ADODB.RECORDSET")
		 OriginSql = "Select * From [KS_Origin] Where OriginName='" & OriginName & "' And OriginType=0"
		 RS.Open OriginSql, conn, 3, 3
		 If RS.EOF And RS.BOF Then
		  RS.AddNew
		  RS("ID") = Year(Now) & Month(Now) & Day(Now) & KS.MakeRandom(5)
		  RS("OriginName") = OriginName
		  RS("Contact") = Contact
		  RS("Telphone") = Telphone
		  RS("UnitName") = UnitName
		  RS("UnitAddress") = UnitAddress
		  RS("Zip") = Zip
		  RS("Email") = Email
		  RS("QQ") = QQ
		  RS("HomePage") = HomePage
		  RS("Note") = Note
		  RS("OriginType") = OriginType
		  RS("AddDate") = Now()
		  RS.Update
		  KS.Echo ("<Script> if (confirm('来源增加成功,继续添加吗?')) { location.href='KS.Origin.asp?Action=Add';} else{top.frames[""MainFrame""].location.reload();top.closeWindow();}</script>")
		 Else
		   Call KS.AlertHistory("数据库中已存在该来源名称!", -1)
		   Exit Sub
		 End If
		 RS.Close
		 End Sub
		 
		 Sub OriginEditSave()
		 With KS
		 ID = KS.G("ID")
		 OriginName = KS.G("OriginName")
		 Contact = KS.G("Contact")
		 Telphone = KS.G("Telphone")
		 UnitName = KS.G("UnitName")
		 UnitAddress = KS.G("UnitAddress")
		 Zip = KS.G("Zip")
		 Email = KS.G("Email")
		 QQ = KS.G("QQ")
		 HomePage = KS.G("HomePage")
		 Note = KS.G("Note")
		 OriginType =KS.G("OriginType")
		  If OriginName = "" Then Call KS.AlertHistory("请输入来源名称!", -1): Exit Sub
		 Set RS = Server.CreateObject("ADODB.RECORDSET")
		  OriginSql = "Select * From [KS_Origin] Where ID='" & ID & "'"
		  RS.Open OriginSql, conn, 1, 3
		  RS("OriginName") = OriginName
		  RS("Contact") = Contact
		  RS("Telphone") = Telphone
		  RS("UnitName") = UnitName
		  RS("UnitAddress") = UnitAddress
		  RS("Zip") = Zip
		  RS("Email") = Email
		  RS("QQ") = QQ
		  RS("HomePage") = HomePage
		  RS("Note") = Note
		  RS.Update
		  RS.Close
		  KS.Echo ("<Script> alert('来源修改成功!');top.frames[""MainFrame""].location.reload();top.closeWindow();</script>")
		 End With
		 End Sub
		 
		 Sub OriginDel()
			Dim ID:ID = KS.G("ID")
			ID = Replace(ID, ",", "','")
			ID = "'" & ID & "'"
			conn.Execute ("Delete From KS_Origin Where ID IN(" & ID & ")")
			Response.Redirect "KS.Origin.asp?Page=" & Page
		 End Sub
End Class
%> 
