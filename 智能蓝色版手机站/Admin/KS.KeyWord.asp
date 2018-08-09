<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_KeyWord
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_KeyWord
        Private KS,Action,ComeUrl,Page,IsSearch
		Private I,totalPut,CurrentPage,KeySql,RS,MaxPerPage,KSCls
		Private Sub Class_Initialize()
		  MaxPerPage =20
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
		    If KS.S("Action")<>"" Then
			.echo "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd"">"
			.echo "<html xmlns=""http://www.w3.org/1999/xhtml"">"
			else
			 .echo "<html>"
			end if
			.echo "<head>"
			.echo "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
			.echo "<title>关键字管理</title>"
			.echo "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
            .echo "<script language='JavaScript'>"
			.echo "var Page='" & CurrentPage & "';"
			.echo "</script>"
			.echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
			.echo "<script language=""JavaScript"" src=""../KS_Inc/Jquery.js""></script>" & vbCrLf
			
             Action=KS.G("Action")
			 IsSearch=KS.ChkClng(KS.G("IsSearch"))
			 If IsSearch="1" Then
				If Not KS.ReturnPowerResult(0, "KMST10019") Then                  '权限检查
					Call KS.ReturnErr(1, "")   
					Response.End()
				End if
			 Else
				If Not KS.ReturnPowerResult(0, "M010004") Then                  '权限检查
					Call KS.ReturnErr(1, "")   
					Response.End()
				End if
			 End If
			 
			 Page=KS.G("Page")
			 Select Case Action
			  Case "Add","Edit"
			   Call KeyWordAddOrEdit()
			  Case "Del"  Call KeyWordDel()
			  Case "DelAllRecord" DelAllRecord
			  Case "DoSave"
			    Call DoSave()
			  Case Else
			   Call KeyWordList()
			 End Select
			.echo "</body>"
			.echo "</html>"
		  End With
		End Sub
			  
			 
		Sub KeyWordList()
		     CurrentPage=KS.ChkClng(Request("page"))
			 If CurrentPage<=0 Then CurrentPage=1			
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
				function KeyWordAdd()
				{
					new parent.KesionPopup().PopupCenterIframe('新增关键字Tags','KS.KeyWord.asp?Action=Add',530,110,'no')
				}
				function EditKeyWord(id)
				{	
				   new parent.KesionPopup().PopupCenterIframe('修改关键字Tags','KS.KeyWord.asp?Page='+Page+'&Action=Edit&ID='+id,530,110,'no')

				}
				function DelKeyWord(id)
				{
				if (confirm('真的要删除该关键字吗?'))
				 location="KS.KeyWord.asp?IsSearch=<%=IsSearch%>&Action=Del&Page="+Page+"&id="+id;
				 SelectedFile='';
				}
				function KeyWordControl(op)
				{  var alertmsg='';
	               var ids=get_Ids(document.myform);
					if (ids!='')
					 {  if (op==1)
						{
						if (ids.indexOf(',')==-1) 
							EditKeyWord(ids)
						  else alert('一次只能编辑一个关键字tags!')	 
						}	
					  else if (op==2)    
						 DelKeyWord(ids);
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
					 alert('请选择要'+alertmsg+'的关键字');
					  }
				}
				function GetKeyDown()
				{ 
				if (event.ctrlKey)
				  switch  (event.keyCode)
				  {  case 90 : location.reload(); break;
					 case 65 : Select(0);break;
					 case 78 : event.keyCode=0;event.returnValue=false; KeyWordAdd();break;
					 case 69 : event.keyCode=0;event.returnValue=false;KeyWordControl(1);break;
					 case 68 : KeyWordControl(2);break;
				   }	
				else	
				 if (event.keyCode==46)KeyWordControl(2);
				}
			</script>
			<%
			.echo "<body topmargin='0' leftmargin='0' onkeydown='GetKeyDown();''>"
			.echo "<ul id='menu_top'>"
		If IsSearch="1" Then
			.echo "<br/><strong>搜索关键词维护</strong> <a href='?issearch=1&order=1'>按搜索次数最多查看</a> | <a href='?issearch=1&order=2'>按搜索次数最少查看</a>"
		Else
			.echo "<li class='parent' onClick=""KeyWordAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加关键字</span></li>"
			.echo "<li class='parent' onClick=""KeyWordControl(1);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>编辑关键字</span></li>"
			.echo "<li class='parent' onClick=""KeyWordControl(2);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>删除关键字</span></li>"
		End If
			.echo "</ul>"
			.echo "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
			.echo ("<form name='myform' method='Post' action='KS.KeyWord.asp'>")
	        .echo ("<input type='hidden' name='action' id='action' value='" & Action & "'>")
			.echo "        <tr>"
		If IsSearch="1" Then
			.echo "          <td class=""sort"" width='35' align='center'>选择</td>"
			.echo "          <td class='sort' align='center'>关键词</td>"
			.echo "          <td width='19%' class='sort' align='center'>搜索频率(次)</td>"
			.echo "          <td width='23%' align='center' class='sort'>最后搜索时间</td>"
			.echo "          <td width='26%' class='sort' align='center'>第一次搜索时间</td>"
		Else
			.echo "          <td class=""sort"" width='35' align='center'>选择</td>"
			.echo "          <td class='sort' align='center'>关键字Tags</td>"
			.echo "          <td width='19%' class='sort' align='center'>使用频率(次)</td>"
			.echo "          <td width='23%' align='center' class='sort'>最后使用时间</td>"
			.echo "          <td width='26%' class='sort' align='center'>添加时间</td>"
		End If
			.echo "  </tr>"
			  
			  Dim Order
			  If Request("Order")="1" Then
			   Order="hits desc,ID Desc"
			  ElseIf Request("Order")="2" Then
			   Order="hits asc,ID Desc"
			  Else
			   Order="AddDate desc,ID Desc"
			  End If
			  
			  Set RS = Server.CreateObject("ADODB.RecordSet")
					   KeySql = "SELECT * FROM [KS_KeyWords] where IsSearch=" & IsSearch & " order by "&Order
					   RS.Open KeySql, conn, 1, 1
					 If RS.EOF And RS.BOF Then
					 Else
								totalPut = RS.RecordCount
								If CurrentPage >1 And (CurrentPage - 1) * MaxPerPage < totalPut Then
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
	        .echo ("<td><select style='height:18px' onchange='set(this.value)' name='setattribute'><option value=0>快速选项...</option><option value='1'>编辑关键字</option><option value='2'>执行删除</option></select></td>")
	        .echo ("</form><td align='right'>")
			Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	        .echo ("</td></tr></table>")
			
			if IsSearch=1 Then
		      .echo ("<form action='KS.KeyWord.asp?action=DelAllRecord&issearch=" & IsSearch & "' method='post' target='_hiddenframe'>")
		      .echo ("<iframe src='about:blank' style='display:none' name='_hiddenframe' id='_hiddenframe'></iframe>")
			  .echo ("<div class='attention'><strong>特别提醒： </strong><br>当站点运行一段时间后,网站的搜索记录表可能存放着大量的记录,为使系统的运行性能更佳,建议一段时间后清理一次。")
		      .echo ("<br /> <strong>删除范围：</strong><br/><label><input onclick=""$('#s1').show();$('#s2').hide();"" name=""deltype"" type=""radio"" value=""1"" checked=""checked""/> 按搜索次数</label><label><input onclick=""$('#s2').show();$('#s1').hide();"" name=""deltype"" type=""radio"" value=""2"" /> 按日期</label> <div id='s1'>搜索次数小于<input name=""searchnum"" size='4' value=10 style='text-align:center'>次</div><div id='s2' style='display:none'>最近<input type='text' name='days' value='100' size='4' style='text-align:center'>天内没有被重新搜索</div><input onclick=""$(parent.frames['FrameTop'].document).find('#ajaxmsg').toggle();"" type=""submit""  class=""button"" value=""执行删除"">")
		     .echo ("</div>")
			 .echo ("</form>")
			End If

            End With
			End Sub
			
			Sub showContent()
			   With KS
					Do While Not RS.EOF
			          .echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & RS("ID") & "' onclick=""chk_iddiv('" & RS("ID") & "')"">"
			          .echo "<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" &RS("ID") & "')"" type='checkbox' id='c"& RS("ID") & "' value='" &RS("ID") & "'></td>"
					  .echo "<td class='splittd' height='20'><span KeyWordID='" & RS("ID") & "' ondblclick=""EditKeyWord(this.KeyWordID)""><img src='Images/key.gif' align='absmiddle'>"
					  .echo "  <span style='cursor:default;'>" & RS("KeyText") & "</span></span></td>"
					  .echo "  <td class='splittd' align='center'>" & RS("Hits") & " </td>"
					  .echo "  <td class='splittd' align='center'><FONT Color=red>" & RS("lastusetime") & "</font> </td>"
					  .echo "  <td class='splittd' align='center'>" & RS("AddDate") & " </td>"
					  .echo "</tr>"
					  I = I + 1
					  If I >= MaxPerPage Then Exit Do
						   RS.MoveNext
					Loop
					  RS.Close
				End With
			End Sub

			
			Sub KeyWordAddOrEdit()
			With KS
			 Dim Action, KeyWordText, ID, PageRS, KeySql,Page
			  ID = KS.ChkClng(KS.G("ID"))
			  Page = KS.G("Page")
			  If ID<>0 Then
				 Set RS = Server.CreateObject("ADODB.RECORDSET")
				 KeySql = "Select * From [KS_KeyWords] Where ID=" & ID
				 RS.Open KeySql, conn, 1, 1
				 If Not RS.EOF Then KeyWordText = RS("KeyText")
				 RS.Close:Set RS=Nothing
			  End If
			
			.echo "<br>" & vbCrLf
			.echo "<table width='45%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
			.echo "  <form  action='KS.KeyWord.asp' method='post' name='KeyWordForm'>" & vbCrLf
			.echo "    <tr>" & vbCrLf
			.echo "      <td height='10' colspan='3' align='right'>&nbsp;</td>" & vbCrLf
			.echo "    </tr>" & vbCrLf
			.echo "    <tr>" & vbCrLf
			.echo "      <td width='23%' align='right' nowrap>关键字Tags：</td>" & vbCrLf
			.echo "      <td width='20' height='30' align='right' nowrap> <div align='center'></div></td>" & vbCrLf
			.echo "      <td width='74%' height='30' nowrap><b>" & vbCrLf
			.echo "        <input name='KeyWordText' class='textbox' type='text' onload='this.focus()' id='KeyWordText' value='" & KeyWordText & "'>" & vbCrLf
			.echo "        </b>* </td>" & vbCrLf
			.echo "    </tr>" & vbCrLf
			.echo "    <input type='hidden' value='DoSave' name='Action'>" & vbCrLf
			.echo "    <input type='hidden' value='" & ID & "' name='ID'>" & vbCrLf
			.echo "    <input type='hidden' value='" & Page & "' name='Page'>" & vbCrLf
			.echo "  </form>" & vbCrLf
			.echo "</table>" & vbCrLf

		 .echo "<div id='save'>"
		 .echo "<li class='parent' onclick=""return(CheckForm())""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/save.gif' border='0' align='absmiddle'>确定保存</span></li>"
		 .echo "<li class='parent' onclick=""parent.closeWindow()""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>关闭取消</span></li>"
		 .echo "</div>"

			
			.echo "<Script Language='javascript'>" & vbCrLf
			.echo "<!--" & vbCrLf
			.echo "function CheckForm()" & vbCrLf
			.echo "{ " &vbcrlf
			.echo "   if ($('#KeyWordText').val()=='')" & vbCrLf
			.echo "    {" & vbCrLf
			.echo "     alert('请输入关键字!');" & vbCrLf
			.echo "     $('#KeyWordText').focus();" & vbCrLf
			.echo "     return false;" & vbCrLf
			.echo "    }" & vbCrLf
			.echo "    $('form[name=KeyWordForm]').submit();" & vbCrLf
			.echo "}" & vbCrLf
			.echo "//-->" & vbCrLf
			.echo "</Script>" & vbCrLf
			.echo "</body>" & vbCrLf
			.echo "</html>" & vbCrLf
			End With
			End Sub
			
			Sub DoSave()
			    Dim RS,ID:ID=KS.ChkClng(KS.G("ID"))
			    Dim KeyWordText:KeyWordText = KS.G("KeyWordText")
				If KeyWordText = "" Then
				   Call KS.AlertHistory("请输入关键字!", -1)
				 End If
				 Set RS = Server.CreateObject("ADODB.RECORDSET")
				 If ID=0 Then
				  If Not Conn.Execute("Select top 1 * From [KS_KeyWords] Where KeyText='" & KeyWordText & "'").Eof Then
				   KS.AlertHintScript "数据库中已存在该关键字!"
				  End If
				 End If
				 KeySql = "Select top 1 * From [KS_KeyWords] Where ID=" & ID
				 RS.Open KeySql, conn, 1, 3
				 If RS.EOF And RS.BOF Then
				  RS.AddNew
				  RS("AddDate") = Now()
				  RS("LastUseTime")=Now()
				  RS("Hits")=0
				  Rs("IsSearch")=0
				 End If
				  RS("KeyText") = KeyWordText
				  RS.Update
				 RS.Close:Set RS=Nothing
				 If ID=0 Then
				  KS.Echo ("<Script> if (confirm('关键字增加成功,继续添加吗?')) { location.href='?Action=Add';} else{top.frames[""MainFrame""].location.reload();parent.closeWindow();}</script>")
				 Else
				  KS.Echo ("<Script> alert('关键字修改成功!');top.frames[""MainFrame""].location.reload();parent.closeWindow();</script>")
				 End If
			End Sub
		
			
			Sub KeyWordDel()
			  Dim ID,Page
			  Page=KS.G("Page")
			  Dim RS:Set RS=Server.CreateObject("ADODB.Recordset")
			  ID = KS.G("ID")
			  ID = Replace(ID, " ", "")
			  RS.Open "Delete from [KS_KeyWords] Where ID in(" & ID & ")", conn, 3, 3
			  Set RS = Nothing
			  Response.Redirect "?IsSearch="&IsSearch&"&Page=" & Page
			End Sub
			
			Sub DelAllRecord()
			  Dim DelType,SQL
			  DelType=KS.ChkClng(Request("DelType"))
			  If DelType=1 Then
			    SQL="Delete From KS_KeyWords Where IsSearch=1 And Hits<=" & KS.ChkClng(Request("SearchNum"))
			  ElseIf DelType=2 Then
			    SQL="Delete From KS_KeyWords Where IsSearch=1 And datediff(" & DataPart_D & ",LastUseTime," & SqlNowString & ")>" & KS.ChkClng(Request("SearchNum"))
			  Else 
			    Exit Sub
			  End If
			  Conn.Execute(SQL)
               KS.echo "<script>$(parent.document).find('#ajaxmsg').toggle();alert('恭喜,删除指定条件的记录成功!');</script>"
			End Sub
End Class
%>
 
