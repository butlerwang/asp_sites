<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%> 
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Field
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Field
        Private KS,Action,ChannelID,Page,ItemName,TableName,KSCls
		Private I, totalPut, CurrentPage, FieldSql, FieldRS,MaxPerPage
		Private FieldName,ID,Contact, Title, Tips, FieldType, DefaultValue, MustFillTF, ShowOnForm, ShowOnUserForm,Options,OrderID,AllowFileExt,MaxFileSize,Width

		Private Sub Class_Initialize()
		  MaxPerPage = 30
		  Set KSCls=New ManageCls
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
		With Response
		.Write "<html>"
		.Write "<head>"
		.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
		.Write "<title>字段管理</title>"
		.Write "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
             Action=KS.G("Action")
			 ChannelID=KS.ChkClng(KS.G("ChannelID"))
			 
			 TableName=KS.C_S(ChannelID,2)
			 If ChannelID=101 Then TableName="KS_User"   '会员表
			 ItemName=KS.C_S(ChannelID,3)
			 Page=KS.G("Page")
			 
			If Not KS.ReturnPowerResult(0, "M010008") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
			End if
			 
			 Select Case Action
			  Case "SetCollect"
			    Call FieldSetCollect()
			  Case Else
			   Call FieldList()
			 End Select
			.Write "</body>"
			.Write "</html>"
		 End With
		End Sub
		
		Sub FieldList()
		 On Error Resume Next
		 CurrentPage = KS.ChkClng(KS.G("page"))
		 If CurrentPage<1 Then CurrentPage=1
		With Response
		.Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		.Write "</head>"
		.Write "<body scroll=no topmargin='0' leftmargin='0'>"
		.Write "<ul id='menu_top'>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_ItemModify.asp?channelid=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>新建项目</span></li>"
		.Write "<li class='parent' onclick='location.href=""Collect_ItemFilters.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/move.gif' border='0' align='absmiddle'>过滤设置</span></li>"
		.Write "<li class='parent' onclick='location.href=""Collect_IntoDatabase.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/save.gif' border='0' align='absmiddle'>审核入库</span></li>"
		.Write "<li class='parent' onclick='location.href=""Collect_ItemHistory.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/Recycl.gif' border='0' align='absmiddle'>历史记录</span></li>"
		.Write "<li class='parent' onclick='location.href=""Collect_Field.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/addjs.gif' border='0' align='absmiddle'>自定义字段</span></li>"
		.Write "<li class='parent' onclick='location.href=""Collect_main.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/back.gif' border='0' align='absmiddle'>回上一级</span></li>"
		.Write "</ul>"
        
		.Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
		
		.Write "<div style='text-align:right'>请按模型设置要启用的自定义字段采集<select id='channelid' name='channelid' onchange=""if (this.value!=0){location.href='?channelid='+this.value;}"">"
			.Write " <option value='0'>---请选择模型---</option>"
	
			If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,Node
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1][@ks6=1||@ks6=2||@ks6=5]")
			  If trim(ChannelID)=trim(Node.SelectSinglenode("@ks0").text) Then
			   .Write "<option value='" &Node.SelectSingleNode("@ks0").text &"' selected>" & Node.SelectSingleNode("@ks1").text & "</option>"

			  Else
			   .Write "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
			  End If
			next
			.Write "</select></div>"
		
		.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		.Write "<form action='Collect_Field.asp?action=SetCollect&channelid=" & ChannelID&"&page="&CurrentPage &"'' name='form1' method='post'>"
		.Write "        <tr class='sort'>"
		.Write "         <td width='80' align='center'>启用采集</td>"
		.Write "          <td width='100' align='center'>字段名称</td>"
		.Write "          <td align='center'>字段别名</td>"		
		.Write "          <td align='center'>归属模型</td>"
		.Write "          <td align='center'>字段类型</td>"
		.Write "          <td align='center'>是否启用采集</td>"
		.Write "          <td align='center'>出现位置</td>"
		.Write "        </tr>"
			 Set FieldRS = Server.CreateObject("ADODB.RecordSet")
				   FieldSql = "SELECT * FROM KS_Field Where fieldtype<>0 and ChannelID=" & ChannelID & " order by orderid asc"
				   FieldRS.Open FieldSql, conn, 1, 1
				 If FieldRS.EOF And FieldRS.BOF Then
				  .Write "<tr><td height='30' class='splittd' style='text-align:center' colspan=10>该模型没有自定义字段!</td></tr>"
				 Else
					        totalPut = FieldRS.RecordCount
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
								FieldRS.Move (CurrentPage - 1) * MaxPerPage
							End If
							Call showContent
			End If
		 .Write " <tr>"
		 .Write "   <td colspan='3'>&nbsp;&nbsp;<input type='submit' class='button' value='批量保存字段采集设置'> </td></form>"
		 .Write "   <td height='35' colspan='4' align='right'>"
		 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		.Write "    </td>"
		.Write " </tr>"
		.Write "</table>"
		.Write "</div>"
		End With
		End Sub
		Sub showContent()
		Dim CollectTF,ShowType
		With Response
		Do While Not FieldRS.EOF
		  Dim RS:Set RS=KS.ConnItem.Execute("Select * From KS_FieldItem where fieldid=" & FieldRS("FieldID"))
		  IF RS.Eof Then
		   CollectTF=false
		   ShowType=0
		  Else
		   ShowType=RS("ShowType")
		   CollectTF=true
		  End If
		  
		  RS.Close:Set RS=Nothing
		 .Write "<tr>"
		 .Write "<td class='splittd' align='center'>&nbsp;&nbsp;"
		 If CollectTF=True Then
		 .Write "<input type='checkbox' name='CField"& FieldRS("FieldID")&"' value='1' checked>"
		 Else
		 .Write "<input type='checkbox' name='CField"& FieldRS("FieldID")&"' value='1'>"
		 End iF
		 .Write "<input type='hidden' name='FieldID' value='" & FieldRS("FieldID") & "'>" & FieldRS("FieldID") & "</td>"
		 .Write "  <td class='splittd'><img src='../Images/Field.gif' align='absmiddle'><span  style='cursor:default;'>" & FieldRS("FieldName") & "</td>"
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
				 End Select
		  If Left(Lcase(FieldRS("FieldName")),3)<>"ks_" Then .Write "<font color=#cccccc>[系统]</font>"
		 .Write "</td>"
		 .Write "   <td align='center' class='splittd'>&nbsp;" 
		 If CollectTF=false Then
		  .Write "未启用"
		 Else
		  .Write "<font color=red>已启用</font>"
		 End If
		 .Write " </td>"
		 
		 '========================增加列表采集开关=================
		 .Write "   <td align='center' class='splittd'>" 
		 .Write "<input type='radio' value='1' name='ShowType"& FieldRS("FieldID")&"'"
		 If ShowType=1 Then .Write " Checked"
		 .Write ">列表页"
		 .Write "<input type='radio' value='0' name='ShowType"& FieldRS("FieldID")&"'"
		 If ShowType=0 Then .Write " Checked"
		 .Write ">内容页"
		 .Write " </td>"
		 '================================================================
		 
		 .Write " </tr>"
								I = I + 1
								If I >= MaxPerPage Then Exit Do
							   FieldRS.MoveNext
							   Loop
								FieldRS.Close
						 
         End With
		 End Sub
		 
		 
		 
		 Sub FieldSetCollect()
			  Dim FieldID:FieldID=KS.G("FieldID")
			  Dim I,FieldIDArr,AllowStr
			  FieldIDArr=Split(FieldID,",")
			  AllowStr=KS.G("CField")
			  For I=0 To Ubound(FieldIDArr)
			   If KS.G("CField" & trim(FieldIDArr(i)))="1" Then
			      Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
				   RS.Open "Select * From KS_FieldItem Where FieldID=" & FieldIDArr(i),KS.ConnItem,1,3
				   If RS.Eof Then 
					RS.AddNew
				   End If
					RS("FieldID")=FieldIDArr(i)
					'==============列表采集开关================================
					RS("ShowType")=KS.ChkClng(KS.G("ShowType"&trim(FieldIDArr(i))))
					'==========================================================
					RS("ChannelID")=KS.ChkClng(KS.G("ChannelID"))
					RS("FieldName")=Conn.Execute("Select FieldName From KS_Field Where FieldID=" & FieldIDArr(i))(0)
					RS("FieldTitle")=Conn.Execute("Select Title From KS_Field Where FieldID=" & FieldIDArr(i))(0)
					RS("OrderID")=Conn.Execute("Select OrderID From KS_Field Where FieldID=" & FieldIDArr(i))(0)
				   RS.Update
				   KS.ConnItem.Execute("Update KS_FieldRules Set ShowType=" & rs("ShowType") & ",ChannelID=" & RS("channelid") & ",FieldName='" & RS("FieldName") & "' Where FieldID=" & FieldIDArr(i))
				   RS.Close:Set RS=Nothing
			   Else
			     KS.ConnItem.Execute("Delete From KS_FieldItem Where FieldID=" & FieldIDArr(i))
			   End If
			  
			  Next
			  Response.Write "<script>alert('批量保存字段成功！');location.href='?ChannelID=" & ChannelID & "&Page=" & Page&"';</script>"
		 End Sub
End Class
%> 
