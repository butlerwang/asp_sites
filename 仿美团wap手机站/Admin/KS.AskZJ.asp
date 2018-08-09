<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_AskZJ
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_AskZJ
        Private KS,KSCls
		Private maxperpage,CurrentPage,TotalPut,SqlParam
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		  maxperpage=20
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		%>
		<html>
		<head>
		<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
		<script src="../KS_Inc/common.js" language="JavaScript"></script>
		<script src="../KS_Inc/jquery.js" language="JavaScript"></script>
		</head>
		<body>
        <div class='topdashed sort'>问答认证专家管理</div>
		<%
		     If Not KS.ReturnPowerResult(0, "WDXT10005") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 response.End
			 End If
				 
		    CurrentPage=KS.ChkClng(Request("page"))
			If CurrentPage<=0 Then CurrentPage=1
			Dim Action
			Action = LCase(Request("action"))
			Select Case Trim(Action)
			Case "del" Call delZJ()
			Case "verify" Call verifyZJ()
			Case "unverify" Call unverifyZJ()
			Case "modify" Call Modify()
			Case "modifysave" Call ModifySave()
			Case "recommend" Call Recommend()
			Case Else
				Call showmain()
			End Select
	   End Sub

		Sub showmain()
		%>
		<div style="margin-top:5px;height:25px;line-height:25px">
		<b>查看：</b> <a href="KS.AskZJ.asp"><font color=#999999>全部</font></a> - <a href="?status=0"><font color=#999999>未审核</font></a> - <a href="?status=1"><font color=#999999>已审核</font></a> 
		</div>
		<table  border="0" align="center" style='border-top:1px solid #cccccc' cellpadding="0" cellspacing="0" width="100%">
		<tr class="sort">
			<td width="5%" noWrap="noWrap">选择</td>
			<td>姓名</td>
			<td>用户名</td>
			<td>电话/手机</td>
			<td>回答数</td>
			<td>申请时间</td>
			<td>照片</td>
			<td>身份证</td>
			<td>执业证</td>
			<td>推荐</td>
			<td>状态</td>
			<td>管理操作</td>
		</tr>
		
		<form name="myform" id="myform" method="post" action="?">
		<input type="hidden" name="action" id="action" value="del">
		<%
		    SqlParam="1=1"
			if request("status")<>"" then
			  SqlParam=" status=" & KS.ChkClng(KS.G("status"))
			end if
			Dim SQLStr,i,RS:SET RS=Server.CreateObject("ADODB.RECORDSET")
			SQLStr=KS.GetPageSQL("KS_AskZJ","ID",MaxPerPage,CurrentPage,1,SqlParam,"*")
			RS.Open SQLStr,Conn,1,1
			If RS.Eof AND RS.Bof Then
			  Response.Write "<tr><td class='splittd' colspan=6 align='center'>对不起, 还没有用户申请认证问答专家!</td></tr>"
		   Else
              totalPut = conn.execute("select count(1) from ks_askzj where " & sqlParam)(0)
			   i=0
			  Do While Not RS.Eof
		   %>
		<tr align="center" onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='u<%=rs("id")%>'>
			<td class="splittd"><input type="checkbox" name="ID" id='c<%=rs("id")%>' value="<%=rs("id")%>"/></td>
			<td class="splittd" align="left"><%=RS("RealName")%></td>
			<td class="splittd"><%=RS("UserName")%></td>
			<td class="splittd"><%=RS("Tel")%></td>
			<td class="splittd"><%=RS("AskDoneNum")%></td>
			<td class="splittd"><%=formatdatetime(rs("adddate"),2)%></td>
			<td class="splittd"><%
			if not ks.isnul(rs("userface")) then 
			  response.write "<a href='" & rs("userface") & "' target='_blank'><img border='0' src='" & rs("userface") & "' width='40' height='40'/></a>"
			else
			  response.write "---"
			end if%></td>
			<td class="splittd"><%
			if not ks.isnul(rs("idcard")) then 
			  response.write "<a href='" & rs("idcard") & "' target='_blank'><img border='0' src='" & rs("idcard") & "' width='40' height='40'/></a>"
			else
			  response.write "---"
			end if%></td>
			<td class="splittd"><%
			if not ks.isnul(rs("ryz")) then 
			  response.write "<a href='" & rs("ryz") & "' target='_blank'><img border='0' src='" & rs("ryz") & "' width='40' height='40'/></a>"
			else
			  response.write "---"
			end if%></td>
			<td class="splittd"><%if rs("recommend")="1" then
			 response.write "<a href=""?action=Recommend&id=" & rs("id") & "&v=0""><font color=blue>√</font></a>"
			 else
			  response.write "<a href=""?action=Recommend&id=" & rs("id") & "&v=1""><font color=red>X</font></a>"
			 end if%>
			</td>
			<td class="splittd"><%if rs("status")="1" then 
			response.write "<a href='?action=unverify&id=" & rs("id") & "'><font color=blue>已审核</font></a>"
			else
			 response.write "<a href='?action=verify&id=" & rs("id") & "'><font color=red>未审核</font></a>"
			end if
			%></td>
			<td class="splittd" noWrap="noWrap"><a href="?action=modify&ID=<%=rs("id")%>">查看编辑</a> | <a href="?action=del&ID=<%=rs("id")%>" onClick="return confirm('删除后将不能恢复，您确定要删除吗?')">删除</a></td>
		</tr>
		<%   i=i+1
		     if i>=MaxPerPage Then Exit Do
			 RS.MoveNext
			Loop
	     End If
		%>
		<tr>
			<td colspan="10">
			&nbsp;&nbsp;<label><input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选</label>
		
				<input class="button" type="submit" name="submit_button1" value="批量删除" onClick="$('action').value='del';return confirm('您确定执行该操作吗?');">
				<input type="submit" value="批量审核" class="button" onClick="$('#action').val('verify');return(confirm('确定批量审核吗?'));">
				
				<input type="submit" value="取消审核" class="button" onClick="$('#action').val('unverify');return(confirm('确定批量取消审核吗?'));">
				
			</td>
		</tr>
		</form>
		<tr>
			<td  align="right" colspan="11" id="NextPageText">
			<%
			Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			%>
			</td>
		</tr>
		
		
		
		</table>
		<%
		End Sub
		
		Sub DelZJ()
		  Dim ID:ID=KS.FilterIds(KS.S("ID"))
		  If Id="" Then
		   KS.AlertHintScript "对不起，您没有选择要删除的记录!"
		  End If
		  Conn.Execute("Delete From KS_AskZJ Where ID in(" & ID & ")")
		  KS.AlertHintScript "恭喜，删除成功!"
		End Sub
		
		Sub verifyZJ()
		  Dim ID:ID=KS.FilterIds(KS.S("ID"))
		  If Id="" Then
		   KS.AlertHintScript "对不起，您没有选择要审核的记录!"
		  End If
		  Conn.Execute("Update KS_AskZJ Set Status=1 Where ID in(" & ID & ")")
		  KS.AlertHintScript "恭喜，审核成功!"
		End Sub
		Sub Recommend()
		  Dim ID:ID=KS.FilterIds(KS.S("ID"))
		  If Id="" Then
		   KS.AlertHintScript "对不起，您没有选择要设置的记录!"
		  End If
		  Conn.Execute("Update KS_AskZJ Set recommend=" & KS.ChkClng(KS.G("V")) & " Where ID in(" & ID & ")")
		  KS.AlertHintScript "恭喜，设置成功!"
		End Sub
		Sub unverifyZJ()
		  Dim ID:ID=KS.FilterIds(KS.S("ID"))
		  If Id="" Then
		   KS.AlertHintScript "对不起，您没有选择要取消审核的记录!"
		  End If
		  Conn.Execute("Update KS_AskZJ Set Status=0 Where ID in(" & ID & ")")
		  KS.AlertHintScript "恭喜，取消审核成功!"
		End Sub
		
		Sub Modify()
		 Dim RS,ID:ID=KS.ChkClng(Request("ID"))
		 If ID=0 Then KS.Die "error!"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "select top 1 * From KS_AskZJ Where ID=" & ID,conn,1,1
		 If RS.Eof And RS.Bof Then
		    RS.Close:Set RS=Nothing
			KS.Die "error!"
		 End If
		%>
		<table  cellspacing="1" cellpadding="3"  width="100%" class="ctable" align="center" border="0">
					  <form action="?Action=modifysave" method="post" name="myform" id="myform">
					    <input type="hidden" value="<%=id%>" name="id"/>
                         <tr class="tdbg">
                            <td width="13%"  class="clefttitle" style='text-align:right'><span style="color: red">* </span> 真实姓名：</td>
                            <td width="27%"><input name="RealName" class="textbox" type="text" id="RealName" value="<%=rs("RealName")%>" size="30" maxlength="50" /></td>
                            <td width="15%"  class="clefttitle" style='text-align:right'><span style="color: red">* </span> 出生年月：</td>
                           <td width="45%"> <%
								%><input type="text" name="birthday" id="birthday" class="textbox" value="<%=rs("Birthday")%>"/></td>
						 </tr>
                         <tr class="tdbg">
                            <td  class="clefttitle" style='text-align:right'>QQ号码：</td>
                            <td><input name="QQ" class="textbox" type="text" id="QQ" value="<%=rs("QQ")%>" size="30" maxlength="50" />
                              </td>
                            <td  class="clefttitle" style='text-align:right'>MSN：</td>
                            <td> <input type="text" name="msn" class="textbox" value="<%=rs("MSN")%>"/></td>
						 </tr>
                         <tr class="tdbg">
                            <td  class="clefttitle" style='text-align:right'>用户照片：</td>
                            <td id="newPreview">
							<%if ks.isnul(rs("UserFace")) then%>
							<%else%>
							<img src="<%=rs("UserFace")%>" width="120" height="120"/>
							<%end if%>
							 	<input type="text" name="UserFace" size="40" value="<%=rs("UserFace")%>"  class="textbox">
							
							 </td>
                            <td  class="clefttitle" style='text-align:right'>性别：</td>
                            <td> <input name="Sex" type="radio" value="男" <% if rs("Sex")="男" then response.write " checked"%>/>男  <input name="Sex" type="radio" value="女" <% if rs("Sex")="女" then response.write " checked"%>/>女</td>
						 </tr>
						 <tr class="tdbg">
                            <td  class="clefttitle" style='text-align:right'><span style="color: red">* </span>电话/手机：</td>
                            <td><input name="Tel" class="textbox" type="text" id="Tel" value="<%=rs("Tel")%>" size="30" maxlength="50" />
                              </td>
                            <td  class="clefttitle" style='text-align:right'>城市：</td>
                            <td> <script src="../plus/area.asp" language="javascript"></script>
							<script language="javascript">
							  <%if rs("Province")<>"" then%>
							  $('#Province').val('<%=rs("province")%>');
								  <%end if%>
							  <%if rs("City")<>"" Then%>
							  $('#City')[0].options[1]=new Option('<%=rs("City")%>','<%=rs("City")%>');
							  $('#City')[0].options(1).selected=true;
							  <%end if%>
							</script></td>
						 </tr>
						 <tr class="tdbg">
						    <td  class="clefttitle" style='text-align:right'><span style="color: red">* </span>擅长分类：</td>
                            <td><input name="SCFL" class="textbox" type="text" id="SCFL" value="<%=rs("SCFL")%>" size="20" maxlength="50" /><span class="msgtips">如：内分泌科</span>
                              </td>
							  <td class="clefttitle" style='text-align:right'>问答分类：</td>
							<td><script src="../<%=KS.ASetting(1)%>category.asp?classid=<%=rs("BigClassID")%>&smallclassid=<%=rs("SmallClassID")%>&SmallerClassID=<%=rs("SmallerClassID")%>" language="javascript"></script>
							
							</td>
                          </tr>
						  
                        
                          
                          
                          <tr class="tdbg">
                            <td  class="clefttitle" style='text-align:right'>所在单位：</td>
                            <td><input name="DanWei" class="textbox" type="text" id="DanWei" value="<%=rs("DanWei")%>" size="30" maxlength="50" /></td>
							<td  class="clefttitle" style='text-align:right'>认证分类：</td>
                            <td><select name="TypeName">
							<option value='0'>--选择认证分类--</option>
							<%
							 dim ii,TypeArr
							 If Not KS.IsNul(KS.ASetting(48)) Then
							 TypeArr=Split(KS.ASetting(48),vbcrlf)
							 for ii=0 to Ubound(TypeArr)
							   IF Trim(rs("TypeName"))=Trim(TypeArr(ii)) Then
							   response.write "<option selected>" & typeArr(ii) & "</option>"
							   Else
							   response.write "<option>" & typeArr(ii) & "</option>"
							   End If
							 next
							 End If
							%>
							</select>
                            </td>
                          </tr>
						  
						  <tr class="tdbg">
						    <td  class="clefttitle" style='text-align:right'>身份证图片：</td>
                            <td colspan=3><input name="IDCard" class="textbox" type="text" id="IDCard" value="<%=rs("IDCard")%>" size=40 /> <span style="color: red">* </span><%if Not KS.IsNul(rs("IDCard")) Then
							 response.write "已上传,<a style='color:red' href='" & rs("IDCard") &"' target='_blank'>浏览</a>"
							end if%>
                            </td>
                          </tr>
						  <tr class="tdbg">
						    <td  class="clefttitle" style='text-align:right'>执业证图片：</td>
                            <td colspan=3><input name="RYZ" class="textbox" type="text" id="RYZ" value="<%=RS("RYZ")%>" size=40 /> <span style="color: red">* </span> <%if Not KS.IsNul(rs("ryz")) Then
							 response.write "已上传,<a style='color:red' href='" & rs("ryz") &"' target='_blank'>浏览</a>"
							end if%>
                            </td>
                          </tr>
						  
                          <tr class="tdbg">
                            <td  class="clefttitle" style='text-align:right'>个人简介：</td>
                            <td colspan=3><textarea name="Intro" class="textbox" cols="80" rows="7" id="Intro" style="width:500px; height:80px"><%=rs("Intro")%></textarea></td>
                          </tr>
                          <tr class="tdbg">
                            <td  class="clefttitle" style='text-align:right'>审核状态：</td>
                            <td colspan=3>
							<input type="radio" name="status" value="0"<%if rs("status")="0" then response.write " checked"%>>未审核
							<input type="radio" name="status" value="1"<%if rs("status")="1" then response.write " checked"%>>已审核
							
							</td>
                          </tr>
                          <tr class="tdbg">
						    <td  class="clefttitle"></td>
                            <td colspan=3><input type="submit" class="button" value="确定保存" /></td>
                          </tr>
		    </form>
            </table>
		<% RS.Close
		Set RS=Nothing
		End Sub
		Sub ModifySave()
           
			 Dim RealName:RealName=KS.DelSql(KS.G("RealName"))
			 Dim Birthday:Birthday=KS.DelSql(KS.G("Birthday"))
			 Dim QQ:QQ=KS.DelSql(KS.G("QQ"))
			 Dim MSN:MSN=KS.DelSql(KS.G("MSN"))
			 Dim Intro:Intro=KS.DelSql(KS.G("Intro"))
			 Dim Tel:Tel=KS.DelSql(KS.G("Tel"))
			 Dim Province:Province=KS.DelSql(KS.G("Province"))
			 Dim City:City=KS.DelSql(KS.G("City"))
			 Dim SCFL:SCFL=KS.DelSql(KS.G("SCFL"))
			 Dim DanWei:DanWei=KS.DelSql(KS.G("DanWei"))
			 Dim Sex:Sex=KS.DelSql(KS.G("Sex"))
			 Dim TypeName:TypeName=KS.DelSql(KS.G("TypeName"))
			 Dim BigClassID:BigClassID=KS.ChkClng(Request("ClassID"))
			 Dim SmallClassID:SmallClassID=KS.ChkClng(Request("SmallClassID"))
			 Dim SmallerClassID:SmallerClassID=KS.ChkClng(Request("SmallerClassID"))
				 

            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select top 1 * From KS_ASKZJ Where ID=" & KS.ChkClng(Request("id")),Conn,1,3
			  IF Not RS.Eof  Then
				 RS("AddDate")=Now
				 RS("Status")=KS.ChkClng(Request("status"))
				 RS("RealName")=RealName
				 RS("Birthday")=Birthday
				 RS("qq")=qq
				 RS("Sex")=Sex
				 RS("Msn")=Msn
				 RS("Tel")=Tel
				 RS("Province")=Province
				 RS("City")=City
				 RS("SCFL")=SCFL
				 RS("DanWei")=DanWei
				 RS("UserFace")=KS.G("UserFace")
				 RS("IDCard")=KS.G("IDCard")
				 RS("RYZ")=KS.G("RYZ")
				 RS("BigClassID")=BigClassID
				 RS("SmallClassID")=SmallClassID
				 RS("SmallerClassID")=SmallerClassID
				 RS("TypeName")=TypeName
				 RS("Intro")=Intro
		 		 RS.Update
				 RS.Close:Set RS=Nothing
				 Response.Write "<script>alert('恭喜，修改成功！');location.href='KS.AskZJ.asp';</script>"
		    Else
				 Response.Write "<script>alert('出错，找不到记录！');location.href='KS.AskZJ.asp';</script>"
			End If
  End Sub
End Class
%>