<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%dim channelid
Dim KSCls
Set KSCls = New Admin_Ask_Class
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Ask_Class
        Private KS,DataArry,TypeFlag
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim Action,DataArry
		Action = LCase(Request("action"))
		TypeFlag=KS.ChkClng(KS.S("TypeFlag"))
		If TypeFlag=1 Then
         If Not KS.ReturnPowerResult(0, "KSMB10003") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 KS.Die ""
		 End If
		Else
			If Not KS.ReturnPowerResult(0, "WDXT10003") Then          '检查是权限
						 Call KS.ReturnErr(1, "")
						 KS.Die ""
			 End If
		End If
		Select Case Trim(Action)
		Case "save"
			Call saveScore()
		Case Else
			Call showmain()
		End Select
		End Sub
		Sub showmain()
			Dim i,iCount,lCount
			iCount=2:lCount=1
		%>
		<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
		<html xmlns="http://www.w3.org/1999/xhtml">
		<head>
		<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
		<script src="../KS_Inc/common.js" language="JavaScript"></script>
		</head>
		<body>
		<div class='topdashed sort'><%IF TypeFlag=1 then response.write "论坛" else response.write "问吧"%>等级头衔设置</div>
		<table id="tablehovered" border="0" align="center" cellpadding="3" cellspacing="1" width="100%">
		<form name="selform" id="selform" method="post" action="?">
		<input type="hidden" name="action" value="save">
		<input type="hidden" name="typeflag" value="<%=typeflag%>"/>
		<tr class='sort'>
			<td width="10%" noWrap="noWrap">等级ID</td>
			<td>用户等级头衔</td>
			<td noWrap="noWrap">颜色</td>
			<td noWrap="noWrap">图标</td>
			<%if TypeFlag=1 then%>
			<td noWrap="noWrap">论坛帖子数</td>
			<%end if%>
			<td noWrap="noWrap">所需积分数</td>
			<%if typeflag=1 then%>
			<td noWrap="noWrap">用户数</td>
			<%end if%>
			<td width="15%" noWrap="noWrap">管理操作</td>
		</tr>
		<%
			Call showScoreList()
			iCount=1:lCount=2
			If IsArray(DataArry) Then
				For i=0 To Ubound(DataArry,2)
					If Not Response.IsClientConnected Then Response.End
		%>
		<tr align="center">
			<td class="splittd"><input type="hidden" name="GradeID" value="<%=DataArry(0,i)%>"><%=DataArry(0,i)%></td>
			<%if DataArry(5,i)="0" then%>
			<td class="splittd"><input class="textbox" type="text" size="20" name="UserTitle<%=DataArry(0,i)%>" value="<%=Server.HTMLEncode(DataArry(1,i))%>" /></td>
			<%else%>
			<td class="splittd"><%=Server.HTMLEncode(DataArry(1,i))%> (<font color=red>系统</font>)<input type="hidden" size="20" name="UserTitle<%=DataArry(0,i)%>" value="<%=Server.HTMLEncode(DataArry(1,i))%>" /></td>
			<%end if%>
			<td class="splittd"><input class="textbox" type="text" size="10" name="color<%=DataArry(0,i)%>" value="<%=DataArry(6,i)%>" /></td>
			<td class="splittd"><input class="textbox" type="text" size="10" name="ico<%=DataArry(0,i)%>" value="<%=DataArry(3,i)%>" />
			<img src="../<%=KS.Setting(66)%>/images/<%=DataArry(3,i)%>" />
			</td>
			<%if typeflag=1 then%>
			  <%if DataArry(5,i)="0" then%>
			<td class="splittd"><input class="textbox" type="text" size="5" name="ClubPostNum<%=DataArry(0,i)%>" value="<%=DataArry(4,i)%>" /></td>
			  <%else%>
			   <td class="splittd">---</td>
			 <%end if%>
			<%end if%>
			 <%if DataArry(5,i)="0" then%>
			<td class="splittd"><input class="textbox" type="text" size="5" name="Score<%=DataArry(0,i)%>" value="<%=DataArry(2,i)%>" /></td>
			  <%else%>
			   <td class="splittd">---</td>
			  <%end if%>
			  
			  <%if typeflag=1 then%>
			<td class="splittd">
			 <a href='KS.User.asp?UserSearch=14&ClubGradeID=<%=DataArry(0,i)%>'>
			<%=conn.execute("select count(1) from ks_user where clubgradeid=" & DataArry(0,i))(0)%> 位
			</a>
			</td>
			  <%end if%>
			<td class="splittd">
			<%if DataArry(5,i)="0" then%>
			 <a href="?x=c&typeflag=<%=typeflag%>&id=<%=DataArry(0,i)%>" onClick="return(confirm('确定删除吗?'))">删除</a>
			<%else%>
			 <a href="#" disabled>删除</a>
			<%end if%>			</td>
		</tr>
		<%
				Next
			End If
			DataArry=Null
		%>
		<tr align="center">
			<td class="tablerow<%=lCount%>" colspan="6">
				<input class="button" type="submit" name="submit_button" value="批量保存设置"/>			</td>
		</tr>
		</form>

		<form action="?x=b&typeflag=<%=typeflag%>" method="post" name="myform" id="form">
		    <tr>
			<td height="25" colspan="7">&nbsp;&nbsp;<strong>&gt;&gt;新增等级头衔</strong><<</td>
		    </tr>
			<tr><td colspan=10 background='images/line.gif'></td></tr>
			<tr valign="middle" class="list"> 
			  <td height="25"></td>
			  <td height="25" align="center"><input name="UserTitle" type="text" class="textbox" id="UserTitle" size="25"></td>
			  <td align="center"><input style="text-align:center" name="color" type="text" value="#000000" class="textbox" id="color" size="10"></td>
			  <td align="center"><input style="text-align:center" name="rank" type="text" value="rank0.gif" class="textbox" id="rank" size="10"></td>
			  <td align="center"><input style="text-align:center" name="clubpostnum" type="text" value="100" class="textbox" size="5"></td>
			  <td height="25" align="center"><input style="text-align:center" name="Score" type="text" value="1000" class="textbox" id="Score" size="5">
分</td>
			  <td height="25" align="center"><input name="Submit3" class="button" type="submit" value="OK,提交"></td>
			</tr>
			<tr><td colspan=10 background='images/line.gif'></td></tr>
		</form>
		</table>
		<div class="attention" style="color:#FF0000">
		<strong>说明：</strong><br/>
		<li>等级图标必须放于<%=KS.Setting(66)%>/images目录下；</li>
		<%If request("flag")="1" then%>
		<li>如果您的站点有开启积分兑换功能，则建议等级与发帖数相关与积分无关,即积分设置为0，以免用户兑换礼品后影响论坛等级；</li>
		<%end if%>
		</div>
		<%
		 Select case request("x")
		   case "b"
		       If KS.G("UserTitle")="" Then Response.Write "<script>alert('请输入等级头衔!');history.back();</script>":response.end
			   If Not Isnumeric(KS.G("Score")) Then Response.Write "<script>alert('积分必须用数字!');history.back();</script>":response.end
			    Dim GradeID:GradeID=KS.ChkClng(Conn.Execute("Select Max(gradeid) From KS_AskGrade")(0))+1
				conn.execute("Insert into KS_AskGrade(GradeID,UserTitle,score,ico,clubpostnum,typeflag,color)values(" & GradeID & ",'" & KS.G("UserTitle") & "','" & KS.ChkClng(KS.G("Score")) & "','" & KS.G("Rank") & "'," & KS.ChkClng(KS.G("clubpostnum")) &"," & typeflag &",'" & KS.S("Color") & "')")

				
				KS.AlertHintScript "恭喜,等级头衔成功!"
		   case "c"
				conn.execute("Delete from KS_AskGrade where GradeID="& KS.ChkClng(KS.G("id")))
				KS.AlertHintScript "恭喜,等级头衔删除成功!"
		End Select
		  
		End Sub
		
		Sub showScoreList()
			Dim Rs,SQL
			SQL="SELECT GradeID,UserTitle,Score,Ico,ClubPostNum,Special,Color FROM [KS_AskGrade] Where TypeFlag=" & TypeFlag & " order by gradeid"
			Set Rs=Conn.Execute(SQL)
			If Not (Rs.BOF And Rs.EOF) Then
				DataArry=Rs.GetRows(-1)
			Else
				DataArry=Null
			End If
			Rs.close()
			Set Rs=Nothing
		End Sub
		
		Sub saveScore()
			Dim Rs,SQL,i
			Dim GradeID,UserTitle,Score,Ico,clubpostnum,Color
			    GradeID=Split(Replace(Request.Form("GradeID")," ",""),",")
                For I=0 To Ubound(GradeID)
				 UserTitle=Replace(Request.Form("UserTitle"&GradeID(I)),"'","")
				 Score=KS.ChkClng(Request.Form("Score"&GradeID(I)))
				 Ico=Request.Form("Ico"&GradeID(I))
				 Color=Request.Form("Color"&GradeID(I))
				 clubpostnum=KS.ChkClng(Request.Form("clubpostnum"&GradeID(I)))
				 If GradeID(I)>0 Then
					Conn.Execute ("UPDATE KS_AskGrade SET clubpostnum=" & clubpostnum &",Ico='" & Ico & "',UserTitle='"&UserTitle&"',Score="&Score&",Color='" & Color &"' WHERE GradeID="&GradeID(I))
				 End If
			   Next
			Call KS.AlertHintScript("恭喜您！保存用户积分等级成功!")
		End Sub
End Class
%>