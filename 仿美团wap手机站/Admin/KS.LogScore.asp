<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New User_ScoreDetail
KSCls.Kesion()
Set KSCls = Nothing

Class User_ScoreDetail
        Private KS,KSCls
		Private MaxPerPage,RS,TotalPut,TotalPages,I,Page,SQL,ComeUrl
		Private Sub Class_Initialize()
		  MaxPerPage=20
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
       Sub Kesion()
          Response.Write "<html>"
			Response.Write"<head>"
			Response.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			Response.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			Response.Write"<script src=""../ks_inc/jquery.js""></script>"
			Response.Write"</head>"
			Response.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
	      If Not KS.ReturnPowerResult(0, "KMUA10005") Then
			  response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
			  Call KS.ReturnErr(1, "")
			End If
			Response.Write"<div class='topdashed sort' style='text-align:left'>会员积分明细: <a href='?'>所有积分明细</a> | <a href='?channelid=1000'>点广告积分明细</a> | <a href='?channelid=1001'>点友情链接积分明细</a></div>"
		ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))
		if KS.G("Action")="del" then
		  Dim Param
		  Select Case KS.ChkClng(KS.G("DelType"))
		   Case 1 Param="datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>11"
		   Case 2 Param="datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>31"
		   Case 3 Param="datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>61"
		   Case 4 Param="datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>91"
		   Case 5 Param="datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>181"
		   Case 6 Param="datediff(" & DataPart_D & ",adddate," & SqlNowString & ")>366"
		  End Select
		  If Param<>"" Then Conn.Execute("Delete From KS_LogScore Where 1=1 and " & Param)
          KS.echo "<script>$(parent.document).find('#ajaxmsg').toggle();alert('已按所给的条件，删除了积分明细的相关记录！');</script>"
		end if
		%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
  <tr class="sort">
    <td width="80" align="center"><strong> 用户名</strong></td>
    <td width="138" align="center"><strong>产生时间</strong></td>
    <td width="111" align="center"><strong>IP地址</strong></td>
    <td width="54"  align="center"><strong>收入</strong></td>
    <td width="54" align="center"><strong>支出</strong></td>
    <td width="59" align="center"><strong>摘要</strong></td>
    <td width="65" align="center"><strong> 余额</strong></td>
    <td width="75" align="center"><strong> 操作员</strong></td>
    <td align="center"><strong>备注</strong></td>
  </tr>
  <%
  Page	= KS.ChkClng(request("page"))
  If Page<=0 Then Page=1
  Set RS=Server.CreateObject("ADODB.RecordSet")
  If KS.ChkClng(Request("ChannelID"))<>0 Then
    Param=" channelid=" & KS.ChkClng(Request("ChannelID"))
  End If
  if request("keyword")<>"" then
    Param=" username='" & request("keyword") & "'"
  end if
  If Param="" Then Param=" 1=1"
  Dim FieldStr,SQLStr
  FieldStr="ID,UserName,AddDate,IP,Score,InOrOutFlag,CurrScore,[User],Descript"
  If DataBaseType=1 Then
					Dim Cmd : Set Cmd = Server.CreateObject("ADODB.Command")
					Set Cmd.ActiveConnection=conn
					Cmd.CommandText="KS_GetsPageRecords"
					Cmd.CommandType=4	
					CMD.Prepared = true 
					Cmd.Parameters.Append cmd.CreateParameter("@tblName",202,1,200)
					Cmd.Parameters.Append cmd.CreateParameter("@fldName",202,1,200)
					Cmd.Parameters.Append cmd.CreateParameter("@pagesize",3)
					Cmd.Parameters.Append cmd.CreateParameter("@pageindex",3)
					Cmd.Parameters.Append cmd.CreateParameter("@ordertype",3)
					Cmd.Parameters.Append cmd.CreateParameter("@strWhere",202,1,1000)
					Cmd.Parameters.Append cmd.CreateParameter("@fieldIds",202,1,1000)
					Cmd("@tblName")="KS_LogScore"
					Cmd("@fldName")= "ID"
					Cmd("@pagesize")=MaxPerPage
					Cmd("@pageindex")=page
					Cmd("@ordertype")=1
					Cmd("@strWhere")=Param
					Cmd("@fieldIds")=FieldStr
					Set Rs=Cmd.Execute
					Set Cmd=Nothing
	 Else
			SQLStr=KS.GetPageSQL("KS_LogScore","ID",MaxPerPage,Page,1,Param,FieldStr)
			Set RS = Server.CreateObject("AdoDb.RecordSet")
			RS.Open SQLStr, conn, 1, 1
	 End If
		   
	If RS.Eof And RS.Bof Then
	 Response.Write "<tr><td colspan=9 align=center height=25 class='tdbg'>还没有积分记录！</td></tr>"
	Else
                    TotalPut=conn.execute("select count(1) from KS_LogScore Where " & Param)(0)

					SQL = rs.GetRows(MaxPerPage)
					rs.Close:set rs=Nothing
					ShowContent
   End If
%>		
</table>
<div>
		<form action="?" name="myform" method="post">
		   <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
			  &nbsp;<strong>按用户搜索=></strong>
			 &nbsp;用户名:<input type="text" class='textbox' name="keyword">
			  &nbsp;<input type="submit" value="开始搜索" class="button" name="s1">
			  </div>
		</form>
		</div>


<div class="attention">
<strong>特别提醒：</strong>
如果积分明细记录太多，影响了系统性能，可以删除一定时间段前的记录以加快速度。
<br />
<iframe src='about:blank' style='display:none' name='_hiddenframe' id='_hiddenframe'></iframe>
<form action="?action=del" target="_hiddenframe" method=post onsubmit="return(confirm('确实要删除有关记录吗？'))">
删除范围：<input name="deltype" type="radio" value=1>
10天前 
    <input name="deltype" type="radio" value="2" />
    1个月前
    <input name="deltype" type="radio" value="3" />
    2个月前
    <input name="deltype" type="radio" value="4" />
    3个月前
    <input name="deltype" type="radio" value="5" />
    6个月前
    <input name="deltype" type="radio" value="6" checked="checked" />
    1年前
    <input type="submit" value="执行删除" onclick="$(parent.frames['FrameTop'].document).find('#ajaxmsg').toggle();" class="button">
	</form>
</div>
<%End Sub
Sub ShowContent
 Dim InScore,OutScore
For i=0 To Ubound(SQL,2)
	%>
  <tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
    <td class="splittd" width="80" align="center"><%=SQL(1,i)%></td>
    <td class="splittd" align="center"><%=SQL(2,i)%></td>
    <td class="splittd" align="center"><%=SQL(3,i)%></td>
    <td class="splittd" align="right"><%if SQL(5,I)=1 Then Response.Write SQL(4,I) & " 分 ":InScore=InScore+SQL(4,I) ELSE Response.Write "-"%></td>
    <td class="splittd" align="right"><%if SQL(5,I)=2 Then Response.Write SQL(4,I) & " 分" :OutScore=OutScore+SQL(4,I) ELSE Response.Write "-"%></td>
    <td class="splittd" align="center"><%if SQL(5,I)=1 Then Response.Write "<font color=red>收入</font>" Else Response.Write "支出"%></td>
    <td class="splittd" align="center"><%=SQL(6,i)%> 分</td>
    <td class="splittd" align="center"><%=SQL(7,i)%></td>
	<td class="splittd"><%=SQL(8,i)%></td>
  </tr>
  <%Next%>
  <tr class='list' onmouseout="this.className='list'" onmouseover="this.className='listmouseover'">    <td class="splittd" colspan='3' align='right'>本页合计：</td>    <td class="splittd" align='right'><%=InScore%>分</td>    <td align='right'><%=OutScore%>分</td>    <td class="splittd" colspan='4'>&nbsp;</td>  </tr> 

  <% Dim totalinScore:totalinScore=conn.execute("Select sum(Score) From KS_LogScore where InOrOutFlag=1")(0)
     Dim TotalOutScore:TotalOutScore=conn.execute("Select sum(Score) From KS_LogScore where InOrOutFlag=2")(0)
  %>
    <tr class='list' onmouseout="this.className='list'" onmouseover="this.className='listmouseover'">    <td class="splittd" colspan='3' align='right'>所有合计：</td>    <td class="splittd" align='right'><%=totalInScore%>分</td>    <td class="splittd" align='right'><%=totalOutScore%>分</td>    <td class="splittd" colspan='4'>&nbsp;</td>  </tr> 
  <%  
  Response.Write "<tr><td colspan=9 align=right class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"">"
  Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
  Response.Write "</td></tr>"
End Sub
				
End Class
%> 
