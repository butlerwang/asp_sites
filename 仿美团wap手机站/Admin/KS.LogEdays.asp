<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_User
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_User
        Private KS
		Private MaxPerPage,RS,TotalPut,TotalPages,I,CurrentPage,SQL,ComeUrl
		Private Sub Class_Initialize()
		  MaxPerPage=20
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
       Sub Kesion()
	    If Not KS.ReturnPowerResult(0, "KMUA10006") Then
			  Call KS.ReturnErr(1, "")
			End If
          Response.Write "<html>"
			Response.Write"<head>"
			Response.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			Response.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			Response.Write"</head>"
			Response.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Response.Write"<div class='topdashed sort'>会员有效期明细</div>"
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
		  If Param<>"" Then Conn.Execute("Delete From KS_LogEdays Where 1=1 and  " & Param)
		  Call KS.Alert("已按所给的条件，删除了有效期明细的相关记录！",ComeUrl)
		end if
		%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
  <tr class="sort">
    <td width="80">用户名</td>
    <td width="138">操作时间</td>
    <td width="111">IP地址</td>
    <td width="71">增加</td>
    <td width="74">减少</td>
    <td width="50">摘要</td>
    <td width="75">操作员</td>
    <td width="200">备注</td>
  </tr>
  <%
  CurrentPage	= KS.ChkClng(request("page"))
  Set RS=Server.CreateObject("ADODB.RecordSet")
    RS.Open "Select ID,UserName,AddDate,IP,Edays,InOrOutFlag,[User],Descript From KS_LogEdays order by ID desc",conn,1,1
	If RS.Eof And RS.Bof Then
	 Response.Write "<tr><td colspan=9 align=center height=25 class='splittd'>找不到相关记录！</td></tr>"
	Else
       TotalPut=rs.recordcount
					if (TotalPut mod MaxPerPage)=0 then
						TotalPages = TotalPut \ MaxPerPage
					else
						TotalPages = TotalPut \ MaxPerPage + 1
					end if
					if CurrentPage > TotalPages then CurrentPage=TotalPages
					if CurrentPage < 1 then CurrentPage=1
					rs.move (CurrentPage-1)*MaxPerPage
					SQL = rs.GetRows(MaxPerPage)
					rs.Close:set rs=Nothing
					ShowContent
   End If
%>		
</table>
<div class="attention">
<table border="0" style="margin-top:20px" width="96%" align=center>
<tr><td><strong>特别提醒：</strong>
如果有效期明细记录太多，影响了系统性能，可以删除一定时间段前的记录以加快速度。但可能会带来会员在查看以前收过费的信息时重复收费（这样会引发众多消费纠纷问题），无法通过有效期明细记录来真实分析会员的消费习惯等问题。
</td></tr>
<form action="?action=del" method=post onsubmit="return(confirm('确实要删除有关记录吗？一旦删除这些记录，会出现会员查看原来已经付过费的收费信息时重复收费等问题。请慎重!'))">
<tr><td>删除范围：<input name="deltype" type="radio" value=1>
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
    <input type="submit" value="执行删除" class="button"></td></tr>
  </form>
</table>
</div>
<%End Sub
Sub ShowContent
For i=0 To Ubound(SQL,2)
	%>
  <tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
    <td width="80" align="center" class="splittd"><%=SQL(1,i)%></td>
    <td align="center" class="splittd"><%=SQL(2,i)%></td>
    <td align="center" class="splittd"><%=SQL(3,i)%></td>
    <td align="center" class="splittd"><%if SQL(5,I)=1 Then Response.Write SQL(4,I) ELSE Response.Write "-"%>天</td>
    <td align="center" class="splittd"><%if SQL(5,I)=2 Then Response.Write SQL(4,I) ELSE Response.Write "-"%>天</td>
    <td align="center" class="splittd"><%if SQL(5,I)=1 Then Response.Write "<font color=red>收入</font>" Else Response.Write "支出"%></td>
    <td align="center" class="splittd"><%=SQL(6,i)%></td>
		<td class="splittd"><%=SQL(7,i)%></td>
  </tr>

  <%Next
  Response.Write "<tr><td colspan=9 align=right class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"">"
  Call KS.ShowPageParamter(totalPut, MaxPerPage, "", True, "条记录", CurrentPage, "")
  Response.Write "</td></tr>"
End Sub
				
End Class
%> 
