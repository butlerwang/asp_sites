<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New PaymentPlatCls
KSCls.Kesion()
Set KSCls = Nothing

Class PaymentPlatCls
        Private KS,Action,KSCls
		Private K, SqlStr,ChannelID,SQL,RS
		
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls= New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
             With KS
		 	     .echo "<html>"
				 .echo "<head>"
				 .echo "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
				 .echo "<title>支付平台管理</title>"
				 .echo "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		         .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
	         	 .echo "<script language=""JavaScript"" src=""../KS_Inc/jQuery.js""></script>" & vbCrLf
					If Not KS.ReturnPowerResult(0, "KMST10001") Then          '检查是否有基本信息设置的权限
					  .echo ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back()';</script>")
					 Call KS.ReturnErr(1, "")
					 .End
					 End If


			  Action=KS.G("Action")
			 Select Case Action
			  Case "Modify"
			    Call DoModify()
			  Case "DoModifySave"
			    Call DoModifySave()
			  Case "DoBatch"
			    Call DoBatch()
			  Case "Disabled"
			    Call DoDisabled()
			  Case Else
			   Call MainList()
			 End Select
			 .echo "</body>"
			 .echo "</html>"
			End With
		End Sub
		
		Sub MainList()
		With KS
		 .echo "</head>"
		%><script language="javascript">
		 function CheckForm()
		 {
		   this.myform.submit();
		 }
		</script>
		<%
		 .echo "<body topmargin='0' leftmargin='0'>"
		 .echo "<ul id='mt'> <div id='mtl'>友情提示：</div><li>"
		 .echo "本系统集成多家在线支付接口，您可以在此管理所有的支付平台 "
		 .echo "</ul>"
		 .echo "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		 .echo(" <form name=""myform"" method=""Post"" action=""?Action=DoBatch"">")
		 .echo "    <tr class='sort'>"
		 .echo "    <td width='30' width='50' align='center'>序号</td>"
		 .echo "    <td align='center'>支付平台</td>"
		 .echo "    <td width='100' align='center'>商家ID</td>"
		 .echo "    <td width='300' align='center'>备注说明</td>"
		 .echo "    <td width='60' align='center'>手续费</td>"
		 .echo "    <td width='40' align='center'>默认</td>"
		 .echo "    <td width='40' align='center'>启用</td>"
		 .echo "    <td width='100' align='center'>管理操作</td>"
		 .echo "    <td width='60' align='center'>申请</td>"
		 .echo "  </tr>"
		 Set RS = Server.CreateObject("ADODB.RecordSet")
         SqlStr = "SELECT ID,OrderID,PlatName,AccountID,Note,MD5Key,Rate,RateByUser,IsDisabled,IsDefault FROM [KS_PaymentPlat] order by orderid"
		 RS.Open SqlStr, conn, 1, 1
		 If Not RS.EOF Then SQL=RS.GetRows(-1)
		 If Not IsArray(SQL) Then
			 .echo "<tr><td  class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"" colspan=8 height='25' align='center'>没有任何支付平台!</td></tr>"
		 Else
			 For K=0 To Ubound(SQL,2)
		       .echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			   .echo "<td class='splittd' align='center'><input type='hidden' value='" & SQL(0,K) & "' name='id'><input type='text' name='orderid' value='" &SQL(1,K) & "' style='width:36px;text-align:center'></td>"
			   .echo " <td class='splittd' height='22' nowrap>&nbsp;<span style='cursor:default;'>"
			   .echo SQL(2,K)
			   .echo "</td>"
			   
			    .echo " <td class='splittd' align='center'>" & SQL(3,K) & "</td>"
			    .echo " <td class='splittd' align='center'>" & SQL(4,K) & "&nbsp;</td>"
			    .echo " <td class='splittd' align='center'>" & SQL(6,K) & "%</td>"
			    .echo " <td class='splittd' align='center'>"
			   If SQL(9,K)=1 Then
			    .echo "<input type='radio' name='IsDefault' value='" & SQL(0,K) & "' checked>"
			   Else
			    .echo "<input type='radio' name='IsDefault' value='" & SQL(0,K) & "'>"
			   End If
			    .echo " </td>"
			    .echo " <td class='splittd' align='center'>" 
			   If SQL(8,K)=1 Then
			    .echo "<input type='checkbox' name='IsDisabled' value='" & SQL(0,K) & "' checked>"
			   Else
			    .echo "<input type='checkbox' name='IsDisabled' value='" & SQL(0,K) & "'>"
			   End If
			    .echo "</td>"
			    .echo " <td class='splittd' align='center'><a href='?Action=Modify&ID=" & SQL(0,K) &"'>修改</a> "
			   If SQL(8,K)=1 Then
			   	 .echo " <a href='?V=0&Action=Disabled&id=" & SQL(0,K) & "'>关闭</a>"
			   Else
			     .echo " <a href='?V=1&Action=Disabled&id=" & SQL(0,K) & "'>启用</a>"
			   End If
			    .echo " </td>"
			    .echo "<td class='splittd' nowrap>"
			   	Select Case SQL(0,K)
			    Case 10  .echo "<a href='http://union.tenpay.com/mch/mch_register.shtml?sp_suggestuser=1202640601' target='_blank'>申请商户"
				case 11  .echo "<a href='http://union.tenpay.com/mch/mch_register_1.shtml?sp_suggestuser=1202640601' target='_blank'>申请商户"
			    Case  1  .echo "<a href='http://merchant3.chinabank.com.cn/register.do' target='_blank'>申请商户</a>"
			    Case  5  .echo "<a href='http://new.xpay.cn/SignUp/Default.aspx' target='_blank'>申请商户</a>"
			    Case  6  .echo "<a href='https://www.cncard.net/products/products.asp' target='_blank'>申请商户</a>"
			    Case  7,9,15  .echo "<a href='http://act.life.alipay.com/systembiz/kesion/' target='_blank'>申请商户</a>"
			    Case  8  .echo "<a href='https://www.99bill.com/website/' target='_blank'>申请商户</a>"
			    Case  2  .echo "<a href='http://www.ipay.cn/home/index.php' target='_blank'>申请商户</a>"
			    Case  4  .echo "<a href='http://www.yeepay.com/' target='_blank'>申请商户</a>"
			    Case  3  .echo "<a href='https://www.ips.com.cn/' target='_blank'>申请商户</a>"
				case 12  .echo "<a href='https://www.paypal.com/c2/cgi-bin/webscr?cmd=_registration-run' target='_blank'>申请商户</a>"
				case 13  .echo "<a href='https://www.paypal.com/cn/cgi-bin/webscr?cmd=_registration-run' target='_blank'>申请商户</a>"
				case 14  .echo "<a href='https://www.umbpay.com/mer/' target='_blank'>申请商户</a>"
			   End Select 
                .echo "</td>"
			    .echo "</tr>"
			Next
								 
		 End If
          .echo "<tr>"
		  .echo " <td colspan='8' height='40'>&nbsp;&nbsp;"
		  .echo " <input type='submit' value='批量保存设置' class='button'><font color=blue>&nbsp;序号越小在前台排在越前面，只有在这里设置启用的支付平台，前台才会显示</font>"
		  .echo " </td>"
		  .echo "</tr>"
		  .echo "</form>"
		  .echo "</table>"
		
		End With
		End Sub

		Sub DoModify()
		 Dim ID:ID=KS.ChkClng(KS.G("ID"))
		 Dim RS:Set RS=Server.CreateOBject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_PaymentPlat Where ID=" & ID,conn,1,1
		 If RS.EOf And RS.Bof Then
		 RS.Close:Set RS=Nothing
		  KS.Echo "<script>alert('参数传递错误!');history.back();</script>"
		  Exit Sub
		 End If
		%>
		<html>
		<head>
		<title>支付平台管理</title>
		<meta http-equiv=Content-Type content="text/html; charset=utf-8">
		<link href="Include/Admin_Style.CSS" type=text/css rel=stylesheet>
		</head>
		<body leftMargin=0 topMargin=0>
		<script language="javascript">
		 function CheckForm()
		 {
		   this.myform.submit();
		 }
		</script>
		<ul id=menu_top>
		<li class=parent onclick=return(CheckForm())><SPAN class=child onMouseOver="this.parentNode.className='parent_border'" onMouseOut="this.parentNode.className='parent'"><img src="images/ico/save.gif" align=absMiddle border=0>确定保存</SPAN></li>
		<li class=parent onClick="location.href='?ChannelID=1';"><SPAN class=child onMouseOver="this.parentNode.className='parent_border'" onMouseOut="this.parentNode.className='parent'"><img src="images/ico/back.gif" align=absMiddle border=0>取消返回</SPAN></li></ul>
		<FORM name=myform onsubmit=return(CheckForm()) action="?action=DoModifySave&ID=<%=rs("ID")%>" method=post>
		  <table class=ctable style=" BORDER-COLLAPSE: collapse" cellSpacing=1 cellPadding=1 width="100%" align=center border=0>
			<tr class=tdbg>
			  <td class=clefttitle noWrap align=right height=25><strong>平台名称：</strong></td>
			  <td align=right width=21 height=30>
				<Input value="<%=rs("PlatName")%>" Class="textbox" name="PlatName"> </td>
				<tr class=tdbg>
				  <td class=clefttitle noWrap align=right height=25><strong>备注说明：</strong></td>
				  <td noWrap height=25>
		                <textarea name="Note" cols="70" rows="5"><%=rs("note")%></textarea>
		            </td>
					<tr class=tdbg>
					  <td class=clefttitle align=right><strong>支付编号：</strong><br>
请填入您在线支付平台申请的商户编号</td>
					  <td>
						<Input id="AccountID" class="textbox" name="AccountID" value="<%=rs("AccountID")%>"> </td>
					</tr>
					<tr class=tdbg>
					  <td class=clefttitle align=right height=25><strong>支付密钥：</strong><br>
请填入您在上述在线支付平台中设置的MD5私钥,部分在线支付平台不需要此项</td>
					  <td height=25>
						<Input class="textbox" name="MD5Key" value="<%=rs("MD5Key")%>"></td>
					</tr>

					<tr class=tdbg>
					  <td class=clefttitle align=right height=25><strong>手续费率：</strong></td>
					  <td noWrap height=25>
						<Input class="textbox" size="6" name="Rate"  value="<%=rs("rate")%>">%
						<br>
						<input type="checkbox" name="RateByUser" value="1"<%if rs("ratebyuser")=1 Then KS.Echo " checked"%>>
						 手续费由付款人额外支付
						</td>
					</tr>
					<tr class=tdbg>
					  <td class=clefttitle align=right><strong>是否启用:</strong></td>
					  <td>
					    <%if rs("isdisabled")=1 Then%>
						<input type="radio" value="0" name="isdisabled">禁用
						<input type="radio" value="1" name="isdisabled" checked>启用
						<%else%>
						<input type="radio" value="0" name="isdisabled" checked>禁用
						<input type="radio" value="1" name="isdisabled">启用
						<%end if%>
						
					  </td>
					</tr>
				  </table>
				  <div style='margin:8px;text-align:center'>
				   <input type='button' onclick='CheckForm()' class='button' value='确定保存'>&nbsp;
				   <input type='button' class='button' value='取消返回' onClick="javascript:location.href='KS.PaymentPlat.asp';">
				  </div>
				   </FORM>
				</body>
				</html>
		<%
		 RS.Close:Set RS=Nothing
		End Sub
		
		Sub DoModifySave()
		  Dim ID:ID=KS.ChkClng(KS.G("ID"))
		  Dim PlatName:PlatName=KS.G("PlatName")
		  Dim Note:Note=KS.G("Note")
		  Dim AccountID:AccountID=KS.G("AccountID")
		  Dim MD5Key:MD5Key=KS.G("MD5Key")
		  Dim Rate:Rate=KS.G("Rate")
		  Dim RateByUser:RateByUser=KS.ChkClng(KS.G("RateByUser"))
		  Dim IsDisabled:IsDisabled=KS.ChkClng(KS.G("IsDisabled"))
		  Dim RS:Set RS=Server.CreateOBject("ADODB.RECORDSET")
		  RS.Open "Select * from KS_PaymentPlat where id=" & ID,conn,1,3
		  If Not RS.Eof Then
		    RS("PlatName") = PlatName
			RS("Note")     = Note
			RS("AccountID")= AccountID
			RS("MD5Key")   = MD5Key
			RS("Rate")     = Rate
			RS("RateByUser")=RateByUser
			RS("IsDisabled")= IsDisabled
			RS.Update
		  End If
		  RS.Close:Set RS=Nothing
		  KS.Alert "恭喜，修改成功！","KS.PaymentPlat.asp" 
		End Sub
		
		Sub DoBatch()
			Dim ID:ID = KS.G("ID")
			Dim OrderID:OrderID=KS.G("OrderID")
			Dim IsDisabled:IsDisabled=KS.G("IsDisabled")
			Dim IsDefault:IsDefault=KS.G("IsDefault")
			Dim ID_Arr:ID_Arr=Split(ID,",")
			Dim OrderID_Arr:OrderID_Arr=Split(OrderID,",")
		    Dim K
			For K=0 TO Ubound(ID_Arr)
			 Conn.Execute("Update KS_PaymentPlat Set OrderID=" & OrderID_Arr(K) & " where id=" & ID_Arr(K))
			 If KS.FoundInArr(IsDisabled, ID_Arr(K), ",")=true Then
			  Conn.Execute("Update KS_PaymentPlat Set IsDisabled=1 where id=" & ID_Arr(K))
			 Else
			  Conn.Execute("Update KS_PaymentPlat Set IsDisabled=0 where id=" & ID_Arr(K))
			 End If
			 If KS.FoundInArr(IsDefault, ID_Arr(K), ",")=true Then
			  Conn.Execute("Update KS_PaymentPlat Set IsDefault=1 where id=" & ID_Arr(K))
			 Else
			  Conn.Execute("Update KS_PaymentPlat Set IsDefault=0 where id=" & ID_Arr(K))
			 End If
			Next
			KS.Alert "恭喜，批量设置成功！" , "KS.PaymentPlat.asp"
		 End Sub
		Sub DoDisabled()
		  Conn.Execute("Update KS_PaymentPlat Set IsDisabled=" & KS.ChkClng(KS.G("V")) & " where id=" & KS.ChkClng(KS.G("ID")))
		  KS.AlertHintScript "恭喜,操作成功!"
		End Sub
End Class
%> 
