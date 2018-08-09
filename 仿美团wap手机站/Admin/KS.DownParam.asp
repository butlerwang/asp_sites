<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Down_Param
KSCls.Kesion()
Set KSCls = Nothing

Class Down_Param
        Private KS,ChannelID
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Sub Kesion()
		 With KS
		  	.echo "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
			.echo "<title>下载基本参数设置</title>"
			.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			.echo "<style type=""text/css"">" & vbCrLf
			.echo "<!--" & vbCrLf
			.echo ".STYLE1 {color: #FF0000}" & vbCrLf
			.echo "-->" & vbCrLf
			.echo "</style>" & vbCrLf
			.echo "</head>"
			
			Dim RS, Action, SQLStr
			Dim DownLb, DownYY, DownSQ, DownPT, JyDownUrl, JyDownWin
			Action = KS.G("Action")
			ChannelID= KS.ChkClng(KS.G("ChannelID"))
			If ChannelID=0 Then ChannelID=3
			If Not KS.ReturnPowerResult(0, "KMST20001") Then Call KS.ReturnErr(1, "")   '下载基本参数设置权限检查
			
			SQLStr = "Select * From KS_DownParam Where ChannelID=" & ChannelID
			Set RS = Server.CreateObject("Adodb.RecordSet")
			If Action = "save" Then
			  RS.Open SQLStr, conn, 1, 3
			  If RS.Eof Then
			   RS.AddNew
			   RS("ChannelID")=ChannelID
			  End If
			  RS("DownLb") = KS.G("DownLB")
			  RS("DownYY") = KS.G("DownYY")
			  RS("DownSQ") = KS.G("DownSQ")
			  RS("DownPT") = KS.G("DownPT")
			  RS.Update
			  .echo ("<script>alert('下载参数修改成功!');</script>")
			  RS.Close
			End If
			 RS.Open SQLStr, conn, 1, 1
			  If Not RS.EOF Then
			   DownLb = RS("DownLb")
			   DownYY = RS("DownYy")
			   DownSQ = RS("DownSQ")
			   DownPT = RS("DownPT")
			  End If
			RS.Close
			
			Set RS = Nothing
			.echo "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.echo "      <div class='topdashed sort'>"
			.echo "      下载参数设置"
			.echo "      </div>"
			
			.echo "<br /><strong>按模型设置:</strong><select id='channelid' name='channelid' onchange=""if (this.value!=0){location.href='?channelid='+this.value;}"">"
			.echo " <option value='0'>---请选择模型---</option>"
	
			If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,Node
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1 and @ks6=3]")
			  If trim(ChannelID)=trim(Node.SelectSinglenode("@ks0").text) Then
			    .echo "<option value='" &Node.SelectSingleNode("@ks0").text &"' selected>" & Node.SelectSingleNode("@ks1").text & "</option>"

			  Else
			   .echo "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
			  End If
			next
			.echo "</select>"

			.echo "<form action=""?ChannelID=" & ChannelID &"&Action=save"" method=""post"" name=""DownParamForm"">"
			.echo "  <table width=""100%"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td height=""30"" colspan=""4"" class='clefttitle' style='text-align:left'>&nbsp;<font color=""#000080""><b>软件性质自定</b></font></td>"
			.echo "    </tr>"
			.echo "    <tr>"
			.echo "      <td width=""25%"" height=""200"" align=""center"" class='tdbg'>设定类别：<br>"
			.echo "        <textarea name=""DownLb"" cols=""20"" rows=""10"" style=""border-style: solid; border-width: 1"">" & DownLb & "</textarea>"
			.echo "        <br>"
			.echo "        <span class=""STYLE1"">说明：每一个类别为一行</span><br></td>"
			.echo "      <td width=""25%"" align=""center"" class='tdbg'>设定语言：<br>"
			.echo "      <textarea name=""DownYy"" cols=""20"" rows=""10"" style=""border-style: solid; border-width: 1"">" & DownYY & "</textarea>"
			.echo "      <br>"
			.echo "      <span class=""STYLE1"">说明：每一种语言为一行</span></td>"
			.echo "      <td width=""25%"" align=""center"" class='tdbg'>授权形式： <br>"
			.echo "      <textarea name=""DownSq"" cols=""20"" rows=""10"" style=""border-style: solid; border-width: 1"">" & DownSQ & "</textarea>"
			.echo "        <br>"
			.echo "        <span class=""STYLE1"">说明：每一种授权方式为一行</span></td>"
			.echo "      <td width=""25%"" align=""center"" class='tdbg'>运行平台：<br>"
			.echo "      <textarea name=""DownPt"" cols=""20"" rows=""10"" style=""border-style: solid; border-width: 1"">" & DownPT & "</textarea>"
			.echo "      <br>"
			.echo "      <span class=""STYLE1"">说明：每一种运行平台为一行</span></td>"
			.echo "    </tr>"
			.echo "  </table>"
			.echo "</form>"
			.echo "</body>"
			.echo "</html>"
			.echo "<Script Language=""javascript"">"
			.echo "<!--" & vbCrLf
			.echo "function CheckForm()" & vbCrLf
			.echo "{ var form=document.DownParamForm;" & vbCrLf
			.echo "    form.submit();" & vbCrLf
			.echo "    return true;" & vbCrLf
			.echo "}" & vbCrLf
			.echo "//-->"
			.echo "</Script>"
			End With
		End Sub

End Class
%> 
