<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.CollectCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Collect_ItemModify
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_ItemModify
        Private KS
		Private KMCObj
		Private ConnItem
		Private SqlItem, RsItem, FoundErr, ErrMsg
		
		Private ItemID, ItemName, WebName, WebUrl, ChannelID, ClassID, SpecialID, LoginType, LoginUrl, LoginPostUrl, LoginUser, LoginPass, LoginFalse, ItemDemo,CharsetCode
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KMCObj=New CollectPublicCls
		  Set ConnItem = KS.ConnItem()
		End Sub
        Private Sub Class_Terminate()
		 Call KS.CloseConnItem()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KMCObj=Nothing
		End Sub
		Sub Kesion()
		ItemID = Trim(Request("ItemID"))
		If ItemID = "" Then
		  ChannelID=KS.ChkClng(KS.G("ChannelID")):CharsetCode="utf-8"
		Else
		   ItemID = CLng(ItemID)
		   SqlItem = "select ItemID,ItemName,CharsetCode,WebName,WebUrl,ChannelID,ClassID,SpecialID,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,ItemDemo From KS_CollectItem where ItemID=" & ItemID
		   Set RsItem = Server.CreateObject("adodb.recordset")
		   RsItem.Open SqlItem, ConnItem, 1, 1
		   If RsItem.EOF And RsItem.BOF Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "<br><li>参数错误，没有找到该项目！</li>"
		   Else
			  ItemName = RsItem("ItemName")
			  CharsetCode=RsItem("CharsetCode")
			  ItemDemo = RsItem("ItemDemo")
			  WebName = RsItem("WebName")
			  WebUrl = RsItem("WebUrl")
			  ChannelID = RsItem("ChannelID")
			  ClassID = RsItem("ClassID")
			  SpecialID = RsItem("SpecialID")
			  LoginType = RsItem("LoginType")
			  LoginUrl = RsItem("LoginUrl")
			  LoginPostUrl = RsItem("LoginPostUrl")
			  LoginUser = RsItem("LoginUser")
			  LoginPass = RsItem("LoginPass")
			  LoginFalse = RsItem("LoginFalse")
		   End If
		   RsItem.Close
		   Set RsItem = Nothing
		End If
		
		If FoundErr = True Then
		  Call KS.AlertHistory(ErrMsg,-1)
		Else
		   'Call KMCObj.GetClassList
		   Call Main
		End If
		
		End Sub
		Sub Main()
		With KS
		  .echo "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd"">"
		  .echo "<html xmlns=""http://www.w3.org/1999/xhtml"">"
		  .echo "<head>"
		  .echo "<title>采集系统</title>"
		  .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		  .echo "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Admin_Style.css"">"
		  .echo "<script src=""../../ks_inc/jquery.js""></script>"
		  .echo "</head>"
		  .echo "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		  .echo "<div class=""topdashed"">"
		  .echo  KMCObj.GetItemLocation(1,ItemID)
		  .echo "</div>"
		  .echo "<br>"
		  .echo "<table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""ctable"" >"
		  .echo "<form method=""post"" action=""Collect_ItemModify2.asp"" name=""myform""  onSubmit=""return(CheckForm(this))"">"
		  .echo "    <tr class='tdbg'>"
		   .echo "     <td width=""20%"" height=""25"" align=""center"" class='clefttitle'>项目名称：</td>"
		   .echo "     <td width=""796"">"
		   .echo "     <input name=""ItemName"" type=""text"" size=""27"" class=""textbox"" maxlength=""30"" value=""" & ItemName & """>&nbsp;&nbsp;<font color=red>*</font>如：新浪网－新闻中心</td>"
		    .echo "  </tr>"
		    .echo "  <tr class='tdbg'>"
		    .echo "    <td width=""20%"" height=""25"" align=""center"" class='clefttitle'> 所属模型：</td>"
		    .echo "    <td width=""796""><select ID=""ChannelID"" name=""ChannelID"" onChange=""SelectClass(this.value)"" style=""Width:200"">"
		    .echo KMCObj.Collect_ShowChannel_Option(ChannelID) & "</select>      </td>"
		    .echo "  </tr>"
		   .echo "   <tr class='tdbg'>"
		   .echo "     <td width=""20%"" height=""25"" align=""center"" class='clefttitle'> 所属栏目：</td>"
		   .echo "     <td width=""796"" ID=""ClassArea""><select name=""ClassID"" ID=""ClassID"" style=""Width:200"">"
		   .echo Replace(KS.LoadClassOption(ChannelID,true),"value='" & ClassID & "'","value='" & ClassID &"' selected") & "</select>      </td>"
		   .echo "   </tr>"
		  .echo "    <tr style=""display:none"">"
		  .echo "      <td width=""20%"" height=""25"" align=""center"" class='clefttitle'> 所属专题：</td>"
		  .echo "      <td width=""796""><input type=""hidden"" value=""0"" name=""specialid"">"
		'call KMCObj.Collect_ShowSpecial_Option(1,0)
		  .echo "      </td>"
		   .echo "   </tr>"
		   .echo "   <tr  class='tdbg' class='tdbg'>"
		   .echo "     <td width=""20%"" height=""25"" align=""center"" class='clefttitle'> 网站名称：</td>"
		   .echo "     <td width=""796"">"
		    .echo "      <input name=""WebName"" type=""text"" class=""textbox"" size=""27"" maxlength=""30"" value=""" & WebName & """>      </td>"
		    .echo "  </tr>"
		     .echo "   <tr  class='tdbg' class='tdbg'>"
		   .echo "     <td width=""20%"" height=""25"" align=""center"" class='clefttitle'> 编码方式：</td>"
		   .echo "     <td width=""796"">"
		   .echo " <select name=""CharsetCode"">"
	       .echo " <option value='auto'>自动检测</option>"
	       .echo " <option value='gb2312' "
		 if CharsetCode="gb2312" then   .echo("selected")
		   .echo " >gb2312</option>"
	       .echo "<option value='utf-8' "
		 if CharsetCode="utf-8" then   .echo("selected")
		   .echo ">utf-8</option>"
	       .echo " </select>"
	        .echo "   </td>"
		    .echo "  </tr>"
		    .echo "  <tr class='tdbg' style=""display:none"">"
		    .echo "    <td width=""20%"" height=""25"" align=""center"" class='clefttitle'> 网站网址：</td>"
		    .echo "    <td width=""796""><input name=""WebUrl"" type=""text"" size=""49"" maxlength=""150"" value=""" & WebUrl & """>      </td>"
		    .echo "  </tr>"
		    .echo " <tr class='tdbg'>"
		    .echo "    <td  class='clefttitle' style=""text-align:right"" height=""25""> 网站登录：</td>"
		    .echo "    <td>"
		    .echo "      <input type=""radio"" value=""0"" name=""LoginType"" "
		  If LoginType = 0 Then   .echo "checked"
		    .echo " onClick=""Login.style.display='none'"">不需要登录<span lang=""en-us"">&nbsp;"
		    .echo "      </span>"
		    .echo "      <input type=""radio"" value=""1"" name=""LoginType"" "
		  If LoginType = 1 Then   .echo "checked"
		    .echo " onClick=""Login.style.display=''"">设置参数      </td>"
		    .echo "   </tr>"
		    .echo " <tr  class='tdbg' id=""Login"""
		  If LoginType = 0 Then   .echo " style=""display:none"" "
		   .echo "      ><td width=""20%"" height=""25""style=""text-align:right""> 登录参数：</td>"
		   .echo "     <td>"
		   .echo "       登录地址：<input name=""LoginUrl""  class=""textbox"" type=""text"" size=""40"" maxlength=""150"" value=""" & LoginUrl & """><br>"
		   .echo "       提交地址：<input name=""LoginPostUrl""  class=""textbox"" type=""text"" size=""40"" maxlength=""150"" value=""" & LoginPostUrl & """><br>"
		   .echo "       用户参数：<input name=""LoginUser""  class=""textbox"" type=""text"" size=""30"" maxlength=""150"" value=""" & LoginUser & """><br>"
		   .echo "       密码参数：<input name=""LoginPass""  class=""textbox"" type=""text"" size=""30"" maxlength=""150"" value=""" & LoginPass & """><br>"
		   .echo "       失败信息：<input name=""LoginFalse""  class=""textbox"" type=""text"" size=""30"" maxlength=""150"" value=""" & LoginFalse & """></td>"
		   .echo "   </tr>"
		   .echo "   <tr class='tdbg'>"
		    .echo "    <td  width=""20%"" height=""25"" align=""center"" class='clefttitle'>项目备注：</td>"
		   .echo "     <td width=""796""><textarea name=""ItemDemo"" cols=""49"" rows=""5"">" & ItemDemo & "</textarea></td>"
		   .echo "   </tr>"
		    .echo "  <tr class='tdbg'>"
		    .echo "    <td height=""35"" colspan=""2"" style=""text-align:center"">"
		    .echo "      <input name=""ItemID"" type=""hidden"" id=""ItemID"" value=""" & ItemID & """>"
		     .echo "     <input name=""Action"" type=""hidden"" id=""Action"" value=""SaveEdit"">"
		     .echo "     <input class='button' name=""Cancel"" type=""button"" id=""Cancel"" value="" 返&nbsp;&nbsp;回 "" onClick=""window.location.href='javascript:history.back();'"">"
		     .echo "     &nbsp;"
		     .echo "   <input  class='button' type=""submit"" name=""Submit"" value=""下&nbsp;一&nbsp;步""></td>"
		     .echo " </tr>"
		  .echo "</form>"
		  .echo "</table>"
		  .echo "</body>"
		  .echo "</html>"
		  .echo "<script>"
		  .echo "function SelectClass(ChannelID)"
		  .echo "{"
		  .echo " if (ChannelID!=0){"
	      .echo "$(parent.document).find(""#ajaxmsg"").toggle();" 
	      .echo "$.get(""../../plus/ajaxs.asp"",{action:""GetClassOption"",channelid:ChannelID},function(data){"
	      .echo "$(parent.document).find(""#ajaxmsg"").toggle();"
	      .echo "$(""select[name=ClassID]"").empty();"
		  .echo "$(""select[name=ClassID]"").append(unescape(data));"
	      .echo " });"
	      .echo "}"
		  .echo "}"
		  .echo "function CheckForm(myform)"
		  .echo "{ if (myform.ItemName.value=='')"
		  .echo "  {"
		  .echo "   alert('请输入项目名称');"
		  .echo "   myform.ItemName.focus();"
		  .echo "   return false;"
		  .echo "  }"
		   .echo "if (myform.ChannelID.value=='0')"
		    .echo "{"
		   .echo "   alert('请选择系统模块!');"
		   .echo "   return false;"
		   .echo " }"
		   .echo "  if (myform.ClassID.value=='0')"
		   .echo " {"
		   .echo "   alert('请选择栏目!');"
		   .echo "   return false;"
		   .echo " }"
		   .echo " if (myform.WebName.value=='')"
		   .echo " {"
		   .echo "  alert('请输入网站名称');"
		   .echo "  myform.WebName.focus();"
		   .echo "  return false;"
		   .echo " }"
		 
		   .echo "return true;"
		  .echo "}"
		  .echo "</script>"
		 End With
		End Sub
End Class
%> 
