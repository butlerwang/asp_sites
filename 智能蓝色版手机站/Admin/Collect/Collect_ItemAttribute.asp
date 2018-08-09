<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.CollectCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Collect_ItemAttribute
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_ItemAttribute
        Private KS,KSCls
		Private KMCObj
		Private ConnItem
		Private SqlItem, RsItem, Action, FoundErr, ErrMsg
		Private ItemID, ItemName, ChannelID, ClassID, SpecialID
		Private PaginationType, MaxCharPerPage, ReadLevel, Stars, ReadPoint, Hits, UpDateType, UpDateTime, PicNews, Rolls
		Private Comment, Recommend, Popular, FnameType, TemplateID
		Private Script_Iframe, Script_Object, Script_Script, Script_Div, Script_Class, Script_Span, Script_Img, Script_Font, Script_A, Script_Html, Script_Table, Script_Tr, Script_Td
		Private CollecListNum, CollecNewsNum,RepeatInto, IntoBase, BeyondSavePic, CollecOrder, Verific, InputerType, Inputer, EditorType, Editor, ShowComment
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		  Set KMCObj=New CollectPublicCls
		  Set ConnItem = KS.ConnItem()
		End Sub
        Private Sub Class_Terminate()
		 Call KS.CloseConnItem()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		 Set KMCObj=Nothing
		End Sub
		Sub Kesion()
			FoundErr = False
			ItemID = Trim(Request("ItemID"))
			Action = Trim(Request("Action"))
			Verific = 1
			Recommend=1
			IntoBase = 1
			
			If ItemID = "" Then
			   FoundErr = True
			   ErrMsg = ErrMsg & "<br><li>参数错误，项目ID不能为空！</li>"
			Else
			   ItemID = CLng(ItemID)
			End If
			
			If FoundErr <> True Then
				  Call GetTest
			End If
			If FoundErr <> True Then
			   Call Main
			End If
			Response.Write "<script>"
			Response.Write "function SelectClass(ChannelID)"
			Response.Write "{"
			Response.Write " document.all.ClassArea.innerHTML='<select ID=""ClassID"" name=""ClassID"" style=""Width:200"">'+ClassArr[ChannelID]+'</select>';"
			Response.Write "}"
			Response.Write "function CheckForm(myform)"
			Response.Write "{ if (myform.ItemName.value=='')"
			Response.Write "  {"
			Response.Write "   alert('请输入项目名称');"
			Response.Write "   myform.ItemName.focus();"
			Response.Write "   return false;"
			Response.Write "  }"
			Response.Write " if (myform.ChannelID.value=='0')"
			Response.Write "  {"
			Response.Write "    alert('请选择系统模块!');"
			Response.Write "    return false;"
			Response.Write "  }"
			 Response.Write "  if (myform.ClassID.value=='0')"
			Response.Write "  {"
			Response.Write "    alert('请选择栏目!');"
			Response.Write "    return false;"
			Response.Write "  }"
			Response.Write "  if (myform.WebName.value=='')"
			 Response.Write " {"
			Response.Write "   alert('请输入网站名称');"
			Response.Write "   myform.WebName.focus();"
			Response.Write "   return false;"
			Response.Write "  }"
			 Response.Write "return true;"
			Response.Write "}"
			Response.Write "</script>"
			End Sub
			
			Sub Main()
			Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd"">"
			Response.Write "<html xmlns=""http://www.w3.org/1999/xhtml"">"
			Response.Write "<head>"
			Response.Write "<title>采集系统</title>"
			Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Admin_Style.css"">"
			Response.Write "<script language='JavaScript' src='../../KS_Inc/common.js'></script>"
			  Call KMCObj.GetClassList
			Response.Write "<style type=""text/css"">" & vbCrLf
			Response.Write "<!--" & vbCrLf
			Response.Write ".STYLE1 {color: #FF0000}" & vbCrLf
			Response.Write ".STYLE4 {color: #0000FF}" & vbCrLf
			Response.Write "-->" & vbCrLf
			Response.Write "</style>"
			Response.Write "</head>"
			Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Response.Write "<div class='topdashed'>"& KMCObj.GetItemLocation(6,ItemID) &"</div>"

			Response.Write "<table align='center' width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""ctable"">"
			Response.Write "<form method=""post"" action=""Collect_ItemSuccess.asp"" name=""myform"">"
			 Response.Write " <br>"
			 Response.Write "   <tr class='tdbg'>"
			 Response.Write "     <td height=""20"" width=""20%"" align=""right"" class='clefttitle'>项目名称：</td>"
			 Response.Write "     <td><input name='ItemName' class='textbox' type='text' id='ItemName' value='" & ItemName & "' size='27' maxlength='30'></td>"
			 Response.Write "   </tr>"
			  Response.Write "  <tr class='tdbg'>"
			  Response.Write "    <td height=""20"" width=""20%"" align=""right"" class='clefttitle'> 所属模型：</td>"
			   Response.Write "   <td height=""20""><input ID=""ChannelID"" name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"" style=""Width:200"">"
			   Response.Write "   <font color=""red"">" & KS.C_S(ChannelID,1) & "</font>     </td>"
			   Response.Write " </tr>"
			  Response.Write "  <tr class='tdbg'>"
			  Response.Write "    <td height=""20"" width=""20%"" align=""right"" class='clefttitle'> 所属栏目：</td>"
			  Response.Write "    <td height=""20"" ID=""ClassArea""><select name=""ClassID"" ID=""ClassID"" style=""Width:200"">" & Replace(KS.LoadClassOption(ChannelID,true),"value='" & ClassID & "'","value='" & ClassID &"' selected") & "</select>      </td>"
			  Response.Write "  </tr>"
			  Response.Write "  <tr class='tdbg' style=""display:none"">"
			  Response.Write "    <td height=""20"" width=""20%"" align=""right"" class='clefttitle'> 所属专题：</td>"
			  Response.Write "    <td><input type=""hidden"" value=""0"" name=""specialid"">"
			  'call KMCObj.Collect_ShowSpecial_Option(ChannelID,SpecialID)
			  Response.write"     </td>"
			  Response.Write "   </tr>"
			  Response.Write "      <tr class='tdbg'>"
			  Response.Write "    <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>录入时间：</td>"
			   Response.Write "   <td><input name=""UpdateType"" type=""radio"" value=""0"" "
			   If UpDateType = 0 Then Response.Write "checked"
			   Response.Write ">当前时间"
			   Response.Write "    &nbsp;<input name=""UpdateType"" type=""radio"" value=""1"" "
			   If UpDateType = 1 Then Response.Write "checked"
			   Response.Write ">标签中的时间"
			   Response.Write "    &nbsp;<input name=""UpdateType"" type=""radio"" value=""2"" "
			   If UpDateType = 2 Then Response.Write "checked"
			   Response.Write ">自定义："
			   Response.Write "    <input name=""UpdateTime"" class='textbox' type=""text"" value=""" & UpDateTime & """>"
			   Response.Write "   　</td>"
			   Response.Write " </tr>"
			   Response.Write " <tr class='tdbg'>"
			   Response.Write "   <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>录入员：</td>"
			   Response.Write "   <td><input name=""InputerType"" type=""radio"" value=""0"" "
			   If InputerType = 0 Then Response.Write "checked"
			   Response.Write ">当前用户"
			   Response.Write "    &nbsp;<input name=""InputerType"" type=""radio"" value=""1"" "
			   If InputerType = 1 Then Response.Write "checked"
			   Response.Write ">指定用户"
				Response.Write "   <input name=""Inputer"" class='textbox' type=""text"" value=""" & Inputer & """>"
				Response.Write "  　</td>"
			   Response.Write " </tr>"
			   Response.Write " <tr class='tdbg' style='display:none'>"
				Response.Write "  <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>责任编辑：</td>"
			   Response.Write "   <td><input name=""EditorType"" type=""radio"" value=""0"" "
			   If EditorType = 0 Then Response.Write "checked"
			   Response.Write ">当前用户"
			   Response.Write "    &nbsp;<input name=""EditorType"" type=""radio"" value=""1"" "
			   If EditorType = 1 Then Response.Write "checked"
			   Response.Write ">指定用户"
			   Response.Write "    <input name=""Editor"" class='textbox' type=""text"" value=""" & Editor & """>"
			   Response.Write "   　</td>"
			   Response.Write " </tr>"
			   Response.Write " <tr class='tdbg'>"
			   Response.Write "   <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>设置属性：</td>"
			   Response.Write "   <td>"
			  
			   Response.Write "     <input name=""Rolls"" type=""checkbox"" value=""1"" "
			   If Rolls = 1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     滚 动"
			   Response.Write "     <input name=""Comment"" type=""checkbox"" value=""1"" "
			   If Comment = 1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     允许评论"
			   Response.Write "     <input name=""Recommend"" type=checkbox value=""1"" "
			   If Recommend = 1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     推 荐"
			   Response.Write "     <input name=""Popular"" type=""checkbox"" value=""1"" "
			   If Popular = 1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     热 门"
			   
			   'Response.Write "     <input name=""Verific"" type=""checkbox"" value=""1"" "
			   'If Verific = 1 Then Response.Write "checked"
			   'Response.Write "checked"
			   'Response.Write ">已审核"
			Response.Write "</td>"
			 Response.Write "   </tr>"
			 Response.Write "   <tr class='tdbg' style='display:none'>"
			 Response.Write "     <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>标题旁显示评论链接：</td>"
			 Response.Write "     <td>"
			 Response.Write "        <input name=""ShowComment"" type=""radio"" id=""ShowComment"" value=""1"" "
			 If ShowComment = 1 Then Response.Write "Checked"
			 Response.Write ">"
			 Response.Write "        显示"
			 Response.Write "        <input name=""ShowComment"" type=""radio"" id=""ShowComment"" value=""0"" "
			 If ShowComment = 0 Then Response.Write "Checked"
			 Response.Write ">         不显示      </td>"
			 Response.Write "   </tr>"
			 Response.Write "   <tr class='tdbg'>"
			 Response.Write "     <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>内容分页方式：</td>"
			 Response.Write "     <td><select name=""PaginationType"">"
			 Response.Write "           <option value=""0"" "
			 If PaginationType = 0 Then Response.Write "selected"
			 Response.Write ">不分页</option>"
			 Response.Write "           <option value=""1"" "
			 If PaginationType = 1 Then Response.Write "selected"
			 Response.Write ">自动分页</option>"
			 Response.Write "           <option value=""2"" "
			 If PaginationType = 2 Then Response.Write "selected"
			 Response.Write ">采用原文分页</option>"
			 Response.Write "         </select>"
			  Response.Write "      自动分页时的每页大约字符数（包含HTML标记）："
			  Response.Write "<input name=""MaxCharPerPage"" class='textbox' type=""text"" value=""" & MaxCharPerPage & """ size=""8"" maxlength=""8"">      "
			  Response.Write "  </td></tr>"
			  
			  
			  Response.Write "  <tr class='tdbg' style=""display:none"">"
			  Response.Write "    <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>阅读等级：</td>"
			 Response.Write "     <td><input type='hidden' value='0' name='ReadLevel'></td>"
			 Response.Write "   </tr>"
			 Response.Write "     <tr  class='tdbg' style=""display:none"">"
			 Response.Write "     <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>阅读点数：</td>"
			 Response.Write "     <td><input name='ReadPoint' type='text' id='ReadPoint' value='" & ReadPoint & "' size='5' maxlength='3'>"
			 Response.Write "     <font color='#0000FF'>如果大于0，则用户阅读此文章时将消耗相应点数。（对游客和管理员无效）</font>      </td>"
			 Response.Write "   </tr>"
			 
			 
			 
			 
			 Response.Write "   <tr class='tdbg'>"
			 Response.Write "     <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>评分等级：</td>"
			 Response.Write "     <td><select name=""Stars"">"
			 Response.Write "           <option value=""★★★★★"" "
			 If Stars = "★★★★★" Then Response.Write "selected"
			 Response.Write ">★★★★★</option>"
			 Response.Write "           <option value=""★★★★"""
			 If Stars = "★★★★" Then Response.Write "selected"
			 Response.Write ">★★★★</option>"
			 Response.Write "           <option value=""★★★"" "
			 If Stars = "★★★" Then Response.Write "selected"
			 Response.Write ">★★★</option>"
			 Response.Write "           <option value=""★★"" "
			 If Stars = "★★" Then Response.Write "selected"
			 Response.Write ">★★</option>"
			 Response.Write "           <option value=""★"" "
			 If Stars = "★" Then Response.Write "selected"
			 Response.Write ">★</option>"
			 Response.Write "         </select>      </td>"
			 Response.Write "   </tr>"
			  
			 Response.Write "   <tr class='tdbg'>"
			 Response.Write "     <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>点击数初始值：</td>"
			 Response.Write "     <td><input name=""Hits"" class='textbox' type=""text"" id=""Hits"" value=""" & Hits & """ size=""10"" maxlength=""10"">"
			 Response.Write "       <span class=""STYLE4"">用于浏览数作弊</span></td>"
			 Response.Write "   </tr>"
			  
			 Response.Write "   <tr class='tdbg'>"
			 Response.Write "     <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>生成扩展名：</td>"
			 Response.Write "     <td>" & KSCls.GetFsoTypeStr(1)
			 Response.Write "   </td>"
			 Response.Write "   </tr>"
			 	    Dim CurrPath:CurrPath=KS.Setting(3)&KS.Setting(90)
		If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath,Len(CurrPath)-1)

			  Response.Write "   <tr class='tdbg' style='display:none'>"
			  Response.Write "    <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>绑定模板：</td>"
			  Response.Write "    <td><input type='text' size='25' name='TemplateID' id='TemplateID' value='" & templateid & "'> <input type='button' name=""Submit"" class=""button"" value=""选择模板..."" onClick=""OpenThenSetValue('../KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle='+escape('选择模板')+'&CurrPath=" & Server.URLEncode(CurrPath) & "',450,350,window,TemplateID);"">"
			  Response.Write "  </td>"
			  Response.Write "  </tr>"
			  Response.Write "  <tr class='tdbg'>"
			  Response.Write "    <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>标签过滤：</td>"
			  Response.Write "    <td>"
			  Response.Write "      <input name=""Script_Iframe"" type=""checkbox"" value=""yes"" "
			  If Script_Iframe = -1 Then Response.Write "checked"
			  Response.Write ">"
			  Response.Write "      Iframe"
			  Response.Write "      <input name=""Script_Object"" type=""checkbox"" value=""yes"" "
			  If Script_Object = -1 Then Response.Write "checked "
			  Response.Write "onclick='return confirm(""确定要选择该标记吗？这将删除正文中的所有Object标记，结果将导致该文章中的所有动漫动画被删除！"");'>"
			  Response.Write "      Object"
			  Response.Write "      <input name=""Script_Script"" type=""checkbox"" value=""yes"" "
			  If Script_Script = -1 Then Response.Write "checked"
			  Response.Write ">"
			   Response.Write "     Script"
			   Response.Write "     <input name=""Script_Div"" type=""checkbox""  value=""yes"" "
			   If Script_Div = -1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     Div"
			   Response.Write "     <input name=""Script_Class"" type=""checkbox""  value=""yes"" "
			   If Script_Class = -1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     Class"
			   Response.Write "     <input name=""Script_Table"" type=""checkbox""  value=""yes"" "
			   If Script_Table = -1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     Table"
			   Response.Write "     <input name=""Script_Tr"" type=""checkbox""  value=""yes"" "
			   If Script_Tr = -1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     Tr"
			   Response.Write "     <br>"
			   Response.Write "     <input name=""Script_Span"" type=""checkbox""  value=""yes"" "
			   If Script_Span = -1 Then Response.Write "checked"
			   Response.Write ">"
				Response.Write "    Span&nbsp;&nbsp;"
				Response.Write "    <input name=""Script_Img"" type=""checkbox"" value=""yes"" "
				If Script_Img = -1 Then Response.Write "checked"
				Response.Write ">"
				Response.Write "    Img&nbsp;&nbsp;&nbsp;"
				Response.Write "    <input name=""Script_Font"" type=""checkbox""  value=""yes"" "
				If Script_Font = -1 Then Response.Write "checked"
				Response.Write ">"
				 Response.Write "   Font&nbsp;&nbsp;"
				Response.Write "    <input name=""Script_A"" type=""checkbox"" value=""yes"" "
				If Script_A = -1 Then Response.Write "checked"
				Response.Write ">"
				 Response.Write "   A&nbsp;&nbsp;"
				 Response.Write "   <input name=""Script_Html"" type=""checkbox"" value=""yes"" "
				 If Script_Html = -1 Then Response.Write "checked"
				 Response.Write " onclick='return confirm(""确定要选择该标记吗？这将删除正文中的所有Html标记，结果将导致该文章的可阅读性降低！"");'>"
				 Response.Write "   Html&nbsp;"
				 Response.Write "   <input name=""Script_Td"" type=""checkbox""  value=""yes"" "
				 If Script_Td = -1 Then Response.Write "checked"
				 Response.Write ">"
				Response.Write "    Td      </td>"
			   Response.Write " </tr>"
			   Response.Write " <tr class='tdbg'>"
			   Response.Write "   <td height=""22"" width=""20%"" align=""right"" class='clefttitle'> 列表深度：</td>"
			   Response.Write "   <td>"
			   Response.Write "     <input name=""CollecListNum"" class='textbox' type=""text"" id=""CollecListNum"" value=""" & CollecListNum & """ size=""10"" maxlength=""10"">&nbsp;&nbsp;&nbsp;"
			   Response.Write "     <font color='#0000FF'>0为所有的列表</font></td>"
			   Response.Write " </tr>"
			   Response.Write " <tr class='tdbg'>"
			   Response.Write "   <td height=""22"" width=""20%"" align=""right"" class='clefttitle'> 采集信息数量：</td>"
			   Response.Write "   <td>"
			  Response.Write "      <input name=""CollecNewsNum"" class='textbox' type=""text"" id=""CollecNewsNum"" value=""" & CollecNewsNum & """ size=""10"" maxlength=""10"">"
			  Response.Write "      &nbsp;&nbsp;"
			  Response.Write "      <font color='#0000FF'>0为所有的文章<span lang=""en-us"">(</span>每一列表的新闻限制数量<span lang=""en-us"">)</span></font></td>"
			  Response.Write "  </tr>"
			  Response.Write "  <tr class='tdbg'>"
			  Response.Write "    <td height=""22"" width=""20%"" align=""right"" class='clefttitle'> 采集选项：</td>"
			  Response.Write "    <td>"
			   Response.Write "     <input name=""PicNews"" type=""checkbox"" value=""1"" "
			   If PicNews = 1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     自动转换为图片文章"

			 Response.Write "       <input name=""RepeatInto"" type=""checkbox"" value=""1"" "
			  If RepeatInto="1" Then Response.Write "checked"
			  Response.Write ">重复记录入库"
			  Response.Write "      <input name=""BeyondSavePic"" type=""checkbox"" value=""1"" "
			  If BeyondSavePic = 1 Then Response.Write "checked"
			  If KMCObj.IsObjInstalled(KS.Setting(99)) = False Then Response.Write "disabled"
			  Response.Write ">"
			  Response.Write "      保存图片"
			  Response.Write "      <input name=""CollecOrder"" type=""checkbox"" value=""yes"" "
			  If CollecOrder = -1 Then Response.Write "checked"
			  Response.Write ">"
			  Response.Write "      倒序采集        </td>"
			  Response.Write "  </tr>"
			  Response.Write "  <tr class='tdbg'>"
			  Response.Write "    <td height=""22"" width=""20%"" align=""right"" class='clefttitle'> 入库选项：</td>"
			  Response.Write "    <td>"

			  Response.Write "      <input name=""IntoBase"" type=""radio"" value=""0"" "
			  If IntoBase = 0 Then Response.Write "checked"
			  Response.Write ">  不直接入库，需要审核(<font color=red>不推荐</font>)<br/>"
			  Response.Write "      <input name=""IntoBase"" type=""radio"" value=""1"" "
			  If IntoBase = 1 Then Response.Write "checked"
			  Response.Write ">  立即写入主数据库<br/>"
			  Response.Write "      <input name=""IntoBase"" type=""radio"" value=""2"" "
			  If IntoBase = 2 Then Response.Write "checked"
			  Response.Write ">  立即写入主数据库并直接生成内容页<br/>"
	
			  Response.Write "              </td>"
			  Response.Write "  </tr>"
			  
			 Response.Write " <tr class='tdbg'>"
			  Response.Write "  <td height=""30"" width=""20%"" align=""right""></td>"
			  Response.Write "  <td style=""text-align:center""><center>"
			  Response.Write "     <input type=""hidden"" value=""" & ItemID & """ name=""ItemID"">"
			  Response.Write "     <input type=""submit""  class='button' value="" 完&nbsp;&nbsp;成 "" name=""submit"">  </center>      </td>"
			  Response.Write "  </tr>"
			Response.Write "</form>"
			Response.Write "</table>"
			Response.Write "</body>"
			Response.Write "</html>"
			End Sub
			Sub GetTest()
			   SqlItem = "Select top 1 * From KS_CollectItem Where ItemID=" & ItemID
			   Set RsItem = Server.CreateObject("adodb.recordset")
			   RsItem.Open SqlItem, ConnItem, 1, 1
			   If RsItem.EOF And RsItem.BOF Then
				  FoundErr = True
				  ErrMsg = ErrMsg & "<br><li>参数错误，找不到该项目</li>"
			   Else
				  ItemName = RsItem("ItemName")
				  ChannelID = RsItem("ChannelID")
				  ClassID = RsItem("ClassID")
				  SpecialID = RsItem("SpecialID")
				  PaginationType = RsItem("PaginationType")
				  MaxCharPerPage = RsItem("MaxCharPerPage")
				  ReadLevel = RsItem("ReadLevel")
				  Stars = RsItem("Stars")
				  ReadPoint = RsItem("ReadPoint")
				  Hits = RsItem("Hits")
				  UpDateType = RsItem("UpdateType")
				  UpDateTime = RsItem("UpdateTime")
				  PicNews = RsItem("PicNews")
				  Rolls = RsItem("Rolls")
				  Comment = RsItem("Comment")
				  Recommend = RsItem("Recommend")
				  Popular = RsItem("Popular")
				  FnameType = RsItem("FnameType")
				  TemplateID = RsItem("TemplateID")
				  Script_Iframe = RsItem("Script_Iframe")
				  Script_Object = RsItem("Script_Object")
				  Script_Script = RsItem("Script_Script")
				  Script_Div = RsItem("Script_Div")
				  Script_Class = RsItem("Script_Class")
				  Script_Span = RsItem("Script_Span")
				  Script_Img = RsItem("Script_Img")
				  Script_Font = RsItem("Script_Font")
				  Script_A = RsItem("Script_A")
				  Script_Html = RsItem("Script_Html")
				  IntoBase = RsItem("IntoBase")
				  RepeatInto = RsItem("RepeatInto")
				  BeyondSavePic = RsItem("BeyondSavePic")
				  CollecOrder = RsItem("CollecOrder")
				  Verific = RsItem("Verific")
				  CollecListNum = RsItem("CollecListNum")
				  CollecNewsNum = RsItem("CollecNewsNum")
				  InputerType = RsItem("InputerType")
				  Inputer = RsItem("Inputer")
				  EditorType = RsItem("EditorType")
				  Editor = RsItem("Editor")
				  ShowComment = RsItem("ShowComment")
				  Script_Table = RsItem("Script_Table")
				  Script_Tr = RsItem("Script_Tr")
				  Script_Td = RsItem("Script_Td")
			   End If
			   RsItem.Close
			   Set RsItem = Nothing
			End Sub
End Class
%> 
