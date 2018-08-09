<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New GetLinkList
KSCls.Kesion()
Set KSCls = Nothing

Class GetLinkList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript,LabelStyle
		Dim show,ClassID, LinkType, ShowStyle, LogoWidth, LogoHeight, ListNumber, TitleLen, ColNumber,RollWidth,RollHeight,RollSpeed,recommend
		FolderID = Request("FolderID")
		With KS
		'判断是否编辑
		LabelID = Trim(Request.QueryString("LabelID"))
		If LabelID = "" Then
		  Action = "Add"
		Else
			Action = "Edit"
		  Dim LabelRS, LabelName
		  Set LabelRS = Server.CreateObject("Adodb.Recordset")
		  LabelRS.Open "Select * From KS_Label Where ID='" & LabelID & "'", Conn, 1, 1
		  If LabelRS.EOF And LabelRS.BOF Then
			 LabelRS.Close
			 Set LabelRS = Nothing
			 .echo ("<Script>alert('参数传递出错!');window.close();</Script>")
			 .End
		  End If
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			FolderID = LabelRS("FolderID")
			Descript = LabelRS("Description")
			LabelContent = LabelRS("LabelContent")
			LabelRS.Close
			Set LabelRS = Nothing
			LabelStyle         = KS.GetTagLoop(LabelContent)
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetLinkList", ""),"}" & LabelStyle &"{/Tag}", "")
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('标签加载出错!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
				show = Node.getAttribute("show")
				ClassID = Node.getAttribute("classid")
			    LinkType = Node.getAttribute("linktype")
				ShowStyle = Node.getAttribute("showstyle")
				LogoWidth = Node.getAttribute("logowidth")
				LogoHeight = Node.getAttribute("logoheight")
				ListNumber = Node.getAttribute("num")
				TitleLen = Node.getAttribute("titlelen")
				recommend=Node.getAttribute("recommend")
				ColNumber = Node.getAttribute("col")
				RollWidth = Node.getAttribute("rollwidth")
				RollHeight= Node.getAttribute("rollheight")
				RollSpeed = Node.getAttribute("rollspeed")
		   End If
		   Set Node=Nothing
		   XMLDoc=Empty
		End If
		If Show="" Then show=0
		If LinkType = "" Then LinkType = 1
		If ShowStyle = "" Then ShowStyle = 2
		If LogoWidth = "" Then LogoWidth = 88
		If LogoHeight = "" Then LogoHeight = 31
		If RollWidth = "" Then LogoWidth = 200
		If RollHeight = "" Then LogoHeight = 150
		If ListNumber = "" Then ListNumber = 0
		If TitleLen = "" Then TitleLen = 30
		If ColNumber = "" Then ColNumber = 7
		If RollSpeed="" Then RollSpeed=5
		If recommend="" Then recommend=0
		If LabelStyle="" Then LabelStyle="[loop={@num}] " & vbcrlf & "<li><a href=""{@url}"">{@sitename}</a></li>" & vbcrlf & "[/loop]"
		.echo "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../ks_inc/Jquery.js"" language=""JavaScript""></script>"
		%>
		<style type="text/css">
		 .field{width:720px;}
		 .field li{cursor:pointer;float:left;border:1px solid #DEEFFA;background-color:#F7FBFE;height:18px;line-height:18px;margin:3px 1px 0px;padding:2px}
		 .field li.diyfield{border:1px solid #f9c943;background:#FFFFF6}
		</style>
		<script language="javascript">
		function SetLogoDisabled(Num)
		{
		if (Num==0)
		{
		 $("input[name=LogoWidth]").attr("disabled",true);
		 $("input[name=LogoHeight]").attr("disabled",true);
		}
		else
		{
		 $("input[name=LogoWidth]").attr("disabled",false);
		 $("input[name=LogoHeight]").attr("disabled",false);
		}
		}
		function SetDisabled(Num)
		{
		 if (Num==1||Num==3)
		 {
		  $("input[name=ColNumber]").attr("disabled",true);
		 }
		 else
		  {
		  $("input[name=ColNumber]").attr("disabled",false);
		  }
		 if (Num==1)
		 { $("#RollArea").show();
		 }else{
		   $("#RollArea").hide();
		 }
		 if (Num==4){
		 $("#DiyArea").show();
		 }else{
		 $("#DiyArea").hide();
		 }
		}
		   function InsertLabel(label)
		{
		  InsertValue(label);
		}
		var pos=null;
		 function setPos()
		 { if (document.all){
				$("#LabelStyle").focus();
				pos = document.selection.createRange();
			  }else{
				pos = document.getElementById("LabelStyle").selectionStart;
			  }
		 }
		 //插入
		function InsertValue(Val)
		{  if (pos==null) {alert('请先定位要插入的位置!');return false;}
			if (document.all){
				  pos.text=Val;
			}else{
				   var obj=$("#LabelStyle");
				   var lstr=obj.val().substring(0,pos);
				   var rstr=obj.val().substring(pos);
				   obj.val(lstr+Val+rstr);
			}
		 }
		function CheckForm()
		{
		    if ($("input[name=LabelName]").val()=='')
			 {
			  alert('请输入标签名称');
			  $("input[name=LabelName]").focus(); 
			  return false
			  }
			var show,LinkType,ShowStyle;
			var ClassID=$("#ClassID").val();
			var LogoWidth=$("input[name=LogoWidth]").val();
			var LogoHeight=$("input[name=LogoHeight]").val();
			var RollWidth=$("input[name=RollWidth]").val();
			var RollHeight=$("input[name=RollHeight]").val();
			var RollSpeed=$("input[name=RollSpeed]").val();
			var ListNumber=$("input[name=ListNumber]").val();
			var TitleLen=$("input[name=TitleLen]").val();
			var ColNumber=$("input[name=ColNumber]").val();
			var show=$("input[name='show']:checked").val();
			var LinkType=$("input[name='LinkType']:checked").val();
			var ShowStyle=$("input[name='ShowStyle']:checked").val();
			var recommend=0;
			if ($("#recommend").attr("checked")==true)
			{
			  recommend=1;
			}
			var tagVal='{Tag:GetLinkList labelid="0" show="'+show+'" classid="'+ClassID+'" linktype="'+LinkType+'" showstyle="'+ShowStyle+'" logowidth="'+LogoWidth+'" logoheight="'+LogoHeight+'" rollwidth="'+RollWidth+'" rollheight="'+RollHeight+'" rollspeed="'+RollSpeed+'" num="'+ListNumber+'" titlelen="'+TitleLen+'" recommend="'+recommend+'" col="'+ColNumber+'"}'+$("#LabelStyle").val()+'{/Tag}'
			$("input[name=LabelContent]").val(tagVal);
			$("#myform").submit();
		}
		</script>
		<%
		.echo "</head>"
		.echo "<body topmargin=""0"" leftmargin=""0"" scroll=no>"
		.echo "<div align=""center"">"
		.echo "<iframe src='about:blank' name='_hiddenframe' style='display:none' id='_hiddenframe' width='0' height='0'></iframe>"
		.echo "<form  method=""post"" id=""myform"" name=""myform"" action=""AddLabelSave.asp"" target='_hiddenframe'>"
		.echo " <input type=""hidden"" name=""LabelContent"" id=""LabelConent"">"
		.echo "   <input type=""hidden"" name=""LabelFlag"" value=""2"">"
		.echo "  <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo "  <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo "  <input type=""hidden"" name=""FileUrl"" value=""GetLinkList.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"" colspan=""2"">链接类别&nbsp;"
		
					  .echo ("<Select Name=""ClassID"" id=""ClassID"" Class=""textbox"">")
					  .echo ("<option Value=""0"">-列出所有分类的站点-</option>")
					  Dim ObjRS
					  Set ObjRS = Server.CreateObject("Adodb.Recordset")
					  ObjRS.Open "Select * From KS_LinkFolder Order BY OrderID,AddDate Desc", Conn, 1, 1
					  Do While Not ObjRS.EOF
					  If ClassID = Trim(ObjRS("FolderID")) Then
					   .echo ("<option value=" & ObjRS("FolderID") & " Selected>" & ObjRS("FolderName") & "</Option>")
					  Else
					   .echo ("<option value=" & ObjRS("FolderID") & ">" & ObjRS("FolderName") & "</Option>")
					  End If
					   ObjRS.MoveNext
					  Loop
					  ObjRS.Close
					  Set ObjRS = Nothing
					  .echo ("</Select>")
					   
		.echo "               </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"" colspan=""2""> 显示路径"

						If show = "0" Then
						 .echo ("<input type=""radio"" name=""show"" value=""0"" Checked>直接链接URL ")
						Else
						 .echo ("<input type=""radio"" name=""show"" value=""0"">直接链接URL ")
						End If
						If show = "1" Then
						 .echo ("<input type=""radio"" name=""show"" value=""1"" Checked>通过ToLink.asp转向(可累计点击次数) ")
						Else
						 .echo ("<input type=""radio"" name=""show"" value=""1"">通过asp链接转向(可累计点击次数) ")
						End If
						 
		.echo "                 </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"" colspan=""2""> 链接类型"
						
						If LinkType = 2 Then
						 .echo ("<input type=""radio"" onclick=""SetLogoDisabled(2)"" name=""LinkType"" value=""2"" Checked>全部链接 ")
						Else
						 .echo ("<input type=""radio"" onclick=""SetLogoDisabled(2)"" name=""LinkType"" value=""2"">全部链接 ")
						End If
						If LinkType = 0 Then
						 .echo ("<input type=""radio"" onclick=""SetLogoDisabled(0)"" name=""LinkType"" value=""0"" Checked>文本链接 ")
						Else
						 .echo ("<input type=""radio"" onclick=""SetLogoDisabled(0)"" name=""LinkType"" value=""0"">文本链接 ")
						End If
						If LinkType = 1 Then
						 .echo ("<input type=""radio"" onclick=""SetLogoDisabled(1)"" name=""LinkType"" value=""1"" Checked>LOGO链接 ")
						Else
						 .echo ("<input type=""radio"" onclick=""SetLogoDisabled(1)"" name=""LinkType"" value=""1"">LOGO链接 ")
						End If
						 
		.echo "                 </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"" colspan=""2"">显示方式"
						
						If ShowStyle = 1 Then
						 .echo ("<input type=""radio"" onclick=""SetDisabled(1)"" name=""ShowStyle"" value=""1"" Checked>向上滚动 ")
						Else
						 .echo ("<input type=""radio"" onclick=""SetDisabled(1)"" name=""ShowStyle"" value=""1"">向上滚动 ")
						End If
						If ShowStyle = 2 Then
						 .echo ("<input type=""radio"" onclick=""SetDisabled(2)"" name=""ShowStyle"" value=""2"" Checked>横向列表 ")
						Else
						 .echo ("<input type=""radio"" onclick=""SetDisabled(2)"" name=""ShowStyle"" value=""2"">横向列表 ")
						End If
						If ShowStyle = 3 Then
						 .echo ("<input type=""radio"" onclick=""SetDisabled(3)"" name=""ShowStyle"" value=""3"" Checked>下拉列表 ")
						Else
						 .echo ("<input type=""radio"" onclick=""SetDisabled(3)"" name=""ShowStyle"" value=""3"">下拉列表 ")
						End If
						If ShowStyle = 4 Then
						 .echo ("<input type=""radio"" onclick=""SetDisabled(4)"" name=""ShowStyle"" value=""4"" Checked>自定义输出样式<span style='color:red'>(新增)</span> ")
						Else
						 .echo ("<input type=""radio"" onclick=""SetDisabled(4)"" name=""ShowStyle"" value=""4"">自定义输出样式<span style='color:red'>(新增)</span> ")
						End If
		 .echo "              </td>"
		 .echo "           </tr>"
		 
		 .echo "<tbody id=""DiyArea"">"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">行 号</li><li onclick=""InsertLabel('{@url}')"">网址</li> <li onclick=""InsertLabel('{@linkid}')"">链接ID</li><li onclick=""InsertLabel('{@sitename}')"">网站名称</li><li onclick=""InsertLabel('{@logo}')"">Logo地址</li><li onclick=""InsertLabel('{@hits}')"">点击数</li><li onclick=""InsertLabel('{@description}')"">介绍</li><li onclick=""InsertLabel('{@adddate}')"">添加时间</li></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'><textarea name='LabelStyle' onkeyup='setPos()' onclick='setPos()' id='LabelStyle' style='width:95%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>使用说明 :</font></strong><br />循环标签[loop=n][/loop]对可以省略,也可以平行出现多对；</td>"
		.echo "            </tr>"
		.echo "           </tbody>"
		 
		 .echo "          <tbody id=""RollArea"">"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"" colspan='2'>区域 : 滚动宽度"
		 .echo "               <input name=""RollWidth"" id=""RollWidth"" class=""textbox""  onBlur=""CheckNumber(this,'滚动宽度');"" type=""text"" id=""RollWidth"" style=""width:50px;"" value=""" & RollWidth & """>滚动高度"
		 .echo "               <input name=""RollHeight""  class=""textbox""  onBlur=""CheckNumber(this,'滚动高度');"" type=""text"" id=""RollHeight"" style=""width:50px;"" value=""" & RollHeight & """> 滚动速度<input name=""RollSpeed"" class=""textbox""  onBlur=""CheckNumber(this,'滚动速度');"" type=""text"" id=""RollSpeed"" style=""width:50px;"" value=""" & RollSpeed & """></td>"
		 .echo "           </tr>"
		 .echo "         </tbody>"
		 
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">Logo宽度"
		 .echo "               <input name=""LogoWidth"" id=""LogoWidth"" class=""textbox""  onBlur=""CheckNumber(this,'Logo宽度');"" type=""text"" id=""LogoWidth"" style=""width:70%;"" value=""" & LogoWidth & """></td>"
		 .echo "             <td height=""30"">Logo高度"
		 .echo "               <input name=""LogoHeight"" id=""LogoHeight"" class=""textbox""  onBlur=""CheckNumber(this,'Logo高度');"" type=""text"" id=""LogoHeight"" style=""width:70%;"" value=""" & LogoHeight & """></td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td width=""50%"" height=""30"">显示数目"
		 .echo "               <input name=""ListNumber"" class=""textbox""  onBlur=""CheckNumber(this,'公告条数');"" type=""text"" id=""ListNumber"" style=""width:100px;"" value=""" & ListNumber & """><font color=""#FF0000"">设置为0时将列出所有友情链接站点</font></td>"
		 .echo "             <td width=""50%"" height=""30""><label><input type='checkbox' value='1' id='recommend' name='recommend'"
		 if recommend="1" then .echo " checked"
		 .echo ">仅显示推荐</label></td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""24"">标题字数"
		 .echo "               <input name=""TitleLen"" class=""textbox""  onBlur=""CheckNumber(this,'标题字数');"" type=""text"" id=""TitleLen"" style=""width:70%;"" value=""" & TitleLen & """></td>"
		 .echo "             <td height=""24""> 显示列数"
		 .echo "               <input name=""ColNumber"" class=""textbox""  onBlur=""CheckNumber(this,'显示列数');"" type=""text"" id=""ColNumber"" style=""width:70%;"" value=""" & ColNumber & """></td>"
		 .echo "           </tr>"
		.echo "                  </table>"	
		 .echo " </form>"
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
		.echo "<script>"
		.echo "SetLogoDisabled(" & LinkType & ");"
		.echo "SetDisabled(" & ShowStyle & ");"
		.echo "</script>"
		End With
		End Sub
End Class
%> 
