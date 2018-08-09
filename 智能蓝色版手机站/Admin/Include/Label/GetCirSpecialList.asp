<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New GetCirSpecialList
KSCls.Kesion()
Set KSCls = Nothing

Class GetCirSpecialList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim TempClassList, InstallDir, CurrPath, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript
		Dim ClassCol, ClassCss, MenuBgType, MenuBg
		Dim ShowClassName, OpenType, Num, IntroLen, TitleLen, RowHeight,SpecialSort, ShowPicFlag, NavType, Navi, MoreLinkType, MoreLink, SplitPic, DateRule, DateAlign, TitleCss, PhotoCss,ShowStyle,PicWidth,PicHeight,PrintType,Col
		Dim ClassPrintType,LabelStyleW,LabelStyle,AjaxOut
		FolderID = Request("FolderID")
		CurrPath = KS.GetCommonUpFilesDir()
		
		With KS
		'判断是否编辑
		LabelID = Trim(Request.QueryString("LabelID"))
		If LabelID = "" Then
		  DateRule="YYYY-MM-DD"
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
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetCirSpecialList", ""),"}" & LabelStyle &"{/Tag}", "")
			
			'response.write labelcontent
			LabelStyleW        = Split(LabelStyle,"§")(0)
			LabelStyle         = Split(LabelStyle,"§")(1)
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('标签加载出错!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
                ClassCol     = Node.getAttribute("classcol")
				ClassCss     = Node.getAttribute("classcss")
				MenuBgType   = Node.getAttribute("menubgtype")
				MenuBg       = Node.getAttribute("menubg")
				Num          = Node.getAttribute("num")
				IntroLen     = Node.getAttribute("introlen")
				TitleLen     = Node.getAttribute("titlelen")
				RowHeight    = Node.getAttribute("rowheight")
				Col          = Node.getAttribute("col")
				OpenType     = Node.getAttribute("opentype")
				NavType      = Node.getAttribute("navtype")
				Navi         = Node.getAttribute("nav")
				MoreLinkType = Node.getAttribute("morelinktype")
				MoreLink     = Node.getAttribute("morelink")
				SplitPic     = Node.getAttribute("splitpic")
				DateRule     = Node.getAttribute("daterule")
				DateAlign    = Node.getAttribute("datealign")
				TitleCss     = Node.getAttribute("titlecss")
				PhotoCss     = Node.getAttribute("photocss")
				ShowStyle    = Node.getAttribute("showstyle")
				PicWidth     = Node.getAttribute("picwidth")
				PicHeight    = Node.getAttribute("picheight")
				PrintType    = Node.getAttribute("printtype")
				AjaxOut      = Node.getAttribute("ajaxout")
                ClassPrintType    = Node.getAttribute("classprinttype")
			End If
			Set Node=Nothing
			Set XMLDoc=Nothing
		End If
		If ShowStyle="" Then ShowStyle=1
		If PrintType="" Then PrintType=2
		If PicWidth="" Then PicWidth=130
		If PicHeight="" Then PicHeight=90
		If Col="" Then Col=1
		If Num = "" Then Num = 10
		If IntroLen = "" Then IntroLen = 200
		If TitleLen = "" Then TitleLen = 30
		If RowHeight= "" Then RowHeight= 22
		If ClassCol = "" Then ClassCol = 2
		If AjaxOut="" Or IsNull(AjaxOut) Then AjaxOut=false
		If ClassPrintType="" Then ClassPrintType=2
		If LabelStyleW="" Then LabelStyleW="<div class=""col"">" & vbcrlf & " <div class=""t""><span><a href=""{@specialclassurl}"" target=""_blank"">更多...</a></span>{@specialclassname}</div>" & vbcrlf & " <ul>{$InnerText}</ul>" & vbcrlf & "</div>"
		If LabelStyle="" Then LabelStyle="[loop={@num}] " & vbcrlf & "<li><a href=""{@specialurl}"" target=""_blank"">{@specialname}</a></li>" & vbcrlf & "[/loop]"
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
		$(document).ready(function(){
		 ChangeClassPrintOutArea($("#ClassPrintType option:selected").val());
		 ChangeOutArea($("#PrintType option:selected").val());
		});
		
		
		function ChangeClassPrintOutArea(Val)
		{
		   if (Val==1)
		   {$("#ClassTable").show();
		    $("#ClassDiy").hide();
		   }else{
		    $("#ClassTable").hide();
		    $("#ClassDiy").show();
		   }
		}
       function InsertLabel(label)
		{
		  InsertValue(label);
		}
		var pos=null;
		var tag=null;
		 function setPos(Tag)
		 {   tag=Tag;
		     if (document.all){
				$("#"+Tag).focus();
				pos = document.selection.createRange();
			  }else{
				pos = document.getElementById("#"+Tag).selectionStart;
			  }
			
		 }
		 //插入
		function InsertValue(Val)
		{  if (pos==null||tag==null) {alert('请先定位要插入的位置!');return false;}
			if (document.all){
				  pos.text=Val;
			}else{
				   var obj=$("#"+tag);
				   var lstr=obj.val().substring(0,pos);
				   var rstr=obj.val().substring(pos);
				   obj.val(lstr+Val+rstr);
			}
		 }		
		
		function ChangeOutArea(Val)
		{
		 if (Val==2){
		  $("#DiyArea").show();
		  $("#TableArea").hide();
		 }
		 else{
		  $("#DiyArea").hide();
		  $("#TableArea").show();
		 }
		}
		function SetMenuBg()
		{if ($("#MenuBgType").val()==0)
		   {
		    $("#MenuBgColor").show();
			$("#MenuBgPic").hide();}
		  else
		  {
		    $("#MenuBgColor").hide();
		    $("#MenuBgPic").show();}
		   }
		function SetNavStatus()
		{
		  if ($("select[name=NavType]").val()==0)
		   {$("#NavWord").show();
			$("#NavPic").hide();
			}else{
		   $("#NavWord").hide();
		   $("#NavPic").show();}
		}
		function SetMoreLinkStatus()
		{
		  if ($("select[name=MoreLinkType]").val()==0){
		    $("#LinkWord").show();
			$("#LinkPic").hide();
			}else{
		   $("#LinkWord").hide();
		   $("#LinkPic").show();}
		}
		function SetLabelFlag(Obj)
		{
		 if (Obj.value=='-1')
		  $("#LabelFlag").val(1);
		  else
		  $("#LabelFlag").val(0);
		}
		function CheckForm()
		{
		    if ($("input[name=LabelName]").val()=='')
			 {
			  alert('请输入标签名称');
			  $("input[name=LabelName]").focus(); 
			  return false
			  }
			var ClassCol=$("#ClassCol").val();
			var ClassCss=$("#ClassCss").val();
			var MenuBgType=1,NavType=1;
			var MenuBg,MenuBgType=$("#MenuBgType").val();
			var OpenType=$("#OpenType").val();
			var Num=$("input[name=Num]").val();
			var IntroLen=$("input[name=IntroLen]").val();
			var TitleLen=$("input[name=TitleLen]").val();
			var Col=$("input[name=Col]").val();
			var RowHeight=$("input[name=RowHeight]").val();
			var Nav,NavType=$("#NavType").val();
			var MoreLink,MoreLinkType=$("#MoreLinkType").val();
			var SplitPic=$("input[name=SplitPic]").val();
			var DateRule=$("#DateRule").val();
			var DateAlign=$("#DateAlign").val();
			var TitleCss=$("input[name=TitleCss]").val();
			var PhotoCss=$("input[name=PhotoCss]").val();
			var ShowStyle=$("#ShowStyle").val();
			var PicWidth=$("input[name=PicWidth]").val();
			var PicHeight=$("input[name=PicHeight]").val();
			var PrintType=$("#PrintType").val();
	    	var ClassPrintType=$("#ClassPrintType").val();
			var AjaxOut=false;
			if ($("#AjaxOut").attr("checked")==true){AjaxOut=true}
			
			if  (Num=='')  Num=10;
			if (IntroLen=='') IntroLen=20
			if  (TitleLen=='') TitleLen=30;
			if  (ClassCol=='') ClassCol=2;
			if  (MenuBgType==0) MenuBg=$("#ColorMenuBg").val()
			 else  MenuBg=$("#PicMenuBg").val();	
			if  (NavType==0) Nav=$("#TxtNavi").val()
			 else  Nav=$("#NaviPic").val();
			if  (MoreLinkType==0) MoreLink=$("#MoreLinkWord").val()
			else  MoreLink=$("#MoreLinkPic").val();
			
			var tagVal='{Tag:GetCirSpecialList labelid="0" classid="0" classprinttype="'+ClassPrintType+'" ajaxout="'+AjaxOut+'" classcol="'+ClassCol+'" classcss="'+ClassCss+'" menubgtype="'+MenuBgType+'" menubg="'+MenuBg+'" num="'+Num+'" introlen="'+IntroLen+'" titlelen="'+TitleLen+'" rowheight="'+RowHeight+'" col="'+Col+'" opentype="'+OpenType+'" navtype="'+NavType+'" nav="'+Nav+'" morelinktype="'+MoreLinkType+'" morelink="'+MoreLink+'" splitpic="'+SplitPic+'" daterule="'+DateRule+'" datealign="'+DateAlign+'" titlecss="'+TitleCss+'" photocss="'+PhotoCss+'" showstyle="'+ShowStyle+'" picwidth="'+PicWidth+'" picheight="'+PicHeight+'" printtype="'+PrintType+'"}';
			tagVal  +=$("#LabelStyleW").val()+'§'+$("#LabelStyle").val()+'{/Tag}';

			$("input[name=LabelContent]").val(tagVal);
			$("#myform").submit();
		}
		</script>
		<%
		.echo "</head>"
		.echo "<body topmargin=""0"" leftmargin=""0"">"
		.echo "<div align=""center"">"
		.echo "<iframe src='about:blank' name='_hiddenframe' style='display:none' id='_hiddenframe' width='0' height='0'></iframe>"
		.echo "<form  method=""post"" id=""myform"" name=""myform"" action=""AddLabelSave.asp"" target='_hiddenframe'>"
		.echo " <input type=""hidden"" name=""LabelContent"" ID=""LabelContent"">"
		.echo " <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""1"">"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo " <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetCirSpecialList.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" colspan='2' height=""24"">&nbsp;&nbsp;&nbsp;&nbsp;<strong>栏目输出格式</strong>&nbsp;"
		.echo " <select class='textbox'  name=""ClassPrintType"" id=""ClassPrintType"" onChange=""ChangeClassPrintOutArea(this.options[this.selectedIndex].value);"">"
        .echo "  <option value=""1"""
		If ClassPrintType=1 Then .echo " selected"
		.echo ">普通(Table)</option>"
        .echo "  <option value=""2"""
		If ClassPrintType=2 Then .echo " selected"
		.echo ">自定义输出样式</option>"
        .echo "</select>"
		.echo "            <font color=green>便于更好的控制,建议选择自定义输出样式</font>"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">采用Ajax输出</label>"

		.echo "</td></tr>"
		
		.echo "         <tbody id=""ClassTable"">"
		.echo "              <tr class='tdbg'>"
		.echo "                <td width=""50%"" align='right' height=""20"">栏目列数"
		.echo "                  <input type=""text"" class=""textbox"" onBlur=""CheckNumber(this,'排列列数');""  style=""width:70%;"" value=""" & ClassCol & """ name=""ClassCol"" id=""ClassCol"">"
		.echo "                </td>"
		.echo "                <td width=""50%"" height=""20"">栏目CSS类名"
						  
		.echo "                <input name=""ClassCss"" class=""textbox"" type=""text"" id=""ClassCss"" value=""" & ClassCss & """></td>"
		.echo "              </tr>"
		.echo "              <tr class='tdbg'>"
		.echo "                <td width=""50%""  align='right' height=""20""> 表头背景"
		.echo "                  <select name=""MenuBgType"" id=""MenuBgType"" class=""textbox"" style=""width:70%;"" onchange=""SetMenuBg()"">"
				  
				  If LabelID = "" Or MenuBgType = "0" Then
					.echo ("<option value=""0"" selected>背景颜色</option>")
					.echo ("<option value=""1"">背景图片</option>")
				   Else
					.echo ("<option value=""0"">背景颜色</option>")
					.echo ("<option value=""1"" selected>背景图片</option>")
				   End If
		.echo "                  </select></td>"
		.echo "                <td width=""50%"" height=""20"">"
				
				If LabelID = "" Or MenuBgType = "0" Then
				  .echo ("<div align=""left"" id=""MenuBgColor""> ")
				  .echo ("<input type=""text"" class=""textbox"" id=""ColorMenuBg"" name=""ColorMenuBg"" style=""width:120;"" value=""" & MenuBg & """>")
				  .echo " <img border=0 id=""ColorMenuBgShow"" src=""../../images/rect.gif"" style=""cursor:pointer;background-Color:" & MenuBg & ";"" onClick=""Getcolor(this,'../../../editor/ksplus/selectcolor.asp?ColorMenuBgShow|ColorMenuBg');"" title=""选取颜色"">"
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=""MenuBgPic"" style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""PicMenuBg"" name=""PicMenuBg"">")
				  .echo ("<input  class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.PicMenuBg);"" name=""Submit3"" value=""选择图片..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.PicMenuBg.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  .echo ("</div>")
				Else
				  .echo ("<div align=""left"" id=""MenuBgColor"" style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""ColorMenuBg"" id=""ColorMenuBg1"" style=""width:120;""> ")
				  .echo " <img border=0 id=""ColorMenuBgShow1"" src=""../../images/rect.gif"" style=""cursor:pointer;background-Color:" & MenuBg & ";"" onClick=""Getcolor(this,'../../../editor/ksplus/selectcolor.asp?ColorMenuBgShow1|ColorMenuBg1');"" title=""选取颜色"">"
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=""MenuBgPic"">")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""PicMenuBg"" name=""PicMenuBg"" value=""" & MenuBg & """>")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.PicMenuBg);"" name=""Submit3"" value=""选择图片..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.PicMenuBg.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  .echo ("</div>")
				End If
				
		.echo "                </td>"
		.echo "              </tr>"
		.echo "              <tr><td colspan=2><hr color=#ff6600 size=1></td></tr>"
		.echo "          </tbody>"
		
	    .echo "           <tbody id=""ClassDiy"">"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='field'>"
		.echo "               <table border='0' width='100%'>"
		.echo "                <tr><td align='center' width='100'><strong>可用标签:</strong></td>"
		.echo "                <td><li onclick=""InsertLabel('{@autoid}')"">行 号</li><li onclick=""InsertLabel('{@specialclassname}')"">专题分类名称</li><li onclick=""InsertLabel('{@specialclassurl}')"">专题分类URL</li><li onclick=""InsertLabel('{@specialclassintro}')"">专题分类介绍(200字)</li><li onclick=""InsertLabel('{@classid}')"">专题分类小ID</li></td>"
		.echo "                 </tr></table>"
		.echo "               </td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'>"
		.echo "                <table border='0' width='100%'><tr><td width='100' align='center'><strong>外循环(分类)</strong><br><font color=blue>必须包含标签{$InnerText}</font></td>"
		.echo "                <td><textarea name='LabelStyleW' onkeyup='setPos(""LabelStyleW"")' onclick='setPos(""LabelStyleW"")' id='LabelStyleW' style='width:100%;height:120px'>" & LabelStyleW & "</textarea></td>"
		.echo "                </tr>"
		.echo "               </table>"
		.echo "             </td>"
		.echo "            </tr>"
		.echo "           </tbody>"		
		
		
		
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">&nbsp;&nbsp;&nbsp;&nbsp;<strong>专题输出格式</strong>&nbsp;"
		.echo " <select class='textbox'  name=""PrintType"" id=""PrintType"" onChange=""ChangeOutArea(this.options[this.selectedIndex].value);"">"
        .echo "  <option value=""1"""
		If PrintType=1 Then .echo " selected"
		.echo ">文本列表样式(Table)</option>"
        .echo "  <option value=""2"""
		If PrintType=2 Then .echo " selected"
		.echo ">自定义输出样式</option>"
        .echo "</select>"
		.echo "             </td> <td><span id='ShowDiyDate'></span> </td>"
		.echo "            </tr>"
		.echo "            <tbody id=""DiyArea"">"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'>"
		.echo "               <table border='0' width='100%'><tr><td width='100' align='center'><strong>可用标签</strong></td>"
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@specialurl}')"">专题链接URL</li> <li onclick=""InsertLabel('{@specialid}')"">专题ID</li><li onclick=""InsertLabel('{@specialname}')"">专题名称</li><li onclick=""InsertLabel('{@specialphotourl}')"">专题图片</li><li onclick=""InsertLabel('{@classid}')"">分类ID</li><li onclick=""InsertLabel('{@specialclassname}')"">分类名称</li><li onclick=""InsertLabel('{@specialclassurl}')"">分类URL</li> <li onclick=""InsertLabel('{@intro}')"">简要介绍</li><li onclick=""InsertLabel('{@photourl}')"">图片地址</li><li onclick=""InsertLabel('{@adddate}')"">添加时间</li><li onclick=""InsertLabel('{@creater}')"">创建人</li></td>"
		.echo "               </tr>"
		.echo "              </table>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'>"
		.echo "               <table border='0' width='100%'><tr><td width='100' align='center'><strong>内循环(专题)</strong></td>"
		.echo "               <td><textarea name='LabelStyle' onkeyup='setPos(""LabelStyle"")' onclick='setPos(""LabelStyle"")' id='LabelStyle' style='width:100%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "               </tr>"
		.echo "              </table>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>使用说明 :</font></strong><br />1、循环标签[loop=n][/loop]对可以省略,也可以平行出现多对；<br /></font>2、支持条件判断语句，格式：<font color=blue>{$IF 条件}</font><font color=red>{成立执行的代码}</font><font color=green>{不成立执行的代码}</font><font color=blue>{/$IF}</font></td>"
		.echo "            </tr>"
		.echo "           </tbody>"	
		
		
	
		.echo "            <tr class='tdbg'>"
		.echo "              <td colspan='2' height=""25"">专题数量"
		.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num2""    style=""width:50px;"" onBlur=""CheckNumber(this,'专题数量');"" value=""" & Num & """> 介绍字数"
		.echo "                <input name=""IntroLen"" class=""textbox"" type=""text"" id=""IntroLen"" style=""width:50px;"" onBlur=""CheckNumber(this,'介绍字数');"" value=""" & IntroLen & """> 行高<input name=""RowHeight"" class=""textbox"" type=""text"" id=""RowHeight"" style=""width:50px;"" onBlur=""CheckNumber(this,'行高');"" value=""" & RowHeight & """> 列数<input name=""Col"" class=""textbox"" type=""text"" id=""Col"" style=""width:50px;"" onBlur=""CheckNumber(this,'列数');"" value=""" & Col & """></td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td width=""50%"" height=""25"">名称字数"
		.echo "                <input name=""TitleLen"" class=""textbox"" onBlur=""CheckNumber(this,'标题字数');"" type=""text""    style=""width:70%;"" value=""" & TitleLen & """>"
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""25"">"
		.echo ReturnOpenTypeStr(OpenType)
		.echo "                 </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""25"">日期格式"
		.echo ReturnDateFormat(DateRule)
		.echo " </td>"
		.echo "              <td height=""25""> <div align=""left"">日期对齐"
		.echo "                  <select name=""DateAlign"" class=""textbox"" id=""DateAlign"" style=""width:70%;"">"
					
					If LabelID = "" Or CStr(DateAlign) = "left" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 .echo ("<option value=""left""" & Str & ">左对齐</option>")
					If CStr(DateAlign) = "center" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 .echo ("<option value=""center""" & Str & ">居中对齐</option>")
					If CStr(DateAlign) = "right" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 .echo ("<option value=""right""" & Str & ">右对齐</option>")
				   
		.echo "                  </select>"
		.echo "                </div></td>"
		.echo "            </tr>"
		
		.echo "       <tbody id=""TableArea"">"
		.echo "             <tr class='tdbg'>"
		.echo "               <td width=""50%"" height=""24"">" &ReturnSpecialStyle(ShowStyle)
		.echo "               </td>"
		.echo "               <td width=""50%"" height=""24"">图片设置 宽"
		.echo "<input name=""PicWidth"" class=""textbox"" type=""text"" id=""PicWidth"" value=""" & PicWidth & """ size=""6"" onBlur=""CheckNumber(this,'图片宽度');"">"
		.echo "                像素 高"
		.echo "<input name=""PicHeight"" class=""textbox"" type=""text"" id=""PicHeight"" value=""" & PicHeight & """ size=""6"" onBlur=""CheckNumber(this,'图片高度');"">"
		.echo "                像素</td>"
		.echo "             </tr>"
		.echo "             <tr class='tdbg'>"
		.echo "               <td height=""24"">名称 CSS"
		.echo "                 <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" onBlur=""CheckBadChar(this,'内容样式');"" value=""" & TitleCss & """></td>"
		.echo "               <td height=""24"">图片 CSS"
		.echo "                 <input name=""PhotoCss"" class=""textbox"" type=""text"" style=""width:70%;"" onBlur=""CheckBadChar(this,'图片样式');"" value=""" & PhotoCss & """></td>"
		.echo "             </tr>"		
		.echo "            <tr class='tdbg'>"
		.echo "              <td width=""50%"" height=""25"">导航类型"
		.echo "                <select name=""NavType"" id=""NavType"" class=""textbox"" style=""width:70%;"" onchange=""SetNavStatus()"">"
				   If LabelID = "" Or CStr(NavType) = "0" Then
					.echo ("<option value=""0"" selected>文字导航</option>")
					.echo ("<option value=""1"">图片导航</option>")
				   Else
					.echo ("<option value=""0"">文字导航</option>")
					.echo ("<option value=""1"" selected>图片导航</option>")
				   End If
				   
		.echo "                </select></td>"
		.echo "              <td width=""50%"" height=""25""> "
				 
				If LabelID = "" Or CStr(NavType) = "0" Then
				  .echo ("<div align=""left"" id=""NavWord""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" id=""TxtNavi"" style=""width:70%;"" value=""" & Navi & """> 支持HTML语法")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"">")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""选择图片..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  .echo ("</div>")
				Else
				  .echo ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" id=""TxtNavi"" style=""width:70%;""> 支持HTML语法")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"" value=""" & Navi & """>")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""选择图片..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  .echo ("</div>")
				End If
		.echo "        </td>"
		.echo "            </tr>"
		.echo "           <tr class='tdbg'>"
		.echo "             <td width=""50%"" height=""25"">更多链接"
		.echo "               <select name=""MoreLinkType"" id=""MoreLinkType"" class=""textbox"" style=""width:70%;"" onchange=""SetMoreLinkStatus()"">"
				  
				  If LabelID = "" Or CStr(MoreLinkType) = "0" Then
					.echo ("<option value=""0"" selected>文字链接</option>")
					.echo ("<option value=""1"">图片链接</option>")
				   Else
					.echo ("<option value=""0"">文字链接</option>")
					.echo ("<option value=""1"" selected>图片链接</option>")
				   End If
				   
		.echo "                </select></td>"
		.echo "              <td width=""50%"" height=""25""> "
				
				If LabelID = "" Or CStr(MoreLinkType) = "0" Then
					.echo ("<div align=""left"" id=""LinkWord""> ")
					.echo ("  <input type=""text"" class=""textbox"" id=""MoreLinkWord"" name=""MoreLinkWord"" style=""width:70%;"" value=""" & MoreLink & """>")
					.echo ("</div>")
					.echo ("<div align=""left"" id=""LinkPic"" style=""display:none""> ")
					.echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""MoreLinkPic"" name=""MoreLinkPic"">")
					.echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.MoreLinkPic);"" name=""Submit3"" value=""选择图片..."">")
					.echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.MoreLinkPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
					.echo ("</div>")
				Else
				   .echo ("<div align=""left"" id=""LinkWord"" style=""display:none""> ")
				   .echo ("<input type=""text"" class=""textbox"" name=""MoreLinkWord"" id=""MoreLinkWord"" style=""width:70%;"">")
				   .echo ("</div>")
				   .echo ("<div align=""left"" id=""LinkPic""> ")
				   .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""MoreLinkPic"" name=""MoreLinkPic"" value=""" & MoreLink & """>")
				   .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.MoreLinkPic);"" name=""Submit3"" value=""选择图片..."">")
				   .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.MoreLinkPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				   .echo ("</div>")
				End If
		.echo "        </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""25"" colspan=""2"">分隔图片"
		.echo "                <input name=""SplitPic"" class=""textbox"" type=""text"" id=""SplitPic"" style=""width:61%;"" value=""" & SplitPic & """ readonly>"
		.echo "                <input class='button' name=""SubmitPic"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.SplitPic);"" type=""button"" id=""SubmitPic2"" value=""选择图片..."">"
		.echo "                <span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.SplitPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>"
		.echo "                <div align=""left""> </div></td>"
		.echo "            </tr>"
		.echo "           </tbody>"
		.echo "                  </table>"	
		.echo "</form>"
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
		End With
		End Sub
End Class
%> 
