<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New GetLastSpecialList
KSCls.Kesion()
Set KSCls = Nothing

Class GetLastSpecialList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		'主体部分
		Public Sub Kesion()
		Dim TempSpecialList, InstallDir, CurrPath, FolderID, LabelContent, Action, LabelID, Str, Descript, LabelFlag,ColNumber
		Dim PerPageNumber, OpenType, IntroLen, TitleLen, NavType, Navi, SplitPic, DateRule, DateAlign, TitleCss, PhotoCss,PageStyle,ShowStyle,PicWidth,PicHeight,PrintType,AjaxOut,LabelStyle,RowHeight
		FolderID = Request("FolderID")
		CurrPath = KS.GetCommonUpFilesDir()
		With KS
		'判断是否编辑
		LabelID = Trim(Request.QueryString("LabelID"))
		If LabelID = "" Then
		  Action = "Add":DateRule="YYYY-MM-DD"
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
			LabelFlag = LabelRS("LabelFlag")
			LabelRS.Close
			Set LabelRS = Nothing
			LabelStyle         = KS.GetTagLoop(LabelContent)
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetLastSpecialList", ""),"}" & LabelStyle&"{/Tag}", "")
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('标签加载出错!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
				AjaxOut          = Node.getAttribute("ajaxout")
				PrintType        = Node.getAttribute("printtype")
			    OpenType         = Node.getAttribute("opentype")
			    PerPageNumber    = Node.getAttribute("num")
			    IntroLen         = Node.getAttribute("introlen")
			    TitleLen         = Node.getAttribute("titlelen")
			    ColNumber        = Node.getAttribute("col")
				RowHeight        = Node.getAttribute("rowheight")
			    NavType          = Node.getAttribute("navtype")
			    Navi             = Node.getAttribute("nav")
			    PageStyle        = Node.getAttribute("pagestyle")
			    SplitPic         = Node.getAttribute("splitpic")
			    DateRule         = Node.getAttribute("daterule")
			    DateAlign        = Node.getAttribute("datealign")
			    TitleCss         = Node.getAttribute("titlecss")
			    PhotoCss         = Node.getAttribute("photocss")
			    ShowStyle        = Node.getAttribute("showstyle")
			    PicWidth         = Node.getAttribute("picwidth")
			    PicHeight        = Node.getAttribute("picheight")
				
			 End If
	   			 Set Node=Nothing
			 Set XMLDoc=Nothing
		End If

		If ShowStyle="" Then ShowStyle=1
		If PrintType="" Then PrintType=1
		If PicWidth="" Then PicWidth=130
		If PicHeight="" Then PicHeight=90
		If PageStyle="" Then PageStyle=1
		If ColNumber="" Then ColNumber=1
		If PerPageNumber = "" Then PerPageNumber = 10
		If IntroLen = "" Then IntroLen = 200
		If TitleLen = "" Then TitleLen = 30
		If RowHeight= "" Then RowHeight= 20
		If AjaxOut="" Then AjaxOut=false
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
		   ChangeOutArea();
		 })
		
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
		 
		function ChangeOutArea(Val)
		{
		  var Val=$("#PrintType").val();
		  if (Val==2){
		   $("#DiyArea").show();
		  }else{
		  $("#DiyArea").hide();
		  }
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
		function CheckForm()
		{
		    if ($("input[name=LabelName]").val()=='')
			 {
			  alert('请输入标签名称');
			  $("input[name=LabelName]").focus(); 
			  return false
			  }
			var PerPageNumber=$("input[name=PerPageNumber]").val(); 
			var OpenType=$("#OpenType").val();
			var IntroLen=$("input[name=IntroLen]").val();
			var TitleLen=$("input[name=TitleLen]").val();
			var ColNumber=$("input[name=ColNumber]").val();
			var RowHeight=$("input[name=RowHeight]").val();
			var Nav,NavType=$("#NavType").val();
			var SplitPic=$("input[name=SplitPic]").val();
			var DateRule=$("#DateRule").val();
			var DateAlign=$("#DateAlign").val();
			var TitleCss=$("input[name=TitleCss]").val();
			var PhotoCss=$("input[name=PhotoCss]").val();
			var PageStyle=$("#PageStyle").val();
			var ShowStyle=$("#ShowStyle").val();
			var PicWidth=$("input[name=PicWidth]").val();
			var PicHeight=$("input[name=PicHeight]").val();
			var PrintType=$("#PrintType").val();
			var AjaxOut=false;
			if ($("#AjaxOut").attr("checked")==true){AjaxOut=true}
			if (IntroLen=='') IntroLen=20
			if  (TitleLen=='') TitleLen=30;
			if  (NavType==0) Nav=$("#TxtNavi").val();
			 else  Nav=$("#NaviPic").val();
			 
			var tagVal='{Tag:GetLastSpecialList labelid="0" pagestyle="'+PageStyle+'" printtype="'+PrintType+'" ajaxout="'+AjaxOut+'" num="'+PerPageNumber+'" opentype="'+OpenType+'" titlelen="'+TitleLen+'" introlen="'+IntroLen+'" rowheight="'+RowHeight+'" col="'+ColNumber+'" navtype="'+NavType+'" nav="'+Nav+'" splitpic="'+SplitPic+'" daterule="'+DateRule+'" datealign="'+DateAlign+'" titlecss="'+TitleCss+'" photocss="'+PhotoCss+'" picwidth="'+PicWidth+'" picheight="'+PicHeight+'" showstyle="'+ShowStyle+'"}'+$("#LabelStyle").val()+'{/Tag}';

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
		.echo " <input type=""hidden"" name=""LabelContent"" id=""LabelContent"">"
		.echo " <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""1"">"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo " <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetLastSpecialList.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">输出格式"
		.echo " <select class='textbox' style='width:70%' name=""PrintType"" id=""PrintType"" onChange=""ChangeOutArea();"">"
        .echo "  <option value=""1"""
		If PrintType=1 Then .echo " selected"
		.echo ">普通Table格式</option>"
        .echo "  <option value=""2"""
		If PrintType=2 Then .echo " selected"
		.echo ">自定义输出样式</option>"
        .echo "</select>"
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""24"">"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">采用Ajax输出</label></td>"
		.echo "            </tr>"
		
		
		.echo "            <tbody id=""DiyArea"">"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">行 号</li><li onclick=""InsertLabel('{@specialurl}')"">专题链接URL</li> <li onclick=""InsertLabel('{@specialid}')"">专题ID</li><li onclick=""InsertLabel('{@specialname}')"">专题名称</li><li onclick=""InsertLabel('{@specialphotourl}')"">专题图片</li><li onclick=""InsertLabel('{@classid}')"">分类ID</li><li onclick=""InsertLabel('{@specialclassname}')"">分类名称</li><li onclick=""InsertLabel('{@specialclassurl}')"">分类URL</li> <li onclick=""InsertLabel('{@intro}')"">简要介绍</li><li onclick=""InsertLabel('{@photourl}')"">图片地址</li><li onclick=""InsertLabel('{@adddate}')"">添加时间</li><li onclick=""InsertLabel('{@creater}')"">创建人</li></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'><textarea name='LabelStyle' onkeyup='setPos()' onclick='setPos()' id='LabelStyle' style='width:95%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>使用说明 :</font></strong><br />循环标签[loop=n][/loop]对可以省略,也可以平行出现多对；</td>"
		.echo "            </tr>"
		.echo "           </tbody>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""28"">每页数量"
		.echo "                <input class=""textbox"" name=""PerPageNumber"" type=""text"" id=""PerPageNumber""    style=""width:70%;"" onBlur=""CheckNumber(this,'每页数量');"" value=""" & PerPageNumber & """></td>"
		.echo "              <td height=""28"">"
					   
		.echo ReturnOpenTypeStr(OpenType)
		
		.echo "         </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td width=""50%"" height=""25"">" &ReturnPageStyle(PageStyle)
		.echo "              <td>排列列数<input type=""text"" class=""textbox"" onBlur=""CheckNumber(this,'排列列数');"" size=5 value=""" & ColNumber & """ name=""ColNumber"" id=""ColNumber""> 行距<input type=""text"" class=""textbox"" onBlur=""CheckNumber(this,'行距');"" size=5 value=""" & RowHeight & """ name=""RowHeight"" id=""RowHeight""><font color=red>若自定义输出样式无效</font></td>"
		.echo "            </tr>"
		
						.echo "            <tr class='tdbg'>"
		.echo "              <td height=""28"">日期格式"
		.echo ReturnDateFormat(DateRule)
		.echo " </td>"
		.echo "              <td height=""28""> <div align=""left"">日期对齐"
		.echo "                  <select class=""textbox"" name=""DateAlign"" id=""DateAlign"" style=""width:70%;"">"
					  
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

		
		
		.echo "            <tr class='tdbg'>"
		.echo "              <td width=""50%"" height=""28"">名称字数"
		.echo "                <input class=""textbox"" name=""TitleLen"" onBlur=""CheckNumber(this,'标题字数');"" type=""text""    style=""width:70%;"" value=""" & TitleLen & """>"
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""28"">介绍字数"
		.echo "                <input class=""textbox"" name=""IntroLen"" type=""text"" id=""IntroLen2""    style=""width:70%;"" onBlur=""CheckNumber(this,'介绍字数');"" value=""" & IntroLen & """></td>"
		.echo "            </tr>"
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
		.echo "             <tr class='tdbg'>"
		.echo "               <td width=""50%"" height=""24"">导航类型"
		.echo "                 <select class=""textbox"" name=""NavType"" id=""NavType"" style=""width:70%;"" onchange=""SetNavStatus()"">"
				   If LabelID = "" Or NavType = "0" Then
					.echo ("<option value=""0"" selected>文字导航</option>")
					.echo ("<option value=""1"">图片导航</option>")
				   Else
					.echo ("<option value=""0"">文字导航</option>")
					.echo ("<option value=""1"" selected>图片导航</option>")
				   End If
				   
		.echo "                 </select></td>"
		.echo "               <td width=""50%"" height=""24"">"
				
				If LabelID = "" Or CStr(NavType) = "0" Then
				  .echo ("<div align=""left"" id=""NavWord""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" id=""TxtNavi"" style=""width:70%;"" value=""" & Navi & """ onBlur=""CheckBadChar(this,'文字导航');""> 支持HTML语法")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"">")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""选择图片..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  .echo ("</div>")
				Else
				  .echo ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" id=""TxtNavi"" style=""width:70%;"" onBlur=""CheckBadChar(this,'文字导航');""> 支持HTML语法")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"" value=""" & Navi & """>")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""选择图片..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  .echo ("</div>")
				End If
		.echo "         </td>"
		.echo "             </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "             <td height=""28"" colspan=""2"">分隔图片"
		.echo "                <input name=""SplitPic"" class=""textbox"" type=""text"" id=""SplitPic"" style=""width:61%;"" value=""" & SplitPic & """ readonly>"
		.echo "                <input class='button' name=""SubmitPic"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.SplitPic);"" type=""button"" id=""SubmitPic2"" value=""选择图片..."">"
		.echo "                <span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.SplitPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>"
		.echo "                            </td>"
		.echo "            </tr>"

		.echo "                  </table>"	
		.echo "    </form>"
		 
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
		End With
		End Sub
End Class
%> 
