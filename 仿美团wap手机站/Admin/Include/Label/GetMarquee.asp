<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New GetMarqueeArticle
KSCls.Kesion()
Set KSCls = Nothing

Class GetMarqueeArticle
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim TempClassList, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript, LabelFlag
		Dim ChannelID,ClassID, IncludeSubClass, MarqueeWidth, MarqueeHeight, DateRule, OpenType, Num, MarqueeBgcolor, MarqueeDirection, TitleLen, MarqueeStyle, OrderStr, MarqueeSpeed, TitleCss, DateCss,SpecialID,DocProperty,NavType,Navi,CurrPath,InstallDir,Attr
		FolderID = Request("FolderID")
		ChannelID=KS.ChkCLng(Request("ChannelID"))
		CurrPath = KS.GetCommonUpFilesDir()
		With KS
		LabelID = Trim(Request.QueryString("LabelID"))
		If LabelID = "" Then
		  ClassID = "0":DateRule="YYYY-MM-DD"
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
			LabelFlag = LabelRS("LabelFlag")
			LabelRS.Close
			Set LabelRS = Nothing
			LabelContent = Replace(Replace(LabelContent, "{Tag:GetMarquee", ""),"}{/Tag}", "")
			'response.write LabelContent
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('标签加载出错!');</Script>")
			 response.End()
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
			  ChannelID          = Node.getAttribute("modelid")
			  ClassID            = Node.getAttribute("classid")
			  IncludeSubClass    = Node.getAttribute("includesubclass")
			  SpecialID          = Node.getAttribute("specialid")
			  DocProperty        = Node.getAttribute("docproperty")
			  Attr               = Node.getAttribute("attr")
			  MarqueeWidth       = Node.getAttribute("marqueewidth")
			  MarqueeHeight      = Node.getAttribute("marqueeheight")
			  MarqueeSpeed       = Node.getAttribute("marqueespeed")
			  MarqueeDirection   = Node.getAttribute("marqueedirection")
			  OpenType           = Node.getAttribute("opentype")
			  OrderStr           = Node.getAttribute("orderstr")
			  TitleLen           = Node.getAttribute("titlelen")
			  MarqueeStyle       = Node.getAttribute("marqueestyle")
			  Num                = Node.getAttribute("num")
			  DateRule           = Node.getAttribute("daterule")
			  MarqueeBgcolor     = Node.getAttribute("marqueebgcolor")
			  TitleCss           = Node.getAttribute("titlecss")
			  DateCss            = Node.getAttribute("datecss")
			  NavType            = Node.getAttribute("navtype")
			  Navi               = Node.getAttribute("nav")
			End If
			Set XMLDoc=Nothing
			Set Node=Nothing
		End If
		If MarqueeWidth = "" Then MarqueeWidth = 450
		If MarqueeHeight = "" Then MarqueeHeight = 20
		If Num = "" Then Num = 10
		If TitleLen = "" Then TitleLen = 30
		If MarqueeSpeed = "" Then MarqueeSpeed = 30
		If LabelID = "" Then MarqueeStyle = 0
		If SpecialID="" Then SpecialID=0
		If ChannelID="" Then ChannelID=0
		If DocProperty = "" Or IsNull(DocProperty) Then DocProperty = "01000"
		If NavType="" Or IsNull(NavType) Then NavType="0"
		.echo "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../ks_inc/jQuery.js"" language=""JavaScript""></script>"
		%>
		<script language="javascript">
		$(document).ready(function(){
		 $("#ChannelID").change(function(){
		    $(parent.document).find('#ajaxmsg').toggle();
			GetAttribute($(this).val());
			$.get('../../../plus/ajaxs.asp',{action:'GetClassOption',channelid:$(this).val()},function(data){
			  $("#ClassList").empty();
			  $("#ClassList").append("<option value='-1' style='color:red'>-当前栏目(通用)-</option>");
			  $("#ClassList").append("<option value='0'>-不指定栏目-</option>");
			  $("#ClassList").append(unescape(data));
			  $(parent.document).find('#ajaxmsg').toggle();
			 })
		   })	
		  $("#MutileClass").click(function(){
		    if ($(this).attr("checked")==true){
		      $("#ClassList").attr("multiple","multiple");
		      $("#ClassList").attr("style","height:60px");
		    }else{
			   $("#ClassList").removeAttr("multiple");
			}
		  });
		   <%if Instr(ClassID,",")<>0 Then%>
		   var searchStr="<%=ClassID%>";
		   $("#MutileClass").attr("checked",true);
		   $("#ClassList").attr("multiple","multiple");
		   $("#ClassList").attr("style","height:60px");
		   setTimeout(function(){ 
		   $("#ClassList>option").each(function(){
		     if($(this).val()=='-1' || $(this).val()=='0')
			  $(this).attr("selected",false)
			 else if (searchStr.indexOf($(this).val())!=-1)
			 { 
			   $(this).attr("selected",true);
			 }
		   });},1);
		  <%end if%>
          <%If LabelID<>"" Then%>
		   GetAttribute($("#ChannelID").val());
		  <%End If%>
		});
		function GetAttribute(channelid){
		    $.get('../../../plus/ajaxs.asp',{action:'GetModelAttr',attr:'<%=attr%>',channelid:channelid},function(data){
			  $("#showattr").html('').html(data)
			 });
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
		function SetDisabled(Num)
		{
		 if (Num==0)
		 $("#MarqueeDirection").attr("disabled",false);
		 else{
		 $("#MarqueeDirection").attr("disabled",true);
		 }
		}
		function SetLabelFlag(Obj)
		{
		 if (Obj.value=='-1')
		  $("#LabelFlag").val(1);
		  else
		  $("#LabelFlag").val(0);
		}
		function SpecialChange(SpecialID)
		{
			if (SpecialID==-1) 
			  $("#ClassArea").hide();
			else
			  $("#ClassArea").show();	
		}
		function CheckForm()
		{  
		
		   if ($("input[name=LabelName]").val()=='')
			 {
			  alert('请输入标签名称');
			  $("input[name=LabelName]").focus(); 
			  return false
			  }
			var ChannelID=$("#ChannelID").val();
			var ClassList='';
		    if ($("#MutileClass").attr("checked")==true){
				$("#ClassList option:selected").each(function(){
					if ($(this).val()!='0' && $(this).val()!='-1')
						if (ClassList=='') 
						 ClassList=$(this).val() 
						else
						 ClassList+=","+$(this).val();
					})
			 }else{
			    ClassList=$("#ClassList").val();
			 }
			var SpecialID=$("select[name=SpecialID]").val();
			if (SpecialID==-1) ClassList=0;
			var DocProperty='';
			 $("input[name=DocProperty]").each(function(){
			     if ($(this).attr("checked")==true){
				  DocProperty=DocProperty+'1'
				 }else{
				  DocProperty=DocProperty+'0'
				 }      
			 })
			 var av='';
		   $("input[name=attr]").each(function(){
		     if ($(this).attr("checked")==true){
			   if (av==''){
			    av=$(this).val();
			   }else{
			    av+='|'+$(this).val();
			   }
			 }
		   });
			var MarqueeWidth=$("#MarqueeWidth").val();
			var MarqueeHeight=$("#MarqueeHeight").val();
			var MarqueeSpeed=$("#MarqueeSpeed").val();
			var MarqueeDirection=$("#MarqueeDirection").val();
			var OpenType=$("#OpenType").val();
			var OrderStr=$("#OrderStr").val();
			var TitleLen=$("#TitleLen").val();
			var MarqueeStyle;
			var Num=$("#Num").val();
			var DateRule=$("#DateRule").val();
			var MarqueeBgcolor=$("#MarqueeBgcolor").val();
			var TitleCss=$("#TitleCss").val();
			var DateCss=$("#DateCss").val();
			var IncludeSubClass=false;
			if ($("#IncludeSubClass").attr("checked")==true) IncludeSubClass=true;
			var MarqueeStyle=$("input[name=MarqueeStyle]:checked").val();
			var Nav,NavType=$("select[name=NavType]").val();
			if  (NavType==0) Nav=$("input[name=TxtNavi]").val();
			 else  Nav=$("input[name=NaviPic]").val();
			if  (Num=='')  Num=10;
			if  (TitleLen=='') TitleLen=30;
			if  (MarqueeSpeed=='') MarqueeSpeed=5;
			
			var tagVal='{Tag:GetMarquee labelid="0" marqueetype="text" modelid="'+ChannelID+'" classid="'+ClassList+'" specialid="'+SpecialID+'" includesubclass="'+IncludeSubClass+'" attr="'+av+'" docproperty="'+DocProperty+'" orderstr="'+OrderStr+'" marqueewidth="'+MarqueeWidth+'" marqueeheight="'+MarqueeHeight+'" opentype="'+OpenType+'" num="'+Num+'" titlelen="'+TitleLen+'" titlecss="'+TitleCss+'" datecss="'+DateCss+'" daterule="'+DateRule+'" marqueedirection="'+MarqueeDirection+'" marqueespeed="'+MarqueeSpeed+'" marqueebgcolor="'+MarqueeBgcolor+'" marqueestyle="'+MarqueeStyle+'" navtype="'+NavType+'" nav="'+Nav+'"}{/Tag}';
		 
			$("#LabelContent").val(tagVal);
			$("#myform").submit();
			
		}
		</script>
		<%
		.echo "</head>"
		.echo "<body topmargin=""0"" leftmargin=""0"" onload=""SpecialChange(" & SpecialID &");"">"
		.echo "<div align=""center"">"
		.echo "<iframe src='about:blank' name='_hiddenframe' style='display:none' id='_hiddenframe' width='0' height='0'></iframe>"
		.echo "<form  method=""post"" id=""myform"" name=""myform"" action=""AddLabelSave.asp"" target='_hiddenframe'>"
		.echo " <input type=""hidden"" name=""LabelContent"" id=""LabelContent"">"
		.echo "  <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""" & LabelFlag & """>"
		.echo " <input type=""hidden"" name=""Action"" id=""Action"" value=""" & Action & """>"
		.echo "  <input type=""hidden"" name=""LabelID"" id=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetMarquee.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
        .echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td colspan=""2"" height=""24"">选择范围"
		.echo "                <select name=""ChannelID"" id=""ChannelID"">"
		.echo "                 <option value=""0"">-所有模型-</option>"
        .LoadChannelOption ChannelID
		.echo "                </select>"
		.echo "                <select class=""textbox"" name=""ClassList"" id=""ClassList"" onChange=""SetLabelFlag(this)"">"
		.echo "                 <option selected value=""-1"" style=""color:red"">- 当前栏目(通用)-</option>"
						
						If ClassID = "0" Then
						   .echo ("<option  value=""0"" selected>- 不指定栏目 -</option>")
						Else
						  .echo ("<option  value=""0"">- 不指定栏目 -</option>")
					   End If
						  .echo Replace(KS.LoadClassOption(ChannelID,false),"value='" & ClassID & "'","value='" & ClassID &"' selected")
						  .echo "</select>"

						  
					If cbool(IncludeSubClass) = True Or LabelID = "" Then
					  Str = " Checked"
					Else
					  Str = ""
					End If
					  .echo "<input type='checkbox' name='MutileClass' id='MutileClass' value='1'>指定多栏目"
					  .echo ("<input name=""IncludeSubClass"" type=""checkbox"" id=""IncludeSubClass"" value=""true""" & Str & ">")
			
		.echo "                  调用子栏目</div></td>"
		.echo "</td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"">所属专题"
		.echo "                <select class=""textbox"" onchange=""SpecialChange(this.value)"" style=""width:70%;"" name=""SpecialID"" id=""SpecialID"">"
		.echo "                <option selected value=""-1"" style=""color:red"">- 当前专题(专题页通用)-</option>"
						 If SpecialID = "0" Then
						   .echo ("<option  value=""0"" selected>- 不指定专题 -</option>")
						   Else
						  .echo ("<option  value=""0"">- 不指定专题 -</option>")
						  End If
		.echo KS.ReturnSpecial(SpecialID)
		.echo "</Select>"
		.echo "                 </td>"
		.echo "              <td height=""30"">属性控制"
		.echo "                <label><input name=""DocProperty"" type=""checkbox"" value=""1"""
		If mid(DocProperty,1,1) = 1 Then .echo (" Checked")
		.echo ">推荐</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" Checked disabled value=""2"">滚动</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""3"""
		If mid(DocProperty,3,1) = 1 Then .echo (" Checked")
		  .echo ">头条</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""4"""
		If mid(DocProperty,4,1) = 1 Then .echo (" Checked")
		  .echo ">热门</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""5"""
		If mid(DocProperty,5,1) = 1 Then .echo (" Checked")
		  .echo ">幻灯</label><span id=""showattr""></span></td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"">滚动速度"
		.echo "                <input type=""text"" class=""textbox"" onBlur=""CheckNumber(this,'滚动速度');""  style=""width:70%;"" value=""" & MarqueeSpeed & """ id=""MarqueeSpeed"" name=""MarqueeSpeed"">"
		.echo "                　　</td>"
		.echo "              <td>滚动方向"
		.echo "                <select class=""textbox"" name=""MarqueeDirection"" id=""MarqueeDirection"" style=""width:160;"">"
					  
					   If MarqueeDirection = "left" Then
						.echo ("<option value=""left"" selected>向左滚动</option>")
					   Else
						.echo ("<option value=""left"">向左滚动</option>")
					   End If
					   If MarqueeDirection = "right" Then
						.echo ("<option value=""right"" selected>向右滚动</option>")
					   Else
						.echo ("<option value=""right"">向右滚动</option>")
					   End If
					   If MarqueeDirection = "up" Then
						.echo ("<option value=""up"" selected>向上滚动</option>")
						Else
						.echo ("<option value=""up"">向上滚动</option>")
						End If
						If MarqueeDirection = "down" Then
						.echo ("<option value=""down"" selected>向下滚动</option>")
						Else
						.echo ("<option value=""down"">向下滚动</option>")
						End If
					   
		.echo "                </select></td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"">"
				 
		.echo ReturnOpenTypeStr(OpenType)
		
		 .echo "               　</td>"
		 .echo "             <td>排序方法"
		 .echo "                <select class='textbox' name='OrderStr' id='OrderStr' style=""width:75%;"">"
					
					If OrderStr = "ID Desc" Then
					.echo ("<option value='ID Desc' selected>文档ID(降序)</option>")
					Else
					.echo ("<option value='ID Desc'>文档ID(降序)</option>")
					End If
					If OrderStr = "ID Asc" Then
					.echo ("<option value='ID Asc' selected>文档ID(升序)</option>")
					Else
					.echo ("<option value='ID Asc'>文档ID(升序)</option>")
					End If
					If OrderStr = "Rnd" Then
					.echo ("<option value='Rnd' style='color:blue' selected>随机显示</option>")
					Else
					.echo ("<option value='Rnd' style='color:blue'>随机显示</option>")
					End If
					
					If OrderStr = "ModifyDate Asc" Then
					.echo ("<option value='ModifyDate Asc' selected>修改时间(升序)</option>")
					Else
					.echo ("<option value='ModifyDate Asc'>修改时间(升序)</option>")
					End If
					If OrderStr = "ModifyDate Desc" Then
					 .echo ("<option value='ModifyDate Desc' selected>修改时间(降序)</option>")
					Else
					 .echo ("<option value='ModifyDate Desc'>修改时间(降序)</option>")
					End If
					If OrderStr = "AddDate Asc" Then
					.echo ("<option value='AddDate Asc' selected>添加时间(升序)</option>")
					Else
					.echo ("<option value='AddDate Asc'>添加时间(升序)</option>")
					End If
					If OrderStr = "AddDate Desc" Then
					 .echo ("<option value='AddDate Desc' selected>添加时间(降序)</option>")
					Else
					 .echo ("<option value='AddDate Desc'>添加时间(降序)</option>")
					End If
					If OrderStr = "Hits Asc" Then
					 .echo ("<option value='Hits Asc' selected>点击数(升序)</option>")
					Else
					 .echo ("<option value='Hits Asc'>点击数(升序)</option>")
					End If
					If OrderStr = "Hits Desc" Then
					  .echo ("<option value='Hits Desc' selected>点击数(降序)</option>")
					Else
					  .echo ("<option value='Hits Desc'>点击数(降序)</option>")
					End If
				   
		.echo "                </select></td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"">标题字数"
		.echo "                <input name=""TitleLen"" id=""TitleLen"" class=""textbox"" onBlur=""CheckNumber(this,'标题字数');"" type=""text""    style=""width:70%;"" value=""" & TitleLen & """>"
		.echo "              </td>"
		.echo "              <td>滚动方式"
				  
					 If MarqueeStyle = 0 Then
					  .echo ("<input type=""radio"" onclick=""SetDisabled(0);"" value=""0"" name=""MarqueeStyle"" checked>默认方式 ")
					  .echo ("<input type=""radio"" onclick=""SetDisabled(1);alert('提示:纵向滚动建议高度设为18px');"" value=""1"" name=""MarqueeStyle"">纵向间隔滚动 ")
					 Else
					  .echo ("<input type=""radio"" onclick=""SetDisabled(0);"" value=""0"" name=""MarqueeStyle"">默认方式 ")
					  .echo ("<input type=""radio"" onclick=""SetDisabled(1);alert('提示:纵向滚动建议高度设为18px');"" value=""1"" name=""MarqueeStyle"" checked>纵向间隔滚动 ")
					 End If
		   
		 .echo "             </td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">文档数量"
		 .echo "               <input class=""textbox"" name=""Num"" type=""text"" id=""Num""  style=""width:70%;"" onBlur=""CheckNumber(this,'文档数量');"" value=""" & Num & """>"
		 .echo "             </td>"
		 .echo "             <td>日期格式"
		 .echo ReturnDateFormat(DateRule)
		 .echo "               </td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">背景颜色"
		 .echo "               <input name=""MarqueeBgcolor"" class=""textbox"" type=""text"" style=""width:50px;"" id=""MarqueeBgcolor"" value=""" & MarqueeBgcolor & """><img border=0 id=""MarqueeBgcolorShow"" src=""../../images/rect.gif"" style=""cursor:pointer;background-Color:" & MarqueeBgcolor & ";"" onClick=""Getcolor(this,'../../../editor/ksplus/selectcolor.asp?MarqueeBgcolorShow|MarqueeBgcolor');"" title=""选取颜色""></td>"
		 .echo "<td> 宽高设置         宽度<input name=""MarqueeWidth"" class=""textbox"" type=""text"" id=""MarqueeWidth"" value=""" & MarqueeWidth & """ size=""6"" onBlur=""CheckNumber(this,'占据宽度');"">像素"
		.echo "                高度<input name=""MarqueeHeight"" class=""textbox"" type=""text"" id=""MarqueeHeight"" value=""" & MarqueeHeight & """ size=""6"" onBlur=""CheckNumber(this,'占据高度');"">像素</td>"
		 .echo "           </tr>"
.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">导航类型"
		.echo "                <select name=""NavType"" style=""width:70%;"" class='textbox' onchange=""SetNavStatus()"">"
				   If NavType = "0" Then
					.echo ("<option value=""0"" selected>文字导航</option>")
					.echo ("<option value=""1"">图片导航</option>")
				   Else
					.echo ("<option value=""0"">文字导航</option>")
					.echo ("<option value=""1"" selected>图片导航</option>")
				   End If
		 .echo "               </select></td>"
		 .echo "             <td width=""50%"" height=""24"">"
			   If NavType = "0" Then
				  .echo ("<div align=""left"" id=""NavWord""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" style=""width:70%;"" value=""" & Navi & """> 支持HTML语法")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"">")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""选择图片..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  .echo ("</div>")
				Else
				  .echo ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" style=""width:70%;""> 支持HTML语法")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"" value=""" & Navi & """>")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""选择图片..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  .echo ("</div>")
				End If
		 .echo "             </td>"
		 .echo "           </tr>"		 
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">标题样式"
		 .echo "               <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" value=""" & TitleCss & """></td>"
		 .echo "             <td>日期样式"
		 .echo "               <input name=""DateCss"" class=""textbox"" type=""text"" id=""DateCss"" style=""width:70%;"" value=""" & DateCss & """>"
		 .echo "              </td>"
		 .echo "           </tr>"

		 .echo "                  </table>"	
		 .echo " </form>"
		  
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
		.echo "<script>"
		.echo "SetDisabled(" & MarqueeStyle & ");"
		.echo "</script>"
		End With
		End Sub
End Class
%> 
