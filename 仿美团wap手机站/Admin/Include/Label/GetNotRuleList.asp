<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New GetNotRuleList
KSCls.Kesion()
Set KSCls = Nothing

Class GetNotRuleList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim TempClassList, InstallDir, CurrPath, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript, LabelFlag
		Dim ChannelID,ClassID, IncludeSubClass,  OpenType, ArticleProperty, RowNumber, RowHeight, ShowNumPerRow, OrderStr,ShowPicFlag, NavType, Navi, MoreLinkType, MoreLink, SplitPic,  TitleCss, PrintType,DocProperty,SpecialID,AjaxOut,Attr
		FolderID = Request("FolderID")
		CurrPath = KS.GetCommonUpFilesDir()

		With KS
		'判断是否编辑
		LabelID = Trim(Request.QueryString("LabelID"))
		If LabelID = "" Then
		  ClassID = "0"
		  Action = "Add"
		Else
			Action = "Edit"
		  Dim LabelRS, LabelName
		  Set LabelRS = Server.CreateObject("Adodb.Recordset")
		  LabelRS.Open "Select * From KS_Label Where ID='" & LabelID & "'", Conn, 1, 1
		  If LabelRS.EOF And LabelRS.BOF Then
			 LabelRS.Close
			 Conn.Close
			 Set Conn = Nothing
			 Set LabelRS = Nothing
			 .echo ("<Script>alert('参数传递出错!');window.close();</Script>")
			 Exit Sub
		  End If
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			FolderID = LabelRS("FolderID")
			Descript = LabelRS("Description")
			LabelContent = LabelRS("LabelContent")
			LabelFlag = LabelRS("LabelFlag")
			LabelRS.Close
			Set LabelRS = Nothing


			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetNotRuleList", ""),"}{/Tag}", "")
             'response.write LabelContent
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('标签加载出错!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
			  ChannelID          = Node.getAttribute("modelid")
			  ClassID            = Node.getAttribute("classid")
			  IncludeSubClass    = Node.getAttribute("includesubclass")
			  OpenType           = Node.getAttribute("opentype")
			  DocProperty        = Node.getAttribute("docproperty")
			  attr               = Node.getAttribute("attr")
			  RowNumber          = Node.getAttribute("rownumber")
			  ShowNumPerRow      = Node.getAttribute("shownumperrow")
			  RowHeight          = Node.getAttribute("rowheight")
			  OrderStr           = Node.getAttribute("orderstr")
			  NavType            = Node.getAttribute("navtype")
			  Navi               = Node.getAttribute("nav")
			  MoreLinkType       = Node.getAttribute("morelinktype")
			  MoreLink           = Node.getAttribute("morelink")
			  SplitPic           = Node.getAttribute("splitpic")
			  TitleCss           = Node.getAttribute("titlecss")
			  PrintType          = Node.getAttribute("printtype")
			  AjaxOut            = Node.getAttribute("ajaxout")
            End If
			XMLDoc=Empty
			Set Node=Nothing
		End If
		If PrintType="" Then PrintType=1
		If RowNumber = "" Then RowNumber = 10
		If RowHeight = "" Then RowHeight = 20
		If ShowNumPerRow = "" Then ShowNumPerRow = 60
		If SpecialID=""  Then SpecialID=0
		If DocProperty = "" Or IsNull(DocProperty) Then DocProperty = "00000"
		If AjaxOut="" Then AjaxOut=false
		.echo "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../KS_Inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../KS_Inc/jquery.js"" language=""JavaScript""></script>"
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
		
		function SetStatus(Obj)
		{
		  if (Obj.value=='0')
		   {
			$("#MoreLinkArea").show();
		   }
		   else
		   $("#MoreLinkArea").hide();
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

			var SpecialID=$("select[name=SpecialID]").val();
			if (SpecialID==-1) ClassList=0;
			var OpenType=$("#OpenType").val();
			var ShowNumPerRow=$("#ShowNumPerRow").val();
			var RowHeight=$("#RowHeight").val();
			var RowNumber=$("#RowNumber").val();
			var OrderStr=$("#OrderStr").val();
			var Nav,NavType=$("#NavType").val();
			var MoreLink,MoreLinkType=$("#MoreLinkType").val();
			var SplitPic=$("input[name=SplitPic]").val();
			var TitleCss=$("input[name=TitleCss]").val();
			var PrintType=$("#PrintType").val();
			var AjaxOut=false;
			if ($("#AjaxOut").attr("checked")==true){AjaxOut=true}
			var IncludeSubClass=false;
			if ($("#IncludeSubClass").attr("checked")==true) IncludeSubClass=true;
			
			if  (ShowNumPerRow=='')  ShowNumPerRow=10;
			if (RowHeight=='') RowHeight=20
			if  (RowNumber=='') RowNumber=30;
			if  (NavType==0) Nav=$("input[name=TxtNavi]").val();
			 else  Nav=$("input[name=NaviPic]").val();
			if  (MoreLinkType==0) MoreLink=$("input[name=MoreLinkWord]").val()
			else  MoreLink=$("input[name=MoreLinkPic]").val();
			
			
			var tagVal='{Tag:GetNotRuleList labelid="0" ajaxout="'+AjaxOut+'" modelid="'+ChannelID+'" classid="'+ClassList+'" specialid="'+SpecialID+'" includesubclass="'+IncludeSubClass+'" attr="'+av+'" docproperty="'+DocProperty+'" orderstr="'+OrderStr+'" opentype="'+OpenType+'" rownumber="'+RowNumber+'" shownumperrow="'+ShowNumPerRow+'" rowheight="'+RowHeight+'" titlecss="'+TitleCss+'" navtype="'+NavType+'" nav="'+Nav+'" splitpic="'+SplitPic+'" printtype="'+PrintType+'" morelinktype="'+MoreLinkType+'" morelink="'+MoreLink+'"}{/Tag}';
			$("#LabelContent").val(tagVal);
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
		.echo " <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""" & LabelFlag & """>"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo " <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetNotRuleList.asp"">"
	    .echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">输出格式"
		.echo " <select class='textbox' style='width:70%' name=""PrintType"" id=""PrintType"">"
        .echo "  <option value=""1"""
		If PrintType=1 Then .echo " selected"
		.echo ">普通Table格式</option>"
        .echo "  <option value=""2"""
		If PrintType=2 Then .echo " selected"
		.echo ">LI格式</option>"
        .echo "</select>"
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""24"">"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">采用Ajax输出</label>"
		.echo "              </td>"
		.echo "            </tr>"
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
		.echo "           <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">所属专题"
		.echo "                <select class=""textbox"" onchange=""SpecialChange(this.value)"" style=""width:70%;"" name=""SpecialID"" id=""SpecialID"">"
		.echo "                <option selected value=""-1"" style=""color:red"">- 当前专题(专题页通用)-</option>"
						 If SpecialID = "0" Then
						   .echo ("<option  value=""0"" selected>- 不指定专题 -</option>")
						   Else
						  .echo ("<option  value=""0"">- 不指定专题 -</option>")
						  End If
		.echo KS.ReturnSpecial(SpecialID)
		.echo "</Select>"
		
		.echo "           </td><td>属性控制"
		.echo "                <label><input name=""DocProperty"" type=""checkbox"" value=""1"""
		If mid(DocProperty,1,1) = 1 Then .echo (" Checked")
		.echo ">推荐</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox""  value=""2"""
		If mid(DocProperty,2,1) = 1 Then .echo (" Checked")
		  .echo ">滚动</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""3"""
		If mid(DocProperty,3,1) = 1 Then .echo (" Checked")
		  .echo ">头条</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""4"""
		If mid(DocProperty,4,1) = 1 Then .echo (" Checked")
		  .echo ">热门</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""5"""
		If mid(DocProperty,5,1) = 1 Then .echo (" Checked")
		  .echo ">幻灯</label>"
		.echo "<span id=""showattr""></span> <td>"
		.echo "              </td>"
		.echo "           </tr>"

		.echo "            <tr class='tdbg'>"
		.echo "              <td width=""50%"" height=""24"">显示行数"
		.echo "                <input name=""RowNumber"" class=""textbox"" type=""text"" id=""RowNumber"" style=""width:70%;"" onBlur=""CheckNumber(this,'查询数量');"" value=""" & RowNumber & """></td>"
		.echo "              <td width=""50%"" height=""24"">每行字数"
		.echo "                <input name=""ShowNumPerRow"" class=""textbox"" id=""ShowNumPerRow"" onBlur=""CheckNumber(this,'标题字数');"" type=""text""    style=""width:30%;"" value=""" & ShowNumPerRow & """>&nbsp;<font color=red>包括标题之间的空格</font>"
		.echo "              </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td width=""50%"" height=""24"">文档行距"
		.echo "                <input name=""RowHeight"" class=""textbox"" type=""text"" id=""RowHeight""    style=""width:70%;"" onBlur=""CheckNumber(this,'文档行距');"" value=""" & RowHeight & """></td>"
		.echo "              <td width=""50%"" height=""24"">排序方法"
		.echo "                <select style=""width:70%;"" class='textbox' name=""OrderStr"" id=""OrderStr"">"
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

		.echo "         </select></td>"
		 .echo "           </tr>"

		.echo "            <tr class='tdbg'>"
		.echo "              <td width=""50%"" height=""24"">导航类型"
		.echo "                <select class='textbox' name=""NavType"" id=""NavType"" style=""width:70%;"" onchange=""SetNavStatus()"">"
				   If LabelID = "" Or CStr(NavType) = "0" Then
					.echo ("<option value=""0"" selected>文字导航</option>")
					.echo ("<option value=""1"">图片导航</option>")
				   Else
					.echo ("<option value=""0"">文字导航</option>")
					.echo ("<option value=""1"" selected>图片导航</option>")
				   End If
		 .echo "               </select></td>"
		 .echo "             <td width=""50%"" height=""24"">"
			   If LabelID = "" Or CStr(NavType) = "0" Then
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
		 .echo "           <tr class='tdbg' id=""MoreLinkArea"""
		 .echo ">"
		 .echo "             <td width=""50%"" height=""24"">更多链接"
		 .echo "               <select class='textbox' name=""MoreLinkType"" id=""MoreLinkType"" style=""width:70%;"" onchange=""SetMoreLinkStatus()"">"
				  If LabelID = "" Or CStr(MoreLinkType) = "0" Then
					.echo ("<option value=""0"" selected>文字链接</option>")
					.echo ("<option value=""1"">图片链接</option>")
				   Else
					.echo ("<option value=""0"">文字链接</option>")
					.echo ("<option value=""1"" selected>图片链接</option>")
				   End If
		.echo "                </select></td>"
		.echo "              <td width=""50%"" height=""24"">"
				If LabelID = "" Or CStr(MoreLinkType) = "0" Then
					.echo ("<div align=""left"" id=""LinkWord""> ")
					.echo ("  <input type=""text"" class=""textbox"" name=""MoreLinkWord"" style=""width:70%;"" value=""" & MoreLink & """>")
					.echo ("</div>")
					.echo ("<div align=""left"" id=""LinkPic"" style=""display:none""> ")
					.echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""MoreLinkPic"" name=""MoreLinkPic"">")
					.echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.MoreLinkPic);"" name=""Submit3"" value=""选择图片..."">")
					.echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.MoreLinkPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
					.echo ("</div>")
				Else
				   .echo ("<div align=""left"" id=""LinkWord"" style=""display:none""> ")
				   .echo ("<input type=""text"" class=""textbox"" name=""MoreLinkWord"" style=""width:70%;"">")
				   .echo ("</div>")
				   .echo ("<div align=""left"" id=""LinkPic""> ")
				   .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""MoreLinkPic"" name=""MoreLinkPic"" value=""" & MoreLink & """>")
				   .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.MoreLinkPic);"" name=""Submit3"" value=""选择图片..."">")
				   .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.MoreLinkPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				   .echo ("</div>")
				End If
		.echo "              </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""24"">分隔图片"
		.echo "                <input name=""SplitPic"" class=""textbox"" type=""text"" id=""SplitPic"" style=""width:150px;"" value=""" & SplitPic & """ readonly>"
		.echo "                <input class='button' name=""SubmitPic"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.SplitPic);"" type=""button"" id=""SubmitPic2"" value=""选择图片..."">"
		.echo "                <span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.SplitPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>"
		.echo "              </td><td>  " & ReturnOpenTypeStr(OpenType) & " </td>"
		.echo "            </tr>"
		
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""24"">标题样式"
		.echo "                <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" value=""" & TitleCss & """></td>"
		.echo "              <td height=""24""></td>"
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
