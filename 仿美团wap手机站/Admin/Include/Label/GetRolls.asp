<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New GetRolls
KSCls.Kesion()
Set KSCls = Nothing

Class GetRolls
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim TempClassList, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript, LabelFlag, PicBorderColor
		Dim ClassID, IncludeSubClass, MarqueeDirection, MarqueeWidth, MarqueeHeight, PicWidth, PicHeight, PicStyle, OpenType, Num, TitleLen, ShowTitle, OrderStr, MarqueeSpeed, TitleCss,SpecialID,DocProperty,Attr
		Dim CurrPath, InstallDir
		Dim ChannelID:ChannelID=KS.G("ChannelID")
		CurrPath = KS.GetCommonUpFilesDir()
		FolderID = Request("FolderID")
		
		With KS
		'判断是否编辑
		LabelID = Trim(KS.G("LabelID"))
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
			LabelContent = Replace(Replace(LabelContent, "{Tag:GetRolls", ""),"}{/Tag}", "")
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
			  MarqueeDirection   = Node.getAttribute("marqueedirection")
			  SpecialID          = Node.getAttribute("specialid")
			  DocProperty        = Node.getAttribute("docproperty")
			  Attr               = Node.getAttribute("attr")
			  PicWidth           = Node.getAttribute("picwidth")
			  PicHeight          = Node.getAttribute("picheight")
			  OrderStr           = Node.getAttribute("orderstr")
			  MarqueeWidth       = Node.getAttribute("marqueewidth")
			  MarqueeHeight      = Node.getAttribute("marqueeheight")
			  OpenType           = Node.getAttribute("opentype")
			  ShowTitle          = Node.getAttribute("showtitle")
			  MarqueeSpeed       = Node.getAttribute("marqueespeed")
			  Num                = Node.getAttribute("num")
			  TitleLen           = Node.getAttribute("titlelen")
			  TitleCss           = Node.getAttribute("titlecss")
			  PicBorderColor     = Node.getAttribute("picbordercolor")
			End If
			Set Node=Nothing
			Set XMLDoc=Nothing		
		End If
		If MarqueeWidth = "" Then MarqueeWidth = 450
		If MarqueeHeight = "" Then MarqueeHeight = 120
		If MarqueeSpeed = "" Then MarqueeSpeed = 30
		If PicWidth = "" Then PicWidth = 130
		If PicHeight = "" Then PicHeight = 90
		If Num = "" Then Num = 10
		If TitleLen = "" Then TitleLen = 30
		If LabelID = "" Then ShowTitle = True
		If SpecialID="" Then SpecialID=0
		If ChannelID="" Then ChannelID=0
		If ShowTitle="" Then ShowTitle=True
		If DocProperty = "" Then DocProperty = "01000"
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
		   SetStatus($("input[name=ShowTitle]:checked").val());
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
		function SetStatus(Value)
		{ 
		 if (Value=='true'|| Value==true)
		  {
		   $("#titleArea").show()
		   }
		 else
		 {
		   $("#titleArea").hide()
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
			var MarqueeDirection=$("#MarqueeDirection").val();
			var OrderStr=$("#OrderStr").val();
			var MarqueeWidth=$("#MarqueeWidth").val();
			var MarqueeHeight=$("#MarqueeHeight").val();
			var OpenType=$("#OpenType").val();
			var PicWidth=$("#PicWidth").val();
			var PicHeight=$("#PicHeight").val();
			var MarqueeSpeed=$("#MarqueeSpeed").val();
			var Num=$("#Num").val();
			var TitleLen=$("#TitleLen").val();
			var TitleCss=$("#TitleCss").val();
			var PicBorderColor=$("#PicBorderColor").val();
			 
			var IncludeSubClass=false;
			if ($("#IncludeSubClass").attr("checked")==true) IncludeSubClass=true;
            var ShowTitle=$("input[name=ShowTitle]:checked").val();
			if  (Num=='')  Num=10;
			if  (TitleLen=='') TitleLen=30;
			var tagVal='{Tag:GetRolls labelid="0" modelid="'+ChannelID+'" classid="'+ClassList+'" specialid="'+SpecialID+'" includesubclass="'+IncludeSubClass+'" attr="'+av+'" docproperty="'+DocProperty+'" orderstr="'+OrderStr+'" marqueewidth="'+MarqueeWidth+'" marqueeheight="'+MarqueeHeight+'" opentype="'+OpenType+'" showtitle="'+ShowTitle+'" picwidth="'+PicWidth+'" picheight="'+PicHeight+'" num="'+Num+'" titlelen="'+TitleLen+'" titlecss="'+TitleCss+'" marqueedirection="'+MarqueeDirection+'" marqueespeed="'+MarqueeSpeed+'" picbordercolor="'+PicBorderColor+'"}{/Tag}';
		 
			$("#LabelContent").val(tagVal);
			$("#myform").submit();
			
		}
		</script>
		<%
		.echo "</head>"
		.echo "<body topmargin=""0"" leftmargin=""0"" onload=""SpecialChange(" & SpecialID &");"" scroll=no>"
		.echo "<div align=""center"">"
		.echo "<iframe src='about:blank' name='_hiddenframe' style='display:none' id='_hiddenframe' width='0' height='0'></iframe>"
		.echo "<form  method=""post"" id=""myform"" name=""myform"" action=""AddLabelSave.asp"" target='_hiddenframe'>"
		.echo " <input type=""hidden"" name=""LabelContent"" id=""LabelContent"">"
		.echo " <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""" & LabelFlag & """>"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """> "
		.echo " <input type=""hidden"" name=""LabelID"" id=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetRolls.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td width=""50%"" height=""24"" colspan=""2"">选择范围"
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
		.echo "            </tr>"
		 .echo "            <tr class='tdbg'>"
		.echo "              <td  width=""50%"" height=""26"">所属专题"
		.echo "                <select class=""textbox"" onchange=""SpecialChange(this.value)"" style=""width:70%;"" name=""SpecialID"" id=""SpecialID"">"
		.echo "                <option selected value=""-1"" style=""color:red"">- 当前专题(专题页通用)-</option>"
						 If SpecialID = "0" Then
						   .echo ("<option  value=""0"" selected>- 不指定专题 -</option>")
						   Else
						  .echo ("<option  value=""0"">- 不指定专题 -</option>")
						  End If
		.echo KS.ReturnSpecial(SpecialID)
		.echo "</Select>"
		.echo " </td>"
		.echo "              <td height=""26"" valign=""top"">属性控制"
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
		  .echo ">幻灯</label>"
		.echo "<span id=""showattr""></span></td>"
		.echo "            </tr>"

		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""26"">滚动方向"
		 .echo "               <select class=""textbox"" name=""MarqueeDirection"" id=""MarqueeDirection"" style=""width:70%;"">"
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
		.echo "              <td height=""26"" valign=""top"">排序方法"
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
		.echo "              <td height=""26"">"
		.echo ReturnOpenTypeStr(OpenType)
		.echo "</td>"
		.echo "              <td height=""26"" valign=""top"">显示标题"
				   
				   If Cbool(ShowTitle) = True Then
					.echo ("<input name=""ShowTitle"" onclick=""SetStatus(true)"" type=""radio"" value=""true"" checked>是")
					.echo ("<input name=""ShowTitle"" onclick=""SetStatus(false)"" type=""radio"" value=""false"">否")
					Else
					  .echo ("<input type=""radio"" onclick=""SetStatus(true)""  value=""true"" name=""ShowTitle"">是")
					  .echo ("<input type=""radio"" onclick=""SetStatus(false)"" value=""false"" name=""ShowTitle"" checked>否")
				   End If
		.echo "        </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""26"">滚动区域 宽度 <input name=""MarqueeWidth"" class=""textbox"" type=""text"" id=""MarqueeWidth"" value=""" & MarqueeWidth & """ size=""6"" onBlur=""CheckNumber(this,'占据宽度');"">像素 高度"
		.echo "                <input name=""MarqueeHeight"" class=""textbox"" type=""text"" id=""MarqueeHeight"" value=""" & MarqueeHeight & """ size=""6"" onBlur=""CheckNumber(this,'占据高度');"">像素"
		.echo " </td>"
		.echo "              <td height=""26"" valign=""top"">图片大小 图片宽度"
		.echo "                <input name=""PicWidth"" class=""textbox"" type=""text"" id=""PicWidth"" value=""" & PicWidth & """ size=""6"" onBlur=""CheckNumber(this,'图片宽度');"">像素  图片高度"
		.echo "                <input name=""PicHeight"" class=""textbox"" type=""text"" id=""PicHeight"" value=""" & PicHeight & """ size=""6"" onBlur=""CheckNumber(this,'图片高度');"">"
		.echo "像素</td>"
		.echo "            </tr>"
		  .echo "              <tr class='tdbg'>"
		  .echo "                <td colspan='2' height=""20"">边框颜色"
		  .echo (" <input type=""text"" class=""textbox"" name=""PicBorderColor"" id=""PicBorderColor"" style=""width:120;"" value=""" & PicBorderColor & """>")
		  .echo (" <img border=0 id=""PicBorderColorShow"" src=""../../images/rect.gif"" style=""cursor:pointer;background-Color:" & PicBorderColor & ";"" onClick=""Getcolor(this,'../../../editor/ksplus/selectcolor.asp?PicBorderColorShow|PicBorderColor');"" title=""选取颜色""> 可留空")
				
		  .echo "                </td>"
		  .echo "              </tr>"
		
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""26"">滚动速度"
		.echo "              <input name=""MarqueeSpeed"" type=""text"" class=""textbox"" id=""MarqueeSpeed""    style=""width:75%;"" onBlur=""CheckNumber(this,'滚动速度');"" value=""" & MarqueeSpeed & """></td>"
		.echo "              <td height=""26"" valign=""top"">列出条数"
		.echo "              <input name=""Num"" class=""textbox"" type=""text"" id=""Num"" style=""width:75%;"" onBlur=""CheckNumber(this,'列出条数');"" value=""" & Num & """></td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg' id='titleArea'>"
		.echo "              <td height=""26"">标题字数"
		.echo "                <input name=""TitleLen"" id=""TitleLen"" class=""textbox"" onBlur=""CheckNumber(this,'标题字数');"" type=""text""    style=""width:75%;"" value=""" & TitleLen & """>              </td>"
		.echo "              <td height=""26"" valign=""top"">标题样式"
		.echo "              <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:75%;"" value=""" & TitleCss & """></td>"
		.echo "            </tr>"
		.echo "         </table>"	
		.echo "  </form>"
		  
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"

		End With
		End Sub
End Class
%> 
