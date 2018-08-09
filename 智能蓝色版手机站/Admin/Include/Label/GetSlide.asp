<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="LabelFunction.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 9.0
' Email: service@EasyTool.CN . QQ:111394,9537636
' Web: http://www.EasyTool.CN http://www.KeSion.cn
' Copyright (C) KeSion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New GetSlide
KSCls.KeSion()
Set KSCls = Nothing

Class GetSlide
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub KeSion()
		Dim TempClassList, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript, LabelFlag
		Dim ClassID, IncludeSubClass, PicWidth, PicHeight, Num, OpenType, ShowTitle,  TitleLen, TitleCss, ChangeTime,SlideType,SpecialID,DocProperty,From,Attr
		FolderID = Request("FolderID")
		Dim ChannelID:ChannelID=KS.G("ChannelID")
		With KS
		
		'判断是否编辑
		LabelID = Trim(Request.QueryString("LabelID"))
		From =KS.S("From")
		If LabelID = "" Then
		  ClassID = "0"
		  Action = "Add"
		Else
			Action = "Edit"
		  Dim LabelRS, LabelName
		  Set LabelRS = Server.CreateObject("Adodb.Recordset")
		  LabelRS.Open "Select TOP 1 * From KS_Label Where ID='" & LabelID & "'", Conn, 1, 1
		  If LabelRS.EOF And LabelRS.BOF Then
			 LabelRS.Close
			 Set LabelRS = Nothing
			 .echo ("<Script>alert('参数传递出错!');window.close();</Script>")
			 .End
		  End If
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			Descript = LabelRS("Description")
			FolderID = LabelRS("FolderID")
			LabelContent = LabelRS("LabelContent")
			LabelFlag = LabelRS("LabelFlag")
			LabelRS.Close
			Set LabelRS = Nothing
			LabelContent = Replace(Replace(LabelContent, "{Tag:GetSlide", ""),"}{/Tag}", "")
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
			  ChannelID          = KS.ChkClng(Node.getAttribute("modelid"))
			  If ChannelID=-1000 Then From="club"
			  ClassID            = Node.getAttribute("classid")
			  IncludeSubClass    = Node.getAttribute("includesubclass")
			  SpecialID          = Node.getAttribute("specialid")
			  PicWidth           = Node.getAttribute("picwidth")
			  PicHeight          = Node.getAttribute("picheight")
			  Num                = Node.getAttribute("num")
			  OpenType           = Node.getAttribute("opentype")
			  ShowTitle          = Node.getAttribute("showtitle")
			  TitleLen           = Node.getAttribute("titlelen")
			  TitleCss           = Node.getAttribute("titlecss")
			  ChangeTime         = Node.getAttribute("changetime")
			  SlideType          = Node.getAttribute("slidetype")
			  DocProperty        = Node.getAttribute("docproperty")
			  Attr               = Node.getAttribute("attr")
		   End If
		   Set Node=Nothing
		   Set XMLDoc=Nothing
		End If
		If ChannelID="" Then ChannelID=0
		If Num = "" Then Num = 5
		If TitleLen = "" Then TitleLen = 30
		If PicWidth = "" Then PicWidth = 200
		If PicHeight = "" Then PicHeight = 200
		If ChangeTime = "" Then ChangeTime = 5000
		If SlideType=0 Then SlideType=2
		If SpecialID="" Then SpecialID=0
		If ShowTitle="" Then ShowTitle=true
		If DocProperty = "" Then DocProperty = "00001"
		.echo "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../KS_Inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../KS_Inc/jQuery.js"" language=""JavaScript""></script>"
		%>
		<script language="javascript">
		$(document).ready(function(){
		 $("#ChannelID").change(function(){
		 
		    $(parent.document).find('#ajaxmsg').toggle();
			GetAttribute($(this).val());
			$.get('../../../plus/ajaxs.asp',{action:'GetClassOption',channelid:$(this).val()},function(data){
			  $("#ClassList").empty().append("<option value='-1' style='color:red'>-当前栏目(通用)-</option>").append("<option value='0'>-不指定栏目-</option>").append(unescape(data));
			  $(parent.document).find('#ajaxmsg').toggle();
			 })
		   })	
		  $("#MutileClass").click(function(){
		    if ($(this).attr("checked")==true){
		      $("#ClassList").attr("multiple","multiple").attr("style","height:60px");
		    }else{
			   $("#ClassList").removeAttr("multiple");
			}
		  });
		  $("#SlideType>option[value=<%=SlideType%>]").attr("selected",true);
		   <%if Instr(ClassID,",")<>0 Then%>
		   var searchStr="<%=ClassID%>";
		   $("#MutileClass").attr("checked",true);
		   $("#ClassList").attr("multiple","multiple").attr("style","height:60px");
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
		{   if ($("input[name=LabelName]").val()=='')
			 {
			  alert('请输入标签名称');
			  $("input[name=LabelName]").focus(); 
			  return false
			  }
		<%If From="club" then%>
		    var ChannelID=-1000;
			var ClassList='0';
			var SpecialID='0';
			var DocProperty='000000';
		<%Else%>	  
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
		<%End If%>
			var PicWidth=$("input[name=PicWidth]").val();
			var PicHeight=$("input[name=PicHeight]").val();
			var Num=$("input[name=Num]").val();
			var OpenType=$("#OpenType").val();
			var TitleLen=$("input[name=TitleLen]").val();
			var TitleCss=$("input[name=TitleCss]").val();
			var ChangeTime=$("input[name=ChangeTime]").val();
			var SlideType=$("#SlideType").val();
			var IncludeSubClass=false;
			if ($("#IncludeSubClass").attr("checked")==true) IncludeSubClass=true;
		    
			var ShowTitle=$("input[name=ShowTitle]:checked").val();
			if  (Num=='')  Num=10;
			if  (TitleLen=='') TitleLen=30;
			
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
		   
			var tagVal='{Tag:GetSlide labelid="0" modelid="'+ChannelID+'" classid="'+ClassList+'" specialid="'+SpecialID+'" includesubclass="'+IncludeSubClass+'" attr="'+av+'" docproperty="'+DocProperty+'" picwidth="'+PicWidth+'" picheight="'+PicHeight+'" num="'+Num+'" opentype="'+OpenType+'" showtitle="'+ShowTitle+'" titlelen="'+TitleLen+'" titlecss="'+TitleCss+'" changetime="'+ChangeTime+'" slidetype="'+SlideType+'"}{/Tag}';
		 
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
		.echo " <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""" & LabelFlag & """> "
		.echo " <input type=""hidden"" name=""Action"" id=""Action"" value=""" & Action & """>"
		.echo "  <input type=""hidden"" name=""LabelID"" id=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetSlide.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		If From="club" Then
		.echo "            <tr class=tdbg>"
		.echo "              <td  height=""24"" colspan=""4"" style=""text-align:center""><strong>小论坛幻灯片调用标签</strong></td></tr>"
		Else
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
		.echo "              <td height=""30"">所属专题"
		.echo "                <select class=""textbox"" onchange=""SpecialChange(this.value)"" style=""width:35%;"" name=""SpecialID"" id=""SpecialID"">"
		.echo "                <option selected value=""-1"" style=""color:red"">- 当前专题(专题页通用)-</option>"
						 If SpecialID = "0" Then
						   .echo ("<option  value=""0"" selected>- 不指定专题 -</option>")
						   Else
						  .echo ("<option  value=""0"">- 不指定专题 -</option>")
						  End If
		.echo KS.ReturnSpecial(SpecialID)
		.echo "</Select>"
        .echo "</td>"
		.echo "              <td width=""50%"" height=""24"">属性控制"
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
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""5"" checked disabled>幻灯</label>"
		
		.echo " <span id=""showattr""></span> </td>"
		.echo "</tr>"
	End If
		.echo "  <tr class='tdbg'>"
		.echo "              <td height=""30"" colspan=""2"">幻灯类型"
		.echo " <select name=""SlideType"" id=""SlideType"">"
		.echo "<option value=""1"">普通JS幻灯</option>"
		.echo "<option value=""2"">flash幻灯1</option>"
		.echo "<option value=""3"">flash幻灯2(Sina)</option>"
		.echo "<option value=""4"">flash幻灯3(SOHU)</option>"
		.echo "<option value=""5"" style=""color:red"">flash幻灯4(推荐)</option>"
		.echo "</select>"
					
        .echo " </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"">查询条数"
		.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num""    style=""width:60px;text-align:center"" onBlur=""CheckNumber(this,'图片数量');"" value=""" & Num & """> 条</td>"
		.echo "              <td height=""30"">图片大小 宽"
		.echo "                <input name=""PicWidth"" class=""textbox"" type=""text"" id=""PicWidth2"" value=""" & PicWidth & """ size=""6"" onBlur=""CheckNumber(this,'图片宽度');"">"
		.echo "                像素 高"
		.echo "                <input name=""PicHeight"" class=""textbox"" type=""text"" id=""PicHeight2"" value=""" & PicHeight & """ size=""6"" onBlur=""CheckNumber(this,'图片高度');"">"
		.echo "                像素</td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"">显示名称"
					  
					If cbool(ShowTitle) = true Then
					.echo ("<input name=""ShowTitle"" id=""ShowTitle"" type=""radio"" value=""true"" checked>显示　")
					.echo ("<input name=""ShowTitle"" id=""ShowTitle"" type=""radio"" value=""false"">不显示")
					Else
					  .echo ("<input type=""radio"" id=""ShowTitle"" value=""true"" name=""ShowTitle"">显示　")
					  .echo ("<input type=""radio"" id=""ShowTitle"" value=""false"" name=""ShowTitle"" checked>不显示")
				   End If
				
		.echo "              </td>"
		 .echo "             <td height=""30"">" &ReturnOpenTypeStr(OpenType) & "</td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">标题字数"
		 .echo "               <input name=""TitleLen"" class=""textbox"" onBlur=""CheckNumber(this,'标题字数');"" type=""text""    style=""width:70%;"" value=""" & TitleLen & """ > "
		 .echo "             </td>"
		 .echo "             <td height=""30""><font color=""#FF0000"">一个汉字=两个英文字符</font></td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">标题样式"
		 .echo "               <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" value=""" & TitleCss & """></td>"
		 .echo "             <td height=""30""><font color=""#FF0000"">已定义的CSS ,要有一定的网页设计基础</font></td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">效果变换间隔时间"
		 .echo "               <input name=""ChangeTime"" class=""textbox"" type=""text"" id=""ChangeTime2"" value=""" & ChangeTime & """  onBlur=""CheckNumber(this,'间隔时间');"">"
		 .echo "             </td>"
		 .echo "             <td height=""30""><font color=""#FF0000"">单位:毫秒</font></td>"
		 .echo "           </tr>"
		.echo "                  </table>"	
		.echo "  </form>"
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
	    End With
		End Sub
End Class
%> 
