<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New GetCirClassList
KSCls.Kesion()
Set KSCls = Nothing

Class GetCirClassList
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
		Dim ChannelID,ClassID, IncludeSubClass, ShowClassName,ClassPrintType, OpenType, DocProperty, Num, RowHeight, TitleLen, OrderStr, ColNumber, ShowPicFlag, NavType, Navi, MoreLinkType, MoreLink, SplitPic, DateRule, DateAlign, TitleCss, DateCss,SpecialID,ShowNewFlag,ShowHotFlag, PrintType,AjaxOut,LabelStyle,IntroLen,LabelStyleW
		Dim PicWidth,PicHeight,PicStyle,PicBorderColor,PicSpacing,IsPic
		Dim ButtonType,PriceType,TypeID,ShowGQType
		Dim ClassCol,ClassCss,MenuBgType,MenuBg
		FolderID = Request("FolderID")
		CurrPath = KS.GetCommonUpFilesDir()
		ChannelID=KS.ChkCLng(Request("ChannelID"))
		
		
		With KS
		'判断是否编辑
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
			 Conn.Close:Set Conn = Nothing
			 Set LabelRS = Nothing
			 .echo ("<Script>alert('参数传递出错!');history.back();</Script>")
			 Exit Sub
		  End If
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			FolderID = LabelRS("FolderID")
			LabelContent = LabelRS("LabelContent")
			LabelFlag = LabelRS("LabelFlag")
			LabelRS.Close:Set LabelRS = Nothing
			LabelStyle         = KS.GetTagLoop(LabelContent)
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetCirClassList", ""),"}" & LabelStyle &"{/Tag}", "")
			LabelStyleW        = Split(LabelStyle,"§")(0)
			LabelStyle         = Split(LabelStyle,"§")(1)
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
			  showclassname      = Node.getAttribute("showclassname")
			  OpenType           = Node.getAttribute("opentype")
			  Num                = Node.getAttribute("num")
			  RowHeight          = Node.getAttribute("rowheight")
			  TitleLen           = Node.getAttribute("titlelen")
			  IntroLen           = Node.getAttribute("introlen")
			  OrderStr           = Node.getAttribute("orderstr")
			  ColNumber          = Node.getAttribute("col")
			  ShowPicFlag        = Node.getAttribute("showpicflag")
			  NavType            = Node.getAttribute("navtype")
			  Navi               = Node.getAttribute("nav")
			  MoreLinkType       = Node.getAttribute("morelinktype")
			  MoreLink           = Node.getAttribute("morelink")
			  SplitPic           = Node.getAttribute("splitpic")
			  DateRule           = Node.getAttribute("daterule")
			  DateAlign          = Node.getAttribute("datealign")
			  TitleCss           = Node.getAttribute("titlecss")
			  DateCss            = Node.getAttribute("datecss")
			  ShowNewFlag        = Node.getAttribute("shownewflag")
			  ShowHotFlag        = Node.getAttribute("showhotflag")
			  PrintType          = Node.getAttribute("printtype")
			  
			  ClassPrintType     = Node.getAttribute("classprinttype")
			  ClassCss           = Node.getAttribute("classcss")
			  ClassCol           = Node.getAttribute("classcol")
			  MenuBgType         = Node.getAttribute("menubgtype")
			  MenuBg             = Node.getAttribute("menubg")
			  AjaxOut            = Node.getAttribute("ajaxout")
			  
			  PicWidth           = Node.getAttribute("picwidth")
			  PicHeight          = Node.getAttribute("picheight")
			  PicStyle           = Node.getAttribute("picstyle")
			  PicBorderColor     = Node.getAttribute("picbordercolor")
			  PicSpacing         = Node.getAttribute("picspacing")
			  
			  ButtonType         = Node.getAttribute("buttontype")
			  PriceType          = Node.getAttribute("pricetype")
			  TypeID             = Node.getAttribute("typeid")
			  ShowGQType         = Node.getAttribute("showgqtype")
              ispic              = Node.getAttribute("ispic")
			End If
            
			Set Node=Nothing
			Set XMLDoc=Nothing
		End If
		
		If ClassCol="" Then ClassCol=2
		
		If PrintType="" Then PrintType=3
		If ClassPrintType="" Then ClassPrintType=2
		If Num = "" Then Num = 10
		If RowHeight = "" Then RowHeight = 20
		If TitleLen = "" Then TitleLen = 30
		If IntroLen = "" Then IntroLen = 50
		If ColNumber = "" Then ColNumber = 1
		If SpecialID=""  Then SpecialID=0
		If ShowNewFlag="" Then ShowNewFlag=False
		If ShowHotFlag="" Then ShowHotFlag=False
		If KS.IsNul(ispic)       Then ispic=False
		If PicWidth="" Then PicWidth=130
		If PicHeight="" Then PicHeight=90
		If PicStyle="" Then PicStyle=1
		If PicSpacing="" Then PicSpacing=2
		If ButtonType="" Then ButtonType=4
		If PriceType="" Then PriceType=0
		If TypeID="" Then TypeID=0
		If ShowGQType="" Or IsNull(ShowGQType) Then ShowGQType=true
		If AjaxOut="" Then AjaxOut=false
		If LabelStyleW="" Then LabelStyleW="<div class=""col"">" & vbcrlf & " <div class=""t""><span><a href=""{@classurl}"" target=""_blank"">更多...</a></span>{@classname}</div>" & vbcrlf & " <ul>{$InnerText}</ul>" & vbcrlf & "</div>"
		If LabelStyle="" Then LabelStyle="[loop={@num}] " & vbcrlf & "<li><a href=""{@linkurl}"" target=""_blank"">{@title}</a></li>" & vbcrlf & "[/loop]"
		.echo "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../ks_inc/Jquery.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../ks_inc/kesion.box.js"" language=""JavaScript""></script>"
		%>
		<style type="text/css">
		 .field{width:720px;}
		 .field li{cursor:pointer;float:left;border:1px solid #DEEFFA;background-color:#F7FBFE;height:18px;line-height:18px;margin:3px 1px 0px;padding:2px}
		 .field li.diyfield{border:1px solid #f9c943;background:#FFFFF6}
		</style>
		<script type="text/javascript">
		var TempFieldStr='';
		var TempDateStr='';
		var TempTitleCss='';
		var GenericPicStyleOption="<option value='1'>①:仅显示缩略图</option><option value='2'>②:缩略图+名称:上下</option><option value='3'>③:缩略图+(名称+简介:上下):左右</option><option value='4'>④:(名称+简介:上下)+缩略图:左右</option>";
						 
		$(document).ready(function(){
		  $("#ChannelID").change(function(){
		    $(parent.document).find('#ajaxmsg').toggle();
			$.get('../../../plus/ajaxs.asp',{action:'GetClassOption',channelid:$(this).val()},function(data){
			  $("#ClassList").empty();
			  $("#ClassList").append("<option value='-1' style='color:red'>-当前栏目(通用)-</option>");
			  $("#ClassList").append("<option value='0'>-不指定栏目-</option>");
			  $("#ClassList").append(unescape(data));
			  SetField($("#ChannelID").val());
			  SetPicStyle($("#ChannelID").val());
			  SetModelParam($("#ChannelID").val());
			  $(parent.document).find('#ajaxmsg').toggle();
			 });
		    });
		   
		  $("#MutileClass").click(function(){
		    if ($(this).attr("checked")==true){
		      $("#ClassList").attr("multiple","multiple");
		      $("#ClassList").attr("style","height:60px");
			  $("#MoreLinkArea").hide();
		    }else{
			   $("#ClassList").removeAttr("multiple");
			   $("#MoreLinkArea").show();
			}
		  });
		  
		  SetPicStyle($("#ChannelID").val()); //填充样式选项
		  $("#PicStyle").change(function(){
		    $("#ViewStylePicArea").html('<img style="border:1px solid #ccc;margin:5px" src="../../Images/View/S'+$(this).val()+'.gif" height="100" width="180" border="0">');
			if ($(this).val()==1){
			 if ($("#ShowPicTitleCss").html()!=null)	TempTitleStr=$("#ShowPicTitleCss").html();
			 $("#ShowPicTitleCss").empty();
			}else{
			$("#ShowPicTitleCss").html(TempTitleStr);
			}
		  });
		  $("#ViewStylePicArea").html('<img style="border:1px outset #ccc;margin:5px" src="../../Images/View/S<%=PicStyle%>.gif" height="100" width="180" border="0">');
		  try{
		  $("#PicStyle>option[value=<%=PicStyle%>]").attr("selected",true);
		  }catch(e){
		  }
		  
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
		  <%
		  end if
		  If LabelID <> "" Then
		   .echo "$('#ChannelID').attr('disabled',true);"
		  End If
		  %>
		  TempFieldStr=$("#ShowFieldArea").html();
		  TempDateStr=$("#ShowTableDate").html();
		  TempTitleStr=$("#ShowTitleCss").html();
		  ChangeOutArea($("#PrintType option:selected").val());
		  ChangeClassPrintOutArea($("#ClassPrintType option:selected").val());
		})
		
		function SetField(channelid)
		{  
		   switch (parseInt(channelid)){
		    case 5:
		     $("#ShowFieldArea").html(TempFieldStr+"<li onclick=\"InsertLabel('{@bigphoto}')\" title=\"商品大图\">商品大图</li><li onclick=\"InsertLabel('{@price_market}')\" title=\"市场价格\">市 场 价</li><li onclick=\"InsertLabel('{@price_member}')\" title=\"商 城 价\">商 城 价</li><li onclick=\"InsertLabel('{@price_vip}')\" title=\"VIP价格\">VIP 价格</li><li onclick=\"InsertLabel('{@score_vip}')\" title=\"赠送积分\">赠送积分</li><li title=\"品牌ID\" onclick=\"InsertLabel('{@brandid}')\">品牌ID号</li><li title=\"品牌名称\" onclick=\"InsertLabel('{@brandname}')\">品牌名称</li><li title=\"品牌英文名\" onclick=\"InsertLabel('{@brandename}')\">品牌英文名</li><br/><li title=\"商品型号\" onclick=\"InsertLabel('{@promodel}')\">商品型号</li><li title=\"商品编号\" onclick=\"InsertLabel('{@proid}')\">商品编号</li><li title=\"商品单位\" onclick=\"InsertLabel('{@unit}')\">商品单位</li>");
			 break;
		    case 7:
		     $("#ShowFieldArea").html(TempFieldStr+"<li onclick=\"InsertLabel('{@movieact}')\" title=\"主要演员\">主要演员</li><li onclick=\"InsertLabel('{@moviedy}')\" title=\"影片导演\">影片导演</li><li title=\"播放时间\" onclick=\"InsertLabel('{@movietime}')\">播放时间</li><li title=\"影片语言\" onclick=\"InsertLabel('{@movieyy}')\">影片语言</li><li title=\"出产地区\" onclick=\"InsertLabel('{@moviedq}')\">出产地区</li><li title=\"所需点数\" onclick=\"InsertLabel('{@readpoint}')\">所需点数</li>");
		     break;
		    case 8:
		     $("#ShowFieldArea").html(TempFieldStr+"<li onclick=\"InsertLabel('{@validdate}')\" title=\"有效期\">有 效 期</li><li onclick=\"InsertLabel('{@typeid}')\" title=\"交易类别\">交易类别</li><li title=\"联系人\" onclick=\"InsertLabel('{@contactman}')\">联 系 人</li><li title=\"公司名称\" onclick=\"InsertLabel('{@companyname}')\">公司名称</li><li title=\"所在省份\" onclick=\"InsertLabel('{@province}')\">所在省份</li><li title=\"所在城市\" onclick=\"InsertLabel('{@city}')\">所在城市</li><li title=\"详细地址\" onclick=\"InsertLabel('{@address}')\">详细地址<li title=\"联系电话\" onclick=\"InsertLabel('{@tel}')\">联系电话</li></li>");
		     break;
			
		   default:
		     $("#ShowFieldArea").html(TempFieldStr);
		   }
		   
		   if ($("#PrintType").val()==4){
		      $(parent.document).find('#ajaxmsg').toggle();
		  	  $.get('../../../plus/ajaxs.asp',{action:'GetFieldOption',channelid:channelid},function(data){
			  $("#ShowFieldArea").html($("#ShowFieldArea").html()+data)
			  $(parent.document).find('#ajaxmsg').toggle();
			 });

		 }
		}
		
		function SetPicStyle(channelid)
		{ 
		   switch (parseInt(channelid))
		   { case 0:
		     case 1:
			 case 2:
			 case 3:
			   $("#PicStyle").empty();
			   $("#PicStyle").append(GenericPicStyleOption);
			  break;
			 case 4:
			   $("#PicStyle").empty();
			   $("#PicStyle").append(GenericPicStyleOption);
			   $("#PicStyle").append("<option value='5'>⑤:缩略图+(名称+类别+作者+时间:上下):左右</option>");
			   $("#PicStyle").append("<option value='6'>⑥:缩略图+(名称+介绍:上下+人气等):左右</option>");
			   break;
			 case 5:
			   $("#PicStyle").empty();
			   $("#PicStyle").append(GenericPicStyleOption);
			   $("#PicStyle").append("<option value='7'>⑤:缩略图+按钮</option>");
			   $("#PicStyle").append("<option value='8'>⑥:缩略图+名称+按钮:上下</option>");
			   $("#PicStyle").append("<option value='9'>⑦:缩略图+名称+价格+按钮:上下</option>");
			   $("#PicStyle").append("<option value='10'>⑧:缩略图+(价格+按钮:上下):左右</option>");
			   $("#PicStyle").append("<option value='11'>⑨:(缩略图+名称)+(价格+按钮):左右</option>");
			   $("#PicStyle").append("<option value='12'>⑩:缩略图+(名称+价格+按钮):左右</option>");
			   break;
			 case 7:
			   $("#PicStyle").empty();
			   $("#PicStyle").append(GenericPicStyleOption);
			   $("#PicStyle").append("<option value='13'>⑤:缩略图+(名称+主演+简介+按钮):左右</option>");
			   $("#PicStyle").append("<option value='14'>⑥:缩略图+(名称+简介+属性):左右</option>");
			   $("#PicStyle").append("<option value='15'>⑦:缩略图+(名称+主演+导演+简介+按钮):左右</option>");
			   break;
			 case 8:
			   $("#PicStyle").empty();
			   $("#PicStyle").append(GenericPicStyleOption);
			   $("#PicStyle").append("<option value='16'>⑤:缩略图+[(标题+地区+时间)+简介]:左右</option>");
			   $("#PicStyle").append("<option value='17'>⑥:缩略图+(名称+简介+属性):左右</option>");
			   break;
			 default:
			  break;
		   }
		}
		function ProductStyle(){
		 return "<tr class='tdbg'><td colspan='2'>按钮样式 <select style='width:160px' name='ButtonType' id='ButtonType'><option value='0'>不显示</option><option value='1'>显示购买按钮</option><option value='2'>显示收藏按钮</option><option value='3'>显示详情按钮</option><option value='4' selected>显示购买+收藏按钮</option><option value='5'>显示购买+详情按钮</option><option value='6'>显示收藏+详情按钮</option><option value='7'>显示购买+详情+收藏按钮</option></select> 价格样式 <select style='width:160px' class='textbox' name='PriceType' id='PriceType'><option value='0' selected>自动显示</option><option value='1'>只显示商城价</option><option value='2'>只显示参考价</option><option value='3'>只显示VIP价</option><option value='4'>参考价+商城价</option><option value='5'>参考价+商城价+VIP价</option><option value='6'>参考价+VIP价</option><option value='7'>商城价+VIP价</option></select></td></tr>";
		}
		function SetModelParam(channelid)
		{
		  if (parseInt(channelid)<=1) 
		    $("#twbz").show() 
		  else $("#twbz").hide();
		  
		  if (parseInt(channelid)==5){
		   if (parseInt($("#PrintType").val())==2)   
		    $("#ModelParamArea").show();
		   else
		    $("#ModelParamArea").hide();
		   $("#ModelParamArea").empty();
		   $("#ModelParamArea").append(ProductStyle());
		   $("#ButtonType>option[value=<%=ButtonType%>]").attr("selected",true);
		   $("#PriceType>option[value=<%=PriceType%>]").attr("selected",true);
		  }
		 else if(parseInt(channelid)==8){
		  $("#ModelParamArea").show();
		  $("#ModelParamArea").empty();
		  
		  $("#ModelParamArea").append("<tr class='tdbg'><td colspan='2'>交易类型 <%= Replace(Replace(KS.ReturnGQType(TypeID,1),"""","\"""),vbcrlf,"\n")%>  <label><input type='checkbox' name='ShowGQType' id='ShowGQType'>显示交易类型</label></td></tr>");
		  $("#TypeID").css("width",120);
		  <%if ChannelID=8 Then%>
		  $("#TypeID>option[value=<%=ButtonType%>]").attr("selected",true);
		  <%if cbool(ShowGQType)=true then .echo "$('#ShowGQType').attr('checked',true);"%>
		  <%End If%>
		 }else{
		   $("#ModelParamArea").hide()
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
      function FieldInsertCode(fieldname,dbtype,dbname)
		{ 
		   if(pos==null) {alert('请先定位插入位置!');return false;}
		   var link="../FieldParam.asp?fieldname=" + fieldname + "&dbtype="+ dbtype + "&dbname=" + dbname+"&datasourcetype=0";
		  var p=new KesionPopup()
		  p.PopupImgDir="../../";
		  p.PopupCenterIframe('插入字段标签',link,350,230,'no');
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
		
		function ChangeOutArea(Val)
		{
		 SetModelParam($("#ChannelID").val());
		 switch (parseInt(Val)){
		  case 2:
		   $("#DiyArea").hide();
		   $("#TableArea").hide();
		   $("#PicArea").show();
		   $("#ShowIntroArea").show();
   		   
		     $("#ShowPicTitleCss").html(TempTitleStr);
		     $("#ShowTitleCss").empty();
		   $("#ViewStylePicArea").html('<img style="border:1px outset #ccc;margin:5px" src="../../Images/View/S'+$("#PicStyle").val()+'.gif" height="100" width="180" border="0">');
		   break;
		  case 3:
		  case 4:
		  $("#DiyArea").show();
		  $("#TableArea").hide();
		  $("#PicArea").hide();
		  $("#ShowDiyDate").html(TempDateStr);
		  $("#ShowTableDate").html('')
		  $("#DateRule").attr("style","width:130px");
		  $("#ShowIntroArea").show();
		  break;
		  default :
		  $("#DiyArea").hide();
		  $("#PicArea").hide();
		  $("#TableArea").show();
		  $("#ShowTableDate").html(TempDateStr);
		  $("#ShowDiyDate").html('')
		  $("#DateRule").attr("style","width:268px");
		  $("#ShowIntroArea").hide();
		  $("#ShowTitleCss").html(TempTitleStr);
		  $("#ShowPicTitleCss").html('');
		  break;
		 }
		 SetField($("#ChannelID").val());
		 
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
			var NavType=1;
			var ShowClassName,ShowPicFlag,ShowNewFlag,ShowHotFlag;
			var OpenType=$("#OpenType").val();
			var Num= $("#Num").val();
			var RowHeight=$("input[name=RowHeight]").val();
			var TitleLen=$("input[name=TitleLen]").val();
			var IntroLen=$("input[name=IntroLen]").val();
			var OrderStr=$("#OrderStr").val();
			var ColNumber=$("input[name=ColNumber]").val();
			var Nav,NavType=$("select[name=NavType]").val();
			var MoreLink,MoreLinkType=$("select[name=MoreLinkType]").val();
			var SplitPic=$("input[name=SplitPic]").val();
			var DateRule= $("#DateRule").val();
			var DateAlign=$("select[name=DateAlign]").val();
			var TitleCss=$("input[name=TitleCss]").val();
			var DateCss=$("input[name=DateCss]").val();
			var PicWidth=$("input[name=PicWidth]").val();
			var PicHeight=$("input[name=PicHeight]").val();
			var PicStyle=$("#PicStyle").val();
			var PicBorderColor=$("input[name=PicBorderColor]").val();
			var PicSpacing=$("input[name=PicSpacing]").val();
			
			var PrintType=$("#PrintType").val();
			var ClassPrintType=$("#ClassPrintType").val();
			var AjaxOut=false;
			if ($("#AjaxOut").attr("checked")==true){AjaxOut=true}
			var IncludeSubClass=false;
			if ($("#IncludeSubClass").attr("checked")==true) IncludeSubClass=true;
			var ShowClassName =false;
			if ($("#ShowClassName").attr("checked")==true) ShowClassName = true
			var ShowPicFlag=false;
			if ($("#ShowPicFlag").attr("checked")==true)  ShowPicFlag= true
			var ShowHotFlag=false;
			if ($("#ShowHotFlag").attr("checked")==true)  ShowHotFlag= true
			var ShowNewFlag=false;
			if ($("#ShowNewFlag").attr("checked")==true)  ShowNewFlag= true
			 var ispic=false;
			if ($("#ispic").attr("checked")==true) ispic=true;
			 
			if  (Num=='')  Num=10;
			if (RowHeight=='') RowHeight=20
			if  (TitleLen=='') TitleLen=30;
			if  (ColNumber=='') ColNumber=1;
			if  (NavType==0) Nav=$("input[name=TxtNavi]").val();
			 else  Nav=$("input[name=NaviPic]").val();
			if  (MoreLinkType==0) MoreLink=$("input[name=MoreLinkWord]").val();
			else  MoreLink=$("input[name=MoreLinkPic]").val();
			
			var ClassCss=$("#ClassCss").val();
			var ClassCol=$("#ClassCol").val();
			var MenuBgType=$("#MenuBgType").val();
			var MenuBg='';
			if  (MenuBgType==0) MenuBg=$("#ColorMenuBg").val()
			 else  MenuBg=$("#PicMenuBg").val();
			
			var tagVal='{Tag:GetCirClassList labelid="0" ispic="'+ispic+'" classprinttype="'+ClassPrintType+'" classcss="'+ClassCss+'" classcol="'+ClassCol+'" menubgtype="'+MenuBgType+'" menubg="'+MenuBg+'" printtype="'+PrintType+'" ajaxout="'+AjaxOut+'" modelid="'+ChannelID+'" classid="'+ClassList+'" includesubclass="'+IncludeSubClass+'" orderstr="'+OrderStr+'" opentype="'+OpenType+'" num="'+Num+'" titlelen="'+TitleLen+'" introlen="'+IntroLen+'" rowheight="'+RowHeight+'" col="'+ColNumber+'" showclassname="'+ShowClassName+'" showpicflag="'+ShowPicFlag+'" shownewflag="'+ShowNewFlag+'" showhotflag="'+ShowHotFlag+'" navtype="'+NavType+'" nav="'+Nav+'" morelinktype="'+MoreLinkType+'" morelink="'+MoreLink+'" splitpic="'+SplitPic+'" daterule="'+DateRule+'" datealign="'+DateAlign+'" titlecss="'+TitleCss+'" datecss="'+DateCss+'" picwidth="'+PicWidth+'" picheight="'+PicHeight+'" picstyle="'+PicStyle+'" picbordercolor="'+PicBorderColor+'" picspacing="'+PicSpacing+'"';
			if (ChannelID==5){
			 var ButtonType=$("#ButtonType").val();
			 var PriceType =$("#PriceType").val();
			 tagVal += ' buttontype="'+ButtonType+'" pricetype="'+PriceType+'"';
			}else if(ChannelID==8){
			 var TypeID=$("#TypeID").val();
			 var ShowGQType=false;
			 if($("#ShowGQType").attr("checked")==true) ShowGQType=true;
			 tagVal += ' typeid="'+TypeID+'" showgqtype="'+ShowGQType+'"';
			}
			tagVal  +='}'+$("#LabelStyleW").val()+'§'+$("#LabelStyle").val()+'{/Tag}';
			
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
		.echo " <input type=""hidden"" name=""LabelContent"">"
		.echo " <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""" & LabelFlag & """>"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo " <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetCirClassList.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td width=""50%"" height=""24"" colspan=""2"">选择范围"
		.echo "                <select name=""ChannelID"" id=""ChannelID"">"
		.echo "                 <option value=""0"">-所有模型-</option>"
        .LoadChannelOption ChannelID
		.echo "                </select>"
		.echo "                <select class=""textbox"" name=""ClassList"" id=""ClassList"" onChange=""SetLabelFlag(this)"">"
		.echo "                 <option "
		if classid="-1" then .echo "selected"
		.echo " value=""-1"" style=""color:red"">- 当前栏目(通用)-</option>"
		.echo "                 <option "
		if classid="-2" then .echo "selected"
		.echo " value=""-2"" style=""color:blue"">- 调用本模型的一级栏目-</option>"

						  .echo Replace(KS.LoadClassOption(ChannelID,false),"value='" & ClassID & "'","value='" & ClassID &"' selected")
						  .echo "</select>"

						  
					If cbool(IncludeSubClass) = True Or LabelID = "" Then
					  Str = " Checked"
					Else
					  Str = ""
					End If
					  .echo "<input type='checkbox' name='MutileClass' id='MutileClass' value='1'>指定多栏目"
					  .echo ("<input name=""IncludeSubClass"" type=""checkbox"" id=""IncludeSubClass"" value=""true""" & Str & ">")
			
		.echo "                  调用子栏目"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">采用Ajax输出</label>"

		.echo "              </div></td>"
		.echo "            </tr>"
	
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
		.echo "            <font color=green>便于更好的控制,建议选择自定义输出样式</font></td></tr>"
		
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
				  .echo ("<input type=""text"" class=""textbox"" name=""ColorMenuBg"" id=""ColorMenuBg"" style=""width:120;""> ")
				  .echo " <img border=0 id=""ColorMenuBgShow1"" src=""../../images/rect.gif"" style=""cursor:pointer;background-Color:" & MenuBg & ";"" onClick=""Getcolor(this,'../../../editor/ksplus/selectcolor.asp?ColorMenuBgShow1|ColorMenuBg');"" title=""选取颜色"">"
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
		.echo "                <tr><td style='text-align:right' width='100'><strong>可用标签:</strong></td>"
		.echo "                <td><li onclick=""InsertLabel('{@autoid}')"">行号</li><li onclick=""InsertLabel('{@classname}')"">栏目名称</li><li onclick=""InsertLabel('{@classurl}')"">栏目URL</li> <li onclick=""InsertLabel('{@classimg}')"">栏目图片地址</li><li onclick=""InsertLabel('{@classintro}')"">栏目介绍(200字)</li><li onclick=""InsertLabel('{@tid}')"">栏目大ID</li><li onclick=""InsertLabel('{@classid}')"">栏目小ID</li></td>"
		.echo "                 </tr></table>"
		.echo "               </td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'>"
		.echo "                <table border='0' width='100%'><tr><td width='100' style='text-align:right'><strong>外循环(栏目)</strong><br><font color=blue>必须包含标签{$InnerText}</font></td>"
		.echo "                <td><textarea name='LabelStyleW' onkeyup='setPos(""LabelStyleW"")' onclick='setPos(""LabelStyleW"")' id='LabelStyleW' style='width:100%;height:120px'>" & LabelStyleW & "</textarea></td>"
		.echo "                </tr>"
		.echo "               </table>"
		.echo "             </td>"
		.echo "            </tr>"
		.echo "           </tbody>"
	
	
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">&nbsp;&nbsp;&nbsp;&nbsp;<strong>文档输出格式</strong>&nbsp;"
		.echo " <select class='textbox'  name=""PrintType"" id=""PrintType"" onChange=""ChangeOutArea(this.options[this.selectedIndex].value);"">"
        .echo "  <option value=""1"""
		If PrintType=1 Then .echo " selected"
		.echo ">文本列表样式(Table)</option>"
        .echo "  <option value=""2"""
		If PrintType=2 Then .echo " selected"
		.echo ">图片列表样式(Table)</option>"
        .echo "  <option value=""3"""
		If PrintType=3 Then .echo " selected"
		.echo ">自定义输出样式(不带自定义字段)</option>"
        .echo "  <option style='color:green' value=""4"""
		If PrintType=4 Then .echo " selected"
		.echo ">自定义输出样式(带自定义字段)</option>"
        .echo "</select>"
		.echo "             </td> <td><span id='ShowDiyDate'></span> <span id='ShowIntroArea'>简介字数<input type='text' class='textbox' style='text-align:center' name='IntroLen' id='IntroLen' value='" & IntroLen & "' size='4'></span><label style='font-weight:bold;color:brown'><input type='checkbox' value='1' name='ispic' id='ispic'"
		if cbool(ispic)=true then .echo " checked"
		.echo "/>仅显示有图片的文档</label></td>"
		.echo "            </tr>"
		.echo "            <tbody id=""DiyArea"">"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'>"
		.echo "               <table border='0' width='100%'><tr><td width='100' style='text-align:right'><strong>可用标签</strong></td>"
		.echo "                <td id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">行 号</li><li onclick=""InsertLabel('{@linkurl}')"">链接URL</li> <li onclick=""InsertLabel('{@id}')"">文档ID</li><li onclick=""InsertLabel('{@title}')"">标 题</li><li onclick=""InsertLabel('{@fulltitle}')"" style='color:red'>不截断标题</li><li onclick=""InsertLabel('{@classname}')"">栏目名称</li><li onclick=""InsertLabel('{@classurl}')"">栏目URL</li> <li onclick=""InsertLabel('{@intro}')"">简要介绍</li><li onclick=""InsertLabel('{@photourl}')"">图片地址</li><li onclick=""InsertLabel('{@adddate}')"">添加时间</li><li onclick=""InsertLabel('{@inputer}')"">录入员</li><li onclick=""InsertLabel('{@hits}')"">点击数</li></td>"
		.echo "               </tr>"
		.echo "              </table>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'>"
		.echo "               <table border='0' width='100%'><tr><td width='100' style='text-align:right'><strong>内循环(文档)</strong></td>"
		.echo "               <td><textarea name='LabelStyle' onkeyup='setPos(""LabelStyle"")' onclick='setPos(""LabelStyle"")' id='LabelStyle' style='width:100%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "               </tr>"
		.echo "              </table>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>使用说明 :</font></strong><br />1、循环标签[loop=n][/loop]对可以省略,也可以平行出现多对；<br />2、输出格式选择不带自定义字段输出的运行效率高于带自定义字段,如果列表没有用到自定义字段请选择不带自定义字段输出</font><br/>3、支持条件判断语句，格式：<font color=blue>{$IF 条件}</font><font color=red>{成立执行的代码}</font><font color=green>{不成立执行的代码}</font><font color=blue>{/$IF}</font></td>"
		.echo "            </tr>"
		.echo "           </tbody>"	
	
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" width=""50%"">排序方法"
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
		.echo "              <td height=""24"">" & ReturnOpenTypeStr(OpenType) & "</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" colspan='2'>文档数量"
		.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num""    style=""width:40px;text-align:center"" onBlur=""CheckNumber(this,'文档数量');"" value=""" & Num & """>条 标题字数<input name=""TitleLen"" class=""textbox"" onBlur=""CheckNumber(this,'标题字数');"" type=""text""    style=""width:40px;;text-align:center"" value=""" & TitleLen & """> 行距"
		.echo "                <input name=""RowHeight"" class=""textbox"" type=""text"" id=""RowHeight2""    style=""width:40px;;text-align:center"" onBlur=""CheckNumber(this,'文档行距');"" value=""" & RowHeight & """>px 列数<input type=""text"" class=""textbox"" onBlur=""CheckNumber(this,'排列列数');""  style=""width:40px;text-align:center"" value=""" & ColNumber & """ name=""ColNumber""> <font color=red>Tips:若自定义样式输出,行距和列数请在您的样式里控制</font></td>"
		.echo "            </tr>"
		
		

		
		
		.echo "           <tbody id='ModelParamArea'></tbody>"

		
		.echo "          <tbody id='TableArea'>"
		.echo "           <tr class=tdbg>"
		 .echo "             <td colspan=2 height=""30"">附加显示 "
				   If cbool(ShowClassName) = true Then
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowClassName"" name=""ShowClassName"" checked>显示栏目</label>")
				   Else
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowClassName"" name=""ShowClassName"">显示栏目</label>")
				   End If
                    .echo "&nbsp;&nbsp;&nbsp;"
					 If cbool(ShowPicFlag) = True Then
					  .echo ("<label id='twbz'><input type=""checkbox"" value=""true"" id=""ShowPicFlag"" name=""ShowPicFlag"" checked>“图文”标志</label>")
					 Else
					  .echo ("<label id='twbz'><input type=""checkbox"" value=""true"" id=""ShowPicFlag"" name=""ShowPicFlag"">“图文”标志</label>")
					 End If
				   .echo "&nbsp;&nbsp;&nbsp;"
					 If  cbool(ShowNewFlag) = True Then
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowNewFlag"" name=""ShowNewFlag"" checked>最新文档标志</label>")
					 Else
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowNewFlag"" name=""ShowNewFlag"">最新文档标志</label>")
					 End If
				 .echo "&nbsp;&nbsp;&nbsp;"
					 If  cbool(ShowHotFlag) = True Then
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowHotFlag"" name=""ShowHotFlag"" checked>显示热门文档标志</label>")
					 Else
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowHotFlag"" name=""ShowHotFlag"">显示热门文档标志</label>")
					 End If
			   
		.echo "       　</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">导航类型"
		.echo "                <select name=""NavType"" style=""width:70%;"" class='textbox' onchange=""SetNavStatus()"">"
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
		 .echo "           <tr  class=tdbg id=""MoreLinkArea"""
		 If Instr(ClassID,",")<>0 Then .echo " style='display:none'"
		 .echo ">"
		 .echo "             <td width=""50%"" height=""24"">更多链接"
		 .echo "               <select name=""MoreLinkType"" style=""width:70%;"" class='textbox' onchange=""SetMoreLinkStatus()"">"
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
					.echo ("  <input type=""text"" class=""textbox"" name=""MoreLinkWord"" style=""width:70%;"" value=""" & MoreLink & """> 支持HTML语法")
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
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" colspan=""2"">分隔图片"
		.echo "                <input name=""SplitPic"" class=""textbox"" type=""text"" id=""SplitPic"" style=""width:61%;"" value=""" & SplitPic & """ readonly>"
		.echo "                <input class='button' name=""SubmitPic"" class='button' onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.SplitPic);"" type=""button"" id=""SubmitPic2"" value=""选择图片..."">"
		.echo "                <span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.SplitPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>"
		.echo "                <div align=""left""> </div></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" id='ShowTableDate'>日期格式"
		.echo ReturnDateFormat(DateRule)
		.echo "                </td>"
		.echo "              <td height=""24"">"
		.echo "                <div align=""left"">日期对齐"
		.echo "                  <select class=""textbox"" name=""DateAlign"" id=""select3"" style=""width:70%;"">"
							
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
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" id=""ShowTitleCss"">标题样式"
		.echo "                <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" value=""" & TitleCss & """></td>"
		.echo "              <td height=""24"">日期样式"
		.echo "                <input name=""DateCss"" class=""textbox"" type=""text"" id=""DateCss"" style=""width:70%;"" value=""" & DateCss & """></td>"
		.echo "            </tr>"
		.echo "              </tbody>"



		.echo "           <tbody id='PicArea'>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"">图片设置 宽"
		.echo "                <input name=""PicWidth"" class=""textbox"" type=""text"" id=""PicWidth"" size='4' value=""" & PicWidth & """>px 高<input name=""PicHeight"" class=""textbox"" type=""text"" id=""PicHeight"" size='4' value=""" & PicHeight & """>px</td>"
		.echo "                <td colspan='2' rowspan='5' id='ViewStylePicArea'>图片显示</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"">显示样式"
		.echo "                <select class='textbox' style='width:230px' name=""PicStyle"" id=""PicStyle"">"
							.echo ("<option value=""1"">①:仅显示缩略图</option>")
							.echo ("<option value=""2"">②:缩略图+名称:上下</option>")
							.echo ("<option value=""3"">③:缩略图+(名称+简介:上下):左右</option>")
							.echo ("<option value=""4"">④:(名称+简介:上下)+缩略图:左右</option>")
						 
		.echo "                </select> <font color=""#FF0000""> =>右边效果预览</font></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"">边框颜色 <input type=""text"" id=""PicBorderColor"" class=""textbox"" name=""PicBorderColor"" style=""width:120;"" value=""" & PicBorderColor & """><img border=0 id=""ColorThumbsBorderShow"" src=""../../images/rect.gif"" style=""cursor:pointer;background-Color:" & PicBorderColor & ";"" onClick=""Getcolor(this,'../../../editor/ksplus/selectcolor.asp?ColorThumbsBorderShow|PicBorderColor');"" title=""选取颜色""> 可留空</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"">图片间距:<input type='text' class='textbox' name='PicSpacing' id='PicSpacing' value='" & PicSpacing & "' size='8' style='text-align:center'> px</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" id=""ShowPicTitleCss""></td>"
		.echo "            </tr>"
		.echo "           </tbody>"

		.echo "         </table>"			 
		.echo "    </form>"
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
		End With
		End Sub
End Class
%> 
