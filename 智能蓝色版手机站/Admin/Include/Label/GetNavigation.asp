<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New GetNavigation
KSCls.Kesion()
Set KSCls = Nothing

Class GetNavigation
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
Dim InstallDir, CurrPath, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript,LabelStyle
Dim TypeFlag, OpenType, NavType, Navi, TitleCss, ColNumber, SplitPic, ChannelID,PrintType,DivID,DivClass,UlID,UlClass,LiID,LiClass
FolderID = Request("FolderID")
CurrPath = KS.GetCommonUpFilesDir()
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
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetNavigation", ""),"}" & LabelStyle&"{/Tag}", "")
	'LabelContent       = Replace(Replace(LabelContent, "{Tag:GetNavigation", ""),"}{/Tag}", "")
	Dim XMLDoc,Node
	Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	 If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
	 Else
			 .echo ("<Script>alert('标签加载出错!');history.back();</Script>")
			 Exit Sub
	  End If
	If  Not Node Is Nothing Then
			ChannelID = Node.getAttribute("channelid")
			NavType = Node.getAttribute("navtype")
			Navi = Node.getAttribute("nav")
			SplitPic = Node.getAttribute("splitpic")
			ColNumber = Node.getAttribute("col")
			OpenType =  Node.getAttribute("opentype")
			TitleCss =  Node.getAttribute("titlecss")
			PrintType=  Node.getAttribute("printtype")
			DivID    =  Node.getAttribute("divid")
			divclass =  Node.getAttribute("divclass")
			ulid     =  Node.getAttribute("ulid")
			ulclass  =  Node.getAttribute("ulclass")
			LIID     =  Node.getAttribute("liid")
			LIClass  =  Node.getAttribute("liclass")
	End If
	Set Node=Nothing
	XMLDoc=Empty
End If
If PrintType="" Then PrintType=2
If Navi = "" Then Navi = " | "
If ColNumber = "" Then ColNumber = 10
If LabelStyle="" Then LabelStyle="[loop={@num}] " & vbcrlf & "<li><a href=""{@classurl}"">{@foldername}</a></li>" & vbcrlf & "[/loop]"
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

<script type="text/javascript">
        $(document).ready(function(){
		  ChangeOutArea($("#PrintType option:selected").val());
		})
		function ChangeOutArea(Val)
		{
		 if (Val==2){
		  $("#TableArea").hide();
		  $("#DivArea").show();
		  $("#TableShow").hide();
		  $("#DiyArea").hide();
		 }else if(Val==3){
		  $("#TableArea").hide();
		  $("#DivArea").hide();
		  $("#TableShow").hide();
		  $("#DiyArea").show();
		 }
		 else{
		  $("#TableArea").show();
		  $("#DivArea").hide();
		  $("#TableShow").show();
		  $("#DiyArea").hide();
		 }
		}
		function SetNavStatus()
		{
		  if ($("select[name=NavType]").val()==0)
		   { $("#NavWord").show();
			 $("#NavPic").hide();
		  }else{
		     $("#NavWord").hide();
		     $("#NavPic").show();
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
			var ChannelID=$("#ChannelID").val();
			var OpenType=$("#OpenType").val();
			var Nav,NavType=$("#NavType").val();
			var SplitPic=$("input[name=SplitPic]").val();
			var ColNumber=$("input[name=ColNumber]").val();
			var TitleCss=$("input[name=TitleCss]").val();
					var PrintType=$("select[name=PrintType]").val();
					var divid=$("input[name=divid]").val();
					var divclass=$("input[name=divclass]").val();
					var ulid=$("input[name=ulid]").val();
					var ulclass=$("input[name=ulclass]").val();
					var liid=$("input[name=liid]").val();
					var liclass=$("input[name=liclass]").val();
			if  (NavType==0) Nav=$("#TxtNavi").val();
			 else  Nav=$("#NaviPic").val();
			var tagVal='{Tag:GetNavigation labelid="0" channelid="'+ChannelID+'" navtype="'+NavType+'" nav="'+Nav+'" splitpic="'+SplitPic+'" col="'+ColNumber+'" opentype="'+OpenType+'" titlecss="'+TitleCss+'" printtype="'+PrintType+'" divid="'+divid+'" divclass="'+divclass+'" ulid="'+ulid+'" ulclass="'+ulclass+'" liid="'+liid+'" liclass="'+liclass+'"}'+$("#LabelStyle").val()+'{/Tag}'
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
		.echo " <input type=""hidden"" name=""LabelContent"" id=""LabelContent"">"
		.echo "   <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""2"">"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo "  <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetNavigation.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"

		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">输出格式"
		.echo " <select class='textbox' style='width:70%' name=""PrintType"" id=""PrintType"" onChange=""ChangeOutArea(this.options[this.selectedIndex].value);if (this.value==3){alert('特别提醒:如果选择范围是空间导航,则设自定义输出样式无效!');}"">"
        .echo "  <option value=""1"""
		If PrintType=1 Then .echo " selected"
		.echo ">普通Table格式</option>"
        .echo "  <option value=""2"""
		If PrintType=2 Then .echo " selected"
		.echo ">DIV+CSS格式</option>"
        .echo "  <option style='color:red' value=""3"""
		If PrintType=3 Then .echo " selected"
		.echo ">自定义输出样式(新增)</option>"
        .echo "</select>"
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""24"">"
		.echo "<div id=""TableArea"""
		If PrintType=2 Then .echo "style=""display:none"""
		.echo "><font color=blue>请选择系统支持的输出格式</font></div><span id=""DivArea"""
		If PrintType<>2 Then .echo "style=""display:none"""
		.echo ">&lt;div id=&quot; <input name=""divid"" type=""text"" value=""" & Divid &""" id=""divid"" size=""6""  style=""border-top-width: 0px;border-right-width: 0px;border-bottom-width: 1px;border-left-width:0px;border-bottom-color: #000000"" class='textbox' title=""DIV调用的ID号，必须在CSS中预先定义且不能为空!"">&quot; class=&quot; <input name=""divclass"" class='textbox' type=""text"" value=""" & Divclass &""" id=""divclass"" size=""6"" style=""border-top-width: 0px;	border-right-width: 0px;border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"" title=""DIV调用的Class名称，请在CSS中预先定义,可以为空!""> &quot;&gt;<span style=""color:blue"">此处留空则不输出标记</span><br> &lt;ul  id=&quot; <input value=""" & ulid &""" name=""ulid"" type=""text"" id=""ulid"" class='textbox' size=""6"" style=""border-top-width: 0px;border-right-width: 0px;border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000""  title=""生成ul调用的ID，请在CSS中预先定义,可以为空!""> &quot; class=&quot; <input class='textbox' value=""" & ulclass &""" name=""ulclass""  type=""text"" id=""ulclass"" size=""6"" style=""border-top-width: 0px;border-right-width: 0px;border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"" title=""生成ul调用的class名称，请在CSS中预先定义。可以为空!"">&quot;&gt;<span style=""color:blue"">此处留空则不输出标记</span><br>&lt;li id=&quot; <input value=""" & liid &""" name=""liid"" type=""text"" id=""liid"" size=""6"" class='textbox' style=""border-top-width: 0px;	border-right-width: 0px;border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"" title=""生成li调用的ID，请在CSS中预先定义。可以为空!"">&quot; class=&quot; <input value=""" & liclass &""" name=""liclass"" class='textbox' type=""text"" id=""liclass"" size=""6"" style=""border-top-width: 0px;border-right-width: 0px;border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000""  title=""生成li调用的class名称，请在CSS中预先定义。可以为空!""> &quot;&gt;</div></td>"
		.echo "            </tr>"
		
		.echo "<tbody id=""DiyArea"">"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">行 号</li><li onclick=""InsertLabel('{@classurl}')"">栏目URL</li> <li onclick=""InsertLabel('{@id}')"">栏目ID</li><li onclick=""InsertLabel('{@classid}')"" title='栏目的ClassID'>栏目小ID</li><li onclick=""InsertLabel('{@foldername}')"">栏目名称</li><li onclick=""InsertLabel('{@classename}')"">栏目英文名称</li><li onclick=""InsertLabel('{@classimg}')"">栏目图片地址</li><li onclick=""InsertLabel('{@classintro}')"">栏目介绍</li></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'><textarea name='LabelStyle' onkeyup='setPos()' onclick='setPos()' id='LabelStyle' style='width:95%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>使用说明 :</font></strong><br />循环标签[loop=n][/loop]对可以省略,也可以平行出现多对；</font></td>"
		.echo "            </tr>"
		.echo "           </tbody>"
		
		
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"" colspan=""2"">选择范围"
.echo " " & ReturnAllChannel(ChannelID)    
.echo "             <font color=red>若选择当前频道通用，则相应的频道将调用各自的子栏目</font></td>"
.echo "            </tr>"

.echo "      <tbody id=""TableShow"">"
.echo "            <tr class='tdbg'>"
.echo "              <td width=""50%"" height=""30"">导航类型"
.echo "                <select class=""textbox"" name=""NavType"" id=""NavType"" style=""width:70%;"" onchange=""SetNavStatus()"">"
            
            If LabelID = "" Or CStr(NavType) = "0" Then
            .echo ("<option value=""0"" selected>文字导航</option>")
            .echo ("<option value=""1"">图片导航</option>")
           Else
            .echo ("<option value=""0"">文字导航</option>")
            .echo ("<option value=""1"" selected>图片导航</option>")
           End If
           
.echo "                </select> </td>"
.echo "              <td>"
        
        If LabelID = "" Or CStr(NavType) = "0" Then
          .echo ("<div align=""left"" id=""NavWord""> ")
          .echo ("<input type=""text"" class=""textbox"" id=""TxtNavi"" name=""TxtNavi"" style=""width:70%;"" value=""" & Navi & """>")
          .echo ("</div>")
          .echo ("<div align=""left"" id=NavPic style=""display:none""> ")
          .echo ("<input type=""text"" class=""textbox"" readonly style=""width:55%;"" id=""NaviPic"" name=""NaviPic"">")
          .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""选择图片"">")
          .echo ("</div>")
        Else
          .echo ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
          .echo ("<input type=""text"" class=""textbox"" id=""TxtNavi"" name=""TxtNavi"" style=""width:70%;"">")
          .echo ("</div>")
          .echo ("<div align=""left"" id=NavPic> ")
          .echo ("<input type=""text"" class=""textbox"" readonly style=""width:55%;"" id=""NaviPic"" name=""NaviPic"" value=""" & Navi & """>")
          .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""选择图片"">")
          .echo ("</div>")
        End If
        
.echo "              </td>"
.echo "            </tr>"
.echo "            <tr class='tdbg'>"
.echo "             <td height=""30"" colspan=""2"">分隔图片"
.echo "                <input name=""SplitPic"" class=""textbox"" type=""text"" id=""SplitPic"" style=""width:61%;"" value=""" & SplitPic & """ readonly>"
.echo "                <input  class='button' name=""SubmitPic"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.SplitPic);"" type=""button"" id=""SubmitPic2"" value=""选择图片..."">"
.echo "                <span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.SplitPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>"
.echo "              </td>"
.echo "            </tr>"
.echo "          </tbody>"
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"">显示列数"
.echo "                <input name=""ColNumber"" class=""textbox"" type=""text"" id=""ColNumber"" style=""width:70%"" value=""" & ColNumber & """></td>"
.echo "              <td height=""30"">"
          
 .echo ReturnOpenTypeStr(OpenType)
    
.echo "              </td>"
.echo "            </tr>"
.echo "           <tr class='tdbg'>"
.echo "              <td height=""30"">标题样式"
.echo "                <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" value=""" & TitleCss & """></td>"
.echo "              <td height=""30""><font color=""#FF0000"">已定义的CSS ,要有一定的网页设计基础</font></td>"
.echo "            </tr>"
.echo "                  </table>"	
.echo "  </form>"
  
.echo "</div>"
.echo "</body>"
.echo "</html>"
End With

End Sub

'取得网站的所有频道及其子栏目
Function ReturnAllChannel(FolderID)
  Dim ChannelStr:ChannelStr = ""
      ChannelStr = "<select class='textbox' name=""ChannelID"" id=""ChannelID"" style=""width:200;border-style: solid; border-width: 1"">"
      ChannelStr = ChannelStr & "<option value=""0"">    -整站导航-  </option>"
	  if FolderID="9999" then
	  ChannelStr = ChannelStr & "<option value=""9999"" style=""color:red"" selected>-当前频道通用-</option>"
	  else
	  ChannelStr = ChannelStr & "<option value=""9999"" style=""color:red"">-当前频道通用-</option>"
	  end if
	 if FolderID="9998" then
	   ChannelStr = ChannelStr & "<option value=""9998"" style=""color:blue"" selected>-同级频道通用-</option>"
	   else
	   ChannelStr = ChannelStr & "<option value=""9998"" style=""color:blue"">-同级频道通用-</option>"
	   end if

		ChannelStr = ChannelStr & "<optgroup  label=""-----个人空间相关导航-----"">"
	  if FolderID="9997" then
	  ChannelStr = ChannelStr & "<option value=""9997"" selected>-空间分类-</option>"
	  else
	  ChannelStr = ChannelStr & "<option value=""9997"">-空间分类-</option>"
	  end if
	  if FolderID="9996" then
	  ChannelStr = ChannelStr & "<option value=""9996"" selected>-日志分类-</option>"
	  else
	  ChannelStr = ChannelStr & "<option value=""9996"">-日志分类-</option>"
	  end if
	  if FolderID="9995" then
	  ChannelStr = ChannelStr & "<option value=""9995"" selected>-圈子分类-</option>"
	  else
	  ChannelStr = ChannelStr & "<option value=""9995"">-圈子分类-</option>"
	  end if
	  if FolderID="9994" then
	  ChannelStr = ChannelStr & "<option value=""9994"" selected>-相册分类-</option>"
	  else
	  ChannelStr = ChannelStr & "<option value=""9994"">-相册分类-</option>"
	  end if

		ChannelStr = ChannelStr & "<optgroup  label=""-----指定到模型-----"">"
		ChannelStr = ChannelStr & ReturnChannelOption(FolderID)
   ChannelStr = ChannelStr & "</Select>"
   ReturnAllChannel = ChannelStr
End Function
	'**************************************************
	'函数名：ReturnChannelOption
	'作  用：显示频道列表。
	'参  数：SelectChannelID ----选择频道ID号
	'返回值：频道列表
	'**************************************************
	Public Function ReturnChannelOption(SelectChannelID)
	  Dim RS:Set RS=Server.CreateObject("ADODB.Recordset")
	  Dim SQL,K,ChannelStr:ChannelStr = ""
	   RS.Open "Select channelid,channelname From [KS_Channel] Where ChannelStatus=1 And ChannelID<>10 and channelid<>9", Conn, 1, 1
	   If RS.EOF And RS.BOF Then
		  RS.Close:Set RS = Nothing:Exit Function
	   Else
	     SQL=RS.GetRows(-1):rs.close:set rs=nothing
	   End iF
		
	    For K=0 To ubound(sql,2)
		  If Cstr(sql(0,k)) = Cstr(SelectChannelID) Then
		  ChannelStr = ChannelStr & "<option selected value=" & sql(0,k) & ">" & sql(1,k) & "</option>"
		 Else
		   ChannelStr = ChannelStr & "<option value=" & sql(0,k) & ">" & sql(1,k) & "</option>"
		 End If
		Next 
		ChannelStr = ChannelStr & "<optgroup  label=""-----指定到具体的栏目(以下列出了整站的导航树)----"">"  
	   For K=0 To Ubound(sql,2)
	        ChannelStr=ChannelStr & Replace(KS.LoadClassOption(sql(0,k),false),"value='" & SelectChannelID & "'","value='" & SelectChannelID &"' selected")
	    Next
	   ReturnChannelOption = ChannelStr
	End Function

End Class
%> 
