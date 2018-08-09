<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New GetAQList
KSCls.Kesion()
Set KSCls = Nothing

Class GetAQList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim InstallDir, CurrPath, FolderID, LabelContent, SplitPic, Action, LabelID, Str, Descript, LabelFlag,ShowUserFace,ShowReward,RewardTF,ZeroTF,JJTF
		Dim ClassID, OpenType, RecommendTF, ShowClass,ShowUserName,Num, ZWLen, TitleLen, InfoSort, ColNumber, Province, NavType, Navi, DateRule, DateAlign, TitleCss, City,ShowStyle, PrintType,AjaxOut,LabelStyle
		FolderID = Request("FolderID")
		CurrPath = KS.GetCommonUpFilesDir()
		

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
			 .echo ("<Script>alert('参数传递出错!');window.close();</Script>")
			 Exit Sub
		  End If
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			FolderID = LabelRS("FolderID")
			Descript = LabelRS("Description")
			LabelContent = LabelRS("LabelContent")
			LabelFlag = LabelRS("LabelFlag")
			LabelRS.Close:Set LabelRS = Nothing
            LabelStyle         = KS.GetTagLoop(LabelContent)
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetAQList", ""),"}" & LabelStyle & "{/Tag}", "")
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('标签加载出错!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
			ClassID = Node.getAttribute("classid")
			ShowStyle=Node.getAttribute("showstyle")
			Province=Node.getAttribute("province")
			City=Node.getAttribute("city")
			RecommendTF=Cbool(Node.getAttribute("recommendtf"))
			OpenType = Node.getAttribute("opentype")
			Num = Node.getAttribute("num")
			TitleLen = Node.getAttribute("titlelen")
			InfoSort = Node.getAttribute("infosort")
			ShowClass= Node.getAttribute("showclass")
			ShowUserName= Node.getAttribute("showusername")
			ShowUserface=Node.getAttribute("showuserface")
			ShowReward=Node.getAttribute("showreward")
			SplitPic= Node.getAttribute("splitpic")
			NavType = Node.getAttribute("navtype")
			Navi = Node.getAttribute("nav")
			DateRule = Node.getAttribute("daterule")
			TitleCss = Node.getAttribute("titlecss")
			PrintType= Node.getAttribute("printtype")
			RewardTF = Node.getAttribute("rewardtf")
			ZeroTF   = Node.getAttribute("zerotf")
			JJTF     = Node.getAttribute("jjtf")
			AjaxOut  = Node.getAttribute("ajaxout")
		   End If
		   Set Node=Nothing
		   XMLDoc=Empty
		End If
		If PrintType="" Then PrintType=1
		If Num = "" Then Num = 20
		If TitleLen = "" Then TitleLen = 30
		If ColNumber = "" Then ColNumber = 1
		If RecommendTF="" Then RecommendTF=False
		If ShowStyle="" Then ShowStyle=2
		If KS.IsNul(ShowClass) Then ShowClass=False
		If KS.IsNul(RewardTF) Then RewardTF=false
		If KS.IsNul(ShowUserName) Then ShowUserName=False
		If KS.IsNul(ShowUserFace) Then ShowUserFace=False
		If KS.IsNul(ShowReward) Then ShowReward=false
		If KS.IsNul(ZeroTF) Then ZeroTF=false
		If KS.IsNul(JJTF) Then JJTF=false
		If AjaxOut="" Then AjaxOUT=false
		If LabelStyle="" Then LabelStyle="<li><a href=""{@aqurl}"">{@title}</a></li>"
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
		<script>
		$(document).ready(function(){
		 $("#classid").attr("size",1);
		 $("#classid>option[value=<%=ClassID%>]").attr("selected",true);
		 $("#smallclassid").css("display","none");
		 try{
		  $("#smallerclassid").attr("size",1);
		  $("#smallerclassid").css("width",80);
		 }catch(e){
		 }
		 ChangeOutArea($("#PrintType").val());
		});
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
		 if (Val==2){
		  $("#DiyArea").show();
		 }
		 else{
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
		

		function CheckForm()
		{
		    if ($("input[name=LabelName]").val()=='')
			 {
			  alert('请输入标签名称');
			  $("input[name=LabelName]").focus(); 
			  return false
			  }
			var ClassID=$('#classid').val();
			var ShowStyle=$('#ShowStyle').val();
			var NavType=1;
			var OpenType=$('#OpenType').val();
			var Num=$('#Num').val();
			var TitleLen=$('input[name=TitleLen]').val();
			var InfoSort=$('select[name=InfoSort]').val();
			var SplitPic=$("#SplitPic").val();
			var Nav,NavType=$('select[name=NavType]').val();
			var DateRule=$('input[name=DateRule]').val();
			var TitleCss=$('input[name=TitleCss]').val();
			var PrintType=$('#PrintType').val();
            var AjaxOut=false;
			if ($("#AjaxOut").attr("checked")==true){AjaxOut=true}
			var RecommendTF=false;
			if ($("#RecommendTF").attr("checked")==true){RecommendTF=true}
			var RewardTF=false;
			if ($("#RewardTF").attr("checked")==true){RewardTF=true}
			var ZeroTF=false;
			if ($("#ZeroTF").attr("checked")==true){ZeroTF=true}
			var JJTF=false;
			if ($("#JJTF").attr("checked")==true){JJTF=true}
            var ShowClass=false;
			if ($("#ShowClass").attr("checked")==true){ShowClass=true}
            var ShowUserName=false;
			if ($("#ShowUserName").attr("checked")==true){ShowUserName=true}
            var ShowUserFace=false;
			if ($("#ShowUserFace").attr("checked")==true){ShowUserFace=true}
            var ShowReward=false;
			if ($("#ShowReward").attr("checked")==true){ShowReward=true}
	
			if  (Num=='')  Num=10;
			if  (TitleLen=='') TitleLen=30;
			if  (NavType==0) Nav=$('#TxtNavi').val()
			 else  Nav=$('#NaviPic').val();
            var tagVal='{Tag:GetAQList labelid="0" classid="'+ClassID+'" showstyle="'+ShowStyle+'" recommendtf="'+RecommendTF+'" rewardtf="'+RewardTF+'" opentype="'+OpenType+'" num="'+Num+'" titlelen="'+TitleLen+'" infosort="'+InfoSort+'" showclass="'+ShowClass+'" zerotf="'+ZeroTF+'" jjtf="'+JJTF+'" showusername="'+ShowUserName+'" showuserface="'+ShowUserFace+'" showreward="'+ShowReward+'" splitpic="'+SplitPic+'" navtype="'+NavType+'" nav="'+Nav+'" titlecss="'+TitleCss+'" daterule="'+DateRule+'" printtype="'+PrintType+'" ajaxout="'+AjaxOut+'"}'+$("#LabelStyle").val()+'{/Tag}';
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
		.echo " <input type=""hidden"" name=""LabelFlag"" value=""" & LabelFlag & """>"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo " <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetAQList.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">输出格式"
		.echo " <select class='textbox' style='width:70%' name=""PrintType"" id=""PrintType"" onChange=""ChangeOutArea(this.value);"">"
        .echo "  <option value=""1"""
		If PrintType=1 Then .echo " selected"
		.echo ">普通输出(Table)</option>"
        .echo " <option value=""2"""
		If PrintType=2 Then .echo " selected"
		.echo ">自定义输出样式</option>"
        
        .echo "</select>"
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""24"">"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">采用Ajax输出</label>"
		.echo"</td>"
		.echo "            </tr>"
		
		.echo "            <tbody id=""DiyArea"">"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">行 号</li><li onclick=""InsertLabel('{@aqurl}')"">问题URL</li> <li onclick=""InsertLabel('{@title}')"">标题</li><li onclick=""InsertLabel('{@username}')"">提问者</li><li onclick=""InsertLabel('{@userface}')"">提问者头像</li><li onclick=""InsertLabel('{@pclassname}')"">上级分类名称</li><li onclick=""InsertLabel('{@aqclassname}')"">分类名称</li><li onclick=""InsertLabel('{@aqclassurl}')"">分类Url</li><li onclick=""InsertLabel('{@aqclassid}')"">分类id</li><li onclick=""InsertLabel('{@adddate}')"">提问时间</li><li onclick=""InsertLabel('{@lastposttime}')"">最后回复时间</li><li onclick=""InsertLabel('{@bestusername}')"">最佳答复人</li><li onclick=""InsertLabel('{@lastposttime}')"">最后回复时间</li><li onclick=""InsertLabel('{@topicmode}')"">状态值</li> <li onclick=""InsertLabel('{@expiredtime}')"">过期时间</li><li onclick=""InsertLabel('{@hits}')"">浏览数</li><li onclick=""InsertLabel('{@reward}')"">悬赏分</a><li onclick=""InsertLabel('{@rewardbyimg}')"" style='color:red;width:80px' title='有悬赏分才显示'>悬赏分带图片</a></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'><textarea name='LabelStyle' onkeyup='setPos()' onclick='setPos()' id='LabelStyle' style='width:95%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>使用说明 :</font></strong><br />循环标签[loop=n][/loop]对可以省略,也可以平行出现多对；</td>"
		.echo "            </tr>"
		.echo "           </tbody>"		
		.echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">指定分类"
		.echo " <select id=""classid"" name=""classid"">"
		.echo "<option value='' selected>-不指定分类-</option>"
	dim rs:Set Rs=Server.CreateObject("ADODB.Recordset")
	Rs.Open "SELECT * FROM KS_AskClass ORDER BY rootid,orders", Conn, 1, 1
	dim i:i = 0
	Do While Not Rs.EOF
		
	    .echo "<option value=" & rs("classid") &">"
		If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;├"
		If Rs("depth") > 1 Then
			For i = 2 To Rs("depth")
				Response.Write "&nbsp;&nbsp;│"
			Next
			Response.Write "&nbsp;&nbsp;├ "
		End If
		.echo Rs("ClassName")
		.echo "</option>"
		Rs.movenext
		i = i + 1
	Loop
	Rs.Close
	Set Rs = Nothing
		.echo "</select>"
		
		
		
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""24"">"
		.echo "                <input name=""RecommendTF"" id=""RecommendTF"" type=""checkbox"" value=""true"""
		If Cbool(RecommendTF) = true Then .echo (" Checked")
		.echo ">仅显示推荐"				  
		.echo "                <input name=""RewardTF"" id=""RewardTF"" type=""checkbox"" value=""true"""
		If Cbool(RewardTF) = true Then .echo (" Checked")
		.echo ">仅显示有悬赏"				  
		.echo "                <input name=""ZeroTF"" id=""ZeroTF"" type=""checkbox"" value=""true"""
		If Cbool(ZeroTF) = true Then .echo (" Checked")
		.echo ">仅显示零回答"				  
		.echo "                <input name=""JJTF"" id=""JJTF"" type=""checkbox"" value=""true"""
		If Cbool(JJTF) = true Then .echo (" Checked")
		.echo ">仅显示已解决"				  
		.echo "                </td>"
		.echo "            </tr>"
		
		.echo "            <tr id=""ClassArea"" class=tdbg style=""display:none"">"
		.echo "              <td colspan='2' height=""24"">显示样式"
		.echo "                <select class='textbox' name=""ShowStyle"" id=""ShowStyle"" style=""width:200px;"">"
		Dim StyleStr
		           If ShowStyle = "1" Then StyleStr = ("<option value=""1"" selected>①:样式一</option>") Else	StyleStr = StyleStr & ("<option value=""1"">①:样式一</option>")
				   If ShowStyle = "2" Then StyleStr = StyleStr & ("<option value=""2"" selected>②:样式二</option>") Else StyleStr = StyleStr & ("<option value=""2"">②:样式二</option>")
				  
		
		
		.echo  StyleStr
		.echo "                  </select></td>"
		.echo "            </tr>"
		
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" width=""50%"">排序方法"
		.echo "                <select style=""width:70%;"" class='textbox' name=""InfoSort"">"
					If InfoSort = "a.TopicID Desc" Then
					 .echo ("<option value='a.TopicID Desc' selected>问题ID降序</option>")
					Else
					 .echo ("<option value='a.TopicID Desc'>问题ID降序</option>")
					End If
					
					If InfoSort = "a.DateAndTime Desc" Then
					 .echo ("<option value='a.DateAndTime Desc' selected>最新提问</option>")
					Else
					 .echo ("<option value='a.DateAndTime Desc'>最新提问</option>")
					End If
					If InfoSort = "LastPostTime Desc" Then
					 .echo ("<option value='a.LastPostTime Desc' selected>最新回答</option>")
					Else
					 .echo ("<option value='a.LastPostTime Desc'>最新回答</option>")
					End If
					If InfoSort = "a.Hits Desc,a.topicid desc" Then
					  .echo ("<option value='a.Hits Desc,a.topicid desc' selected>点击数最多</option>")
					Else
					  .echo ("<option value='a.Hits Desc,a.topicid desc'>点击数最多</option>")
					End If
					If InfoSort = "a.reward Desc,a.topicid desc" Then
					  .echo ("<option value='a.reward Desc,a.topicid desc' selected>悬赏分最高</option>")
					Else
					  .echo ("<option value='a.reward Desc,a.topicid desc'>悬赏分最高</option>")
					End If

		.echo "         </select></td>"
		.echo "              <td height=""24"">" & ReturnOpenTypeStr(OpenType) & "</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">问题数量"
		.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num""    style=""width:70%;"" onBlur=""CheckNumber(this,'问题数量');"" value=""" & Num & """></td>"
		.echo "              <td width=""50%"" height=""24"">标题长度"
		.echo "                <input name=""TitleLen"" class=""textbox"" onBlur=""CheckNumber(this,'标题长度');"" type=""text""    style=""width:50px;"" value=""" & TitleLen & """><font color=blue>一个汉字算两个字符</font>"
		.echo "              </td>"
		 .echo "           </tr>"
		
		.echo "           <tr class=tdbg>"
		 .echo "             <td colspan=2 height=""30"">附加显示 "
				   If cbool(ShowClass) = True Then
					  .echo ("<input type=""checkbox"" value=""true"" id=""ShowClass"" name=""ShowClass"" checked>显示分类名称")
				   Else
					  .echo ("<input type=""checkbox"" value=""true"" id=""ShowClass"" name=""ShowClass"">显示分类名称")
				   End If
                      .echo "&nbsp;&nbsp;&nbsp;&nbsp;"
					 If  cbool(ShowUserName) = True Then
					  .echo ("<input type=""checkbox"" value=""true"" id=""ShowUserName"" name=""ShowUserName"" checked>显示提问者")
					 Else
					  .echo ("<input type=""checkbox"" value=""true"" id=""ShowUserName"" name=""ShowUserName"">显示提问者")
					 End If
                      .echo "&nbsp;&nbsp;&nbsp;&nbsp;"
					 If  cbool(ShowUserFace) = True Then
					  .echo ("<input type=""checkbox"" value=""true"" id=""ShowUserFace"" name=""ShowUserFace"" checked>显示提问者头像")
					 Else
					  .echo ("<input type=""checkbox"" value=""true"" id=""ShowUserFace"" name=""ShowUserFace"">显示提问者头像")
					 End If
                      .echo "&nbsp;&nbsp;&nbsp;&nbsp;"
					 If  cbool(ShowReward) = True Then
					  .echo ("<input type=""checkbox"" value=""true"" id=""ShowReward"" name=""ShowReward"" checked>显示悬赏分")
					 Else
					  .echo ("<input type=""checkbox"" value=""true"" id=""ShowReward"" name=""ShowReward"">显示悬赏分")
					 End If
				 
		.echo "       　</td>"
		.echo "            </tr>"
		
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" colspan=""2"">分隔图片"
		.echo "                <input name=""SplitPic"" class=""textbox"" type=""text"" id=""SplitPic"" style=""width:61%;"" value=""" & SplitPic & """ readonly>"
		.echo "                <input class='button' name=""SubmitPic"" class='button' onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.SplitPic);"" type=""button"" id=""SubmitPic2"" value=""选择图片..."">"
		.echo "                <span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.SplitPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>"
		.echo "                <div align=""left""> </div></td>"
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
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" id=""TxtNavi"" style=""width:70%;"" value=""" & Navi & """> 支持HTML语法")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"">")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,$('#NaviPic')[0]);"" name=""Submit3"" value=""选择图片..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:$('#NaviPic').val('');"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  .echo ("</div>")
				Else
				  .echo ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" id=""TxtNavi"" style=""width:70%;""> 支持HTML语法")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"" value=""" & Navi & """>")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,$('#NaviPic')[0]);"" name=""Submit3"" value=""选择图片..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:$('#NaviPic').val('');"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">清除</span>")
				  .echo ("</div>")
				End If
		 .echo "             </td>"
		 .echo "           </tr>"
		

		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"">日期格式"
		.echo ReturnDateFormat(DateRule)
		.echo "                </td>"
		.echo "              <td height=""24"">"
		.echo "                <div align=""left"">标题样式<input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" value=""" & TitleCss & """>"
		.echo "                </div></td>"
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