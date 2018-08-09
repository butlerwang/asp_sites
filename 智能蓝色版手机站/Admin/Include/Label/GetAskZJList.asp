<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New GetAskZJList
KSCls.Kesion()
Set KSCls = Nothing

Class GetAskZJList
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
Dim InstallDir, CurrPath, FolderID, LabelContent, Action, LabelID, Str, Descript,Recommend,LabelFlag,classtype
Dim TypeFlag, Num, IntroLen,ChannelID,PrintType,AjaxOut,LabelStyle,ClassID,OrderStr,BigClassID,SmallClassID,SmallerClassID,DateRule
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
  LabelRS.Open "Select top 1 * From KS_Label Where ID='" & LabelID & "'", Conn, 1, 1
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
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetAskZJList", ""),"}" & LabelStyle&"{/Tag}", "")
			' response.write LabelContent
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('标签加载出错!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
			    classtype        = Node.getAttribute("classtype")
			    BigClassID       = Node.getAttribute("bigclassid")
				SmallClassID     = Node.getAttribute("smallclassid")
				SmallerClassID   = Node.getAttribute("smallerclassid")
				DateRule         = Node.getAttribute("daterule")
				Num              = Node.getAttribute("num")
				IntroLen         = Node.getAttribute("introlen")
				AjaxOut          = Node.getAttribute("ajaxout")
				OrderStr         = Node.getAttribute("orderstr")
				recommend        = Node.getAttribute("recommend")
			End If
			XMLDoc=Empty
			Set Node=Nothing
    
End If
		If IntroLen="" Then IntroLen=50
		If Num = "" Then Num = 10
		If recommend="" Then recommend=0
		if classtype="" then classtype=1
		If BigClassID="" Then BigClassID=0
		If SmallClassID="" Then SmallClassID=0
		If SmallerClassID="" Then SmallerClassID=0
		If LabelStyle="" Then LabelStyle="[loop={@num}] " & vbcrlf & "<li>问答专家：<a href=""{@spaceurl}"" target=""_blank"">{@username}</a> <a href=""{@askzjurl}"" target=""_blank"">咨询</a></li>" & vbcrlf & "[/loop]"
		
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
	

	function CheckForm()
		{
		    if ($("input[name=LabelName]").val()=='')
			 {
			  alert('请输入标签名称');
			  $("input[name=LabelName]").focus(); 
			  return false
			  }
	var ClassID=document.myform.classid.value;
	var SmallClassID=document.myform.smallclassid.value;
	var SmallerClassID=document.myform.smallerclassid.value;
	var Num=document.myform.Num.value;
	var IntroLen=document.myform.IntroLen.value;
	var DateRule=document.myform.DateRule.value;
	var OrderStr=$("#OrderStr").val();
	var AjaxOut=false;
	if ($("#AjaxOut").attr("checked")==true){AjaxOut=true}
			
	if (Num=='') Num=10
	var recommend=$("input[name=recommend]:checked").val();
	var classtype=$("input[name=classtype]:checked").val();
	
	var tagVal='{Tag:GetAskZJList labelid="0" ajaxout="'+AjaxOut+'" classtype="'+classtype+'" recommend="'+recommend+'" bigclassid="'+ClassID+'" smallclassid="'+SmallClassID+'" smallerclassid="'+SmallerClassID+'" num="'+Num+'" orderstr="'+OrderStr+'" daterule="'+DateRule+'" introlen="'+IntroLen+'"}'+$("#LabelStyle").val()+'{/Tag}';
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
		.echo "   <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""" & LabelFlag & """>"
		.echo " <input type=""hidden"" name=""Action"" id=""Action"" value=""" & Action & """>"
		.echo "  <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetAskZJList.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"

		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">输出格式"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">采用Ajax输出</label></td><td>日期格式：" & ReturnDateFormat(DateRule) & "</td>"
		.echo "            </tr>"
		
		
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"">显示条数"
.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num"" style=""width:50px"" value=""" & Num & """>个 简介字数<input name=""IntroLen"" class=""textbox"" type=""text"" id=""IntroLen"" style=""width:50px"" value=""" & IntroLen & """><font color=red>如果不想控制，请设置为“0”</font></td>"

.echo "              <td height=""30"">排序方式"
.echo "                <select style=""width:70%;"" class='textbox' name=""OrderStr"" id=""OrderStr"">"
					If OrderStr = "ID Desc" Then
					.echo ("<option value='ID Desc' selected>专家ID(降序)</option>")
					Else
					.echo ("<option value='ID Desc'>专家ID(降序)</option>")
					End If
					If OrderStr = "ID Asc" Then
					.echo ("<option value='ID Asc' selected>专家ID(升序)</option>")
					Else
					.echo ("<option value='ID Asc'>专家ID(升序)</option>")
					End If
					If OrderStr = "AskDoneNum Asc" Then
					 .echo ("<option value='AskDoneNum Asc' selected>回答数(升序)</option>")
					Else
					 .echo ("<option value='AskDoneNum Asc'>回答数(升序)</option>")
					End If
					If OrderStr = "AskDoneNum Desc" Then
					  .echo ("<option value='AskDoneNum Desc' selected>回答数(降序)</option>")
					Else
					  .echo ("<option value='AskDoneNum Desc'>回答数(降序)</option>")
					End If
	

		.echo "         </select></td>"
.echo "            </tr>"	



		.echo "            <tr class='tdbg' id='spaceclass'>"
		.echo "              <td height=""30"">所属分类"
		.echo "                <input type='radio' name='classtype' onclick=""$('#sssss').hide();"""
		if classtype="1" then .echo " checked"
		.echo " value='1'><font color=blue>当前分类通用</font>"
		.echo "                <input type='radio' name='classtype' onclick=""$('#sssss').show();"""
		if classtype="2" then .echo " checked"
		.echo " value='2'>指定具体分类"
		.echo "<div id='sssss'"
		if classtype="1" then .echo " style='display:none'"
		.echo ">选择分类：<script src=""" & KS.Setting(3) &KS.ASetting(1) &"category.asp?classid=" & BigClassID &"&smallclassid=" & SmallClassID &"&SmallerClassID=" & SmallerClassID &""" language=""javascript""></script></div>"
		%>
				
				</td><td colspan="">仅显示推荐：<input type='radio' name='recommend' value='0'<%if recommend="0" then response.write " checked"%>>否 <input type='radio' name='recommend' value='1'<%if recommend="1" then response.write " checked"%>>是
		<%
				  
.echo "                </td>"

.echo "            </tr>"
		
		
		
		.echo "            <tbody>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">行 号</li><li onclick=""InsertLabel('{@askzjurl}')"">咨询Url</li> <li onclick=""InsertLabel('{@spaceurl}')"">专家空间URL</li><li onclick=""InsertLabel('{@username}')"">专家名称</li> <li onclick=""InsertLabel('{@realname}')"">真实姓名</li><li onclick=""InsertLabel('{@adddate}')"">认证时间</li><li onclick=""InsertLabel('{@askdonenum}')"">回答数</li><li onclick=""InsertLabel('{@userface}')"">头像</li><li onclick=""InsertLabel('{@qq}')"">QQ号</li><li onclick=""InsertLabel('{@tel}')"">电话</li><li onclick=""InsertLabel('{@province}')"">省份</li><li onclick=""InsertLabel('{@city}')"">城市</li><li onclick=""InsertLabel('{@intro}')"">简介</li><li onclick=""InsertLabel('{@askclassname}')"">一级分类</li><li onclick=""InsertLabel('{@asksubclassname}')"">二级分类</li></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'><textarea name='LabelStyle' onkeyup='setPos()' onclick='setPos()' id='LabelStyle' style='width:95%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>使用说明 :</font></strong><br />循环标签[loop=n][/loop]对可以省略,也可以平行出现多对；</td>"
		.echo "            </tr>"
		.echo "           </tbody>"


.echo "                  </table>"	
.echo "  </form>"
  
.echo "</div>"
.echo "</body>"
.echo "</html>"
End With

End Sub
End Class
%> 
