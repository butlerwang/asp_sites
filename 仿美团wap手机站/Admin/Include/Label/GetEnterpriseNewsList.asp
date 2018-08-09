<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<!--#include file="LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New GetSpaceList
KSCls.Kesion()
Set KSCls = Nothing

Class GetSpaceList
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
Dim InstallDir, CurrPath, FolderID, LabelContent, Action, LabelID, Str, Descript,CallBySpace
Dim TypeFlag, Num, TitleLen,ChannelID,PrintType,AjaxOut,LabelStyle,ClassID,OrderStr,BigClassID,SmallClassID,DateRule
FolderID = Request("FolderID")
CurrPath = KS.GetCommonUpFilesDir()
With KS
'判断是否编辑
LabelID = Trim(Request.QueryString("LabelID"))
If LabelID = "" Then
  Action = "Add"
  DateRule="YYYY-MM-DD"
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
    LabelRS.Close
    Set LabelRS = Nothing
            LabelStyle         = KS.GetTagLoop(LabelContent)
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetEnterpriseNewsList", ""),"}" & LabelStyle&"{/Tag}", "")
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
			    ClassID          = Node.getAttribute("bigclassid")
				SmallClassID     = Node.getAttribute("smallclassid")
				DateRule         = Node.getAttribute("daterule")
				Num              = Node.getAttribute("num")
				TitleLen         = Node.getAttribute("titlelen")
				AjaxOut          = Node.getAttribute("ajaxout")
				OrderStr         = Node.getAttribute("orderstr")
				CallBySpace        = Node.getAttribute("callbyspace")
			End If
			XMLDoc=Empty
			Set Node=Nothing
    
End If
		If TitleLen="" Then TitleLen=0
		If KS.IsNUL(CallBySpace) Then CallBySpace=False
		If Num = "" Then Num = 10
		If LabelStyle="" Then LabelStyle="[loop={@num}] " & vbcrlf & "<li><a href=""{@newsurl}"" target=""_blank"">{@title}</a></li>" & vbcrlf & "[/loop]"
		
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
	var ClassID=document.myform.ClassID.value;
	var SmallClassID=document.myform.SmallClassID.value;
	var Num=document.myform.Num.value;
	var TitleLen=document.myform.TitleLen.value;
	var DateRule=document.myform.DateRule.value;
	var OrderStr=$("#OrderStr").val();
	var AjaxOut=false;
	if ($("#AjaxOut").attr("checked")==true){AjaxOut=true}
	var CallBySpace=false;
	if ($("#CallBySpace").attr("checked")==true) CallBySpace=true;
			
	if (Num=='') Num=10
	
	var tagVal='{Tag:GetEnterpriseNewsList labelid="0" ajaxout="'+AjaxOut+'" callbyspace="'+CallBySpace+'" bigclassid="'+ClassID+'" smallclassid="'+SmallClassID+'" num="'+Num+'" orderstr="'+OrderStr+'" daterule="'+DateRule+'" titlelen="'+TitleLen+'"}'+$("#LabelStyle").val()+'{/Tag}';
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
		.echo "   <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""2"">"
		.echo " <input type=""hidden"" name=""Action"" id=""Action"" value=""" & Action & """>"
		.echo "  <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetSpaceList.asp"">"
		.echo ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"

		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">输出格式"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">采用Ajax输出</label>"
		.echo "<strong><font color=blue><input type='checkbox' name='CallBySpace' onclick=""if (this.checked){alert('提示：勾选后该标签只能放在我个人/企业空间模板');}"" id='CallBySpace' value='1'"
		If cbool(CallBySpace)=true Then .echo " checked"
		.echo ">用于在个人/企业空间模板调用</font></strong>"
		.echo "</td><td>日期格式：" & ReturnDateFormat(DateRule) & "</td>"
		.echo "            </tr>"
		
		
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"">显示条数"
.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num"" style=""width:50px"" value=""" & Num & """>个 显示字数<input name=""TitleLen"" class=""textbox"" type=""text"" id=""TitleLen"" style=""width:50px"" value=""" & TitleLen & """><font color=red>如果不想控制，请设置为“0”</font></td>"

.echo "              <td height=""30"">排序方式"
.echo "                <select style=""width:70%;"" class='textbox' name=""OrderStr"" id=""OrderStr"">"
					If OrderStr = "ID Desc" Then
					.echo ("<option value='ID Desc' selected>新闻ID(降序)</option>")
					Else
					.echo ("<option value='ID Desc'>新闻ID(降序)</option>")
					End If
					If OrderStr = "ID Asc" Then
					.echo ("<option value='ID Asc' selected>新闻ID(升序)</option>")
					Else
					.echo ("<option value='ID Asc'>新闻ID(升序)</option>")
					End If

					
					
					If OrderStr = "B.Hits Asc" Then
					 .echo ("<option value='B.Hits Asc' selected>点击数(升序)</option>")
					Else
					 .echo ("<option value='B.Hits Asc'>点击数(升序)</option>")
					End If
					If OrderStr = "Hits Desc" Then
					  .echo ("<option value='B.Hits Desc' selected>点击数(降序)</option>")
					Else
					  .echo ("<option value='B.Hits Desc'>点击数(降序)</option>")
					End If

		.echo "         </select></td>"
.echo "            </tr>"	


        .echo "           <tr class='tdbg' style='display:none'>"
		.echo "            <td>显示类型:"
		.echo "             <input type='radio' name='ShowType' value='1' onclick=""location.href='?t=1&LabelID=" & KS.S("LabelID") & "&Action=" & KS.S("Action") & "&enterprisetype=0'"">产品行业大类"
		.echo "             <input type='radio' name='ShowType' value='2' onclick=""location.href='?t=1&LabelID=" & KS.S("LabelID") & "&Action=" & KS.S("Action") & "&enterprisetype=1'"">装饰行业大类"
		.echo "            </td>"
		.echo "              <td height=""30"">"
	  
        .echo "</td>"
		.echo "           </tr>"
		
		.echo "            <tr class='tdbg' id='spaceclass'>"
		.echo "              <td height=""30"" colspan='2'>所属行业"
		%>
		<%
		dim rsb,sqlb,enterprisetype:enterprisetype=KS.S("enterprisetype")
		
		If KS.S("Action")="Edit" and ks.s("t")<>"1" Then
		  if conn.execute("select top 1 * from ks_enterpriseclass where id=" & KS.ChkClng(classid)).eof then
		      enterprisetype=1
		  end if
		End If
		enterprisetype=0  '兼容定制

		dim rss,sqls,count
		set rss=server.createobject("adodb.recordset")
		if enterprisetype="1" then
		sqls = "select * from KS_enterpriseClass_zs Where parentid<>0 order by orderid"
		Else
		sqls = "select * from KS_enterpriseClass Where parentid<>0 order by orderid"
		End If
		rss.open sqls,conn,1,1
		%>
          <script language = "JavaScript">
		var onecount;
		subcat = new Array();
				<%
				count = 0
				do while not rss.eof 
				%>
		subcat[<%=count%>] = new Array("<%= trim(rss("id"))%>","<%=trim(rss("parentid"))%>","<%= trim(rss("classname"))%>");
				<%
				count = count + 1
				rss.movenext
				loop
				rss.close
				%>
		onecount=<%=count%>;
		function changelocation(locationid)
			{
			document.myform.SmallClassID.length = 0;
			document.myform.SmallClassID.options[document.myform.SmallClassID.length] = new Option("--小类不限--","0");
 
			var locationid=locationid;
			var i;
			for (i=0;i < onecount; i++)
				{
					if (subcat[i][1] == locationid)
					{ 
						document.myform.SmallClassID.options[document.myform.SmallClassID.length] = new Option(subcat[i][2], subcat[i][0]);
					}        
				}
			}    
		
		</script>
		 <select class="face" name="ClassID" id="ClassID" onChange="changelocation(document.myform.ClassID.options[document.myform.ClassID.selectedIndex].value)" size="1">
		<option value='0'>--大类不限--</option>
		<% 
		set rsb=server.createobject("adodb.recordset")
		if enterprisetype="1" then
        sqlb = "select * from ks_enterpriseClass_zs where parentid=0 order by orderid"
		else
		sqlb = "select * from ks_enterpriseclass where parentid=0 order by orderid"
		end if
        rsb.open sqlb,conn,1,1
		if rsb.eof and rsb.bof then
		else
		    do while not rsb.eof
					  If trim(ClassID)=trim(rsb("id")) then
					  %>
                    <option value="<%=trim(rsb("id"))%>" selected><%=trim(rsb("ClassName"))%></option>
                    <%else%>
                    <option value="<%=trim(rsb("id"))%>"><%=trim(rsb("ClassName"))%></option>
                    <%end if
		        rsb.movenext
    	    loop
		end if
        rsb.close
			%>
                  </select>
                  <font color=#ff6600>&nbsp;*</font>
                  <select class="face" id="SmallClassID" name="SmallClassID">
				   <option value='0'>--小类不限--</option>
                    <%dim rsss,sqlss
						set rsss=server.createobject("adodb.recordset")
						if enterprisetype="1" then
						sqlss="select * from ks_enterpriseclass_zs where parentid="& ks.chkclng(ClassID)&" order by orderid"
						else
						sqlss="select * from ks_enterpriseclass where parentid="& ks.chkclng(ClassID)&" order by orderid"
						end if
						rsss.open sqlss,conn,1,1
						if not(rsss.eof and rsss.bof) then
						do while not rsss.eof
							  if trim(SmallClassID)=trim(rsss("id")) then%>
							<option value="<%=rsss("id")%>" selected><%=rsss("ClassName")%></option>
							<%else%>
							<option value="<%=rsss("id")%>"><%=rsss("ClassName")%></option>
							<%end if
							rsss.movenext
						loop
					end if
					rsss.close
					%>
                </select>
		<%
				  
.echo "                </td>"

.echo "            </tr>"
		
		
		
		.echo "            <tbody>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">行 号</li><li onclick=""InsertLabel('{@newsurl}')"">新闻URL</li> <li onclick=""InsertLabel('{@title}')"">新闻标题</li><li onclick=""InsertLabel('{@adddate}')"">添加时间</li><li onclick=""InsertLabel('{@hits}')"">浏览数</li><li onclick=""InsertLabel('{@username}')"">用户名</li><li onclick=""InsertLabel('{@userid}')"">用户ID</li></td>"
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
