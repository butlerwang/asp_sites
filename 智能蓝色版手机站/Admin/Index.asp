<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
Const CheckNewVersion=true   '是否检测获得官方最新版本信息
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"

Dim KSCls
Set KSCls = New Admin_Index
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Index
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  CheckChannelStatus
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		
		Sub CheckChannelStatus()
		 if application("setmodelstatus")<>ChannelNotOnStr then
		 conn.execute("update ks_channel set channelstatus=0 where channelid<100 and channelid in(" & channelNotOnStr & ")")
		 'conn.execute("update ks_channel set channelstatus=1 where channelid not in(" & channelNotOnStr & ")")
		 application("setmodelstatus")=ChannelNotOnStr
		 Call KS.DelCahe(KS.SiteSN & "_ChannelConfig")
		 end if
		End Sub

		Public Sub Kesion()
		    Call CheckSetting()
			Select Case KS.G("Action")
			 Case "Main" Call KS_Main()
			 Case "ver"  Call GetRemoteVer()
			 Case "copyright" Call CopyRight()
			 Case "setTips" Call setTips()
			 Case Else  Call KS_Index()
			End Select
		End Sub
		
		Sub setTips() 
		  Call KS.settingsave(0,KS.G("v"))
		  KS.Die "success"
		End Sub
		
		Sub KS_Index()
		With Response
		.Write 	"<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
		.Write "<head>"
		.Write "<title>" & KS.Setting(0) & "---网站后台管理</title>"
		.Write "<link href=""images/frame.CSS"" rel=""stylesheet"" type=""text/css"">"
         .Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">" & vbCrLf
		.Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		.Write "<script language='JavaScript' src='../KS_Inc/jquery.js'></script>"
		.Write "<script src=""../ks_inc/kesion.box.js""></script>"
		.Write "<script language=""JavaScript"">" & vbCrLf
		.Write "<!--" & vbCrLf
		.Write "   //保存复制,移动的对象,模拟剪切板功能" & vbCrLf
		.Write "  function CommonCopyCutObj(ChannelID, PasteTypeID, SourceFolderID, FolderID, ContentID)" & vbCrLf
		.Write "  {" & vbCrLf
		.Write "   this.ChannelID=ChannelID;             //频道ID" & vbCrLf
		.Write "   this.PasteTypeID=PasteTypeID;         //操作类型 0---无任何操作,1---剪切,2---复制" & vbCrLf
		.Write "   this.SourceFolderID=SourceFolderID;   //所在的源目录" & vbCrLf
		.Write "   this.FolderID=FolderID;               //目录ID" & vbCrLf
		.Write "   this.ContentID=ContentID;             //文章或图片等ID" & vbCrLf
		.Write "  }" & vbCrLf
		.Write "  function CommonCommentBack(FromUrl)" & vbCrLf
		.Write "  {" & vbCrLf
		.Write "    this.FromUrl=FromUrl;             //保存来源页的地址" & vbCrLf
		.Write "  }" & vbCrLf
		.Write "  //初始化对象实例" & vbCrLf
		.Write " var CommonCopyCut=null;" & vbCrLf
		.Write " var CommonComment=null;" & vbCrLf
		.Write " var DocumentReadyTF=false;" & vbCrLf
		.Write "$(document).ready(function(){" & vbcrlf
		.Write "$(""#leftframe"").height(document.body.clientHeight-$(""#topframe"").height()-$(""#bottomframe"").height());" & vbcrlf
		.Write "$(""#MainFrame"").height($(""#leftframe"").height()-30);" &vbcrlf
		.Write "    if (DocumentReadyTF==true) return;" & vbCrLf
		.Write "    CommonCopyCut=new CommonCopyCutObj(0,0,0,'0','0');" & vbCrLf
		.Write "    CommonComment=new CommonCommentBack(0);" & vbCrLf
		.Write "    DocumentReadyTF=true;" & vbCrLf
		.Write "});" &vbcrlf

		.Write " function out(src){"& vbcrlf
		.Write " if(confirm('确定要退出吗？')){"& vbcrlf
		.Write " return true;	}"& vbcrlf
		.Write "  return false;"& vbcrlf
		.Write " }"& vbcrlf
		.Write "function modifyPass(){" &vbcrlf
		.Write " var p=new parent.KesionPopup();" &vbcrlf
		.Write "     p.FadeInTime=p.FadeOutTime=800;" &vbcrlf
		.Write "     p.PopupCenterIframe('修改后台登录密码','KS.Admin.asp?Action=SetPass',520,265,'auto');" & vbcrlf
		.Write "}" & vbcrlf
		.Write " function getNewMessage()"& vbcrlf
		.Write " {"& vbcrlf
		.Write "  var url = '';"   & vbCrLf
		.Write "  jQuery.get(url,{action:'GetAdminMessage'},function(d){jQuery('#newmessage').html(d);});" & vbCrLf
		.Write " }"& vbCrlf
		
		.Write "function setCookieTips(tf){" & vbCrLf
		.Write "var v=0;" & vbCrlf
		.Write "if (tf){ v=0;}else{v=1;} "& vbCrlf
		.Write "jQuery.ajax({ " & vbCrlf
		.Write "url: ""index.asp""," & vbCrlf
		.Write "cache: false," & vbCrlf
		.Write "data: ""action=setTips&v=""+v," & vbCrlf
		.Write "success: function(d){ if (d!='success'){alert(d);}}});" & vbCrLf
		.Write "}" & vbCrLf
		.Write "setTimeout('getNewMessage()', 3000);" & vbCrLf
		
		.Write "//-->" & vbCrLf
		.Write "</script>" & vbCrLf
		
		.Write "</head>" & vbCrLf
		.Write "<body style=""overflow:hidden"" scroll=""no"">"
		%>
	       <table width="100%" height="100%" cellpadding="0" cellspacing="0" border="0">
		   <tr>
		    <td colspan="2" id="topframe" height="70" class="head">
		<%
		.Write "<div id='ajaxmsg' style='text-align:center;background-color: #ffffee;border: 1px #f9c943 solid;position:absolute; z-index:1; left: 200px; top: 5px;display:none;'> <img src='images/loading.gif'> 请稍候,正在执行您的请求...  </div>"
			.Write "        <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "          <tr>"
			.Write "            <td rowspan='2' width='160' style='color:#fff;text-align:center'><span  class='logo'>智能管理系统</span><br/>版本号：V9.05</td>"			
			.Write "            <td><table width=""100%"" celspacing=""0"" cellpadding=""0"" height=""100%"" border=""0"">"
			.Write "             <tr>"
			 Dim KSAnnounceDisplayFlag
			 If Instr(KS.Setting(16),"1")=0 Then
			  .Write "                <td class=""font_text"" width=""40%""><script language=""JavaScript"" src=""../ks_inc/time/3.js""></script></td>"

			  KSAnnounceDisplayFlag=" style=""display:none"""
			 Else
			  KSAnnounceDisplayFlag=""
			 End If
			 .write "                 <td " & KSAnnounceDisplayFlag & " class=""font_text"" align=""right""><font color=#ffffff>官方公告：</font></td>"
			 .Write "                 <td " & KSAnnounceDisplayFlag & "  width=""40%"">"
			 .Write "暂无公告！"
			 .Write "</td>"
			 
			.Write "                <td class=""font_text"" align=""right""> [<a href=""" & KS.GetDomain &""" target=""_blank"" class=""white"">网站首页</a>]"
			If KS.ReturnPowerResult(0, "KMUA10010") Then
			.Write "[<a href=""#"" onClick=""modifyPass()"" class=""white"">修改密码</a>] "
			End If
			If KS.ReturnPowerResult(0, "KMST20000") Then
			.Write "[<a href=""KS.CleanCache.asp"" target=""MainFrame"" class=""white"">更新缓存</a>] "
			End If
			.WRite "[<a href=""Login.asp?Action=LoginOut"" target=""_top"" onClick=""return out(this)""  class=""white"">安全退出</a>]"
			
			.Write "               </td>"
			.Write "              </tr>"
			.Write "            </table></td>"
			.Write "          </tr>"
			%>
			<tr>
			<td>
			<div>
			     <ul id="TabPage">
					<li class="Selected" id="left_tab1" title="内容管理" onClick="javascript:showleft(1);" name="left_tab1">内容</li>
					<li id="left_tab2" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"sysset0")>0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(2);" title="系统管理">设置</li>		
					<li id="left_tab3" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"subsys1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(3);" title="相关操作">相关</li>
					<li id="left_tab4" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"model1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(4);" title="模型管理">模型</li>
					<li id="left_tab5" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"lab1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(5);" title="标签">标签</li>
					<li id="left_tab6">备用</li>
					<li id="left_tab7" title="插件" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"other1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(7);" name="left_tab7">插件</li>
			     </ul><span><a style=" display:block; height:30px;line-height:30px; font-size:14px; color:#fff; padding:0 10px; margin-right:10px; float:right; background:#cc0000; font-weight:900;" target="MainFrame" href="index.asp?action=Main">后台首页</a></span>
			 </div>	
			<%
			.Write "        </table>"
			
	%>
			 	
			 </td>
			 </tr>
		   <tr>
		    <td id="leftframe" style="width:300px;_width:250px;" rowspan="2" valign="top">
			 <div style="overflow-x:hidden;overflow-y:scroll;height:100%">
			<%
			CALL ks_left
			%>
			 </div>
			</td>
			<td valign="top" width="10000">
			 <%
			   Dim MainUrl:MainUrl="Index.asp?action=Main"
			   If Not KS.IsNul(Request("From")) Then
			       MainUrl=Request("From")
			   End If
			 %>
			  <iframe src="<%=MainUrl%>" noresize name="MainFrame" id="MainFrame"" frameborder="no" scrolling="auto"" marginwidth="0"  marginheight="0" width="100%" height="100%"></iframe>
				 
			 </td>
		   </tr>
           <tr>
			 <td height="30" valign="top">
				<iframe src="KS.Split.asp?ButtonSymbol=Disabled&OpStr=<%=Server.URLEncode("系统管理中心 >> 首页")%>" name="BottomFrame" ID="BottomFrame" frameborder="no" height="30"  scrolling="no" width="100%" marginwidth="0" marginheight="0"></iframe>
			 </td>
		  </tr>		   
		   <tr>
		     <td colspan="2" id="bottomframe" height="32">
			 <%
			 KS_Foot
			 %>
			 </td>
		   </tr>
		  </table>
	<%
		.Write "</body>" & vbCrLf
		.Write "</html>" & vbCrLf
		  If KS.S("C")="1" Then
					 On Error Resume Next
					 Dim FileContent
					 FileContent=KS.ReadFromFile("../KS_Inc/ajax.js")
					 FileContent=GetAjaxInstallDir(FileContent,installdir)
					 Call KS.WriteTOFile("../KS_Inc/ajax.js", FileContent)
					 If Err Then
					  err.clear
					 End If
		  End If	
		End With
		End Sub
		

		Function GetAjaxInstallDir(Content,byval installdir)
			 Dim regEx, Matches, Match
			 Set regEx = New RegExp
			 regEx.Pattern="var installdir='[\S]*';"
			 regEx.IgnoreCase = True
			 regEx.Global = True
			 Set Matches = regEx.Execute(Content)
			 If Matches.count > 0 Then
			  GetAjaxInstallDir=Replace(content,Matches.item(0),"var installdir='" & installdir & "';")
			 Else
			  GetAjaxInstallDir="var installdir='/';"
			 end if
		End Function
		
		
		Public Sub KS_Left()
		Dim SQL,I,ModelXML
		Dim RSC:Set RSC=Conn.Execute("Select ChannelID,ChannelName,ChannelTable,ItemName,BasicType,ModelEname,ChannelStatus From KS_Channel Order By ChannelID")
		If Not RSC.Eof Then
		  SQL=RSC.GetRows(-1)
		  Set ModelXML=KS.ArrayToxml(SQL,RSC,"row","ModelXML")
		End If
		RSC.Close:Set RSC=Nothing
		
		
		on error resume next

		With Response
		.Write "<script language=""javascript"">"
		.Write " var ChannelID=null;" & vbcrlf
		.Write " var BasicType=null;" & vbcrlf
		For I=0 To Ubound(SQL,2)
		 .Write " var SearchPower" & SQL(0,I) & "='" & KS.ReturnPowerResult(SQL(0,I), "M"&SQL(0,I)&"10008")&"';    //搜索权限" & vbCrLf
       Next
		.Write " var SearchSpecialPower='" & KS.ReturnPowerResult(0, "KMSP10004") & "';    //搜索专题权限" & vbCrLf
		.Write " var SearchLinkPower='" & KS.ReturnPowerResult(0, "KMCT10001") & "';       //搜索友情链接的权限" & vbCrLf
		.Write " var SearchAdminPower='" & KS.ReturnPowerResult(0, "KMUA10001") & "';      //搜索管理员权限" & vbCrLf
		.Write " var SearchSysLabelPower='" & KS.ReturnPowerResult(0, "KMTL10001") & "';   //搜索系统函数标签权限" & vbCrLf
		.Write " var SearchDIYFunctionLabelPower='" & KS.ReturnPowerResult(0, "KMTL10002") & "';   //搜索自定义函数标签权限" & vbCrLf
		.Write " var SearchFreeLabelPower='" & KS.ReturnPowerResult(0, "KMTL10003") & "';  //搜索自定义静态标签权限" & vbCrLf
		.Write " var SearchSysJSPower='" & KS.ReturnPowerResult(0, "KMTL10004") & "';      //搜索系统JS权限" & vbCrLf
		.Write " var SearchFreeJSPower='" & KS.ReturnPowerResult(0, "KMTL10005") & "';     //搜索自由JS权限" & vbCrLf
		.Write "</script>"
		.Write "<script language=""JavaScript"" src=""Include/SetFocus.js""></script>"
		%>
		<script language="JavaScript">
		var normal='#26517B';     //color;
		var zindex=10000;         //z-index;
		var openTF=false;
		var width=140,height=window.document.body.offsetHeight-15;
		var left=0,top=0,title='搜索小助理';
		var SearchBodyStr=''
						   +'<table style="width:140px;height:100%" border="0" cellspacing="0" cellpadding="0">'
						   +'<form name="searchform" target="MainFrame" method="post">'
						   +'<tr> '
						   +'<td height="25"><strong>按下面任意或全部条件进行搜索</strong></td>'
						   +' </tr>'
						   +'<tr><td height="25">全部或部分关键字</td></tr>'
						   +'<tr><td height="25"><input style="width:90%" type="text" id="KeyWord" name="KeyWord"></td></tr>'
						   +'  <tr><td height="25">搜索范围</td></tr>'
						   +'  <tr><td height="25"> <select style="width:95%" id="SearchArea" name="SearchArea" onchange="SetSearchTypeOption(this.options[this.selectedIndex].text)">'
						   +'     </select></td></tr>' 
						   +'<tr><td height="25">搜索类型</td></tr>'
						   +'<tr><td height="25"><select style="width:95%" id="SearchType" name="SearchType">'
						   +'</select></td></tr>'
						   +'  <tr id="DateArea" onclick="setstatus(this)" style="cursor:pointer"><td height="25"><strong>什么时候修改的?</strong></td></tr>'
						   +'  <tr style="display:none"><td height="25">开始日期<input type="text" style="width:80%" name="StartDate" id="StartDate">'
						   +'  </td></tr>'
						   +'  <tr style="display:none"><td height="25">结束日期<input type="text" style="width:80%" name="EndDate" id="EndDate">'
						   +' </td></tr>'
						   +'  <tr><td height="40" align="center"><input type="submit" name="SearchButton" value="开始搜索" onclick="return(SearchFormSubmit())"></td></tr>'
						   +'</form>'
						   +'  <tr><td><br/><br/><strong>使用说明:</strong></td></tr>'
						   +'  <tr><td valign="top" style="height:250px;padding-right:10px;"> ① 您可以利用本搜索助理来搜索文章、图片、下载Flash、专题、标签、JS等,但不能搜索（目录）诸如频道名称、栏目名称，标签目录等<br/>'
						   +'  ② 按 <font color=red>Ctrl+F</font> 可以快速进行打开或关闭搜索小助理<br/><br/><br/><br/></td></tr>'
						   +'</table>'
				var str=""
					   +"<div id='SearchBox' style='display:none;z-index:" + zindex + ";width:" + width + "px;height:" + height + "px;left:" + left + "px;top:" + top + "px;background-color:" + normal + ";color:black;font-size:12px;font-family:Verdana, Arial, Helvetica, sans-serif;position:absolute;cursor:default;border:7px solid " + normal + ";'>"
					   + "<div style='background-color:" + normal + ";width:" + (width) + "px;height:22px;color:white;'>"
									   + "<span style='width:" + (width-2*12-4) + ";padding-left:3px;font-weight:bold;'>" + title + "</span>"
									   + "<span id='Close' style='padding:50px;cursor:hand;padding-right:0px;width:20;border-width:0px;color:white;font-family:webdings;' onclick='CloseSearchBox()'>r</span>"
					   + "</div>"
					   + "<div style='width:140;overflow:hidden;height:" + (height-20-4) + ";background-color:white;line-height:14px;word-break:break-all;padding:6px;'>" + SearchBodyStr + "</div>"
					   + "</div>"
					   + "<div style='display:none;width:" + width + "px;height:" + height + "px;top:" + top + "px;left:" + left + "px;z-index:" + (zindex-1) + ";position:absolute;background-color:black;filter:alpha(opacity=40);'></div>";
		//关闭;
		function CloseSearchBox(){$("#SearchBox").hide('slow');openTF=false;SearchBodyStr=null;str=null;}
		function initial()
		{if (!openTF){document.body.insertAdjacentHTML("beforeEnd",str);openTF=true;}
		}
		//初始化;
		function initializeSearch(SearchArea,sChannelID,sBasicType)
		{
		 initial();
		 initialSearchAreaOption(SearchArea);
		 ChannelID=sChannelID;
		 BasicType=sBasicType;
		if (jQuery('#SearchBox')[0].style.display=='none')
		 {jQuery('#SearchBox').show('fast');
		  if (document.forms[0].disabled==false) document.forms[0].focus();
		 }
		 else
		 jQuery('#SearchBox').hide('fast');
		}
		<%
		 Dim ModelList,ModelEList,ChannelIDList
		 For I=0 To Ubound(SQL,2)
		  If SQL(0,I)<>6 and SQL(6,I)=1 Then
			  ModelList=ModelList & "'" & SQL(1,I) & "',"
			  ModelElist=ModelElist & "'" & SQL(4,I) & "',"
			  ChannelIDList=ChannelIDList & "'" & SQL(0,I) &"',"
		  End If
		 Next
		%>
		var sTextArr,ChannelIDArr;
		function initialSearchAreaOption(SearchArea)
		{	 var EF=false;
			 sTextArr=new Array(<%=ModelList%>'专题中心','友情链接站点','系统函数标签','自定义函数标签','自定义静态标签','系统 JS','自由 JS','管理员')
			 ChannelIDArr=new Array(<%=ChannelIDList%>'专题中心','友情链接站点','系统函数标签','自定义函数标签','自定义静态标签','系统 JS','自由 JS','管理员')
			 var valueArr=new Array(<%=ModelElist%>'Special','Link','SysLabel','DIYFunctionLabel','FreeLabel','SysJS','FreeJS','Manager')
			  for(var i=0;i<valueArr.length;++i)
			   if (SearchArea==sTextArr[i]){ 
				  EF=true;
				  break;
				 }
			  if (!EF) return false; 
			  jQuery('#KeyWord').val('');
			  jQuery('#SearchArea').empty();
			  for (var i=0;i<sTextArr.length;++i)
				{
				   if (SearchArea==sTextArr[i]){
					jQuery('#SearchArea').append("<option value='"+valueArr[i]+"' selected>"+sTextArr[i]+"</option>");
					}else{
					jQuery('#SearchArea').append("<option value='"+valueArr[i]+"'>"+sTextArr[i]+"</option>");
					}
				} 
			//进行权限检查,对没有权限的搜索模块,进行屏蔽	
			 var n=0;
			for (var i=1000;i<sTextArr.length;++i)
			   {   var removeTF=false;
				   if (valueArr[i]!=SearchArea)
				  { 
				  
				  <%For I=0 To Ubound(SQL,2)
				    If SQL(6,I)=1 Then 
				   %>
				  if (SearchPower<%=SQL(0,i)%>=='False')
					   removeTF=true;
				  <%
				    End If
				  NEXT%>
		 
					if (valueArr[i]=='Special' && SearchSpecialPower=='False') removeTF=true;
					if (valueArr[i]=='Link' && SearchLinkPower=='False') removeTF=true;
					if (valueArr[i]=='SysLabel' && SearchSysLabelPower=='False') removeTF=true;
					if (valueArr[i]=='DIYFunctionLabel' && SearchDIYFunctionLabelPower=='False') removeTF=true;
					if (valueArr[i]=='FreeLabel' && SearchFreeLabelPower=='False') removeTF=true;
					if (valueArr[i]=='SysJS' && SearchSysJSPower=='False')  removeTF=true;
					if (valueArr[i]=='FreeJS' && SearchFreeJSPower=='False')  removeTF=true;
					if (valueArr[i]=='Manager' && SearchAdminPower=='False') removeTF=true;
				   }
				  if (removeTF==true)  
					{document.all.SearchArea.options.remove(i-n);
					 n++;
					}	
			   }
			SetSearchTypeOption(SearchArea); 
		}
		function SetSearchTypeOption(AreaType)
		{	
			  //改变选择范围时，取得正确的模型ID
			  for(var i=0;i<sTextArr.length;++i)
			   if (AreaType==sTextArr[i]) 
				{ 
				  ChannelID=ChannelIDArr[i];
				  break;
				 }

			var TextArr=new Array();
			jQuery('#SearchType').empty();
		  switch (AreaType)
		  {
		   <%For I=0 To Ubound(SQL,2)
		      If SQL(6,I)=1 Then 
			%>
			case '<%= SQL(1,I)%>':
				 if (SearchPower<%= SQL(0,I)%>=='False')          //搜索权限检查
				 {
				  DisabledSearchFluctuation(true);
				  return;
				 }
				 else
				 {
				  DisabledSearchFluctuation(false);
				  jQuery('#DateArea').show();
				  TextArr=new Array('<%=SQL(3,I)%>标题','<%=SQL(3,I)%>内容','<%=SQL(3,I)%>关键字','<%=SQL(3,I)%>作者','<%=SQL(3,I)%>录入')
				  }
				  break;
		   <% End If
		   Next%>
			case '专题中心':
				 if (SearchSpecialPower=='False')        //搜索专题权限检查
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }
				 else
				 {
				  DisabledSearchFluctuation(false);
				  jQuery('#DateArea').show();
				  TextArr=new Array('专题名称','简要说明')
				 }
				 break;
			case '友情链接站点':
				 if (SearchLinkPower=='False'){DisabledSearchFluctuation(true); return; }
				 else{ DisabledSearchFluctuation(false);jQuery('#DateArea').show();TextArr=new Array('站点名称','站点描述');}
				 break;
			case '系统函数标签':
				 if (SearchSysLabelPower=='False'){DisabledSearchFluctuation(true);return;}
				 else{DisabledSearchFluctuation(false);jQuery('#DateArea').show();TextArr=new Array('系统标签名称','系统标签描述');}
				 break;
			case '自定义函数标签':
				 if (SearchDIYFunctionLabelPower=='False')       //搜索自定义函数标签权限检查
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }else{
				  DisabledSearchFluctuation(false);
				 jQuery('#DateArea').show();
				 TextArr=new Array('自定义函数标签名称','自定义函数标签描述')
				 }
				 break;
			case '自定义静态标签':
				 if (SearchFreeLabelPower=='False')       //搜索自定义静态标签权限检查
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }else{
				 jQuery('#DateArea').show()
				 TextArr=new Array('自定义静态标签名称','自定义静态标签描述','自定义静态标签内容')
				 }
				 break;
			case '系统 JS':
				 if (SearchSysJSPower=='False')       //搜索系统JS权限检查
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }else{
				 jQuery('#DateArea').show();
				 TextArr=new Array('系统JS 名称','系统JS 描述','系统JS 文件名')
				 }
				 break;
			case '自由 JS' :
				 if (SearchFreeJSPower=='False')       //搜索自由JS权限检查
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }else{
				 jQuery('#DateArea').show();
				 TextArr=new Array('自由JS 名称','自由JS 描述','自由JS 文件名')
				 }
				 break;
			case '管理员':	 
				  if (SearchAdminPower=='False')          //搜索管理员权限检查
				 {
				  DisabledSearchFluctuation(true);
				  return;
				 }else{
				  DisabledSearchFluctuation(false);
				 jQuery('#DateArea').show();
				 TextArr=new Array('管理员名称','管理员简介')
				}
				break;
		  }
		  for (var i=0;i<TextArr.length;++i){
			jQuery('#SearchType').append("<option value='"+i+"'>"+TextArr[i]+"</option>");
			}
		}
		function setstatus(Obj)
		  {var today=new Date()
			if (Obj.nextSibling.style.display=='none')
			 {
			  Obj.nextSibling.style.display='';
			  jQuery('#StartDate').val(today.getYear()+'-'+(today.getMonth()+1)+'-'+today.getDate());
			 }
			else 
			{
			 Obj.nextSibling.style.display='none';
			 jQuery('#StartDate').val('');
			 }
			if (Obj.nextSibling.nextSibling.style.display=='none')
			{
			 Obj.nextSibling.nextSibling.style.display='';
			  jQuery('#EndDate').val(today.getYear()+'-'+(today.getMonth()+1)+'-'+today.getDate());
			}
			else 
			 {
			 Obj.nextSibling.nextSibling.style.display='none';
			 jQuery('#EndDate').val('');
			 }
		  }
		 function SearchFormSubmit()
		  { var form=document.forms[0];
			if (form.elements[0].value=='')
			 {
			   alert('请输入关键字!')
			   form.elements[0].focus();
			   return false;
			 }
		   switch (form.elements[1].value)
			{
			  case '1':
			  case '2':
			  case '3':
			  case '4':
			  case '5':
			  case '7':
			  case '8':
				   form.action="KS.ItemInfo.asp?ChannelID="+ChannelID;
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("文档搜索管理 >> <font color=red>搜索结果</font>")+'&ButtonSymbol=Search';
				   break;
			  case 'Special':
				   form.action="KS.Special.asp?Action=SpecialList";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("专题管理 >> <font color=red>搜索专题结果</font>")+'&ButtonSymbol=SpecialSearch';
				   break;
			  case 'Link':
				   form.action="KS.FriendLink.asp";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("常规管理 >> 友情链接管理 >> <font color=red>搜索友情链接站点结果</font>")+'&ButtonSymbol=LinkSearch';
				   break;
			  case 'SysLabel'  :
				   form.action="Include/Label_Main.asp?LabelType=0";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("标签管理 >> <font color=red>搜索系统函数标签结果</font>")+'&ButtonSymbol=SysLabelSearch';
				   break;
			 case 'DIYFunctionLabel'  :
				   form.action="Include/Label_Main.asp?LabelType=5";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("标签管理 >> <font color=red>搜索自定义函数标签结果</font>")+'&ButtonSymbol=DIYFunctionSearch';
				   break;
			  case 'FreeLabel'  :
				   form.action="Include/Label_Main.asp?LabelType=1";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("标签管理 >> <font color=red>搜索自由标签结果</font>")+'&ButtonSymbol=FreeLabelSearch';
				   break;
			  case 'SysJS'     :
				   form.action="Include/JS_Main.asp?JSType=0";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("JS管理 >> <font color=red>搜索系统JS结果</font>")+'&ButtonSymbol=SysJSSearch';
				   break;
			  case 'FreeJS'     :
				   form.action="Include/JS_Main.asp?JSType=1";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("JS管理 >> <font color=red>搜索自由JS结果</font>")+'&ButtonSymbol=FreeJSSearch';
				   break;
			  case 'Manager'     :
				   form.action="KS.Admin.asp";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("管理员管理 >> <font color=red>搜索管理员结果</font>")+'&ButtonSymbol=ManagerSearch';
				   break;
			}
			form.submit();
		  }
		function DisabledSearchFluctuation(Flag)
		{ if (Flag==true)
		   document.all.KeyWord.value='对不起,权限不足!'; 
		  var AllBtnArray=document.body.getElementsByTagName('INPUT'),CurrObj=null;
			for (var i=0;i<AllBtnArray.length;i++)
			{
				CurrObj=AllBtnArray[i];
				CurrObj.disabled=Flag;
			}
			AllBtnArray=document.body.getElementsByTagName('SELECT'),CurrObj=null;
			for (var i=0;i<AllBtnArray.length;i++)
			{
				CurrObj=AllBtnArray[i];
				CurrObj.disabled=Flag;
			}
		}
		</script>
		<table border=0 cellPadding=0 cellSpacing=0>
		  <tr>
		  <td class="lefttop"></td>
		  </tr>
		  <tr vAlign=top>
		
			<td align="center" class="boxright">
			 
			    <div>
				  <div id="menubox">
					<ul class="leftbox" id="dleft_tab1">
					 <% dim n:n=0%>
					 
					 <!--------------内容管理 start-------------------->
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);">&nbsp;&nbsp;<a href="javascript:void(0)">内容管理</a></DIV>
					  <div class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					     <div class="modellist">
						  <ul>
					  <%
					   For I=0 To Ubound(SQL,2)
					      If SQL(6,I)=1 Then 
						   IF instr(KS.C("ModelPower"),sql(5,i) & "0")=0 and SQL(0,I)<>6 and SQL(0,I)<>9 And SQL(0,I)<>10 Then
						   Dim ItemManageUrl
						   Select Case  SQL(4,I)
							Case 1 :ItemManageUrl="KS.Article.asp"
							Case 2 :ItemManageUrl="KS.Picture.asp"
							Case 3 :ItemManageUrl="KS.Down.asp"
							Case 4 :ItemManageUrl="KS.Flash.asp"
							Case 5 :ItemManageUrl="KS.Shop.asp"
							Case 7 :ItemManageUrl="KS.Movie.asp"
							Case 8 :ItemManageUrl="KS.Supply.asp"
						   End Select
						  
						   %>
						   <li>
						   <a href="javascript:void(0)" title="<%=SQL(1,I)%>" onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red><%=SQL(3,I)%>管理</font>','ViewFolder','KS.ItemInfo.asp?ChannelID=<%=SQL(0,I)%>',<%=SQL(0,I)%>);"><%=KS.Gottopic(SQL(1,I),8)%></a> <span style="cursor:pointer" onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>添加<%=SQL(3,I)%></font>','AddInfo','<%=ItemManageUrl%>?Action=Add&ChannelID=<%=SQL(0,I)%>',<%=SQL(0,I)%>);"><img alt="添加<%=SQL(3,I)%>" src="images/add.gif" border="0" align="absmiddle"></span><%if KS.ReturnPowerResult(SQL(0,I), "M"&SQL(0,I)&"10012") then%> <span style="cursor:pointer" onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>签收<%=SQL(3,I)%></font>','Disabled','KS.ItemInfo.asp?ComeFrom=Verify&ChannelID=<%=SQL(0,I)%>',<%=SQL(0,I)%>);"><img alt="签收<%=SQL(3,I)%>" src="images/accept.gif" border="0" align="absmiddle"></span>
						   <%end if%>
						   </li>
						   <%
						   End If
						 End If
					   Next
					   %>
					    </ul>
					     </div> 
						    <div id='classOpen' style="margin-top:5px;"></div>
						  
                          <div class="modelxg">
						  <script type="text/javascript">
						   var toggle=getCookie("ctips");
						   if (toggle==null) toggle='show';
							$(document).ready(function(){
							TipsToggle(toggle);
							})
						   function TipsToggle(f){
						    setCookie("ctips",f);
							 if (f=='hide'){
							 jQuery("#modelxg").hide('fast');
							 jQuery("#classOpen").html("<img style='cursor:pointer' id='classOpen' onclick='TipsToggle(\"show\")' src='images/left_down.gif' align='absmiddle' title='展开'>");
							 }else{
							 jQuery("#modelxg").show('fast');
							 jQuery("#classOpen").html("<img style='cursor:pointer' id='classOpen' onclick='TipsToggle(\"hide\")' src='images/left_up.gif' title='收藏' align='absmiddle'>");						
                              	 }
						   }
						  </script>
						  
                           <div  id="modelxg">
						    <ul>
						   <%If KS.ReturnPowerResult(0, "M010001") Then %>
						   <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'内容管理 >> <font color=red><%=SQL(3,I)%>栏目管理</font>','Disabled','KS.Class.asp');">栏目管理</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'栏目管理 >> <font color=red>添加栏目</font>','Go','KS.Class.asp?Action=Add&FolderID=1','');">添加</a></li>
						   <%End If%>
						   <%If KS.ReturnPowerResult(0, "M010002") Then %>
						    <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'内容管理 >> <font color=red>评论管理</font>','Disabled','KS.Comment.asp');">评论管理</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'内容管理 >> <font color=red>评论管理</font>','Disabled','KS.Comment.asp?ComeFrom=Verify');">审核</a> </li>
							<%End If%>
							<%If KS.ReturnPowerResult(0, "M010003") Then %>
						    <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'内容管理 >> <font color=red>专题管理</font>','Disabled','KS.Special.asp');">全站专题管理</a> </li>
							<%End If%>
							<%If KS.ReturnPowerResult(0, "M010004") Then %>
						    <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'内容管理 >> <font color=red>关键字Tags管理</font>','Disabled','KS.KeyWord.asp');">关键字Tags管理</a> </li>
							<%End If%>
                            <%If KS.ReturnPowerResult(0, "M010005") or KS.ReturnPowerResult(0, "M010006") Then %>
							<li>
							<%If KS.ReturnPowerResult(0, "M010005") Then%><a href='javascript:void(0)' onClick="SelectObjItem1(this,'内容管理 >> <font color=red>批量设置</font>','Disabled','KS.ItemInfo.asp?Action=SetAttribute');">批量设置</a><%end if%><%If KS.ReturnPowerResult(0, "M010006") then%> <a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'相关操作 >> <font color=red>浏览回收站</font>','ViewFolder','KS.ItemInfo.asp?ComeFrom=RecycleBin','');">回 收 站</a><%end if%></li>
							<%End If%>
						   <%If KS.ReturnPowerResult(0, "M010007") Then %>
						    <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'内容管理 >> <font color=red>一键快速生成HTML</font>','Disabled','include/refreshquick.asp');">一键快速生成HTML</a> </li>
						   <%end if%>
						   <%If KS.ReturnPowerResult(0, "M010008") Then %>
						   <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'内容管理 >> <font color=red>信息采集管理</font>','Disabled','Collect/Collect_Main.asp?ChannelID=1');">信息采集管理</a> </li>
						   <%End if%>
						    </ul>
						   </div>
						

						   
						 </div>
					 </div>
					<!--------------内容管理 end-------------------->  
					
					
					<!--------------商城管理 start-------------------->
				  <%
				  IF instr(lcase(KS.C("ModelPower")),"shop0")=0 or KS.C("SuperTf")=1 Then
					 If Not ModelXML Is Nothing Then
					  If ModelXML.documentElement.SelectNodes("row[@channelid=5 and @channelstatus=1]").length<>0 Then
					   N=N+1
					 %>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">商城管理</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					   <ul>
					    <%If KS.ReturnPowerResult(5, "M510020") Then %>				
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>处理24小时内订单</font>','Disabled','KS.ShopOrder.asp?searchtype=1&ChannelID=5');"><font color=red>处理24小时内订单</font></a></li>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>处理所有订单</font>','Disabled','KS.ShopOrder.asp?ChannelID=5');">处理所有订单</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M510014") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>资金明细查询</font>','Disabled','KS.LogMoney.asp?ChannelID=<%=SQL(0,I)%>');">资金明细查询</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M510015") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>发退货查询</font>','Disabled','KS.LogDeliver.asp?ChannelID=5');">发退货查询</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M510016") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>发退货查询</font>','Disabled','KS.LogInvoice.asp?ChannelID=5');">开发票查询</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M510017") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>客户统计</font>','Disabled','KS.ShopStats.asp?Action=Custom');">销售数据统计</a></li>
						 <%End If%>
						  <%If KS.ReturnPowerResult(5, "M520004") Then %>
						 ====================
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>送货方式管理</font>','Disabled','KS.Delivery.asp?ChannelID=5');">送货&付款方式</a></li>
						 <%end if%>
					   	  <%If KS.ReturnPowerResult(5, "M510018") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'品牌管理 >> <font color=red>品牌管理</font>','Disabled','KS.ShopBrand.asp');">品牌管理</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'品牌管理 >> <font color=red>添加品牌</font>','Go','KS.ShopBrand.asp?Action=Add&FolderID=0',5);">添加</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'品牌管理 >> <font color=red>生成品牌的JS菜单</font>','Go','KS.ShopBrand.asp?Action=Create&FolderID=0',5);">生成</a></li>	
						 <%end if%>	
						 <%If KS.ReturnPowerResult(5, "M520006") Then %>
						  <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>商品规格列表管理</font>','Disabled','KS.ShopSpecification.asp?ChannelID=5');">商品规格列表管理</a> </li>
						 <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520003") Then %>
						  <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>厂商管理</font>','Disabled','KS.Author.asp?ChannelID=5');">厂商管理</a> </li>
						  <%end if%>
						 
						 ====================
						  <%If KS.ReturnPowerResult(5, "M530001") Then %>
						  <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'团购系统 >> <font color=red>团购管理首页</font>','Disabled','KS.GroupBuy.asp');">团购管理首页</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'团购系统 >> <font color=red>团购管理首页</font>','Go','KS.GroupBuy.asp?Action=Add');">添加</a></li>
						  <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520008")  and IsBusiness Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>抢购商品管理</font>','Disabled','KS.Shop.asp?action=LimitBuy&channelid=5');">限时/限量抢购管理</a></li>
						 <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520009") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>限时抢购商品管理</font>','Disabled','KS.Shop.asp?action=BundleSale&channelid=5');">捆绑销售商品管理</a></li>
						 <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520010") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>换购商品管理</font>','Disabled','KS.Shop.asp?action=ChangedBuy&channelid=5');">换购商品管理</a></li>
						 <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520012") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>商品库存报警管理</font>','Disabled','KS.Shop.asp?action=StockAlarm&channelid=5');">商品库存报警管理</a></li>
						 <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520011") and IsBusiness Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>超值礼包管理</font>','Disabled','KS.Shop.asp?action=Package&channelid=5');">超值礼包管理</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M510005") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>批量调价助手</font>','Disabled','KS.ItemInfo.asp?action=SetAttribute&channelid=5');">批量调价助手</a></li>
						 <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520007") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>优惠券管理</font>','Disabled','KS.ShopCoupon.asp');">优惠券管理</a></li>
						 <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520014") and IsBusiness Then %>
                          ====================
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>快递单模板管理</font>','Disabled','KS.ShopExpress.asp');">快递单模板管理</a></li>
						 <%End If%>
						 </ul>
					 </DIV>
					 <!--------------商城管理 End-------------------->
					<% End If
					End If
				  End If
					%>
					
					<!--------------招聘求职 start-------------------->
					 <%
				   IF (instr(lcase(KS.C("ModelPower")),"job0")=0 or KS.C("SuperTf")=1) and IsBusiness Then
					 If Not ModelXML Is Nothing Then
					  If ModelXML.documentElement.SelectNodes("row[@channelid=10 and @channelstatus=1]").length<>0 Then
					   N=N+1
					  %>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">招聘求职</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					  <ul>
					  <%
					  If KS.ReturnPowerResult(10, "M1010005") Then   
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>行业职位设置</font>'") & ",'disabled','KS.Jobhy.asp');"">行业职位设置</a></li>"
					  End If
					  If KS.ReturnPowerResult(10, "M1010006") Then   
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>毕业院校设置</font>'") & ",'disabled','KS.JobSchool.asp');"">毕业院校设置</a></li>"
					  
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>简历模板管理</font>'") & ",'disabled','KS.JobTemplate.asp');"">简历模板管理</a></li>"
						  Response.Write "<div class=""clear""></div>==================="
					  End If
					  If KS.ReturnPowerResult(10, "M1010002") Then  
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>招聘单位管理</font>'") & ",'disabled','KS.JobCompany.asp');"">招聘单位管理</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & Server.UrlEncode("求职招聘 >> <font color=red>审核招聘单位</font>") & "','disabled','KS.JobCompany.asp?ComeFrom=Verify');"">招聘单位审核</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>添加招聘单位</font>'") & ",'disabled','KS.JobCompany.asp?Action=Add');"">添加招聘单位</a></li>"
					  End If
					 If KS.ReturnPowerResult(10, "M1010004") Then  
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>招聘职位管理</font>'") & ",'disabled','KS.Jobzw.asp');"">招聘职位管理</a></li>"
					  End If
					  If KS.ReturnPowerResult(10, "M1010003") Then  
						  Response.Write "<div class=""clear""></div>==================="
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>个人简历管理</font>'") & ",'disabled','KS.JobResume.asp');"">个人简历管理</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>审核个人简历</font>'") & ",'disabled','KS.JobResume.asp?ComeFrom=Verify');"">个人简历审核</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & Server.Urlencode("求职招聘 >> <font color=red>添加个人简历</font>'") & ",'disabled','KS.JobResume.asp?Action=Add');"">添加个人简历</a></li>"
					  
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & Server.Urlencode("求职招聘 >> <font color=red>教育背景管理</font>'") & ",'disabled','KS.JobEdu.asp');"">教育背景管理</a></li>"
					 End If
					  %>
					 <ul>
					 </DIV>
                    <!--------------招聘求职 end--------------------> 
					<% End If
					End If
				   End If
					%>
					
					
					<!--------------考试系统 start-------------------->
				 <%
				IF (instr(lcase(KS.C("ModelPower")),"mnkc0")=0 or KS.C("SuperTf")=1) and IsBusiness Then
					 If Not ModelXML Is Nothing Then
					  If ModelXML.documentElement.SelectNodes("row[@channelid=9 and @channelstatus=1]").length<>0 Then
					   N=N+1
					   %>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">考试系统</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					  <ul>
					 <%
					      Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("考试系统 >> <font color=red>试卷管理</font>'") & ",'disabled','mnkc/mnkc_sort.asp');"">试卷管理</a>&nbsp;<a href='javascript:void(0)' onClick=""SelectObjItem1(this,'考试系统 >> <font color=red>添加试卷</font>','Go','mnkc/mnkc_add_first.asp');"">添加</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("考试系统 >> <font color=red>试卷题库管理</font>'") & ",'disabled','mnkc/mnkc_tk.asp');"">题库管理</a> <a href=""javascript:SelectObjItem1(this,'" & server.urlencode("考试系统 >> <font color=red>从题库选题组卷</font>'") & ",'disabled','mnkc/xtzj.asp');"">组卷</a></li>"
						  If KS.ReturnPowerResult(9, "M910001") Then 
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("考试系统 >> <font color=red>试卷分类管理</font>'") & ",'disabled','mnkc/mnkc_class.asp');"">试卷分类管理</a></li>"
						  End If
					      Response.Write "<li><a href='mnkc/mnkc_score.asp' target='MainFrame'>考试成绩管理</a></li>"
						  If KS.ReturnPowerResult(9, "M910007") Then 
						  Response.Write "==================="
						  Response.Write "<li><a href='mnkc/refreshindex.asp' target='MainFrame'>发布试卷首页等</a></li>"
						  Response.Write "<li><a href='mnkc/refreshallcalss.asp?type=all' target='MainFrame'>发布所有分类</a></li>"
						  Response.Write "<li><a href='mnkc/RefreshSJ.asp' target='MainFrame'>发布所有试卷页</a></li>"
						  End If
					 %>
					 </ul>
					 </DIV>
                    <!--------------考试系统 end--------------------> 
				    <%
					 End If
					End If
			   End If
					%>
					<%IF instr(lcase(KS.C("ModelPower")),"bbs0")=0 or KS.C("SuperTf")=1 Then%>
					<%N=N+1%>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">论坛系统</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					  <ul>
					 <%If KS.ReturnPowerResult(0, "KSMB10000") Then%>
					<li><a href="KS.GuestBook.asp?Action=Main"  target="MainFrame" title="论坛帖子管理">帖子管理</a> <a href="KS.GuestBook.asp?Action=Recycle"  target="MainFrame" title="帖子回收站">回收站</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMB10001") Then%>
					<li><a href="KS.GuestBoard.asp"  target="MainFrame" title="版面分类管理">论坛版面分类管理</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMB10003") Then%>
					<li><a href="KS.GuestTable.asp"  target="MainFrame" title="当前数据表管理">当前数据表管理</a></li>
					<%end if%>
					   <%If KS.ReturnPowerResult(0, "KSMB10004") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'问答系统 >> <font color=red>等级头衔设置</font>','Disabled','KS.AskGrade.asp?typeflag=1');" title="等级头衔设置">等级头衔</a> <a href="javascript:void(0)" onClick="SelectObjItem1(this,'论坛系统 >> <font color=red>勋章管理</font>','Disabled','KS.GuestMedal.asp');" title="勋章管理">勋章管理</a>
					   </li>
					   <%end if%>
					   </ul>
					 </DIV>
					<%end if%>
					
					
					<!--------------问答系统 start-------------------->
					<%IF instr(lcase(KS.C("ModelPower")),"ask0")=0 or KS.C("SuperTf")=1 Then%>
					 <%N=N+1%>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">问答系统</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					  <ul>
					 <%If KS.ReturnPowerResult(0, "WDXT10000") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'问答系统 >> <font color=red>问答参数设置</font>','SetParam','KS.AskSetting.asp');" title="问答参数设置">问答参数设置</a></li>
					   <%end if%>
					   <%If KS.ReturnPowerResult(0, "WDXT10001") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'问答系统 >> <font color=red>问题列表管理</font>','SetParam','KS.AskList.asp');" title="问题列表管理">问题列表管理</a></li>
					   <%End If%>
					   <%If KS.ReturnPowerResult(0, "WDXT10004") Then%>
					   <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'问答系统 >> <font color=red>审核问题回答管理</font>','Disabled','KS.AskList.asp?action=verifyanswer');">审核回答管理</a>
					   <%end if%>
					   <%If KS.ReturnPowerResult(0, "WDXT10002") Then%>
					   <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'问答系统 >> <font color=red>分类管理</font>','Disabled','KS.AskClass.asp');">分类管理</a>
					   <a href='javascript:void(0)' onClick="SelectObjItem1(this,'问答系统 >> <font color=red>添加问答分类</font>','GO','KS.AskClass.asp?action=add');">添加</a></li>
					   <%end if%>
					   <%If KS.ReturnPowerResult(0, "WDXT10003") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'问答系统 >> <font color=red>等级头衔设置</font>','Disabled','KS.AskGrade.asp');" title="等级头衔设置">等级头衔设置</a></li>
					   <%end if%>
					   <%If KS.ReturnPowerResult(0, "WDXT10005") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'问答系统 >> <font color=red>专家认证管理</font>','Disabled','KS.AskZJ.asp');" title="专家认证管理">专家认证管理</a></li>
                       <%end if%>
					   </li>
					   </ul>
					 </DIV>
				   <%End If%>
                    <!--------------问答系统 end--------------------> 
					
					<!--------------空间系统 start-------------------->
				   <%IF instr(lcase(KS.C("ModelPower")),"space0")=0 or KS.C("SuperTf")=1 Then%>
					 <%N=N+1%>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">空间门户</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					 
					    	 <div style=" border:#ff6600 1px dotted;width:115px; height:21px; line-height:21px;margin-left:5px;margin-bottom:2px; margin-top:2px;text-align:left;padding-left:5px; font-size:14px;font-weight:bold; color:#ff6600;"><img src="images/ico_friend.gif">&nbsp;个人空间</div>
							  <ul>
						 <%If cbool(KS.ReturnPowerResult(0, "KSMS10000")) Then%>
						<li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'个人空间 >> <font color=red>空间参数设置</font>','SetParam','KS.SpaceSetting.asp');" title="空间参数设置">空间参数设置</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10001") Then%>
						<li><a href="KS.Space.asp"  target="MainFrame" title="所有空间管理">所有空间管理</a></li>
						<li><a href="KS.Space.asp?showtype=1"  target="MainFrame" title="个人空间管理">个人空间管理</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10002") Then%>
						<li><a href="KS.Spacelog.asp"  target="MainFrame" title="空间博文管理">空间博文管理</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS20016") Then %>
					    <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'空间系统 >> <font color=red>微博数据管理</font>','Disabled','KS.UserLog.asp');" title="微博数据管理">微博数据管理</a></li>
					  <%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10003") Then%>
						<li><a href="KS.SpaceAlbum.asp"  target="MainFrame" title="用户相册管理">用户相册管理</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10004") Then%>
						<li><a href="KS.SpaceTeam.asp"  target="MainFrame" title="用户圈子管理">用户圈子管理</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10005") Then%>
						<li><a href="KS.SpaceMessage.asp"  target="MainFrame" title="用户留言管理">用户留言管理</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10007") Then%>
						<li><a href="KS.SpaceMusic.asp"  target="MainFrame" title="用户歌曲管理">用户歌曲管理</a></li>
						<%end if%>
						</ul>
						 <div style=" border:#ff6600 1px dotted;width:115px; height:21px; line-height:21px;margin-left:5px; text-align:left;padding-left:5px; font-size:14px;font-weight:bold; color:#ff6600;"><img src="images/ico_home.gif">&nbsp;企业空间</div>
						 <ul>
						<%If KS.ReturnPowerResult(0, "KSMS10008") Then%>
					  <li><a href="KS.EnterPrise.asp"  target="MainFrame" title="企业信息管理">企业空间管理</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMS10009") Then%>
					  <li><a href="KS.EnterPriseNews.asp"  target="MainFrame" title="企业新闻管理">企业新闻管理</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMS10010") Then%>
					  <li><a href="KS.EnterPrisePro.asp"  target="MainFrame" title="企业产品管理">企业产品管理</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMS10012") Then%>
					  <li><a href="KS.EnterPriseClass.asp"  target="MainFrame" title="行业分类管理">行业分类管理</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMS10013") Then%>
					  <li><a href="KS.EnterPriseAD.asp"  target="MainFrame" title="行业广告管理">行业广告管理</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMS10011") Then%>
					  <li><a href="KS.EnterPriseZS.asp"  target="MainFrame" title="荣誉证书管理">荣誉证书管理</a></li>
					 <%end if%>
						</ul>
					 </DIV>
					<%End If%>
                    <!--------------空间系统 end--------------------> 

					
						 
					
					</ul>
					
					
					
					
					<ul class="leftbox" id="dleft_tab2" style="display:none;">
					
					<%IF instr(lcase(KS.C("ModelPower")),"sysset10")=0 or KS.C("SuperTf")=1 Then%>
					   <div class="dt">系统设置</div>
					   <div class="dc">
					<%If KS.ReturnPowerResult(0, "KMST10001") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'系统设置 >> <font color=red>基本信息设置</font>','SetParam','KS.Setting.asp');" title="基本信息设置">基本信息设置</a></li>
					 <%end if%>
					      
						 <%If Not ModelXML Is Nothing Then
					       If ModelXML.documentElement.SelectNodes("row[@basictype=3 and @channelstatus=1]").length<>0 Then
						 %>
						  <%If KS.ReturnPowerResult(0, "KMST20001") Then %>
						  <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red><%=SQL(3,I)%>参数设置</font>','SetParam','KS.DownParam.asp?ChannelID=<%=SQL(0,I)%>');">下载参数设置</a></li>
						  <%End If%>
						 
						 <%If KS.ReturnPowerResult(0, "KMST20002") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>下载服务器管理</font>','Disabled','KS.DownServer.asp?ChannelID=<%=SQL(0,I)%>');">下载服务器管理</a>
						 <%end if%>
						 
					
						<%
						  End If
						End If
						
						If Not ModelXML Is Nothing Then
					       If ModelXML.documentElement.SelectNodes("row[@channelid=7 and @channelstatus=1]").length<>0 Then
						 %>
						  <%If KS.ReturnPowerResult(0, "KMST20003") and IsBusiness Then %>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>参数设置</font>','SetParam','KS.MovieParam.asp?ChannelID=7');">影视参数设置</a></li>
						  <%End If%>
						
						 <%If KS.ReturnPowerResult(0, "KMST20004") and IsBusiness Then %>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>影视服务器管理</font>','Disabled','KS.MediaServer.asp?TypeID=2&ChannelID=7');">影视服务器管理</a></li>
						 <%end if%>
					    <%
						   End If
						End If
						
						If Not ModelXML Is Nothing Then
					       If ModelXML.documentElement.SelectNodes("row[@channelid=8 and @channelstatus=1]").length<>0 Then
						 %>
						 <%If KS.ReturnPowerResult(0, "KMST20005") and IsBusiness Then %>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>供求交易类型管理</font>','Disabled','KS.SupplyType.asp');">供求交易类型管理</a></li>
						  <%End If%>
					  <%  End If
					   End If
					   %>
					 
					 <%If KS.ReturnPowerResult(0, "KMST10003") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'系统设置 >> <font color=red>在线支付平台管理</font>','SetParam','KS.PaymentPlat.asp');"  title="整合系统设置">在线支付平台管理</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KMST10002") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'系统设置 >> <font color=red>整合系统设置</font>','SetParam','KS.API.asp');"  title="整合系统设置">API通用整合设置</a></li>
					 <%end if%>
					   </div>
				<%End If%>
					   
				<%IF instr(lcase(KS.C("ModelPower")),"sysset20")=0 or KS.C("SuperTf")=1 Then%>	
				   
					<%
					If KS.CheckDir("../3G/") Then
					If KS.ReturnPowerResult(0, "KSO10000")  Then %>
					  <div class="dt">3G版参数配置</div>
					  <div class="dc">
                       <li><a href="#" onClick="SelectObjItem1(this,'3G版系统管理 >> <font color=red>3G版基本参数设置</font>','SetParam','../3g/Setting.asp');" title="3G版基本参数设置">3G版基本参数设置</a></li>
					   <li><a href="#"  onClick="SelectObjItem1(this,'3G版系统管理 >> <font color=red>3G版自定义页面管理</font>','Disabled','../3g/setting.asp?action=template');">3G版自定义页面</a></li>
					  </div>
					<%end if
					Else
					 Check3G
					End If
					%>
					   
					   <div class="dt">辅助管理</div>
					   <div class="dc">
						 <%If KS.ReturnPowerResult(0, "KMST10015") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'辅助管理 >> <font color=red>来源管理</font>','Disabled','KS.Origin.asp');">来源管理</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(0, "KMST10016") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'辅助管理 >> <font color=red>作者管理</font>','Disabled','KS.Author.asp?ChannelID=0');">作者管理</a> </li>
						 <%end if%>

						 <%If KS.ReturnPowerResult(0, "KMST10017") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'辅助管理 >> <font color=red>省市管理</font>','Disabled','KS.Province.asp');">地区管理</a> </li>
						 <%end if%>

					  <%If KS.ReturnPowerResult(0, "KMST10004") Then%>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'辅助管理 >> <font color=red>内容关键字设置</font>','Disabled','KS.InnerLink.asp');">内容关键字设置</a></li>
                      <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMST10019") Then%>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'辅助管理 >> <font color=red>搜索关键词维护</font>','Disabled','KS.KeyWord.asp?issearch=1');">搜索关键词维护</a></li>
                      <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMST10020") Then%>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'辅助管理 >> <font color=red>定时任务管理</font>','Disabled','KS.Task.asp?action=manage');">定时任务管理</a></li>
                      <%end if%>
					  
                       <%If KS.ReturnPowerResult(0, "KMST10014") Then %>
					     <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>投票记录管理</font>','Disabled','KS.PhotoVote.asp?ChannelID=<%=SQL(0,I)%>');">图片投票记录管理</a>	</li>
					   <%End If%>

					  <%If KS.ReturnPowerResult(0, "KMST10006") Then%>
					   	<li><a href="KS.Log.asp"  target="MainFrame" title="站点文件管理">后台日志管理</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMST10007") Then%>
					   <li><a href="KS.Database.asp?Action=BackUp"  target="MainFrame" title="数据库维护">数据库维护</a></li>
					   <%end if%>
					   <%If KS.ReturnPowerResult(0, "KMST10008") Then%>
					   <li><a href="KS.DataReplace.asp"  target="MainFrame" title="数据库字段替换">数据库字段替换</a></li>
					   <%end if%>
                       <%If KS.ReturnPowerResult(0, "KMST10018") Then%>
					   <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'辅助管理 >> <font color=red>上传文件管理</font>','Disabled','KS.AdminFiles.asp');">上传文件管理</a></li>
					   <%end if%>					   
					   <%If KS.ReturnPowerResult(0, "KMST10009") Then%>
					   <li><a href="KS.Database.asp?Action=ExecSql"  target="MainFrame" title="在线执行SQL语句">在线执行SQL语句</a></li>
					   <%end if%>
					   <%If KS.ReturnPowerResult(0, "KMST10011") Then%>
					   <li><a href="KS.Setting.asp?Action=CopyRight"  target="MainFrame" title="服务器参数探测">服务器参数探测</a></li>
					   <%end if%>
					   <%If KS.ReturnPowerResult(0, "KMST10012") Then%>
					   <li><a href="KS.CheckMM.asp"  target="MainFrame" title="在线检测木马">在线检测木马</a></li>
					   <%end if%>
					   </div>
				<%end if%>
					</ul>
					
					
					
					<ul class="leftbox" id="dleft_tab3" style="display:none;">
					<%If KS.ReturnPowerResult(0, "KSMS10006") Then %>
					<div class="dt">自定义表单</div>
					  <div class="dc">
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'自定义表单 >> <font color=red>表单项目管理</font>','Disabled','KS.Form.asp');">自定义表单管理</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'自定义表单 >> <font color=red>添加表单项目</font>','GO','KS.Form.asp?action=Add');">添加表单项目</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'自定义表单 >> <font color=red>表单项目调用代码</font>','Disabled','KS.Form.asp?action=total');">表单项目调用代码</a></li>
					  </div>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20014") Then%>
					<div class="dt">PK系统</div>
					  <div class="dc">
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'观点PK系统 >> <font color=red>PK主题管理</font>','Disabled','KS.PKZT.asp');">PK主题管理</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'观点PK系统 >> <font color=red>PK用户观点管理</font>','Disabled','KS.PKGD.asp');">PK用户观点管理</a></li>
					  </div>
					<%end if%>
					
                    <div class="dt">
					其它系统
					</div>
					<div class="dc">
						 <%If KS.ReturnPowerResult(0, "KSMS20010") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'其它系统 >> <font color=red>积分兑换商品</font>','Disabled','KS.MallScore.asp');">积分兑换商品</a></li>
						 <%End If%>
					<%If KS.ReturnPowerResult(0, "KSMS20009") Then %>
					<li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'其它系统 >> <font color=red>Digg管理</font>','Disabled','KS.DiggList.asp');">文档Digg管理</a></li>
					<%End If%>
					<%If KS.ReturnPowerResult(0, "KSMS20008") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'其它系统 >> <font color=red>心情指数管理</font>','Disabled','KS.Mood.asp');">心情指数管理</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'其它系统 >> <font color=red>点评系统管理</font>','Disabled','KS.Mood.asp?TypeFlag=1');">点评系统管理</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20001") Then%>
					<li><a href="KS.FriendLink.asp"  target="MainFrame" title="友情链接管理">友情链接管理</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20002") Then%>
					<li><a href="KS.Announce.asp"  target="MainFrame" title="网站公告管理">网站公告管理</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20000") Then%>
					<li><a href="KS.FeedBack.asp"  target="MainFrame" title="投诉及反馈管理">投诉及反馈管理</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20003") Then%>
					<li><a href="KS.Vote.asp"  target="MainFrame" title="站内调查管理">站内调查管理</a></li>
					<%end if%>
					
					<%If KS.ReturnPowerResult(0, "KSMS20005") Then%>
					<li><a href="KS.Online.asp"  target="MainFrame" title="站点计数器管理">站点计 数 器</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20006") Then%>
					<li><a href="KS.Ads.asp"  target="MainFrame" title="广告系统管理">广告系统管理</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20007") Then%>
					<li><a href="KS.PromotedPlan.asp"  target="MainFrame" title="推广计划管理">推广计划管理</a></li>
					<%end if%>
					</div>
					</ul>
					
					
					
										
					<ul class="leftbox" id="dleft_tab4" style="display:none">
					<div class="dt">模型管理</div>
					 <div class="dc">
					 <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'模型管理 >> <font color=red>模型管理首页</font>','Disabled','KS.Model.asp');">模型管理首页</a></li>
					 <%If KS.ReturnPowerResult(0, "KSMM10000") Then%>
					 <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'模型管理 >> <font color=red>添加新模型</font>','Go','KS.Model.asp?action=Add');">添加新模型</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMM10004") Then%>
					 <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'模型管理 >> <font color=red>模型信息统计</font>','Go','KS.Model.asp?action=total');">模型信息统计</a></li>
					 <%end if%>
					 </div>
					 <%If KS.ReturnPowerResult(0, "KSMM10003") Then%>
					<div class="dt">模型字段管理</div>
					 <div class="dc">
					  <%For I=0 To UBound(SQL,2)
					   if SQL(6,I)=1 AND SQL(0,I)<>6 and SQL(0,I)<>9 and SQL(0,I)<>10 Then
					  %>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'模型管理 >> <font color=red>字段管理</font>','Disabled','KS.Field.asp?ChannelID=<%=SQL(0,I)%>',<%=SQL(0,I)%>);"><%=SQL(1,I)%>字段</a></li>					  
					  <%
					  End iF
					 Next%>
					</div>
					<div class="dt">管理列表菜单</div>
					 <div class="dc">
					  <%For I=0 To UBound(SQL,2)
					   if SQL(6,I)=1 AND SQL(0,I)<>6 and SQL(0,I)<>9 and SQL(0,I)<>10 Then
					  %>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'模型管理 >> <font color=red>管理列表管理</font>','Disabled','KS.Model.asp?action=ManageMenu&ChannelID=<%=SQL(0,I)%>',<%=SQL(0,I)%>);"><%=SQL(1,I)%>列表菜单</a></li>					  
					  <%
					  End iF
					 Next%>
					</div>
					<%end if%>
					</ul>

                    <ul class="leftbox" id="dleft_tab5" style="display:none">
					 <div class="dt">标签管理</div>
					 <div class="dc">
					<%
					If KS.ReturnPowerResult(0, "KMTL10001") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>系统函数标签</font>','FunctionLabel','Include/Label_Main.asp?LabelType=0');"">系统函数标签</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10002") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>自定义SQL函数标签</font>','DiyFunctionLabel','Include/Label_Main.asp?LabelType=5');"">自定义SQL函数标签</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10003") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>自定义静态标签</font>','FreeLabel','Include/Label_Main.asp?LabelType=1');"">自定义静态标签</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10010") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>通用循环标签</font>','FreeLabel','Include/Label_Main.asp?LabelType=6');"">通用循环列表标签</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10011") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>自定义生成XML文档</font>','DiyFunctionLabel','Include/Label_Main.asp?LabelType=7');"">自定义生成XML文档</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10004") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>自定义JS管理</font>','SysJSList','include/JS_Main.asp?JSType=0');"">系统JS管理</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10005") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>自定义JS管理</font>','FreeJSList','include/JS_Main.asp?JSType=1');"">自定义JS管理</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMSL10008") Then
					  .Write "<li><a href='KS.ClassMenu.asp'  target='MainFrame' title='生成顶部菜单'>生成顶部菜单</a></li>"
					end if
					If KS.ReturnPowerResult(0, "KMSL10009") Then
					  .Write "<li><a href='KS.TreeMenu.asp'  target='MainFrame' title='生成树形菜单'>生成树形菜单</a></li>"
					End If

		              .write "</div>"
					  .write "<div class='dt'>模板管理</div>"
					  .write "<div class='dc'>"
					If KS.ReturnPowerResult(0, "KMTL10006") Then
						.Write ("<li id='s_1'><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'模板标签管理 >> <font color=red>自定义页面管理</font>','Disabled','KS.DIYPage.asp');"">自定义页面管理</a></li>")
				    End If
					If KS.ReturnPowerResult(0, "KMTL10007") Then
						.Write ("<li id='s_1'><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'模板标签管理 >> <font color=red>所有模板管理</font>','Disabled','KS.Template.asp');"">所有模板管理</a></li>")
					End If
					 %>	
					 </div>
					</ul>
					
					<ul class="leftbox" id="dleft_tab6" style="display:none">
					
					  <div class="dt">
					   用户管理					  </div>
					  <div class="dc">
					  <%If KS.ReturnPowerResult(0, "KMUA10001") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>管理员管理</font>','Disabled','KS.Admin.asp');">管理员管理</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10002") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>注册用户管理</font>','Disabled','KS.User.asp');" title="注册用户管理">注册用户管理</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>添加用户</font>','Disabled','KS.User.asp?Action=Add');" title="添加用户">添加用户</a></li>
					  
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10016") and IsBusiness Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>实名认证管理</font>','Disabled','KS.UserRZ.asp');" title="实名认证管理" style="color:red">实名认证管理</a></li>
					  <%end if%>
					  
					  <%If KS.ReturnPowerResult(0, "KMUA10004") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>用户组管理</font>','Disabled','KS.UserGroup.asp');" title="用户组管理">用户组管理</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10003") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>用户短信管理</font>','Disabled','KS.UserMessage.asp');" title="用户短信管理">用户短信管理</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10009") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>发送邮件管理</font>','Disabled','KS.UserMail.asp');" title="发送邮件管理">发送邮件管理</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10012") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员字段管理</font>','Disabled','KS.Field.asp?ChannelID=101');" title="会员字段管理">会员字段管理</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10013") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员表单管理</font>','Disabled','KS.UserForm.asp');" title="会员表单管理">会员表单管理</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10015") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员使用记录</font>','Disabled','KS.UserUseLog.asp');" title="会员使用记录">会员使用记录</a></li>
					  <%end if%>
					  
					  
					  </div>
					  <div class="dt">
					   账务明细管理					 
					  </div>
					  <div class="dc">
					  <%If KS.ReturnPowerResult(0, "KMUA10005") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员点券明细</font>','Disabled','KS.LogPoint.asp');" title="会员点券明细">会员点券明细</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10006") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员有效期明细</font>','Disabled','KS.LogEdays.asp');" title="会员有效期明细">会员有效期明细</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10007") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员资金明细</font>','Disabled','KS.LogMoney.asp');" title="会员资金明细">会员资金明细</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员积分明细</font>','Disabled','KS.LogScore.asp');" title="会员积分明细">会员积分明细</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10008") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>线下充值卡管理</font>','Disabled','KS.Card.asp?cardtype=0');" title="线下充值卡管理">线下充值卡管理</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>线上充值卡管理</font>','Disabled','KS.Card.asp?cardtype=1');" title="线上充值卡管理">线上充值卡管理</a></li>
					  <%end if%>
					  </div>
					  <%If KS.ReturnPowerResult(0, "KMUA10002") Then %>
					  <div class="dt">
					   快速查找用户					  </div>
					  <div class="dc">
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','KS.User.asp?UserSearch=5');"><font color=#ff6600>24小时内登录</a></font></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','KS.User.asp?UserSearch=6');">24小时内注册</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','KS.User.asp?UserSearch=1');"> 被锁住的用户</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','KS.User.asp?UserSearch=3');">待审批会员</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','KS.User.asp?UserSearch=4');">待邮件验证</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','KS.User.asp?UserSearch=2');">所有管理员用户</a></li>
                      </div>
					<%end if%>
					</ul>
					<ul class="leftbox" id="dleft_tab7" style="display:none">

					<%If KS.ReturnPowerResult(0, "KSO10003") Then %>
					  <div class="dt">WSS统计插件</div>
					  <div class="dc">
					   <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'WSS 统计插件 >> <font color=red>WSS 设置</font>','Disabled','../plus/wss/wss.asp');">WSS 设置</a></li>
					   <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'WSS 统计插件 >> <font color=red>WSS 设置</font>','Disabled','../plus/wss/wss.asp?action=show');">查看统计</a></li>
					  </div>
				<%end if%>
				 <%If KS.ReturnPowerResult(0, "KSO10002") Then %>
					  <div class="dt">bShare分享插件 </div>
					  <div class="dc">
					   <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'bShare分享插件 >> <font color=red>bShare分享插件设置</font>','Disabled','../plus/bshare/bshare.asp');">bShare插件设置</a></li>
					   <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'bShare分享插件 >> <font color=red>bShare分享数据统计</font>','Disabled','../plus/bshare/bshare.asp?action=getdata');">bShare数据统计</a></li>
					  </div>
					<%end if%>
				 <%If KS.ReturnPowerResult(0, "KSO10001") Or KS.ReturnPowerResult(0, "KSO10004") Then %>
					  <div class="dt">辅助管理插件</div>
					  <div class="dc">
					   <%If KS.ReturnPowerResult(0, "KSO10001") Then %>
					   <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'数据导入插件 >> <font color=red>数据批量导入管理</font>','Disabled','KS.Import.asp');">数据批量导入管理</a></li>
					   <%End If%>
					   <%If KS.ReturnPowerResult(0, "KSO10004") Then %>
					   <li><a href="KS.Tools.asp"  target="MainFrame" title="一键管理工具">一键管理工具</a></li>
					   <%End If%>
					  </div>
				<%end if%>
				    </ul>
				  </div>
					
					<div></div>
			</div><!--menubox-->			</td>
		  </tr>
		</table>
		<SCRIPT type="text/javascript">
		function fHideFocus(tName){
		aTag=document.getElementsByTagName(tName);
		for(i=0;i<aTag.length;i++)aTag[i].onfocus=function(){this.blur();};
		}
		fHideFocus("A");
		var id = getCookie("cltips");  //默认选中的ID
		if (id==''){id=1};
		document.getElementById("subTable"+id).className = "show";
		document.getElementById("td_"+id).className = "left_menu_selected";
		var cache_id = id;
		function switchShow(id,tag){
		    setCookie("cltips",id);
		    document.getElementById("td_"+id).className='left_menu_selected';
			for(var i=1; i<=<%=n%>; i++){
			   if (i!=id)
				document.getElementById("td_"+i).className='left_menu';
		     }
			var tObj = document.getElementById("subTable"+id);
			var	cObj = document.getElementById("subTable"+cache_id);
			if(tag){
				if(tObj) tObj.className =(tObj.className=='hid') ? "show" : "hid";
			}else{
				if(tObj) tObj.className = "show";
			}
			if(cache_id != id){
				cache_id = id;
				if(cObj)cObj.className = "hid";
			}
			event.cancelBubble = true;
		}
		function showleft(id)
		{ 
		 document.getElementById("left_tab"+id).className='Selected';
		 var oItem = document.getElementById("TabPage").getElementsByTagName("li"); 
			for(var i=1; i<=oItem.length; i++){if (i!=id){document.getElementById("left_tab"+i).className='';} }
			var dvs=$(".leftbox");
			for (var i=0;i<dvs.length;i++){if (dvs[i].id==('dleft_tab'+id)){$("#"+dvs[i].id).show('fast');}else{$("#"+dvs[i].id).hide('fast');}}
		}
		</SCRIPT>
<%
        If Session("ShowCount")="" Then
		.Write " <ifr" & "ame src=""http://ww" &"w.k" &"e" & "si" &"on." & "co" & "m" & "/WebS" & "ystem/Co" & "unt.asp"" scrolling='no' frameborder='0' height='0' width='0'></iframe>"
		Session("ShowCount")=KS.C("AdminName")
		End If
	    End With
		End Sub
		Function bytes2BSTR(vIn)
		Dim i,ThisCharCode,NextCharCode
		Dim strReturn:strReturn = ""
		For i = 1 To LenB(vIn)
			ThisCharCode = AscB(MidB(vIn,i,1))
			If ThisCharCode < &H80 Then
				strReturn = strReturn & Chr(ThisCharCode)
			Else
				NextCharCode = AscB(MidB(vIn,i+1,1))
				strReturn = strReturn & Chr(CLng(ThisCharCode) * &H100 + CInt(NextCharCode))
				i = i + 1
			End If
		Next
		bytes2BSTR = strReturn
		End Function
		Function getfile(RemoteFileUrl)
		On Error Resume Next 
		Dim Retrieval:Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
		With Retrieval
		 .Open "Get", RemoteFileUrl, false, "", ""
		 .Send
		 If .Readystate<>4 then Exit Function
		 getfile =bytes2BSTR(.responseBody)
		End With
		If Err Then
		Err.clear
		getfile="error!"
		End if
		Set Retrieval = Nothing
		end function

		Function GetTrueDomain(domain)
				Dim x:x = split(domain,".")
				Dim sdomain:sdomain= ""
				Dim start:start = 2
				Dim k :k= 1
				if ubound(x)<=1 then GetTrueDomain=domain:exit function
				if (ubound(x) >= 3) then start = 3
				dim i:i=start
				do while i > 0
					if (i=start) then
						sdomain = sdomain & x(ubound(x)-start+k)
					else
						sdomain = sdomain & "." & x(ubound(x)-start+k)
					end if
					k=k+1
					i=i-1
				loop
				GetTrueDomain=sdomain
		end function

		Sub GetRemoteVer()         response.write getfile()
		End Sub
		Sub CopyRight()
		  If Request.ServerVariables("SERVER_NAME")="127.0.0.1" or Request.ServerVariables("SERVER_NAME")="localhost" Then
		  Else
			  If KS.IsNul(Session("CheckCopyRight")) Then
			   Session("CheckCopyRight")=getfile()
			  End If
			   
			  If Not KS.IsNul(Session("CheckCopyRight")) and Session("CheckCopyRight")<>"error!" Then
				If Session("CheckCopyRight")="true" Then
				  KS.Echo escape("")
				Else
				 If IsBusiness=False Then
				  KS.Echo escape("")
				 Else
				  KS.Echo escape("")
				 End If
				End If
			  End If
		  End If
		End Sub

  Public Sub KS_Main()
		   
		   Dim TipStr,SafetyTips:SafetyTips=KS.ReadSetting(0)
		 
		   If SafetyTips="1" Then
			   If EnableSiteManageCode=false Then
				TipStr="<li style=""height:24px;line-height:24px"">您没有启用管理认证码，建议您打开conn.asp将EnableSiteManageCode的值设置为True；</li>"
			   ElseIf SiteManageCode="8888"  Then
				TipStr=TipStr & "<li style=""height:24px;line-height:24px""><img src=""images/gif-0760.gif"" align=""absmiddle""> 您后台管理认证密码为系统默认值：<span style=""color:red"">8888</span>,建议您及时打开conn.asp里修改；</li>"
			   End If
			   If KS.CheckDir("../admin") Then
		       TipStr=TipStr & "<li style=""height:24px;line-height:24px""><img src=""images/gif-0760.gif"" align=""absmiddle""> 您的网站后台管理目录为：<span style=""color:red"">admin </span>，出于安全的考虑，我们建议您修改目录名；</li>"
			   End If
			   
			   If DataBaseType=0 Then
			    If instr(lcase(DBPath),"ks_data/kesioncms9.mdb")<>0 then
		         TipStr=TipStr & "<li style=""height:24px;line-height:24px""><img src=""images/gif-0760.gif"" align=""absmiddle""> 您的数据库名称为系统默认名称：<span style=""color:red"">" & DBPath & "</span>,出于安全的考虑，我们建议您修改数据库名称；</li>"
				end if
			   End If
			   
			   If Lcase(KS.C("AdminPass"))="469e80d32c0559f8" Then
		         TipStr=TipStr & "<li style=""height:24px;line-height:24px""><img src=""images/gif-0760.gif"" align=""absmiddle""> 您的后台管理员密码为系统默认值：<span style=""color:red"">admin888</span>,出于安全的考虑，我们建议您及时修改后台登录密码；</li>"
			   End If
			   
			   If TipStr<>"" Then
		    TipStr=TipStr & "<div style=""margin-top:16px;margin-bottom:20px;text-align:right""><label style=""color:#999""><input onclick=""parent.setCookieTips(this.checked)""  type=""checkbox"" name=""nottips"" id=""notips"" value=""1"">我知道了，下次进入后台不再提醒</label></div>"
			   End If
		   End If
		   
		   %>
           <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
			<html xmlns="http://www.w3.org/1999/xhtml">
			<head>
			<script src="../ks_inc/jquery.js" type="text/javascript"></script>
			<script src="../ks_inc/lhgdialog.js"></script>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
			<style type="text/css">
			a{color:#555;}
			.position{ border-bottom:1px #83B5CD solid;background:url(images/titlebg.png); height:36px; font-size:13px; color:#555;line-height:36px; padding-left:10px;}
			.title{ background:#FBFDFF;border-top:2px solid #E1EEFF; line-height:28px; font-weight:bold;height:28px; color:#555; margin-left:20px;margin-right:20px;text-decoration:none;font-size:14px; margin-top:10px; padding-left:10px; padding-top:8px;}
			.title img{ padding-top:5px; padding-right:6px;}
			
			.nr{ height:auto; color:#555; text-decoration:none;font-size:12px; line-height:22px; padding-left:10px;margin-left:20px;margin-right:20px;}
			.nr ul{ padding:0px;margin:0px;}
			.nr li{text-alilgn:left;list-style-type:none;}
			.l {width:350px;float:left}
			.r {width:380px;float:right}
			.l h2,.r h2{font-size:13px;color:#ff6600}
			.box{clear:both}
			.newbox1{float:left;width:49%;}
			.newbox2{float:right;width:50%;}
			
			<%
			If Instr(KS.Setting(16),"2")=0 Then
			 KS.Echo ".bbs{display:none}"
			End If
			%>
			.bbs li{list-style-image:url(images/37.gif)}
			</style>
			<script type="text/javascript">
			function showbigpic(){
				var box=$.dialog({title:'查看软件登记证书：',content: '<style>.zs{width:460px;}.zs li img{border:1px solid #000;margin:5px;width:199px;height:220px;}.zs li{width:200px;float:left;margin:10px;}</style><strong>软件版权证书：</strong><br/><div class="zs"><ul><li><a href="#cmsv9zs.jpg" target="_blank"><img src="#cmsv9zs.jpg"/></a></li><li><a href="#eshopv9zs.jpg" target="_blank"><img src="#eshopv9zs.jpg"/></a></li><li><a href="#examv9zs.jpg" target="_blank"><img src="#examv9zs.jpg"/></a></li><li><a href="#v5dj.jpg" target="_blank"><img src="#v5dj.jpg"/></a></li></ul></div>',max:false,min: false});
			}
			$(document).ready(function(){
			  <%If SafetyTips="1" and TipStr<>"" Then%>
			   var p=new parent.KesionPopup();
               p.FadeInTime=800;
			   p.FadeOutTime=800;
			   p.popupTips('<span style="font-weight:bold;font-size:16px"><img align="absmiddle" src="images/ico/back.gif">安全提醒</span>','<div style="font-size:12px;height:160px;"><br/><ul><%=TipStr%></ul></div>',600,400);
			  <%End If%>
			 <%if CheckNewVersion Then%>
			  $.get('index.asp',{action:'ver'},function(d){$('#versioninfo').html(d);});
			  <%if request.ServerVariables("SERVER_NAME")<>"localhost" and request.ServerVariables("SERVER_NAME")<>"127.0.0.1" then%>
			  $.get('index.asp',{timestamp:new Date().getTime(),action:'copyright'},function(d){$('#currversion').html(unescape(d))});
			 <%End If%>
			  //检查是否存在升级文件
			  $.ajax({
			  url: "KS.Update.asp",
			  cache: false,
			  data: "action=check",
			  success: function(d){
			        d=unescape(d);
					switch (d){
					 case 'enabled':
					  $("#updateInfo").html("<font color='green'>对不起,您没有开启自动检测最新版本功能!</font>");
					  break;
					 case 'false':
					  $("#updateInfo").html("<font color='green'>当前已经是最新版本!</font>");
					  break;
					 case 'localversionerr':
					  $("#updateInfo").html("<font color='green'>加载本地xml版本文件出错,请检查<%=KS.Setting(89)%>include/version.xml文件是否存在!</font>");
					  break;
					 case 'remoteversionerr':
					  $("#updateInfo").html("<font color='green'>读取服务器文件出错,请检查<%=KS.Setting(89)%>ks.update.asp文件的配置是否正确或稍候再试!</font>");
					  break;
					 case 'unallow':					  $("#updateInfo").html("<font color='green'>系统检查到有可更新文件,但不支持在线升级,请到官方站(<a href='#' target='_blank'>#</a>)下载升级文件!</font>");
					  break;
					 case 'unallowversion':
					  $("#updateInfo").html("<font color='green'>系统检查到有可更新文件,但由于您的版本号与最新版本号不对应,不支持在线升级,请根据您当前使用的版本到官方站(<a href='#' target='_blank'>#</a>)下载升级文件手工升级!</font>");
					  break;
					 default:
					    $("#updateInfo").html("<font color='red'>系统检查到有可升级文件!</font>");
						new parent.KesionPopup().PopupCenterIframe('<img align="absmiddle" src="images/ico/back.gif">KesionCMS 升级提醒','KS.update.asp?action=showupdateinfo',700,350,'auto')
					  break;
					}
			  }
		 	 });
			  <%End If%>
			 });
           </script>
			</head>
			
			<body scroll=no>
<style type="text/css">
<!--
.admin_nav{ padding:15px;}
.admin_nav a{ float:left;background:#CEE7FF; display:block; width:33%; height:35px; padding:0; text-align:center; line-height:35px; color:#003366; font-size:16px; border-bottom:1px solid #fff; border-right:1px solid #fff; text-decoration:none; font-family:Verdana, Arial, Helvetica, sans-serif; overflow:hidden;}
.admin_nav span{background:#84C1FF; display:block; width:99%; height:35px; padding:0; text-align:center; line-height:35px; color:#003366; font-size:16px; border-bottom:1px solid #fff; border-right:1px solid #fff; text-decoration:none; font-family:Verdana, Arial, Helvetica, sans-serif; overflow:hidden;}
.admin_nav span em{ color:#f00; font-style:normal; background:#FF9;}

-->
</style>
<div class="admin_nav">
<span>欢迎来到管理后台，请谨慎操作！</span>
<a href="KS.ItemInfo.asp?ChannelID=113&ID=20137865226380">网站图标管理 &gt;&gt;</a>
<a href="KS.ItemInfo.asp?ChannelID=115&ID=20137700396941">企业信息设置 &gt;&gt;</a>
<a href="KS.Form.asp?ItemID=5&action=resulthp">在线留言管理 &gt;&gt;</a>

</div>
			</body>
			</html>

          <%
				Conn.Close:Set Conn = Nothing
			End Sub
			
			Public Sub KS_Foot()
		     With Response
				.Write "<div id='foot'>"
				.Write "<div id='co' align=""center"" onClick=""ChangeLeftFrameStatu();"" title=""全屏/半屏"" style=""cursor:pointer;""><font color=red>×</font> 关闭左栏</div>"
				.Write "<div id='footmenu'><span style='float:left'>快速通道=>：</span>"
				.Write "<a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'系统管理中心 >> <font color=red>首页</font>','disabled','index.asp?action=Main');"">后台首页</a>"

				If KS.ReturnPowerResult(0, "KMTL20000") Then
				.Write "<a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'发布中心 >> <font color=red>发布管理首页</font>','disabled','Include/refreshindex.asp');"">发布首页</a>"
				End If
				.Write "<a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'发布中心 >> <font color=red>发布管理首页</font>','disabled','Include/RefreshHtml.asp?ChannelID=1');"">发布管理</a>"
				
				If KS.ReturnPowerResult(0, "KMTL10007") Then
				.Write "<a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'模板标签管理 >> <font color=red>模板管理</font>','disabled','KS.Template.asp');"">模板管理</a>"
				End If
				If KS.ReturnPowerResult(0, "KMST10001") Then
				.Write "<a href='javascript:void(0)' onClick=""SelectObjItem1(this,'系统设置 >> <font color=red>基本信息设置</font>','SetParam','KS.Setting.asp');"" title='基本信息设置'>基本信息设置</a>"
				End If
				If Instr(KS.C("ModelPower"),"model1")>0 Or KS.C("SuperTF")="1" then
				.Write "<a href='javascript:void(0)' onClick=""SelectObjItem1(this,'模型管理 >> <font color=red>模型管理首页</font>','SetParam','KS.Model.asp');"">模型管理</a>"
				End If
				If KS.ReturnPowerResult(0, "KMUA10011") Then
				.Write "<a href='javascript:void(0)' onClick=""SelectObjItem1(this,'用户管理 >> <font color=red>稿件统计</font>','SetParam','KS.UserProgress.asp');"">稿件统计</a>"
			    End If
				.Write "</div>"
				.Write "<div id='footcopyright'>欢迎进入智能管理系统，感谢您的支持！</div>"
				.Write "</div>"
				.Write "<SCRIPT language=javascript>"
				.Write "    var screen=false;"
				.Write "    function ChangeLeftFrameStatu()"
				.Write "    {"
				.Write "        if(screen==false)"
				.Write "        {"
				.Write "            $('#leftframe').hide();"
				.Write "            screen=true;"
				.Write "            self.co.innerHTML = ""√ 打开左栏"""
				.Write "        }"
				.Write "        else if(screen==true)"
				.Write "        {"
				.Write "            $('#leftframe').show();"
				.Write "           screen=false;"
				.Write "            self.co.innerHTML = ""<font color=red>×</font> 关闭左栏"""
				.Write "        }"
				.Write "    }"
				.Write "</SCRIPT>"
			End With
		End Sub
		Sub CheckSetting()
			 dim strDir,strAdminDir,InstallDir
			 strDir=Trim(request.ServerVariables("SCRIPT_NAME"))
			 strAdminDir=split(strDir,"/")(Ubound(split(strDir,"/"))-1) & "/"
			 InstallDir=left(strDir,instr(lcase(strDir),"/"&Lcase(strAdminDir)))
					
			If Instr(UCASE(InstallDir),"/W3SVC")<>0 Then
			   InstallDir=Left(InstallDir,Instr(InstallDir,"/W3SVC"))
			End If
		 If KS.Setting(2)<>KS.GetAutoDoMain or KS.Setting(3)<>InstallDir Then
			
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select Setting From KS_Config",conn,1,3
		  Dim SetArr,SetStr,I
		  SetArr=Split(RS(0),"^%^")
		  For I=0 To Ubound(SetArr)
		   If I=0 Then 
			SetStr=SetArr(0)
		   ElseIf I=2 Then
			SetStr=SetStr & "^%^" & KS.GetAutoDomain
		   ElseIf I=3 Then
			SetStr=SetStr & "^%^" & InstallDir
		   Else
			SetStr=SetStr & "^%^" & SetArr(I)
		   End If
		  Next
		  RS(0)=SetStr
		  RS.Update
		  RS.Close:Set RS=Nothing
		  Call KS.DelCahe(KS.SiteSn & "_Config")
		  Call KS.DelCahe(KS.SiteSn & "_Date")
		 End If
		End Sub
		
		Sub Check3G()
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select WapSetting From KS_Config",conn,1,3
		  Dim SetArr,SetStr,I
		  SetArr=Split(RS(0),"^%^")
		  For I=0 To Ubound(SetArr)
		   If I=0 Then 
			SetStr=0
		   Else
			SetStr=SetStr & "^%^" & SetArr(I)
		   End If
		  Next
		  RS(0)=SetStr
		  RS.Update
		  RS.Close:Set RS=Nothing
		  Call KS.DelCahe(KS.SiteSn & "_Config")
		  Call KS.DelCahe(KS.SiteSn & "_Date")
		End Sub

End Class
%> 
