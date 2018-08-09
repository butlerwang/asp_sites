<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<!--#include file="../KS_Cls/Kesion.KeyCls.asp"-->
<%
	response.cachecontrol="no-cache"
	response.addHeader "pragma","no-cache"
	response.expires=-1
	response.expiresAbsolute=now-1
	Response.CharSet="utf-8"

Dim KSCls
Set KSCls = New Admin_Down
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Down
        Private KS,KSCls
		'=====================================定义本页面全局变量=====================================
		Private ID, I, totalPut, Page, RS,ComeFrom,TempStr
		Private KeyWord, SearchType, StartDate, EndDate, SearchParam,MaxPerPage,SpecialID
		Private T, TitleStr,Subtitle, AttributeStr
		Private FolderID, TemplateID,WapTemplateID,Action,FileName
		Private DownID, Title, DownVerSion, PhotoUrl, BigPhoto,DownContent, DownUrls, Recommend,IsTop
		Private Popular, Verific, Comment,Slide,Rolls,Strip, ChangesUrl, KeyWords, Author, Origin, AddDate, Rank, Hits, HitsByDay, HitsByWeek, HitsByMonth
		Private CurrPath, InstallDir, UpPowerFlag,Inputer
		Private DownLb, DownYY, DownSQ, DownPT, DownSize, SizeUnit, YSDZ, ZCDZ, JYMM
		Private ComeUrl,SqlStr,Errmsg,Makehtml,Tid,Fname,KSRObj
		Private ReadPoint,ChargeType,PitchTime,ReadTimes,InfoPurview,arrGroupID,DividePercent
		Private ChannelID,FieldXML,FieldNode,FNode,FieldDictionary
		Private SEOTitle,SEOKeyWord,SEODescript
		'=============================================================================================
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		Public Sub Kesion()
		ChannelID=KS.ChkClng(KS.G("ChannelID"))
		Session("ThumbnailsConfig")=KS.C_S(ChannelID,46)
		Call KSCls.LoadModelField(ChannelID,FieldXML,FieldNode)

		'收集搜索参数
		KeyWord = KS.G("KeyWord")
		SearchType = KS.G("SearchType")
		StartDate = KS.G("StartDate")
		EndDate = KS.G("EndDate")
		ComeFrom   = KS.G("ComeFrom")
		SearchParam = "ChannelID=" & ChannelID
		If KeyWord<>"" Then SearchParam=SearchParam & "&KeyWord=" & KeyWord
		If SearchType<>"" Then  SearchParam=SearchParam & "&SearchType=" & SearchType
		If StartDate<>"" Then SearchParam=SearchParam & "&StartDate=" & StartDate 
		If EndDate<>"" Then SearchParam=SearchParam & "&EndDate=" & EndDate
		If KS.S("Status")<>"" Then SearchParam=SearchParam & "&Status=" & KS.S("Status")
		If ComeFrom<>"" Then SearchParam=SearchParam & "&ComeFrom=" & ComeFrom
		
		ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))

			Action = Trim(KS.G("Action"))
			If action="geturls" then
			   dim folderpath:folderpath=KS.G("folderpath")
			   dim subfolder:subfolder=ks.chkclng(ks.g("subfolder"))
			   dim ext:ext=request("ext")
			   call geturls(folderpath,subfolder,ext)
			   exit sub
			ElseIf action="SelectFolder" Then
	           dim uploadDir:uploadDir=KS.Setting(3) & KS.Setting(91)
			   if request("path")<>"" then
			     uploadDir=request("path")
			   end if
			   call folderlist(uploadDir)
			   exit sub
			end if
			Page = KS.G("page")
				
			IF KS.G("Method")="Save" Then
				 Call DownSave()
			Else 
				 Call DownAdd()
			End If
	End Sub
	
	Sub geturls(byval folderpath,subfolder,ext)
	  if right(folderpath,1)<>"/" then folderpath=folderpath & "/"
	  dim FsoObj,FolderObj,FileObj,FsoItem,fext
	  Set FsoObj = KS.InitialObject(KS.Setting(99))
	  Set FolderObj = FsoObj.GetFolder(Server.MapPath(folderpath))
	  Set FileObj = FolderObj.Files
	  For Each FsoItem In FileObj
		  fext=Mid(FsoItem.name, InStrRev(FsoItem.name, ".")) '分离出扩展名
		  If KS.FoundInArr(ext,fext,"|") then
	      KS.Echo folderpath & FsoItem.name &vbcrlf
		  end if
	  Next
	  
	  '子目录
	  If subfolder=1 then
		  if FolderObj.subfolders.count>0 then
			For Each FsoItem In FolderObj.subfolders
			  call geturls (folderpath & fsoitem.name,subfolder,ext)
			Next
		  end if
	  end if
	End Sub
	
	sub folderlist(uploadDir)
	  dim ParentPath,tp
	  if uploadDir<>"" and uploadDir<>KS.Setting(3) & KS.Setting(91) then
	        if right(uploaddir,1)="/" then tp=left(uploaddir,len(uploaddir)-1) else tp=uploaddir
	        ParentPath = Mid(tp, 1, InStrRev(tp, "/") - 1)
	  end if
	  dim FsoObj,FolderObj,FileObj,FsoItem,fext
	  Set FsoObj = KS.InitialObject(KS.Setting(99))
	  Set FolderObj = FsoObj.GetFolder(Server.MapPath(uploadDir))
	  Set FileObj = FolderObj.Files
		ks.echo "<html>"
		ks.echo "<head>"
		ks.echo "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
		ks.echo "<title>添加</title>"
		ks.echo "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>" & vbCrlf
		ks.echo "<script language='JavaScript' src='../KS_Inc/common.js'></script>" & vbCrlf
		ks.echo "<script language='JavaScript' src='../KS_Inc/Jquery.js'></script>" & vbCrlf
	   ks.echo "<table border='0' width='100%'>"
	   if FolderObj.subfolders.count>0 then
			 if ParentPath<>"" then
			   ks.echo "<tr><td><img src='Images/arrow.gif' align='absmiddle' /><a href='KS.Down.asp?action=SelectFolder&channelid=" & channelid & "&path=" & ParentPath & "/'>回上一级</a></td></tr>"
			 end if
		  For Each FsoItem In FolderObj.subfolders
			ks.echo "<tr><td><img src='images/folder/folder.gif' align='absmiddle' /><a href='KS.Down.asp?action=SelectFolder&channelid=" & channelid & "&path=" & uploadDir & fsoitem.name & "/'>" & FsoItem.Name & "</a></td><td><label><input type='checkbox' onclick=""parent.returnFolder('" &uploadDir & fsoItem.Name & "/')"">选择</label></td></tr>"
		  Next
	   else
	     if ParentPath<>"" then
	       ks.echo "<tr><td><img src='Images/arrow.gif' align='absmiddle' /><a href='KS.Down.asp?action=SelectFolder&channelid=" & channelid & "&path=" & ParentPath & "/'>没有子目录了,点此返回</a></td></tr>"
		 end if
	   end if
	   ks.echo "</table>"
	end sub
	
	Sub DownAdd()
	  'On Error Resume Next
		With Response
		CurrPath = KS.GetUpFilesDir()
		
		Set RS = Server.CreateObject("ADODB.RecordSet")
		If Action = "Add" Then
		  FolderID = Trim(KS.G("FolderID"))
		  
		  If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10002") Then          '检查是否有添加下载的权限
		   .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "&channelid=" & channelid&"';</script>")
		   Call KS.ReturnErr(2, "KS.Down.asp?Page=" & Page & "&ID=" & FolderID&"&channelid=" & channelid)
		   Exit Sub
		  End If
		  Hits = 0:HitsByDay = 0:HitsByWeek = 0:HitsByMonth = 0:Comment = 1:IsTop=0:Strip=0
		  ReadPoint=0:PitchTime=24:ReadTimes=10
		  KeyWords = Session("keywords")
		  Author = Session("Author")
		  Origin = Session("Origin")
		  DownPT = "Win9x/NT/2000/XP":SizeUnit = "KB":YSDZ = "http://":ZCDZ = "http://"
		ElseIf Action = "Edit" Or Action="Verify" Then
		   Set RS = Server.CreateObject("ADODB.RECORDSET")
		   RS.Open "Select top 1 * From " & KS.C_S(ChannelID,2) & " Where ID=" & KS.ChkClng(KS.G("ID")), conn, 1, 1
		   If RS.EOF And RS.BOF Then
			Call KS.Alert("参数传递出错!", "KS.Down.asp")
			Set KS = Nothing
			Response.End
			Exit Sub
		   End If
			DownID = Trim(RS("ID"))
			FolderID = Trim(RS("Tid"))
			
			If Action = "Edit" And Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10003") Then     '检查是否有编辑下载的权限
			RS.Close
			Set RS = Nothing
			 If KeyWord = "" Then
			  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "&channelid=" &channelid & "';</script>")
			  Call KS.ReturnErr(1, "KS.Down.asp?ChannelID=" & channelid & "&Page=" & Page & "&ID=" & FolderID)
			 Else
			  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=" & server.URLEncode(KS.C_S(ChannelID,1)&" >> <font color=red>搜索" & ks.c_s(channelid,3) &"结果</font>") & "&ButtonSymbol=DownSearch';</script>")
			  Call KS.ReturnErr(1, "KS.Down.asp?channelid=" &channelid & "&Page=" & Page & "&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate)
			 End If
			 Exit Sub
		   End If
		   If Action="Verify" And Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10012") Then          '审核前台会员投稿的下载
			  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "&channelid= "& channelid & "';</script>")
			  Call KS.ReturnErr(1, "KS.Down.asp?channelid=" & channelid & "&Page=" & Page & "&ID=" & FolderID)
		   End If
		   
			Title       = Trim(RS("title"))
			DownVerSion = Trim(RS("DownVerSion"))
			PhotoUrl    = Trim(RS("PhotoUrl")) '缩略图
			BigPhoto    = Trim(RS("BigPhoto")) '大图
			DownLb      = Trim(RS("DownLB"))
			DownYY      = Trim(RS("DownYY"))
			DownSQ      = Trim(RS("DownSQ"))
			DownPT      = Trim(RS("DownPT"))
			DownSize    = Trim(RS("DownSize"))
			SizeUnit    = Right(DownSize, 2)
			DownSize    = Replace(DownSize, SizeUnit, "")
			If DownSize = "0" Then
			 DownSize = ""
			End If
			YSDZ        = Trim(RS("YSDZ"))
			ZCDZ        = Trim(RS("ZCDZ"))
			JYMM        = Trim(RS("JYMM"))
			DownUrls    = Trim(RS("DownUrls"))
			DownContent = Trim(RS("DownContent")) : If KS.IsNul(DownContent) Then DownContent=" "
			Recommend   = CInt(RS("Recommend"))
			Popular     = CInt(RS("Popular"))
			Verific     = CInt(RS("Verific"))
			Comment     = CInt(RS("Comment"))
			IsTop       = RS("IsTop")
			Rolls       = RS("Rolls")
			Strip       = RS("Strip")
			Slide       = RS("Slide")
			AddDate     = CDate(RS("AddDate"))
			Rank        = Trim(RS("Rank"))
			TemplateID  = RS("TemplateID")
			WapTemplateID=RS("WapTemplateID")
			Hits        = Trim(RS("Hits"))
			HitsByDay   = Trim(RS("HitsByDay"))
			HitsByWeek  = Trim(RS("HitsByWeek"))
			HitsByMonth = Trim(RS("HitsByMonth"))
			KeyWords    = Trim(RS("KeyWords"))
			Author      = Trim(RS("Author"))
			Origin      = Trim(RS("Origin"))
			FolderID    = RS("Tid")
			FileName    = RS("Fname")
			ReadPoint   = RS("ReadPoint")
			ChargeType  = RS("ChargeType")
			PitchTime   = RS("PitchTime")
			ReadTimes   = RS("ReadTimes")
			InfoPurview = RS("InfoPurview")
			arrGroupID  = RS("arrGroupID")
			DividePercent=RS("DividePercent")
			SEOTitle    = RS("SEOTitle")
			SEOKeyWord  = RS("SEOKeyWord")
			SEODescript = RS("SEODescript")
			   '自定义字段
				Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype!=0]")
				If diynode.length>0 Then
					Set FieldDictionary=KS.InitialObject("Scripting.Dictionary")
					For Each FNode In DiyNode
					   FieldDictionary.add lcase(FNode.SelectSingleNode("@fieldname").text),RS(FNode.SelectSingleNode("@fieldname").text)
					   If FNode.SelectSingleNode("showunit").text="1" Then
					   FieldDictionary.add lcase(FNode.SelectSingleNode("@fieldname").text) &"_unit",RS(FNode.SelectSingleNode("@fieldname").text&"_Unit")
					   End If
					Next
				End If
		End If
		'取得上传权限
		UpPowerFlag = KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10009")
		

		'取得下载参数
		 Dim DownLBList, DownYYList, DownSQList, DownPTList, RSP, DownLBStr, LBArr, YYArr, SQArr, PTArr, DownYYStr, DownSQStr, DownPTStr
		  Set RSP = Server.CreateObject("Adodb.RecordSet")
		  RSP.Open "Select top 1 * From KS_DownParam Where ChannelID=" & ChannelID, conn, 1, 1
		  If Not RSP.Eof Then
		   DownLBStr = RSP("DownLB")
		   DownYYStr = RSP("DownYY")
		   DownSQStr = RSP("DownSQ")
		   DownPTStr = RSP("DownPT")
		  End If
		  RSP.Close:Set RSP = Nothing
		  '下载类别
		 ' DownLBList="<option value="""" selected> </option>"
		  LBArr = Split(DownLBStr, vbCrLf)
		  For I = 0 To UBound(LBArr)
		   If LBArr(I) = DownLb Then
			DownLBList = DownLBList & "<option value='" & LBArr(I) & "' Selected>" & LBArr(I) & "</option>"
		   Else
			DownLBList = DownLBList & "<option value='" & LBArr(I) & "'>" & LBArr(I) & "</option>"
		   End If
		  Next
		  '下载语言
		  YYArr = Split(DownYYStr, vbCrLf)
		  For I = 0 To UBound(YYArr)
		   If YYArr(I) = DownYY Then
			DownYYList = DownYYList & "<option value='" & YYArr(I) & "' Selected>" & YYArr(I) & "</option>"
		   Else
			DownYYList = DownYYList & "<option value='" & YYArr(I) & "'>" & YYArr(I) & "</option>"
		   End If
		  Next
		'下载授权
		  SQArr = Split(DownSQStr, vbCrLf)
		  For I = 0 To UBound(SQArr)
		   If SQArr(I) = DownSQ Then
			DownSQList = DownSQList & "<option value='" & SQArr(I) & "' Selected>" & SQArr(I) & "</option>"
		   Else
			DownSQList = DownSQList & "<option value='" & SQArr(I) & "'>" & SQArr(I) & "</option>"
		   End If
		  Next
		'下载平台
		  PTArr = Split(DownPTStr, vbCrLf)
		  For I = 0 To UBound(PTArr)
			DownPTList = DownPTList & "<a href='javascript:SetDownPT(""" & PTArr(I) & """)'>" & PTArr(I) & "</a>/"
		  Next
        .Write"<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">"
		.Write "<html xmlns=""http://www.w3.org/1999/xhtml"">"
		.Write "<head>"
		.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
		.Write "<title>添加</title>"
		.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>" & vbCrlf
		.Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>" & vbCrlf
		.Write "<script language='JavaScript' src='../KS_Inc/Jquery.js'></script>" & vbCrlf
		.Write "<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>" & vbCrlf
		.Write "<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & vbCrlf
		.Write "<script language=""javascript"" src=""../KS_Inc/popcalendar.js""></script>" & vbCrlf
		.Write "<script type=""text/javascript"" src=""../editor/ckeditor.js"" mce_src=""../editor/ckeditor.js""></script>"
		.Write "<script language='javascript' src='../ks_inc/kesion.box.js'></script>"
		.Write "<script>var DownUrls='" & DownUrls & "';</script>"
        %>
		<script language="javascript">
		$(document).ready(function(){
		      $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",false);
			  $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",false);
			 <%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='keywords']/showonform").text="1" Then%>
			  $('#KeyLinkByTitle').click(function(){GetKeyTags();});
			 <%End If%>

		});
        function GetKeyTags()
		{
			  var text=escape($('input[name=Title]').val());
			  if (text!=''){
				  $('#KeyWords').val('请稍等,系统正在自动获取tags...').attr("disabled",true);
				  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){$('#KeyWords').val(unescape(data)).attr("disabled",false);});
			  }else{alert('对不起,请先输入内容!');}
		}
        function SetDownPT(addTitle){
			var str=$('#DownPT').val();
			if ($('#DownPT').val()=="") {
				$('#DownPT').val($('#DownPT').val()+addTitle);
			}else{
				if (str.substr(str.length-1,1)=="/"){
					$('#DownPT').val($('#DownPT').val()+addTitle);
				}else{
					$('#DownPT').val($('#DownPT').val()+"/"+addTitle);
				}
			}
			$('#DownPT').focus();
		}		
		
		function SetDownUrlByUpLoad(DownUrlStr,FileSize)
		{  var flag=false;
		   <%If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='nature']/showonform").text="1" Then%>
		   if (FileSize!=0)
		    { 
			  if (FileSize/1024/1024>1)
			  {
			   $("input[name=SizeUnit]")[1].checked=true;
			   $('#DownSize').val((FileSize/1024/1024).toFixed(2)); 
			  }
			  else{
			  document.getElementById('DownSize').value=(FileSize/1024).toFixed(2);
			   document.all.SizeUnit[0].checked=true;
			  }
			 }
		  <%end if%>
		  if ($('input[name=no]').val()==1){		    
		   $('input[name=DownAddress1]').val(DownUrlStr); 
		   return;
		   }else
		   {for(var i=1;i<=$('input[name=no]').val();i++)
		    {
			  if ($('input[name=DownAddress'+i+']').val()=='')
		       {$('input[name=DownAddress'+i+']').val(DownUrlStr);  return; flag=true;}
			}
		   }
		   if (flag==false){$('input[name=DownAddress1]').val(DownUrlStr);  }
		}
		function SelectAll(){$("#SpecialID>option").each(function(){$(this).attr("selected",true); });}
		function UnSelectAll(){$("#SpecialID>option").each(function(){$(this).attr("selected",false);});}
		function GetFileNameArea(f){$('#filearea').toggle(f);}
		function GetTemplateArea(f){$('#templatearea').toggle(f);}
		function SubmitFun()
		{ 
		  if ($('input[name=title]').val()==""){
				alert("请输入<%=KS.C_S(ChannelID,3)%>名称！");
				$('input[name=title]').focus();
				return;
		  }
		 if ($("#tid option:selected").val()=='0'){
			    alert('请选择所属栏目!');
			    return false;
		   }
			   <%
			  Call LFCls.ShowDiyFieldCheck(FieldXML,1)
			     %>
		   for (var i=1;i<=$('input[name=no]').val();i++){
			  if($('input[name=DownAddress'+i+']').val()!=''&&$('input[name=DownAddress'+i+']').val()!='del')
			  {  var downname='立即下载';
			     if ($('input[name=DownName'+i+']').val()!='') downname=$('input[name=DownName'+i+']').val();
				 downname=downname.replace('|','');
				 var dv=$('select[name=serverid'+i+']').val()+'|'+downname+'|'+$('input[name=DownAddress'+i+']').val();
			   if ($('input[name=DownUrls]').val()=='') $('input[name=DownUrls]').val(dv);
			  else $('input[name=DownUrls]').val($('input[name=DownUrls]').val()+'|||'+dv);
			  }
			}
		 if ($("input[name=addtype]:checked").val()==1){
		   if ($('#AddressUrls').val()==''){
			 alert('请输入批量地址列表');
			 $('#AddressUrls').focus();
			 return;
		   }
		}else{
			if($('#DownUrls').val()=='')
			{
			 alert('请输入下载地址');
			 $('#DownAddress1').focus();
			 return;
			}
		}

			$('#myform').submit();
			$(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
			$(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
		}	
		</script>
		<%
		.Write "</head>"
		.Write "<body leftmargin='0' topmargin='0' marginwidth='0' onkeydown='if (event.keyCode==83 && event.ctrlKey) SubmitFun();' marginheight='0'>"
		.Write "<div align='center'>"
			.Write "<ul id='menu_top'>"
			.Write "<li onclick=""return(SubmitFun())"" class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/save.gif' border='0' align='absmiddle'>确定保存</span></li>"
			.Write "<li onclick=""history.back();"" class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>取消返回</span></li>"
		    .Write "</ul>"
			
			.Write "<div class=tab-page id=DownPane>"
			.Write " <SCRIPT type=text/javascript>"
			.Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""DownPane"" ), 1 )"
			.Write " </SCRIPT>"
				 
			.Write " <div class=tab-page id=basic-page>"
			.Write "  <H2 class=tab>基本信息</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "		 tabPane1.addTabPage( document.getElementById( ""basic-page"" ) );"
			.Write "	</SCRIPT>"
			.Write " <TABLE width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"
		.Write "    <form  action='?Method=Save&ChannelID=" & ChannelID & "' method='post' id='myform' name='myform' onsubmit='return(SubmitFun())'>"
		.Write "      <input type='hidden' value='" & DownID & "' name='DownID'>"
		.Write "      <input type='hidden' value='" & Action & "' name='Action'>"
		.Write "      <input type='hidden' name='Page' value='" & Page & "'>"
		.Write "      <input type='hidden' name='KeyWord' value='" & KeyWord & "'>"
		.Write "      <input type='hidden' name='SearchType' value='" & SearchType & "'>"
		.Write "      <Input type='hidden' name='StartDate' value='" & StartDate & "'>"
		.Write "      <input type='hidden' name='EndDate' value='" & EndDate & "'>"
		
		'收集签收下载的参数
		.Write "      <Input type='hidden' name='DownStatus' value='" & KS.G("DownStatus") & "'>"
		.Write "      <input type='hidden' name='ID' value='" & KS.G("ID") & "'>"
		.Write "      <input type='hidden' name='Inputer' value='" &Inputer & "'>"
		
		If KS.G("ID")="" Then
		%>
		  <script type="text/javascript">
		    function selectFolder(){
		        new KesionPopup().PopupCenterIframe('选择批量添加的目录','KS.Down.asp?action=SelectFolder&channelid=<%=channelid%>',450,400,'auto')
			}
			function returnFolder(path){
			 if (path!=''){
			 $('#folderpath').val(path);
			 }
			 top.frames["MainFrame"].closeWindow();
			 getAddress();
			}
		    function setAddMode(){ 
			   if ($("input[name=addtype]:checked").val()==1){
			      $("#add1").show();
				  $("#add0").hide();
			      $("input[name=title]").val('{$FileName}');
			      $("input[name=KeyWords]").val('{$FileName}');
				  $("#DownSize").val('{$Size}');
					 insertHTMLToEditor('{$FileName}');
			   }else{
			      $("#add1").hide();
				  $("#add0").show();
			      $("input[name=title]").val('');
			      $("input[name=KeyWords]").val('');
				  $("#DownSize").val('');
				  CKEDITOR.instances.Content.setData('');
			   }
			}
			function getAddress(){
			  if ($('#subfolder[checked=true]').val()==undefined){
			   subfolder=0;
			  }else{
			   subfolder=1;
			  }
			  $(parent.document).find("#ajaxmsg").toggle();
			  $.ajax({
			  url: "ks.down.asp",
			  cache: false,
			  data: 'action=geturls&channelid=<%=channelid%>&folderpath='+$('#folderpath').val()+'&subfolder='+subfolder+'&ext='+$('#ext').val(),
			  success: function(d){
			   $(parent.document).find("#ajaxmsg").toggle();
			   $('#AddressUrls').val(d);
			  }})
			}
			function insertHTMLToEditor(codeStr) 
			{   
			  CKEDITOR.instances.Content.insertHtml(codeStr);
			} 
		  </script>
		<%
		.Write "              <tr class='tdbg' style='height:25px'>"
		.Write "                <td height='20' width='80' class='clefttitle'><div align='right'><font color='#FF0000'><strong>添加方式:</strong></font></div></td>"
		.Write "                <td><label><input name='addtype' type='radio' onclick='setAddMode()' checked value='0'>直接添加</label> <label><input name='addtype' type='radio' onclick='setAddMode()' value='1'>批量添加</label>"
		.Write "</td></tr>"
		End If


	For Each FNode In FieldNode
	    If KS.ChkClng(FNode.SelectSingleNode("fieldtype").text)<>0 Then
			.Write   KSCls.GetDiyField(ChannelID,FieldXML,FNode,FieldDictionary,0) '自定义字段
		Else
		 Dim XTitle:XTitle=FNode.SelectSingleNode("title").text
	     Select Case lcase(FNode.SelectSingleNode("@fieldname").text)
	       case "title"
				.Write "  <tr class='tdbg' style='height:25px'>"
				.Write "   <td height='20' width='80' class='clefttitle'><div align='right'><font color='#FF0000'><strong>" & XTitle & ":</strong></font></div></td>"
				.Write "  <td><input name='title' type='text'  class='textbox' value='" & Title & "' size=50><font color='#FF0000'>*</font>"
				If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='version']/showonform").text="1" Then
				.Write "&nbsp;&nbsp;版本号&nbsp;&nbsp;<input name='DownVerSion' type='text'  class='textbox' value='" & DownVerSion & "' size=10>"
				End If
				If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='pub']/showonform").text="1" Then
					.Write "<label><input type='checkbox' name='MakeHtml' value='1' checked>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='pub']/title").text & "</label>"
					End IF
				.Write "   </td>"
				.Write " </tr>" &vbcrlf
		  case "tid"
                .Write " <tr class='tdbg'>"
				.Write " <td class='clefttitle'><div align='right'><strong>" & XTitle & ":</strong></div></td>"
				.Write " <td><input type='hidden' name='OldClassID' value='" & FolderID & "'> "
					.Write " <select size='1' name='tid' id='tid' style='width:160px'>"
					.Write " <option value='0'>--请选择栏目--</option>"
					.Write Replace(KS.LoadClassOption(ChannelID,true),"value='" & FolderID & "'","value='" & FolderID &"' selected") & " </select>"
				If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attribute']/showonform").text="1" Then
				.Write " " & XTitle & " <label><input name='Recommend' type='checkbox' id='Recommend' value='1'"
				If Recommend = 1 Then .Write (" Checked")
				.Write ">推荐</label><label><input name='Popular' type='checkbox' id='Popular' value='1'"
				If Popular = 1 Then .Write (" Checked")
				.Write ">热门</label><label><input name='Rolls' type='checkbox' id='Rolls' value='1'"
				If Rolls = 1 Then .Write (" Checked")
				.Write ">滚动</label><label><input name='Strip' type='checkbox' id='Strip' value='1'"
				If Strip = 1 Then .Write (" Checked")
				.Write ">头条</label><label><input name='Slide' type='checkbox' id='Slide' value='1'"
				If Slide = 1 Then .Write (" Checked")
				.Write ">幻灯</label><label><input name='IsTop' type='checkbox' id='IsTop' value='1'"
				If IsTop = 1 Then .Write (" Checked")
				.Write ">固顶</label><label><input name='Comment' type='checkbox' id='Comment' value='1'"
				If Comment = 1 Then .Write (" Checked")
				.Write ">允许评论</label>" 
				Call KSCls.GetDiyAttribute(FieldXML,FieldDictionary)
				.Write "</td></tr>" &vbcrlf
				End If		  
		case "nature" 
				.Write "  <tr class='tdbg'>"
				.Write "  <td align='right' class='clefttitle'><strong>" & XTitle & ":</strong></td><td>类别:<select name='DownLB'>"
				.Write DownLBList
				.Write "  </select>&nbsp; 语言:<select name='DownYY' size='1'>"
				.Write DownYYList
				.Write "</select>&nbsp; 授权:<select name='DownSQ' size='1'>"
				.Write DownSQList
				.Write "</select>&nbsp;&nbsp;文件大小:<input maxlength='20' class='textbox' type='text' size=9 id='DownSize' name='DownSize' value='" & DownSize & "'>&nbsp;"
				If SizeUnit = "KB" Then
				.Write " <input name=""SizeUnit"" type=""radio"" value=""KB"" checked id=""kb""><label for=""kb"">KB</label> " & vbCrLf
				.Write " <input type=""radio"" name=""SizeUnit"" value=""MB"" id=""mb""><label for=""mb"">MB</label> " & vbCrLf
				Else
				.Write " <input name=""SizeUnit"" type=""radio"" value=""KB""  id=""kb""><label for=""kb"">KB</label> " & vbCrLf
				.Write " <input type=""radio"" name=""SizeUnit"" value=""MB"" checked id=""mb""><label for=""mb"">MB</label> " & vbCrLf
				End If
				.Write "  </td></tr>" &vbcrlf
		case "platform"
				.Write " <tr class='tdbg'>"
				.Write "    <td align='right' class='clefttitle'><strong>" & XTitle & ":</strong></td><td><input type='text' class='textbox' size=79 name='DownPT' id='DownPT' value='" & DownPT & "'><br>"
				.Write "    <font color='#808080'>平台选择"
				.Write DownPTList
				.Write "</font></td>"
				.Write "</tr>" &vbcrlf
		case "photourl"
				.Write " <tr class='tdbg'>"
				.Write "  <td align='right' class='clefttitle'><strong>" & XTitle & ":</strong></td>"
				.Write "  <td> <input name='PhotoUrl' type='text' id='PhotoUrl' size='45' value='" & PhotoUrl & "' class='textbox'><input type='hidden' value='" & BigPhoto & "' name='BigPhoto' id='BigPhoto'>"
				.Write "   &nbsp;<input class='button' type='button' name='Submit' value='选择图片...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID &"&CurrPath=" & CurrPath & "/DownPhoto',550,290,window,$('#PhotoUrl')[0]);document.myform.BigPhoto.value=document.myform.PhotoUrl.value;""> <input class='button' type='button' name='Submit' value='远程抓图...' onClick=""OpenThenSetValue('Include/Frame.asp?FileName=SaveBeyondfile.asp&PageTitle='+escape('抓取远程图片')+'&ItemName=图片&CurrPath=" & CurrPath & "/DownPhoto',300,100,window,$('#PhotoUrl')[0]);"">"
				.Write "                  <input class=""button""  type='button' name='Submit' value='裁剪...' onClick=""if($('#PhotoUrl').val()==''){alert('请选择图片或是上传后再使用此功能');return false;}else{OpenImgCutWindow(1,'" & KS.Setting(3) & "',$('#PhotoUrl').val())}"">  "
				.Write "  </td></tr>" & vbcrlf
		case "uploadphoto"
			   If CBool(UpPowerFlag) = True Then
				.Write "<tr class='tdbg'>"
				.Write "  <td align='right' height='25' class='clefttitle'><strong><font color=blue>" & XTitle & ":</font></strong></td>"
				.Write "  <td align='left'><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=Pic' frameborder=0 scrolling=no width='100%' height='30'></iframe>"
				.Write " </td></tr>"& vbcrlf
				End If
		case "keywords"
		        .Write " <tr class='tdbg'>"
				.Write "  <td class='clefttitle'><div align='right'><strong>" & XTitle & ":</strong></div></td>"
				.Write "  <td height='50'> <input name='KeyWords' type='text' id='KeyWords' class='textbox' value='" & KeyWords & "' size=40> <="
				.Write "  <select name='SelKeyWords' style='width:150px' onChange='InsertKeyWords(document.getElementById(""KeyWords""),this.options[this.selectedIndex].value)'>"
				.Write "<option value="""" selected> </option><option value=""Clean"" style=""color:red"">清空</option>"
				.Write KSCls.Get_O_F_D("KS_KeyWords","KeyText","IsSearch=0 Order BY AddDate Desc")
				.Write " </select>"
				.Write " <br />【<a href=""#"" id=""KeyLinkByTitle"" style=""color:green"">根据" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='title']/title").text & "自动获取Tags</a>】<input type='checkbox' name='tagstf' value='1' checked>写入Tags表  <span class='help' onclick=""new KesionPopup().mousepop('什么是Tag','Tag(标签)是一种更为灵活、有趣的分类方式，您可以为每篇文章添加一个或多个Tag(标签)，你可以看到网站上所有和您使用了相同Tag的内容，由此和他人产生更多的联系。Tag体现了群体的力量，使得内容之间的相关性和用户之间的交互性大大增强。多个Tag请用英文逗号隔开',300)"">帮助</span>"
				.Write " </td></tr>"& vbcrlf
		case "address"
		  %>
		      <script language="javascript">
				  function setid() {
				 str='';
				 if($('input[name=no]').val()=='')	 $('input[name=no]').val(1);
				<%If Action="Edit" Then
				 Response.Write "for(i=$('input[name=topnum]').val();i<=$('input[name=no]').val();i++)"
				Else
				 Response.Write "for(i=2;i<=$('input[name=no]').val();i++)"
				End If
				%>
				 str+='<input type="text" name="DownName'+i+'" class="textbox" value="下载地址'+i+'" size="8">-'+'<select style="width:140" name="serverid'+i+'"><%=SelDownServer(0)%></select><input class="textbox" type="text" name="DownAddress'+i+'" size="40" value="">&nbsp;<input type="button" class="button" value="选择下载地址..." name="button1" onClick="OpenThenSetValue(\'Include/SelectPic.asp?ChannelID=<%=ChannelID%>&CurrPath=<%=CurrPath%>/DownUrl\',550,290,window,document.myform.DownAddress'+i+');"><br>';
				 $('#upid').html(str);
			 }
			 </script>
				<tbody id="add0">
				<tr class="tdbg">
					  <td align="right" class='clefttitle'><b>地址数:</b></td>
					  <td><input type="text" class="textbox" name="no"value="1" size=2>&nbsp;&nbsp;<input type="button" name="Button" class="button" onClick="setid();" value="添加下载地址数"> <span><font color=blue>下载服务器路径 + 下载文件名称 = 完整下载地址</font>,<font color=red>删除某个地址,请在地址里输入 "del"或留空。</font></span>
				  </td>
				</tr>
				<tr class="tdbg">
				  <td align="right" class="clefttitle"><b><%=XTitle%>:</b><br></td>
					  <td><input type="hidden" name="DownUrls" id="DownUrls">
			     <%	
			   If Action="Edit" Or Action="Verify" Then
				Dim DownUrlsArr:DownUrlsArr=Split(DownUrls,"|||")
				For I=1 To Ubound(DownUrlsArr)+1
					Dim UrlsParam:UrlsParam=Split(DownUrlsArr(I-1),"|")
					Response.Write "<input name=""DownName" & I & """ type=""text"" size=""8"" value=""" & UrlsParam(1) &""">-"
					Response.Write "<select name=""serverid" & I & """ size=""1"" style=""width:140"">"
					Response.Write SelDownServer(UrlsParam(0))
					Response.Write "</select>"
					Response.Write "<input name=""DownAddress" & I & """ value=""" & UrlsParam(2) & """ type=""text"" size=""40"">&nbsp;"
					Response.Write "<input type=""button"" class=""button"" value=""选择下载地址..."" name=""button1"" onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID &"&CurrPath=" & CurrPath & "/DownUrl',550,290,window,document.myform.DownAddress" & I & ");"">"
					Response.Write "<br>"
				Next
					Response.Write "<script>$('input[name=no]').val(" & I-1 & ");</script>"
					Response.Write "<input type='hidden' name='topnum' value='" & I & "'>"
			   Else
				Response.Write "<input name=""DownName1"" class=""textbox"" type=""text"" size=""8"" value=""下载地址1"">-"
				Response.Write "<select name=""serverid1"" size=""1"" style=""width:140"">"
				Response.Write SelDownServer(0)
				Response.Write "</select>"
				Response.Write "<input class=""textbox"" name=""DownAddress1"" id=""DownAddress1"" type=""text"" size=""40"">&nbsp;"
				Response.Write "<input type=""button"" class=""button"" value=""选择下载地址..."" name=""button1"" onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID &"&CurrPath=" & CurrPath & "/DownUrl',550,290,window,$('input[name=DownAddress1]')[0]);"">"
				Response.Write "<br>"
			  End If
				.Write "<span id=""upid""></span></td>"
				.Write "</tr>"
				
				 If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='uploadsoft']/showonform").text="1" Then
					If CBool(UpPowerFlag) = True Then
					.Write "<tr><td height=25 align='right' class='clefttitle'><strong>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='uploadsoft']/title").text & ":</strong></td>"
					.Write "<td align='left'><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='KS.UpFileForm.asp?UpType=UpByBar&ChannelID=" & ChannelID & "' frameborder=0 scrolling=no width='300' height='30'></iframe></td></tr>"
					End If
				 End If
				.Write "</tbody>"
				.Write "<tbody id='add1' style='display:none'>"
				.Write "<tr class='tdbg'><td height=25 align='right' class='clefttitle'><strong>批量地址列表:</strong></td>"
				.Write "<td>输入目录<input type='text' name='folderpath' id='folderpath' size='26' value='/uploadfiles/" & year(now) & "-" & right("0" & month(now),2) & "/'> <input type='button' onclick='selectFolder()' value='选择...' class='button' /> <label><input type='checkbox' name='subfolder' id='subfolder' value='1' checked>包含子栏目</label>  扩展名限定<input type='text' name='ext' id='ext' value='.rar|.exe|.jpg' size=16> <input type='button' class='button' onclick='getAddress()' value='自动获取'><br/>地址标题:<input type='text' name='AddressName' value='下载地址1' /> "
				.Write "选择服务器：<select name=""serverids"" size=""1"" style=""width:140"">"
				.Write SelDownServer(0)
				.Write "</select>"

				.Write "<br/><textarea name='AddressUrls' id='AddressUrls' style='width:400px;height:200px'></textarea><br/><font color=blue>格式:一行对应一个软件地址</font></td></tr>"
				.Write "</tbody>"	
         case "author"
		        .Write " <tr class='tdbg'>"
				.Write "  <td class='clefttitle'><div align='right'><strong>" & XTitle & ":</strong></div></td>"
				.Write "  <td> <input name='author' type='text' id='author' value='" & Author & "' size=30 class='textbox'>                 <=【<font color='blue'><font color='#993300' onclick='$(""#author"").val(""未知"")' style='cursor:pointer;'>未知</font></font>】【<font color='blue'><font color='#993300' onclick=""$('#author').val('佚名')"" style='cursor:pointer;'>佚名</font></font>】【<font color='blue'><font color='red' onclick=""$('#author').val('" & KS.C("AdminName") & "')"" style='cursor:pointer;'>" & KS.C("AdminName") & "</font></font>】"
								 If Author <> "" And Author <> "未知" And Author <> KS.C("AdminName") And Author <> "佚名" Then
								  .Write ("【<font color='blue'><font color='#993300' onclick=""$('#author').va('" & Author & "')"" style='cursor:pointer;'>" & Author & "</font></font>】")
								 End If
								  .Write ("<select name='SelAuthor' style='width:100px' onChange=""$('#author').val(this.options[this.selectedIndex].value)"">")
				.Write "<option value="""" selected> </option><option value="""" style=""color:red"">清空</option>"
				.Write KSCls.Get_O_F_D("KS_Origin","OriginName","ChannelID=0 and OriginType=1 Order BY AddDate Desc")
				.Write " </select> &nbsp; </td></tr>" & vbcrlf
		case "origin"
                .Write "<tr class='tdbg'>"
				.Write "  <td class='clefttitle'><div align='right'><strong>" & XTitle & ":</strong></div></td>"
				.Write "  <td> <input name='Origin' type='text' id='Origin' value='" & Origin & "' size=30 class='textbox'>                 <=【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('不详');"" style='cursor:pointer;'>不详</font></font>】【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('本站原创')"" style='cursor:pointer;'>本站原创</font></font>】【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('互联网')"" style='cursor:pointer;'>互联网</font></font>】"
								  If Origin <> "" And Origin <> "不详" And Origin <> "本站原创" And Origin <> "互联网" Then
								  .Write ("【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('" & Origin & "')"" style='cursor:pointer;'>" & Origin & "</font></font>】 ")
								   End If
								  .Write ("<select name='selOrigin' style='width:100px' onChange=""$('#Origin').val(this.options[this.selectedIndex].value)"">")
				.Write "<option value="""" selected> </option><option value="""" style=""color:red"">清空</option>"
				.Write KSCls.Get_O_F_D("KS_Origin","OriginName","OriginType=0 Order BY AddDate Desc")
				.Write " </select> </td></tr>" &vbcrlf
		case "content"
				.Write " <tr class='tdbg'>"
				.Write "  <TD align='right' class='clefttitle'><strong>" & XTitle & ":</strong><br><input name='BeyondSavePic' type='checkbox' value='1'><font color=green>自动下载<br>简介里的图片</font></td>"
				.Write " <td><textarea ID='Content' name='Content' style='display:none'>" & Server.HTMLEncode(DownContent) & "</textarea><script type=""text/javascript"">"
				.Write "CKEDITOR.replace('Content', {width:""98%"",height:""160px"",toolbar:""Basic"",filebrowserBrowseUrl :""Include/SelectPic.asp?from=ckeditor&channelid=" & ChannelID &"&Currpath="& KS.GetUpFilesDir() &""",filebrowserWindowWidth:650,filebrowserWindowHeight:290});"
				.Write "</script>"
				.Write "</TD></tr>"
		case "ysdz"
				.Write "<tr class='tdbg'>"
				.Write " <td  align='right' class='clefttitle'><strong>" & XTitle & ":</strong></td>"
				.Write "  <td> <input name='YSDZ' type='text' id='YSDZ' value='" & YSDZ & "' size='50' class='textbox'>"
				.Write " </td></tr>" &vbcrlf
		case "zcdz"
				.Write " <tr class='tdbg'>"
				.Write " <td  align='right' class='clefttitle'><strong>" & XTitle & ":</strong></td>"
				.Write " <td> <input name='ZCDZ' type='text' id='ZCDZ' value='" & ZCDZ & "' size='50' class='textbox'>"
				.Write " </td></tr>"&vbcrlf
		case "jymm"
				.Write " <tr class='tdbg'>"
				.Write "  <td  align='right' class='clefttitle'><strong>" & XTitle & ":</strong></td>"
				.Write "  <td> <input name='JYMM' type='text' id='JYMM' value='" & JYMM & "' size='50' class='textbox'>"
				.Write " </td></tr>" &vbcrlf
	  End Select
	End If
Next
	.Write "</table>"
	.Write "</div>"
		
	If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attroption']/showonform").text="1" Then	
		   .Write " <div class=tab-page id=classoption-page>"
		   .Write "  <H2 class=tab>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='attroption']/title").text &"</H2>"
		   .Write "	<SCRIPT type=text/javascript>"
		   .Write "				 tabPane1.addTabPage( document.getElementById( ""classoption-page"" ) );"
		   .Write "	</SCRIPT>"
            .Write "<TABLE style='margin:1px' width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"
			If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='special']/showonform").text="1" Then
			.Write "           <tr class='tdbg'>"
			.Write "              <td class='clefttitle' align='right'><strong>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='special']/title").text & ":</strong></td>"
			.Write "              <td>"
	        Call KSCls.Get_KS_Admin_Special(ChannelID,KS.ChkClng(KS.G("ID")))
			.write "</td>"
			.Write "           </tr>"
		  End If
		  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='adddate']/showonform").text="1" Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td class='clefttitle'><div align='right'><strong>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='adddate']/title").text & ":</strong></div></td>"
			.Write "                <td>"
			If Action <> "Edit" Then
			.Write ("<input name='AddDate' type='text' onclick=""popUpCalendar(this, this, dateFormat,-1,-1)"" id='AddDate' value='" & Now() & "' size='50'  class='textbox'>")
			Else
			.Write ("<input name='AddDate' type='text' onclick=""popUpCalendar(this, this, dateFormat,-1,-1)"" id='AddDate' value='" & AddDate & "' size='50'  readonly class='textbox'>")
			End If
			.Write "                  <b><a href='#' onClick=""popUpCalendar(this, $('input[name=AddDate]').get(0), dateFormat,-1,-1)""><img src='Images/date.gif'  border='0' align='absmiddle' title='选择日期'></a>日期格式：年-月-日 时：分：秒"
			.Write "               </td>"
			.Write "             </tr>"
	    End If
		If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='rank']/showonform").text="1" Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td class='clefttitle'><div align='right'><strong>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='rank']/title").text & ":</strong></td>"
			.Write "                <td><select name='rank'>"
			If Rank = "★" Then .Write "<option  selected>★</option>" Else .Write "<option>★</option>"
			If Rank = "★★" Then .Write "<option  selected>★★</option>" Else .Write "<option>★★</option>"
			If Rank = "★★★" Or Action = "Add" Then .Write "<option  selected>★★★</option>" Else .Write "<option>★★★</option>"
			If Rank = "★★★★" Then .Write "<option  selected>★★★★</option>" Else .Write "<option>★★★★</option>"
			If Rank = "★★★★★" Then .Write "<option  selected>★★★★★</option>" Else .Write "<option>★★★★★</option>"
			.Write "</select>&nbsp;请为" & KS.C_S(ChannelID,3) & "评定阅读等级"
			.Write "               </td>"
			.Write "             </tr>"
	   End If
	 If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='hits']/showonform").text="1" Then
				.Write "              <tr class='tdbg'>"
				.Write "                <td class='clefttitle'><div align='right'><strong>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='hits']/title").text & ":</strong></td><td>本日：<input name='HitsByDay' type='text' id='HitsByDay' value='" & HitsByDay & "' size='6' style='text-align:center' class='textbox'> 本周：<input name='HitsByWeek' type='text' id='HitsByWeek' value='" & HitsByWeek & "' size='6' style='text-align:center' class='textbox'> 本月：<input name='HitsByMonth' type='text' id='HitsByMonth' value='" & HitsByMonth & "' size='6' style='text-align:center' class='textbox'> 总计：<input name='Hits' type='text' id='Hits' value='" & Hits & "' size='6' style='text-align:center' class='textbox'>&nbsp;初数点击数作弊" 
				.Write "                  </td>"
				.Write "              </tr>"
	 End If
	 If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='template']/showonform").text="1" Then
			 .Write "             <tr class='tdbg'>"
			 .Write "               <td class='clefttitle'><div align='right'><strong>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='template']/title").text & ":</strong></div></td>"
			.Write "                <td> "
			IF Action <> "Edit" and  Action<>"Verify" Then
			.Write " <input type='radio' name='templateflag' onclick='GetTemplateArea(false);' value='2' checked>继承栏目设定<input type='radio' onclick='GetTemplateArea(true);' name='templateflag' value='1'>自定义"
			.Write "<div id='templatearea' style='display:none'>"
			If KS.WSetting(0)="1" Then .Write "<strong>WEB模板</strong> "
			.Write "<input id='TemplateID' name='TemplateID' readonly size=30 class='textbox' value='" & TemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]") 
			If KS.WSetting(0)="1" Then 
			.Write "<br/><strong>3G版模板</strong> "
			.Write "<input id='WapTemplateID' name='WapTemplateID' readonly size=30 class='textbox' value='" & WapTemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#WapTemplateID')[0]") 
			End If
			.Write "</div>"
			Else
			.Write "<div id='templatearea'>"
			If KS.WSetting(0)="1" Then .Write "<strong>WEB模板</strong> "
			.Write "<input id='TemplateID' name='TemplateID' readonly maxlength='255' size=30 class='textbox' value='" & TemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]")
			If KS.WSetting(0)="1" Then 
			.Write "<br/><strong>3G版模板</strong> "
			.Write "<input id='WapTemplateID' name='WapTemplateID' readonly size=30 class='textbox' value='" & WapTemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#WapTemplateID')[0]") 
			End If
			.Write "</div>"
			End If
			.Write "                </td>"
			.Write "             </tr>"
	  End If
	  
	  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='fname']/showonform").text="1" Then
				.Write "             <tr class='tdbg'>"
				.Write "               <td class='clefttitle'><div align='right'><strong>" & FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='fname']/title").text & ":</strong></td><td>"
				IF Action = "Edit" or Action="Verify" Then
				.Write "<input name='FileName' type='text' id='FileName' readonly  value='" & FileName & "' size='25' class='textbox'> <font color=red>不能改</font>"
				Else
				.Write "<input type='radio' value='0' name='filetype' onclick='GetFileNameArea(false);' checked>自动生成 <input type='radio' value='1' name='filetype' onclick='GetFileNameArea(true);' >自定义"
				.Write "<div id='filearea' style='display:none'><input name='FileName' type='text' id='FileName'   value='" & FileName  & "' size='25' class='textbox'> <font color=red>可带路径,如 help.html,news/news_1.shtml等</font></div>"
				End IF
				 .Write "                  </td>"
				 .Write "             </tr>"
		 End If
End If
			.Write "</table>"
			.Write "</div>"
			
	     If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='seooption']/showonform").text="1" Then
		  KSCls.LoadSeoOption ChannelID,FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='seooption']/title").text,SEOTitle,SEOKeyWord,SEODescript
         End If			
 
	     If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='chargeoption']/showonform").text="1" Then
	      KSCls.LoadChargeOption ChannelID,ChargeType,InfoPurview,arrGroupID,ReadPoint,PitchTime,ReadTimes,DividePercent
         End If
		 If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='relativeoption']/showonform").text="1" Then
          KSCls.LoadRelativeOption ChannelID,KS.ChkClng(KS.G("ID"))
		 End If
 		 .Write "</form>"
		 .Write " </div>"
		 .Write "</body>"
		 .Write "</html>"
		 End With
		 if rs.state=1 then rs.close:Set rs=nothing
		End Sub
		
		Sub DownSave()
         ' On Error Resume Next
		  Dim SelectInfoList,HasInRelativeID
          With Response
			Page        = KS.G("Page")
			Action      = KS.G("Action") ' Add添加新下载 Edit编辑下载
			Title       = KS.G("Title")
			DownVerSion = KS.G("DownVerSion")
			DownLb      = KS.G("DownLb")
			DownYY      = KS.G("DownYY")
			DownSQ      = KS.G("DownSQ")
			DownPT      = KS.G("DownPT")
			DownSize    = KS.G("DownSize")
			If DownSize = "" Or Not IsNumeric(DownSize) Then DownSize = 0
			DownSize    = DownSize & KS.G("SizeUnit")
			YSDZ        = KS.G("YSDZ")
			ZCDZ        = KS.G("ZCDZ")
			JYMM        = KS.G("JYMM")
			PhotoUrl    = KS.G("PhotoUrl")
			BigPhoto    = KS.G("BigPhoto")
			DownContent = Request.Form("Content")
			Hits        = KS.ChkClng(KS.G("Hits"))
			HitsByDay   = KS.ChkClng(KS.G("HitsByDay"))
			HitsByWeek  = KS.ChkClng(KS.G("HitsByWeek"))
			HitsByMonth = KS.ChkClng(KS.G("HitsByMonth"))
			DownUrls    = Trim(KS.G("DownUrls"))
			Recommend   = KS.ChkClng(KS.G("Recommend"))
			Popular     = KS.ChkClng(KS.G("Popular"))
			IsTop       = KS.ChkClng(KS.G("IsTop"))
			Comment     = KS.ChkClng(KS.G("Comment"))
			Slide       = KS.ChkClng(KS.G("Slide"))
			Rolls       = KS.ChkClng(KS.G("Rolls"))
			Strip       = KS.ChkClng(KS.G("Strip"))
			Makehtml    = KS.ChkClng(KS.G("Makehtml"))
			Tid         = KS.G("Tid")
			SpecialID   = Replace(KS.G("SpecialID")," ",""):SpecialID = Split(SpecialID,",")
			SelectInfoList = Replace(KS.G("SelectInfoList")," ","")
			KeyWords    = KS.G("KeyWords")
			Author      = KS.G("Author")
			Origin      = KS.G("Origin")
			AddDate     = KS.G("AddDate")
			If Not IsDate(AddDate) Then AddDate=Now
			Rank        = Trim(KS.G("Rank"))
			'SEO优化选项
			SEOTitle    = KS.G("SEOTitle")
			SEOKeyWord  = KS.G("SEOKeyWord")
			SEODescript = KS.G("SEODescript")
			'收费选项
			ReadPoint   = KS.ChkClng(KS.G("ReadPoint"))
		    ChargeType  = KS.ChkClng(KS.G("ChargeType"))
			PitchTime   = KS.ChkClng(KS.G("PitchTime"))
			ReadTimes   = KS.ChkClng(KS.G("ReadTimes"))
			InfoPurview = KS.ChkClng(KS.G("InfoPurview"))
			arrGroupID  = KS.G("GroupID")
			DividePercent=KS.G("DividePercent"):IF Not IsNumeric(DividePercent) Then DividePercent=0
             
				TemplateID = KS.G("TemplateID")
				WapTemplateID=KS.G("WapTemplateID")
				Dim FnameType:FnameType=KS.C_C(TID,23)
				 If KS.ChkClng(KS.G("filetype"))=0 Then
					If Action = "Add" OR Action="Verify" Then
						Fname=KS.GetFileName(KS.C_C(TID,24), Now, FnameType)
					 End If
				 Else
				     Fname=KS.G("FileName")
				 End If
				 If KS.ChkClng(KS.G("TemplateFlag"))=2 Or TemplateID="" Then TemplateID=KS.C_C(TID,5):WapTemplateID=KS.C_C(TID,22)

				Call KSCls.CheckDiyField(FieldXML,ErrMsg)  '检查自定义字段
			 
			If Title = "" Then
			 .Write ("<script>alert('" & KS.C_S(ChannelID,3) & "名称不能为空!');history.back(-1);</script>")
			 Exit Sub
			End If
			
			Dim addtype:addtype=KS.ChkClng(KS.G("addtype"))
			
			Set RS = Server.CreateObject("ADODB.RecordSet")
			If Tid = "" Then ErrMsg = ErrMsg & "[" & KS.C_S(ChannelID,3) & "类别]必选! \n"
			If Title = "" Then ErrMsg = ErrMsg & "[" & KS.C_S(ChannelID,3) & "标题]不能为空! \n"
			If Title <> "" And Tid <> "" And Action = "Add" And addtype=0 Then
			  SQLStr = "select top 1 * from " & KS.C_S(ChannelID,2) & " where Title='" & Title & "' And Tid='" & Tid & "'"
			   RS.Open SQLStr, conn, 1, 1
				If Not RS.EOF Then
				 ErrMsg = ErrMsg & "该类别已存在此项" & KS.C_S(ChannelID,3) & "! \n"
			   End If
			   RS.Close
			End If
			If ErrMsg <> "" Then
			   .Write ("<script>alert('" & ErrMsg & "');history.back(-1);</script>")
			   .End
			Else
				 If KS.ChkClng(KS.G("BeyondSavePic")) = 1 Then
						Dim SaveFilePath
						SaveFilePath = KS.GetUpFilesDir & "/"
						KS.CreateListFolder (SaveFilePath)
						DownContent= KS.ReplaceBeyondUrl(DownContent, SaveFilePath)
				 End If
			      If KS.ChkClng(KS.G("TagsTF"))=1 Then Call KSCls.AddKeyTags(KeyWords)
				  Verific=1
				  If Action = "Add" Then
				    
					Dim AddressUrls:AddressUrls=Trim(KS.G("AddressUrls"))
				    If addtype=0 Then  AddressUrls=Title
					Dim KK,AddressUrlsArr:AddressUrlsArr=Split(AddressUrls,vbcrlf)
					Dim EndNum,ALen:ALen=Ubound(AddressUrlsArr)
					If AddressUrlsArr(Alen)="" Then EndNum=Alen-1 Else EndNum=Alen
			For KK=0 To EndNum
			  If Not KS.IsNul(AddressUrlsArr(kk)) Then
						SQLStr = "select top 1 * from " & KS.C_S(ChannelID,2) & " where 1=0"
						RS.Open SQLStr, conn, 1, 3
						RS.AddNew
						If addtype=0 then
							RS("Title") = Title
							RS("DownSize") = DownSize
							RS("KeyWords") = KeyWords
							RS("DownContent") = DownContent
							RS("DownUrls") = DownUrls
						else
						    dim wlfilename:wlfilename=split(split(AddressUrlsArr(kk),"/")(ubound(split(AddressUrlsArr(kk),"/"))),".")(0)
							RS("Title") = replace(Title,"{$FileName}",wlfilename)
							If instr(request("downsize"),"{$Size}")>0 then
							  dim softsize:softsize=KS.GetFieSize(Server.MapPath(AddressUrlsArr(kk)))/1024
							  if softsize>1024 then
							   RS("DownSize")=round(softsize/1024,2) & "MB"
							  else
							   RS("DownSize")=round(softsize,2) & "KB"
							  end if
							else
							 RS("DownSize") = DownSize
							end if
							if instr(keywords,"{$FileName}")>0 then
							 Dim MaxLen:MaxLen=30
							 Dim WS:Set WS=New Wordsegment_Cls
							 RS("KeyWords") = replace(KeyWords,"{$FileName}",WS.SplitKey(wlfilename,4,MaxLen))
							 Set WS=Nothing
							else
							 RS("KeyWords") = keywords
							end if
							 RS("DownContent") = replace(DownContent,"{$FileName}",wlfilename)
							 RS("DownUrls") = KS.ChkClng(Request("serverids")) & "|" & KS.G("AddressName") & "|" & AddressUrlsArr(kk)
						end if
						
						RS("DownVerSion") = DownVerSion
						RS("DownLB")      = DownLb
						RS("DownYY")      = DownYY
						RS("DownSQ")      = DownSQ
						RS("DownPT")      = DownPT
						RS("YSDZ")        = YSDZ
						RS("ZCDZ")        = ZCDZ
						RS("JYMM")        = JYMM
						RS("PhotoUrl")    = PhotoUrl
						RS("BigPhoto")    = BigPhoto
						RS("Recommend")   = Recommend
						RS("Popular")     = Popular
						RS("Slide")       = Slide
						RS("Rolls")       = Rolls
						RS("Strip")       = Strip
						RS("Verific")     = 1
						RS("Comment")     = Comment
						RS("IsTop")       = IsTop
						RS("Tid")         = Tid
						RS("Author")      = Author
						RS("Origin")      = Origin
						RS("AddDate")     = AddDate
						RS("ModifyDate")  = AddDate 
						RS("Rank")        = Rank
						RS("Slide")       = Slide
						RS("TemplateID")  = TemplateID
						RS("WapTemplateID") = WapTemplateID
						RS("Hits")        = Hits
						RS("HitsByDay")   = HitsByDay
						RS("HitsByWeek")  = HitsByWeek
						RS("HitsByMonth") = HitsByMonth
						RS("Fname")       = Fname
						RS("Inputer")     = KS.C("AdminName")
						RS("RefreshTF")   = Makehtml
						RS("DelTF")       = 0
						RS("ReadPoint")   =	ReadPoint
						RS("ChargeType")  = ChargeType
						RS("PitchTime")   = PitchTime
						RS("ReadTimes")   = ReadTimes
						RS("InfoPurview") = InfoPurview
						RS("arrGroupID")  = arrGroupID
						RS("DividePercent")=DividePercent
						RS("SEOTitle")    = SEOTitle
						RS("SEOKeyWord")  = SEOKeyWord
						RS("SEODescript") = SEODescript
						Call KSCls.AddDiyFieldValue(RS,FieldXml)
						RS.Update
					
					   '写入Session,添加下一篇下载调用
					   Session("KeyWords") = KeyWords
					   Session("Author") = Author
					   Session("Origin") = Origin
						RS.MoveLast
						  If Left(Ucase(Fname),2)="ID" Or addtype<>0 Then
						   RS("Fname") = RS("ID") & FnameType
						   RS.Update
						  End If
						 For I=0 To Ubound(SpecialID)
							Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & RS("ID") & "," & ChannelID & ")")
						 Next
						 
						 Call KSCls.UpdateRelative(ChannelID,RS("ID"),SelectInfoList,0)
						 Call LFCls.AddItemInfo(ChannelID,RS("ID"),RS("Title"),Tid,RS("DownContent"),rs("KeyWords"),PhotoUrl,AddDate,KS.C("AdminName"),Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific,RS("Fname"))
						  '关联上传文件
						  Call KS.FileAssociation(ChannelID,RS("ID"),PhotoUrl & BigPhoto & DownContent & DownUrls,0)
						  
						  If Makehtml = 1 Then
							  If addType=0 Or KK=EndNum Then
								Call RefreshHtml(1)
							  ElseIf KS.C_S(ChannelID,7)=1 Or KS.C_S(ChannelID,7)=2 Then 
								.Write "<div style='display:none'><iframe src=""Include/RefreshHtmlSave.Asp?ChannelID=" & ChannelID &"&Types=Content&RefreshFlag=ID&ID=" & RS("ID") &""" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
							 End If
					     ElseIf Makehtml <>1 and KK=EndNum Then  
						   Call RefreshHtml(1)
						 End If
					    RS.Close
				   End If
				 Next

				 Set RS = Nothing
				ElseIf Action = "Edit" Or Action="Verify" Then
				If Action="Verify" Then 
					 Call KS.ReplaceUserFile(DownContent ,ChannelID)
					 Call KS.ReplaceUserFile(PhotoUrl,ChannelID)
					 Call KS.ReplaceUserFile(BigPhoto,ChannelID)
					 Call KS.ReplaceUserFile(DownUrls,ChannelID)
				End If
				DownID = KS.ChkClng(Request("DownID"))
				SQLStr = "SELECT top 1 * FROM " & KS.C_S(ChannelID,2) & " Where ID=" & DownID
					RS.Open SQLStr, conn, 1, 3
					If RS.EOF And RS.BOF Then
					 .Write ("<script>alert('参数传递出错!');history.back(-1);</script>")
					 .End
					End If
					RS("Title")       = Title
					RS("DownVerSion") = DownVerSion
					RS("DownLB")      = DownLb
					RS("DownYY")      = DownYY
					RS("DownSQ")      = DownSQ
					RS("DownPT")      = DownPT
					RS("DownSize")    = DownSize
					RS("YSDZ")        = YSDZ
					RS("ZCDZ")        = ZCDZ
					RS("JYMM")        = JYMM
					RS("PhotoUrl")    = PhotoUrl
					RS("BigPhoto")    = BigPhoto
					RS("DownContent") = DownContent
					RS("DownUrls")    = DownUrls
					RS("Recommend")   = Recommend
					RS("Popular")     = Popular
					RS("Comment")     = Comment
					RS("Slide")       = Slide
					RS("Rolls")       = Rolls
					RS("Strip")       = Strip
					RS("IsTop")       = IsTop
					RS("Tid")         = Tid
					RS("KeyWords")    = KeyWords
					RS("Author")      = Author
					RS("Origin")      = Origin
					RS("AddDate")     = AddDate
					RS("ModifyDate")  = Now 
					RS("Rank")        = Rank
					RS("Slide")       = Slide
					RS("TemplateID")  = TemplateID
					RS("WapTemplateID") = WapTemplateID
					If Makehtml = 1 Then
					 RS("RefreshTF")  = 1
					End If
					RS("Hits")        = Hits
					RS("HitsByDay")   = HitsByDay
					RS("HitsByWeek")  = HitsByWeek
					RS("HitsByMonth") = HitsByMonth
					RS("ReadPoint")   =	ReadPoint
				    RS("ChargeType")  = ChargeType
				    RS("PitchTime")   = PitchTime
				    RS("ReadTimes")   = ReadTimes
					RS("InfoPurview") = InfoPurview
					RS("arrGroupID")  = arrGroupID
					RS("DividePercent")=DividePercent
					If Action="Verify" Then  Inputer=RS("Inputer")
					RS("Verific")     = 1
					RS("SEOTitle")    = SEOTitle
					RS("SEOKeyWord")  = SEOKeyWord
					RS("SEODescript") = SEODescript
					Call KSCls.AddDiyFieldValue(RS,FieldXml)
					RS.Update
			       RS.MoveLast
			       If TID<>Request.Form("OldClassID") Then
					     Call KSCls.DelInfoFile(ChannelID,Request.Form("OldClassID"), "",RS("Fname"),RS("ID"))
				   End If
					Conn.Execute("Delete From KS_SpecialR Where InfoID=" & DownId & " and channelid=" & ChannelID)
					For I=0 To Ubound(SpecialID)
					Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & DownId & "," & ChannelID & ")")
					Next
					Call KSCls.UpdateRelative(ChannelID,DownID,SelectInfoList,1)
					Call LFCls.UpdateItemInfo(ChannelID,DownId,Title,Tid,DownContent,KeyWords,PhotoUrl,AddDate,Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific)
	 				  '关联上传文件
					  Call KS.FileAssociation(ChannelID,DownID,PhotoUrl & BigPhoto & DownContent & DownUrls,1)
					
				  Call RefreshHtml(2)
				  RS.Close:Set RS = Nothing
					IF Action="Verify" Then     '如果是审核投稿下载，对用户，进行加积分等，并返回签收下载管理
							  '对用户进行增值，及发送通知操作
							  IF Inputer<>"" And Inputer<>KS.C("AdminName") Then Call KS.SignUserInfoOK(ChannelID,Inputer,Title,DownId)
							 .Write ("<script> parent.frames['MainFrame'].focus();alert('恭喜，" & KS.C_S(ChannelID,3) & "成功签收!');location.href='KS.ItemInfo.asp?ComeFrom=Verify&ChannelID=" & ChannelID & "&Page=" & Page & "&DownStatus=" & KS.G("DownStatus")&"';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr=" & server.URLEncode(KS.C_S(ChannelID,1)&" >> <font color=red>签收会员" & KS.C_S(ChannelID,3)) & "</font>';</script>") 
							 
				       End If
					If KeyWord <> "" Then
						 .Write ("<script> parent.frames['MainFrame'].focus();setTimeout(function(){alert('" & KS.C_S(ChannelID,3) & "修改成功!');location.href='KS.ItemInfo.asp?ChannelID=" & ChannelID & "&Page=" & Page & "&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=DownSearch&OpStr=" & server.URLEncode(KS.C_S(ChannelID,1)&" >> <font color=red>搜索结果</font>")&"';},2500);</script>")
					End If
				End If
			End If
		  End With  		
		End Sub
		
		Sub RefreshHtml(Flag)
			     Dim TempStr,EditStr,AddStr
			    If Flag=1 Then
				  TempStr="添加":EditStr="修改" & KS.C_S(ChannelID,3) & "":AddStr="继续添加" & KS.C_S(ChannelID,3) & ""
				Else
				  TempStr="修改":EditStr="继续修改" & KS.C_S(ChannelID,3) & "":AddStr="添加" & KS.C_S(ChannelID,3) & ""
				End If
			    With Response
				     .Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
					 .Write "<script language='JavaScript' src='../KS_Inc/Jquery.js'></script>"
					 .Write " <Br><br><br><table align='center' width=""95%"" height='200' class='ctable' cellpadding=""1"" cellspacing=""1"">"
					  .Write "	  <tr class=""sort""> "
					  .Write "		<td  height=""28"" colspan=2>系统操作提示信息</td>" & vbcrlf
					  .Write "	  </tr>"
                      .Write "    <tr class='tdbg'>"
					  .Write "          <td align='center'><img src='images/succeed.gif'></td>"
					  .Write "<td><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;恭喜，" & TempStr &"" & KS.C_S(ChannelID,3) & "成功！</b><br>"
					   If Makehtml = 1 Then
					      .Write "<div style=""margin-top:15px;border: #E7E7E7;height:220; overflow: auto; width:100%"">" 
					    If KS.C_S(ChannelID,7)=1 Or KS.C_S(ChannelID,7)=2 Then
						  	 .Write "<div><iframe src=""Include/RefreshHtmlSave.Asp?ChannelID=" & ChannelID &"&Types=Content&RefreshFlag=ID&ID=" & RS("ID") &""" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
						  Else
						  .Write "<div style=""height:25px""><li>由于" & KS.C_S(ChannelID,1) & "没有启用生成HTML的功能，所以ID号为 <font color=red>" & RS("ID") & "</font>  的" & KS.C_S(ChannelID,3) & "没有生成!</li></div> "
						  End If
						  
							If KS.C_S(ChannelID,7)<>1 Then
							  .Write "<div style=""height:25px""><li>由于" & KS.C_S(ChannelID,1) & "的栏目页没有启用生成HTML的功能，所以ID号为 <font color=red>" & TID & "</font>  的栏目没有生成!</li></div> "
							Else
							 If KS.C_S(ChannelID,9)<>1 Then
								  Dim FolderIDArr:FolderIDArr=Split(left(KS.C_C(Tid,8),Len(KS.C_C(Tid,8))-1),",")
								  For I=0 To Ubound(FolderIDArr)
								  .Write "<div align=center><iframe src=""Include/RefreshHtmlSave.Asp?ChannelID=" & ChannelID &"&Types=Folder&RefreshFlag=ID&FolderID=" & FolderIDArr(i) &""" width=""100%"" height=""90"" frameborder=""0"" allowtransparency='true'></iframe></div>"
								   Next
							 End If
						   End If
					   If Split(KS.Setting(5),".")(1)="asp"or KS.C_S(ChannelID,9)<>3 Then
					    '.Write "<div style=""margin-left:140;color:blue;height:25px""><li>由于 <a href=""" & KS.GetDomain & """ target=""_blank""><font color=red>网站首页</font></a> 没有启用生成HTML的功能或发布选项没有开启，所以没有生成!</li></div>"
					   Else
					     .Write "<div align=center><iframe src=""Include/RefreshIndex.asp?RefreshFlag=Info"" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
					   End If
					   .Write "</div>"
					   End If
					  .Write   "</td></tr>"
					  .Write "	  <tr class='tdbg'>"
					  .Write "		<td colspan=2 height=""25"" style='text-align:right'>【<a href=""#"" onclick=""location.href='KS.Down.asp?ChannelID=" & ChannelID &"&Page=" & Page & "&Action=Edit&KeyWord=" & KeyWord &"&SearchType=" & SearchType &"&StartDate=" & StartDate & "&EndDate=" & EndDate &"&ID=" &RS("ID") & "';""><strong>" & EditStr &"</strong></a>】&nbsp;【<a href=""#"" onclick=""location.href='KS.Down.asp?ChannelID=" & ChannelID &"&Action=Add&FolderID=" & Tid & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr="+server.URLEncode("添加" & KS.C_S(ChannelID,3)) & "&ButtonSymbol=AddInfo&FolderID=" & Tid & "';""><strong>" & AddStr & "</strong></a>】&nbsp;【<a href=""#"" onclick=""location.href='KS.ItemInfo.asp?ID=" & Tid & "&ChannelID=" & ChannelID &"&Page=" & Page&"&keyword=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & Tid & "';""><strong>" & KS.C_S(ChannelID,3) & "管理</strong></a>】&nbsp;【<a href=""" & KS.GetDomain & "Item/Show.asp?m=" & ChannelID & "&d=" & RS("ID") & """ target=""_blank""><strong>预览" & KS.C_S(ChannelID,3) & "内容</strong></a>】</td>"
					  .Write "	  </tr>"
					  .Write "	</table>"				
			End With
		End Sub
		Private Function SelDownServer(intdownid)
			Dim rsobj,SQL
			intdownid = KS.ChkClng(intdownid)
			SelDownServer= "<option value=""0"""
			If intdownid = 0 Then SelDownServer=SelDownServer & " selected"
			SelDownServer=SelDownServer & ">↓不使用下载服务器↓</option>"
			SQL = "SELECT downid,DownloadName,depth,rootid FROM KS_DownSer WHERE depth=0 And ChannelID="& ChannelID
			Set rsobj = conn.Execute(SQL)
			Do While Not rsobj.EOF
				SelDownServer=SelDownServer & "<option value=""" & rsobj("downid") & """"
				If intdownid = rsobj("downid") Then SelDownServer=SelDownServer & " selected"
				SelDownServer=SelDownServer & ">" & rsobj(1) & "</option>"
				rsobj.movenext
			Loop
			rsobj.Close:Set rsobj = Nothing
		End Function

End Class
%> 
