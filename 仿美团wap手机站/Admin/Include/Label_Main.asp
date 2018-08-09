<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<!--#include file="Label/LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New Label_Main
KSCls.Kesion()
Set KSCls = Nothing

Class Label_Main
        Private KS
		Private LabelSql, LabelRS, FolderID, LabelID, ChannelID, Channel, Action
		Private i, totalPut, CurrentPage, LabelType,UPFolderRS, ParentID,ItemName
		Private KeyWord, SearchType, StartDate, EndDate
		'搜索参数集合
		Private SearchParam
		Private MaxPerPage
		Private Row 
		Private Sub Class_Initialize()
		  MaxPerPage = 48
		  Row = 8
		  Set KS=New PublicCls
		  Call KS.DelCahe(KS.SiteSn & "_labellist")
		  Call KS.DelCahe(KS.SiteSn & "_ReplaceFreeLabel")
		  Call KS.DelCahe(KS.SiteSn & "_jslist")
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
			 '采集搜索信息
			KeyWord = KS.G("KeyWord")
			SearchType = KS.G("SearchType")
			StartDate = KS.G("StartDate")
			EndDate = KS.G("EndDate")
			SearchParam = "KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate
			FolderID = KS.G("FolderID"):If FolderID = "" Then FolderID = "0"
			LabelType = KS.G("LabelType"):If LabelType = "" Then LabelType = 0
			If LabelType=7 Then 
			 ItemName="XML文档" 
			ElseIf LabelType=5 Then
			 ItemName="SQL标签" 
			ELSE 
			 ItemName="标签"
			END IF
			If LabelType = 0 Then
				If Not KS.ReturnPowerResult(0, "KMTL10001") Then                '系统函数标签的权限检查
				  Call KS.ReturnErr(1, "")
				  Response.End
				End If
			ElseIf LabelType = 1 Then
				If Not KS.ReturnPowerResult(0, "KMTL10003") Then                '自定义静态标签的权限检查
				  Call KS.ReturnErr(1, "")
				  Response.End
				End If
			ElseIf LabelType = 5 Then
				If Not KS.ReturnPowerResult(0, "KMTL10002") Then                '自定义函数标签的权限检查
				  Call KS.ReturnErr(1, "")
				  Response.End
				End If
			ElseIf LabelType = 6 Then
				If Not KS.ReturnPowerResult(0, "KMTL10010") Then                '自定义函数标签的权限检查
				  Call KS.ReturnErr(1, "")
				  Response.End
				End If
			ElseIf LabelType = 7 Then
				If Not KS.ReturnPowerResult(0, "KMTL10011") Then                '自定义生成xml的权限检查
				  Call KS.ReturnErr(1, "")
				  Response.End
				End If
			End If
			
		If Not IsEmpty(KS.G("page")) And KS.G("page") <> "" Then
			  CurrentPage = CInt(KS.G("page"))
		Else
			  CurrentPage = 1
		End If
		Set UPFolderRS = Conn.Execute("select * from KS_LabelFolder where ID ='" & FolderID & "'")
		If Not UPFolderRS.EOF Then
		 ParentID = UPFolderRS("ParentID")
		End If
		UPFolderRS.Close:Set UPFolderRS = Nothing
		
		Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd"">"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<title>标签列表</title>"
		Response.Write "</head>"
		Response.Write "<link href=""Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"

			Action = KS.G("Action")
			Select Case  Action
			 Case "SetPasteParam"  Call SetPasteParam()
			 Case "PasteSave"	   Call LabelPasteSave()
			 Case "LabelDel"	   Call LabelDel()
			 Case "LabelFolderDel" Call LabelFolderDel()
			 Case "LabelOut"       Call LabelOut()
			 Case "Doexport"	   Call Doexport()
			 Case "LabelIn"		   Call LabelIn()
			 Case "LabelIn2"	   Call LabelIn2()
			 Case "Doimport"	   Call Doimport()
			 Case "CreateXML"      Call FsoXML()
			 Case Else      	   Call LabelMainList()
			End Select
        End Sub 
		
		'生成XML文档
		Sub FsoXML()
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "select * from KS_Label Where LabelType=7",conn,1,1
		  Do While Not RS.Eof
		    Call CreateXML(RS("id"))
		  RS.MoveNext
		  Loop
		  RS.Close
		  Set RS=Nothing
		  KS.AlertHintScript "恭喜，生成所有XML文档成功!"
		End Sub
		
		Sub LabelMainList()	
		Response.Write "<script language=""JavaScript"">" & vbCrLf
		Response.Write "var FolderID='" & FolderID & "';         //目录ID" & vbCrLf
		Response.Write "var ParentID='" & ParentID & "'; //父栏目ID" & vbCrLf
		Response.Write "var Page='" & CurrentPage & "';   //当前页码" & vbCrLf
		Response.Write "var KeyWord='" & KeyWord & "';    //关键字" & vbCrLf
		Response.Write "var SearchParam='" & SearchParam & "';  //搜索参数集合" & vbCrLf
		Response.Write "var LabelType=" & LabelType & ";" & vbCrLf
		Response.Write "</script>" & vbCrLf
		Response.Write "<script language=""JavaScript"" src=""../../ks_inc/Common.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../../ks_inc/jQuery.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../../ks_inc/lhgdialog.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""ContextMenu.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""SelectElement.js""></script>"
		%>
		<script language="javascript">
		var DocElementArrInitialFlag=false;
		var DocElementArr = new Array();
		var DocMenuArr=new Array();
		var SelectedFile='',SelectedFolder='';
		$(document).ready(function(){
			if (DocElementArrInitialFlag) return;
			InitialDocElementArr('FolderID','LabelID');
			InitialDocMenuArr();
			DocElementArrInitialFlag=true;
		})
		function InitialDocMenuArr()
		{  
		   if (KeyWord=='')
			{ 
			 DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.AddFolder();",'新建目录(N)','disabled');
			 DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.AddLabel('');",'新建<%=ItemName%>(M)','disabled');
			 DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			}
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'全 选(A)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Edit('');",'编 辑(E)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Delete('');",'删 除(D)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Paste();",'克 隆(V)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('parent.ChangeUp();','后 退(B)','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Reload('');",'刷 新(Z) ','');
		}
		function DocDisabledContextMenu()
		{ 
		   var TempDisabledStr=''; 
		   if (FolderID=='0') TempDisabledStr='后 退(B),';
			DisabledContextMenu('FolderID','LabelID',TempDisabledStr+'编 辑(E),删 除(D),克 隆(V)','克 隆(V)','','编 辑(E),克 隆(V)','编 辑(E),克 隆(V)','克 隆(V)')
		}
		function ChangeUp()
		{
		 if (FolderID=='0') return;
		 location.href='Label_Main.asp?LabelType='+LabelType+'&FolderID='+ParentID;
		 if (LabelType==0)
			$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> 系统函数标签')+'&ButtonSymbol=FunctionLabel&LabelFolderID='+ParentID;
		 else if(LabelType==5)
		 	$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> 自定义SQL函数标签')+'&ButtonSymbol=DIYFunctionLabel&LabelFolderID='+ParentID;
		 else if(LabelType==7)
		 	$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> 自定义XML标签')+'&ButtonSymbol=DIYFunctionLabel&LabelFolderID='+ParentID;
		 else
		   $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> 自定义静态标签')+'&ButtonSymbol=FreeLabel&LabelFolderID='+ParentID;
		}
		function OpenLabelFolder(FolderID)
		{
			location.href='Label_Main.asp?LabelType='+LabelType+'&FolderID='+FolderID;
		   if (LabelType==0)
			 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> 系统函数标签')+'&ButtonSymbol=FunctionLabel&LabelFolderID='+FolderID;
			else if (LabelType==5)
			 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> 自定义函数标签')+'&ButtonSymbol=DIYFunctionLabel&LabelFolderID='+FolderID;
			else
			 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> 自定义静态标签')+'&ButtonSymbol=FreeLabel&LabelFolderID='+FolderID;
		}
		var box='';
		function AddFolder()
		{
		   box=$.dialog({title:"新建标签目录",content:"url:include/LabelFolder.asp?LabelType="+LabelType+"&FolderID="+FolderID,width:650,height:360});

		}
		function AddLabel(TempUrl)
		{ 
		   if (LabelType==1){
				location.href=TempUrl+'LabelAdd.asp?mode=text&LabelType=1&Action=AddNew&FolderID='+FolderID;
				$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> <font color=red>添加自定义静态标签</font>')+'&ButtonSymbol=LabelAdd';
				}
		   else if(LabelType==5){
				location.href=TempUrl+'LabelSQL.asp?LabelType=5&FolderID='+FolderID;
				$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+('标签管理 >> <font color=red>添加自定义函数标签</font>')+'&ButtonSymbol=DIYFunctionStep1';		
}
		   else if(LabelType==6){
				location.href=TempUrl+'CirLabel.asp?LabelType=6&FolderID='+FolderID;
				$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> <font color=red>添加循环标签</font>')+'&ButtonSymbol=LabelAdd';}
		   else if(LabelType==7){
				location.href=TempUrl+'LabelXML.asp?LabelType=7&FolderID='+FolderID;
				$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+('标签管理 >> <font color=red>生成自定义XML</font>')+'&ButtonSymbol=DIYFunctionStep1';		
}
		   else
			 { 
			    location.href=TempUrl+'AddFunctionLabel.asp?FolderID='+FolderID;
				$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> <font color=red>添加系统函数标签</font>')+'&ButtonSymbol=Go';
			  }
		}
		function EditLabel(TempUrl,id)
		{ if (LabelType==1)
				if (KeyWord=='')
				 {	location.href=TempUrl+'LabelAdd.asp?mode=text&LabelType=1&page='+Page+'&Action=EditLabel&LabelID='+id;
					$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> <font color=red>修改自定义静态标签</font>')+'&ButtonSymbol=LabelAdd';
				 }
				else
				   { location.href=TempUrl+'LabelAdd.asp?mode=text&LabelType=1&page='+Page+'&Action=EditLabel&'+SearchParam+'&LabelID='+id;
					$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> 搜索自定义静态标签结果 >> <font color=red>修改自定义静态标签</font>')+'&ButtonSymbol=LabelAdd';
				   }
			else if(LabelType==5)
			   if (KeyWord=='')
				 {	location.href=TempUrl+'LabelSQL.asp?LabelType=5&page='+Page+'&Action=Edit&LabelID='+id;
					$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> <font color=red>修改自定义函数标签</font>')+'&ButtonSymbol=DIYFUNCTIONSTEP1';
				 }
				else
				   { location.href=TempUrl+'LabelSQL.asp?LabelType=5&page='+Page+'&Action=Edit&'+SearchParam+'&LabelID='+id;
					$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> 搜索自定义函数标签结果 >> <font color=red>修改自定义函数标签</font>')+'&ButtonSymbol=DIYFUNCTIONSTEP1';
					}
		    else if(LabelType==6){
			        location.href=TempUrl+'CirLabel.asp?mode=text&LabelType=1&page='+Page+'&Action=EditLabel&LabelID='+id;
					$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> <font color=red>修改循环标签</font>')+'&ButtonSymbol=LabelAdd';
			}
			else if(LabelType==7)
			   if (KeyWord=='')
				 {	location.href=TempUrl+'LabelXML.asp?LabelType=7&page='+Page+'&Action=Edit&LabelID='+id;
					$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> <font color=red>修改XML文档</font>')+'&ButtonSymbol=DIYFUNCTIONSTEP1';
				 }
				else
				   { location.href=TempUrl+'LabelXML.asp?LabelType=7&page='+Page+'&Action=Edit&'+SearchParam+'&LabelID='+id;
					$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> 搜索XML文档结果 >> <font color=red>修改自定义XML文档</font>')+'&ButtonSymbol=DIYFUNCTIONSTEP1';
					}
			else
			 {	
			 
			 	location.href=TempUrl+'EditFunctionLabel.asp?LabelID='+id;
				$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> <font color=red>修改系统函数标签</font>')+'&ButtonSymbol=GoSave';
			 }
		}
		function AddByText()
		{
		 location.href='LabelAdd.asp?LabelType=1&Action=AddNew&FolderID='+FolderID;
		 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> <font color=red>添加自定义静态标签</font>')+'&ButtonSymbol=LabelAdd';
		}
		function EditByText(id)
		{ 	GetSelectStatus('FolderID','LabelID');
			 if (SelectedFile!='')
					 {
					 if (SelectedFile.indexOf(',')==-1) 
					  {
					  location.href='LabelAdd.asp?LabelType=1&page='+Page+'&Action=EditLabel&'+SearchParam+'&LabelID='+SelectedFile;
			$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('标签管理 >> 搜索自定义静态标签结果 >> <font color=red>修改自定义静态标签</font>')+'&ButtonSymbol=LabelAdd';
					   }
					else alert('一次只能够编辑一个标签');
					}
				 else
				   alert('请选择您要编辑的标签!');
			SelectedFile='';
		}
		function LabelOut(){
		location.href='?Action=LabelOut&LabelType='+LabelType;
		}
		function LabelIn(){
		location.href='?Action=LabelIn&LabelType='+LabelType;
		}
		function CreateXML(){
		location.href='?Action=CreateXML&LabeltYPE='+LabelType;
		}
		function EditFolder(ID){
		 box=$.dialog({title:"编辑标签目录",content:"url:include/LabelFolder.asp?Action=EditFolder&FolderID="+ID,width:650,height:360});

		}
		function Edit(TempUrl)
		{   GetSelectStatus('FolderID','LabelID');
			if (!((SelectedFile=='')&&(SelectedFolder=='')))
				{
					if (SelectedFolder!='')
					{ 
					if (TempUrl=='Folder'||TempUrl=='')
					 if (SelectedFolder.indexOf(',')==-1) 
					  { EditFolder(SelectedFolder);
					 }
					else alert('一次只能够编辑一个标签目录');
				   }
				   if (SelectedFile!='')
					 {
					 if (TempUrl!='Folder'||TempUrl=='')
					 {if (SelectedFile.indexOf(',')==-1) 
					   EditLabel(TempUrl,SelectedFile);
					else alert('一次只能够编辑一个标签');
					 }
					}
				}
			else 
			{
			alert('请选择要编辑的标签或目录');
			}
			SelectedFile='';
			SelectedFolder='';
		}
		function Delete(TempUrl)
		{   GetSelectStatus('FolderID','LabelID');
			if (!((SelectedFile=='')&&(SelectedFolder=='')))
				{  
					if (confirm('删除确认:\n\n真的要执行删除操作吗?'))
					  { if (SelectedFolder!='')
					   if (TempUrl=='Folder'||TempUrl=='')
						 location='Label_Main.asp?Action=LabelFolderDel&ID='+SelectedFolder+'&FolderID='+FolderID+'&LabelType='+LabelType;
					  if (SelectedFile!='')  
						if (TempUrl!='Folder'||TempUrl=='')
						location=TempUrl+'Label_Main.asp?Action=LabelDel&Page='+Page+'&ID='+SelectedFile+'&FolderID='+FolderID+'&LabelType='+LabelType;
					}	
				}
			else alert('请选择要删除的标签目录或标签');
		   SelectedFile='';
		   SelectedFolder='';
		}
		function DelFolder(ID){
		    if (confirm('删除确认:\n\n真的要执行删除标签目录操作吗?')){
			location='Label_Main.asp?Action=LabelFolderDel&ID='+ID+'&FolderID='+FolderID+'&LabelType='+LabelType;
			}
		}
		function DelLabel(ID){
		    if (confirm('删除确认:\n\n真的要执行删除标签操作吗?')){
			location='Label_Main.asp?Action=LabelDel&Page='+Page+'&ID='+ID+'&FolderID='+FolderID+'&LabelType='+LabelType;
			}
		}
		
		function Paste()
		{
		GetSelectStatus('FolderID','LabelID');
			if (SelectedFile!='')  
			  OpenWindow('LabelFrame.asp?Url=Label_Main.asp&Action=SetPasteParam&PageTitle='+escape('请输入新<%=ItemName%>名称')+'&LabelType=<%=LabelType%>&LabelID='+SelectedFile,350,120,window);
			else alert('请选择要克隆的标签');
			SelectedFile='';
			SelectedFolder='';
			Reload('');
		}
		function GetKeyDown()
		{
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 :  Reload(''); break;
			 case 65 : SelectAllElement();break;
			 case 66 : event.keyCode=0;event.returnValue=false;ChangeUp();break;
			 case 78 : event.keyCode=0;event.returnValue=false; AddFolder();break;
			 case 77 : event.keyCode=0;event.returnValue=false; AddLabel('');break;
			 case 69 : event.keyCode=0;event.returnValue=false;Edit('');break;
			 case 86 : event.keyCode=0;event.returnValue=false;Paste();break;
			 case 68 : Delete('');break;
			 case 70 :event.keyCode=0;event.returnValue=false;
			 if (LabelType==0)
				parent.initializeSearch('系统函数标签')
			 else if(LabelType==5)
				parent.initializeSearch('自定义函数标签')
			 else
			   parent.initializeSearch('自定义静态标签')
		 }	
		else if (event.keyCode==46)
		Delete('');
		}
		function Reload(TempUrl)
		{
		location.href=TempUrl+'Label_Main.asp?FolderID='+FolderID+'&page='+Page+'&LabelType='+LabelType+'&'+SearchParam;
		}
		</script>
		<%
		Response.Write "<body  topmargin=""0"" leftmargin=""0"" OnClick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
           Response.Write "<ul id='menu_top'>"			 
			 If KeyWord = "" Then
			  Response.Write "<li class='parent' onclick=""AddLabel('')""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>添加" & ItemName&"</span></li>"
			  Response.Write "<li class='parent' onclick=""AddFolder();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>添加目录</span></li>"
			  Response.Write "<li class='parent' onclick=""Edit('Folder');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/as.gif' border='0' align='absmiddle'>编辑目录</span></li>"
			  Response.Write "<li class='parent' onclick=""Delete('Folder');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/del.gif' border='0' align='absmiddle'>删除目录</span></li>"
			  Response.Write "<li class='parent' onclick=""parent.initializeSearch("
			If LabelType = 0 Then Response.Write ("'系统函数标签'") Else If LabelType=5 Then Response.Write("'自定义函数标签'") Else If LabelType=7 Then Response.Write("'自定义生成XML文档'") Else Response.Write ("'自定义静态标签'")
			  Response.Write ");""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/s.gif' border='0' align='absmiddle'>搜索助理</span></li>"

			If LabelType=1 Then
			  Response.Write "<li class='parent' onclick=""AddByText();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>可视添加</span></li>"
			  Response.Write "<li class='parent' onclick=""EditByText();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/as.gif' border='0' align='absmiddle'>可视编辑</span></li>"
			End If
			
			If LabelType=7 Then
			  Response.Write "<li class='parent' onclick=""CreateXML();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/move.gif' border='0' align='absmiddle'>一键生成所有XML</span></li>"
			Else
			  Response.Write "<li class='parent' onclick=""LabelIn();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/move.gif' border='0' align='absmiddle'>导入</span></li>"
			  Response.Write "<li class='parent' onclick=""LabelOut();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/verify.gif' border='0' align='absmiddle'>导出</span></li>"
		   End If
			  Response.Write "<li class='parent' onclick=""ChangeUp();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/verify.gif' border='0' align='absmiddle'>回上一级</span></li>"

			 
			 Else
				If LabelType = 0 Then
				   Response.Write ("<img src='../Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('Label_Main.asp?LabelType=0','../Template_Left.asp','../KS.Split.asp?ButtonSymbol=FunctionLabel&OpStr=标签管理 >> <font color=red>系统函数标签</font>')"">系统标签首页</span>")
				ElseIf LabelType=5 Then
				  Response.Write ("<img src='../Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('Label_Main.asp?LabelType=5','../Template_Left.asp','../KS.Split.asp?ButtonSymbol=FunctionLabel&OpStr=标签管理 >> <font color=red>自定义函数标签</font>')"">自定义函数标签首页</span>")
				Else
				   Response.Write ("<img src='../Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('Label_Main.asp?LabelType=1','../Template_Left.asp','../KS.Split.asp?ButtonSymbol=FunctionLabel&OpStr=标签管理 >> <font color=red>自定义静态标签</font>')"">自定义静态标签首页</span>")
				End If
			   Response.Write (">>> 搜索结果: ")
				 If StartDate <> "" And EndDate <> "" Then
					Response.Write ("标签更新日期在 <font color=red>" & StartDate & "</font> 至 <font color=red> " & EndDate & "</font>&nbsp;&nbsp;&nbsp;&nbsp;")
				 End If
				Select Case SearchType
				 Case 0
				  Response.Write ("名称含有 <font color=red>" & KeyWord & "</font> 的标签")
				 Case 1
				  Response.Write ("描述中含有 <font color=red>" & KeyWord & "</font> 的标签")
				 Case 2
				  Response.Write ("内容中含有 <font color=red>" & KeyWord & "</font> 的标签")
				 End Select
			 End If
		Response.Write "</ul>"
	
		Response.Write "<div style="" height:98%; overflow: auto; width:100%"" align=""center"">"
		Response.Write "  <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "      <tr>"
		Response.Write "    <td height=""6""></td>"
		Response.Write "      </tr>"
			  
			  Dim FolderSql, Param
			  Param = " Where LabelType=" & LabelType
			If KeyWord <> "" Then
				FolderSql = "SELECT ID,FolderName,Description,OrderID as LabelFlag,OrderID FROM [KS_LabelFolder] where 1=0"
				Select Case SearchType
					Case 0
					  Param = Param & " AND LabelName like '%" & KeyWord & "%'"
					Case 1
					 Param = Param & " AND Description like '%" & KeyWord & "%'"
					Case 2
					 Param = Param & " AND LabelContent like '%" & KeyWord & "%'"
				End Select
				If StartDate <> "" And EndDate <> "" Then
					 Param = Param & " And (AddDate>=#" & StartDate & "# And AddDate<=#" & DateAdd("d", 1, EndDate) & "#)"
			   End If
			Else
			  FolderSql = "SELECT ID,FolderName,Description,OrderID as LabelFlag,OrderID FROM [KS_LabelFolder] where  FolderType=" & LabelType & " And ParentID='" & FolderID & "'"
			  Param = Param & " AND FolderID='" & FolderID & "'"
			End If
			Param = Param & " ORDER BY OrderID ,foldername"
		Set LabelRS = Server.CreateObject("ADODB.recordset")
		LabelRS.Open FolderSql & " UNION all Select ID,LabelName,Description,LabelFlag,OrderID from [KS_Label] " & Param, Conn, 1, 1
		If LabelRS.EOF And LabelRS.BOF Then
		Else
					        totalPut = LabelRS.RecordCount
		
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									LabelRS.Move (CurrentPage - 1) * MaxPerPage
							End If
							Call showContent
		End If
		
		Response.Write "    </table>"
		Response.Write "    </div>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
		 Sub showContent()
		 %>
		 <style>
		 .labellist{}
		 .labellist td{position:relative;}
		 .labellist .m{display:none;position:absolute;right:5px;top:5px}
		 .labellist td.td{border:1px solid #fff;}
		 .labellist td.td:hover{border:1px solid #E4E4E4;background:#FBFDFF;}
		 .labellist td.td:hover .m{display:block;position:absolute;right:5px;top:5px}
		 </style>
		 <%
		 Do While Not LabelRS.EOF
		   Response.Write "<tr>"
		   Response.Write " <td>"
		   Response.Write " <table class=""labellist"" width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		   Response.Write "     <tr>"
				  
				  Dim T, TitleStr, FolderName, ShortName, LabelTypeStr
					   For T = 1 To Row
					   If Not LabelRS.EOF Then
						  If LabelRS(4) = 0 Then
							  FolderName = LabelRS(1)
							  ShortName = KS.ListTitle(FolderName, 24)
							  TitleStr = " TITLE='名 称:" & FolderName & "&#13;&#10;类 型:标签目录'"
						  Else
							  FolderName = LabelRS(1)
							  ShortName = KS.ListTitle(Replace(Replace(Replace(Replace(FolderName, "{LB_", ""), "}", ""),"{SQL_",""),"{XML_",""), 24)
							  If LabelType = 1 Then
								 LabelTypeStr = "自定义静态标签"
							  ElseIf LabelType=5 Then
							     LabelTypeStr = "自定义函数标签"
							  ElseIf LabelType=7 Then
							     LabelTypeStr = "自定义生成XML文档"
							  Else
								 LabelTypeStr = "系统函数标签"
							  End If
							  TitleStr = " TITLE='名 称:" & FolderName & "&#13;&#10;类 型:" & LabelTypeStr & "'"
						  End If
						   Response.Write ("<td class=""td"" width=""" & CInt(100 / Row) & "%"" Style=""cursor:default"" align=""center""" & TitleStr & ">")
						   response.write "<span class=""m"">"
						   If LabelRS(4) = 0 Then
						   response.write "<a href=""javascript:;"" onclick=""EditFolder('" & LabelRS(0) & "')"" style=""color:green;"" title=""修改标签目录"">E</a> <a href=""javascript:;"" onclick=""DelFolder('" & LabelRS(0) & "');"" style=""color:red;""　title=""删除标签目录"">X</a>"
						   Else
						   response.write "<a href=""javascript:;"" onclick=""EditLabel('','" & LabelRS(0) & "')"" style=""color:green;"" title=""修改标签"">E</a> <a href=""javascript:;"" onclick=""DelLabel('" & LabelRS(0) & "');"" style=""color:red;"" title=""删除标签"">X</a>"
						   End If
						   response.write "</span>"
						   
						If LabelRS(4) = 0 Then
						   Response.Write ("<span onmousedown=""mousedown(this);"" FolderID=""" & LabelRS(0) & """ style=""POSITION:relative;"" OnDblClick=""OpenLabelFolder('" & LabelRS(0) & "')"">")
						Else
						   Response.Write ("<span onmousedown=""mousedown(this);"" LabelID=""" & LabelRS(0) & """ style=""POSITION:relative;"" onDblClick=""EditLabel('','" & LabelRS(0) & "');"">")
						End If
					 If LabelRS(4) = 0 Then
					   Response.Write ("<img src=""../Images/Folder/folder.gif""> ")
					 ElseIf LabelType = 1 Then
					   Response.Write ("<img src=""../Images/Label/Label3.gif"">")
					 ElseIF LabelType=5 Then
					  Response.Write ("<img src=""../Images/Label/Label5.gif"">")
					 ElseIF LabelType=7 Then
					  Response.Write ("<img src=""../Images/Label/Label7.gif"">")
					 Else
					   Response.Write ("<img src=""../Images/Label/Label" & LabelRS(3) & ".gif"">")
					 End If
					Response.Write ("<span style=""display:block;height:16;padding:0px 0px 0px 0px;margin:1px;width:80%;cursor:default"">" & ShortName & "</span>")
					Response.Write ("</span>")
					Response.Write ("</td>")
					i = i + 1
					   If LabelRS.EOF Or i >= MaxPerPage Then Exit For
					   LabelRS.MoveNext
					 Else
					  Exit For
					 End If
				Next
				'不到7个单元格,则进行补空
				Do While T <= Row
				 Response.Write ("<td width=70>&nbsp;</td>")
				 T = T + 1
				 Loop
				   
		   Response.Write "     </tr>"
		   Response.Write "     <tr><td colspan=" & Row & " height=10></td></tr>"
		   Response.Write "   </table></td>"
		   Response.Write "</tr>"
		
				  If i >= MaxPerPage Then Exit Do
				  If LabelRS.EOF Then Exit Do
				Loop
				  LabelRS.Close
				  Conn.Close
		
		  Response.Write " <td  align=""right"">"
			 Call KS.ShowPageParamter(totalPut, MaxPerPage, "Label_Main.asp", True, "个", CurrentPage, "LabelType=" & LabelType & "&FolderID=" & FolderID & "&" & SearchParam)
		  Response.Write " </td>"
		  Response.Write "    </tr>"
		End Sub
		
		'克隆标签的名称
		Sub SetPasteParam()
		Dim LabelID:LabelID=KS.G("LabelID")
		Dim NewLabelName,FileName
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select top 1 * From KS_Label Where ID='" & LabelID & "'",conn,1,1
		If Not RS.Eof Then
		 NewLabelName = "复制_" & Replace(Replace(Replace(Replace(RS("LabelName"), "{LB_", ""),"{SQL_",""),"{XML_",""), "}", "")
		 LabelType    = RS("LabelType")
		 FileName     = year(now) & month(now) & day(now)&hour(now)&minute(now)&second(now) &".xml"
		End If
		RS.Close:Set RS=Nothing
		Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd"">"
        Response.Write "<html xmlns=""http://www.w3.org/1999/xhtml"">"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<title>标签类型</title>"
		Response.Write "<link href=""Admin_Style.css"" rel=""stylesheet"">"
		Response.Write "<script language=""JavaScript"" src=""../../ks_inc/jquery.js""></script>"
		Response.Write "</head>"
		Response.Write "<body style=""background: #EAF0F5;"" scroll=no topmargin=""0"" leftmargin=""0"">"
		Response.Write "  <form id=""LabelPasteForm"" method=""post"" action=""?Action=PasteSave"">"
		Response.Write "  <input type=""hidden"" value=""" & LabelID & """ name=""LabelID"">"
		Response.Write "  <table width=""96%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "    <tr>"
		Response.Write "      <td>"
		Response.Write "      <FIELDSET align=center>"
		Response.Write "      <LEGEND align=left>"
        Response.Write "         克隆"& ItemName
		Response.Write "       </LEGEND>"

		Response.Write "  <table width=""100%"" height=""30"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "    <tr>"
		Response.Write "      <td height=""40"">"
		Response.Write "       克隆" & ItemName &"的名称："
		Response.Write "        <input type=""text"" name=""NewLabelName"" id='NewLabelName' size='30' class='textbox' value=""" & NewLabelName & """>"
		Response.Write "        <input type=""hidden"" name=""labelType"" value=""" & Conn.Execute("Select LabelType From KS_Label Where ID='" & LabelID & "'")(0) &""">"
		Response.Write "      </td>"
		Response.Write "    </tr>"
		If LabelType=7 Then
		Response.Write "    <tr>"
		Response.Write "      <td height=""40"">"
		Response.Write "       生成" & ItemName &"文件名：" & KS.Setting(127)
		Response.Write "        <input type=""text"" name=""FileName"" id=""FileName"" size='20' class='textbox' value=""" & FileName & """> <span class='tips'>必须以.xml结束</span>"
		Response.Write "      </td>"
		Response.Write "    </tr>"
		End If
		Response.Write "    <tr>"
		Response.Write "      <td height=""40"" align=""center"">"
		Response.Write "        <input type=""submit"" name=""Submit""  class=""button"" onclick=""return(CheckForm())"" value="" 确 定 "">"
		Response.Write "        <input type=""button"" name=""Submit2""  class=""button"" onclick=""window.close()"" value="" 取 消 "">"
		Response.Write "      </td>"
		Response.Write "    </tr>"
		Response.Write "  </table>"
		Response.Write "          </FIELDSET>"
		Response.Write "          </td></tr></table>"
		Response.Write "  </form>"

		Response.Write "</body>"
		Response.Write "</html>"
		Response.Write "<script Language=""javascript"">" & vbCrLf
		Response.Write "function CheckForm()" & vbCrLf
		Response.Write "{" & vbCrLf
		Response.Write "    if ($('#NewLabelName').val()=='')" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""请给新克隆的标签取个名称!"");" & vbCrLf
		Response.Write "     $('#NewLabelName').focus();" & vbCrLf
		Response.Write "    return false;" & vbCrLf
		Response.Write "    }"
		If LabelType=7 Then
		Response.Write "    if ($('#FileName').val()=='')" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""请输入xml文件名称!"");" & vbCrLf
		Response.Write "     $('#FileName').focus();" & vbCrLf
		Response.Write "    return false;" & vbCrLf
		Response.Write "    }"
		Response.Write "if ($('#FileName').val().toLowerCase().indexOf('.xml')==-1){" & vbcrlf
		Response.Write " alert('XML文件名必须以.xml为扩展名!');$('#FileName').focus(); return false;}" & vbcrlf
		End If
		Response.Write "    return true;" & vbCrLf
		Response.Write "}" & vbCrLf
		Response.Write "</script>"
		End Sub
		'保存克隆
		Sub LabelPasteSave()
		  Dim LabelID:LabelID=KS.G("LabelID")
		  Dim NewLabelName:NewLabelName=KS.G("NewLabelName")
		  If KS.G("LabelType")=5 Then
		  NewLabelName = "{SQL_" & NewLabelName & "}"
		  ElseIf KS.G("LabelType")=7 Then
		  NewLabelName = "{XML_" & NewLabelName & "}"
		  Else
		  NewLabelName = "{LB_" & NewLabelName & "}"
		  End IF
		  Dim FileName:FileName=Request("FileName")
		  If LabelType=7 Then
		    If FileName="" Or Right(Lcase(FileName),4)<>".xml" Then
			   Call KS.AlertHistory("XML文件名必须以.xml结束!", -1)
			   Set KS = Nothing
			   Exit Sub
			End If
          End If
		  Dim LabelRS:Set LabelRS=Server.CreateObject("ADODB.RECORDSET")
		  LabelRS.Open "Select TOP 1 LabelName From KS_Label Where LabelName='" & NewLabelName & "'", Conn, 1, 1
		  If Not LabelRS.Eof Then 
		     LabelRS.Close:Set LabelRS=Nothing
		     Call KS.Alert(ItemName & "名称已存在，请输入其它名称!","Label_Main.asp?LabelType=" & LabelType & "&Action=SetPasteParam&LabelID=" & LabelID)
		  End If
		    LabelRS.Close
			LabelRS.Open "Select top 1 * From KS_Label Where ID='" & LabelID & "'",Conn,1,1
			If Not LabelRS.Eof Then
			    Dim NewRS:Set NewRS=Server.CreateObject("ADODB.RECORDSET")
				NewRS.Open "Select * From KS_Label",Conn,1,3
				NewRS.AddNew
				  NewRS("ID")           = Year(Now()) & KS.MakeRandom(10)
				  NewRS("LabelName")    = NewLabelName
				  NewRS("FileName")     = FileName
				  NewRS("LabelContent") = LabelRS("LabelContent")
				  NewRS("Description") = LabelRS("Description")
				  NewRS("FolderID")    = LabelRS("FolderID")
				  NewRS("OrderID")     = LabelRS("OrderID")
				  NewRS("LabelType")   = LabelRS("LabelType")
				  NewRS("LabelFlag")   = LabelRS("LabelFlag")
				  NewRS("AddDate")     = Now
				  NewRS.Update
				  If LabelType=7 Then
				    CreateXML NewRS("id") 
				  End If
				  NewRS.Close:Set NewRS=Nothing
				  LabelRS.Close:Set LabelRS=Nothing
				  Response.Write "<script>window.close();</script>"
			Else
			  Response.Write "<script>alert('克隆失败!');window.close();</script>"
			End If
		End Sub
		
		'删除标签目录
		Sub LabelFolderDel()
		   Dim RS,K, ID, ParentID, FolderSql,LabelFolderID
		   Set RS=Server.CreateObject("ADODB.Recordset")
		   ID=Request("ID")
		   If Id="" Then KS.Die "error!"
		   ID = Split(ID, ",")     '获得要删除目录的ID集合
			For K = LBound(ID) To UBound(ID)
			  FolderSql = "select id from [KS_LabelFolder] where TS like '%" & ID(K) & ",%'"
			  RS.Open FolderSql, Conn, 1, 1
			If Not RS.Eof Then
			  Do While Not RS.Eof
			    Dim RSI:Set RSI=Conn.Execute("Select id FROM KS_Label WHERE FolderID='" & rs("id") & "' and LabelType=7")
				 Do While NOT RSI.Eof 
				   DelXML RSI(0)
				 RSI.MoveNext
				 Loop
				 RSI.Close:Set RSI=Nothing
				 Conn.Execute("DELETE  FROM KS_Label WHERE FolderID='" & rs("id") & "'")
				 conn.execute("DELETE  FROM [KS_LabelFolder] where id='" &rs("id")&"'")
			   RS.MoveNext
			  Loop
			  RS.Close
			End If  
			
		  Next
		   Set RS=Nothing
		  Response.Write "<script>location.href='Label_Main.asp?LabelType=" & LabelType & "&Folderid=" & ParentID & "'</script>"
		End Sub
		'删除标签
		Sub LabelDel()
			Dim K, ID,Page
			Page = KS.G("Page")
			ID = Split(Request("id"), ",") '获得要删除标签的ID集合
			For K = LBound(ID) To UBound(ID)
			  DelXML ID(K)
			  Conn.Execute("Delete FROM KS_Label WHERE ID='" & ID(K) & "'")
			Next
			Response.Write "<script>location.href='Label_Main.asp?Page=" & Page & "&LabelType=" & LabelType & "&FolderID=" & FolderID & "';</script>"

		End Sub
		
		'删除XML文件
		Sub DelXML(id)
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "select top 1 * From KS_Label Where LabelType=7 and ID='" & ID & "'",conn,1,1
		  If Not RS.EOf Then
		      Dim CreatePath
			  If KS.Setting(127)<>"/" Then
			  CreatePath=KS.Setting(3) & KS.Setting(127)
			  Else
			  CreatePath=KS.Setting(3)
			  End If
			  CreatePath=CreatePath & RS("FileName")
              Call KS.DeleteFile(CreatePath)
		  End If
		  RS.Close :  Set RS=Nothing
		eND Sub 
		
		Sub LabelOut()
		Response.Write "<body>"
		Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""   class=""sortbutton"">"
		Response.Write "  <tr>"
		Response.Write "    <td height=""23"" align=""left"">管理导航：<a href='?Action=LabelIn&LabelType=" & LabelType & "'>标签导入</a> | <a href='?Action=LabelOut&LabelType=" & LabelType & "'>导出功能</a></td>"
		Response.Write " </tr>"
		Response.Write "</table>"

		  LabelType=KS.G("LabelType")
		  %>
		  <Script language="Javascript">
		  var ClassArr = new Array();
		  <%
			Response.Write "ClassArr[0] =new Array(""" & GetLabelOption(0,conn) & """);" & vbcrlf
			Response.Write "ClassArr[1] =new Array(""" & GetLabelOption(1,conn) & """);" & vbcrlf
			Response.Write "ClassArr[5] =new Array(""" & GetLabelOption(5,conn) & """);" & vbcrlf
			Response.Write "ClassArr[9999] =new Array(""" & GetLabelOption(0,conn)&GetLabelOption(1,conn)&GetLabelOption(5,conn) & """);" & vbcrlf
		  %>
		  </Script>
		  <form name='myform' method='post' action='Label_main.asp'>  
		  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>    
		  <tr class='title'>       
		  <td colspan="2" height='22' align='center'><strong>标签导出</strong></td>    
		  </tr>    
		  <tr class='tdbg'>
			  <td width="100" style="text-align:right">选择类型：</td>
			  <td width="820" style="text-align:left">
			  <select id="LabelType" name="LabelType" onChange="SelectClass(this.value)">
			  <option value="9999">全部标签</option>
			  <option value="0"<%IF LabelType="0" Then Response.write " selected"%>>系统函数标签</option>
			  <option value="5"<%IF LabelType="5" Then Response.write " selected"%>>自定义SQL标签</option>
			  <option value="1"<%IF LabelType="1" Then Response.write " selected"%>>自定义静态标签</option>
		    </select>
			</td>
		  </tr>    
		  <tr class='tdbg'>      
		  <td colspan="2" align='center'>        
		    <table width="100%" border='0' cellpadding='0' cellspacing='0'>          
			   <tr>           
			     <td width="90" style="text-align:right">标签列表：</td>
				 <td width="54%" ID="ClassArea"><select name='LabelID' size='2' multiple style='height:300px;width:450px;'>
				 </select></td>                
				  <td align='left'>&nbsp;&nbsp;&nbsp;
				   <input type='button' class="button"  name='Submit' value=' 选定所有 ' onclick='SelectAll()'>    <br><br>&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' class="button" name='Submit' value=' 取消选定 ' onclick='UnSelectAll()'><br><br><br><b>&nbsp;提示：按住“Ctrl”或“Shift”键可以多选</b></td>      
			 </tr>     
			 <tr height='30'>        
			 <td colspan='2'>　目标数据库：
			     <input name='LabelMdb' type='text' class="textbox" id='LabelMdb' value='<%=KS.Setting(3)%>Label.mdb' size='20' maxlength='50'>
			 &nbsp;&nbsp;此操作将清空目标数据库</td>      
			 </tr>      
		    <tr height='50'>        
			 <td colspan='3' style='text-align:center'><input type='submit' class="button" name='Submit' value='导出选中的标签' onClick="document.myform.Action.value='Doexport';">   
			<input type='submit' class="button" name='Submit' value='一键导出所有系统函数标签' onClick="document.myform.ExportType.value='0';">
			<input type='submit' class="button" name='Submit' value='一键导出所有自定义SQL标签' onClick="document.myform.ExportType.value='5';">
			<input type='submit' class="button" name='Submit' value='一键导出所有自定义静态标签' onClick="document.myform.ExportType.value='1';">
			           <input name='Action' type='hidden' id='Action' value='Doexport'> 
					   <input name='ExportType' type='hidden' id='ExportType' value=''>         </td>        </tr>    </table>   
		    </td> </tr></table></form>
		  <script language='javascript'>
		  SelectClass(<%=LabelType%>);
		function SelectClass(LabelType)
		{ document.all.ClassArea.innerHTML='<select name="LabelID" size="2" multiple style="height:300px;width:450px;">'+ClassArr[LabelType]+'</select>';
		}
		function SelectAll(){
		  for(var i=0;i<document.myform.LabelID.length;i++){
			document.myform.LabelID.options[i].selected=true;}
		}
		function UnSelectAll(){
		  for(var i=0;i<document.myform.LabelID.length;i++){
			document.myform.LabelID.options[i].selected=false;}
		}
		</script>
		  <%
		End Sub
		Function GetLabelOption(LabelType,DBC)
		  Dim AllLabel,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select * From KS_Label Where LabelType=" & LabelType,DBC,1,1
		  Do While Not RS.Eof 
			AllLabel=AllLabel & "<option value='" & RS("ID") & "'>" & RS("LabelName") & "</option>"
			RS.MoveNext
		  Loop
          RS.Close:Set RS=Nothing
		  GetLabelOption=AllLabel
		End Function
		'导出操作
		Sub Doexport()
		 Dim LabelID:LabelID="'"& Replace(Replace(KS.G("LabelID")," ",""),",","','") & "'"
		 Dim LabelMdb:LabelMdb=KS.G("LabelMdb")
		 If InStr(lcase(LabelMdb),".asp")>0 or InStr(lcase(LabelMdb),".asa")>0 or InStr(lcase(LabelMdb),".php")>0 or InStr(lcase(LabelMdb),".cer")>0 or InStr(lcase(LabelMdb),".cdx")>0 or right("00000"&lcase(LabelMdb),4)<>".mdb" Then
			Call KS.AlertHistory("导出数据库文件名格式不正确，数据库扩展名必段是.mdb!", -1)
			Set KS = Nothing:Response.End
		 End If
		 

		 Dim rs:set rs=server.createobject("adodb.recordset")
		 Dim sqlstr,n
		   n=0
		 If Request("ExportType")<>"" Then
		   sqlstr="select ID,LabelName,LabelContent,Description,FolderID,OrderID,LabelType,LabelFlag,AddDate,FileName from ks_label Where LabelType=" & KS.ChkClng(request("ExportType"))
		 Else
		   sqlstr="select ID,LabelName,LabelContent,Description,FolderID,OrderID,LabelType,LabelFlag,AddDate,FileName from ks_label where id in(" & LabelID & ")"
		 End if
		         'on error resume next
			     if CreateDatabase(LabelMdb)=true then
						Dim DataConn:Set DataConn = Server.CreateObject("ADODB.Connection")
	                    DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(LabelMdb)
						If not Err Then
						   If Checktable("KS_Label",DataConn)=true Then
						     DataConn.Execute("drop table KS_Label")
						   end if
				             Dataconn.execute("CREATE TABLE [KS_Label] ([LabelID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,[ID] varchar(50) Not Null,[LabelName] varchar(255) Not Null,[LabelContent] text not null,[Description] text null,[FolderID] varchar(100) not null,[OrderID] int not null,[LabelType] int not null,[LabelFlag] int not null,[AddDate] date not null,[FileName] varchar(255))")
						  rs.open sqlstr,conn,1,1
						 if not rs.eof then
						   	Dim RST:Set RST=Server.CreateObject("ADODB.RECORDSET")
						   do while not rs.eof
							  n=n+1
						      'DataConn.Execute("Insert Into KS_Label(ID,LabelName,LabelContent,Description,FolderID,OrderID,LabelType,LabelFlag,AddDate) values('" & rs(0) & "','" & rs(1) & "','" &rs(2) & "','" & rs(3) & "','" & rs(4) & "'," & rs(5) & "," & rs(6) & "," & rs(7) & ",'" & rs(8) & "')")
							  RST.Open "Select * From KS_Label where 1=0",DataConn,1,3
							  RST.AddNew
							    RST("ID")=rs(0)
								RST("LabelName")=rs(1)
								RST("LabelContent")=rs(2)
								RST("Description")=rs(3)
								RST("FolderID")=rs(4)
								RST("OrderID")=rs(5)
								RST("LabelType")=rs(6)
								RST("LabelFlag")=rs(7)
								RST("AddDate")=rs(8)
								RST("FileName")=rs(9)
							  RST.Update
							  RST.Close
							  rs.movenext
						   loop
						   Set RST=Nothing
						 end if
                          rs.close:set rs=nothing
						End if
						DataConn.Close:Set DataConn=Nothing
				 end if
				response.write "<br><br><br><div align=center>操作完成!成功导出了 <font color=red>" & n & "</font> 个标签！<a href=" & LabelMdb & ">请点击这里下载</a>(右键目标另存为) <input type='button' value=' 返回 ' class='button' onclick=""history.back();""/> </div><br><br><br><br><br><br><br>"

		End Sub
		
		Sub LabelIn()
		Response.Write "<body>"
		Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""   class=""sortbutton"">"
		Response.Write "  <tr>"
		Response.Write "    <td height=""23"" align=""left"">管理导航：<a href='?Action=LabelIn&LabelType=" & LabelType & "'>标签导入</a> | <a href='?Action=LabelOut&LabelType=" & LabelType & "'>导出功能</a></td>"
		Response.Write " </tr>"
		Response.Write "</table>"

		%>
		<form name='myform' method='post' action='Label_Main.asp?LabelType=<%=KS.G("LabelType")%>'>  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>    <tr class='title'>       <td height='22' align='center'><strong>标签导入（第一步）</strong></td>    </tr>   
		 <tr class='tdbg'>      
		 <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;请输入要导入的标签数据库的文件名： <input name='LabelMdb' class="textbox" type='text' id='LabelMdb' value='<%=KS.Setting(3)%>Label.mdb' size='20' maxlength='50'>  
		 <input name='Submit' class="button" type='submit' id='Submit' value=' 下一步 '>        <input name='Action' type='hidden' id='Action' value='LabelIn2'>      </td>    </tr>  </table></form>
		<%
		End Sub
		
		Sub LabelIn2()
		Response.Write "<body>"
		Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""   class=""sortbutton"">"
		Response.Write "  <tr>"
		Response.Write "    <td height=""23"" align=""left"">管理导航：<a href='?Action=LabelIn&LabelType=" & LabelType & "'>标签导入</a> | <a href='?Action=LabelOut&LabelType=" & LabelType & "'>导出功能</a></td>"
		Response.Write " </tr>"
		Response.Write "</table>"

		on error resume next
		LabelType=KS.G("LabelType")
		Dim LabelMdb:LabelMdb=KS.G("LabelMdb")
		Dim DataConn:Set DataConn = Server.CreateObject("ADODB.Connection")
	    DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(LabelMdb)
		%>
		<form name='myform' method='post' action='Label_Main.asp'>  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>    <tr class='title'>       <td height='22' align='center'><strong>标签导入（第二步）</strong></td>    </tr>    <tr class='tdbg'>       <td height='100' align='center'>        <br>        <table border='0' cellspacing='0' cellpadding='0'>          
		<%
		If Err Then 
		Err.Clear:Set DataConn = Nothing:Response.Write "<tr><td>数据库路径不正确，连接出错</td></tr>":Response.End
		else
		 	%>
		  <Script language="Javascript">
		  var ClassArr = new Array();
		  <%
			Response.Write "ClassArr[0] =new Array(""" & GetLabelOption(0,DataConn) & """);" & vbcrlf
			Response.Write "ClassArr[1] =new Array(""" & GetLabelOption(1,DataConn) & """);" & vbcrlf
			Response.Write "ClassArr[5] =new Array(""" & GetLabelOption(5,DataConn) & """);" & vbcrlf
			Response.Write "ClassArr[9999] =new Array(""" & GetLabelOption(0,DataConn)&GetLabelOption(1,DataConn)&GetLabelOption(5,DataConn) & """);" & vbcrlf
		  %>
		  </Script>
		<tr> <td><strong>选择要导入的标签的分类：</strong><select disabled="disabled" id="LabelType" name="LabelType" onChange="SelectClass(this.value)">
			  <option value="9999">全部标签</option>
			  <option value="0"<%IF LabelType="0" Then Response.write " selected"%>>系统函数标签</option>
			  <option value="5"<%IF LabelType="5" Then Response.write " selected"%>>自定义函数标签</option>
			  <option value="1"<%IF LabelType="1" Then Response.write " selected"%>>自定义静态标签</option>
		    </select></td></tr>   
  		<tr>
		<td height="30"><strong>重名处理方式：</strong> 
		<input type="radio" value="2" name="cl" checked>标签重名自动重命名导入
		<input type="radio" value="0" name="cl">标签重名跳过
		<input type="radio" value="1" name="cl">标签重名覆盖
		</td>
		</tr>  

		<tr>
		<td id="ClassArea"> 
		<select name='LabelID' size='2' multiple style='height:300px;width:350px;'> </select>
		</td>
		</tr>  
		<%end if%>                <tr><td colspan='3' height='5'></td></tr>                  <tr>                    <td height='25' align='center'><b> 提示：按住“Ctrl”或“Shift”键可以多选</b></td>                  </tr>    <tr><td colspan='3' height='25' align='center'>
		
		导入到的目录：<%=ReturnLabelFolderTree("0", LabelType)%>
		
		<input type='submit' name='Submit' class='button' value=' 导入选中的标签 ' onClick="document.myform.Action.value='Doimport';" >      
		   <input type='submit' name='Submit' class='button' value=' 全部导入 ' onClick="document.myform.ExportType.value='<%=LabelType%>';" >       </td></tr>               </table>               <input name='LabelMdb' type='hidden' id='LabelMdb' value='<%=LabelMdb%>'>               <input name='Action' type='hidden' id='Action' value='Doimport'>   <input name='ExportType' type='hidden' id='ExportType' value=''>              <br>            </td>          </tr>       
		</table></form>
		<script language='javascript'>
		  SelectClass(<%=LabelType%>);
		function SelectClass(LabelType)
		{ document.all.ClassArea.innerHTML='<select name="LabelID" size="2" multiple style="height:300px;width:350px;">'+ClassArr[LabelType]+'</select>';
		}
   </script>
		<%
		dataconn.close:set dataconn=nothing
		End Sub
		'导入操作
		Sub Doimport()
			'on error resume next
			Dim n:n=0
			Dim m:m=0
			Dim k:k=0
			Dim t:t=0
			Dim LabelMdb:LabelMdb=KS.G("LabelMdb")
			Dim NewLabelID,cl:cl=KS.G("cl")
			Dim ClassID:ClassID=KS.G("ParentID")
			Dim LabelID:LabelID="'"& Replace(Replace(KS.G("LabelID")," ",""),",","','")& "'"
			Dim DataConn:Set DataConn = Server.CreateObject("ADODB.Connection")
			DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(LabelMdb)
			If Err Then 
			Err.Clear:Set DataConn = Nothing:Response.Write "<tr><td>数据库路径不正确，连接出错</td></tr>":Response.End
			else
			 Dim rs:set rs=server.createobject("adodb.recordset")
			 if request("ExportType")<>"" then
			 rs.open "select * from ks_label where labeltype=" & KS.ChkClng(request("ExportType")),dataconn,1,1
			 else
			 rs.open "select * from ks_label where ID in(" & LabelID & ")",dataconn,1,1
			 end if
			 Dim rsa:set rsa=server.createobject("adodb.recordset")
			 do while not rs.eof 
			  rsa.open "select * from ks_label where labelname='" & rs("labelname") & "'",conn,1,3
			  if rsa.eof then
			     rsa.addnew
				  Do While True
					'生成ID  年+10位随机
					NewLabelID = Year(Now()) & KS.MakeRandom(10)
					Dim RSCheck:Set RSCheck = Conn.Execute("Select ID from [KS_Label] Where ID='" & NewLabelID & "'")
					 If RSCheck.EOF And RSCheck.BOF Then
					  RSCheck.Close:Set RSCheck = Nothing:Exit Do
					 End If
				  Loop
			     rsa("ID")=NewLabelID
				 rsa("LabelName")=rs("LabelName")
				 rsa("LabelContent")=rs("LabelContent")
				 rsa("Description")=rs("Description")
				' rsa("FolderID")=rs("folderid")
				 rsa("FolderID")=ClassID
				 rsa("OrderID")=rs("OrderID")
				 rsa("LabelType")=rs("LabelType")
				 rsa("LabelFlag")=rs("LabelFlag")
				 rsa("AddDate")=rs("AddDate")
				 rsa("FileName")=rs("filename")
				 n=n+1
				rsa.update
			  else   '重名处理
			   if cl="1" then
				 rsa("LabelContent")=rs("LabelContent")
				 rsa("Description")=rs("Description")
				 rsa("OrderID")=rs("OrderID")
				 rsa("LabelType")=rs("LabelType")
				 rsa("LabelFlag")=rs("LabelFlag")
				 rsa("AddDate")=rs("AddDate")
				 rsa("FileName")=rs("filename")
				 m=m+1
				rsa.update
			   elseif cl=2 then  '重名自动命名
			     Do While True
					'生成ID  年+10位随机
					NewLabelID = Year(Now()) & KS.MakeRandom(10)
					Set RSCheck = Conn.Execute("Select ID from [KS_Label] Where ID='" & NewLabelID & "'")
					 If RSCheck.EOF And RSCheck.BOF Then
					  RSCheck.Close:Set RSCheck = Nothing:Exit Do
					 End If
				  Loop
				 rsa.addnew
			     rsa("ID")=NewLabelID
				 rsa("LabelContent")=rs("LabelContent")
				 rsa("Description")=rs("Description")
				 rsa("OrderID")=rs("OrderID")
				 rsa("LabelType")=rs("LabelType")
				 rsa("LabelFlag")=rs("LabelFlag")
				 rsa("AddDate")=rs("AddDate")
				 rsa("FileName")=rs("filename")
				 rsa("FolderID")=ClassID
				 rsa("LabelName")=replace(rs("LabelName"),"}","_new}")
				 t=t+1
				 rsa.update
			   else
			    k=K+1
			   end if
			  end if
			   rsa.close
			  rs.movenext
			 loop
			 rs.close:set rs=nothing
			 set rsa=nothing
			end if
			response.write "<br><br><br><div align=center>操作完成!成功导入了 <font color=red>" & n & "</font> 个标签,覆盖了 <font color=red>" & m & "</font> 个标签,重命名了 <font color=red>" & t & "</font> 个标签，重名跳过了 <font color=red>" & k & "</font> 个标签！  </div><br><br><br><br><br><br><br>"
           dataconn.close:set dataconn=nothing
		End Sub
		Function CreateDatabase(dbname)
		      if KS.CheckFile(dbname) then CreateDatabase=true:exit function
				dim objcreate :set objcreate=Server.CreateObject("adox.catalog") 
				if err.number<>0 then 
					set objcreate=nothing 
					CreateDatabase=false
					exit function 
				end if 
				'建立数据库 
				objcreate.create("data source="+server.mappath(dbname)+";provider=microsoft.jet.oledb.4.0") 
				if err.number<>0 then 
					CreateDatabase=false
					set objcreate=nothing 
					exit function
				end if 
				CreateDatabase=true
		End Function
		'检查数据表是否存在	
		Function Checktable(TableName,DataConn)
			On Error Resume Next
			DataConn.Execute("select * From " & TableName)
			If Err.Number <> 0 Then
				Err.Clear()
				Checktable = False
			Else
				Checktable = True
			End If
		End Function

End Class
%> 
