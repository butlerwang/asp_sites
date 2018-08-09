<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<%
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim KS:Set KS=New PublicCls

Dim TaskXML,TaskNode,Node,N,TaskUrl,Taskid,Action
set TaskXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
TaskXML.async = false
TaskXML.setProperty "ServerHTTPRequest", true 
TaskXML.load(Server.MapPath(KS.Setting(3)&"Config/task.xml"))
Set TaskNode=TaskXML.DocumentElement.SelectNodes("item[@isenable=1]")


Action=Request.QueryString("Action")
Select Case Action
  case "manage"Manage
  case "add","modify" add
  case "DoSave" DoSave
  case "ModifySave" ModifySave
  case "del" del
  case "taskitem" taskitem
  case else
    Call Task()
End Select

Sub taskitem()
  Dim tasktype:tasktype=KS.ChkClng(KS.G("tasktype"))
  Dim SQLStr,RS,selectid
  selectid=Request("selectid")
  Select Case TaskType
     case 1
	  SQLStr="Select ItemID,ItemName From KS_CollectItem Order By ItemID desc"
	  Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open SQLStr,KS.ConnItem,1,1
	  KS.Echo Escape("<br/><strong>选择要定时采集的项目</strong><br/>")
	  KS.Echo "<select name=""taskid"" id=""taskid"" size=10 multiple style=""width:240px"">"
	  Do While Not RS.EOF
		  If KS.FoundInArr(selectid,RS(0),",") Then
		   KS.Echo Escape("<option value=""" & RS(0) & """ selected>" & RS(1) & "</option>")
		  Else
		   KS.Echo Escape("<option value=""" & RS(0) & """>" & RS(1) & "</option>")
		  End If
	   RS.MoveNext
	  Loop
	  KS.Echo "</select>"
	  KS.Echo Escape("<br/><font color=red>可以按住ctrl键进行多选</font><br/>")
	  RS.Close
	  Set RS=Nothing
	 Case 2,3
	  KS.Echo Escape("<br/><strong>选择要定时发布的栏目</strong><br/>")
	  KS.Echo "<select name=""taskid"" id=""taskid"" size=10 multiple style=""width:240px"">"
	   Dim i,Str,IDArr:IDArr=Split(selectid,",")
	   Str=KS.LoadClassOption(0,false)
	   For I=0 To Ubound(IDArr)
	    str=Replace(str,"value='" & IDArr(i) & "'","value='" & IDArr(i) &"' selected")
	   Next
	  KS.Echo Escape(str)
	  KS.Echo "</select>"
	  KS.Echo Escape("<br/><font color=red>可以按住ctrl键进行多选</font><br/>")
      If TaskType=3 Then
	   KS.Echo Escape("限定最新添加的<input type=""text"" name=""limitnum"" size=""4"" value=""50"" style=""text-align:center"">篇文档")
	  End If	 
  End Select
End Sub


Sub manage()
%>
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
<title>定时任务管理</title>
<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>
<script language="JavaScript" src="../KS_Inc/Jquery.js"></script>
<script LANGUAGE="JavaScript"> 
<!-- 
function openwin() { 
window.open ("KS.Task.asp", "newwindow", "height=450, width=550, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no, location=no, status=no")
} 
--> 
</script>
</head>
<body>
<ul id='mt'> <div id='mtl'>定时任务管理：</div><li><a href="?action=add"><img src="images/ico/as.gif" border='0' align='absmiddle'>添加任务</a></li></ul>
	  <table width="100%" align='center' border="0" cellpadding="0" cellspacing="0">
      <tr class="sort">
	    <td>任务ID</td>
	    <td>任务名称</td>
		<td>任务类型</td>
		<td>执行周期</td>
		<td>执行时间</td>
		<td>状 态</td>
		<td>管理操作</td>
	  </tr>
<%
  If TaskXML.DocumentElement.SelectNodes("item").length=0 Then
      Response.Write "<tr class='list'><td colspan=7 height='25' class='splittd' align='center'>您没有添加定时任务!</td></tr>"
  Else
	  N=0
	  For Each Node In TaskXML.DocumentElement.SelectNodes("item")
	  %>
			  <tr  onmouseout="this.className='list'" onMouseOver="this.className='listmouseover'">               
			   <td class='splittd' height="30" align="center"><%=Node.SelectSingleNode("@id").text%></td>
			   <td class='splittd' height="30"><%=Node.SelectSingleNode("name").text%></td>
			   <td class='splittd' align="center">
			   <%
			   Select Case Node.SelectSingleNode("tasktype").text
				case "1" Response.write "采集"
				case "0" Response.Write "生成首页"
				case "2" Response.Write "生成栏目页"
				case "3" Response.Write "生成内容页"
			   End Select
			   %>
			   </td>
			   <td class='splittd' align="center">
				<font color=red>
			   <%
				if Node.SelectSingleNode("starttype").text="1" then
				 response.write "每天"
				ElseIf Node.SelectSingleNode("starttype").text="3" then
				 response.write "时间段"
				Else
				 response.write "每周 "
				 Select Case  Node.SelectSingleNode("week").text
				  case 0 response.write "星期日"
				  case 1 response.write "星期一"
				  case 2 response.write "星期二"
				  case 3 response.write "星期三"
				  case 4 response.write "星期四"
				  case 5 response.write "星期五"
				  case 6 response.write "星期六"
				 End Select
				End If
				%>
				</font>
			   </td>
			   <td class='splittd' align="center">
			   <%
			    If Node.SelectSingleNode("starttype").text="3" then
			     response.write " " & KS.GotTopic(Node.SelectSingleNode("time").text,45) &"..."
			    else
			     response.write " " & Node.SelectSingleNode("time").text
				End If
			   %></td>
			   <td align="center" class="splittd">
				<%
				 if node.selectSingleNode("@isenable").text="1" then
				  response.write "<font color=blue>开启</font>"
				 else
				  response.write "<font color=green>关闭</font>"
				 end if
				%>
			   </td>
			   <td class='splittd' align="center">
				 <a href="?action=modify&itemid=<%=Node.SelectSingleNode("@id").text%>">修改</a> | <a href="?action=del&itemid=<%=Node.SelectSingleNode("@id").text%>" onClick="return(confirm('确定删除该任务吗?'))">删除</a>
			   </td>
			  </tr>
	  <%
		n=n+1
	  Next
  End If
  %>
		
	  </table>
       <br/>
	   <div style="text-align:center">
	    <input name="Submit" type="button" onClick="openwin()" class="button" value="开始启动">
	   </div>
</body>
</html>
<%
End Sub

Sub Add()
 Dim ItemID:ItemID=KS.ChkClng(Request("ItemID"))
 Dim Node,TaskName,TaskType,StartType,time,week,taskid,limitnum,remark,Isenable,act,ChannelID
 Isenable=1
 starttype=1
 ChannelID=1
 limitnum=50
 act="DoSave"
 If ItemID<>0 Then
   Set Node=TaskXML.DocumentElement.SelectSingleNode("item[@id=" & ItemID & "]")
   If Not Node Is Nothing Then
    Isenable=Node.getAttribute("isenable")
    TaskName=Node.childNodes(0).text
	StartType=Node.childNodes(1).text
	week=Node.childNodes(2).text
	time=Node.childNodes(3).text
    TaskType=Node.childNodes(4).text
	taskid=Node.childNodes(5).text
	limitnum=Node.childNodes(6).text
	remark=Node.childNodes(7).text
	channelid=Node.childNodes(8).text
	Act="ModifySave"
   End If
 End If
%>
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
<title>定时任务管理</title>
<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>
<script language="JavaScript" src="../KS_Inc/Jquery.js"></script>
<script>
 $(document).ready(function()
 {
   $("#starttype").change(function()
   {
     if($(this).val()==2)
	  {
	   $("#weekarea").show();
	  }else{
	   $("#weekarea").hide();
	  }
	  if($(this).val()==3)
	  {
	   $("#time").attr("multiple","multiple");
	   $("#time").attr("style","width:200px;height:150px");
	  }else{
	   $("#time").removeAttr("multiple");
	   $("#time").removeAttr("style");
	  }
	  
	  
   });
   
   $("#tasktype").change(function(){
     getTaskID()
   });
   
   $("#channelid").change(function(){
     getClass();
   });
   
   <%if itemid<>0 and tasktype="1" then%>
    getTaskID();
   <%end if%>

   
 });
 
 function getClass()
 {
      $.get('../plus/ajaxs.asp',{action:'GetClassOption',channelid:$("#channelid option:selected").val()},function(data){
	     $("#typearea").html('<br/><b>选择要定时发布的栏目</b><br/><select name="taskid" id="taskid" size=10 multiple style="width:240px"></select><br/><font color=red>可以按住ctrl键进行多选</font><br/>限定最新添加的<input type="text" name="limitnum" size="4" value="50" style="text-align:center">篇文档');
	     $("#taskid").empty();
		 $('#taskid').append(unescape(data));
	  })
 
 }
 
 function getTaskID()
 {
    if ($("#tasktype option:selected").val()==3)
	{
	 $("#channelarea").show();
	 <%If itemid=0 then%>
	  $("#channelid option[value=0]").attr("selected",true);
	 <%end if%>
	}else{
	 $("#channelarea").hide();
	}
	

 
    if ($("#tasktype option:selected").val()!=undefined && $("#tasktype option:selected").val()!=0&& $("#tasktype option:selected").val()!=3 && $("#tasktype option:selected").val()!='')
	{
	   $.get("KS.Task.asp",{action:"taskitem",tasktype:$("#tasktype option:selected").val(),selectid:"<%=taskid%>"},function(r){
	     $("#typearea").html(unescape(r));
	   });
	 }else{
	 $("#typearea").html('');
	 }
 }
 
 function CheckForm()
 {
 
   if ($("#TaskName").val()=='')
   {
     alert('请输入任务名称!');
	 $("#TaskName").focus();
	 return false
   }
   if ($("#tasktype").val()=='')
   {
     alert('请选择可执行的任务!');
	 $("#tasktype").focus();
	 return false;
   }
   if ($("#tasktype").val()!=0)
   { 
      if ($("#taskid option:selected").val()=='' || $("#taskid option:selected").val()==undefined)
	  {
	   if ($("#tasktype").val()==1)
	   alert('请选择采集项目!');
	   else
	   alert('请选择栏目!');
	   return false;
	  }
   }
   return true;
 }
</script>
<body>
<div class='topdashed sort'>添加/编辑定时任务</div>
<br/>
   <form name="myform" action="KS.Task.asp?action=<%=Act%>" method="post" id="myform">
	  <table width='90%' style="margin:4px" align="center" BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>任务名称:</strong></td>
		   <td><input type="text" name="TaskName" id="TaskName" value="<%=TaskName%>"> 如:定时生成首页</td>
		  </tr>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>可执行任务:</strong></td>
		   <td>
		   <select name="tasktype" id="tasktype">
		     <option value="">--选择任务--</option>
		     <option value="1"<%if tasktype="1" then response.write " selected"%>>定时采集</option>
		     <option value="0"<%if tasktype="0" then response.write " selected"%>>生成首页</option>
		     <option value="2"<%if tasktype="2" then response.write " selected"%>>生成栏目页</option>
		     <option value="3"<%if tasktype="3" then response.write " selected"%>>生成内容页</option>
		   </select>
		   
		    <span id="channelarea"<%if tasktype<>"3" then%> style="display:None"<%end if%>>
		    <strong>选择模型</strong><select id='channelid' name='channelid'>
			<option value='0'>---请选择模型---</option>
			<%
			If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,MNode
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each MNode In ModelXML.documentElement.SelectNodes("channel")
			 if MNode.SelectSingleNode("@ks21").text="1" and MNode.SelectSingleNode("@ks0").text<>"6" and MNode.SelectSingleNode("@ks0").text<>"9" and MNode.SelectSingleNode("@ks0").text<>"10" And MNode.SelectSingleNode("@ks7").text<>"0" Then
			  If Trim(ChannelID)=Trim(MNode.SelectSingleNode("@ks0").text) Then
			  KS.echo "<option value='" &MNode.SelectSingleNode("@ks0").text &"' selected>" & MNode.SelectSingleNode("@ks1").text & "</option>"
			  Else
			  KS.echo "<option value='" &MNode.SelectSingleNode("@ks0").text &"'>" & MNode.SelectSingleNode("@ks1").text & "</option>"
			  End If
			 End If
			next
			
			%>
			</select>
			</span>
			
		   <div id="typearea">
		   <%if tasktype="3" or tasktype="2" then%>
		    <br/><b>选择要定时发布的栏目</b><br/><select name="taskid" id="taskid" size=10 multiple style="width:240px">
			<%
			   Dim Str,IDArr:IDArr=Split(taskid,",")
			   if tasktype="3" then
			   Str=KS.LoadClassOption(ChannelID,false)
			   else
			   Str=KS.LoadClassOption(0,false)
			   end if
			   For I=0 To Ubound(IDArr)
				str=Replace(str,"value='" & IDArr(i) & "'","value='" & IDArr(i) &"' selected")
			   Next
			   KS.Echo str
			
			%>
			</select><br/><font color=red>可以按住ctrl键进行多选</font>
			<%if tasktype="3" then%>
			<br/>限定最新添加的<input type="text" name="limitnum" class="textbox" size="4" value="<%=limitnum%>" style="text-align:center">篇文档
			<%end if%>
		   <%end if%>
		   
		   </div>
		   
		   </td>
		  </tr>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>执行周期:</strong></td>
		   <td>
		   <select name="starttype" id="starttype">
		     <option value="1"<%if starttype="1" then response.write " selected"%>>每天</option>
		     <option value="2"<%if starttype="2" then response.write " selected"%>>每周</option>
		     <option value="3"<%if starttype="3" then response.write " selected"%>>按时间段</option>
		   </select>
		   <span id="weekarea"<%if starttype<>"2" then response.write " style='display:none'"%>>
		    <select name="week" id="week">
			 <option value="0"<%if week="0" then response.write " selected"%>>星期日</option>
			 <option value="1"<%if week="1" then response.write " selected"%>>星期一</option>
			 <option value="2"<%if week="2" then response.write " selected"%>>星期二</option>
			 <option value="3"<%if week="3" then response.write " selected"%>>星期三</option>
			 <option value="4"<%if week="4" then response.write " selected"%>>星期四</option>
			 <option value="5"<%if week="5" then response.write " selected"%>>星期五</option>
			 <option value="6"<%if week="6" then response.write " selected"%>>星期六</option>
			</select> 
		   </span>
		   </td>
		  </tr>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>执行时间:</strong></td>
		   <td>
		   <%if starttype="3" then%>
		    <select name="time" id="time" style="width:200px;height:150px" multiple>
		   <%else%>
		    <select name="time" id="time">
		   <%end if%>
			<%dim i,Ta,Time_S : Time_S=CDate("00:00")
			 for i=1 to 144
			    Ta=Split(Time_S,":")
				If KS.FoundInArr(Time,Ta(0) & ":" & Ta(1),",") Then
				 Response.Write "<option value="""& Ta(0) & ":" & Ta(1) &""" selected>"& Ta(0) & "点" & Ta(1) &"分</option>"
				Else
				 Response.Write "<option value="""& Ta(0) & ":" & Ta(1) &""">"& Ta(0) & "点" & Ta(1) &"分</option>"
				End If
			    Time_S = CDate(Time_S) + CDate("00:10")
			 next  
			 %>
			</select>
			
		   </td>
		  </tr>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>是否启用:</strong></td>
		   <td>
		    <input type="radio" name="Isenable" value="0"<%if Isenable="0" then response.write " checked"%>>不启用
		    <input type="radio" name="Isenable" value="1"<%if Isenable="1" then response.write " checked"%>>启用
		   </td>
		  </tr>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>简要说明:</strong></td>
		   <td>
		    <textarea name="remark" style="width:350px;height:80px" class="textbox"><%=Remark%></textarea>
		   </td>
		  </tr>
      </table>

        <br/>
		<div style="text-align:center">
		 <Input type="hidden" value="<%=itemid%>" name="itemid">
		 <input type="submit" value="保存设置" class="button" onClick="return(CheckForm())">
		</div>
		</form>
</body>
</html>
<%
End Sub

'保存
Sub DoSave()
 Dim TaskName:TaskName=KS.G("TaskName")
 Dim tasktype:tasktype=KS.ChkClng(Request.Form("tasktype"))
 Dim starttype:starttype=KS.ChkClng(Request.Form("starttype"))
 Dim week:week=KS.ChkClng(Request.Form("week"))
 Dim time:time=replace(request.form("time")," ","")
 Dim Isenable:Isenable=KS.ChkClng(Request.Form("Isenable"))
 Dim ChannelID:ChannelID=KS.ChkClng(Request.Form("channelid"))
 Dim TaskID:TaskID=Replace(Request.Form("TaskID")," ","")
 Dim limitnum:limitnum=KS.ChkClng(Request.Form("limitnum"))
 Dim Remark:Remark=Request.Form("Remark")
 
 Dim ItemID
 '取得唯一任务ID号
 If TaskXML.DocumentElement.SelectNodes("item").length<>0 Then
   ItemID=TaskXML.DocumentElement.SelectNodes("item").length+1
 Else
   ItemID=1
 End If
 
 Dim NodeStr,brstr
     brstr=chr(13)&chr(10)&chr(9)
     NodeStr="<item isenable=""" & IsEnable & """ id=""" & ItemID &""">" &brstr
	 NodeStr=NodeStr & "<name>" & TaskName & "</name>"&brstr
	 NodeStr=NodeStr & "<starttype>" & StartType & "</starttype>" &brstr
	 NodeStr=NodeStr & "<week>" & Week & "</week>"&brstr
	 NodeStr=NodeStr & "<time>" & Time & "</time>"&brstr
	 NodeStr=NodeStr & "<tasktype>" & TaskType & "</tasktype>"&brstr
	 NodeStr=NodeStr & "<taskid>" & TaskID & "</taskid>"&brstr
	 NodeStr=NodeStr & "<limitnum>" & limitnum & "</limitnum>" & brstr
	 NodeStr=NodeStr & "<remark><![CDATA[ " & Remark & "]]></remark>" & brstr
	 NodeStr=NodeStr & "<channelid>" & ChannelID &"</channelid>" & brstr
	 NodeStr=NodeStr & " </item>"&brstr
	 Dim XML2:set XML2 = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
     XML2.LoadXml(NodeStr)
	 Dim NewNode:set NewNode=XML2.documentElement
	 
	 Dim TN:Set TN=TaskXML.DocumentElement
	 TN.appendChild(NewNode)
	 TaskXML.Save(Server.MapPath(KS.Setting(3)&"Config/task.xml"))
	 Response.Write "<script>if (confirm('恭喜,定时任务添加成功!')){location.href='?action=add'}else{location.href='?action=manage'}</script>"
End Sub

'保存修改
Sub ModifySave()
 Dim TaskName:TaskName=KS.G("TaskName")
 Dim tasktype:tasktype=KS.ChkClng(Request.Form("tasktype"))
 Dim starttype:starttype=KS.ChkClng(Request.Form("starttype"))
 Dim week:week=KS.ChkClng(Request.Form("week"))
 Dim time:time=replace(request.form("time")," ","")
 Dim Isenable:Isenable=KS.ChkClng(Request.Form("Isenable"))
 Dim TaskID:TaskID=Replace(Request.Form("TaskID")," ","")
 Dim limitnum:limitnum=KS.ChkClng(Request.Form("limitnum"))
 Dim Remark:Remark=Request.Form("Remark")
 Dim ItemID:ItemID=KS.ChkClng(Request.Form("ItemID"))
 Dim ChannelID:ChannelID=KS.ChkClng(Request.Form("channelid"))
 Dim Node
 Set Node=TaskXML.DocumentElement.SelectSingleNode("item[@id=" & ItemID & "]")
 Node.Attributes.getNamedItem("isenable").text=isenable
 Node.childnodes(0).text=TaskName
 Node.childNodes(1).text=StartType
 Node.childNodes(2).text=week
 Node.childNodes(3).text=time
 Node.childNodes(4).text=TaskType
 Node.childNodes(5).text=taskid
 Node.childNodes(6).text=limitnum
 Node.childNodes(7).text=remark
 Node.childNodes(8).text=channelid
	 TaskXML.Save(Server.MapPath(KS.Setting(3)&"Config/task.xml"))
	 Response.Write "<script>alert('恭喜,定时任务修改成功!');location.href='?action=manage'</script>"
End Sub

Sub Del()
  Dim ItemID:ItemID=KS.ChkClng(Request("itemid"))
  If ItemID=0 Then KS.AlertHintScript "对不起,参数出错!"
  Dim DelNode,Node,ID
  Set DelNode=TaskXML.DocumentElement.SelectSingleNode("item[@id=" & ItemID & "]")
  If DelNode Is Nothing  Then
   KS.AlertHintScript "对不起,参数出错!"
  End If
  TaskXML.DocumentElement.RemoveChild(DelNode)
  
  '更新比当前任务ID大的ID号,依次减一
  For Each Node In TaskXML.DocumentElement.SelectNodes("item")
     ID=KS.ChkClng(Node.SelectSingleNode("@id").text)
	 If ID>ItemID Then
	    Node.SelectSingleNode("@id").text=ID-1
	 End If
  Next
  '保存
  TaskXML.Save(Server.MapPath(KS.Setting(3)&"Config/task.xml"))
  KS.AlertHintScript "恭喜,定时任务已删除!"
End Sub

Sub Task()
%>
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
<title>定时任务监控中...请不要关闭本窗口!!!</title>
<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>
<script language="JavaScript" src="../KS_Inc/Jquery.js"></script>
<script> 
$(document).ready(function(){ 
 Timer()
});
var itemLen=<%=TaskNode.length%>;
var taskItem = new Array();
var taskUrl = new Array();
var taskTime = new Array();
var taskStartType = new Array();
var taskWeek = new Array();
<%
N=0
For Each Node In TaskNode
  taskid=Node.SelectSingleNode("taskid").text
  Select Case Node.SelectSingleNode("tasktype").text
    case "0" TaskUrl="Include/RefreshIndex.asp?f=task" 
    case "1" TaskUrl="Collect/Collect_ItemCollection.asp?f=task&Action=Start&CollecType=1&itemid=" & taskid
    case "2" TaskUrl="Include/RefreshHtmlSave.Asp?f=task&Types=Folder&RefreshFlag=IDS&ID=" & taskid
    case "3" TaskUrl="Include/RefreshHtmlSave.asp?f=task&Types=Content&RefreshFlag=Folder&TotalNum=" &Node.SelectSingleNode("limitnum").text & "&ChannelID="& Node.SelectSingleNode("channelid").text & "&FolderID=" & "'" & replace(taskid,",","','") & "'"
  End Select
  
  Response.Write "taskItem[" & n & "]='" & Node.SelectSingleNode("name").text &"';" &vbcrlf
  Response.Write "taskUrl[" & n & "]=""" & TaskUrl &""";" &vbcrlf
  Response.Write "taskTime[" & n & "]='" & Node.SelectSingleNode("time").text &"';" &vbcrlf
  Response.Write "taskStartType[" & n & "]='" & Node.SelectSingleNode("starttype").text &"';" &vbcrlf
  Response.Write "taskWeek[" & n & "]='" & Node.SelectSingleNode("week").text &"';" &vbcrlf
  N=N+1
Next

%>
function timeClock(){ 
	var today=new Date();
	var year =today.getYear();
	var month=today.getMonth()+1;
	var day=today.getDate();
	var h = today.getHours();
	var m = today.getMinutes();
	var s = today.getSeconds();
	var endTime=year+'-'+month+'-'+day+' '+h+":"+m+":"+s;
	$("#currTime").html(endTime);
	
	
	//检测时间
	for(var i=0;i<taskItem.length;i++)
	{
	   //倒计时
	    var djs;
	    if (taskStartType[i]==1)
		{
		  djs=year+"/"+month+"/"+day+" "+taskTime[i];
		}else{
		  djs=year+"/"+month+"/"+day+" "+taskTime[i];
		}
	   
	    BirthDay=new Date(djs);//改成你的计时日期
		today=new Date();
		timeold=(BirthDay.getTime()-today.getTime());
		sectimeold=timeold/1000
		secondsold=Math.floor(sectimeold);
		msPerDay=24*60*60*1000
		e_daysold=timeold/msPerDay
		daysold=Math.floor(e_daysold);
		e_hrsold=(e_daysold-daysold)*24;
		hrsold=Math.floor(e_hrsold);
		e_minsold=(e_hrsold-hrsold)*60;
		minsold=Math.floor((e_hrsold-hrsold)*60);
		seconds=Math.floor((e_minsold-minsold)*60);
		
	   var beginTime=today.getYear()+'-'+(today.getMonth()+1)+'-'+today.getDate()+' '+taskTime[i];
	   var ct= comptime(beginTime,endTime)
		if (taskStartType[i]==1){
		$("#lea"+i).html(hrsold+"小时"+minsold+"分"+seconds+"秒");
		}else if(taskStartType[i]==3){
		
		}
		else{
		  var leaday=taskWeek[i]-today.getDay();
		  if (ct>=0)
		  { if (leaday==0)
		     leaday=6;
			else
			 leaday=leaday-1;
		  }
		  if (leaday<0) leaday=leaday+6;
		$("#lea"+i).html(leaday+"天"+hrsold+"小时"+minsold+"分"+seconds+"秒");
		}

	   
	   
	   //检测执行
	 if(taskStartType[i]==3){
	     var harr=taskTime[i].split(',');
		 for(var k=0;k<harr.length;k++)
		 {
		    var beginTime=today.getYear()+'-'+(today.getMonth()+1)+'-'+today.getDate()+' '+harr[k];
			var ct= comptime(beginTime,endTime)
			if (ct==0)
		    { 
			 window.open(taskUrl[i]); 
		    }
		 }
	  }else{
		   var beginTime=today.getYear()+'-'+(today.getMonth()+1)+'-'+today.getDate()+' '+taskTime[i];
		   var ct= comptime(beginTime,endTime)
		   if (ct==0)
		   { 
			 if (parseInt(taskStartType[i])==1 || (parseInt(taskStartType[i])==2 && today.getDay()==parseInt(taskWeek[i]))){
			 window.open(taskUrl[i]);
			 }
		   }
	  }
	}
} 

//注意：在js中实现按时调用必须是这种方式——定时调用自己
function Timer()
{
timeClock();
setTimeout("Timer()", 1000); // 循环定时调用 
}

//比较时间 格式 yyyy-mm-dd hh:mi:ss
function comptime(beginTime,endTime){ 
var beginTimes=beginTime.split(' ')[0].split('-');
var endTimes=endTime.split(' ')[0].split('-');
beginTime=beginTimes[1]+'-'+beginTimes[2]+'-'+beginTimes[0]+' '+beginTime.split(' ')[1];
endTime=endTimes[1]+'-'+endTimes[2]+'-'+endTimes[0]+' '+endTime.split(' ')[1];
var a =(Date.parse(endTime)-Date.parse(beginTime))/3600/1000;
return a;
}

</script>
</head>
<body>
<br/>
<br/>

  <table width='98%' align="center" BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" align='center'><strong>共有 <font color=red><%=TaskNode.Length%></font> 个定时任务,当前时间是:<span id='currTime' style="color:green"></span></strong></td>
		  </tr>
  <table>
  <%
  N=0
  For Each Node In TaskNode
  %>
	  <table width='98%' style="table-layout:fixed;margin:4px" align="center" BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>任务名称:</strong></td>
		   <td><%=Node.SelectSingleNode("name").text%></td>
		  </tr>
		  <tr class='tdbg'>               
		   <td class='clefttitle'  width="100" align='right'><strong>执行时间:</strong></td>
		   <td style="word-wrap:break-word;">
		   <font color=red>
		   <%
		    if Node.SelectSingleNode("starttype").text="1" then
			 response.write "每天"
			ElseIf Node.SelectSingleNode("starttype").text="3" then
			 response.write "指定以下时间段"
			Else
		     response.write "每周 "
			 Select Case  Node.SelectSingleNode("week").text
			  case 0 response.write "星期日"
			  case 1 response.write "星期一"
			  case 2 response.write "星期二"
			  case 3 response.write "星期三"
			  case 4 response.write "星期四"
			  case 5 response.write "星期五"
			  case 6 response.write "星期六"
			 End Select
			End If
			
			 response.write " " & Node.SelectSingleNode("time").text
			%>
			执行
			</font>
			<%if Node.SelectSingleNode("starttype").text="3" then%>
			
			<%else%>
			 离执行时间还剩:
			<%end if%><span id="lea<%=N%>" style='color:blue'></span>
			</td>
		  </tr>
	  </table>
  <%
    n=n+1
  Next
  %>
</body>
</html>
<%
End Sub


Set KS=Nothing
CloseConn
%>