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
If Not KS.ReturnPowerResult(0, "KSMB10002") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
End iF
Dim TableXML,Node,N,TaskUrl,Taskid,Action
'Set TableXML=LFCls.GetXMLFromFile("task")
set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
TableXML.async = false
TableXML.setProperty "ServerHTTPRequest", true 
TableXML.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))

Action=Request.QueryString("Action")
Select Case Action
  case "DoSave" DoSave
  case "ModifySave" ModifySave
  case "del" del
  case else
    Manage
End Select


Sub manage()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
<title>论坛数据表管理</title>
<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>
<script language="JavaScript" src="../KS_Inc/Jquery.js"></script>
</head>
<body>
<ul id='mt'> <div id='mtl'>论坛数据表管理</div></ul>
	  <table width="100%" align='center' border="0" cellpadding="0" cellspacing="0">
	  <form name="myform" action="KS.GuestTable.asp?action=ModifySave" method="post">
      <tr class="sort">
	    <td>序号</td>
	    <td>表名称</td>
	    <td>类型</td>
		<td>当前默认</td>
		<td>记录数</td>
		<td>说明</td>
		<td>管理操作</td>
	  </tr>
<%
  If TableXML.DocumentElement.SelectNodes("item").length=0 Then
      Response.Write "<tr class='list'><td colspan=7 height='25' class='splittd' align='center'>您没有添加小论坛数据表!</td></tr>"
  Else
	  N=0
	  For Each Node In TableXML.DocumentElement.SelectNodes("item")
	  %>
			  <tr  onmouseout="this.className='list'" onMouseOver="this.className='listmouseover'">               
			   <td class='splittd' height="30" align="center"><%=Node.SelectSingleNode("@id").text%></td>
			   <td class='splittd' height="30"><%=Node.SelectSingleNode("tablename").text%></td>
			   <td class='splittd' style='text-align:center' height="30"><%
			   if Node.SelectSingleNode("@issys").text="1" then
			    response.write "<span style='color:red'>系统</span>"
			   else
			    response.write "<span style='color:green'>自定义</span>"
			   end if
			   %></td>
			   <td class='splittd' align="center">
			   <%
				 if node.selectSingleNode("@isdefault").text="1" then
				  response.write "<input type='radio' name='isdefault' value='" & Node.SelectSingleNode("@id").text & "' checked>"
				 else
				  response.write "<input type='radio' name='isdefault' value='" & Node.SelectSingleNode("@id").text & "'>"
				 end if
				%>
			   </td>
			   <td class='splittd' align="center">
			   <%
			     dim num
				 num=conn.execute("select count(1) from " & Node.SelectSingleNode("tablename").text)(0)
				 response.write "<font color='#ff6600'>" & num & "</font>"
			   %>
			   </td>
			   <td class='splittd' align="center">
			   <%=Node.SelectSingleNode("descript").text%>
			   </td>
			   
			   <td class='splittd' align="center">
			    <%if node.selectSingleNode("@isdefault").text="1" or num>0 or Node.SelectSingleNode("@issys").text="1" then%>
				 <span style="color:#999999">删除</span>
				<%else%>
				 <a href="?action=del&itemid=<%=Node.SelectSingleNode("@id").text%>" onClick="return(confirm('确定删除该任务吗?'))">删除</a>
				<%end if%>
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
	    <input name="Submit" type="submit"  class="button" value="批量设置">
		
	   </div>
	 </form>
	   <br/>
	   
	   <script type="text/javascript">
	    function check(){
		 var tobj=$("#TableName");
		 if (tobj.val()==''){
		  alert('请输入数据表名!');
		  tobj.focus();
		  return false;
		 }
		 if (tobj.val().toLowerCase().indexOf('ks_guest_')==-1){
		  alert('数据表名必须与KS_Guest_开头!');
		  tobj.focus();
		  return false;
		 }
		 return true;
		}
	   </script>
	  
</body>
</html>
<%
End Sub



Set KS=Nothing
CloseConn
%>