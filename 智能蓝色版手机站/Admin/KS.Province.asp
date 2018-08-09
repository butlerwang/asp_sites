<!--#include file="../Conn.asp" -->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp" -->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="include/session.asp" -->
<% 
Dim KSCls
Set KSCls = New Admin_Province
KSCls.Kesion()
Set KSCls = Nothing
Class Admin_Province
        Private KS,KSCls,TypeId,TypeName,News,Action
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
			If Not KS.ReturnPowerResult(0, "KMST10017") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
			End if
 %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
<link href="include/admin_style.css" rel="stylesheet" type="text/css">
<script src="../KS_Inc/common.js"></script>
<script src="../KS_Inc/jQuery.js"></script>
			<script language="javascript">
			    function set(v){
				 if (v==1)
				 AreaControl(1);
				 else if (v==2)
				 AreaControl(2);
				}
				function AreaAdd(){ new parent.KesionPopup().PopupCenterIframe('新增地区','KS.Province.asp?Action=add&parentid=<%=ks.s("parentid")%>',630,200,'no');}
				function EditArea(id){
					 new parent.KesionPopup().PopupCenterIframe('编辑地区','KS.Province.asp?Action=add&ID='+id,630,200,'no')
				}
				function DelArea(id){
				if (confirm('真的要删除该地区吗?'))
				 $('form[name=myform]').submit();
				}
				function AreaControl(op)
				{  var alertmsg='';
	               var ids=get_Ids(document.myform);
					if (ids!='')
					 {  if (op==1)
						{
						if (ids.indexOf(',')==-1) 
							EditArea(ids)
						  else alert('一次只能编辑一个地区!')	 
						}	
					  else if (op==2)    
						 DelArea(ids);
					 }
					else 
					 {
					 if (op==1)
					  alertmsg="编辑";
					 else if(op==2)
					  alertmsg="删除"; 
					 else
					  {
					  alertmsg="操作" 
					  }
					 alert('请选择要'+alertmsg+'的地区');
					  }
				}
				function GetKeyDown()
				{ 
				if (event.ctrlKey)
				  switch  (event.keyCode)
				  {  case 90 : location.reload(); break;
					 case 65 : Select(0);break;
					 case 78 : event.keyCode=0;event.returnValue=false; AreaAdd();break;
					 case 69 : event.keyCode=0;event.returnValue=false;AreaControl(1);break;
					 case 68 : AreaControl(2);break;
				   }	
				else	
				 if (event.keyCode==46)AreaControl(2);
				}
			</script>
</head>
<body  onkeydown='GetKeyDown();' onselectstart='return false;'>

<%
Action = KS.S("action")
Select Case Action
 Case "add"
  Call Add_Submit()
 Case "Save"
  Call Add_Submit_Save()
 Case "del"
  Call Del_Submit()
 Case else
  Call Main()
End Select

End Sub
sub main
 Response.Write "<ul id='menu_top'>"
 Response.Write "<li class='parent' onClick=""AreaAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加地区</span></li>"
 Response.Write "<li class='parent' onClick=""AreaControl(1);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>编辑地区</span></li>"
 Response.Write "<li class='parent' onClick=""AreaControl(2);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>删除地区</span></li>"
 Response.Write "</ul>"

%>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <form name='myform' method='Post' action='KS.Province.asp'>
  <input type="hidden" value="del" name="action">
            <tr align="center" bgcolor="#f5f5f5"> 
                <td width="10%" height="25" class="sort">编 号</td>
                <td width="20%" height="25" class="sort">地区名称</td>
                <td width="8%" class="sort">顺序</td>
                <td height="25" class="sort">管理操作</td>
              </tr>
              <% 
Set Rs = Server.CreateObject("ADODB.recordset")
If KS.S("ParentID")<>"" Then
SQL = "Select * From [KS_Province] Where Parentid="& KS.ChkClng(KS.S("ParentID")) & " Order by orderid Asc"
Else
SQL = "Select * From [KS_Province] Where Parentid=0 Order by orderid Asc"
END iF
Rs.Open SQL,Conn,1,1

Rs.Pagesize = 30
Psize       = Rs.PageSize
PCount      = Rs.PageCount
RCount      = Rs.RecordCount

Page = Cint(Request.QueryString("Page"))
If Page < 1 Then
 Page = 1
Elseif Page > PCount Then
 Page = PCount
End if
Thepage = (Page-1)*Psize
If Not Rs.Eof Then Rs.AbsolutePage = Page

For i = 1 to Psize
 If Rs.Eof Then Exit For
 ID     = Rs("ID")
 City   = Rs("City")
 e_City   = Rs("e_City")
 orderid = Rs("orderid")		  
%>
              <tr align="center" onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='u<%=RS("ID")%>' onClick="chk_iddiv('<%=RS("ID")%>')"> 
                <td width="12%" height="25" class="splittd"><input name="id" onClick="chk_iddiv('<%=ID%>')" type='checkbox' id='c<%=ID%>' value='<%=ID%>'></td>
                <td class="splittd"><%= City %></td>
                <td class="splittd"><%= orderid %></td>
                <td class="splittd"><a href="?action=del&ID=<%= ID %>" onClick="return confirm('是否删除该记录');">删除</a> 
                  <a href="javascript:EditArea(<%=id%>)">编辑</a> 
				  <% if rs("parentid")=0 then%>
                  <a href="?parentid=<%= ID %>&City=<%= City %>">下属地区</a> 
				  <%end if%>
                </td>
              </tr>
              <% 
 Rs.Movenext
Next

				  
%>
		  <tr>
		   <td colspan=3>
		   <div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a>
		   
		   <input type="submit" class="button" value="删除选中" onClick="return(confirm('此操作不可逆,确定删除吗?'))">
		    </div>
		   </td>
		            </form>  
 <td colspan=5>
	  
	  <%
	  Call KS.ShowPage(RCount, Psize, "", Page,true,true)
	  %> </td>
	  </tr>
</table>
</body>
</html>
<% 
Rs.Close
Set Rs = Nothing
End Sub

Sub Add_Submit()
Dim City,e_city,parentid,orderid,id,filtertf
If KS.ChkClng(KS.S("ID"))<>0 Then
 Dim RS:Set RS=Conn.Execute("select * from KS_Province where ID=" & KS.ChkClng(KS.S("ID")))
 If Not RS.Eof Then
  ID=rs("id")
  City=rs("City")
  e_City=rs("e_city")
  parentid=rs("parentid")
  orderid=rs("orderid")
  filtertf=rs("filtertf")
 End If
 RS.Close:Set RS=Nothing
Else
 on error resume next
 Parentid=ks.chkclng(ks.s("parentid"))
 if parentid<>0 then
  orderid=conn.execute("select max(orderid) from KS_Province Where ParentID=" & ParentID)(0)+1
 else
 orderid=1
 end if
 filtertf=1
End If
%>
<script language="javascript">
CheckForm=function()
{
if ($('input[name=City]').val()=='')
{alert('请输入地区名称');
$('input[name=City]').focus()
return false;
}
$("form[name=myform]").submit();
}
</script>

		  <table width="100%" border="0" cellspacing="1" cellpadding="0" class="CTable">
              <form action="?action=Save" method="post" name="myform">
			  <input type="hidden" name="ID" value="<%=id%>">
                <tr class="tdbg"> 
                  <td height="25" align="right" class='clefttitle'>所在城份：</td>
                  <td> <select name="parentid" id="parentid">
                      <option value="0">-作为一级省份-</option>
                      <% 
				  SQL = "Select ID,City From [KS_Province] Where Parentid=0 order by orderid"
				  Set Rs = Conn.Execute(SQL)
				  While Not Rs.Eof
				    if trim(parentid)=trim(rs(0)) then 
					 %>
                      <option value="<%= Rs("ID") %>" selected><%= Rs("City") %></option>
                      <% 
				    else
					 %>
                      <option value="<%= Rs("ID") %>"><%= Rs("City") %></option>
                      <% 
					end if
				   Rs.Movenext
				  Wend
				  Rs.Close
				   %>
                    </select> </td>
                </tr>
                <tr class="tdbg"> 
                  <td width="100" height="25" align="right" class='clefttitle'><p>地区名称：</p></td>
                  <td><input name="City" value="<%=City%>" type="text" size="12">
                    (如：北京)</td>
                </tr>
                <tr class="tdbg" style="display:none"> 
                  <td width="100" height="25" align="right" class='clefttitle'>拼音代码：</td>
                  <td><input name="e_city" value="<%=e_city%>" type="text" size="12">
                    (如：beijing)</td>
                </tr>
                <tr class="tdbg"> 
                  <td width="100" height="25" align="right" class='clefttitle'>是否当模型的筛选项：</td>
                  <td>
				  <label><input type="radio" name="filtertf" value="1"<%if filtertf="1" then response.write " checked"%>>是</label>
				  <label><input type="radio" name="filtertf" value="0"<%if filtertf="0" then response.write " checked"%>>否</label>
				  </td>
                </tr>

                <tr class="tdbg">
                  <td height="25" align="right" class='clefttitle'>排列位置：</td>
                  <td><input name="orderid" type="text" id="suppername" value="<%=orderid%>" size="12"></td>
                </tr>
                
              </form>
            </table>
			
<ul id='save'>
<li class='parent' onClick="return(CheckForm())"><span class=child onMouseOver="this.parentNode.className='parent_border'" onMouseOut="this.parentNode.className='parent'"><img src='images/ico/save.gif' border='0' align='absmiddle'>确定保存</span></li>
<li class='parent' onClick="parent.closeWindow()"><span class=child onMouseOver="this.parentNode.className='parent_border'" onMouseOut="this.parentNode.className='parent'"><img src='images/ico/back.gif' border='0' align='absmiddle'>取消返回</span></li>
</ul>

<%
End Sub

Sub Add_Submit_Save()
 Dim Rs,ID,filtertf,e_city,orderid
 ID=KS.ChkClng(KS.S("ID"))
 City = KS.S("City")
 e_City = KS.S("e_City")
 Parentid = KS.S("Parentid")
 orderid  = KS.ChkClng(KS.S("orderid"))
 filtertf = KS.ChkClng(KS.S("filtertf"))
 
 '//检测是否输入类别名称
 If  City = "" Then
  Response.write "<script>alert('必须输入名称！');history.back();</script>"
  Response.End()
 End if
 
  '//检测是否有同名类别名
 Set Rs = Conn.Execute("Select * from [KS_Province] where ID<>" & ID & " and City='"&City&"' and ParentID="&ParentID&"")
 If Not Rs.Eof Then
  Rs.close
  Set Rs = Nothing
  Response.write "<script>alert('该地区已经存在！');history.back();</script>"
  Response.End()
 End if
 Rs.close
 Set Rs = Nothing

 '//插入记录
 If ID=0 Then
 Conn.Execute ("Insert Into [KS_Province] (City,e_City,Parentid,orderid,filtertf) values ('"&City&"','"&e_City&"',"&Parentid&","&orderid&"," & filtertf &")")
 Else
 Conn.Execute ("Update [KS_Province] set City='" & City & "',e_City='" & E_city & "',Parentid=" & ParentID & ",orderid=" & orderid&",filtertf=" & filtertf &" where id="  & ID)
 End If
 closeconn  
 Call KS.CreateAreaCache()
 If Id=0 Then
 Response.write "<script>if (confirm('添加成功,继续添加吗?')){location.href='KS.Province.asp?action=add&parentid=" & parentid & "';}else{top.frames[""MainFrame""].location.reload();top.closeWindow();}</script>"
 ELse
 Response.write "<script>alert('修改成功！');top.frames[""MainFrame""].location.reload();top.closeWindow();</script>"
 end if
 Response.End()
End Sub

'//删除记录
Sub Del_Submit()
 Dim ID
 ID = KS.FilterIDS(KS.S("ID"))
 Conn.Execute("Delete From [KS_Province] Where Parentid in(" & ID & ") or ID in("&ID & ")")
 closeconn
 Call KS.CreateAreaCache()
 KS.AlertHintScript ("恭喜,删除成功!")
End Sub

End Class
 %>