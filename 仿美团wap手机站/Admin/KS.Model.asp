<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Model
KSCls.LoadKesion()
Set KSCls = Nothing


Class Admin_Model
        Private KS,KSCls,I
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
        %>
		<!--#include file="../ks_cls/UserFunction.asp"-->
		<%

		Public Sub LoadKesion()
		  'If Not KS.ReturnPowerResult(0, "model1") Then          '检查权限
			'Call KS.ReturnErr(1, "")
			'.End
		 ' End If
		 If KS.G("Action")="createtemplate" Then
			  response.cachecontrol="no-cache"
			  response.addHeader "pragma","no-cache"
			  response.expires=-1
			  response.expiresAbsolute=now-1
			  Response.CharSet="utf-8"
			  Dim KSUser,ChannelID,FieldXML,FieldNode
			  ChannelID=KS.ChkClng(KS.S("ChannelID"))
			  Call KSUser.LoadModelField(ChannelID,FieldXML,FieldNode)
			  Set KSUser=New UserCls
			  Call GetInputForm(true,ChannelID,FieldXML,FieldNode,"",0,KSUser,"")
			  Set KSUser=Nothing
			  Response.End()
		 End If
		  With Response
		    .Write "<html>"
			.Write "<title>模型基本参数设置</title>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			.Write "<script src=""../ks_inc/JQuery.js"" language=""JavaScript""></script>"
			.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "</head>"
			.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.Write "<ul id='menu_top'>"
			.Write "<li class='parent' onclick=""location.href='?action=Add';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Go&OpStr=系统设置 >> <font color=red>系统模型管理</font>';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加模型</span></li>"
			.Write "<li class='parent' onclick=""location.href='KS.Model.asp?action=total';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/set.gif' border='0' align='absmiddle'>信息统计</span></li>"
			If KS.G("Action")="" Then
			.Write "<li class='parent' disabled><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>管理首页</span></li>"
			Else
			.Write "<li class='parent' onclick=""location.href='KS.Model.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>管理首页</span></li>"
			End IF
			.Write "</ul>"

		  Select Case KS.G("Action")
		   Case "SetChannelParam"
				If Not KS.ReturnPowerResult(0, "KSMM10005") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		             Call SetChannelParam()
			    End If 
		   Case "Edit","Add"
		       If KS.G("Action")="Add" Then
		       If Not KS.ReturnPowerResult(0, "KSMM10000") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		            Call ChannelAddOrEdit()
			    End If
			  Else
		       If Not KS.ReturnPowerResult(0, "KSMM10001") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		            Call ChannelAddOrEdit()
			    End If
			  End If
		   Case "EditSave"
		        Call ChannelSave()
		   Case "Del"
		       If Not KS.ReturnPowerResult(0, "KSMM10002") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		          Call ChannelDel()
			    End If
		   Case "SetSearch"
		        Call SetSearch()
		   Case "ManageMenu"
		       If Not KS.ReturnPowerResult(0, "KSMM10002") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		          Call ManageMenu()
			    End If
		   Case "total"
		        If Not KS.ReturnPowerResult(0, "KSMM10004") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		          Call Total()
			    End If
		   Case Else
		       Call Main()
		  End Select
		  End With
		End Sub
		
		Sub SetSearch()
		 Dim AllowField:AllowField="'title','author','origin','keywords','intro','area'"
		 Dim ChannelID:channelid=KS.ChkClng(Request("channelid"))
		 Dim RS,FieldXML,XMLStr,Node,TemplateFile,isrewrite,maxperpage
		 Dim tj,check,xsz,ssz,title
		 If ChannelID=0 Then KS.Die "error!"
		 If Request("flag")="dosave" Then
		      dim ctid:ctid=KS.ChkClng(Request("Ctid"))
			  XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
			  XMLStr=XMLStr&"<field>" &vbcrlf
			  XMLStr=XMLStr&" <template><![CDATA[" & Request("TemplateFile") &"]]></template>" &vbcrlf
			  XMLStr=XMLStr&" <isrewrite>" & KS.ChkClng(Request("isrewrite")) &"</isrewrite>" &vbcrlf
			  XMLStr=XMLStr&" <maxperpage>" & KS.ChkClng(Request("maxperpage")) &"</maxperpage>" &vbcrlf
			  if ctid=1 then
			    XMLStr=XMLStr&" <item name=""tid"" enabled=""true"">" &vbcrlf
			  Else
			    XMLStr=XMLStr&" <item name=""tid"" enabled=""false"">" &vbcrlf
			  End If
			    XMLStr=XMLStr&"  <title>" & request("titletid") & "</title>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldname>tid</fieldname>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldtype>-1</fieldtype>" &vbcrlf
			    XMLStr=XMLStr&"  <condition>dy</condition>" &vbcrlf
			    XMLStr=XMLStr&"  <showvalue>0</showvalue>" &vbcrlf
			    XMLStr=XMLStr&"  <searchvalue>0</searchvalue>" &vbcrlf
			    XMLStr=XMLStr&" </item>" &vbcrlf
			  '商城模块
			 If KS.ChkClng(KS.C_S(ChannelID,6))=5 Then
				  If KS.ChkClng(Request("cprice_member"))=1 Then
					 XMLStr=XMLStr&" <item name=""price_member"" enabled=""true"">" &vbcrlf
				  Else
					 XMLStr=XMLStr&" <item name=""price_member"" enabled=""false"">" &vbcrlf
				  End If
			    XMLStr=XMLStr&"  <title>" & request("titleprice_member") & "</title>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldname>price_member</fieldname>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldtype>-1</fieldtype>" &vbcrlf
			    XMLStr=XMLStr&"  <condition>fw</condition>" &vbcrlf
			    XMLStr=XMLStr&"  <showvalue>" & request("xszprice_member") &"</showvalue>" &vbcrlf
			    XMLStr=XMLStr&"  <searchvalue>" & request("sszprice_member") &"</searchvalue>" &vbcrlf
			    XMLStr=XMLStr&" </item>" &vbcrlf
				  If KS.ChkClng(Request("cbrandid"))=1 Then
					 XMLStr=XMLStr&" <item name=""brandid"" enabled=""true"">" &vbcrlf
				  Else
					 XMLStr=XMLStr&" <item name=""brandid"" enabled=""false"">" &vbcrlf
				  End If
			    XMLStr=XMLStr&"  <title>" & request("titlebrandid") & "</title>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldname>brandid</fieldname>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldtype>-1</fieldtype>" &vbcrlf
			    XMLStr=XMLStr&"  <condition>dys</condition>" &vbcrlf
			    XMLStr=XMLStr&"  <showvalue>0</showvalue>" &vbcrlf
			    XMLStr=XMLStr&"  <searchvalue>0</searchvalue>" &vbcrlf
			    XMLStr=XMLStr&" </item>" &vbcrlf
			 ElseIf KS.ChkClng(KS.C_S(ChannelID,6))=8 Then
				  If KS.ChkClng(Request("ctypeid"))=1 Then
					 XMLStr=XMLStr&" <item name=""typeid"" enabled=""true"">" &vbcrlf
				  Else
					 XMLStr=XMLStr&" <item name=""typeid"" enabled=""false"">" &vbcrlf
				  End If
			    XMLStr=XMLStr&"  <title>" & request("titletypeid") & "</title>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldname>typeid</fieldname>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldtype>-1</fieldtype>" &vbcrlf
			    XMLStr=XMLStr&"  <condition>dys</condition>" &vbcrlf
			    XMLStr=XMLStr&"  <showvalue>0</showvalue>" &vbcrlf
			    XMLStr=XMLStr&"  <searchvalue>0</searchvalue>" &vbcrlf
			    XMLStr=XMLStr&" </item>" &vbcrlf
				  If KS.ChkClng(Request("cbrandid"))=1 Then
					 XMLStr=XMLStr&" <item name=""brandid"" enabled=""true"">" &vbcrlf
				  Else
					 XMLStr=XMLStr&" <item name=""brandid"" enabled=""false"">" &vbcrlf
				  End If
			    XMLStr=XMLStr&"  <title>" & request("titlebrandid") & "</title>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldname>brandid</fieldname>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldtype>-1</fieldtype>" &vbcrlf
			    XMLStr=XMLStr&"  <condition>dys</condition>" &vbcrlf
			    XMLStr=XMLStr&"  <showvalue>0</showvalue>" &vbcrlf
			    XMLStr=XMLStr&"  <searchvalue>0</searchvalue>" &vbcrlf
			    XMLStr=XMLStr&" </item>" &vbcrlf
			 End If	
			  Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "select * from ks_field Where (fieldname in(" & AllowField & ") or fieldtype<>0) and ChannelID=" & channelid & " order by orderid,fieldid",conn,1,1
			  If Not RS.Eof Then
					Do While Not RS.Eof
						 If KS.ChkClng(Request("c" & rs("fieldname")))=1 Then
							XMLStr=XMLStr&" <item name=""" & rs("fieldname") &""" enabled=""true"">" &vbcrlf
						 Else
							XMLStr=XMLStr&" <item  name=""" & rs("fieldname") &""" enabled=""false"">" &vbcrlf
						 End If
							XMLStr=XMLStr&"  <title>" & request("title"&rs("fieldname")) & "</title>" &vbcrlf
							XMLStr=XMLStr&"  <fieldtype>" & request("fieldtype"&rs("fieldname")) & "</fieldtype>" &vbcrlf
							XMLStr=XMLStr&"  <fieldname>" & rs("fieldname") & "</fieldname>" &vbcrlf
							XMLStr=XMLStr&"  <condition>" & request("tj"&rs("fieldname")) & "</condition>" &vbcrlf
							XMLStr=XMLStr&"  <showvalue>" & request("xsz"&rs("fieldname")) & "</showvalue>" &vbcrlf
							XMLStr=XMLStr&"  <searchvalue>" & request("ssz"&rs("fieldname")) & "</searchvalue>" &vbcrlf
							XMLStr=XMLStr&" </item>" &vbcrlf
					 RS.MoveNext
					Loop
			  End If
			  
			  '排序字段
			 If KS.ChkClng(Request("corderid"))=1 Then
					XMLStr=XMLStr&" <orderitem name=""id"" enabled=""true"">" &vbcrlf
			 Else
					XMLStr=XMLStr&" <orderitem  name=""id"" enabled=""false"">" &vbcrlf
			 End If
					XMLStr=XMLStr&"  <uptitle>" & request("uptitleid") & "</uptitle>" &vbcrlf
					XMLStr=XMLStr&"  <downtitle>" & request("downtitleid") & "</downtitle>" &vbcrlf
					XMLStr=XMLStr&" </orderitem>" &vbcrlf
			 If KS.ChkClng(Request("corderadddate"))=1 Then
					XMLStr=XMLStr&" <orderitem name=""adddate"" enabled=""true"">" &vbcrlf
			 Else
					XMLStr=XMLStr&" <orderitem  name=""adddate"" enabled=""false"">" &vbcrlf
			 End If
					XMLStr=XMLStr&"  <uptitle>" & request("uptitleadddate") & "</uptitle>" &vbcrlf
					XMLStr=XMLStr&"  <downtitle>" & request("downtitleadddate") & "</downtitle>" &vbcrlf
					XMLStr=XMLStr&" </orderitem>" &vbcrlf
			  
			  Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "select * from ks_field Where (fieldtype=4 or fieldtype=12 or fieldtype=5) and ChannelID=" & channelid & " order by orderid,fieldid",conn,1,1
			  If Not RS.Eof Then
					Do While Not RS.Eof
						 If KS.ChkClng(Request("corder" & rs("fieldname")))=1 Then
							XMLStr=XMLStr&" <orderitem name=""" & rs("fieldname") &""" enabled=""true"">" &vbcrlf
						 Else
							XMLStr=XMLStr&" <orderitem  name=""" & rs("fieldname") &""" enabled=""false"">" &vbcrlf
						 End If
							XMLStr=XMLStr&"  <uptitle>" & request("uptitle"&rs("fieldname")) & "</uptitle>" &vbcrlf
							XMLStr=XMLStr&"  <downtitle>" & request("downtitle"&rs("fieldname")) & "</downtitle>" &vbcrlf
							XMLStr=XMLStr&" </orderitem>" &vbcrlf
					 RS.MoveNext
					Loop
			  End If
			'顶部菜单筛选选项
			Dim optionnum:optionnum=KS.ChkClng(Request("optionnum"))
			If optionnum<>0 Then
			  For I=1 To optionnum
			    if request("option"&i)<>"" and request("optionsql"&i)<>"" then
					XMLStr=XMLStr&" <optionitem  name=""" & i &""">" &vbcrlf
					XMLStr=XMLStr&"  <title>" & request("option"&i) & "</title>" &vbcrlf
					XMLStr=XMLStr&"  <sqlparam><![CDATA[" & request("optionsql"&i) & "]]></sqlparam>" &vbcrlf
					XMLStr=XMLStr&" </optionitem>" &vbcrlf
				end if
			  Next	
			End If
			
			  
		   XMLStr=XMLStr &" </field>" &vbcrlf
		   Call KS.WriteTOFile(KS.Setting(3) & "config/filtersearch/s" & ChannelID & ".xml",xmlstr)
		   RS.Close :Set RS=Nothing
		   KS.Die "<script>alert('恭喜，[" & KS.C_S(ChannelID,1) & "]筛选参数配置成功!');location.href='KS.Model.asp?Action=SetSearch&ChannelID=" & channelid &"'</script>"
		 End If
		 	
			set FieldXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			FieldXML.async = false
			FieldXML.setProperty "ServerHTTPRequest", true 
			FieldXML.load(Server.MapPath(KS.Setting(3)& "config/filtersearch/s" & ChannelID & ".xml"))
			if FieldXML.parseError.errorCode<>0 Then
				XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
			    XMLStr=XMLStr&"<field>" &vbcrlf
				XMLStr=XMLStr&" <template><![CDATA[{@TemplateDir}/" & KS.C_S(ChannelID,1) & "/筛选模板.html]]></template>" &vbcrlf
				XMLStr=XMLStr&" <isrewrite>0</isrewrite>" &vbcrlf
				XMLStr=XMLStr&" <maxperpage>20</maxperpage>" &vbcrlf
			    XMLStr=XMLStr&" <item name=""tid"" enabled=""true"">" &vbcrlf
			    XMLStr=XMLStr&"  <title>分类</title>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldname>tid</fieldname>" &vbcrlf
			    XMLStr=XMLStr&"  <fieldtype>-1</fieldtype>" &vbcrlf
			    XMLStr=XMLStr&"  <condition>dy</condition>" &vbcrlf
			    XMLStr=XMLStr&"  <showvalue>0</showvalue>" &vbcrlf
			    XMLStr=XMLStr&"  <searchvalue>0</searchvalue>" &vbcrlf
			    XMLStr=XMLStr&" </item>" &vbcrlf
				XMLStr=XMLStr&" <orderitem name=""id"" enabled=""true"">" &vbcrlf
				XMLStr=XMLStr&"  <uptitle>按" & KS.C_S(ChannelID,3) & "ID升序</uptitle>" &vbcrlf
				XMLStr=XMLStr&"  <downtitle>按" & KS.C_S(ChannelID,3) & "ID降序</downtitle>" &vbcrlf
				XMLStr=XMLStr&" </orderitem>" &vbcrlf
				XMLStr=XMLStr&" <optionitem  name=""1"">" &vbcrlf
				XMLStr=XMLStr&"  <title>所有" & KS.C_S(ChannelID,3) & "</title>" &vbcrlf
				XMLStr=XMLStr&"  <sqlparam><![CDATA[1=1]]></sqlparam>" &vbcrlf
				XMLStr=XMLStr&" </optionitem>" &vbcrlf
				XMLStr=XMLStr&" <optionitem  name=""2"">" &vbcrlf
				XMLStr=XMLStr&"  <title>推荐" & KS.C_S(ChannelID,3) & "</title>" &vbcrlf
				XMLStr=XMLStr&"  <sqlparam><![CDATA[recommend=1]]></sqlparam>" &vbcrlf
				XMLStr=XMLStr&" </optionitem>" &vbcrlf
			    XMLStr=XMLStr&"</field>" &vbcrlf
                Call KS.WriteTOFile(KS.Setting(3) & "config/filtersearch/s" & ChannelID & ".xml",xmlstr)
			    FieldXML.load(Server.MapPath(KS.Setting(3)& "config/filtersearch/s" & ChannelID & ".xml"))
			End If

		%>
		<table width='100%' border='0' cellspacing='0' cellpadding='0'>  
		 <tr>
		  <td height='25' class='sort' colspan="10">[<span style='color:red'><%=KS.C_S(ChannelID,1)%></span>]筛选参数设置</td> 
		 </tr>
		 <form name="ManageMenuForm" id="ManageMenuForm" action="KS.Model.asp" method="post">
		 <%
		  if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("template")
				 If Not Node Is Nothing Then
				  TemplateFile=Node.Text
				 End If
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("isrewrite")
				 If Not Node Is Nothing Then
				  isrewrite=Node.Text
				 End If
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("maxperpage")
				 If Not Node Is Nothing Then
				  maxperpage=Node.Text
				 End If
		  end if
          if KS.IsNul(TemplateFile) Then TemplateFile="{@TemplateDir}/" & KS.C_S(ChannelID,1) & "/筛选模板.html"
		  If KS.IsNul(isrewrite) Then isrewrite=0
		  If KS.IsNul(maxperpage) Then maxperpage=20
		 %>
		 <tr class='tdbg'>
		   <td colspan=10 height='40' style="text-align:left">&nbsp;<strong>绑定模板：</strong><input type='text' name='TemplateFile' id='TemplateFile' class="textbox" value="<%=TemplateFile%>" size="40"/> <%=KSCls.Get_KS_T_C("TemplateFile")%>
		   <%if isrewrite="1" then%>
		   <a href="../search/c-<%=channelid%>" target="_blank">点此预览</a>
		   <%else%>
		   <a href="../item/?c-<%=channelid%>" target="_blank">点此预览</a>
		   <%end if%>
		   </td>
		 </tr>
		 <tr class='tdbg'>
		   <td colspan=10 height='40' style="text-align:left">&nbsp;<strong>是否启用伪静态：</strong>
		   <label><input type="radio" name="isrewrite" value="0"<%If isrewrite="0" then response.write " checked"%>/>不开启</label>
		   <label><input type="radio" name="isrewrite" value="1"<%If isrewrite="1" then response.write " checked"%>/>开启(<font color=green>需要服务器支持Rewrite组件</font>)</label>
		   
		   搜索结果每页显示<input type="text" name="maxperpage" class="textbox" value="<%=maxperpage%>" style="text-align:center;width:40px"/>条
		   
		   </td>
		 </tr>
		 <tr class="tdbg">
		   <td style='text-align:center;width:40px' class='sort'>启用</td>
		   <td class='sort'>供选字段</td>
		   <td style='text-align:center' class='sort'>名称</td>
		   <td style='text-align:center' class='sort'>条件</td>
		   <td style='text-align:center' class='sort'>显示的值</td>
		   <td style='text-align:center' class='sort'>搜索的值</td>
		 </tr>
		 <input type="hidden" name="action" value="SetSearch" />
		 <input type="hidden" name="channelid" value="<%=ChannelID%>"/>
		 <input type="hidden" name="flag" value="dosave"/>
		 <%
		 if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("item[@name='tid']")
				 If Not Node Is Nothing Then
				  title=node.selectsinglenode("title").text
				  check=node.selectsinglenode("@enabled").text
				 Else
				  title="分类"
				  check=false
				 End If
		  end if
		 %>
		 <tr class="tdbg" onmouseout="this.className='list'" onmouseover="this.className='listmouseover'" id='utid'>
		  <td class='splittd'  onclick="chk_iddiv('tid')" style='text-align:center' height="30"><input <%if check then response.write " checked"%> onclick="chk_iddiv('tid')" type='checkbox' name='ctid' value="1" /></td>
		  <td class='splittd'  onclick="chk_iddiv('tid')" width="130">所属栏目<span class='tips'>(tid)</span></td>
		  <td class="splittd"><input type="text" name="titletid" size="8" value="<%=title%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'>---</td>
		  <td style='text-align:center' class='splittd'>---</td>
		  <td style='text-align:center' class='splittd'>---</td>
		 </tr>
		 <%if KS.ChkClng(KS.C_S(ChannelID,6))=5 then
		 if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("item[@name='price_member']")
				 If Not Node Is Nothing Then
				  title=node.selectsinglenode("title").text
				  check=node.selectsinglenode("@enabled").text
				  xsz=node.selectsinglenode("showvalue").text
				  ssz=node.selectsinglenode("searchvalue").text
				 Else
				  title="商城价":check=false
				  xsz="0-10元,10-100元,100-300元,300-500元,500-1000元,1000元以上"
				  ssz="0-10,10-100,100-300,300-500,500-1000,1000-100000"
				 End If
		  end if
		 
		 %>
		 <tr class="tdbg" onmouseout="this.className='list'" onmouseover="this.className='listmouseover'" id='uprice_member'>
		  <td class='splittd'  onclick="chk_iddiv('price_member')" style='text-align:center' height="30"><input <%if check then response.write " checked"%> onclick="chk_iddiv('price_member')" type='checkbox' name='cprice_member' value="1" /></td>
		  <td class='splittd'  onclick="chk_iddiv('price_member')" width="130">商城价<span class='tips'>(price_member)</span></td>
		  <td class="splittd"><input type="text" name="titleprice_member" size="8" value="<%=title%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'>&nbsp;<select name="tjprice_member">
		  <option value='fw' selected>范围(数字型)</option>
		  </select></td>
		  <td style='text-align:center' class='splittd'><input type="text" style="width:220px" name="xszprice_member" value="<%=xsz%>"  class="textbox"/>
			<div class='tips'>多个用英文逗号隔开如:0-10元,10-100元,100-1000元,1000元以上</div></td>
		  <td style='text-align:center' class='splittd'><input type="text" style="width:220px" name="sszprice_member" value="<%=ssz%>"  class="textbox"/>
			<div class='tips'>多个用英文逗号隔开如:0-10,10-100,100-1000,1000-100000</div></td>
		 </tr>
		 <%
		 if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("item[@name='brandid']")
				 If Not Node Is Nothing Then
				  title=node.selectsinglenode("title").text
				  check=node.selectsinglenode("@enabled").text
				 Else
				  title="品牌":check=false
				 End If
		  end if
		 %>
		 <tr class="tdbg" onmouseout="this.className='list'" onmouseover="this.className='listmouseover'" id='ubrandid'>
		  <td class='splittd'  onclick="chk_iddiv('brandid')" style='text-align:center' height="30"><input <%if check then response.write " checked"%> onclick="chk_iddiv('brandid')" type='checkbox' name='cbrandid' value="1" /></td>
		  <td class='splittd'  onclick="chk_iddiv('brandid')" width="130">所属品牌<span class='tips'>(brandid)</span></td>
		  <td class="splittd"><input type="text" name="titlebrandid" size="8" value="<%=title%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'>---</td>
		  <td style='text-align:center' class='splittd'>---</td>
		  <td style='text-align:center' class='splittd'>---</td>
		 </tr>
		 <%elseif KS.ChkClng(KS.C_S(ChannelID,6))=8 then
			 if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
					 Set Node=FieldXML.DocumentElement.SelectSingleNode("item[@name='typeid']")
					 If Not Node Is Nothing Then
					  title=node.selectsinglenode("title").text
					  check=node.selectsinglenode("@enabled").text
					 Else
					  title="交易类别":check=false
					 End If
			  end if
		 %>
		 <tr class="tdbg" onmouseout="this.className='list'" onmouseover="this.className='listmouseover'" id='ubrandid'>
		  <td class='splittd'  onclick="chk_iddiv('typeid')" style='text-align:center' height="30"><input <%if check then response.write " checked"%> onclick="chk_iddiv('typeid')" type='checkbox' name='ctypeid' value="1" /></td>
		  <td class='splittd'  onclick="chk_iddiv('typeid')" width="130">交易类别<span class='tips'>(typeid)</span></td>
		  <td class="splittd"><input type="text" name="titletypeid" size="8" value="<%=title%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'>---</td>
		  <td style='text-align:center' class='splittd'>---</td>
		  <td style='text-align:center' class='splittd'>---</td>
		 </tr>
		 <%
		 end if
		 
		 
		  set rs=server.CreateObject("adodb.recordset")
		  rs.open "select * from ks_field where (fieldname in(" & AllowField & ") or fieldtype<>0) and channelid=" & channelid & " order by orderid,fieldid",conn,1,1
		  do while not rs.eof
		   	if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("item[@name='" & RS("FieldName") &"']")
				 If Not Node Is Nothing Then
				  check=node.selectsinglenode("@enabled").text
				  tj=node.selectsinglenode("condition").text
				  xsz=node.selectsinglenode("showvalue").text
				  ssz=node.selectsinglenode("searchvalue").text
				  title=Node.SelectSingleNode("title").text
				 Else
				  title=rs("title")
				  check=false
				  tj=""
				  xsz=""
				  ssz=""
				 End If
			end if

		  %>
		 <tr class="tdbg" onmouseout="this.className='list'" onmouseover="this.className='listmouseover'" id='u<%=rs("fieldname")%>'>
		 <input type="hidden" name="fieldtype<%=rs("fieldname")%>" value="<%=rs("fieldtype")%>"/>
		  <td class='splittd' onclick="chk_iddiv('<%=rs("fieldname")%>')" style='text-align:center' height="30"><input onclick="chk_iddiv('<%=rs("fieldname")%>')"  type='checkbox' name='c<%=rs("fieldname")%>' value="1"<%if check then response.Write(" checked")%> /></td>
		  <td class='splittd' onclick="chk_iddiv('<%=rs("fieldname")%>')"><%=rs("title")%><span class='tips'>(<%=rs("fieldname")%>)</span></td>
		  <td class='splittd'><input  size="8" type="text" class='textbox' name="title<%=rs("fieldname")%>" value="<%=title%>"/></td>
		  <td style='text-align:center' class='splittd'>
		  <select name="tj<%=rs("fieldname")%>">
		    <%if rs("fieldtype")<>4 and rs("fieldtype")<>12 then%>
		    <option value='dy'<%if tj="dy" then response.write " selected"%>>等于(字符型)</option>
			<%else%>
		    <option value='dys'<%if tj="dys" then response.write " selected"%>>等于(数字型)</option>
		    <option value='fw'<%if tj="fw" then response.write " selected"%>>范围(数字型)</option>
			<%end if%>
			<%if rs("fieldtype")<>3 and rs("fieldtype")<>11 and rs("fieldtype")<>6 and rs("fieldtype")<>7 and rs("fieldtype")<>4 then%>
		    <option value='like'<%if tj="like" then response.write " selected"%>>包含(字符型)</option>
			<%end if%>
		  </select>
		  </td>
		  <td style='text-align:center' class='splittd tips'>
		   <%if rs("fieldtype")=3 or rs("fieldtype")=11 or rs("fieldtype")=6 or rs("fieldtype")=7 then%>
		    <input type="hidden" name="xsz<%=rs("fieldname")%>" value="0">自动显示，字段里设置的选项
		   <%else%>
		    <input type="text" style="width:220px" name="xsz<%=rs("fieldname")%>" value="<%=xsz%>"  class="textbox"/>
			<div class='tips'>多个用英文逗号隔开如:免费,收费,试听</div>
		   <%end if%>
		  </td>
		  <td style='text-align:center' class='splittd tips'>
		  	<%if rs("fieldtype")=3 or rs("fieldtype")=11 or rs("fieldtype")=6 or rs("fieldtype")=7 then%>
		    <input type="hidden" name="ssz<%=rs("fieldname")%>" value="0">自动显示，字段里设置的选项
		   <%else%>
		    <input type="text" style="width:220px" name="ssz<%=rs("fieldname")%>" value="<%=ssz%>"  class="textbox"/>
			<div class='tips'>多个用英文逗号隔开如:0,1,2</div>
		   <%end if%>

		  </td>
		 </tr>
		  <%
		   rs.movenext
		  loop
		  rs.close
		  set rs=nothing
		 %>
		 </table>
		<table width='100%' border='0' cellspacing='0' cellpadding='0'> 
		 <tr>
		   <td class="clefttitle" colspan="10" style="text-align:left;height:28px;padding-left:4px;"><strong>排序设置：</strong></td>
		 </tr>
		 <tr class="tdbg">
		   <td style='text-align:center;width:40px' class='sort'>启用</td>
		   <td class='sort'>供选字段</td>
		   <td style='text-align:center' class='sort'>升序名称</td>
		   <td style='text-align:center' class='sort'>降序名称</td>
		 </tr>
		 <%
		 if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("orderitem[@name='id']")
				 If Not Node Is Nothing Then
				  uptitle=node.selectsinglenode("uptitle").text
				  downtitle=node.selectsinglenode("downtitle").text
				  check=node.selectsinglenode("@enabled").text
				 Else
				  uptitle="按"& KS.C_S(ChannelID,3) &"ID号升序"
				  downtitle="按"& KS.C_S(ChannelID,3) &"ID号降序"
				  check=false
				 End If
		  end if
		 %>
		 <tr class="tdbg" onmouseout="this.className='list'" onmouseover="this.className='listmouseover'" id='uorderid'>
		  <td class='splittd'  onclick="chk_iddiv('orderid')" style='text-align:center' height="30"><input onclick="chk_iddiv('orderid')" <%if cbool(check)=true then response.write " checked"%> type='checkbox'  name='corderid' value="1" /></td>
		  <td class='splittd'  onclick="chk_iddiv('orderid')" width="130"><%=KS.C_S(ChannelID,3)%><span class="tips">(id号)</span></td>
		  <td style='text-align:center' class="splittd"><input type="text" name="uptitleid" size="28" value="<%=uptitle%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'><input type="text" name="downtitleid" size="28" value="<%=downtitle%>" class="textbox"/></td>
		 </tr>
		 <%
		 if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("orderitem[@name='adddate']")
				 If Not Node Is Nothing Then
				  uptitle=node.selectsinglenode("uptitle").text
				  downtitle=node.selectsinglenode("downtitle").text
				  check=node.selectsinglenode("@enabled").text
				 Else
				  uptitle="按添加时间升序"
				  downtitle="按添加时间降序"
				  check=false
				 End If
		  end if
		 %>
		 <tr class="tdbg" onmouseout="this.className='list'" onmouseover="this.className='listmouseover'" id='uorderadddate'>
		  <td class='splittd'  onclick="chk_iddiv('orderadddate')" style='text-align:center' height="30"><input onclick="chk_iddiv('orderadddate')" <%if cbool(check)=true then response.write " checked"%> type='checkbox'  name='corderadddate' value="1" /></td>
		  <td class='splittd'  onclick="chk_iddiv('orderadddate')" width="130">添加时间<span class="tips">(adddate)</span></td>
		  <td style='text-align:center' class="splittd"><input type="text" name="uptitleadddate" size="28" value="<%=uptitle%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'><input type="text" name="downtitleadddate" size="28" value="<%=downtitle%>" class="textbox"/></td>
		 </tr>
		 <%
		  dim uptitle,downtitle
		  set rs=server.CreateObject("adodb.recordset")
		  rs.open "select * from ks_field where (fieldtype=4 or fieldtype=12 or fieldtype=5) and channelid=" & channelid & " order by orderid,fieldid",conn,1,1
		  do while not rs.eof
		   	if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
				 Set Node=FieldXML.DocumentElement.SelectSingleNode("orderitem[@name='" & RS("FieldName") &"']")
				 If Not Node Is Nothing Then
				  check=cbool(node.selectsinglenode("@enabled").text)
				  uptitle=node.selectsinglenode("uptitle").text
				  downtitle=node.selectsinglenode("downtitle").text
				 Else
				  Uptitle="按" & rs("title") & "升序"
				  downtitle="按" & rs("title") & "降序"
				  check=false
				 End If
			end if
		  
		  %>
		 <tr class="tdbg" onmouseout="this.className='list'" onmouseover="this.className='listmouseover'" id='uorder<%=rs("fieldname")%>'>
		  <td class='splittd'  onclick="chk_iddiv('order<%=rs("fieldname")%>')" style='text-align:center' height="30"><input onclick="chk_iddiv('order<%=rs("fieldname")%>')" <%if cbool(check)=true then response.write " checked"%> type='checkbox'  name='corder<%=rs("fieldname")%>' value="1" /></td>
		  <td class='splittd'  onclick="chk_iddiv('order<%=rs("fieldname")%>')" width="130"><%=rs("title")%><span class="tips">(<%=rs("fieldname")%>)</span></td>
		  <td style='text-align:center' class="splittd"><input type="text" name="uptitle<%=rs("fieldname")%>" size="28" value="<%=uptitle%>" class="textbox"/></td>
		  <td style='text-align:center' class='splittd'><input type="text" name="downtitle<%=rs("fieldname")%>" size="28" value="<%=downtitle%>" class="textbox"/></td>
		 </tr>
		  <%
          rs.movenext
		  loop
		  rs.close
		  set rs=nothing
		 %>
		 </table>
		<table width='100%' border='0' cellspacing='0' cellpadding='0'> 
		 <tr>
		   <td class="clefttitle" colspan="10" style="text-align:left;height:28px;padding-left:4px;"><strong>顶部选项卡筛选项：</strong>  <input type='button' class="button" value="添加一个选项" onclick="doadd(1)"/>
		   </td>
		 </tr>
		 <tr class="tdbg">
		   <td style='padding-left:5px' class='sort' width="20%">选项卡名称</td>
		   <td class='sort' width="80%">SQL查询条件</td>
		 </tr>
		 <tr>
		  <td colspan="2" id="addvote">
		    <table width="100%"  cellpadding="0" cellspacing="0">
			<%
			dim nn,ii
			ii=1
			Set Node=FieldXML.DocumentElement.SelectNodes("optionitem")
			If Node.Length>0 Then
				 For Each nn In FieldXML.DocumentElement.SelectNodes("optionitem")
				 %>
			 <tr class="tdbg">
			   <td style='padding-left:5px' width="20%" class='splittd'><input type='text' name='option<%=ii%>' value="<%=NN.selectSingleNode("title").text%>" class="textbox" />
			  <%if ii=1 then%><div class='tips'>如：推荐信息</div><%end if%>
			   </td>
			   <td class='splittd' width="80%"><input type='text' size="50" name='optionsql<%=ii%>' value="<%=server.HTMLEncode(NN.selectSingleNode("sqlparam").text)%>" class="textbox" />
			   <%if ii=1 then%><div class='tips'>如只需要显示推荐的信息可以输入recommend=1,显示推荐和头条的可以输入 recommend=1 and strip=1等。</div><%end if%>
			   </td>
			 </tr>
				 <%
				 ii=ii+1
				 Next
				 II=II-1
			Else	 
			%>
			 <tr class="tdbg">
			   <td style='padding-left:5px' width="20%" class='splittd'><input type='text' name='option1' value="" class="textbox" />
			   <div class='tips'>如：推荐信息</div>
			   </td>
			   <td class='splittd' width="80%"><input type='text' size="50" name='optionsql1' value="" class="textbox" />
			   <div class='tips'>如只需要显示推荐的信息可以输入recommend=1,显示推荐和头条的可以输入 recommend=1 and strip=1等。</div>
			   </td>
			 </tr>
		 <%end if%>
		 
			 </table>
			 
			 <input type='hidden' name='optionnum'  id='optionnum' value='<%=ii%>'/>

		  </td>
		</tr>
		 <tr class="tdbg">
		   <td colspan=2 style='padding-left:5px'><span class="tips">说明：要删除某个选项，请留空然后保存即可。</span></td>
		 </tr>
		 <tr class='tdbg'>
		   <td colspan=10 class='clefttitle' height='30' style="text-align:center"><Input type='submit' value='保存设置' class='button'/></td>
		 </tr>
		 </form>
		 </table>
	<script type="text/javascript">
    function doadd(num)
    {var i;
    var str="";
    var j=0;
	var optionnum=$("#optionnum").val();
	var id=0;
    for(i=1;i<=num;i++)
    {
	 id=parseInt(optionnum)+i;
     str=str+"<tr class='tdbg'><td style='padding-left:5px' width='20%' class='splittd'><input type=text name=option"+id+" class='textbox'></td><td class='splittd' width=80%><input type=text name='optionsql"+id+"' size=50 class='textbox'></td></tr>";
    }
     jQuery("#addvote").html(jQuery("#addvote").html()+"<table width=100% border=0 cellspacing=0 cellpadding=0>"+str+"</table>");
	 $("#optionnum").val(parseInt(optionnum)+1)
    }
    </script>

		 <br/><br/>
		 <div class="attention">
<strong>特别提醒：</strong><br/>
    <li>条件可取 等于（字符型）、等于（数字型）、范围（数字型），包含（字符型）,当指定范围时,还要指定搜索的值,搜索值之间用逗号隔开,如1-20,表示大于1,小于等于20
</li>
	<li>显示的值:即显示在搜索页面供选择的项                     可以带单位,如 1-20万,20-30万</li>
	<li>搜索的值:即用于供数据库搜索的值,搜索值用英文逗号分开      不能带单位,如 1-20,20-30</li>
    
</div>
		 <%
		End Sub
		
		Sub ManageMenu()
		   Dim FieldRS,ChannelID,FieldSql,Doc,Node,XmlFields,XmlFieldArr,Fi,From,xmlname
		   ChannelID=KS.ChkClng(KS.S("ChannelID"))
		   From=KS.S("From")
		   if From="user" then xmlname="usermodelfield" else xmlname="managemodelfield"

		  If KS.G("flag")="dosave" then
		    If KS.IsNul(KS.S("hasfield")) Then
			  'KS.AlertHintScript "对不起,您没有选择供选字段!"
			  'Exit Sub
			End If
		 	set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(Server.MapPath(KS.Setting(3)&"Config/" & xmlname & ".xml"))
			Set Node=Doc.documentElement.selectSingleNode("/modelfield/model[@name='" & ChannelID & "']")
			
			 if not node is nothing then  Doc.DocumentElement.RemoveChild(Node)
			 Set Node=Doc.documentElement.appendChild(Doc.createNode(1,"model",""))
			 Node.attributes.setNamedItem(Doc.createNode(2,"name","")).text=channelid
			 Node.text=Replace(KS.S("hasfield")," ","")
			Doc.Save(Server.MapPath(KS.Setting(3)&"Config/" & xmlname & ".xml"))
			Application(KS.SiteSN&"_Config"&xmlname)=empty
			Response.Write "<script>alert('恭喜,管理列表菜单配置成功!');</script>"
          End If
		    XmlFields=LFCls.GetConfigFromXML(xmlname,"/modelfield/model",ChannelID)
			If Not KS.IsNul(XmlFields) Then
			 XmlFieldArr=Split(XmlFields,",")
			End If
		 %>
		 <script type="text/javascript">
		 function doOrder(select,sequence)                    //将上、下两个方法合并成一个 
			{ 
			   if (!select||select.selectedIndex==-1)            //如果没有选择列表项，不进行任何操作 
				return false; 
			   with (select) 
			   {   
				   var newIndex = selectedIndex + sequence;      //获取移动后的索引 
				   var oldIndex = selectedIndex;               //旧索引 
				   if (newIndex>=options.length||newIndex<0||sequence==0||newIndex<0) //判断是否超出边界 
				   { 
					   return false; 
				   } 
				   options[newIndex].swapNode(options[oldIndex]) //交换指定索引处的节点 
				} 
			
				return true; 
			} 
			function doUp() 
			{ 
			   doOrder(document.all.hasfield,-1);              //向上移动的方法 
			} 
			
			function doDown() 
			{ 
			   doOrder(document.all.hasfield,1);                //向下移动的方法 
			} 
            function doSelectAll(){
			  $("#hasfield option").attr("selected",true);
			  $("#ManageMenuForm").submit();
			}
			function add(){
			   var alloptions = $("#selectfield option");
			   var so = $("#selectfield option:selected");
			   var a = (so.get(so.length-1).index == alloptions.length-1)? so.prev().attr("selected",true):so.next().attr("selected",true);
                
				if (!$("#hasfield option[value="+so.val()+"]").attr("selected")){
				 $("#hasfield").append(so);
				 }else{
				 so.remove();}
			}
			function del(){
				var alloptions = $("#hasfield option");
				 var so = $("#hasfield option:selected");
				 var a = (so.get(so.length-1).index == alloptions.length-1)? so.prev().attr("selected",true):so.next().attr("selected",true);
			   
				$("#selectfield").append(so);

			}
		 </script>
		 <table width='100%' border='0' cellspacing='0' cellpadding='0'>  
		 <tr>
		  <td height='25' class='sort' colspan="4">[<span style='color:red'><%=KS.C_S(ChannelID,1)%></span>]模型<%if From="user" then response.write "会员中心" else response.write "后台" %>管理列表菜单设置</td> 
		 </tr> 
		 <tr><td height=45 colspan="4">
		 <div class="options">
		  <ul>
		  <li<%If from="user" then response.write " class='curr'"%>><a href="KS.Model.asp?action=ManageMenu&ChannelID=<%=ChannelID%>&from=user">设置会员中心管理列表菜单</a></li>
		  <li<%If from="" then response.write " class='curr'"%>><a href="KS.Model.asp?action=ManageMenu&ChannelID=<%=ChannelID%>">设置后台管理列表菜单</a></li>
		  </ul>
         </div>
		 </td></tr>
		 <tr class="tdbg">
		   <td class='sort'>供选字段</td>
		   <td class='sort'>&nbsp;</td>
		   <td class='sort'>已选字段</td>
		   <td class='sort'>&nbsp;</td>
		 </tr>
		 <form name="ManageMenuForm" id="ManageMenuForm" action="KS.Model.asp" method="post">
		 <input type="hidden" name="action" value="ManageMenu" />
		 <input type="hidden" name="channelid" value="<%=ChannelID%>"/>
		 <input type="hidden" name="flag" value="dosave"/>
		 <input type="hidden" name="from" value="<%=from%>"/>
		 <tr class="tdbg">
		 <td class="clefttitle" width="290">
		  <select name="selectfield" id="selectfield" multiple="multiple" style="width:280px" size="16">
		   <%if instr(lcase(XmlFields),"|inputer")=0 then%>
		   <option>录入员|Inputer</option>
		   <%end if%>
		   <%if instr(lcase(XmlFields),"|refreshtf")=0 then%>
		   <option>生成标志|refreshtf</option>
		   <%end if%>
		   <%if instr(lcase(XmlFields),"|adddate")=0 then%>
		   <option>更新时间|AddDate</option>
		   <%end if%>
		   <%if instr(lcase(XmlFields),"|modeltype")=0 then%>
		   <option>类型|ModelType</option>
		   <%end if%>
		   <%if instr(lcase(XmlFields),"|attribute")=0 then%>
		   <option>文档属性|Attribute</option>
		   <%end if%>
		   <%if instr(lcase(XmlFields),"|hits")=0 then%>
		   <option>点击数|Hits</option>
		   <%end if%>
		   <%if instr(lcase(XmlFields),"|author")=0 and channelid<>5 and channelid<>7 and channelid<>8 then%>
		   <option>作者|Author</option>
		   <%end if%>
		   <%if instr(lcase(XmlFields),"|keywords")=0 then%>
		   <option>关键字|KeyWords</option>
		   <%end if%>
		   <%if instr(lcase(XmlFields),"|rank")=0 and channelid<>8 then%>
		   <option>等级|Rank</option>
		   <%end if%>
		   <%
		   if channelid<>5 and channelid<>8 then
			 if instr(lcase(XmlFields),"|readpoint")=0 then
			  response.write "<option>所需费用|ReadPoint</option>"
			 end if
		   end if
		   Select Case KS.ChkClng(KS.C_S(channelid,6))
		    case 1
			 if instr(lcase(XmlFields),"|fulltitle")=0 then
			  response.write "<option>完整标题|FullTitle</option>"
			 end if
			 if instr(lcase(XmlFields),"|origin")=0 then
			  response.write "<option>来源|Origin</option>"
			 end if
			 if instr(lcase(XmlFields),"|province")=0 then
			  response.write "<option>省份|Province</option>"
			 end if
			 if instr(lcase(XmlFields),"|city")=0 then
			  response.write "<option>城市|City</option>"
			 end if
			case 3
			 if instr(lcase(XmlFields),"|downlb")=0 then
			  response.write "<option>类别|DownLB</option>"
			 end if
			 if instr(lcase(XmlFields),"|downyy")=0 then
			  response.write "<option>语言|DownYY</option>"
			 end if
			 if instr(lcase(XmlFields),"|downsq")=0 then
			  response.write "<option>授权|DownSQ</option>"
			 end if
			 if instr(lcase(XmlFields),"|downpt")=0 then
			  response.write "<option>运行平台|DownPT</option>"
			 end if
			 if instr(lcase(XmlFields),"|ysdz")=0 then
			  response.write "<option>演示地址|YSDZ</option>"
			 end if
			 if instr(lcase(XmlFields),"|zcdz")=0 then
			  response.write "<option>注册地址|ZCDZ</option>"
			 end if
			 if instr(lcase(XmlFields),"|hitsbyday")=0 then
			  response.write "<option>日下载数|HitsByDay</option>"
			 end if
			 if instr(lcase(XmlFields),"|hitsbyweek")=0 then
			  response.write "<option>周下载数|HitsByWeek</option>"
			 end if
			 if instr(lcase(XmlFields),"|hitsbymonth")=0 then
			  response.write "<option>月下载数|HitsByMonth</option>"
			 end if
			case 7

			 if instr(lcase(XmlFields),"|movieact")=0 then
			  response.write "<option>演员|MovieAct</option>"
			 end if
			 if instr(lcase(XmlFields),"|moviedy")=0 then
			  response.write "<option>导演|MovieDY</option>"
			 end if
			 if instr(lcase(XmlFields),"|movieyy")=0 then
			  response.write "<option>语言|MovieYY</option>"
			 end if
			 if instr(lcase(XmlFields),"|moviedq")=0 then
			  response.write "<option>地区|MovieDQ</option>"
			 end if
			 if instr(lcase(XmlFields),"|movietime")=0 then
			  response.write "<option>时长|MovieTime</option>"
			 end if
			 if instr(lcase(XmlFields),"|screentime")=0 then
			  response.write "<option>上映时间|ScreenTime</option>"
			 end if
		   case 5
			 if instr(lcase(XmlFields),"|unit")=0 then
			  response.write "<option>单位|Unit</option>"
			 end if
			 if instr(lcase(XmlFields),"|proid")=0 then
			  response.write "<option>商品编号|Proid</option>"
			 end if
			 if instr(lcase(XmlFields),"|totalnum")=0 then
			  response.write "<option>库存量|TotalNum</option>"
			 end if
			 if instr(lcase(XmlFields),"|price")=0 then
			  response.write "<option>参考价|Price</option>"
			 end if
			 if instr(lcase(XmlFields),"|price_member")=0 then
			  response.write "<option>商城价|Price_Member</option>"
			 end if
		  case 8
			 if instr(lcase(XmlFields),"|price")=0 then
			  response.write "<option>价格|price</option>"
			 end if
			 if instr(lcase(XmlFields),"|contactman")=0 then
			  response.write "<option>联系人|ContactMan</option>"
			 end if
			 if instr(lcase(XmlFields),"|tel")=0 then
			  response.write "<option>电话|Tel</option>"
			 end if
			 if instr(lcase(XmlFields),"|companyname")=0 then
			  response.write "<option>公司|CompanyName</option>"
			 end if
			 if instr(lcase(XmlFields),"|address")=0 then
			  response.write "<option>地址|Address</option>"
			 end if
			 if instr(lcase(XmlFields),"|province")=0 then
			  response.write "<option>省份|province</option>"
			 end if
			 if instr(lcase(XmlFields),"|city")=0 then
			  response.write "<option>城市|City</option>"
			 end if
			 if instr(lcase(XmlFields),"|zip")=0 then
			  response.write "<option>邮编|Zip</option>"
			 end if
			 if instr(lcase(XmlFields),"|fax")=0 then
			  response.write "<option>传真|Fax</option>"
			 end if
			 if instr(lcase(XmlFields),"|email")=0 then
			  response.write "<option>邮箱|email</option>"
			 end if
		  end select
		   %>
		   <optgroup  style="color:red" label="=====用户自定义字段====="></optgroup>
		   <%
		    Set FieldRS = Server.CreateObject("ADODB.RecordSet")
			FieldSql = "SELECT FieldName,Title FROM KS_Field Where ChannelID=" & ChannelID & " order by orderid asc"
			FieldRS.Open FieldSql, conn, 1, 1
            Do While Not FieldRS.Eof
			 if instr(lcase(XmlFields),"|" & lcase(FieldRS("FieldName")))=0 then
			 response.write "<option style='color:green'>" & FieldRS("Title") & "|" & FieldRS("FieldName") & "</option>"
			 end if
			FieldRS.MoveNext
			Loop
			FieldRS.Close
			Set FieldRS=Nothing
		   %>
		  </select>
		 </td>
		 <td align="center" width="80">&nbsp;
		  <input type="button" value="添加>>" onClick="add()" class="button"/><br/><br/>
		  &nbsp;
		  <input type="button" value="<<移除" onClick="del()" class="button"/>
		 </td>
		 <td>
		  <select name="hasfield" id="hasfield"  multiple size="16" style="width:290px;">
		  <%
			If IsArray(XmlFieldArr) Then
			  For Fi=0 To Ubound(XmlFieldArr)
			    response.write "<option selected>" &XmlFieldArr(fi) & "</option>" 
			  Next
			End If
		  %>
		  </select>
		 </td>
		 <td align="center" width="80" align="left">
		  <input type="button" value="上移↑" onClick="doUp()" class="button"/><br/><br/>
		  <input type="button" value="下移↓" onClick="doDown()" class="button"/>
		 </td>
		 </tr>
		 <tr class='tdbg'>
		   <td colspan=4 height='30' style="text-align:center"><Input type='button' onClick="doSelectAll()" value='保存设置' class='button'/></td>
		 </tr>
		 </table>
		 <br/><br/>
		 <div class="attention">
<strong>特别提醒：</strong>
管理列表显示的字段越少则查询显示速度会越快,一般不常用的字段建议不要选择。
</div>
		 <%
		End Sub
 
		Sub Main()
		   With Response
			.Write "<script>"
			.Write "$(document).ready(function(){"
			.Write "parent.frames['BottomFrame'].Button1.disabled=true;"
			.Write "parent.frames['BottomFrame'].Button2.disabled=true;"
			.Write "})</script>"
			 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select * From KS_Channel where channelid  not in(" & ChannelNotOnStr &") Order By ChannelID",conn,1,1
		    .Write "<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.Write "<tr height='25' class='sort'>"
			.Write "  <td width='50' align=center>ID</td><td align=center>模型名称</td><td align=center>数据表</td><td align=center>类型</td><td align=center>项目名称</td><td align=center>项目单位</td><td align=center>状态</td>"
			If KS.WSetting(0)="1" Then .Write "<td align=center>3G版</td>"
			.Write "<td align=center>运行模式</td><td align=center>↓操作</td>"
			.Write "</tr>"
		  Do While Not RS.Eof 
		    .Write "<tr height='23' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			.Write "<td align=center class='splittd'>" & RS("ChannelID")&"</td>"
			.Write "<td align=center class='splittd'>" & RS("ChannelName") &"</td>"
			.Write "<td class='splittd'>" & RS("ChannelTable") & "</td>"
			.Write "<td align='center' class='splittd'>"
			 If RS("ChannelID")>=100 Then
			  .Write "<font color=blue>自定义</font>"
			 Else
			  .Write "<font color='#999999'>系统</font>"
			 End If
			.Write "</td>"
			.Write "<td align=center class='splittd'>" & RS("ItemName") & "</td>"
			.Write "<td align=center class='splittd'>" & RS("ItemUnit") & "</td>"
			.Write "<td align=center class='splittd'>" 
			  If RS("ChannelStatus")="1" Then .Write "正常" Else .Write "<font color=red>已禁用</font>"
			.Write "</td>"
			If KS.WSetting(0)="1" Then
			.Write "<td align=center class='splittd'>" 
			  If RS("WapSwitch")="1" Then .Write "正常" Else .Write "<font color=red>已禁用</font>"
			.Write "</td>"
		    End If
			.Write "<td align=center class='splittd'>"
			 if rs("channelid")=10 Then 
			    If KS.JSetting(42)="1" Then
				 .Write "<font color=red>伪静态</font>"
				Else
				 .Write "<font color=red>动态asp</font>"
				End If
			 Else
				  If RS("FsoHtmlTF")="1" Then 
				  .Write "生成Html" 
				  ElseIf RS("FsoHtmlTF")="0" Then
				  .Write "<font color=red>动态asp</font>"
				   If RS("StaticTF")<>0 Then .Write "<i>(伪)</i>"
				  Else
				  .Write "<font color=blue>部分生成</font>"
				  If RS("StaticTF")<>0 Then .Write "<i>(伪)</i>"
				  End If
			 End If
			.Write "</td>"

			.Write "<td align=center class='splittd'>"
			if rs("channelid")=10 Then
			.Write "<a href='KS.JobSetting.asp' onclick=""$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=GoSave&OpStr=系统设置 >> <font color=red>系统模型管理</font>';"">修改</a>｜"
			Elseif rs("channelid")=1 or (Instr(channelNotOnStr,rs("channelid"))=0 and rs("channelid")<>10) then
			.Write "<a href='?action=Edit&ChannelID=" & rs("ChannelID") & "' onclick=""$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=GoSave&OpStr=系统设置 >> <font color=red>系统模型管理</font>';"">修改</a>｜"
			else
			.Write "<font color=#a7a7a7>修改</font>｜"
			end if
			 If RS("ChannelID")>=100 Then
			 .Write "<a href='?action=Del&ChannelID=" & rs("ChannelID") & "' onclick='return(confirm(""此操作不可逆，确定删除吗？""))'>删除</a>｜"
			 Else
			 .Write "<font color=#a7a7a7>删除</font>｜"
			 End If
			 
			 IF rs("channelid")=1 or (rs("ChannelID")<>6 and rs("channelid")<>10  and Instr(channelNotOnStr,rs("channelid"))=0) then
			 .Write "<a href='#' onClick=""SelectObjItem1(this,'模型管理 >> <font color=red>模型字段管理</font>','Disabled','KS.Field.asp?ChannelID=" & rs("ChannelID") & "');"">字段管理</a>｜"
			 else
			 .Write "<font color=#a7a7a7>字段管理</font>｜"
			 end if
			 If Instr(channelNotOnStr,rs("channelid"))=0 or rs("channelid")=1 then
			 If RS("ChannelStatus")="1" Then .Write "<a href='?Action=SetChannelParam&Flag=ChannelOpenOrClose&ChannelID=" & RS("ChannelID") & "'>禁用</a>" Else .Write "<a href='?Action=SetChannelParam&Flag=ChannelOpenOrClose&ChannelID=" & RS("ChannelID") & "'>开启</a>"
			 else
			 .Write "<font color=#a7a7a7>开启</font>"
			 end if
			 If rs("channelid")<>6  and rs("channelid")<>10 Then 
			   .Write "｜<a href='?Action=SetSearch&ChannelID=" & RS("ChannelID") & "'>筛选</a>"
			 else
			 .Write "｜<font color=#a7a7a7>筛选</font>"
			 end if
			.Write "</td></tr>"
			RS.MoveNext 
		  Loop
		    .Write "</table>"
		   RS.Close:Set RS=Nothing
		    .Write "</body>"
			.Write "</html>"
		  End With
		End Sub
		
		Sub Total()
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select * From KS_Channel Where ChannelStatus=1 order by channelid asc",conn,1,1
		   With Response
		  	.Write "<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.Write "<tr height='25' class='sort'>"
			.Write " <td align=center colspan=6>各模型信息统计</td>"
			.Write "</tr>"

		  Do While Not RS.Eof
			.Write "<tr height='25' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
		   If RS("ChannelID")=6 Then
		   ElseIf RS("ChannelID")=9 Then
		   ElseIf RS("ChannelID")=10 Then
		    .Write "<td width='140' align=center><img src='images/37.gif'>&nbsp;<b>" & RS("ChannelName") & "</b></td><td width='140'>简历总数：<font color='#ff0000'>" & Conn.Execute("select Count(ID) from KS_Job_Resume")(0) & "</font> 个</td><td width='140'>职位总数：<font color='red'>" & Conn.Execute("select count(id) from KS_Job_zw")(0) & "</font> 个 </td><td width='140'>单位总数：<font color='red'>" & Conn.Execute("select Count(ID) from KS_Job_Company")(0) & "</font> 家 </td>"
		   Else
			.Write "<td width='140' align=center><img src='images/37.gif'>&nbsp;<b>" & RS("ChannelName") & "</b></td><td width=150>频道总数: <font color=#ff0000>" & Conn.Execute("select count(id) from ks_class where channelid=" & RS("ChannelID") & " and tj=1")(0) & "</font> 个</td><td width=150>" & RS("ItemName") & "总数: <font color=blue>" & conn.Execute("Select Count(ID) From " & RS("ChannelTable") & " Where DelTF=0")(0) & " </font>" & RS("ItemUnit") & "</td><td>待审" & RS("ItemName") & ":<font color=green>" & conn.Execute("Select Count(ID) From " & RS("ChannelTable") & " Where  Verific=0")(0) & " </font>" & RS("ItemUnit") & "</td><td></td><td></td>"
		  End If
			.Write "</tr><tr><td colspan=10 background='images/line.gif'></td></tr>"
		    RS.MoveNext
		  Loop
		   .Write "</table>"
		  End With
		  RS.Close:Set RS=Nothing
		End Sub
		
		'模型设置
		Sub SetChannelParam()
		   With Response
			   Dim ChannelID:ChannelID=KS.ChkClng(KS.G("ChannelID"))
			   If ChannelID=0 Then .Redirect "?": Exit Sub
			   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select * From KS_Channel Where ChannelID=" & ChannelID,Conn,1,3
			   If RS.Eof Then
				 RS.Close:Set RS=Nothing
				.Redirect "?": Exit Sub
			   End If
		     If KS.G("Flag")="ChannelOpenOrClose" Then
			   If RS("ChannelStatus")=1 Then 
				  if conn.execute("select count(channelstatus) from ks_channel where channelstatus=1")(0)=1 then
				   rs.close:set rs=nothing
				   .Write "<script>alert('对不起，请至少保持一个模型是开启状态！');history.back();</script>"
				   .end
				   else
					RS("ChannelStatus")=0 
				   End If
			   Else 
			    RS("ChannelStatus")=1
			   end if
			   if channelid=10 then
			     Dim RSJ:Set RSJ=Conn.Execute("Select JobSetting From KS_Config")
				 If Not RS.Eof Then
				    Dim i,JArr,JobSetting:JobSetting=RSJ(0)
					Jarr=split(JobSetting,"^%^")
					For i=0 To Ubound(jarr)
					  If I=0 Then
					    JobSetting=RS("ChannelStatus")
					  Else
					    JobSetting=JobSetting & "^%^" & jarr(I)
					  End If
					Next
					Conn.Execute("Update KS_Config Set JobSetting='"& replace(JobSetting,"'","''") &"'")
				 End If
				 Set RSJ=Nothing
			   end if
			 End If
			 RS.Update
			 RS.Close:Set RS=Nothing
			 .Write "<script>top.location.href='index.asp?from=KS.Model.asp';</script>"
		   End With
		End Sub
		
		Sub ChannelAddOrEdit()
		Dim SqlStr, RS, InstallDir, FsoIndexFile,StaticTF, FsoIndexExt,FsoListNum,i,ThumbnailsConfig
		Dim ChannelName,ModelEname,ChannelTable,ChannelStatus,WapSwitch,WapSearchTemplate,ItemName,ItemUnit,FsoFolder,Descript,ModelIco,ModelShortName,CommentTime
		Dim FsoHtmlTF,BasicType,MaxPerPage,UserTF,UserClassStyle,UserEditTF
		Dim UpFilesTF,UpfilesDir,UserUpFilesTF,UserUpfilesDir,UserSelectFilesTF,UpfilesSize,AllowUpPhotoType,AllowUpFlashType,AllowUpMediaType,AllowUpRealType,AllowUpOtherType
		Dim  UserAddMoney,UserAddPoint,UserAddScore,RefreshFlag,InfoVerificTF,VerificCommentTF,CommentVF,CommentLen,CommentTemplate,ChargeType,DiggByVisitor,DiggByIP,DiggRepeat,DiggPerTimes
		Dim FsoContentRule,FsoClassListRule,FsoClassPreTag,LatestNewDay,PubTimeLimit,AnnexPoint
		Dim ChannelID:ChannelID = KS.ChkClng(KS.G("ChannelID"))
		
	'	On Error Resume Next
	   If KS.G("Action")="Edit" Then
			SqlStr = "select * from KS_Channel Where ChannelID=" & ChannelID
			Set RS = Server.CreateObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1,1
			ChannelName   = RS("ChannelName")
			ModelEname    = RS("ModelEname")
			ChannelTable  = RS("ChannelTable")
			ItemName      = RS("ItemName")
			ItemUnit      = RS("ItemUnit")
			ChannelStatus = RS("ChannelStatus")
			StaticTF      = RS("StaticTF")
			FsoFolder     = RS("FsoFolder")
			FsoListNum    = RS("FsoListNum")
			WapSwitch     = RS("WapSwitch")
			Descript      = RS("Descript")
			BasicType     = RS("BasicType")
			UserTF        = RS("UserTF")
			UserClassStyle= RS("UserClassStyle")
			UserEditTF    = RS("UserEditTF")
			FsoHtmlTF     = RS("FsoHtmlTF")
			UpFilesTF     = RS("UpFilesTF")
			UpfilesDir    = RS("UpfilesDir")
			UserUpFilesTF = RS("UserUpFilesTF")
			UserUpfilesDir= RS("UserUpfilesDir")
			UserSelectFilesTF =RS("UserSelectFilesTF")
			UpfilesSize   = RS("UpfilesSize")
			AllowUpPhotoType = RS("AllowUpPhotoType")
			AllowUpFlashType = RS("AllowUpFlashType")
			AllowUpMediaType = RS("AllowUpMediaType")
			AllowUpRealType  = RS("AllowUpRealType")
			AllowUpOtherType = RS("AllowUpOtherType")
			ThumbnailsConfig = RS("ThumbnailsConfig")&"|0|||||||||||||"
			
			UserAddMoney     = RS("UserAddMoney")
			UserAddPoint     = RS("UserAddPoint")
			UserAddScore     = RS("UserAddScore")
			RefreshFlag      = RS("RefreshFlag")
			MaxPerPage       = RS("MaxPerPage")
			InfoVerificTF    = RS("InfoVerificTF")
			VerificCommentTF = RS("VerificCommentTF")
			CommentVF        = RS("CommentVF")
			CommentLen       = RS("CommentLen")
			CommentTemplate  = RS("CommentTemplate")
			WapSearchTemplate= RS("WapSearchTemplate")
			ChargeType       = RS("ChargeType")
			AnnexPoint       = RS("AnnexPoint")
			DiggByVisitor    = RS("DiggByVisitor")
			DiggByIP         = RS("DiggByIP")
			DiggRepeat       = RS("DiggRepeat")
			DiggPerTimes     = RS("DiggPerTimes")
			FsoContentRule   = RS("FsoContentRule")
			FsoClassListRule = RS("FsoClassListRule")
			FsoClassPreTag   = RS("FsoClassPreTag")
			LatestNewDay     = RS("LatestNewDay")
			PubTimeLimit     = RS("PubTimeLimit")
			ModelShortName   = RS("ModelShortName")
			ModelIco         = RS("ModelIco")
		Else
			  ChannelStatus =1 
			  ThumbnailsConfig="0.3|130|90|||||||||"
			  UpfilesDir  = "Upfiles/"
			  UserUpfilesDir = "User/"
			  UpfilesSize = 1024
			  BasicType   = 1
			  MaxPerPage=20
			  FsoFolder="html/"
			  RefreshFlag=2
			  InfoVerificTF=1
			  VerificCommentTF=0
			  UserTF=1
			  UserEditTF=0
			  UserClassStyle=1
			  UpFilesTF=1
			  AllowUpPhotoType = "gif|jpg|png"
			  AllowUpFlashType = "swf"
			  AllowUpMediaType = "mid|mp3|wmv|asf|avi|mpg"
			  AllowUpRealType  = "ram|rm|ra"
			  AllowUpOtherType = "rar|doc|zip"
			  WapSwitch = 1
			  ChargeType=1
			  FsoListNum=3
			  DiggByVisitor    = 0
			  DiggRepeat       = 0
			  DiggPerTimes     = 1
			  FsoClassPreTag="list"
			  FsoClassListRule = "1"
			  FsoContentRule   = "{$ClassEname}_{$ClassID}_"
			  LatestNewDay     = 3 
			  PubTimeLimit     = 20
			  AnnexPoint       = 0
			  ModelIco         = "/user/images/icon13.png"
		End If
			  ThumbnailsConfig=Split(ThumbnailsConfig&"|||||||||||||||||","|")
			  If Ubound(ThumbnailsConfig)<2 Then
			   ThumbnailsConfig(0)=0.3
			   ThumbnailsConfig(1)=130
			   ThumbnailsConfig(2)=90
			   ThumbnailsConfig(3)=0
			   ThumbnailsConfig(4)=0
			  End IF
		With Response
		.Write "<html>"&_
		"<title>模型基本参数设置</title>" &_
		"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">" &_
		"<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"&_
		"<script src=""../KS_Inc/JQuery.js"" language=""JavaScript""></script>"&_
		"<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>" & _
		"<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & _
		"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"&_
		"<body>" &_
		"<table width='100%' border='0' cellspacing='0' cellpadding='0'>"&_
		"  <tr>"&_
		"	<td height='25' class='sort'>网站模型管理</td>"&_
		" </tr>"&_
		" <tr><td height=5></td></tr>"&_
		"</table>" & _
		"<div class=tab-page id=modelpane>"& _
		"<form id=""myform"" name=""myform"" method=""post"" action=""KS.Model.asp?Action=EditSave&ChannelID=" & ChannelID & """ onSubmit=""return(CheckForm())"">" & _
        " <SCRIPT type=text/javascript>"& _
        "   var tabPane1 = new WebFXTabPane( document.getElementById( ""modelpane"" ), 1 )"& _
        " </SCRIPT>"& _
             
		" <div class=tab-page id=site-page>"& _
		"  <H2 class=tab>基本信息</H2>"& _
		"	<SCRIPT type=text/javascript>"& _
		"				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"& _
		"	</SCRIPT>" & _
		"<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle'> <div align=""right""><strong>模型状态：</strong></div></td>"
		.Write "      <td height=""30""><input type=""radio"" name=""ChannelStatus"" value=""1"" "
		If ChannelStatus = 1 Then .Write (" checked")
		.Write ">"
		.Write "正常"
		.Write "  <input type=""radio"" name=""ChannelStatus"" value=""0"" "
		If ChannelStatus = 0 Then .Write (" checked")
		.Write ">"
		.Write "关闭</td>"
		.Write "    </tr>"
		If KS.WSetting(0)="1" Then
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle'> <div align=""right""><strong>3G版状态：</strong></div></td>"
		.Write "      <td height=""30""><input type=""radio"" name=""WapSwitch"" value=""1"" "
		If WapSwitch = 1 Then .Write (" checked")
		.Write ">"
		.Write "正常"
		.Write "  <input type=""radio"" name=""WapSwitch"" value=""0"" "
		If WapSwitch = 0 Then .Write (" checked")
		.Write ">"
		.Write "关闭</td>"
		.Write "    </tr>"
		End If
%>
		<script type="text/javascript">
		 $(document).ready(function(){
		   $("input[name=FsoHtmlTF]").click(function(){
		     FsoDisplay();
			});
		   FsoDisplay();
		 });
		 function FsoDisplay()
		 {
		   var FsoHtmlTF=$("input[name=FsoHtmlTF][checked=true]").val();
		   if (FsoHtmlTF==0){
		    $("#fsoarea").hide();
			$("#staticarea").show();
		   }else if(FsoHtmlTF==1){
		    $("#fsoarea").show();
		    $("#staticarea").hide();
		   }else{
		    $("#fsoarea").show();
			$("#staticarea").show();
		   }
		 }
	
		 function CheckForm()
		 {  
		  if ($("input[name=ChannelName]").val()=="")
		  {
		    $("input[name=ChannelName]").focus();
		   alert('请输入模型名称');
		   return false;
		  }
		  if ($("input[name=ModelEname]").val()=="")
		  {
		   $("input[name=ModelEname]").focus()
		   alert('请输入模型的目录名称');
		   return false;
		  }
		  if ($("input[name=ChannelTable]").val()=="")
		  {
		     $("input[name=ChannelTable]").focus()
			 alert('请输入数据名！');
			 return false;
		  }
		  if ($("input[name=ItemName]").val()=="")
		  {
		     $("input[name=ItemName]").focus()
			 alert('请输入项目名称！');
			 return false;
		  }
		  if ($("input[name=ItemUnit]").val()=="")
		  {
		     $("input[name=ItemUnit]").focus();
			 alert('请输入项目单位！');
			 return false;
		  }
		  if ($("input[name=FsoFolder]").val()=="")
		  {
		     $("input[name=FsoFolder]").focus();
			 alert('请输入模型目录！');
			 return false;
		  }
		  $("#myform").submit();
		 }
		 function GetTable(val)
		 { 
		    $.get('../plus/ajaxs.asp', { foldername: escape($('input[name=ChannelName]').val()), action: 'Ctoe' },function(data){
			$('input[name=ChannelTable]').val(unescape(data));
		    $('input[name=ModelEname]').val(unescape(data));
		    $('input[name=FsoFolder]').val('html/'+data+'/');
		    $('input[name=UserUpfilesDir]').val(data+'/');
		    $('input[name=UpfilesDir]').val('upfiles/'+data+'/');
			 });
		 }
		</script>
		<style type="text/css">
		 .textbox{
		 border:0px;border-bottom:1px solid #000;width:60px;background:transparent
		 }
		.tips {color: #999999;padding:2px}
		.txt {color: #666;border:1px solid #ccc;height:22px;line-height:22px}
		textarea {color: #666;border:1px solid #ccc;}
		</style>
		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>模型名称：</strong></div></td>      
			<td height="30"> <input class="txt" name="ChannelName" type="text" <%If KS.G("Action")<>"Edit" Then Response.Write " onkeyup='GetTable(this.value)'"%> value="<%=ChannelName%>" size="30"></td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>模型目录：</strong></div></td>      
			<td height="30"> <input class="txt" name="ModelEname" type="text"<%If KS.G("Action")="Edit" Then Response.Write " Disabled"%> value="<%=ModelEname%>" size="30"> <span class="tips">*只能用字母和数字的组合，且不能修改</span></td> 
		</tr>

		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">     
		<td height="30" class="clefttitle"align="right"><div><strong>数据表名称：</strong></div></td>     
		<td height="30"><%If KS.G("Action")="Add" Then Response.Write " KS_U_" %><input name="ChannelTable" id='ChannelTable' type="text" value="<%=ChannelTable%>" class="txt" size="14"<%If KS.G("Action")="Edit" Then Response.Write " Disabled"%>><font class="tips">说明：创建数据表后无法修改，并且用户创建的数据表以"KS_U_"开头</font></td>   
		</tr> 
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>基 类 型：</strong></div></td>      
			<td height="30"> 
			<select name="BasicType" id="BasicType" <%If KS.G("Action")="Edit" Then Response.Write " Disabled"%>>
			 <option value=1<%if BasicType="1" Then Response.Write " selected"%>>文章类型</option>
			 <option value=2<%if BasicType="2" Then Response.Write " selected"%>>图片类型</option>
			 <option value=3<%if BasicType="3" Then Response.Write " selected"%>>软件类型</option>
			 <%if instr(ChannelNotOnStr,"4")=0 then%>
			<option value=4<%if BasicType="4" Then Response.Write " selected"%>>动漫类型</option>
			 <%end if%>
			 <%If KS.G("Action")="Edit" Then%>
			 <option value=5<%if BasicType="5" Then Response.Write " selected"%>>商城类型</option>
			 <option value=6<%if BasicType="6" Then Response.Write " selected"%>>音乐类型</option>
			 <option value=7<%if BasicType="7" Then Response.Write " selected"%>>影视类型</option>
			 <option value=8<%if BasicType="8" Then Response.Write " selected"%>>供求类型</option>
			 <option value=9<%if BasicType="9" Then Response.Write " selected"%>>考试类型</option>
			 <%End If%>
			</select>
			</td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>项目名称：</strong></div></td>      
			<td height="30"> <input class="txt" name="ItemName" id="ItemName" type="text" value="<%=ItemName%>" size="30"> <span class="tips">*如：文章、图片、软件等项</span></td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>项目单位：</strong></div></td>      
			<td height="30"> <input name="ItemUnit" type="text" value="<%=ItemUnit%>" class="txt" size="8"> <span class="tips">*如：篇、个、本等</span></td> 
		</tr>
		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="clefttitle" align="right"><div><strong>模型描述：</strong></div></td>      
			<td height="30"> <textarea name="Descript" rows=4 cols=80><%=Descript%></textarea></td> 
		</tr>
		</table>
		</div>
		
		
		<%
		If ChannelID=9 Then
		.Write " <div class=tab-page id=fso-page>"
		.Write "  <H2 class=tab>系统选项</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""fso-page"" ) );"
		.Write "	</SCRIPT>"

		.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
	    .Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='clefttitle' align='right'><div><strong>生成的总目录：</strong></div></td>"      
		.Write "	<td height='30'> <input class='txt' name='FsoFolder' type='text' value='" & FsoFolder & "' size='20'><span class='tips'>*用于生成静态html存放的目录，只能以字母和数字的组合,必须以""/""结束</span></td> "
		.Write "</tr>"
	    .Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='clefttitle' align='right'><div><strong>日常练习免费练习次数：</strong></div></td>"      
		.Write "	<td height='30'> <input class='txt' style='text-align:center' name='FsoListNum' type='text' value='" & FsoListNum & "' size='6'><span class='tips'>次</span></td> "
		.Write "</tr>"
	    .Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='clefttitle' align='right'><div><strong>日常练习收费设置：</strong></div></td>"      
		.Write "	<td height='30'><span class='tips'>超过上面设置的免费次数将按<input class='txt' style='text-align:center' name='FsoClassListRule' type='text' value='" &FsoClassListRule & "' size='3'>元/次收取</span></td> "
		.Write "</tr>"
	    .Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='clefttitle' align='right'><div><strong>日常练习随机抽题数：</strong></div></td>"      
		.Write "	<td height='30'> <input maxlength='3' class='txt' style='text-align:center' name='FsoHtmlTF' type='text' value='" & FsoHtmlTF & "' size='6'><span class='tips'> 题</span></td> "
		.Write "</tr>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td height=""30"" width='160' class='clefttitle'><div align=""right""><strong>后台添加试卷启用word存图选项：</strong></div></td>"
		.Write "      <td height=""30"">"
		
		.Write "  <input type=""radio"" name=""FsoContentRule"" value=""1"" "
		If FsoContentRule = "1" Then .Write (" checked")
		.Write ">启用&nbsp;&nbsp;"

		.Write "<input type=""radio"" name=""FsoContentRule"" value=""0"" "
		If FsoContentRule = "0" Then .Write (" checked")
		.Write ">"
		.Write "不启用<br/><span class='tips'>说明启用此功能，需要安装word存图插件(<a href='#' target='_blank'>查看</a>)，并且服务器需要支持asp.net 2.0环境。</span>"
		
		.Write "     </td>"
		.Write "    </tr>"
        .Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='clefttitle' align='right'><div><strong>查看试题解释需要消费：</strong></div></td>"      
		.Write "	<td height='30'> <label><input type='radio' name='jsxh'"
		if KS.ChkClng(ThumbnailsConfig(6))="0" then .write " checked"
		.Write " value='0'>不需要</label><br/>"
		.Write "<label><input type='radio' name='jsxh'"
		if KS.ChkClng(ThumbnailsConfig(6))="1" then .write " checked"
		.Write " value='1'>看一份试卷只需消费一次</label><br/>"
		.Write "<label><input type='radio' name='jsxh'"
		if KS.ChkClng(ThumbnailsConfig(6))="2" then .write " checked"
		.Write " value='2'>看一题消费一次</label>"
		.Write "<div>查看试题解释每次消费<input class='txt' type='text' name='jsxhnum' size='4' style='text-align:center' value='" & ThumbnailsConfig(7) &"'/>元  间隔<input class='txt' type='text' name='rxhhour' size='4' style='text-align:center' value='" & ThumbnailsConfig(8) &"'/>小时后需要重复收复，不重复收费输入0"
		
		.Write "</td></tr>"
	    .Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='clefttitle' align='right'><div><strong>答题鼓励语言设置：</strong></div></td>"      
		.Write "	<td height='30'> 正确：<textarea class='txt' style='width:260px;height:120px' name='rightyy'>" & ThumbnailsConfig(9)& "</textarea><br/>错误：<textarea class='txt' style='width:260px;height:120px' name='wrongyy'>" & ThumbnailsConfig(10)& "</textarea><br/><span class='tips'>一行一个鼓励语言。</span></td> "
		.Write "</tr>"
		.Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='clefttitle' align='right'><div><strong>考试需要输验证码：</strong></div></td>"      
		.Write "	<td height='30'> <label><input type='radio' name='ksyzm'"
		if KS.ChkClng(ThumbnailsConfig(11))="0" then .write " checked"
		.Write " value='0'>不需要</label>&nbsp;"
		.Write "<label><input type='radio' name='ksyzm'"
		if KS.ChkClng(ThumbnailsConfig(11))="1" then .write " checked"
		.Write " value='1'>需要</label><br/>"
		
		.Write "</td></tr>"
		.Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='clefttitle' align='right'><div><strong>考试首页分类从第：</strong></div></td>"      
		.Write "	<td height='30'> <select name=""sjfl"">"
		 for i=1 to conn.execute("select max(tj) from ks_sjclass")(0)
		  if KS.ChkClng(ThumbnailsConfig(12))=i then
		  .write "<option value=" & i & " selected>第" & I & "层</option>"
		  else
		  .write "<option value=" & i & ">第" & I & "层</option>"
		  end if
		 next
		.Write "<select>生成。<span class='tips'> 指用标签{$GetClass}调用生成的分类，从第几层开始生成，建议选择第二层</span></td></tr>"

		.Write "</table>"
		.Write "</div>"
		Else
		.Write " <div class=tab-page id=fso-page>"
		.Write "  <H2 class=tab>生成选项</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""fso-page"" ) );"
		.Write "	</SCRIPT>"

		.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
		.Write "    <tr class='tdbg'>"
		.Write "      <td height=""30"" class='clefttitle'><div align=""right""><strong>本模型运行模式：</strong></div></td>"
		.Write "      <td height=""30"">"
		
		.Write "  <input type=""radio"" name=""FsoHtmlTF"" value=""0"" "
		If FsoHtmlTF = 0 Then .Write (" checked")
		.Write ">动态asp<br/>"

		.Write "<input type=""radio"" name=""FsoHtmlTF"" value=""1"" "
		If FsoHtmlTF = 1 Then .Write (" checked")
		.Write ">"
		.Write "栏目页及内容页都生成HTML<br/>"
		
		.Write "<input type=""radio"" name=""FsoHtmlTF"" value=""2"" "
		If FsoHtmlTF = 2 Then .Write (" checked")
		.Write ">"
		.Write "栏目页不生成,内容页生成HTML(<font color=red>推荐</font>)<br/>"
		
		
		.Write "     </td>"
		.Write "    </tr>"
		.Write "    <tbody id='staticarea'>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td height='25' class='clefttitle' align=""right""><strong>伪静态设置：</strong></td>"
		.Write "      <td>"
		
		.Write "  <input type=""radio"" name=""StaticTF"" value=""0"" "
		If StaticTF = 0 Then .Write (" checked")
		.Write ">"
		.Write "不启用"
		.Write "  <input type=""radio"" name=""StaticTF"" value=""1"" "
		If StaticTF = 1 Then .Write (" checked")
		.Write ">"
		.Write "伪静态(带问号,不需要装组件)"
		.Write "  <input type=""radio"" name=""StaticTF"" value=""2"" "
		If StaticTF = 2 Then .Write (" checked")
		.Write ">"
		.Write "伪静态(需要装ISAPI_Rewrite组件)"
        .Write "<br /><font class='tips'>这里需要设置不生成静态才有效,建议流量大的网站直接启用全部生成静态,而不是使用伪静态</font>"
		.Write "      </td>"
		.Write "    </tr>"
		.Write " </tbody>"
		.Write "  <tbody id='fsoarea'>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' align=""right""><strong>后台添加文档，同时发布选项：</strong></td>"
		.Write "      <td height=""30""> <input type=""radio"" name=""RefreshFlag"" value=""1"" "
		If RefreshFlag = 1 Then .Write (" checked")
		.Write ">"
		.Write "仅发布内容页 <br>"
		.Write "          <input type=""radio"" name=""RefreshFlag"" value=""2"" "
		If RefreshFlag = 2 Then .Write (" checked")
		.Write ">发布栏目页+内容页<font color=red>(建议)</font><br>"		
		.Write "          <input type=""radio"" name=""RefreshFlag"" value=""3"" "
		If RefreshFlag = 3 Then .Write (" checked")
		.Write ">发布首页+栏目页+内容页"
		.Write "        </td>"
		.Write "    </tr>"	
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle'><div align=""right""><strong>自动生成列表分页数：</strong></div></td>"
		.Write "      <td height=""30""><input class='txt' type='text' value='" & FsoListNum & "' name='FsoListNum' size='6' style='text-align:center'><font class='tips'>这里设置生成栏目列表分页时自动生成的分页数，如果你的网站数据量较大，建议输入一个较小的数字，小数据量的网站可以不用限制，直接设置为0</font></td>"
		.Write "    </tr>"	
	    .Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='clefttitle' align='right'><div><strong>生成的总目录：</strong></div></td>"      
		.Write "	<td height='30'> <input class='txt' name='FsoFolder' type='text' value='" & FsoFolder & "' size='20'><span class='tips'>*用于生成静态html存放的目录，只能以字母和数字的组合,必须以""/""结束</span></td> "
		.Write "</tr>"
		.Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='clefttitle' align='right'><div><strong>生成的栏目页规则：</strong></div></td>"   
		.Write "<td>"   
		.Write "<input type=""radio"" name=""FsoClassListRule"" value=""1"" "
		If FsoClassListRule = 1 Then .Write (" checked")
		.Write ">按目录级别顺序结构生成列表页<br>"
		.Write " &nbsp;<font color=blue>如：第1页为/article/aaa/bbb/ccc/index.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/article/aaa/bbb/ccc/index_2.html</font>"
		.Write "  <br><input type=""radio"" name=""FsoClassListRule"" value=""2"" "
		If FsoClassListRule = 2 Then .Write (" checked")
		.Write ">所有栏目页都生成在模型总生成目录下面<font color=red>(有利于SEO)</font><br>"
		.Write " &nbsp;<font color=green>如栏目ID号为100则生成如下：</font><font color=blue><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第1页为/总生成目录/list_100.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/总生成目录/list_100_2.html</font>"
		.Write " <br>&nbsp;<font color=green>如栏目ID号为101则生成如下：</font><font color=blue><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第1页为/总生成目录/list_101.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/总生成目录/list_101_2.html</font>"
		.Write "  <br><input type=""radio"" name=""FsoClassListRule"" value=""3"" "
		If FsoClassListRule = 3 Then .Write (" checked")
		.Write ">本模型下的一级栏目生成在本频道下的Index.html,子栏目按如下规则生成<br>"
		.Write " &nbsp;<font color=green>如一级栏目 ""教育频道"" 英文名称：""edu"",那么生成如下：</font><font color=blue><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第1页为/总生成目录/edu/index.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/总生成目录/edu/index_2.html</font>"
		.Write " <br>&nbsp;<font color=green>二级及以上的栏目(即""教育频道"")下的栏目,如栏目ID号为101则生成如下：</font><font color=blue><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第1页为/总生成目录/edu/list_101.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/总生成目录/edu/list_101_2.html</font>"
		.Write "  <br><input type=""radio"" name=""FsoClassListRule"" value=""4"" "
		If FsoClassListRule = 4 Then .Write (" checked")
		.Write ">所有栏目页都生成在模型总生成目录下面<font color=red>(新增）</font><br>"
		.Write " <font color=blue>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第1页为/总生成目录/自定义列表前缀.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/总生成目录/自定义列表前缀_2.html</font>"
		
		.Write "</td>"
		.Write "	<td height='30'> </td> "
		.Write "</tr>"

		.Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='clefttitle' align='right'><div><strong>生成的列表页的前缀字符：</strong></div></td>"      
		.Write "	<td height='30'> <input class='txt' name='FsoClassPreTag' type='text' value='" & FsoClassPreTag & "' size='30'> <span class='tips'>*如list,show等</span><br/><span class='tips'>可用标签：<br/>{$ClassEname}-本栏目英文名<br/>{$ClassID}-本栏目小ID<br/> {$BigClassID}-本栏目大ID<br/>{$TopClassEname}-一级栏目英文名</td> "
		.Write "</tr>"


		.Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='clefttitle' align='right'><div><strong>生成的内容页目录规则：</strong></div></td>"      
		.Write "	<td height='30'> <input class='txt' name='FsoContentRule' type='text' value='" & FsoContentRule & "' size='30'>&nbsp;"
		.Write "     <select name='srule' onchange='if (this.value!=""""){ $(""input[name=FsoContentRule]"").val(this.value);}'><option value=''>------快速选择内容页生成结构------</option>"
		.Write "     <option value='View_'>View_</option>"
		.Write "     <option value='{$ClassDir}'>{$ClassDir}</option>"
		.Write "     <option value='{$ChannelEname}/{$ClassEname}_{$ClassID}_'>{$ChannelEname}/{$ClassEname}_{$ClassID}_</option>"
		.Write "     <option value='{$ClassEname}_{$ClassID}_'>{$ClassEname}_{$ClassID}_(推荐)</option>"
		.Write "     <option value='{$ClassDir}{$ClassEname}_{$ClassID}_'>{$ClassDir}{$ClassEname}_{$ClassID}_</option>"
		.Write "   </select><br><font color=red>可选项（允许留空）</font><br><span class='tips'>可用标签：一级频道名称{$ChannelEname},栏目路径{$ClassDir} {$ClassID} {$ClassEname} {$InfoID}</span><br> "
		.Write "   <br>"
		.Write " </td></tr>"
		.Write "</tbody>"
        .Write "</table>"
		.Write "</div>"
	End If
		.Write " <div class=tab-page id=upfile-page>"
		.Write "  <H2 class=tab>上传选项</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""upfile-page"" ) );"
		.Write "	</SCRIPT>"

		.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
		.Write "    <tr class='tdbg'>"
		.Write "      <td width='160' class='clefttitle'><div align=""right""><strong>管理员是否允许上传文件：</strong></div></td>"
		.Write "      <td height=""30""><input type=""radio"" name=""UpFilesTF"" value=""1"" "
		If UpFilesTF = 1 Then .Write (" checked")
		.Write ">"
		.Write "允许"
		.Write "  <input type=""radio"" name=""UpFilesTF"" value=""0"""
		If UpFilesTF = 0 Then .Write (" checked")
		.Write ">"
		.Write "  不允许</td>"
		.Write "    </tr>"
		.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
		.Write "            <td height='30' width='160' align='right' class='clefttitle'><strong>缩略图选项：</strong><br><font color=blue>当基本信息设置开启自动生成缩略图功能时才有效</font></td>" & vbCrLf
		.Write "            <td height='28'>&nbsp;"
					    
		.Write "黄金分割点：<input class='txt' type='text' value='" & ThumbnailsConfig(0) & "' name='GoldenPoint' size='4' style='text-align:center'> 宽度：<input class='txt' type='text' value='" & ThumbnailsConfig(1) & "' name='ThumbsWidth' size='8' style='text-align:center'>px 高度：<input class='txt' type='text' value='" & ThumbnailsConfig(2) & "' name='ThumbsHeight' size='8' style='text-align:center'>px"
		.Write "              <br/> <font color=red>tips:如果高度设置为0,则生成的高度将由您设置的宽度自动约束决定(类似photoshop软件的自动约束)</font></td>"
		.Write "          </tr>" & vbCrLf
		.Write "    <tr class='tdbg' style='display:none'>"
		.Write "      <td  width='160' class='clefttitle'><div align=""right""><strong>后台文件上传目录：</strong></div></td>"
		.Write "      <td height=""30""> <input name=""UpfilesDir"" class='txt' type=""text"" id=""UpfilesDir"" value=""" & UpfilesDir & """ size=""30""></td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td  width='160' class='clefttitle'><div align=""right""><strong>会员中心上传设置：</strong></div></td>"
		.Write "      <td height=""30"">"
		.Write "  <input type=""radio"" name=""UserUpFilesTF"" value=""0"""
		If UserUpFilesTF = 0 Then .Write (" checked")
		.Write ">关闭上传<br/>"
		.Write "<input type=""radio"" name=""UserUpFilesTF"" value=""1"" "
		If UserUpFilesTF = 1 Then .Write (" checked")
		.Write ">只允许会员上传<br/>"
		.Write "<input type=""radio"" name=""UserUpFilesTF"" value=""2"" "
		If UserUpFilesTF = 2 Then .Write (" checked")
		.Write ">允许所有人上传，包括游客(匿名投稿)"
		
		.Write "  </td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg' style='display:none'>"
		.Write "      <td  width='160' class='clefttitle'><div align=""right""><strong>会员文件上传目录：</strong></div></td>"
		.Write "      <td height=""30""> <input class='txt' name=""UserUpfilesDir"" type=""text"" id=""UserUpfilesDir"" value=""" & UserUpfilesDir & """ size=""30""><br><b>提示：</b><br><font color=red>1、会员目录构成规则：系统设置的总上传目录/User/会员名称;<br>2、上传目录必须以/结束;</font></td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td  width='160' class='clefttitle'><div align=""right""><strong>允许会员选择上传文件：</strong></div></td>"
		.Write "      <td height=""30""><input type=""radio"" name=""UserSelectFilesTF"" value=""1"" "
		If UserSelectFilesTF = 1 Then .Write (" checked")
		.Write ">"
		.Write "允许"
		.Write "  <input type=""radio"" name=""UserSelectFilesTF"" value=""0"""
		If UserSelectFilesTF = 0 Then .Write (" checked")
		.Write ">"
		.Write "  不允许</td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td  width='160' class='clefttitle'><div align=""right""><strong>允许上传的最大文件大小：</strong></div></td>"
		.Write "      <td height=""30""><input name=""UpfilesSize"" class='txt' onBlur=""CheckNumber(this,'允许上传最大文件大小');"" type=""text"" id=""UpfilesSize"" value=""" & UpfilesSize & """ size=""10"">"
		.Write "      KB 　 <font color='#ff0000'>提示：1 KB = 1024 Byte，1 MB = 1024 KB</font></td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td  width='160' class='clefttitle'><div align=""right""><strong>允许上传的文件类型：</strong><BR>"
		.Write "          <font color='#ff0000'>多种文件类型之间以""|""分隔</font></div></td>"
		.Write "      <td height=""30""><table width=""98%"" border=""0"">"
		.Write "        <tr>"
		.Write "          <td width=""19%"" height=""25"" align=""right"">图片类型:</td>"
		.Write "          <td width=""81%""><input class='txt' name=""AllowUpPhotoType"" type=""text"" id=""AllowUpPhotoType"" value=""" & AllowUpPhotoType & """ size=""30""></td>"
		.Write "        </tr>"
		.Write "        <tr>"
		.Write "          <td height=""25"" align=""right"">Flash 文件:</td>"
		.Write "          <td><input class='txt' name=""AllowUpFlashType"" type=""text"" id=""AllowUpFlashType"" value=""" & AllowUpFlashType & """ size=""30""></td>"
		.Write "        </tr>"
		.Write "        <tr>"
		.Write "          <td height=""25"" align=""right"">Windows 媒体: </td>"
		.Write "          <td><input class='txt'  name=""AllowUpMediaType"" type=""text"" id=""AllowUpMediaType"" value=""" & AllowUpMediaType & """ size=""30""></td>"
		.Write "        </tr>"
		.Write "        <tr>"
		.Write "          <td height=""25"" align=""right"">Real 媒体: </td>"
		.Write "          <td><input class='txt' name=""AllowUpRealType"" type=""text"" id=""AllowUpRealType"" value=""" & AllowUpRealType & """ size=""30""></td>"
		.Write "        </tr>"
		.Write "        <tr>"
		.Write "          <td height=""25"" align=""right"">其它文件:</td>"
		.Write "          <td><input class='txt' name=""AllowUpOtherType"" type=""text"" id=""AllowUpOtherType"" value=""" & AllowUpOtherType & """ size=""30""></td>"
		.Write "        </tr>"
		.Write "      </table></td>"
		.Write "    </tr>"
        .Write "</table>"
		.Write "</div>"

		.Write " <div class=tab-page id=tougao-page>"
		.Write "  <H2 class=tab>投稿选项</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "	  tabPane1.addTabPage( document.getElementById( ""tougao-page"" ) );"
		.Write "	</SCRIPT>"
		.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' style=""text-align:right""><strong>前台会员投稿总开关：</strong><br><font color=red>只有启用，会员才可在本模型里投稿。</font></td>"
		.Write "      <td height=""30"">"
		.Write "          <input type=""radio"" name=""UserTF"" value=""0"" "
		If UserTF = 0 Then .Write (" checked")
		.Write ">关闭会员投稿 <br>"
		.Write " <input type=""radio"" name=""UserTF"" value=""1"" "
		If UserTF = 1 Then .Write (" checked")
		.Write ">只允许注册会员可以投稿,具体可投稿的栏目依栏目设置而定<br/>"
		.Write " <input type=""radio"" name=""UserTF"" value=""2"" "
		If UserTF = 2 Then .Write (" checked")
		.Write ">允许所有用户投稿（包括游客）,具体可投稿的栏目依栏目设置而定<br>"
		.Write "        </td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' style=""text-align:right""><strong>会员中心投稿菜单显示名称：</strong></td>"
		.Write "      <td height=""30""> <input class='txt' name=""ModelShortName"" type=""text"" id=""ModelShortName"" value=""" & ModelShortName & """ size=""16""> <span class='tips'>如：文章，新闻，房源等，建议取两个汉字名称</span></td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' style=""text-align:right""><strong>会员中心投稿菜单图标地址：</strong></td>"
		.Write "      <td height=""30""> <input class='txt' name=""ModelIco"" type=""text"" id=""ModelIco"" value=""" & ModelIco & """ size=""30""> <span class='tips'>如：/user/images/ico1.gif</span></td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg' style='color:blue'>"
		.Write "      <td class='clefttitle' style=""text-align:right""><strong>新注册会员：</strong></td>"
		.Write "      <td height=""30""> <input class='txt' name=""PubTimeLimit"" type=""text"" id=""PubTimeLimit"" value=""" & PubTimeLimit & """ size=""6"">分钟后才可以在此模型投稿</td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' style=""text-align:right""><strong>会员投稿增加：</strong></td>"
		.Write "      <td height=""30""> 资金<input class='txt' style='text-align:center' name=""UserAddMoney"" type=""text"" id=""UserAddMoney"" value=""" & UserAddMoney & """ size=""6"">元  点券<input class='txt' style='text-align:center' name=""UserAddPoint"" type=""text"" id=""UserAddPoint"" value=""" & UserAddPoint & """ size=""6"">点  积分<input class='txt'  name=""UserAddScore"" type=""text"" id=""UserAddScore"" value=""" & UserAddScore & """ style='text-align:center' size=""6"">分<br/><font color=green>为0时不增加,可设置成负数,表示投稿要消费</font></td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' style=""text-align:right""><strong>允许会员刷新添加时间：</strong></td>"
		.Write "      <td height=""30""> "
		.Write " <input type=""radio"" name=""RefreshTimeTF"" value=""0"" "
		If ThumbnailsConfig(3) = "0" Then .Write (" checked")
		.Write ">不允许 "
		.Write "          <input type=""radio"" name=""RefreshTimeTF"" value=""1"" "
		If ThumbnailsConfig(3) = "1" Then .Write (" checked")
		.Write ">允许</td>"
		.Write "    </tr>"
		

		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' style=""text-align:right""><strong>审核过的稿件是否允许修改：</strong></td>"
		.Write "      <td height=""30"">"
		.Write " <input type=""radio"" name=""UserEditTF"" value=""0"" "
		If UserEditTF = 0 Then .Write (" checked")
		.Write ">不允许<font color=red>(建议)</font><br>"
		.Write "          <input type=""radio"" name=""UserEditTF"" value=""1"" "
		If UserEditTF = 1 Then .Write (" checked")
		.Write ">允许，但修改后自动转为未审(<font color=red>如果投稿要增加积分等,会导致重复收费</font>)<br>"
		.Write "          <input type=""radio"" name=""UserEditTF"" value=""2"" "
		If UserEditTF =2 Then .Write (" checked")
		.Write ">允许，修改后仍为已审状态（不推荐,<font color=red>如果投稿要增加积分等,会导致重复收费</font>）"
        .Write "      </td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' style=""text-align:right""><strong>投稿栏目显示方式：</strong></td>"
		.Write "      <td height=""30"">"
		.Write " <input type=""radio"" name=""UserClassStyle"" value=""0"" "
		If UserClassStyle = 0 Then .Write (" checked")
		.Write ">仅显示有权限的栏目（下拉方式）<br>"
		.Write "          <input type=""radio"" name=""UserClassStyle"" value=""1"" "
		If UserClassStyle = 1 Then .Write (" checked")
		.Write ">仅显示有权限的栏目（跳窗方式）<br>"
		.Write "          <input type=""radio"" name=""UserClassStyle"" value=""2"" "
		If UserClassStyle = 2 Then .Write (" checked")
		.Write ">树型显示本模型下所有栏目,不允许投稿的栏目用灰色显示（跳窗方式）<br/>"
		.Write "          <input type=""radio"" name=""UserClassStyle"" value=""3"" "
		If UserClassStyle = 3 Then .Write (" checked")
		.Write "><font color=blue>多级联动下拉(<font color=red>新增</font>)（适合于只有二至三级栏目结构的模型）</font>"
        .Write "      </td>"
		.Write "    </tr>"
		
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' style=""text-align:right""><strong>会员中心发布的信息是否需要审核：</strong></td>"
		.Write "      <td height=""30"">"
		.Write " <input type=""radio"" name=""InfoVerificTF"" value=""0"" "
		If InfoVerificTF = 0 Then .Write (" checked")
		.Write ">需要后台人工审核<br>"
		.Write "          <input type=""radio"" name=""InfoVerificTF"" value=""1"" "
		If InfoVerificTF = 1 Then .Write (" checked")
		.Write ">不需要审核（但不直接生成内容页HTML）<br>"
		.Write "          <input type=""radio"" name=""InfoVerificTF"" value=""2"" "
		If InfoVerificTF = 2 Then .Write (" checked")
		.Write ">不需要审核（当有启用生成静态HTML，直接生成内容页）<br>"

		.Write "      </td>"
		.Write "    </tr>"
		
		If BasicType=1 Or BasicType=2 Or BasicType=3 Then
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' height='30'><strong>自动生成投稿录入表单:</strong></td>"
		.Write "      <td><label><input type='checkbox' name='autocreate' id='autocreate' value='1' onClick=""LoadTemplate(this.checked)"">自动生成</label> <font color=red>提示：第一次生成模板，可以点此自动生成！</font></td>"
		.Write "    </tr>"
		%>
		<script language = 'JavaScript'>
					function LoadTemplate(v)
					{   
					   if (v==true)
					    { 
							$.ajax({
								  url: 'KS.Model.asp',
								  cache: false,
								  data: "action=createtemplate&channelid=<%=ChannelID%>",
								  success: function(s){
									  $('#Content').val(s);
								  }
								});
							 return; 
						}
						else
						{
						  $('#Content').val('');
						}
					}	

		            function show_ln(txt_ln,txt_main){
			            var txt_ln  = document.getElementById(txt_ln);
			            var txt_main  = document.getElementById(txt_main);
			            txt_ln.scrollTop = txt_main.scrollTop;
			            while(txt_ln.scrollTop != txt_main.scrollTop)
			            {
				            txt_ln.value += (i++) + '\n';
				            txt_ln.scrollTop = txt_main.scrollTop;
			            }
			            return;
		            }
		            function editTab(){
			            var code, sel, tmp, r
			            var tabs=''
			            event.returnValue = false
			            sel =event.srcElement.document.selection.createRange()
			            r = event.srcElement.createTextRange()
			            switch (event.keyCode){
				            case (8) :
				            if (!(sel.getClientRects().length > 1)){
					            event.returnValue = true
					            return
				            }
				            code = sel.text
				            tmp = sel.duplicate()
				            tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)
				            sel.setEndPoint('startToStart', tmp)
				            sel.text = sel.text.replace(/\t/gm, '')
				            code = code.replace(/\t/gm, '').replace(/\r\n/g, '\r')
				            r.findText(code)
				            r.select()
				            break
			            case (9) :
				            if (sel.getClientRects().length > 1){
					            code = sel.text
					            tmp = sel.duplicate()
					            tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)
					            sel.setEndPoint('startToStart', tmp)
					            sel.text = '\t'+sel.text.replace(/\r\n/g, '\r\t')
					            code = code.replace(/\r\n/g, '\r\t')
					            r.findText(code)
					            r.select()
				            }else{
					            sel.text = '\t'
					            sel.select()
				            }
				            break
			            case (13) :
				            tmp = sel.duplicate()
				            for (var i=0; tmp.text.match(/[\t]+/g) && i<tmp.text.match(/[\t]+/g)[0].length; i++) tabs += '\t'
				            sel.text = '\r\n'+tabs
				            sel.select()
				            break
			            default  :
				            event.returnValue = true
				            break
				            }
			            }
		            //-->
		            </script>
		<tr class='tdbg'>
		      <td class='clefttitle' align="right"><strong>录入表单模板：</strong>
			   <br/><br/>
			   <font color="#999999">不想自定义可以留空,否则添加/变更字段需要重新生成表单模板</font>
			  </td>
		     <td height="280" nowrap>
			 <textarea id='txt_ln' name='rollContent' cols='6' style='overflow:hidden;height:280px;background-color:highlight;border-right:0px;text-align:right;font-family: tahoma;font-size:12px;font-weight:bold;color:highlighttext;cursor:default;' readonly><%
			Dim XmlForm:XmlForm=LFCls.GetConfigFromXML("modelinputform","/inputform/model",ChannelID)
			If KS.IsNul(XmlForm) Then XmlForm=""
			 
		 Dim N
		 For N=1 To 3000
			Response.Write N & "&#13;&#10;"
		 Next
		 On Error Resume Next
		 %>
		 </textarea>
		 <textarea name='Content' id="Content" style="width:570px;height:280px" ROWS='15' onkeydown='editTab()' onscroll="show_ln('txt_ln','Content')" wrap='on'><%=Server.HTMLEncode(XmlForm)%></textarea>
			 </td>
		   </tr>
		<%
		End If

        .Write "</table>"
		.Write "</div>"
		
		If ChannelID<>9 Then
		.Write " <div class=tab-page id=digg-page>"
		.Write "  <H2 class=tab>Digg选项</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""digg-page"" ) );"
		.Write "	</SCRIPT>"
		.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' width='160'><div align=""right""><strong>是否允许游客DIGG：</strong></div></td>"
		.Write "      <td height=""30"">"
		.Write " <input type=""radio"" name=""DiggByVisitor"" value=""1"" "
		If DiggByVisitor = 1 Then .Write (" checked")
		.Write ">允许"
		.Write "          <input type=""radio"" name=""DiggByVisitor"" value=""0"" "
		If DiggByVisitor = 0 Then .Write (" checked")
		.Write ">不允许"
		.Write "      </td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' width='160'><div align=""right""><strong>是否启用单IP限制：</strong><br><font color=red>若启用单IP限制，每个IP的用户只能对每个项目Digg一次</font></div></td>"
		.Write "      <td height=""30"">"
		.Write " <input type=""radio"" name=""DiggByIP"" value=""1"" "
		If DiggByIP = 1 Then .Write (" checked")
		.Write ">启用"
		.Write "          <input type=""radio"" name=""DiggByIP"" value=""0"" "
		If DiggByIP = 0 Then .Write (" checked")
		.Write ">不启用"
		.Write "      </td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' width='160'><div align=""right""><strong>会员是否允许重复DIGG：</strong><br><font color=red>启用IP限制时，始终不允许</font></div></td>"
		.Write "      <td height=""30"">"
		.Write " <input type=""radio"" name=""DiggRepeat"" value=""1"" "
		If DiggRepeat = 1 Then .Write (" checked")
		.Write ">允许"
		.Write "          <input type=""radio"" name=""DiggRepeat"" value=""0"" "
		If DiggRepeat = 0 Then .Write (" checked")
		.Write ">不允许"
		.Write "      </td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' width='160'><div align=""right""><strong>次数选项：</strong></div></td>"
		.Write "      <td height=""30"">"
		.Write "         每DIGG一下自动增加<input type=""text"" class='textbox' size=""6"" style=""text-align:center""  name=""DiggPerTimes"" value=""" & DiggPerTimes & """>次 "
		.Write "      </td>"
		.Write "    </tr>"
		
        .Write "</table>"
        .Write "</div>"	
	End If	 
		 
		.Write " <div class=tab-page id=detail-page>"
		.Write "  <H2 class=tab>其它参数</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""detail-page"" ) );"
		.Write "	</SCRIPT>"

		.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
		.Write "<input type=""hidden"" value=""Edit"" name=""Flag"">"

		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' width='160'><div align=""right""><strong>本模型计费方式：</strong></div></td>"
		.Write "      <td height=""30""> <input type=""radio"" name=""ChargeType"" value=""0"" "
		If ChargeType = 0 Then .Write (" checked")
		.Write ">"
		.Write "       " & KS.Setting(45)
		.Write "          <input type=""radio"" name=""ChargeType"" value=""1"" "
		If ChargeType = 1 Then .Write (" checked")
		.Write ">"
		.Write "        资金(人民币)"		
		.Write "          <input type=""radio"" name=""ChargeType"" value=""2"" "
		If ChargeType = 2 Then .Write (" checked")
		.Write ">"
		.Write "        积分      <br/><span style='color:red'>如文章/图片/下载等设置需要消费才可以查看,将以这里设置的计费标准扣费,一旦设置建议不要修改,此次设置对商城模型无效</span> </td>"
		.Write "    </tr>"	
		.Write "    <tr class='tdbg'>"
		.Write "     <td class='clefttitle' width='160'><div align=""right""><strong>下载本模型附件费用：</strong></div></td>"
		.Write "     <td height=""30""> <input class='txt' type=""text"" size=8 name=""AnnexPoint"" value=""" & AnnexPoint & """> 24小时内下载不重复扣费,不限制请输入0</td>"
		.Write "    </tr>"
			
	
		.Write "    <tr class='tdbg'>"
		.Write "     <td class='clefttitle' width='160'><div align=""right""><strong>最新信息标志：</strong></div></td>"
		.Write "     <td height=""30""> <input class='txt' type=""text"" size=8 name=""LatestNewDay"" value=""" & LatestNewDay & """>天内添加的信息标志为最新信息</td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg'>"
		.Write "     <td class='clefttitle' width='160'><div align=""right""><strong>后台每页显示：</strong></div></td>"
		.Write "     <td height=""30""> <input class='txt' type=""text"" size=8 name=""MaxPerPage"" value=""" & MaxPerPage & """>条信息</td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td width='160' class='clefttitle'><div align=""right""><strong>本模型启用回收站功能：</strong></td>"
		.Write "      <td height=""30""> "
		.Write " <input type=""radio"" name=""DelTF"" value=""1"" "
		If ThumbnailsConfig(5) = "1" Then .Write (" checked")
		.Write ">不启用 "
		.Write "          <input type=""radio"" name=""DelTF"" value=""0"" "
		If ThumbnailsConfig(5) = "0" Then .Write (" checked")
		.Write ">启用<br/><font color=green>启用回收站后，则删除文档将放入回收站里，可以在回收站中还原。</font></td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' align=""right"" width='160'><div><strong>评论设置：</strong></div></td>"
		.Write "      <td height=""30""><table width=""98%"" border=""0"">"
		.Write "     <tr valign=""middle"">"
		.Write "      <td width=""25%"" height=""30"" width='160'>"
		.Write "        <div align=""right"">本模型评论系统设置：</div></td>"
		.Write "      <td height=""30"">"
		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""0"" "
		If VerificCommentTF = 0 Then .Write (" checked")
		.Write ">关闭本模型的所有信息评论<br>"

		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""1"" "
		If VerificCommentTF = 1 Then .Write (" checked")
		.Write ">本模型只允许会员评论，且评论内容需要后台的审核<br>"
		
		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""2"" "
		If VerificCommentTF = 2 Then .Write (" checked")
		.Write ">本模型只允许会员评论，且评论内容不需要后台审核<br>"
		
		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""3"" "
		If VerificCommentTF = 3 Then .Write (" checked")
		.Write ">本模型允许会员，游客评论，且评论内容需要后台审核<br>"
		
		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""4"" "
		If VerificCommentTF = 4 Then .Write (" checked")
		.Write ">本模型允许会员，游客评论，且评论内容不需要后台审核<br/>"
		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""5"" "
		If VerificCommentTF = 5 Then .Write (" checked")
		.Write ">本模型开放所有用户评论，但会员评论不需审核，游客需要审核"

		

		.Write "             </td>"
		.Write "    </tr>"		
		.Write "    <tr valign=""middle"">"
		.Write "      <td width='160' height=""30"">"
		.Write "        <div align=""right"">评论需要验证码：</div></td>"
		.Write "      <td height=""30""> <input type=""radio"" name=""CommentVF"" value=""1"" "
		If CommentVF = 1 Then .Write (" checked")
		.Write ">"
		.Write "        是"
		.Write "          <input type=""radio"" name=""CommentVF"" value=""0"" "
		If CommentVF = 0 Then .Write (" checked")
		.Write ">"
		.Write "          否        </td>"
		.Write "    </tr>"
		.Write "    <tr valign=""middle"">"
		.Write "      <td height=""30""> <div align=""right"">评论字数控制："
		.Write "        </div></td>"
		.Write "      <td width=""63%"" height=""30""> <input class='txt' name=""CommentLen"" type=""text"" value=""" & CommentLen & """ size=""6"">个字,同一个IP<input class='txt' name=""CommentTime"" type=""text"" value=""" & ThumbnailsConfig(4) & """ size=""6"">小时内对同一篇文档只能发表一次。不限制请输入""0""</td>"
		.Write "    </tr>"		
		.Write "    <tr valign=""middle"">"
		.Write "      <td height=""30""><div align=""right"">评论页模板：</div></td>"
		.Write "      <td height=""30""><input class='txt' name=""CommentTemplate"" id=""CommentTemplate"" type=""text"" value=""" & CommentTemplate & """ size=""25"">&nbsp;" & KSCls.Get_KS_T_C("$('#CommentTemplate')[0]") & "</td>"
		.Write "    </tr>"			
		.Write "      </table></td>"
		.Write "    </tr>"	
		If KS.WSetting(0)="1" Then
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle'><div align=""right""><strong>WAP搜索页模板：</strong></div></td>"
		.Write "      <td height=""30""> <input class='txt' name=""WapSearchTemplate"" id=""WapSearchTemplate"" type=""text"" value=""" & WapSearchTemplate & """ size=""25"">&nbsp;" & KSCls.Get_KS_T_C("$('#WapSearchTemplate')[0]") & " </td>"
		.Write "    </tr>"
		End If
		.Write "  </table>"
		.Write "</div>"
		.Write "</form>"
		End With
		End Sub
		
		Sub ChannelSave()
		    Dim ModelEname,ThumbnailsConfig,ChannelTable,I,OpName,ItemName,ChannelID:ChannelID=KS.ChkClng(KS.G("ChannelID"))
            If KS.IsNul(KS.G("ChannelName")) Then
				   Call KS.AlertHistory("请输入模型名称!",-1)
				   Exit Sub
			End If
            If KS.IsNul(KS.G("ModelEName")) And OpName="添加" Then
				   Call KS.AlertHistory("请输入模型英文名称!",-1)
				   Exit Sub
			End If
			ItemName=KS.G("ItemName")
			ThumbnailsConfig=Request.Form("GoldenPoint") & "|" & KS.ChkClng(Request.Form("ThumbsWidth")) & "|" & KS.ChkClng(Request.Form("ThumbsHeight")) & "|" & KS.ChkClng(Request.Form("RefreshTimeTF"))& "|" & KS.ChkClng(Request.Form("CommentTime")) & "|" & KS.ChkClng(Request.Form("DelTF"))& "|" & KS.ChkClng(Request("jsxh")) & "|" & Request("jsxhnum") & "|" & KS.ChkClng(Request("rxhhour")) & "|" & request("rightyy") & "|" & request("wrongyy")& "|" & KS.ChkClng(request("ksyzm")) & "|" & ks.chkclng(request("sjfl"))

		    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_Channel Where ChannelID=" & ChannelID,Conn,1,3
			If  RS.Eof And RS.Bof Then
			    RS.AddNew
				 OpName       = "添加"
				 ChannelTable = "KS_U_" & KS.G("ChannelTable")
				 ModelEname   = Replace(Replace(Replace(KS.G("ModelEname"), "\","/"), " ",""), "'","")
				 If not Conn.Execute("Select ModelEname From KS_Channel Where ModelEname='" & ModelEname & "'").eof Then
				   Call KS.AlertHistory("系统已存在该目录名称，请重输!",-1)
				   Exit Sub
				 End If
				 If not Conn.Execute("Select ChannelTable From KS_Channel Where ChannelTable='" & ChannelTable & "'").eof Then
				   Call KS.AlertHistory("系统已存在该数据表，请重输!",-1)
				   Exit Sub
				 End If
				 Dim sChannelID:sChannelID=Conn.Execute("select Max(ChannelID) From KS_Channel")(0)+1
				 If sChannelID<100 Then sChannelID=sChannelID+100
				RS("ChannelID")    = sChannelID
				RS("BasicType")    = KS.ChkClng(KS.G("BasicType"))
				RS("ChannelTable") = ChannelTable
				RS("ModelEname")   =ModelEname
			Else
			    OpName="修改"
			End If
				RS("ChannelName")= KS.G("ChannelName")
				RS("ItemName")   = ItemName
				RS("ItemUnit")   = KS.G("ItemUnit")
				RS("FsoFolder")  = KS.G("FsoFolder")
				RS("Descript")   = KS.G("Descript")
				if KS.ChkClng(RS("BasicType"))=1 Or KS.ChkClng(RS("BasicType"))=2 then  RS("CollectTF")=1
				RS("ChannelStatus") = KS.G("ChannelStatus")
				If KS.WSetting(0)="1" Then
				RS("WapSwitch")     = KS.ChkClng(KS.G("WapSwitch"))
				End If
				RS("FsoHtmlTF")     = KS.ChkClng(KS.G("FsoHtmlTF"))
				RS("StaticTF")      = KS.ChkClng(KS.G("StaticTF"))
				RS("FsoListNum")    = KS.ChkClng(KS.G("FsoListNum"))
				RS("UpfilesDir")    = KS.G("UpfilesDir")
				RS("UserUpfilesDir") = KS.G("UserUpfilesDir")
				RS("UpFilesTF")     = KS.G("UpFilesTF")
				RS("UserSelectFilesTF")=KS.G("UserSelectFilesTF")
				'If KS.G("UpfilesDir") <> "" Then Call KS.CreateListFolder(KS.Setting(3) & KS.G("UpfilesDir"))
				
				RS("UserUpFilesTF") = KS.G("UserUpFilesTF")
				'If KS.G("UserUpfilesDir") <> "" Then Call KS.CreateListFolder(KS.Setting(3) & KS.G("UserUpfilesDir"))
				
				RS("ThumbnailsConfig")=ThumbnailsConfig
	            RS("UserTF") = KS.ChkClng(KS.G("UserTF"))
				RS("UserEditTF")  = KS.ChkClng(KS.G("UserEditTF"))
				RS("UserClassStyle") = KS.ChkClng(KS.G("UserClassStyle"))
				RS("UpfilesSize") = KS.ChkClng(KS.G("UpfilesSize"))
				RS("AllowUpPhotoType") = KS.G("AllowUpPhotoType")
				RS("AllowUpFlashType") = KS.G("AllowUpFlashType")
				RS("AllowUpMediaType") = KS.G("AllowUpMediaType")
				RS("AllowUpRealType") = KS.G("AllowUpRealType")
				RS("AllowUpOtherType") = KS.G("AllowUpOtherType")
				RS("VerificCommentTF") = KS.G("VerificCommentTF")
				RS("LatestNewDay")     = KS.ChkClng(KS.G("LatestNewDay"))
				RS("CommentVF")    = KS.ChkClng(KS.G("CommentVF"))
				RS("CommentLen")   = KS.ChkClng(KS.G("CommentLen"))
				RS("CommentTemplate") = KS.G("CommentTemplate")
				RS("WapSearchTemplate")= KS.G("WapSearchTemplate")
				RS("InfoVerificTF") = KS.ChkClng(KS.G("InfoVerificTF"))
				RS("MaxPerPage")   = KS.ChkClng(KS.G("MaxPerPage"))
				RS("RefreshFlag")  = KS.ChkClng(KS.G("RefreshFlag"))
				RS("FsoContentRule")=KS.G("FsoContentRule")
				RS("FsoClassListRule")=KS.ChkClng(KS.G("FsoClassListRule"))
				RS("FsoClassPreTag")=KS.G("FsoClassPreTag")
				RS("ModelIco")=KS.G("ModelIco")
				RS("ModelShortName")=KS.G("ModelShortName")

				'会员积分
				RS("UserAddMoney") = KS.ChkClng(KS.G("UserAddMoney"))
				RS("UserAddPoint") = KS.ChkCLng(KS.G("UserAddPoint"))
				RS("UserAddScore") = KS.ChkClng(KS.G("UserAddScore"))
				RS("PubTimeLimit") = KS.ChkClng(KS.G("PubTimeLimit"))
				RS("ChargeType") = KS.ChkClng(KS.G("ChargeType"))
				RS("AnnexPoint") = KS.ChkClng(KS.G("AnnexPoint"))
				RS("DiggByVisitor")= KS.ChkClng(KS.G("DiggByVisitor"))
				RS("DiggByIP")     = KS.ChkClng(KS.G("DiggByIP"))
				RS("DiggRepeat")= KS.ChkClng(KS.G("DiggRepeat"))
				RS("DiggPerTimes")= KS.ChkClng(KS.G("DiggPerTimes"))
				RS.Update
				ChannelID=RS("ChannelID")
				ChannelTable=RS("ChannelTable")
				Dim BasicType:BasicType=RS("BasicType")
				RS.Close
				If BasicType=1 or BasicType=2 Or BasicType=3 Then
				    Dim Doc,Node,CDATASection
					set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
					Doc.async = false
					Doc.setProperty "ServerHTTPRequest", true 
					Doc.load(Server.MapPath(KS.Setting(3)&"Config/modelinputform.xml"))
					Set Node=Doc.documentElement.selectSingleNode("/inputform/model[@name='" & ChannelID & "']")
					 if not node is nothing then  Doc.DocumentElement.RemoveChild(Node)
					 Set Node=Doc.documentElement.appendChild(Doc.createNode(1,"model",""))
					 Node.attributes.setNamedItem(Doc.createNode(2,"name","")).text=channelid
					 Set   CDATASection   = Doc.createCDATASection(Request.Form("Content")) 
					 Node.appendChild   CDATASection 
					Doc.Save(Server.MapPath(KS.Setting(3)&"Config/modelinputform.xml"))
					Application(KS.SiteSN&"_Configmodelinputform")=empty
               End If
				
				
				
				If OpName="添加" Then
				 
				'建立新表
				dim sql
			    Select Case KS.ChkClng(KS.G("BasicType"))
			    Case 1
				sql="CREATE TABLE ["&ChannelTable&"] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_"&ChannelTable&" PRIMARY KEY,"&_
						"TID nvarchar(22),"&_
						"KeyWords nvarchar(255),"&_
						"TitleType nvarchar(30),"&_
						"Title nvarchar(255),"&_
						"FullTitle nvarchar(255),"&_
						"Intro ntext,"&_
						"ShowComment tinyint Default 0,"&_
						"TitleFontColor nvarchar(30),"&_
						"TitleFontType nvarchar(30),"&_
						"ArticleContent ntext,"&_
						"PageTitle ntext,"&_
						"Author nvarchar(30),"&_
						"Origin nvarchar(40),"&_
						"Rank nvarchar(10),"&_
						"Hits int Default 0,"&_
						"HitsByDay int Default 0,"&_
						"HitsByWeek int Default 0,"&_
						"HitsByMonth int Default 0,"&_
						"LastHitsTime datetime,"&_
						"AddDate datetime,"&_
						"ModifyDate datetime,"&_
						"JSID nvarchar(200),"&_
						"TemplateID nvarchar(255),"&_
						"WapTemplateID nvarchar(255)," &_
						"Fname nvarchar(200),"&_
						"RefreshTF tinyint default 0,"&_
						"Inputer nvarchar(50),"&_
						"PhotoUrl nvarchar(150),"&_
						"PicNews tinyint default 0,"&_
						"Changes tinyint default 0,"&_
						"Recommend tinyint Default 0,"&_
						"Rolls tinyint Default 0,"&_
						"Strip tinyint Default 0,"&_
						"Popular tinyint Default 0,"&_
						"Verific tinyint Default 0,"&_
						"Slide tinyint Default 0,"&_
						"Comment tinyint Default 0,"&_
						"IsTop tinyint Default 0,"&_
						"IsVideo tinyint Default 0,"&_
						"DelTF tinyint Default 0,"&_
						"PostID int Default 0,"&_
						"PostTable varchar(100),"&_
						"IsSign tinyint Default 0,"&_
						"SignUser nvarchar(255),"&_
						"SignDateLimit tinyint Default 0,"&_
						"SignDateEnd datetime,"&_
						"Province nvarchar(100),"&_
						"City nvarchar(100),"&_
						"MapMarker nvarchar(255),"&_
						"InfoPurview tinyint Default 0,"&_
						"ArrGroupID nvarchar(100),"&_
						"ReadPoint int Default 0,"&_
						"ChargeType tinyint Default 0,"&_
						"PitchTime int Default 24,"&_
						"ReadTimes int Default 10,"&_
						"DividePercent int Default 0,"&_
						"SEOTitle varchar(255),"&_
						"SEOKeyWord ntext,"&_
						"SEODescript ntext"&_
						")"
				Conn.Execute(sql)
				KS.ConnItem.Execute(sql)
				'添加索引
				Call AddIndex(ChannelTable, "[TID]", "[TID]")
				Call AddIndex(ChannelTable, "[Verific]", "[verific]")
				Call AddIndex(ChannelTable, "[deltf]", "[deltf]")
				Call AddIndex(ChannelTable, "[adddate]", "[adddate]")
				Call AddIndex(ChannelTable, "[hits]", "[hits]")
				'Call AddIndex(ChannelTable, "[specialid]", "[specialid]")
			 Case 2
				sql="CREATE TABLE ["&ChannelTable&"] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_"&ChannelTable&" PRIMARY KEY,"&_
						"Tid nvarchar(22),"&_
						"KeyWords nvarchar(255),"&_
						"Title nvarchar(255),"&_
						"showstyle tinyint default 0,"&_
						"pagenum int default 10,"&_
						"PhotoUrl nvarchar(255),"&_
						"PicUrls ntext,"&_
						"PictureContent ntext,"&_
						"Author nvarchar(30),"&_
						"Origin nvarchar(40),"&_
						"Rank nvarchar(10),"&_
						"LastHitsTime smalldatetime," &_
						"Hits int Default 0,"&_
						"HitsByDay int Default 0,"&_
						"HitsByWeek int Default 0,"&_
						"HitsByMonth int Default 0,"&_
						"AddDate smalldatetime,"&_
						"ModifyDate datetime,"&_
						"JSID nvarchar(200),"&_
						"TemplateID nvarchar(255),"&_
						"WapTemplateID nvarchar(255)," &_
						"Fname nvarchar(200),"&_
						"RefreshTF tinyint default 0,"&_
						"Inputer nvarchar(50),"&_
						"Recommend tinyint Default 0,"&_
						"Rolls tinyint Default 0,"&_
						"Strip tinyint Default 0,"&_
						"Popular tinyint Default 0,"&_
						"Verific tinyint Default 0,"&_
						"Slide tinyint Default 0,"&_
						"Comment tinyint Default 0,"&_
						"IsTop tinyint Default 0,"&_
						"Score int Default 0,"&_
						"MapMarker nvarchar(255),"&_
						"DelTF tinyint Default 0,"&_
						"InfoPurview tinyint Default 0,"&_
						"ArrGroupID nvarchar(100),"&_
						"ReadPoint int Default 0,"&_
						"ChargeType tinyint Default 0,"&_
						"PitchTime int Default 24,"&_
						"ReadTimes int Default 10,"&_
						"DividePercent int Default 0,"&_
						"SEOTitle varchar(255),"&_
						"SEOKeyWord ntext,"&_
						"SEODescript ntext"&_
						")"
				Conn.Execute(sql)
				KS.ConnItem.Execute(sql)
				'添加索引
				Call AddIndex(ChannelTable, "[TID]", "[TID]")
				Call AddIndex(ChannelTable, "[Verific]", "[verific]")
				Call AddIndex(ChannelTable, "[deltf]", "[deltf]")
				Call AddIndex(ChannelTable, "[adddate]", "[adddate]")
				Call AddIndex(ChannelTable, "[hits]", "[hits]")
				'Call AddIndex(ChannelTable, "[specialid]", "[specialid]")
				
			 Case 3
				sql="CREATE TABLE ["&ChannelTable&"] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_"&ChannelTable&" PRIMARY KEY,"&_
						"Tid nvarchar(22),"&_
						"KeyWords nvarchar(255),"&_
						"Title nvarchar(255),"&_
						"DownVersion nvarchar(50),"&_
						"DownLB nvarchar(100),"&_
						"DownYY nvarchar(100),"&_
						"DownSQ nvarchar(100),"&_
						"DownPT nvarchar(100),"&_
						"DownSize nvarchar(100),"&_
						"YSDZ nvarchar(100),"&_
						"ZCDZ nvarchar(100),"&_
						"JYMM nvarchar(100),"&_
						"PhotoUrl nvarchar(200),"&_
						"BigPhoto nvarchar(200),"&_
						"DownUrls ntext,"&_
						"DownContent ntext,"&_
						"Author nvarchar(50)," & _
						"Origin nvarchar(40),"&_
						"Rank nvarchar(10),"&_
						"LastHitsTime smalldatetime," &_
						"Hits int Default 0,"&_
						"HitsByDay int Default 0,"&_
						"HitsByWeek int Default 0,"&_
						"HitsByMonth int Default 0,"&_
						"AddDate datetime,"&_
						"ModifyDate datetime,"&_
						"JSID nvarchar(200),"&_
						"TemplateID nvarchar(255),"&_
						"WapTemplateID nvarchar(255)," &_
						"Fname nvarchar(200),"&_
						"RefreshTF tinyint default 0,"&_
						"Inputer nvarchar(50),"&_
						"Recommend tinyint Default 0,"&_
						"Rolls tinyint Default 0,"&_
						"Strip tinyint Default 0,"&_
						"Popular tinyint Default 0,"&_
						"Verific tinyint Default 0,"&_
						"Slide tinyint Default 0,"&_
						"Comment tinyint Default 0,"&_
						"IsTop tinyint Default 0,"&_
						"DelTF tinyint Default 0,"&_
						"InfoPurview tinyint Default 0,"&_
						"ArrGroupID nvarchar(100),"&_
						"ReadPoint int Default 0,"&_
						"ChargeType tinyint Default 0,"&_
						"PitchTime int Default 24,"&_
						"ReadTimes int Default 10,"&_
						"DividePercent int Default 0,"&_
						"SEOTitle varchar(255),"&_
						"SEOKeyWord ntext,"&_
						"SEODescript ntext"&_
						")"
				Conn.Execute(sql)
				'添加索引
				Call AddIndex(ChannelTable, "[TID]", "[TID]")
				Call AddIndex(ChannelTable, "[Verific]", "[verific]")
				Call AddIndex(ChannelTable, "[deltf]", "[deltf]")
				Call AddIndex(ChannelTable, "[adddate]", "[adddate]")
				Call AddIndex(ChannelTable, "[hits]", "[hits]")
			 Case 4
				sql="CREATE TABLE ["&ChannelTable&"] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_"&ChannelTable&" PRIMARY KEY,"&_
						"Tid nvarchar(22),"&_
						"KeyWords nvarchar(255),"&_
						"Title nvarchar(255),"&_
						"PhotoUrl nvarchar(255),"&_
						"FlashUrl varchar(255),"&_
						"FlashContent ntext,"&_
						"Author nvarchar(30),"&_
						"Origin nvarchar(40),"&_
						"Rank nvarchar(10),"&_
						"LastHitsTime smalldatetime," &_
						"Hits int Default 0,"&_
						"HitsByDay int Default 0,"&_
						"HitsByWeek int Default 0,"&_
						"HitsByMonth int Default 0,"&_
						"AddDate datetime,"&_
						"ModifyDate datetime,"&_
						"JSID nvarchar(200),"&_
						"TemplateID nvarchar(255),"&_
						"WapTemplateID nvarchar(255)," &_
						"Fname nvarchar(200),"&_
						"RefreshTF tinyint default 0,"&_
						"Inputer nvarchar(50),"&_
						"Recommend tinyint Default 0,"&_
						"Rolls tinyint Default 0,"&_
						"Strip tinyint Default 0,"&_
						"Popular tinyint Default 0,"&_
						"Verific tinyint Default 0,"&_
						"Slide tinyint Default 0,"&_
						"Comment tinyint Default 0,"&_
						"IsTop tinyint Default 0,"&_
						"Score int Default 0,"&_
						"MapMarker nvarchar(255),"&_
						"DelTF tinyint Default 0,"&_
						"InfoPurview tinyint Default 0,"&_
						"ArrGroupID nvarchar(100),"&_
						"ReadPoint int Default 0,"&_
						"ChargeType tinyint Default 0,"&_
						"PitchTime int Default 24,"&_
						"ReadTimes int Default 10,"&_
						"DividePercent int Default 0,"&_
						"SEOTitle varchar(255),"&_
						"SEOKeyWord ntext,"&_
						"SEODescript ntext"&_
						")"
				Conn.Execute(sql)
				KS.ConnItem.Execute(sql)
				'添加索引
				Call AddIndex(ChannelTable, "[TID]", "[TID]")
				Call AddIndex(ChannelTable, "[Verific]", "[verific]")
				Call AddIndex(ChannelTable, "[deltf]", "[deltf]")
				Call AddIndex(ChannelTable, "[adddate]", "[adddate]")
				Call AddIndex(ChannelTable, "[hits]", "[hits]")
				'Call AddIndex(ChannelTable, "[specialid]", "[specialid]")
				
			 End Select
				
				
				
				If KS.ChkClng(KS.G("BasicType"))=3 Then
				 Call KS.CreateListFolder(KS.Setting(3) & KS.G("UpfilesDir")&"DownPhoto/")
				 Call KS.CreateListFolder(KS.Setting(3) & KS.G("UpfilesDir")&"DownUrl/")
				End IF
				  
                 
				'  If Err<>0 Then
				'	Conn.RollBackTrans
				'	Call KS.AlertHistory("出错！出错描述：" & replace(err.description,"'","\'"),-1):response.end
				'  Else
				'	Conn.CommitTrans
				  'End If
				  Call KS.DelCahe(KS.SiteSN & "_selectallowclass")
				  Call KS.DelCahe(KS.SiteSN & "_selectclass")
				  Call KS.DelCahe(KS.SiteSN & "_classpath")
				  Call KS.DelCahe(KS.SiteSN & "_classnamepath")				     
				End If


				Call KS.DelCahe(KS.SiteSN & "_ChannelConfig")
				
				Call KSCLS.CreateModelField(ChannelID)     '初始化模型字段

				
				Response.Write ("<script>alert('KesionCMS系统提醒您：\n\n1、模型配置信息" & OpName & "成功；\n\n2、为了使配置生效，请及时更新缓存；');top.location.href='index.asp?from=KS.Model.asp';</script>")
			
		End Sub
	
		Sub DelColumn(TableName,ColumnName)
		On Error Resume Next
		Conn.Execute("Alter Table "&TableName&" Drop "&ColumnName&"")
		End Sub
		
		Sub DelTable(TableName,C)
			On Error Resume Next
			C.Execute("Drop Table "&TableName&"")
		End Sub
		
		Sub AddIndex(ByVal TableName, ByVal IndexName, ByVal ValueText)
			On Error Resume Next
			Conn.Execute("CREATE INDEX " & IndexName & " ON " & TableName & "(" & ValueText & ")")
		End Sub
		
		
		Sub ChannelDel()
		  Dim ChannelID:ChannelID=KS.ChkClng(KS.G("ChannelID"))
		  Call DelTable(KS.C_S(ChannelID,2),Conn)
		  
		  '删除采集数据库里的相关字段和表
		  If KS.C_S(ChannelID,6)="1" Or KS.C_S(ChannelID,6)="2" or KS.C_S(ChannelID,6)="5" Then  Call DelTable(KS.C_S(ChannelID,2),KS.ConnItem)
		  KS.ConnItem.Execute("Delete From KS_FieldItem Where ChannelID=" & ChannelID)
		  KS.ConnItem.Execute("Delete From KS_FieldRules Where ChannelID=" & ChannelID)
		  '=================================
		  
		  Conn.Execute("Delete From KS_Comment Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_DownParam Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_DownSer Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_Origin Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_Channel Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_Class Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_Field Where ChannelID=" & ChannelID)
		  
		  Call KS.DeleteFile(KS.Setting(3) &"config/fielditem/field_"&ChannelID&".xml")
		  
		  '删除录入表单的模板
			Dim Doc,Node,CDATASection
			set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(Server.MapPath(KS.Setting(3)&"Config/modelinputform.xml"))
			Set Node=Doc.documentElement.selectSingleNode("/inputform/model[@name='" & ChannelID & "']")
			if not node is nothing then  Doc.DocumentElement.RemoveChild(Node)
			Doc.Save(Server.MapPath(KS.Setting(3)&"Config/modelinputform.xml"))
			Application(KS.SiteSN&"_Configmodelinputform")=empty

		  
		  		 Call KS.DelCahe(KS.SiteSN & "_selectallowclass")
				 Call KS.DelCahe(KS.SiteSN & "_selectclass")
				 Call KS.DelCahe(KS.SiteSN & "_classpath")
				 Call KS.DelCahe(KS.SiteSN & "_classnamepath")
				 Call KS.DelCahe(KS.SiteSN & "_ChannelConfig")
		  Response.Write "<script>alert('模型删除成功!');top.location.href='index.asp?from=KS.Model.asp';</script>" 
		End Sub
		
End Class
%> 

