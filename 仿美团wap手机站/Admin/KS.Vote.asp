<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Vote
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Vote
        Private KS,KSCls
		Private I, totalPut, CurrentPage, SqlStr, RSObj
        Private MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
			If Not KS.ReturnPowerResult(0, "KSMS20003") Then
			  Call KS.ReturnErr(1, "")
			  exit sub
			End If
			Select Case KS.G("Action")
			 Case "Add","Edit" Call VoteAdd()
			 Case "Del" Call VoteDel()
			 Case "Set" Call VoteSet()
			 Case Else Call MainList()
			End Select
			
	  End Sub
	  
	  Sub MainList()
			If Request("page") <> "" Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
			With Response
			.Write "<html>"
			.Write "<head>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.Write "<title>站点调查</title>"
			.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script language=""JavaScript"">" & vbCrLf
			.Write "var Page='" & CurrentPage & "';" & vbCrLf
			.Write "</script>" & vbCrLf
			.Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
			.Write "<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>"
			%>
			<script language="javascript">
			$(document).ready(function(){
				
		      $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
			  $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
		     })
			
			
			function VoteAdd()
			{
				location.href='KS.Vote.asp?Action=Add';
				window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=调查主题管理中心 >> <font color=red>添加新调查主题</font>&ButtonSymbol=VoteAddSave';
			}
			function EditVote(id)
			{
			   if (id=='') id=get_Ids(document.myform);
			   if (id==''){
				 alert('请选择要编辑的调查主题!');
				}else if(id.indexOf(',')==-1){
				location="KS.Vote.asp?Action=Edit&VoteID="+id;
				window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=调查主题管理中心 >> <font color=red>编辑调查主题</font>&ButtonSymbol=VoteEdit';
				}else{
				alert('一次只能编辑一个调查主题!');
				}
			}
			function DelVote(id)
			{
			 if (id=='') id=get_Ids(document.myform);
			 if (id==''){
			   alert('请先选择要删除的调查主题!')
			 }else if  (confirm('真的要删除选中的调查主题吗?')){
			 location="KS.Vote.asp?Action=Del&Page="+Page+"&id="+id;
			 }
			}
			</script>
			<%
			.Write "</head>"
			.Write "<body topmargin=""0"" leftmargin=""0"">"
		    .Write "<ul id='menu_top'>"
			.Write "<li class='parent' onclick=""VoteAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加调查</span></li>"
			.Write "<li class='parent' onclick=""EditVote('');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>编辑调查</span></li>"
			.Write "<li class='parent' onclick=""DelVote('');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>删除调查</span></li>"
			.Write "</ul>"
			.Write "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""1"">"
			.Write "<form name='myform' action='KS.Vote.asp' method='post'>"
			.Write "<input type='hidden' name='action' value='Del'>"
			.Write "  <tr>"
			.Write "          <td width=""35"" height=""25"" class=""sort"">选择</td>"
			.Write "          <td height=""25"" class=""sort""align=""center"">调查主题</td>"
			.Write "          <td width=""100"" class=""sort"" align=""center"">论坛发起</td>"
			.Write "          <td width=""100"" class=""sort"" align=""center"">发起人</td>"
			.Write "          <td width=""120"" align=""center"" class=""sort"">时间</td>"
			.Write "          <td width=""100"" class=""sort"" align=""center"">是否最新</td>"
			.Write "          <td width=""120"" class=""sort"" align=""center"">管理操作</td>"
			.Write "  </tr>"
			 
			 Set RSObj = Server.CreateObject("ADODB.RecordSet")
					   SqlStr = "SELECT * FROM KS_Vote order by NewestTF desc,AddDate desc"
					   RSObj.Open SqlStr, Conn, 1, 1
					 If RSObj.EOF And RSObj.BOF Then
					   .Write "<tr><td height='30' class='splittd' align='center' colspan='6'>还没有添加调查主题!</td></tr>"
					 Else
						        totalPut = RSObj.RecordCount
								If CurrentPage < 1 Then	CurrentPage = 1
								If CurrentPage > 1 and  (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
								Else
										CurrentPage = 1
								End If
								Call showContent
				End If
				
			.Write "    </td>"
			.Write "  </tr>"
			.Write "</table>"
			.Write "</body>"
			.Write "</html>"
			End With
			End Sub
			Sub showContent()
			  Dim ID
			  With Response
					Do While Not RSObj.EOF
					   ID=RSObj("id")
					   .Write ("<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" &ID & "' onclick=""chk_iddiv('" & ID & "')"">")
				       .Write ("<td class='splittd' align=center><input type='hidden' value='" & ID & "' name='VoteID'><input name='id'  onclick=""chk_iddiv('" & ID & "')"" type='checkbox' id='c"& ID & "' value='" & ID & "'></td>")
					  .Write "  <td class='splittd'  height='20'> &nbsp;&nbsp; <span VoteID='" & ID & "' ondblclick=""EditVote(this.VoteID)""><img src='Images/37.gif' align='absmiddle'> "
					  .Write    KS.GotTopic(RSObj("Title"), 50) & "</span> "
					  .Write "  </td>"
					  .Write "  <td class='splittd'  align='center'>" 
					   If KS.ChkClng(RSOBj("TopicID"))=0 Then
					     .Write "否"
					   Else
					     .Write "<a href='" & KS.GetClubShowUrl(RSObj("TopicID")) & "' style='color:green' target='_blank'>是</a>"
					   End If
					  .Write " </td>"
					  .Write "  <td class='splittd'  align='center'>" & RSObj("UserName") & " </td>"
					  .Write "  <td class='splittd'  align='center'><FONT Color=red>" & RSObj("AddDate") & "</font> </td>"
					  If RSObj("NewestTF") = 1 Then
					   .Write "  <td class='splittd' align='center'><font color=red>是</font></td>"
					  Else
					   .Write "  <td class='splittd' align='center'>否</td>"
					  End If
					   .Write "  <td class='splittd' align='center'><a href=""javascript:EditVote('"&Id&"');"">修改</a> | <a href=""javascript:DelVote('"&Id&"');"" >删除</a> | <a href=""../?do=vote&id=" & id & """ target=""_blank"">查看</a></td>"
					  .Write "</tr>"

					I = I + 1
					  If I >= MaxPerPage Then Exit Do
						   RSObj.MoveNext
					Loop
					  RSObj.Close
					  Conn.Close
					 .Write "</table><table width='100%'><tr><td><div style='margin:5px'><b>选择：</b><a href='javascript:void(0)' onclick='Select(0)'>全选</a> -  <a href='javascript:void(0)' onclick='Select(1)'>反选</a> - <a href='javascript:void(0)' onclick='Select(2)'>不选</a> <input type='submit' class='button' value='删 除' onclick=""return(confirm('确定删除选中的调查主题吗?'))""></form></td><td height='26' colspan='2' align='right'>"
					 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			  End With
			End Sub
			
			Sub VoteDel()
			 Dim ID,IDArr,I
			 ID=KS.S("ID")
			 If KS.IsNul(ID) Then Call KS.AlertHintScript("请选择要删除的主题!")
			 IDArr=Split(KS.FilterIds(ID),",")
			 For I=0 To Ubound(IDArr)
			 KS.DeleteFile(KS.Setting(3)&"config/voteitem/vote_" & IDArr(i) &".xml")
			 Conn.Execute("delete from KS_Vote where ID="&Clng(IDArr(i)))
			 Conn.Execute("delete from KS_PhotoVote where channelid=-1 and InfoID='"&Clng(IDArr(i))&"'")
			 Next
			 Response.redirect "KS.Vote.asp?Page="&KS.G("Page")
			End Sub
			
			Sub VoteSet()
				conn.execute "Update KS_Vote set NewestTF=0 where NewestTF=1"
				conn.execute "Update KS_Vote set NewestTF=1 Where ID=" & Clng(KS.G("VoteID"))
				Response.Write "<script language='JavaScript' type='text/JavaScript'>alert('设置成功！');location.href='KS.Vote.asp?Page=" & KS.G("Page") & "';</script>"

			End Sub
			
			Sub VoteAdd()
				With Response
				.Write "<html>"
				.Write "<head>"
				.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
				.Write "<title>调查管理-添加主题</title>"
				.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
				.Write "<script src=""../KS_Inc/jQuery.js"" language=""JavaScript""></script>"
				.Write "</head>"
				.Write "<body topmargin=""0"" leftmargin=""0"">"
	
				.Write "  <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
				.Write "        <tr>"
				.Write "          <td width=""44%"" height=""25"" class=""sort"">"
				.Write "          <div align=""center""><strong>添 加 调 查 主 题</strong></div></td>"
				.Write "        </tr>"
				.Write "      </table>"
	           
			   dim timelimit,Title,VoteTime,NewestTF,rs,sql,voteid,ItemArr,VoteNumArr,i,XMLStr
			   Dim VoteType,timebegin,timeend,nmtp,AllowGroupID,ipnum,ipnumS,templateid,Status,editnum
			   timelimit=0:nmtp=0:Status=1:editnum=0:timebegin=now:timeend=dateadd("m",1,now)
			   templateid="{@TemplateDir}/投票页.html"
			   
				
				Title=trim(request.form("Title"))
				VoteTime=trim(request.form("VoteTime"))
				if VoteTime="" then VoteTime=now()
				NewestTF=trim(request("NewestTF"))
				
				ItemArr=Split(request("item"),",")
				VoteNumArr=Split(Request("VoteNum"),",")
				
				if Title<>"" then
					sql="select top 1 * from KS_Vote Where ID=" & ks.chkclng(request("voteid"))
					Set rs= Server.CreateObject("ADODB.Recordset")
					rs.open sql,conn,1,3
					if rs.eof then
					rs.addnew
					 rs("TopicID")=0
					 rs("VoteNums")=0
					end if
					rs("Title")=Title
					rs("timelimit")=KS.ChkClng(KS.G("TimeLimit"))
					If IsDate(Request("TimeBegin")) Then
					rs("TimeBegin")=Request("TimeBegin")
					Else
					rs("TimeBegin")=Now
					End If
					If IsDate(Request("TimeEnd")) Then
					 rs("TimeEnd")=Request("TimeEnd")
					Else
					 rs("TimeEnd")=Now
					End If
					rs("nmtp")=KS.ChkClng(Request("nmtp"))
					rs("groupids")=request.form("allowgroupid")
					rs("ipnum")=KS.ChkClng(Request.Form("ipnum"))
					rs("ipnums")=KS.ChkClng(Request.Form("ipnums"))
					rs("templateid")=request.form("templateid")
					rs("status")=KS.ChkClng(Request.Form("status"))
					rs("AddDate")=VoteTime
					rs("VoteType")=request("VoteType")
					rs("UserName")=KS.C("AdminName")
					if NewestTF="" then NewestTF=0
					rs("NewestTF")=NewestTF
					rs.update
					rs.movelast
					voteid=rs("id")
					rs.close
					if NewestTF=1 then conn.execute "Update KS_Vote set NewestTF=0 where NewestTF=1 and id<>" & voteid

					
					XMLStr="<?xml version=""1.0"" encoding=""utf-8"" ?>" &vbcrlf
					XMLStr=XMLStr&" <vote>" &vbcrlf
					for i=0 to ubound(ItemArr)
					  if trim(Itemarr(i))<>"" Then
					    XMLStr=XMLStr & "  <voteitem id=""" & i+1 &""">"&vbcrlf
						XMLStr=XMLStr & "    <name><![CDATA[" & Itemarr(i) &"]]></name>" &vbcrlf
						XMLStr=XMLStr & "    <num>" & KS.ChkClng(VoteNumArr(i)) &"</num>" &vbcrlf
					    XMLStr=XMLStr & "  </voteitem>"&vbcrlf
						
					  End If
					Next
					XMLStr=XMLStr &" </vote>" &vbcrlf
					Call KS.WriteTOFile(KS.Setting(3) & "config/voteitem/vote_" & voteid & ".xml",xmlstr)
			        Application(KS.SiteSN&"_Configvote_"&voteid)=null
					set rs=nothing
					call CloseConn()
					if ks.chkclng(request("voteid"))=0 then
					 ks.die "<script>if (confirm('恭喜，投票项目添加成功，继续添加吗？')){location.href='KS.Vote.asp?action=Add';}else{location.href='KS.Vote.asp';}</script>"
					else
					 ks.die "<script>alert('恭喜，投票项目修改成功!');location.href='KS.Vote.asp';</script>"
					end if
				end if
				 End With
				 
			if KS.ChkClng(request("voteid"))<>0 Then
			  Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "select top 1 * from ks_vote where id=" & KS.ChkClng(request("voteid")),conn,1,1
			  If Not RS.Eof Then
			    title    = RS("Title")
				VoteType = RS("VoteType")
				NewestTF = RS("NewestTF")
				timelimit= RS("timelimit")
				timebegin= RS("timebegin")
				timeEnd  = RS("timeEnd")
				nmtp     = RS("nmtp")
				AllowGroupID = RS("GroupIDs")
				ipnum    = RS("ipnum")
				ipnumS   = RS("ipnumS")
				templateid = RS("templateid")
				status=rs("status")
			  End If
			End If	 
				 
				%>
	

				<form method="POST" name="voteform" action="KS.Vote.asp?Action=Add">
				<input type="hidden" name="voteid" value="<%=request("voteid")%>">
						<table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" class="ctable">
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>主题名称：</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							<input name="Title" type="text" size="40" value="<%=title%>" maxlength="50">
							如：你对本站的哪些栏目较感兴趣!</td>
						  </tr>
                          <tr class="tdbg"> 
							<td height="25" align="right" class="clefttitle"><strong>调查类型：</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
										<select name="VoteType" id="VoteType">
											<option value="Single"<%If VoteType="Single" Then Response.Write " selected"%>>单选</option>
											<option value="Multi"<%If VoteType="Multi" Then Response.Write " selected"%>>多选</option>
									</select>
										<input name="NewestTF" type="checkbox" id="NewestTF" value="1"<%If NewestTF="1" Then Response.Write " checked"%> />
	设为最新调查</td>
						  </tr>						  
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>投票项目：</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							
							 <table border="0" cellpadding="0" cellspacing="0" style="margin-left:5px;" width="80%">
     
                 <tr>
                  <td colspan="3" height="30px">
							投票扩展数量: 
						  <input name="vote_num" type="text" id="votenum" value="1" size="5" style="text-align:center"> 
						  <input type="button" name="Submit52" value="增加选项" class="button" onclick="javascript:doadd(jQuery('#votenum').val());"> 
							  
							  </td>
							 </tr>
							 <tr bgcolor='#DBEAF5'>
							 <td width='9%' height='20'> <div align='center'>编号</div></td>
							 <td width='65%'> <div align='center'>项目名称</div></td>
							 <td style='width: 100px'> <div align='center'>投票数</div></td>
							 </tr>
							 <tr>
							  <td colspan="3" id="addvote">
							  <%if request("voteid")<>"" then
							    Dim VoteXML,TaskNode,Node,N,TaskUrl,Taskid,Action
								set VoteXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
								VoteXML.async = false
								VoteXML.setProperty "ServerHTTPRequest", true 
								VoteXML.load(Server.MapPath(KS.Setting(3)&"Config/voteitem/vote_" & request("voteid")& ".xml"))
								Dim TempStr
								 if VoteXML.readystate=4 and VoteXML.parseError.errorCode=0 Then
									editnum=VoteXml.DocumentElement.SelectNodes("voteitem").length
									 For Each Node In VoteXml.DocumentElement.SelectNodes("voteitem")
									  tempstr=tempstr & "<tr><td width=9% height=20> <div align=center><input type=hidden name=id value=" & Node.getAttribute("id") & ">" & Node.getAttribute("id") & "</div></td><td width='65%'> <div align=center><input type=text name=item size=40 value='" & trim(Node.childNodes(0).text) & "'></div></td><td width='26%'> <div align=center><input type=text name=votenum style=text-align:center value='" & Node.childNodes(1).text & "' size=6></div></td></tr>"
									 Next
								 end if
							    end if
								response.write "<table width=100% border=0 cellspacing=1 cellpadding=3>"
								response.write tempstr
								response.write "</table>"
							  %>
							  
							  
							  
							  </td>
							 </tr>
							</table>
							<input name="editnum" type="hidden" id="editnum" value="<%=editnum%>"> 

							<script type="text/javascript">
    function doadd(num)
    {var i;
    var str="";
    var oldi=0;
    var j=0;
    oldi=parseInt(jQuery('#editnum').val());
    for(i=1;i<=num;i++)
    {
    j=i+oldi;
    str=str+"<tr><td width=9% height=20> <div align=center><input type=hidden name=id value=0>"+j+"</div></td><td width=65%> <div align=center><input type=text name=item size=40></div></td><td width=26%> <div align=center><input type=text name=votenum style='text-align:center' value=0 size=6></div></td></tr>";
    }
     jQuery("#addvote").html(jQuery("#addvote").html()+"<table width=100% border=0 cellspacing=1 cellpadding=3>"+str+"</table>");
        jQuery('#editnum').val(j);
    }
	<%If request("voteid")="" Then%>
	doadd(8);
	<%end if%>
    </script>
							</td>
						  </tr>
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>启用时间限制：</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							<label><input type='radio' name='timelimit' onclick="$('#time').hide();" value='0'<%IF timelimit="0" Then Response.Write " checked"%>>不启用</albe>
							<label><input type='radio' name='timelimit' onclick="$('#time').show();" value='1'<%IF timelimit="1" Then Response.Write " checked"%>>启用</label>
							</td>
						  </tr>
						  <tbody id='time'<%if timelimit="0" then response.write " style='display:none'"%>>
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>时间限制：</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							有效期 从<input type='text' name='timebegin' value='<%=timebegin%>'>到
							<input type='text' name='timeend' value='<%=timeend%>'> 时间格式为:YYYY-MM-DD HH:mm:ss
							</td>
						  </tr>
						  </tbody>
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>匿名投票：</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							<label><input type='radio' name='nmtp' value='0'<%If nmtp="0" Then Response.Write " checked"%>>允许匿名投票</label>
							<label><input type='radio' name='nmtp' value='1'<%If nmtp="1" Then Response.Write " checked"%>>只允许会员投票</label>
							</td>
						  </tr>
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>限定用户组：</strong>
							<br/>不限制请不要选
							</td>
							<td colspan="3" bgcolor="#EEF8FE">
							<%=KS.GetUserGroup_CheckBox("AllowGroupID",AllowGroupID,5)%>
							</td>
						  </tr>
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>同一IP：</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							一天内最多可以投<input type="text" name='ipnum' value='<%=IPNUM%>' size='3' style='text-align:center'>次 ,总共可以投<input type="text" name='ipnums' value='<%=IPNums%>' size='3' style='text-align:center'>次。tips:不限制请输入0
							</td>
						  </tr>
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>投票页模板：</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							 <input type="text" name="templateid" value="<%=templateid%>" size="40" id="templateid">
							 <%=KSCls.Get_KS_T_C("document.getElementById('TemplateID')")	%>
							</td>
						  </tr>
						  <tr class="tdbg"> 
							<td width="101" height="30" align="right" class="clefttitle"><strong>状态：</strong></td>
							<td colspan="3" bgcolor="#EEF8FE">
							<label><input type='radio' name='status' value='0'<%if status="0" then response.write " checked"%>>关闭</label>
							<label><input type='radio' name='status' value='1'<%if status="1" then response.write " checked"%>>正常</label>
							</td>
						  </tr>
									
									
							  </table>
							</form>
						</td>
					</tr>
	</table>
	<br/>
	<script>
	 function CheckForm()
	 { var form=document.voteform;
	  if (form.Title.value=='')
	   {
		 alert('请输入调查主题!');
		  form.Title.focus();
		 return false;
	   }
	   $("input[name='item']").each(function(){
	     $(this).val($(this).val().replace(/,/g,'，'));
	   });
	  document.voteform.submit();
	 }
	</script>
<%
			End Sub
			
End Class
%>
 
