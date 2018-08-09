<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_DiggList
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_DiggList
        Private KS,Action,Page,KSCls
		Private I, totalPut, CurrentPage,MaxPerPage, SqlStr,ChannelID,ItemName,ItemName1,RS
		Private OriginName, ID, Sex, Birthday, Telphone, UnitName, UnitAddress, Zip, Email, QQ, HomePage, Note, OriginType
		
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		  Set KSCls= New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
             With KS
		 	    .echo "<html>"
				.echo "<head>"
				.echo "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
				.echo "<title>Digg管理</title>"
				.echo "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		        .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
		        .echo "<script language=""JavaScript"" src=""../KS_Inc/Jquery.js""></script>" & vbCrLf
               Action=KS.G("Action")
			   ChannelID=KS.ChkClng(KS.G("ChannelID"))
			    If ChannelID=0 Then ChannelID=1
				If Not KS.ReturnPowerResult(0, "KSMS20009") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
				End iF

			 Page=KS.G("Page")
			 ItemName=KS.C_S(ChannelID,3)
			 Select Case Action
			  Case "Del" ItemDelete
			  Case "DiggDel" DiggDel
			  Case "DelAllRecord" DelAllRecord
			  Case "ShowCode" ShowCode
			  Case Else MainList()
			 End Select
			.echo "</body>"
			.echo "</html>"
			End With
		End Sub
		
		Sub MainList()
			If Not IsEmpty(Request("page")) Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
		With KS
%>	   		
     <SCRIPT language=javascript>
		function DelDiggList()
		{
			var ids=get_Ids(document.myform);
			if (ids!='')
			 { 
				if (confirm('真的要删除选中的记录吗?'))
				{
				$("#myform").action="KS.DiggList.asp?ChannelID=<%=ChannelID%>&Action=Del&show=<%=KS.G("show")%>&DiggID="+ids;
				$("#myform").submit();
				}
			}
			else 
			{
			 alert('请选择要删除的评论!');
			}
		}
		function DelDigg(){if (confirm('真的要删除选中的记录吗?')){$("#myform").submit();}}
		function show(t,m,d){new parent.KesionPopup().PopupCenterIframe('查看详情[<font color=red>'+t+'</font>]记录','KS.DiggList.asp?action=list&infoid='+d+'&ChannelID='+m,750,440,'auto')}
		function ShowCode(){new parent.KesionPopup().PopupCenterIframe('查看Digg调用代码','KS.DiggList.asp?action=ShowCode',750,440,'no')}

		</SCRIPT>

	   <%
	
		.echo "</head>"
		
		.echo "<body topmargin='0' leftmargin='0'>"
		If KS.S("Action")="list" Then Call DiggDetail() : Exit Sub
		.echo "<ul id='mt'> <div id='mtl'>快速查看：</div><li>"
		.echo "<a href='javascript:ShowCode()'><img src='images/ico/s.gif' align='absmiddle' border='0'>调用代码</a> | "
		.echo "<a href='?ChannelID=" & ChannelID &"&show=1'>顶数最多</a> | <a href='?ChannelID=" & ChannelID &"&show=2'>顶数最少</a> | <a href='?ChannelID=" & ChannelID &"&show=3'>踩数最多</a> | <a href='?ChannelID=" & ChannelID &"&show=4'>踩数最少</a> | <a href='?ChannelID=" & ChannelID &"&show=5'>推荐时间升</a> | <a href='?ChannelID=" & ChannelID &"&show=6'>推荐时间降</a> "
		.echo "</ul>"
		.echo "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		.echo(" <form name=""myform"" id=""myform"" method=""Post"" action=""KS.DiggList.asp?Action=Del"">")
		.echo "    <tr class='sort'>"
		.echo "    <td width='30' align='center'>选中</td>"
		.echo "    <td align='center'>被推荐的文档</td>"
		.echo "    <td width='10%' align='center'>顶</td>"
		.echo "    <td width='10%' align='center'>踩</td>"
		.echo "    <td width='20%' align='center'>最后推荐时间</td>"
		.echo "    <td width='15%' align='center'>最后推荐的用户</td>"
		.echo "    <td width='8%' align='center'>浏览</td>"
		.echo "  </tr>"
		 Set RS = Server.CreateObject("ADODB.RecordSet")
		   Dim Param:Param=" where a.ChannelID=b.ChannelID"
		   Select Case KS.ChkClng(KS.G("Show"))
		     case 1  Param=Param & " Order By A.DiggNum Desc,A.DiggID Desc"
			 case 2  Param=Param & " Order By A.DiggNum Asc,A.DiggID Desc"
			 case 3  Param=Param & " Order By A.CDiggNum Desc,A.DiggID Desc"
			 case 4  Param=Param & " Order By A.CDiggNum Asc,A.DiggID Desc"
			 case 5  Param=Param & " order by a.LastDiggTime asc,A.DiggID Desc"
			 Case Else  Param=Param & " order by a.LastDiggTime desc,A.DiggID Desc"
		   End Select
		   
				   SqlStr = "SELECT a.*,b.title FROM [KS_DiggList] a inner join KS_ItemInfo b on a.infoid=b.infoid " & Param
				   RS.Open SqlStr, conn, 1, 1
				 If RS.EOF And RS.BOF Then
				  .echo "<tr><td  class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"" colspan=6 height='25' align='center'>没有推荐文档!</td></tr>"
				 Else
					        totalPut = RS.RecordCount
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							End If
							Call showContent
			End If
		  .echo "  </td>"
		  .echo "</tr>"

		 .echo "</table>"
		 .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
		 .echo ("<tr><td width='170'><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> </div>")
		 .echo ("</td>")
	     .echo ("<td><input type=""button"" value=""删除选中的文档"" onclick=""DelDiggList();"" class=""button""></td>")
	     .echo ("</form><td align='right'>")
			 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	     .echo ("</td></tr></form></table>")
		 .echo ("<form action='KS.Digglist.asp?action=DelAllRecord' method='post' target='_hiddenframe'>")
		 .echo ("<iframe src='about:blank' style='display:none' name='_hiddenframe' id='_hiddenframe'></iframe>")
		 .echo ("<div class='attention'><strong>特别提醒： </strong><br>当站点运行一段时间后,网站的digg记录表可能存放着大量的记录,为使系统的运行性能更佳,建议一段时间后清理一次。")
		 .echo ("<br /> <strong>删除范围：</strong><input name=""deltype"" type=""radio"" value=1>10天前 <input name=""deltype"" type=""radio"" value=""2"" /> 1个月前 <input name=""deltype"" type=""radio"" value=""3"" />2个月前 <input name=""deltype"" type=""radio"" value=""4"" />3个月前 <input name=""deltype"" type=""radio"" value=""5"" /> 6个月前 <input name=""deltype"" type=""radio"" value=""6"" checked=""checked"" /> 1年前  <input onclick=""$(parent.frames['FrameTop'].document).find('#ajaxmsg').toggle();"" type=""submit""  class=""button"" value=""执行删除"">")
		 .echo ("</div>")
		 .echo ("</form>")
		End With
		End Sub
		Sub showContent()
		  With KS
			 Do While Not RS.EOF
			.echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & RS("DiggID") & "' onclick=""chk_iddiv('" & RS("DiggID") & "')"">"
		   .echo "<td class='splittd' align='center'><input name='id' onclick=""chk_iddiv('" &RS("diggID") & "')"" type='checkbox' id='c"& RS("DiggID") & "' value='" &RS("DiggID") & "'></td>"
		  .echo " <td class='splittd' height='22'><img src='Images/folder/TheSmallWordNews1.gif' align='absmiddle'><span style='cursor:default;' title='" & RS("Title") & "'>"
		   .echo  KS.Gottopic(RS("Title"),36) & "</td>"
		   
		   .echo " <td class='splittd' align='center'>" & RS("DiggNum") & " 次</td>"
		   .echo " <td class='splittd' align='center'>" & RS("CDiggNum") & " 次</td>"
		   .echo " <td class='splittd' align='center'>" & RS("LastDiggTime") & "</td>"
		   .echo " <td class='splittd' align='center'>" & RS("LastDiggUser") & " </td>"
		   .echo " <td class='splittd' align='center'><a href='javascript:void(0)' onclick=""show('" & RS("Title") & "'," & RS("ChannelID") & "," & RS("Infoid") & ")"">记录</a> <a href='../item/show.asp?m=" & rs("channelid") & "&d=" & rs("infoid") & "' target='_blank'>浏览</a> </td>"
		   .echo "</tr>"
							  I = I + 1
								If I >= MaxPerPage Then Exit Do
							   RS.MoveNext
							   Loop
								RS.Close
								 
		  End With
		 End Sub
		 
		 '调用代码
		 Sub ShowCode()
		 %>
	   <table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr class='sort'>
           <td>&nbsp;&nbsp;
               <div align="left"><strong> 效果一</strong> 将以下代码复制到内容页模板即可 </div></td>
         </tr>
         <tr>
           <td align="center" nowrap="nowrap"><style type="text/css">
			.mark {overflow:hidden;padding:15px 0 20px 111px; clear:both;}
			#mark0, #mark1 {background:url(../images/default/mark.gif) no-repeat -189px 0;border:0;cursor:pointer;float:left;height:48px;margin:0;overflow:hidden;padding:0;position:relative;width:189px;}
			#mark1 {background-position:-378px 0;margin-left:10px;}
			#barnum1, #barnum2 {color:#333333;font-family:arial;font-size:10px;font-weight:400;left:70px;line-height:12px;position:absolute;top:30px;}
			.bar {background-color:#FFFFFF;border:1px solid #40A300;height:5px;left:9px;overflow:hidden;position:absolute;text-align:left;top:32px;width:55px;}
			.bar div {background:transparent url(../images/default/bar_Footbg.gif) repeat-x ;height:5px;overflow:hidden; margin:0;}
			#mark1 .bar {border-color:#555555;}
			#mark1 .bar div {background:transparent url(../images/default/Barbg.gif) repeat-x ;}
			</style>
               <div class="mark">
                 <div  onfocus="this.blur()" onmouseout="this.style.backgroundPosition='-189px 0'" onmouseover="this.style.backgroundPosition='0 0'" id="mark0" style="background-position: -189px 0pt;">
                   <div class="bar">
                     <div style="width: 0px;" id="digzcimg"></div>
                   </div>
                   <span id="barnum1"><span id="perz">0%</span> (<span id="s">0</span>)</span> </div>
                 <div  onfocus="this.blur()" onmouseout="this.style.backgroundPosition='-378px 0'" onmouseover="this.style.backgroundPosition='-567px 0'" id="mark1" style="background-position: -378px 0pt;">
                   <div class="bar">
                     <div style="width: 0px;" id="digcimg"></div>
                   </div>
                   <span id="barnum2"><span id="perc1">0%</span> (<span id="ca">0</span>)</span> </div>
               </div>
             <textarea onmouseover="javascript:this.select();" name="textarea" style="width:650px;height:90px"><style type="text/css">
	.mark {overflow:hidden;padding:15px 0 20px 111px; clear:both;}
	#mark0, #mark1 {background:url({$GetSiteUrl}images/default/mark.gif) no-repeat -189px 0;border:0;cursor:pointer;float:left;height:48px;margin:0;overflow:hidden;padding:0;position:relative;width:189px;}
	#mark1 {background-position:-378px 0;margin-left:10px;}
	#barnum1, #barnum2 {color:#333333;font-family:arial;font-size:10px;font-weight:400;left:70px;line-height:12px;position:absolute;top:30px;}
	.bar {background-color:#FFFFFF;border:1px solid #40A300;height:5px;left:9px;overflow:hidden;position:absolute;text-align:left;top:32px;width:55px;}
	.bar div {background:transparent url({$GetSiteUrl}images/default/bar_Footbg.gif) repeat-x ;height:5px;overflow:hidden; margin:0;}
	#mark1 .bar {border-color:#555555;}
	#mark1 .bar div {background:transparent url({$GetSiteUrl}images/default/Barbg.gif) repeat-x ;}
</style>
<script language="JavaScript" src="{$GetSiteUrl}ks_inc/digg.js" type="text/javascript"></script> 
	<div class="mark">
	 <div onClick="digg({$ChannelID},{$InfoID},'{$GetSiteUrl}');" onfocus="this.blur()" onMouseOut="this.style.backgroundPosition='-189px 0'" onMouseOver="this.style.backgroundPosition='0 0'" id="mark0" style="background-position: -189px 0pt;">
	 <div class="bar"><div style="width: 0px;" id="digzcimg"></div></div>
	 <span id="barnum1"><span id="perz{$InfoID}">0%</span> (<span id="s{$InfoID}">0</span>)</span>
     </div>
	 <div onClick="cai({$ChannelID},{$InfoID},'{$GetSiteUrl}');" onfocus="this.blur()" onMouseOut="this.style.backgroundPosition='-378px 0'" onMouseOver="this.style.backgroundPosition='-567px 0'" id="mark1" style="background-position: -378px 0pt;">
	<div class="bar"><div style="width: 0px;" id="digcimg"></div></div>
	 <span id="barnum2"><span id="perc{$InfoID}">0%</span> (<span id="c{$InfoID}">10</span>)</span>
	</div>
</div>
<script language="JavaScript" type="text/javascript">show_digg({$ChannelID},{$InfoID},'{$GetSiteUrl}');</script>
         </textarea></td>
         </tr>
       </table>
	   <table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr class='sort'>
           <td>&nbsp;&nbsp;
               <div align="left"><strong> 效果二</strong> 将以下代码复制到内容页模板即可 </div></td>
         </tr>
         <tr>
           <td align="center" nowrap="nowrap">
		   
<div style="width:153px">									
<div style="float:left;BACKGROUND: url(../images/default/ding_bg.gif) no-repeat; WIDTH: 58px; HEIGHT: 62px">
<div style="PADDING-TOP: 7px; TEXT-ALIGN: center"><SPAN id=s{$InfoID} style="FONT-WEIGHT: bold; COLOR: #fff">0</SPAN></div>
                                    <div id=d{$InfoID} style="padding-top:25px;HEIGHT: 25px; TEXT-ALIGN: center"><a href="#">顶一下</a> </div></div>
									
<div style="float:right;BACKGROUND: url(../images/default/ding_bg.gif) no-repeat; WIDTH: 58px; HEIGHT: 62px">
<div style="PADDING-TOP: 7px; TEXT-ALIGN: center"><SPAN id=c{$InfoID} style="FONT-WEIGHT: bold; COLOR: #fff">0</SPAN></div>
                                    <div id=d{$InfoID} style="padding-top:25px;HEIGHT: 25px; TEXT-ALIGN: center"><a href="#">踩一下</a> </div></div>
</div>
			<div style="clear:both"></div>
             <textarea onmouseover="javascript:this.select();" name="textarea2" style="width:650px;height:90px"><script language="JavaScript" src="{$GetSiteUrl}ks_inc/digg.js" type="text/javascript"></script> 
<div style="width:153px">									
<div style="float:left;BACKGROUND: url({$GetSiteUrl}images/default/ding_bg.gif) no-repeat; WIDTH: 58px; HEIGHT: 62px">
<div style="PADDING-TOP: 7px; TEXT-ALIGN: center"><SPAN id=s{$InfoID} style="FONT-WEIGHT: bold; COLOR: #fff"></SPAN></div>
                                    <div id=d{$InfoID} style="padding-top:15px;HEIGHT: 25px; TEXT-ALIGN: center"><a href="javascript:digg({$ChannelID},{$InfoID},'{$GetSiteUrl}');">顶一下</a> </div></div>
									
<div style="float:right;BACKGROUND: url({$GetSiteUrl}images/default/ding_bg.gif) no-repeat; WIDTH: 58px; HEIGHT: 62px">
<div style="PADDING-TOP: 7px; TEXT-ALIGN: center"><SPAN id=c{$InfoID} style="FONT-WEIGHT: bold; COLOR: #fff"></SPAN></div>
                                    <div id=d{$InfoID} style="padding-top:15px;HEIGHT: 25px; TEXT-ALIGN: center"><a href="javascript:cai({$ChannelID},{$InfoID},'{$GetSiteUrl}');">踩一下</a> </div></div>
<script language="JavaScript" type="text/javascript">show_digg({$ChannelID},{$InfoID},'{$GetSiteUrl}');</script>
</div>
             </textarea></td>
         </tr>
       </table>
	   <%
		 End Sub
		 
		 Sub ItemDelete()
			Dim ID:ID = KS.G("ID")
			If KS.IsNul(ID) Then KS.AlertHintScript "对不起，没有选择记录！"
			conn.Execute ("Delete From KS_Digg Where DiggID IN(" & ID & ")")
			conn.execute("delete from ks_digglist where diggid in(" &id & ")")
		    response.redirect request.servervariables("http_referer") 
		 End Sub
		 
		 Sub DiggDel()
			Dim ID:ID = KS.G("ID")
			Dim IDArr:IDArr=Split(id,",")
			Dim I
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			For I=0 To Ubound(IDArr)
			 RS.Open "Select DiggID From KS_Digg Where ID=" & IDArr(i),conn,1,3
			 If Not RS.Eof Then
			  'Conn.Execute("Update KS_DiggList Set DiggNum=DiggNum-1 Where DiggID=" & RS(0))
			  RS.Delete
			 End iF
			 RS.Close
			Next
			Set RS=Nothing
		    response.redirect request.servervariables("http_referer") 
		 End Sub
		 
		 Sub DelAllRecord()
		  Dim Param
		  Select Case KS.ChkClng(KS.G("DelType"))
		   Case 1 Param="datediff(" & DataPart_D & ",DiggTime," & SqlNowString & ")>11"
		   Case 2 Param="datediff(" & DataPart_D & ",DiggTime," & SqlNowString & ")>31"
		   Case 3 Param="datediff(" & DataPart_D & ",DiggTime," & SqlNowString & ")>61"
		   Case 4 Param="datediff(" & DataPart_D & ",DiggTime," & SqlNowString & ")>91"
		   Case 5 Param="datediff(" & DataPart_D & ",DiggTime," & SqlNowString & ")>181"
		   Case 6 Param="datediff(" & DataPart_D & ",DiggTime," & SqlNowString & ")>366"
		  End Select
   		  If Param<>"" Then Conn.Execute("Delete From KS_Digg Where " & Param)
          KS.echo "<script>$(parent.document).find('#ajaxmsg').toggle();alert('恭喜,删除指定日期digg记录成功!');</script>"
		 End Sub
		 
		 Sub DiggDetail()
		  With KS
			 .echo "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
			 .echo "    <tr class='sort'>"
			 .echo "    <td width='30' align='center'>选中</td>"
			 .echo "    <td align='center'>被推荐的文档</td>"
			 .echo "    <td width='15%' align='center'>推荐用户</td>"
			 .echo "    <td width='8%' align='center'>类型</td>"
			 .echo "    <td width='20%' align='center'>推荐时间</td>"
			 .echo "    <td width='15%' align='center'>推荐用户IP</td>"
			 .echo "    <td width='8%' align='center'>浏览</td>"
			 .echo "  </tr>"
			.echo "<form name=""myform"" id=""myform"" method=""Post"" action=""KS.DiggList.asp?action=DiggDel&ChannelID=" & ChannelID & """>"
              MaxPerPage=12
			  Dim RS,XML,Node
			  Set RS = Server.CreateObject("ADODB.RecordSet")
			  Dim Param:Param=" where b.ChannelID="& ChannelID
			  IF KS.ChkClng(KS.G("InfoID"))<>0 Then  Param=Param & " And a.InfoID=" & KS.ChkClng(KS.G("InfoID"))
				  Param=Param & " order by a.DiggTime desc"
				   SqlStr = "SELECT a.*,b.title FROM [KS_Digg] a inner join [KS_ItemInfo] b on a.infoid=b.infoid " & Param
				   
					  RS.Open SqlStr, conn, 1, 1
					 If RS.EOF And RS.BOF Then
					   RS.Close:Set RS=Nothing
					   .echo "<tr><td  class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"" colspan=6 height='25' align='center'>没有推荐记录!</td></tr>"
					 Else
						        totalPut = RS.RecordCount
								
								If CurrentPage >1 And (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
								End If
								Set XML=.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","root")
								RS.Close
								Set RS=Nothing
								If IsObject(XML) Then 
								  For Each Node In XML.DocumentElement.SelectNodes("row")
			                         .echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & Node.SelectSingleNode("@id").text & "' onclick=""chk_iddiv('" & Node.SelectSingleNode("@id").text & "')"">"
									 .echo "<td class='splittd' align='center'><input name='id' onclick=""chk_iddiv('" & Node.SelectSingleNode("@id").text & "')"" type='checkbox' id='c"& Node.SelectSingleNode("@id").text & "' value='" & Node.SelectSingleNode("@id").text & "'></td>"
									 .echo " <td height='22' class='splittd'><img src='Images/folder/TheSmallWordNews1.gif' align='absmiddle'><span style='cursor:default;'>"
									 .echo  Node.SelectSingleNode("@title").text & "</td>"
									   
									 .echo " <td class='splittd' align='center'>" & Node.SelectSingleNode("@username").text & "</td>"
									 .echo " <td class='splittd' align='center'>" 
									 if Node.SelectSingleNode("@diggtype").text = "0" then
									  .echo "顶"
									 else
									  .echo "踩"
									 end if
									 .echo " </td>"
									 .echo " <td class='splittd' align='center'>" & Node.SelectSingleNode("@diggtime").text & "</td>"
									 .echo " <td align='center' class='splittd'>" & Node.SelectSingleNode("@userip").text & " </td>"
									 .echo " <td class='splittd' align='center'><a href='../item/show.asp?m=" & Node.SelectSingleNode("@channelid").text &"&d=" & Node.SelectSingleNode("@infoid").text & "' target='_blank'>浏览</a> </td>"
									 .echo "</tr>"
								  Next
								End If
								Set XML=Nothing
				End If
				
			  .echo "  </td>"
			  .echo "</tr>"
	
			  .echo "</table>"
			  .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
			 .echo ("<tr><td width='160'><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> </div>")
			 .echo ("</td>")
			 .echo ("<td><input type=""button"" value=""删除记录"" onclick=""DelDigg();"" class=""button""></td>")
			 .echo ("</form><td align='right'>")
				 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			 .echo ("</td></tr></form></table>")
		End With
		 
		 End Sub
End Class
%> 
