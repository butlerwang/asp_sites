<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<!--#include file="Label/LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New CirLabel
KSCls.Kesion()
Set KSCls = Nothing

Class CirLabel
        Private KS,TempClassList, InstallDir, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript
		Dim ChannelID,ShowClassName, ArticleListNumber, RowHeight, TitleLen, ArticleSort, ShowPicFlag,DateRule, DateAlign,ShowNewFlag,ShowHotFlag, PrintType,XslContent
		 Dim LabelRS, LabelName,innersql,outsql

		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Call KS.DelCahe(KS.SiteSn & "_cirlabellist")
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()

		FolderID = Request("FolderID")
		ChannelID=KS.ChkCLng(Request("ChannelID"))
		If ChannelID=0 Then ChannelID=1
		
		With KS
				 .echo "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd""><html xmlns=""http://www.w3.org/1999/xhtml"">"
				.echo "<head>"
				.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
				.echo "<link href=""admin_style.css"" rel=""stylesheet"">"
				.echo "<script src=""../../ks_inc/Common.js"" language=""JavaScript""></script>"
				.echo "<script src=""../../ks_inc/jquery.js"" language=""JavaScript""></script>"
                .echo "<script src='../../ks_inc/lhgdialog.js'></script>"
		If KS.G("Action")="DoSave" Then Call DoSave()
        
		'判断是否编辑
		LabelID = Trim(Request.QueryString("LabelID"))
		If LabelID = "" Then
		 outsql="SELECT TOP 10 ID,FolderName FROM [KS_Class] Where ChannelID=1 and ClassType=1 ORDER BY FolderOrder"
		 innersql="SELECT TOP 10 id,title,adddate FROM [KS_Article] Where Tid='{R:ID}' Order By ID Desc"
		 XslContent="<?xml version=""1.0"" encoding=""utf-8""?>"&vbcrlf & _  
                     "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">"&vbcrlf & _  
					 "<xsl:output method=""xml"" omit-xml-declaration=""yes"" indent=""yes"" version=""4.0""/>" & vbcrlf & _
					 "<xsl:template match=""/"">"&vbcrlf & _ 
					 " <div class=""class_loop"">"&vbcrlf & _ 
					 "  <xsl:for-each select=""xml/outerlist/outerrow"">"&vbcrlf & _ 
					 "   <div class=""loop_content"">"&vbcrlf & _  
					 "     <div class=""loop_title"">"&vbcrlf & _  
					 "       <span class=""classname""><a href=""{@classlink}""><xsl:value-of select=""@foldername"" disable-output-escaping=""yes"" /></a></span>"&vbcrlf & _ 
					 "       <span class=""class_more""><a href=""{@classlink}"">更多</a></span>"&vbcrlf & "     </div>" &vbcrlf & _ 
					 "     <div class=""loop_list"">"&vbcrlf & _ 
					 "      <ul>"&vbcrlf & _ 
					 "      <xsl:for-each select=""innerlist/innerrow"">"&vbcrlf & _  
					 "      <li><a href=""{@linkurl}"" title=""{@title}"" target=""_blank"">{KS:CutText(<xsl:value-of select=""@title"" disable-output-escaping=""yes""/>,20,""..."")}</a></li>"&vbcrlf & _  
					 "      </xsl:for-each>"&vbcrlf & _ 
					 "      </ul>"&vbcrlf & _ 
					 "     </div>"&vbcrlf & _
					 "  </div>"&vbcrlf &vbcrlf & _ 
					 "   </xsl:for-each>"&vbcrlf & _  
					 "   </div>"&vbcrlf & _  
					 "</xsl:template>"&vbcrlf & _  
					 "</xsl:stylesheet>"
		Else
		  Set LabelRS = Server.CreateObject("Adodb.Recordset")
		  LabelRS.Open "Select * From KS_Label Where ID='" & LabelID & "'", Conn, 1, 1
		  If LabelRS.EOF And LabelRS.BOF Then
			 LabelRS.Close
			 Set LabelRS = Nothing
			 .echo ("<Script>alert('参数传递出错!');window.close();</Script>")
			 .End
		  End If
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			FolderID = LabelRS("FolderID")
			Descript = Split(LabelRS("Description"),"@@@")
			LabelContent = LabelRS("LabelContent")
			LabelRS.Close
			Set LabelRS = Nothing
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetCirList", ""),"}{/Tag}", "")
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('标签加载出错!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
			ChannelID=Node.getAttribute("channelid")
			DateRule = Node.getAttribute("daterule")
			End If 
			XmlDoc=Empty
			Set Node=Nothing
			OutSql=Descript(0)
			InnerSql=Descript(1)
			XslContent=Descript(2)
		End If
		If PrintType="" Then PrintType=1
		%>
		<script language="javascript">

		function CheckForm()
		{   if ($('input[name=LabelName]').val()=='')
			 {
			  alert('请输入标签名称');
			  $('input[name=LabelName]').focus(); 
			  return false
			  }
			var ChannelID=1;
			var DateRule=document.myform.DateRule.value;
	
			document.myform.LabelContent.value=	'{Tag:GetCirList labelid="0" channelid="'+$('#channelid').val()+'" daterule="'+DateRule+'"}{/Tag}';
			document.myform.submit();
		}
		</script>
		<%
		.echo "</head>"
		.echo "<body topmargin=""0"" leftmargin=""0"">"
		.echo "<div align=""center"">"
		.echo "<form  method=""post"" name=""myform"" action=""CirLabel.asp"">"
		.echo " <input type=""hidden"" name=""LabelContent"">"
		.echo " <input type=""hidden"" name=""LabelFlag"" value=""1"">"
		.echo " <input type=""hidden"" name=""Action"" value=""DoSave"">"
		 .echo " <table width='100%' height='25' border='0' cellpadding='0' cellspacing='1' bgcolor='#efefef' class='sort'>"
		 .echo "       <tr><td><div align='center'><font color='#990000'>"
		 .echo " 通用循环标签"
		 .echo "    </font></div></td></tr>"
		 .echo "    </table>"
%>
       <table border='0' cellspacing='1' cellpadding='1' width='98%' align='center' class='ctable'>
		  <form action="?" method="post" name="myform">
		   <input name='lbtf' type='hidden'>
		   <input type='hidden' name='labelid' value='<%=labelid%>'>
		  <tr class='tdbg'>
		    <td class='clefttitle' align='right'><strong>标签名称:</strong></td>
		    <td><input name="LabelName" value="<%=LabelName%>" onblur='testlabelname()' style="width:200;"> <font color=red>*</font><span id='labelmessage'></span>例如标签名称：&quot;循环文章列表&quot;，则在模板中调用：<font color="#FF0000">&quot;{LB_循环文章列表}&quot;</font>。</td>
		  </tr>
		  <tr class='tdbg'>
		   <td width='100'  height="30" class='clefttitle' align='right'><strong>标签目录:</strong></td>
		   <td><%=ReturnLabelFolderTree(FolderID, 6)%><font color=""#FF0000"">请选择标签归属目录，以便日后管理标签</font></td>
		  </tr>
		
<%
		
		.echo "  <tr class=tdbg>"
		.echo "    <td height=""24"" align='right' class='clefttitle'><strong>输出格式:</strong></td>"
		.echo "    <td><span style='display:none'><select class='textbox'  name=""PrintType"">"
        .echo "  <option value=""1"""
		If PrintType="1" Then .echo " selected"
		.echo ">普通格式</option>"
        .echo " <option value=""3"""
		If PrintType="3" Then .echo " selected"
		.echo ">Ajax输出</option>"
        
        .echo "</select></span>"
		.echo "       内层SQL查询的模型<select name='channelid' id='channelid'><option value='0'>--选择模型--</option>"
		Dim RSC:Set RSC=conn.execute("select ChannelID,ChannelName From KS_Channel Where ChannelStatus=1 and channelid<>6 and channelid<>9 and channelid<>10 order by channelid")
		do while not rsc.eof
		 if trim(channelid)=trim(rsc(0)) then
		 .echo "<option value='" & rsc(0) & "' selected>" & rsc(1) & "</option>"
		 else
		 .echo "<option value='" & rsc(0) & "'>" & rsc(1) & "</option>"
		 end if
		rsc.movenext
		loop
		rsc.close
		set rsc=nothing
		.echo " </select>日期样式 "
		.echo ReturnDateFormat(DateRule)
		.echo "               "
		.echo "<br><font color=green>tips:当使用标签@linkurl时，必须正确选择模型，否则得不到正确的信息url</font></td>"


		.echo "              </tr>"

		.echo "            <tr class='tdbg'>"
		.echo "            <td class='clefttitle' align='right'><b>SQL语句：</b></td>"
		.echo "            <td> <strong>外层SQL语句：</strong><textarea name='outsql' style='width:90%;height:40px'>" & outsql & "</textarea>"
		.echo "            <br><strong>内层SQL语句：</strong>"
		.echo "            <textarea name='innersql' style='width:90%;height:40px'>" & innersql & "</textarea>"
		.echo "<br>SQL语句可用标签：当前栏目ID<font color=red>{$CurrClassID}</font>;当前栏目ID及子目录ID集<font color=red>{$CurrClassChildID}</font>;<br>内层SQL可以用标签<font color=green>{R:字段名}</font>与外层SQL标签关联。"
		.echo "            </td>"
		.echo "            </tr>"
		
		.echo "            <tr class='tdbg'>"
		.echo "            <td class='clefttitle' align='right'><strong>XSLT标签说明：</strong></td>"
		.echo "            <td>预设标签：<font color=red> @classlink</font> 得到栏目链接 <font color=red>@linkurl</font> 得到信息链接<br> 标签的构造规则：根据所查询的字段名前加<font color=blue>@</font>组成，且所有字段名必须小写；<br>如：select top 10 <font color=red>id,title</font> from ks_article 则可用字段为<font color=red>@id</font>及<font color=red>@title</font>两个。"
		.echo "            </td>"
		.echo "            </tr>"
		
		.echo "  <tr class=tdbg>"
		.echo "    <td height=""24"" align='right' class='clefttitle'><strong>XSLT样式:</strong></td>"
		.echo "    <td><textarea name='xslContent' style='width:98%;height:250px'>" & XslContent & "</textarea>"
		.echo "    <br><font color=blue>说明：xlst样式必须严格按照xslt语法编写。<br>内置截取字符长度函数<font color=red>{KS:CutText(title,len,'...')} </font> </font><br>内置函数使用说明：<br><font color=red>title</font>要截取的内容;<br><font color=red>len</font>截取字符数，一个汉字算两个字符;<br> <font color=red>...</font>显示被截取后的省略符</td>"
		.echo "           </tr>"

		.echo "                  </table>"	
		.echo "</form>"
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
		End With
		End Sub
		
		'保存
		Sub DoSave()
					LabelName = KS.G("LabelName")
					LabelID  = KS.G("LabelID")
					Descript = Request("LabelIntro")
					LabelContent = Trim(Request.Form("LabelContent"))
					FolderID = KS.G("ParentID")
					If LabelName = "" Then
					   Call KS.AlertHistory("标签名称不能为空!", -1)
					   Set KS = Nothing
					   Exit Sub
					End If
					
					If LabelContent = "" Then
					  Call KS.AlertHistory("标签内容不能为空!", -1)
					  Set KS = Nothing
					  Exit Sub
					End If
					LabelName = "{LB_" & LabelName & "}"
					Set LabelRS = Server.CreateObject("Adodb.RecordSet")
					LabelRS.Open "Select LabelName From [KS_Label] Where ID<>'" & LabelID & "' and LabelName='" & LabelName & "'", Conn, 1, 1
					If Not LabelRS.EOF Then
					  Call KS.AlertHistory("标签名称已经存在!", -1)
					  LabelRS.Close
					  Conn.Close
					  Set LabelRS = Nothing
					  Set Conn = Nothing
					  Set KS = Nothing
					 Exit Sub
					Else
						LabelRS.Close
						LabelRS.Open "Select * From [KS_Label] Where ID='" & LabelID & "'", Conn, 1, 3
						If LabelRS.Eof Then
						 LabelRS.AddNew
						  Do While True
							'生成ID  年+12位随机
							LabelID = Year(Now()) & KS.MakeRandom(10)
							Dim RSCheck:Set RSCheck = Conn.Execute("Select ID from [KS_Label] Where ID='" & LabelID & "'")
							 If RSCheck.EOF And RSCheck.BOF Then
							  RSCheck.Close
							  Set RSCheck = Nothing
							  Exit Do
							 End If
						  Loop
						 LabelRS("ID") = LabelID
						 LabelRS("AddDate") = Now
						 LabelRS("LabelType") = 6
						 LabelRS("OrderID") = 1
						 LabelRS("LabelFlag") = 6
						End If
						 LabelRS("LabelName") = LabelName
						 LabelRS("LabelContent") = LabelContent
						 LabelRS("Description") = Request("OutSql") & "@@@" & Request("InnerSQL") & "@@@" & Request("xslContent")
						 LabelRS("FolderID") = FolderID
						 LabelRS.Update
						 If KS.G("LabelID")="" Then
						  Call KS.FileAssociation(1021,1,LabelContent&Request("xslContent"),1)
						 ks.die "<script>$.dialog.confirm('恭喜，添加标签成功,继续添加标签吗?',function(){location.href='CirLabel.asp?Action=AddNew&LabelType=6&FolderID=" & FolderID & "';},function(){$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & FolderID & "&OpStr=标签管理 >> 循环标签&ButtonSymbol=FreeLabel';parent.frames['MainFrame'].location.href='Label_Main.asp?LabelType=6&FolderID=" & FolderID & "';});</script>"
						Else
						 	 '遍历所有标签内容，找出所有标签的图片
							 Dim Node,UpFiles,RCls
							 UpFiles=LabelContent&Request("xslContent")
							 if Not IsObject(Application(KS.SiteSN&"_labellist")) Then
								 Set RCls=New Refresh
								 Call Rcls.LoadLabelToCache()
								 Set Rcls=Nothing
							 End If
							 For Each Node in Application(KS.SiteSN&"_labellist").DocumentElement.SelectNodes("labellist")
								   UpFiles=UpFiles & Node.Text
							 Next
							 Call KS.FileAssociation(1021,1,UpFiles,1)
							 '遍历及入库结束
				   	         KS.Echo "<script>$.dialog.tips('<br/>恭喜，标签修改成功!',1,'success.gif',function(){$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & FolderID & "&OpStr=标签管理 >> 循环标签&ButtonSymbol=FreeLabel';parent.frames['MainFrame'].location.href='Label_Main.asp?LabelType=6&FolderID=" & FolderID & "';});</script>"
						End If
					End If
			End Sub
End Class
%> 
