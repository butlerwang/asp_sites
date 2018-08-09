<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.KeyCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
Server.ScriptTimeOut=9999999

Dim KSCls
Set KSCls = New Tools
KSCls.Kesion()
Set KSCls = Nothing

Class Tools
        Private KS,KSCls,ChannelID
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		Sub Kesion()
		 With KS
			If Not KS.ReturnPowerResult(0, "KSO10004") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
			End if
			
			select case KS.S("Action")
			  case "setKeyWords" setKeyWords:response.end
			  case "relativeDoc" relativeDoc:response.end
			  case "getDocImage" getDocImage:response.end
			  case "DocContent" DocContent:response.End
			  case "checkDocFname" checkDocFname:response.end
			  case "UpClubData" UpClubData:response.end
			  case "UpGuestBoard" UpGuestBoard:response.end
			  case "UpUserPostNum" UpUserPostNum:response.end
			end select
			.echo "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd"">"
			.echo "<html xmlns=""http://www.w3.org/1999/xhtml"">"
			.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			.echo "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			%>
			<script type="text/javascript">
			 function checkModelType(channelid)
			 {
			   if (channelid==0) {alert('请选择文章类模型!');return false;}
			   $.get("../plus/ajaxs.asp",{action:"getModelType",channelid:channelid},function(t){
			     if (t!=1)
				 {
				  alert('对不起,你选择的基类型不是文章!');
				  return false;
				 }else{ document.DocImage.submit();}
			   });
			 }
			</script>
			<%
			.echo "</head>"
			
			.echo "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.echo "      <div class='topdashed sort'> 一键相关管理工具</div>"
			
			.echo "<div class='attention'><strong>操作说明：</strong><br>"
			.echo "       1、文章条数可以输入""0"" 表示全部执行<br /> "
			.echo "      2、当你的文档较多时，运行此功能可能需要较长时间并在执行期间会占用一些服务器资源，建议选择夜间访问人数少时执行。</div>"
			
			.echo " <table width=""99%"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td style='text-align:left' height=""30"" class='clefttitle'>&nbsp;<font color=""#000080""><b>自动关联相关文档</b></font></td>"
			.echo "    </tr>"
			.echo "    <tr>"
			.echo "      <td width=""25%"" class='tdbg' style='padding:10px'><strong>功能说明：</strong>"
			.echo "       本操作能将文档与文档之间通过各自设定的关键词Tags进行自动关联,以方便供相关文档标签调用时，将关联文档调用出来<br/>"
			.echo "<form action=""KS.Tools.asp?action=relativeDoc"" method=""post"" name=""keyform"" id=""keyform"">"
			.echo "<strong>选项配置：</strong>仅执行最新添加的<input type='text' class='textbox' value='500' size='4' style='text-align:center' name='docNum' id='docNum'>篇文档 选择模型:<select id='channelid' name='channelid'>"
			.echo " <option value='0'>---请选择模型---</option>"
			.LoadChannelOption 0
			
			.echo "</select>      <table border='0' width='98%' align'center'>"
			.echo "       <tr><td><input type='submit' class='button' value='一键自动关联'></td></tr>"
			.echo "      </table>"
			.echo "        </td></form>"
			.echo "    </tr>"
			.echo "  </table>"
			
			.echo " <table width=""99%"" style=""margin-top:5px"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td style='text-align:left' height=""30"" class='clefttitle'>&nbsp;<font color=""#000080""><b>自动提取内容第一张图片为文档首页图片</b></font></td>"
			.echo "    </tr>"
			.echo "    <tr>"
			.echo "      <td width=""25%"" class='tdbg' style='padding:10px'><strong>功能说明：</strong>"
			.echo "       本操作能从基类型为""文章类""且没有设置缩略图的文档内容中的自动提取第一张图片为做为文档的图片,从而自动转为图片文档,供前台标签调用。<br/>"
			.echo "<form action=""KS.Tools.asp?action=getDocImage"" method=""post"" name=""DocImage"" id=""DocImage"">"
			.echo "<strong>选项配置：</strong>仅执行最新添加的<input type='text' class='textbox' value='500' size='4' style='text-align:center' name='docNum' id='docNum'>篇文档 选择模型:<select id='channelid' name='channelid'>"
			.echo " <option value='0'>---请选择模型---</option>"
			
			Dim ModelXML,Node,Pstr:Pstr="@ks21=1 and @ks6=1"
				If Not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.documentElement.SelectNodes("channel[" & Pstr & "]")
				  .echo "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & Node.SelectSingleNode("@ks1").text & "|" & Node.SelectSingleNode("@ks2").text & "</option>"
			Next
			
			
			.echo "</select> "
			.echo "      <table border='0' width='98%' align'center'>"
			.echo "       <tr><td><input type='button' onclick=""checkModelType(document.DocImage.channelid.value)"" class='button' value='一键自动提取'></td></tr>"
			.echo "      </table>"
			.echo "        </td></form>"
			.echo "    </tr>"
			.echo "  </table>"
			
			
			.echo " <table width=""99%"" style=""margin-top:5px"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td style='text-align:left' height=""30"" class='clefttitle'>&nbsp;<font color=""#000080""><b>自动删除无文章内容的记录</b></font></td>"
			.echo "    </tr>"
			.echo "    <tr>"
			.echo "      <td width=""25%"" class='tdbg' style='padding:10px'><strong>功能说明：</strong>"
			.echo "       本操作能从基类型为""文章类""且文章内容为空的记录删除掉,执行此操作前强烈建议先备份数据库。<br/>"
			.echo "<form action=""KS.Tools.asp"" method=""post"" name=""cform"" id=""cform"">"
			.echo "<input type='hidden' name='action' id='action' value='DocContent'/><strong>选项配置：</strong>选择模型:<select id='channelid' name='channelid'>"
			.echo " <option value='0'>---请选择模型---</option>"
			
			Pstr="@ks21=1 and @ks6=1"
			    Dim FieldXML,FieldNode
				If Not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.documentElement.SelectNodes("channel[" & Pstr & "]")
				     Call KSCls.LoadModelField(Node.SelectSingleNode("@ks0").text,FieldXML,FieldNode)
				  If FieldXML.DocumentElement.selectsinglenode("fielditem[@fieldname='content']/showonform").text="1" Then
				  .echo "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & Node.SelectSingleNode("@ks1").text & "|" & Node.SelectSingleNode("@ks2").text & "</option>"
				  End If
			Next
			
			
			.echo "</select> "
			.echo "      <table border='0' width='98%' align'center'>"
			.echo "       <tr><td><input type='submit' onclick=""if(document.cform.channelid.value==0){alert('请选择模型!');return false;}"" class='button' value='一键自动删除无内容记录'></td></tr>"
			.echo "      </table>"
			.echo "        </td></form>"
			.echo "    </tr>"
			.echo "  </table>"

			.echo " <table width=""99%"" style=""margin-top:5px"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td style='text-align:left' height=""30"" class='clefttitle'>&nbsp;<font color=""#000080""><b>自动修正文档文件名</b></font></td>"
			.echo "    </tr>"
			.echo "<form action=""KS.Tools.asp?action=checkDocFname"" method=""post"" name=""DocFname"" id=""DocFname"">"
			.echo "    <tr>"
			.echo "      <td width=""25%"" class='tdbg' style='padding:10px'><strong>功能说明：</strong>"
			.echo "       本操作能自动检测文档生成静态Html的文件名是否合法,如果不合法将自动修正,以免在生成静态操作时出错。<br/><br/>"
			.echo "      <table border='0' width='98%' align'center'>"
			.echo "       <tr><td><input type='submit' class='button' value='一键修正文件名'></td></tr>"
			.echo "      </table>"
			.echo "        </td></form>"
			.echo "    </tr>"
			.echo "  </table>"
			
			.echo " <table width=""99%"" style=""margin-top:5px"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td style='text-align:left' height=""30"" class='clefttitle'>&nbsp;<font color=""#000080""><b>自动设置关键字</b></font></td>"
			.echo "    </tr>"
			.echo "    <tr>"
			.echo "      <td width=""25%"" class='tdbg' style='padding:10px'><strong>功能说明：</strong>"
			.echo "       本操作能自动检测文档是否有关键词Tags，如果没有设置的自动按文档标题自动生成关键词Tags。<br/>"
			.echo "<form action=""KS.Tools.asp?action=setKeyWords"" method=""post"" name=""keyform"" id=""keyform"">"
			.echo "<strong>选项配置：</strong>仅执行最新添加的<input type='text' class='textbox' value='500' size='4' style='text-align:center' name='docNum' id='docNum'>篇文档 选择模型:<select id='channelid' name='channelid'>"
			.echo " <option value='0'>---请选择模型---</option>"
			.LoadChannelOption 0
			
			.echo "</select>  <label><input type='checkbox' name='setall' value='1'>对所有文档执行重置关键词</label>"
			.echo "      <table border='0' width='98%' align'center'>"
			.echo "       <tr><td><input type='submit' onclick=""if (document.keyform.channelid.value=='0'){alert('请选择要设置的模型!');return false;}"" class='button' value='一键设置关键字'></td></tr>"
			.echo "      </table>"
			.echo "        </td></form>"
			.echo "    </tr>"
			.echo "  </table>"
			
		If KS.Setting(56)="1" Then
			.echo " <a name=""club""></a><table width=""99%"" style=""margin-top:5px"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td style='text-align:left' height=""30"" class='clefttitle'>&nbsp;<font color=""#000080""><b>一键重新统计论坛总数据</b></font></td>"
			.echo "    </tr>"
			.echo "    <tr>"
			.echo "      <td width=""25%"" class='tdbg' style='padding:10px'><strong>功能说明：</strong>"
			.echo "       这里将重新计算整个论坛的帖子主题和回复数，今日帖子，最后加入用户等，建议每隔一段时间运行一次。<br/>"
			.echo "<form action=""KS.Tools.asp?action=UpClubData"" method=""post"" name=""clubform"" id=""clubform"">"
			
			Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
			
			.echo "<strong>选项配置：</strong>主题数:<input style='text-align:center' class='textbox' type='text' name='TopicNum' value='" & doc.documentElement.attributes.getNamedItem("topicnum").text & "' size='5'> 总帖子数：<input style='text-align:center' class='textbox' type='text' name='PostNum' value='" & doc.documentElement.attributes.getNamedItem("postnum").text & "' size='5'> 今日发帖：<input style='text-align:center' class='textbox' type='text' name='TodayNum' value='" & doc.documentElement.attributes.getNamedItem("todaynum").text & "' size='5'> 昨日发帖数：<input style='text-align:center' class='textbox' type='text' name='YesterDayNum' value='" & doc.documentElement.attributes.getNamedItem("yesterdaynum").text & "' size='5'> 总会员：<input style='text-align:center' class='textbox' type='text' name='TotalUserNum' value='" & doc.documentElement.attributes.getNamedItem("totalusernum").text & "' size='5'> 最高发帖数：<input style='text-align:center' class='textbox' type='text' name='MaxDayNum' value='" & doc.documentElement.attributes.getNamedItem("maxdaynum").text & "' size='5'> 最高在线人数：<input style='text-align:center' class='textbox' type='text' name='MaxOnline' value='" & doc.documentElement.attributes.getNamedItem("maxonline").text & "' size='5'> 最高人数发生时间：<input style='text-align:center' class='textbox' type='text' name='MaxOnlineDate' value='" & doc.documentElement.attributes.getNamedItem("maxonlinedate").text & "' size='25'> 新会员：<input style='text-align:center' class='textbox' type='text' name='NewRegUser' value='" & doc.documentElement.attributes.getNamedItem("newreguser").text & "' size='15'><br/><label><input type='checkbox' name='UpNum' value='1' onclick=""if(this.checked){alert('选中后将自动计算论坛的总数据!');}"">自动重新计算</label> "
			Set Doc=Nothing
			.echo "      <table border='0' width='98%' align'center'>"
			.echo "       <tr><td><input type='submit' class='button' value='一键重计论坛总数据'></td></tr>"
			.echo "      </table>"
			.echo "        </td></form>"
			.echo "    </tr>"
			.echo "  </table>"
			.echo " <table width=""99%"" style=""margin-top:5px"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td style='text-align:left' height=""30"" class='clefttitle'>&nbsp;<font color=""#000080""><b>一键重新统计论坛版面数据</b></font></td>"
			.echo "    </tr>"
			.echo "    <tr>"
			.echo "      <td width=""25%"" class='tdbg' style='padding:10px'><strong>功能说明：</strong>"
			.echo "       这里将重新计算论坛各个版面的帖子主题数和回复数，今日帖子，最后回复信息等，建议每隔一段时间运行一次。<br/>"
			.echo "<form action=""KS.Tools.asp?action=UpGuestBoard"" method=""post"" name=""boardform"" id=""boardform"">"
			.echo "      <table border='0' width='98%' align'center'>"
			.echo "       <tr><td><input type='submit' class='button' value='一键重计版面数据'></td></tr>"
			.echo "      </table>"
			.echo "        </td></form>"
			.echo "    </tr>"
			.echo "  </table>"
			.echo " <table width=""99%"" style=""margin-top:5px"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td style='text-align:left' height=""30"" class='clefttitle'>&nbsp;<font color=""#000080""><b>一键更新重新统计用户发帖数</b></font></td>"
			.echo "    </tr>"
			.echo "    <tr>"
			.echo "      <td width=""25%"" class='tdbg' style='padding:10px'><strong>功能说明：</strong>"
			.echo "       这里将重新计算用户的发帖数及精华帖子数，建议每隔一段时间运行一次。<br/>"
			.echo "<form action=""KS.Tools.asp?action=UpUserPostNum"" method=""post"" name=""boardform"" id=""boardform"">"
			.echo "      <table border='0' width='98%' align'center'>"
			.echo "       <tr><td><input type='submit' class='button' value='一键更新用户论坛数据'></td></tr>"
			.echo "      </table>"
			.echo "        </td></form>"
			.echo "    </tr>"
			.echo "  </table>"
			.echo "</body>"
			.echo "</html>"
		End If
	End With
End Sub
		
'自动关联文档
Sub relativeDoc()
     Call main()
     Dim KeyWords,ChannelID,InfoID,Param,TopStr,SqlStr,TotalPut
	 If KS.ChkClng(Request("docNum"))<>0 Then TopStr=" top " & KS.ChkClng(Request("docNum"))
	 If KS.ChkClng(Request("ChannelID"))<>0 Then Param=" Where ChannelID=" &KS.ChkClng(Request("ChannelID")) 
     Dim RSI:Set RSI=Server.CreateObject("ADODB.RECORDSET")
	 SqlStr="Select " & TopStr & " KeyWords,ChannelID,InfoID,title From KS_ItemInfo" & Param & " Order By Id Desc"
	 RSI.Open SqlStr,conn,1,1
	 If Not RSI.Eof Then 
        Dim TotalNum
		If KS.ChkClng(KS.S("TotalNum"))<>0 Then TotalNum=KS.S("TotalNum") Else TotalNum=RSI.Recordcount
		If TotalNum>RSI.Recordcount Then TotalNum=RSI.Recordcount
		Dim Key,NowNum,CurrNowNum:CurrNowNum=KS.ChkClng(KS.G("CurrNowNum"))
		If CurrNowNum=0 Then CurrNowNum=1
		RSI.Move(CurrNowNum-1)
		For NowNum=CurrNowNum To TotalNum
			KeyWords=RSI(0)
			ChannelID=RSI(1)
			InfoID=RSI(2)
			SqlKeyWordStr=""
			If Not KS.IsNul(KeyWords) Then 
				Dim KeyWordsArr, I, SqlKeyWordStr
				KeyWordsArr = Split(KeyWords, ",")
				 For I = 0 To UBound(KeyWordsArr)
							 If SqlKeyWordStr = "" Then
									SqlKeyWordStr = " instr(keywords,'" & KeyWordsArr(I) & "')>0 "
							 Else
									SqlKeyWordStr = SqlKeyWordStr & "or instr(keywords,'" & KeyWordsArr(I) & "')>0 "
							 End If

				Next
				
				Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
				RS.Open "Select top 30 ChannelID,InfoID,Title From KS_ItemInfo Where ChannelID=" & ChannelID & " And InfoID<>" & InfoID & " and (" & SqlKeyWordStr & ")",conn,1,1
				Do While Not RS.Eof
				  Conn.Execute("Delete From KS_ItemInfoR Where ChannelID=" & ChannelID & " and InfoID=" & InfoID & " and RelativeID=" & RS(1) & " And RelativeChannelID=" & RS(0))
				  Conn.Execute("Insert Into KS_ItemInfoR(ChannelID,InfoID,RelativeChannelID,RelativeID) values(" & ChannelID &"," & InfoID & "," & RS(0) & "," & RS(1) & ")")
				 RS.MoveNext
				Loop
				RS.Close:Set RS=Nothing
	  End If
			Call InnerJS(NowNum,TotalNum,"文档")
			if Not Response.IsClientConnected then Exit FOR
			If TotalNum>1 and NowNum Mod 100=0 Then
					rsi.close:set rsi=nothing
					if KS.ChkClng(KS.S("ChannelID"))=0 Then
					ShowPause "TotalNum=" & TotalNum & "&CurrNowNum=" & NowNum+1 & "&docNum=" & KS.S("docNum") & "&action=relativeDoc"
					Else
					ShowPause "TotalNum=" & TotalNum & "&CurrNowNum=" & NowNum+1 & "&docNum=" & KS.S("docNum") & "&ChannelID=" & ChannelID & "&action=relativeDoc"
					End If
			End If
		   RSI.MoveNext
		 Next
	 response.write "<script>sss.innerHTML='<strong>恭喜，执行完毕！<a href=""KS.Tools.asp"" style=""color:#ff6600"">点此返回</a></strong>';</script>"
	Else
		response.write "<script>sss.innerHTML='<strong>找不到记录,<a href=""KS.Tools.asp"" style=""color:#ff6600"">点此返回</a></strong>';</script>"
	End If
	  RSI.Close:Set RSI=Nothing
End Sub

'提取文档第一张图片
Sub getDocImage()
 call main()
 Dim successnum,ChannelID:ChannelID=KS.ChkClng(Request("ChannelID"))
 If ChannelID=0 Then Response.End
     if KS.ChkClng(KS.S("successnum"))>0 Then successnum=KS.ChkClng(KS.S("successnum")) Else successnum=0
     Dim Content,PhotoUrl,TopStr
	 If KS.ChkClng(Request("docNum"))<>0 Then TopStr=" top " & KS.ChkClng(Request("docNum"))
	 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
	 RS.Open "Select " & TopStr &" PhotoUrl,PicNews,ArticleContent,ID From " & KS.C_S(Channelid,2) & " Where PicNews=0 Order by Id desc",conn,1,3
	 If Not RS.Eof Then
			Dim TotalNum
			If KS.ChkClng(KS.S("TotalNum"))<>0 Then TotalNum=KS.S("TotalNum") Else TotalNum=RS.Recordcount
			Dim Matches,Key,NowNum,CurrNowNum:CurrNowNum=KS.ChkClng(KS.G("CurrNowNum"))
			If CurrNowNum=0 Then CurrNowNum=1
			RS.Move(CurrNowNum-1)
			For NowNum=CurrNowNum To TotalNum
			  Dim regEx:Set regEx = New RegExp
			  regEx.IgnoreCase = True
			  regEx.Global = True
			  regEx.Pattern = "src\=.+?\.(gif|jpg)"
			  Content=KS.HtmlCode(rs(2))
			  Set Matches = regEx.Execute(Content)
			  If regEx.Test(Content) Then
			   PhotoUrl=Lcase(Matches(0).value)
			   PhotoUrl=replace(PhotoUrl,"src=","")
			   PhotoUrl=replace(PhotoUrl,"""","")
			   PhotoUrl=replace(PhotoUrl,"'","")
			   RS(0)=PhotoUrl
			   RS(1)=1
			   RS.Update
			   Conn.Execute("Update KS_ItemInfo Set PhotoUrl='" & PhotoUrl & "' Where ChannelID=" &ChannelID & " And InfoId=" & RS(3))
			   successnum=successnum+1
			  End If
			  Call InnerJS(NowNum,TotalNum,KS.C_S(ChannelID,4))
			  if Not Response.IsClientConnected then Exit FOR
			  If TotalNum>1 and NowNum Mod 100=0 Then
				rs.close:set rs=nothing
				if KS.ChkClng(KS.S("ChannelID"))=0 Then
				ShowPause "TotalNum" & TotalNum & "&successnum=" & successnum & "&CurrNowNum=" & NowNum+1 & "&docNum=" & KS.S("docNum") & "&action=getDocImage"
				Else
				ShowPause "TotalNum" & TotalNum & "&successnum=" & successnum & "&CurrNowNum=" & NowNum+1 & "&docNum=" & KS.S("docNum") & "&ChannelID=" & ChannelID & "&action=getDocImage"
				End If			 
			  End If
			 RS.MoveNext
		Next
		response.write "<script>sss.innerHTML='<strong>恭喜，执行完毕！成功设置了 <font color=red>" & successnum & "</font> 篇文档为图片文档！<a href=""KS.Tools.asp"" style=""color:#ff6600"">点此返回</a></strong>';</script>"
	 Else
		response.write "<script>sss.innerHTML='<strong>找不到记录,<a href=""KS.Tools.asp"" style=""color:#ff6600"">点此返回</a></strong>';</script>"
	End If
 RS.Close : Set RS=Nothing
End Sub

'删除无内容的记录
Function DocContent()
 Dim successnum,Param,SqlStr,TotalPut,NowNum,TotalNum
 ChannelID=KS.ChkClng(KS.S("ChannelID"))
 If ChannelID=0 Then KS.AlertHintScript "请选择模型!"
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
  Param=" where articlecontent='' or articlecontent is null"
  TotalNum=Conn.Execute("Select count(id) From " & KS.C_S(ChannelID,2) & " " & Param)(0)
  Conn.Execute("delete From KS_ItemInfo Where ChannelID=" &ChannelID & " And InfoID in(select id From " & KS.C_S(ChannelID,2) & " " & Param & ")")
  Conn.Execute("delete From " & KS.C_S(ChannelID,2) & " " & Param)
  
  response.write "<script>alert('恭喜，执行完毕,成功删除 " & TotalNum & " 条记录！');location.href='KS.Tools.asp';</script>"
End Function

'检测文档文件名
Function checkDocFname()
 call main()
 Dim successnum,Param,SqlStr,TopStr,TotalPut,NowNum
 If KS.ChkClng(Request("docnum"))<>0 Then TopStr=" top " & KS.ChkClng(Request("docnum"))
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 If DataBaseType=1 Then
  Param=" where fname is null or charindex('.',fname)=0"
 Else
  Param=" where fname is null or instr(fname,'.')=0"
 End If
 If ChannelID=0 Then
  SqlStr="Select" & TopStr & " Fname,InfoID,ChannelID From KS_ItemInfo " & Param & " Order By ID Desc"
 Else
  SqlStr="Select" & TopStr & " Fname,ID From " & KS.C_S(ChannelID,2) & " " & Param & " Order By ID Desc"
 End If
 RS.Open SqlStr,conn,1,3
 If Not (RS.Eof or rs.bof) Then
		Dim TotalNum:TotalNum=RS.Recordcount
		For NowNum=1 To TotalNum
			RS(0)=RS(1) & ".html"
			RS.Update
			If ChannelID=0 Then
			 Conn.Execute("Update " & KS.C_S(RS(2),2) & " Set Fname='" & RS(0) & "' Where ID=" & RS(1))
			Else
			 Conn.Execute("Update KS_ItemInfo Set Fname='" & RS(0) & "' Where ChannelID=" & ChannelID & " And InfoID=" & RS(1))
			End If
			Call InnerJS(NowNum,TotalNum,"条")
			RS.Movenext
	    Next
		response.write "<script>sss.innerHTML='<strong>恭喜，执行完毕,成功修复<font color=red>" & TotalNum & "</font>条记录！<a href=""KS.Tools.asp"" style=""color:#ff6600"">点此返回</a></strong>';</script>"
 Else
	response.write "<script>sss.innerHTML='<strong>找不到记录,<a href=""KS.Tools.asp"" style=""color:#ff6600"">点此返回</a></strong>';</script>"
 End If
 RS.Close
 Set RS=Nothing
End Function
		
'执行关键词获取
Sub setKeyWords()
   on error resume next
			 Dim ChannelID:ChannelID=KS.ChkClng(Request("ChannelID"))
			 If ChannelID=0 Then KS.AlertHintScript "请选择模型!"
			 Call main()
			 Dim Param,SqlStr,TopStr,TotalPut
			 If KS.ChkClng(Request("docNum"))<>0 Then TopStr=" top " & KS.ChkClng(Request("docNum"))
			 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			  If KS.S("setall")<>"1" Then
			  SqlStr="Select" & TopStr & " Title,KeyWords,ID From " & KS.C_S(ChannelID,2) & " where keywords='' or keywords is null Order By ID Desc"
			  Else
			  SqlStr="Select" & TopStr & " Title,KeyWords,ID From " & KS.C_S(ChannelID,2) & " Order By ID Desc"
			  End If
			 RS.Open SqlStr,conn,1,1
			 If Not (RS.Eof or rs.bof) Then
					Dim WS:Set WS=New Wordsegment_Cls
					Dim TotalNum
					If KS.ChkClng(KS.S("TotalNum"))<>0 Then TotalNum=KS.S("TotalNum") Else TotalNum=RS.Recordcount
					Dim Key,NowNum,CurrNowNum:CurrNowNum=KS.ChkClng(KS.G("CurrNowNum"))
					If CurrNowNum=0 Then CurrNowNum=1
					RS.Move(CurrNowNum-1)
					For NowNum=CurrNowNum To TotalNum
						Key=WS.SplitKey(RS(0),4,20)
						Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set KeyWords='" & Key & "' Where ID=" & rs(2))
						Conn.Execute("Update KS_ItemInfo Set KeyWords='" & Key & "' Where ChannelID=" & ChannelID & " And InfoID=" & RS(2))
						Call InnerJS(NowNum,TotalNum,KS.C_S(ChannelID,4))
						'if Not Response.IsClientConnected then Exit FOR
						'If TotalNum>1 and NowNum Mod 100=0 Then  '不能暂停，会出错
						'     rs.close:set rs=nothing
						'	 ShowPause "TotalNum=" & TotalNum &"&CurrNowNum=" & NowNum+1 & "&docNum=" & KS.S("docNum") & "&ChannelID=" & ChannelID & "&action=setKeyWords&setall=" & ks.s("setall")
							 
						'End If
					 RS.MoveNext
					Next
					Set WS=Nothing
					response.write "<script>sss.innerHTML='<strong>恭喜，执行完毕！<a href=""KS.Tools.asp"" style=""color:#ff6600"">点此返回</a></strong>';</script>"
			 Else
					response.write "<script>sss.innerHTML='<strong>找不到记录,<a href=""KS.Tools.asp"" style=""color:#ff6600"">点此返回</a></strong>';</script>"
			 End If
			 RS.Close
			 Set RS=Nothing
End Sub

'重计论坛总数据
Sub UpClubData()
 Dim TopicNum:TopicNum=KS.ChkClng(KS.G("TopicNum"))
 Dim PostNum:PostNum=KS.ChkClng(KS.G("PostNum"))
 Dim TodayNum:TodayNum=KS.ChkClng(KS.G("TodayNum"))
 Dim YesterDayNum:YesterDayNum=KS.ChkClng(KS.G("YesterDayNum"))
 Dim TotalUserNum:TotalUserNum=KS.ChkClng(KS.G("TotalUserNum"))
 Dim MaxDayNum:MaxDayNum=KS.ChkClng(KS.G("MaxDayNum"))
 Dim MaxOnline:MaxOnline=KS.ChkClng(KS.G("MaxOnline"))
 Dim MaxOnlineDate:MaxOnlineDate=KS.G("MaxOnlineDate")
 Dim NewRegUser:NewRegUser=KS.G("NewRegUser")
 If Not IsDate(MaxOnlineDate) Then KS.AlertHIntScript "最高在线日期不对!"
 If KS.ChkClng(KS.G("UpNum"))=1 Then
  	Dim TableXML,Node
	set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	TableXML.async = false
	TableXML.setProperty "ServerHTTPRequest", true 
	TableXML.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
	PostNum=0:TodayNum=0:YesterDayNum=0
	TopicNum=Conn.Execute("Select Count(1) From KS_GuestBook")(0)
	TotalUserNum=Conn.Execute("Select Count(1) From KS_User")(0)
	NewRegUser=Conn.Execute("Select top 1 UserName From KS_User Order by UserID Desc")(0)
    For Each Node In TableXML.DocumentElement.SelectNodes("item")
			PostNum=PostNum+Conn.Execute("Select count(1) From " & Node.SelectSingleNode("tablename").text)(0)
			TodayNum=TodayNum+Conn.Execute("Select count(1) From " & Node.SelectSingleNode("tablename").text & " Where year(ReplayTime)=" & Year(Now) & " And month(ReplayTime)=" & Month(Now) & " And Day(ReplayTime)=" & Day(Now))(0)
			YesterDayNum=YesterDayNum+Conn.Execute("Select count(1) From " & Node.SelectSingleNode("tablename").text & " Where datediff(" & DataPart_D & ",ReplayTime," & SQLNowString & ")=1")(0)
	Next
 End If
 
 Dim doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	Doc.async = false
	Doc.setProperty "ServerHTTPRequest", true 
	Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
	  doc.documentElement.attributes.getNamedItem("topicnum").text=TopicNum
	  doc.documentElement.attributes.getNamedItem("postnum").text=PostNum
	  doc.documentElement.attributes.getNamedItem("todaynum").text=TodayNum
	  doc.documentElement.attributes.getNamedItem("yesterdaynum").text=YesterDayNum
	  doc.documentElement.attributes.getNamedItem("totalusernum").text=TotalUserNum
	  doc.documentElement.attributes.getNamedItem("maxdaynum").text=MaxDayNum
	  doc.documentElement.attributes.getNamedItem("maxonline").text=MaxOnline
	  doc.documentElement.attributes.getNamedItem("maxonlinedate").text=MaxOnlineDate
	  doc.documentElement.attributes.getNamedItem("newreguser").text=NewRegUser
   doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
  KS.AlertHIntScript "恭喜，论坛总数据重新成功!"
End Sub
		
'统计版面数据
Sub UpGuestBoard()
    Call main()
	Dim TableXML,Node,NowNum,TotalNum,TaskUrl,Taskid,Action
	Dim TopicNum,PostNum,TodayNum,N_LastPost,RST
	set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	TableXML.async = false
	TableXML.setProperty "ServerHTTPRequest", true 
	TableXML.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
	NowNum=0
	Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "Select id From KS_GuestBoard Where ParentID<>0 Order By OrderID,id",conn,1,1
	If Not (RS.Eof or rs.bof) Then
	   TotalNum=RS.Recordcount
	  Do While Not RS.Eof
	     NowNum=NowNum+1
		 TopicNum=Conn.Execute("Select count(id) From KS_GuestBook Where BoardID=" & RS(0))(0)
		 PostNum=0 : TodayNum=0
		 For Each Node In TableXML.DocumentElement.SelectNodes("item")
			PostNum=PostNum+Conn.Execute("Select count(1) From KS_GuestBook A INNER JOIN " & Node.SelectSingleNode("tablename").text & " B ON A.ID=B.TopicID Where A.BoardID=" & RS(0))(0)
			TodayNum=TodayNum+Conn.Execute("Select count(1) From KS_GuestBook A INNER JOIN " & Node.SelectSingleNode("tablename").text & " B ON A.ID=B.TopicID Where A.BoardID=" & RS(0) & " And year(B.ReplayTime)=" & Year(Now) & " And month(B.ReplayTime)=" & Month(Now) & " And Day(B.ReplayTime)=" & Day(Now))(0)
		 Next
		 Set RST=Conn.Execute("select top 1 * From KS_GuestBook Where BoardID=" & rs(0) &" order by id desc")
		 If Not RST.Eof Then
			 N_LastPost=RST("id")&"$"& now & "$" & Replace(left(RST("subject"),200),"$","") & "$" & RST("LastReplayUser") & "$" &RST("LastReplayUserID")&"$$"
		 Else
			 N_LastPost="0$"& now & "$无$无$0$$"
		 End If
		 RST.Close:Set RST=Nothing
         Conn.Execute("Update KS_GuestBoard Set PostNum=" & PostNum & ",TopicNum=" & TopicNum & ",TodayNum=" & TodayNum & " Where ID= "& RS(0))
	     Call InnerJS(NowNum,TotalNum,KS.C_S(ChannelID,4))
	  RS.MoveNext
	  Loop
	  	response.write "<script>sss.innerHTML='<strong>恭喜，执行完毕！<a href=""KS.Tools.asp"" style=""color:#ff6600"">点此返回</a></strong>';</script>"

	Else
		response.write "<script>sss.innerHTML='<strong>您还没有建版面,<a href=""KS.Tools.asp"" style=""color:#ff6600"">点此返回</a></strong>';</script>"
	 End If
	 RS.Close
	 Set RS=Nothing
	Application(KS.SiteSN&"_ClubBoard")=Empty
	
End Sub

'重计用户发帖数
Sub UpUserPostNum()
    Call main()
	Dim TableXML,Node,TotalNum,BestTopicNum,TaskUrl,Taskid,Action
	Dim TopicNum,PostNum,TodayNum,N_LastPost,RST
	set TableXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	TableXML.async = false
	TableXML.setProperty "ServerHTTPRequest", true 
	TableXML.load(Server.MapPath(KS.Setting(3)&"Config/clubtable.xml"))
	NowNum=0
	Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "Select userid From KS_User Order By UserID",conn,1,1
	If Not (RS.Eof or rs.bof) Then
		If KS.ChkClng(KS.S("TotalNum"))<>0 Then TotalNum=KS.S("TotalNum") Else TotalNum=RS.Recordcount
		If TotalNum>RS.Recordcount Then TotalNum=RS.Recordcount
		Dim Key,NowNum,CurrNowNum:CurrNowNum=KS.ChkClng(KS.G("CurrNowNum"))
		If CurrNowNum=0 Then CurrNowNum=1
		RS.Move(CurrNowNum-1)
		For NowNum=CurrNowNum To TotalNum
	     NowNum=NowNum+1
		 BestTopicNum=Conn.Execute("Select count(id) From KS_GuestBook Where UserID=" & RS(0) & " And isbest=1")(0)
		 PostNum=0 
		 For Each Node In TableXML.DocumentElement.SelectNodes("item")
			PostNum=PostNum+Conn.Execute("Select count(1) From " & Node.SelectSingleNode("tablename").text & " Where UserID=" & RS(0))(0)
		 Next
		 
         Conn.Execute("Update KS_User Set PostNum=" & PostNum & ",BestTopicNum=" & BestTopicNum & " Where UserID= "& RS(0))
	     Call InnerJS(NowNum,TotalNum,"条")
		 if Not Response.IsClientConnected then Exit FOR
			  If TotalNum>1 and NowNum Mod 100=0 Then
				rs.close:set rs=nothing
				ShowPause "TotalNum=" & TotalNum & "&CurrNowNum=" & NowNum+1 & "&action=UpUserPostNum"
		 End If
	  RS.MoveNext
	  Next
	  
	  	response.write "<script>sss.innerHTML='<strong>恭喜，执行完毕！<a href=""KS.Tools.asp"" style=""color:#ff6600"">点此返回</a></strong>';</script>"

	Else
		response.write "<script>sss.innerHTML='<strong>没有找到合适的会员,<a href=""KS.Tools.asp"" style=""color:#ff6600"">点此返回</a></strong>';</script>"
	 End If
	 RS.Close
	 Set RS=Nothing
End Sub	
		
Sub Main()
		  With KS
		  .echo ("<html>")
		  .echo ("<head>")
		  .echo ("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">")
		  .echo ("<title>系统信息</title>")
		  .echo ("<script src='../ks_inc/jquery.js'></script>")
		  .echo ("</head>")
		  .echo ("<link rel=""stylesheet"" href=""include/Admin_Style.css"">")
		  '.echo ("<body oncontextmenu=""return false;"" scroll=no style='background-color:transparent'>")
		  		.echo "      <div class='topdashed sort'> 一键相关管理工具</div>"
				.echo "<br><br><br><table id=""BarShowArea"" width=""400"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1"">"
				.echo "<tr> "
				.echo "<td bgcolor=000000>"
				.echo " <table width=""400"" border=""0"" cellspacing=""0"" cellpadding=""1"">"
				.echo "<tr> "
				.echo "<td bgcolor=ffffff height=9><img src=""images/114_r2_c2.jpg"" width=100% height=10 id=img2 name=img2 align=absmiddle></td></tr></table>"
				.echo "</td></tr></table>"
				.echo "<table width=""550"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1""><tr> "
				.echo "<td align=center> 执行进度:<span id=txt2 name=txt2 style=""font-size:9pt"">100</span><span id=txt4 style=""font-size:9pt"">%</span></td></tr> "
				.echo "<tr><td align=center id='sss'>总共需要执行 <font color=red><b id=t1>0</b></font> <span id=t0></span>,<font color=red><b>在此过程中请勿刷新此页面！！！</b></font> 系统正在执行第 <font color=red><b id=t2>0</b></font> <span id=t00></span></td></tr>"
				.echo "</table>"
			

			 .echo ("</div>")
		 

		 .echo ("<table width=""100%""   border=""0"" cellpadding=""0"" cellspacing=""0"">")
		 .echo (" <tr>")
		 .echo ("   <td height=""50"" id=""fsohtml"">")
		' .echo (FsoHtmlList)
		 .echo ("      </td>")
		 .echo ("   </tr>")
		 .echo ("</table>")
		 .echo ("</body>")
		 .echo ("</html>")
		 End With
		End Sub
		
		Sub InnerJS(NowNum,TotalNum,itemname)
		  With KS
				.echo "<script>var p=" & FormatNumber(NowNum / TotalNum * 100, 2, -1) & ";if (p>100) p=100;"
				'.echo "fsohtml.innerHTML='" & FsoHtmlList & "';" & vbCrLf
				.echo "img2.width=" & Fix((NowNum / TotalNum) * 400) & ";txt2.innerHTML=p;t1.innerHTML=" & TotalNum &";t2.innerHTML=" & NowNum & ";t0.innerHTML=t00.innerHTML='" & ItemName &"';img2.title=""(" & NowNum & ")"";" & vbCrLf
				'.echo "txt3.innerHTML=""总共需要执行 <font color=red><b>" & TotalNum & "</b></font> " & itemname & ",<font color=red><b>在此过程中请勿刷新此页面！！！</b></font> 系统正在执行第 <font color=red><b>" & NowNum & "</b></font> " & itemname & """;" & vbCrLf
				.echo "</script>" & vbCrLf
				Response.Flush
		  End With
		End Sub
		
		Sub ShowPause(param)
		    ks.echo "<script>"
			ks.echo "fsohtml.innerHTML='<div style=""text-align:cdenter""><div style=""margin:10px;height:80px;padding:8px;border:1px dashed #cccccc;text-align:left;""><img src=""images/succeed.gif"" align=""left""><br>&nbsp;&nbsp;&nbsp;&nbsp;<b>温馨提示：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;以免过度占用服务器资源，系统暂停2秒后继续<img src=""../images/default/wait.gif""><br>&nbsp;&nbsp;&nbsp;&nbsp;如果2秒后没有继续，请点此<a href=""KS.Tools.asp?" & param &"""><font color=red>继续</font></a>或点此<a href=""KS.Tools.asp""><font color=red>停止</font></a>!</div></div>';" & vbCrLf
			ks.echo "</script>" &vbcrlf
			ks.die "<meta http-equiv=""refresh"" content=""2;url=KS.Tools.asp?" & Param & """>"
			'response.flush
		End Sub

End Class
%> 
