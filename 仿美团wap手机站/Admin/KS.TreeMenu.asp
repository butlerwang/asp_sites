<%@language=vbscript CODEPAGE="65001" %>
<%
Option Explicit
Response.buffer = True
Server.ScriptTimeout=9999999
%>
<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KS:Set KS=New PublicCls
Dim strInstallDir,ComeUrl
If Not KS.ReturnPowerResult(0, "KMSL10009") Then          
	'Call KS.ReturnErr(1, "")
	Response.End
End If

ComeUrl=Request.ServerVariables("http_referer")
strInstallDir=KS.Setting(3)

Dim ChannelUrl, UseCreateHTML,  ListFileType, FileExt_List

Dim hf, strTopMenu, pNum, pNum2, OpenTyKS_Class, strMenuJS
Dim ObjInstalled, FSO
ObjInstalled = KS.IsObjInstalled(KS.Setting(99))
If ObjInstalled = True Then
    Set FSO = KS.InitialObject(KS.Setting(99))
End If
Response.Write "<html><head><title>顶部栏目菜单管理</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf
Response.Write "<link href='include/Admin_Style.css' rel='stylesheet' type='text/css'></head>" & vbCrLf
Response.Write "<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
Response.Write "<div class='topdashed' style='text-align:center'>"
Response.Write "<b>生成树型菜单</b>"		
Response.Write "</div>"


Dim Action:Action=KS.G("Action")
If Action = "Create" Then
    Call Create_RootClass_Menu
Else
    Call Create_Tree()
End If
Response.Write "</body></html>" & vbCrLf

Sub Create_Tree()
    Response.Write "<form method='POST' action='?Action=Create' id='myform' name='myform'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"
	Response.Write "  <tr  class='title'>"
    Response.Write "    <td height='22' colspan='6'><strong>树型菜单参数设置</strong> </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>选择频道：</strong></td>"
    Response.Write "    <td>"
    Response.Write ReturnAllChannel()
    Response.Write "    </td>"
	Response.Write " </tr>"
	Response.Write " <tr class='tdbg'>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>生成样式：</strong></td>"
    Response.Write "    <td>"
    Response.Write "      <select name='fsostyle' onchange=""if (this.value==2) document.all.cols.style.display='';else document.all.cols.style.display='none';"">"
	Response.Write "        <option value=1>样式一</option>"
	Response.Write "        <option value=2>样式二</option>"
	Response.Write "      </select>"
    Response.Write "    </td>"
	Response.Write "</tr>"
	Response.Write "<tbody id='cols' style='display:none'>"
	Response.Write " <tr class='tdbg'>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>生成列数：</strong></td>"
    Response.Write "    <td>"
    Response.Write "      <input type='text' name='col' value='2' size=""6"">列"
    Response.Write "    </td>"
	Response.Write "</tr>"
	Response.Write "</tbody>"
	Response.Write "<tr class='tdbg'>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>生成文件名：</strong></td>"
    Response.Write "    <td>"
    Response.Write "      <input name='JsFileName' type='text' id='JsFileName' value='Tree.js' size='10' maxlength='10'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
	Response.Write "</table>"
    Response.Write "<br><div style='text-align:center'><input type='submit' name='Submit' value=' 生成树型导航 ' class='button'></div>"
	Response.Write "</form>"
End Sub



Sub Create_RootClass_Menu()
    If KS.ChkCLng(KS.S("fsostyle"))=1 Then
    strTopMenu = TreeList 
	Else
	strTopMenu = HtreeList
	End If
    Call KS.WriteTOFile(KS.Setting(3) & KS.Setting(93) & KS.G("JsFileName"), strTopMenu)
	Response.Write "<br><table width='100%' border='0' cellspacing='1' cellpadding='2' class='ctable'>"
    Response.Write "  <tr class='sort'>"
    Response.Write "    <td height='22' align='center'><strong> 生 成 树 形 导 航 菜 单 </strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='150'>"
	Response.Write "<br><p align='center'><font color=red><b>恭喜您！树形导航菜单成功生成,请按以下提示完成最好操作。</b></font></p>"
    Response.Write "<p><b>将以下代码复制到在模板里要显示的地方。</b></p>"
	Response.Write "<input name='s2' value='&lt;script language=&quot;javascript&quot; type=&quot;text/javascript&quot; src=&quot;" & KS.Setting(3) & KS.Setting(93) & KS.G("JsFileName") & "&quot;&gt;&lt;/script&gt;' size='80'>&nbsp;<input class=""button"" onClick=""jm_cc('s2')"" type=""button"" value=""复制到剪贴板"" name=""button1"">"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
 %>
 <script>
function jm_cc(ob)
{
	var obj=MM_findObj(ob); 
	if (obj) 
	{
		obj.select();js=obj.createTextRange();js.execCommand("Copy");}
		alert('复制成功，粘贴到你要调用的模板里即可!');
	}
	function MM_findObj(n, d) { //v4.0
  var p,i,x;
  if(!d) d=document;
  if((p=n.indexOf("?"))>0&&parent.frames.length)
   {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);
   }
  if(!(x=d[n])&&d.all) x=d.all[n];
  for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}
  </script>
 <%
End Sub


    
Function EncodeJS(str)
    EncodeJS = Replace(Replace(Replace(Replace(Replace(str, Chr(10), ""), "\", "\\"), "'", "\'"), vbCrLf, "\n"), Chr(13), "")
End Function


'取得网站的所有频道及其子栏目
Function ReturnAllChannel()
	  Dim RS:Set RS=KS.InitialObject("ADODB.Recordset")
	  Dim SQL,K,ChannelStr:ChannelStr = ""
	   ChannelStr = "<select class='textbox' name=""ChannelID"" style=""width:200;border-style: solid; border-width: 1""><option value='0'>---不指定栏目---</option>"
	   RS.Open "Select channelid,channelname From [KS_Channel] Where ChannelStatus=1", Conn, 1, 1
	   If RS.EOF And RS.BOF Then
		  RS.Close:Set RS = Nothing:Exit Function
	   Else
	     SQL=RS.GetRows(-1):rs.close:set rs=nothing
	   End iF
		
	    For K=0 To ubound(sql,2)
		   ChannelStr = ChannelStr & "<option value=" & sql(0,k) & ">" & sql(1,k) & "</option>"
		Next 
		ChannelStr = ChannelStr & "<optgroup  label=""-----指定到具体的栏目(以下列出了整站的导航树)----"">"  
	   For K=0 To Ubound(sql,2)
	        ChannelStr=ChannelStr & KS.LoadClassOption(sql(0,k),false)
	    Next
	   ReturnAllChannel = ChannelStr &"</select>"
End Function
	
	Function TreeList()
				Dim RS,TreeStr,ID,i,Param,ChannelID
				ChannelID=KS.S("ChannelID")
				 If Len(Channelid)>4 Then
					 Param=" and a.tn='" & ChannelID & "'"
				 Else
				   If ChannelID<>"0" Then  Param=" And tj=1 and B.ChannelID=" & KS.ChkCLng(KS.S("ChannelID"))
				 End If
				TreeStr="document.writeln('<img src=""" & KS.Setting(3) & "images/tree/home.gif"" align=""absmiddle""><a href=""/"">网站首页</a>');" & vbcrlf
				Set  RS=KS.InitialObject("ADODB.Recordset")
				RS.Open ("select ID,FolderName from KS_Class A,KS_Channel B Where A.ChannelID=B.ChannelID And B.ChannelStatus=1 "  & Param & " Order BY root,folderorder"), Conn, 1, 1
				Do While Not RS.EOF
				  i=i+1:ID = Trim(RS(0))
				  if i=rs.recordcount then
				   if ks.chkclng(ks.g("channelid"))=0 then
				     TreeStr = TreeStr  & "document.writeln('<div><img src=""" & KS.Setting(3) & "images/tree/m2.gif"" align=""absmiddle""><img src=""" & KS.Setting(3) & "images/tree/folderopen.gif"" align=""absmiddle""> " & KS.GetClassNP(rs(0))& "</div>');" & vbnewline
				   elseif conn.execute("select id from ks_class where tn='" & rs(0) & "'").eof then
				     TreeStr = TreeStr  & "document.writeln('<div><img src=""" & KS.Setting(3) & "images/tree/m2.gif"" align=""absmiddle""><img src=""" & KS.Setting(3) & "images/tree/folderopen.gif"" align=""absmiddle""> " & KS.GetClassNP(rs(0))& "</div>');" & vbnewline
				   else
				     TreeStr = TreeStr  & "document.writeln('<div><img src=""" & KS.Setting(3) & "images/tree/m1.gif"" align=""absmiddle""><img src=""" & KS.Setting(3) & "images/tree/folderopen.gif"" align=""absmiddle""> " & KS.GetClassNP(rs(0))& "</div>');" & vbnewline
				   end if
				  TreeStr = TreeStr & ReturnSubList(ID,1)
				  else
				  TreeStr = TreeStr  & "document.writeln('<div><img src=""" & KS.Setting(3) & "images/tree/m1.gif"" align=""absmiddle""><img src=""" & KS.Setting(3) & "images/tree/folderopen.gif"" align=""absmiddle""> " & KS.GetClassNP(rs(0))& "</div>');" & vbnewline
				  TreeStr = TreeStr & ReturnSubList(ID,0)
				  end if
				  
				RS.MoveNext
			   Loop
			   RS.Close:Set RS = Nothing
			 TreeList=TreeStr
	End Function
	
	Public Function ReturnSubList(ParentID,flag)
	  Dim SubTypeList, RS, SpaceStr, k, Total,ID,TJ,I
	  Set RS = conn.execute("Select ID,FolderName,TJ,Child from KS_Class Where TN='" & ParentID & "' Order BY root,folderorder")
	  Dim SQL
	  If Not RS.Eof Then
	  SQL=RS.GetRows(-1):RS.Close:Set RS = Nothing
	  dim num:num=0
	  For I=0 To Ubound(SQL,2)
	   SpaceStr = "document.writeln('<div>"
		TJ = CInt(SQL(2,I))
		Num = Num + 1

		If KS.ChkClng(KS.S("ChannelID"))<>0 Then
			For k = 1 To TJ - 1
			  If k = 1 And k <> TJ - 1 Then
			  SpaceStr = SpaceStr & "<img src=""" & KS.Setting(3) & "images/tree/l4.gif"" align=""absmiddle"">"
			  ElseIf k = TJ - 1 Then
				If Num = Ubound(SQL,2)+1 Then
					 SpaceStr = SpaceStr & "<img src=""" & KS.Setting(3) & "images/tree/l2.gif"" align=""absmiddle"">"
				Else
					 SpaceStr = SpaceStr & "<img src=""" & KS.Setting(3) & "images/tree/l1.gif"" align=""absmiddle"">"
				End If
			  Else
			   SpaceStr = SpaceStr & "<img src=""" & KS.Setting(3) & "images/tree/l4.gif"" align=""absmiddle"">"
			  End If
			Next
			If SQL(3,I)<>0 Then
	     SubTypeList = SubTypeList &SpaceStr & "<img src=""" & KS.Setting(3) & "images/tree/folderopen.gif"" align=""absmiddle""> " & KS.GetClassNP(SQL(0,I)) 
		    Else
	     SubTypeList = SubTypeList &SpaceStr & "<img src=""" & KS.Setting(3) & "images/tree/file.gif"" align=""absmiddle""> " & KS.GetClassNP(SQL(0,I)) 
			End If
			
	   Else
		
			For k = 1 To TJ - 1
			 if flag=1 then
			  SpaceStr = SpaceStr & "&nbsp;&nbsp;&nbsp;"
			 else
			  SpaceStr = SpaceStr & "<img src=""" & KS.Setting(3) & "images/tree/l4.gif"" align=""absmiddle"">"
			  end if
			Next
			
			If I=Ubound(SQL,2) Then
			  if conn.execute("select id from ks_class where tn='" & sql(0,i) & "'").eof  then
			   SubTypeList = SubTypeList & SpaceStr & "<img src=""" & KS.Setting(3) & "images/tree/L2.gif"" align=""absmiddle"">"
			  else
			   SubTypeList = SubTypeList & SpaceStr & "<img src=""" & KS.Setting(3) & "images/tree/m1.gif"" align=""absmiddle"">"
			  end if
			Else
			  SubTypeList = SubTypeList & SpaceStr & "<img src=""" & KS.Setting(3) & "images/tree/L1.gif"" align=""absmiddle"">"
			End If
			If SQL(3,I)<>0 Then
	       SubTypeList = SubTypeList &"<img src=""" & KS.Setting(3) & "images/tree/folderopen.gif"" align=""absmiddle""> " & KS.GetClassNP(SQL(0,I)) 
			Else
	       SubTypeList = SubTypeList &"<img src=""" & KS.Setting(3) & "images/tree/file.gif"" align=""absmiddle""> " & KS.GetClassNP(SQL(0,I)) 
		    End If
			
	   End If
		
	   SubTypeList = SubTypeList &"</div>');"& vbnewline
	   SubTypeList = SubTypeList & ReturnSubList(SQL(0,I),flag)
	  Next
	 End If
	  ReturnSubList = SubTypeList
	End Function	
	
	'横向
	Function HtreeList()
	   Dim RS,TreeStr,ID,i,Param,ChannelID
	   ChannelID=KS.S("ChannelID")
	   If Len(Channelid)>4 Then
	     Param=" and a.tn='" & ChannelID & "'"
	   Else
				If KS.S("ChannelID")<>"0" Then  Param="  and B.ChannelID=" & KS.ChkCLng(ChannelID)
				IF KS.S("ChannelID")<>"8" Then Param=Param &"  And tj=1" Else Param=Param & " and tj=2"
	   End If
				Set  RS=KS.InitialObject("ADODB.Recordset")
				RS.Open ("select ID from KS_Class A,KS_Channel B Where A.ChannelID=B.ChannelID And B.ChannelStatus=1 "  & Param & " Order BY root,folderorder"), Conn, 1, 1
				TreeStr=TreeStr & "document.writeln('<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""3"">');" & vbcrlf
				Do While Not RS.EOF
				  TreeStr=TreeStr & "document.writeln('<tr>');" & vbcrlf
				  For I=1 To KS.ChkClng(KS.G("Col"))
				   TreeStr = TreeStr & "document.writeln('<td valign=""top"" width=""" & 100 / KS.ChkCLng(KS.G("Col")) & "%"">');" & vbcrlf
				   TreeStr = TreeStr & "document.writeln('<div class=""classtitle"" style=""font-weight:bold""><img src=""" & KS.Setting(3) & "images/default/arrow_r.gif"" align=""absmiddle"">&nbsp;" & KS.GetClassNP(rs(0))& "</div>');" & vbnewline 
				   TreeStr = TreeStr & SubList(RS(0))
				   TreeStr = TreeStr & "document.writeln('</td>');" & vbcrlf
				   RS.MoveNext
				   If RS.EOF Then Exit For
				  Next
				   TreeStr = TreeStr & "document.writeln('</tr>');" & vbcrlf
				  if rs.eof then exit do
				Loop
				TreeStr =TreeStr & "document.writeln('</table>');" & vbcrlf
				RS.Close:Set RS=Nothing
		HtreeList=TreeStr
	End Function	
	
	Function SubList(ParentID)
	  Dim RS:Set RS=Conn.Execute("select id from ks_class where tn='" & ParentID & "' order by root,folderorder")
	  Dim SQL,I
	  If Not RS.Eof Then
	     SQL=RS.GetRows(-1)
		 SubList="document.writeln('<div class=""list"">"
		 For I=0 To Ubound(SQL,2)
		   SubList=SubList & KS.GetClassNP(SQL(0,I)) & "&nbsp;"
		   If I <> Ubound(SQL,2) Then SubList=SubList & "<img src=""" & KS.Setting(3) & "images/nl.gif"" align=""absmiddle"">&nbsp;"
		   
		 Next
		 SubList=SubList & "</div>');"& vbcrlf
	  End IF
	  RS.Close:Set RS=Nothing
	End Function 
%>
