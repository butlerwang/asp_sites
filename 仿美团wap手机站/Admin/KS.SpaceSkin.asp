<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdminiStratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New BlogUserSkin
KSCls.Kesion()
Set KSCls = Nothing

Class BlogUserSkin
        Private KS,flag,KSCls
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
		dim action:Action=trim(request("Action"))
		flag=KS.chkclng(KS.g("flag"))

		 With Response
			  .Write "<html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			  .Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../ks_inc/Common.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../ks_inc/kesion.box.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../ks_inc/jquery.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""location.href='KS.Space.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>空间管理</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.SpaceSkin.asp?flag=" & flag & "';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>模板管理</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.Space.asp?action=class';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>空间分类</span></li>"
			  .Write "</ul>"
		End With
		
		select case Action
		    case "newtemplate","modifytext"
			    call textTemplate()
			case "saveaddtext"
			    call saveaddtext()
			case "savetext"
			    call savetext()	
			case "savedefault"
				call savedefault()
			case "delconfig"
				call delconfig()
			case else
				call showconfig()
		end select
	 End Sub

sub showconfig()
dim rs:set rs=conn.execute("select * From KS_BlogTemplate where flag=" & flag & " order by usertag")
%>
<script type="text/javascript">
$(document).ready(function(){
 $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
 $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
})
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"0>
<form name="form2" method="post" action="KS.SpaceSkin.asp?action=savedefault&flag=<%=KS.g("flag")%>">
  <tr class="sort"> 
      <td width="6%"><div align="center">ID</div></td>
      <td width="20%" ><div align="center">名称</div></td>
      <td width="15%" ><div align="center">作者</div></td>
      <td width="12%" ><div align="center">默认模版</div></td>
      <td width="12%" ><div align="center">类型</div></td>
      <td width="47%" > 
        <div align="center">模版管理</div></td>
  </tr>
      <% 
while not rs.eof	  
%> 
    <tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
            <td class="splittd" align="center"><%= rs("id") %>&nbsp;</td>
            <td class="splittd" align="center"><%= rs("TemplateName") %></td>
            <td class="splittd" align="center"><%= rs("TemplateAuthor") %></td>
            <td class="splittd" align="center"> <input name="radiobutton" type="radio" value='<%=rs("id")%>' <%if rs("isdefault")="true" then response.Write "checked" %>></td>
			<td class="splittd" align="center">
			 <% if rs("usertag")=1 then
			     response.write "<font color=blue>用户上传</font>"
				else
				 response.write "<font color=red>系统自带</font>"
			 end if%>
			</td>
            
      <td class="splittd" width="40%"> <div align="center"><a href="../space/showtemplate.asp?templateid=<%=rs("id")%>" target="_blank">预览</a>
	  　<a onclick="$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('空间门户 >> <font color=red>修改模板</font>')+'&ButtonSymbol=GOSave';location.href='KS.SpaceSkin.asp?action=modifytext&id=<%=rs("id")%>&flag=<%=KS.g("flag")%>'" href="#">修改模版</a>　<a href="KS.SpaceSkin.asp?action=delconfig&id=<%=rs("id")%>&flag=<%=KS.g("flag")%>" onclick=return(confirm("确定要删除这个模版吗？"))>删除模版</a></div>
	  </td>
    </tr>
      <%
rs.movenext
wend
%>
    <td height="40" colspan="5" align="center">  
        <div align="center">
          <input type="submit" name="Submit" class="button" value="保存设置">&nbsp;&nbsp;
		  <input type="button" name="Submit1" class="button" value="添加新模板" onClick="$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('空间门户 >> <font color=red>添加模板</font>')+'&ButtonSymbol=GO';location.href='?Action=newtemplate&flag=<%=KS.g("flag")%>';">
      </div></td>
  </tr>
</form>  
</table>
<%
	set rs=nothing
end sub
'添加新模板
sub TextTemplate()
   Dim CurrPath
   CurrPath=KS.GetCommonUpFilesDir()
  dim templatename,templateauthor,templatemain,templatesub,Action,templatepic,GroupID
 if KS.g("action")="modifytext" then
  dim rs:set rs=server.createobject("adodb.recordset")
  rs.open "select * from KS_BlogTemplate Where ID="&KS.chkclng(KS.G("id")),conn,1,1
  if not rs.eof then
   templatename=rs("templatename")
   templateauthor=rs("templateauthor")
   templatepic=rs("templatepic")
   templatemain=rs("templatemain")
   templatesub=rs("templatesub")
   GroupID=rs("GroupID")
  end if
  Action="savetext"
 else
  Action="saveaddtext"
 end if
%>
<script>
 function CheckForm()
 {
    if (document.myform.TemplateName.value=='')
	{
	  alert('请输入模板名称!');
	  document.myform.TemplateName.focus();
	  return false;
	}
    if (document.myform.TemplateMain.value=='')
	{
	  alert('请输入主模板的内容!');
	  document.myform.TemplateMain.focus();
	  return false;
	}
    if (document.myform.TemplateSub.value=='')
	{
	  alert('请输入副模板的内容!');
	  document.myform.TemplateSub.focus();
	  return false;
	}
	 document.myform.submit();
 }
function ShowIframe(flag)
        {   
		 onscrolls=false;
         new KesionPopup().PopupCenterIframe("查看空间站点的可用标签","../editor/ksplus/spacelabel.asp?flag="+flag,590,340,'no')
       }
function InsertLabel(obj,Val)
{ return false;
  $(obj).focus();
  var str = document.selection.createRange();
  str.text = Val; 
  closeWindow();
 }
</script>
   <div style='height:35px;line-height:35px;text-align:center;font-weight:bold'>模板注册</div>
  <table width="100%" border="0" align="center" class="ctable" cellpadding="2" cellspacing="1">
<form method="POST" action="KS.SpaceSkin.asp?ID=<%=KS.G("id")%>&flag=<%=KS.g("flag")%>&action=<%=Action%>" id="myform" name="myform">
    <tr class="tdbg">
      <td align="right" class="clefttitle" height="25"><strong>模版名称：</strong></td>
	  <td><input name="TemplateName" type="text" id="TemplateName" value="<%=templatename%>"></td>
	</tr>
	<tr class="tdbg">
	  <td align="right" class="clefttitle"><strong>模板作者：</strong></td>
	  <td><input name="TemplateAuthor" type="text" id="TemplateAuthor" value="<%=templateauthor%>"> </td>
	</tr>
	<tr class="tdbg">
	   <td align="right" class="clefttitle"><strong>预 览 图：</strong></td>
	   <td><input type="text" name="TemplatePic" value="<%=templatepic%>">&nbsp;<input class='button' type='button' name='Submit' value='选择预览图...' onClick="OpenThenSetValue('Include/SelectPic.asp?Currpath=<%=CurrPath%>',550,290,window,document.all.TemplatePic);">
		</td>
    </tr>
	
	 <tr> 
	  <td height="25" class="clefttitle" align="right"><strong>首页模板：</strong></td>
      <td height="25" class="tdbg">
	  <input type="text" name="TemplateMain" id='TemplateMain' size='25' value="<%=templateMain%>"> <%=KSCls.Get_KS_T_C("$('#TemplateMain')[0]")%>
      </td>
    </tr>

	
    <tr class="tdbg"> 
	  <td height="25" class="clefttitle" align="right"><strong>其它页框架模板：</strong></td>
      <td height="25">
	  <input type="text" name="TemplateSub" id='TemplateSub' size='25' value="<%=templateSub%>"> <%=KSCls.Get_KS_T_C("$('#TemplateSub')[0]")%> 
      </td>
    </tr>
    <tr class="tdbg"> 
	  <td height="25" class="clefttitle" align="right">
	  <strong>允许使用本模板的用户组：</strong>
	  </td>
      <td height="25">
	 <%=KS.GetUserGroup_CheckBox("GroupID",GroupID,4)%>
      </td>
    </tr>
    <tr class="tdbg"> 
	  <td height="25"  colspan="2" >
	     <script type="text/javascript">
		 var num=1;
		 function addImg(obj)
           { 	  if (num>=100) {alert('最多只能定义100张图片!');return;}
		          num++;
                  var src  = obj.parentNode.parentNode;
                  var idx  = rowindex(src);
                  var tbl  = document.getElementById('gallery-table');
                  var row  = tbl.insertRow(idx+1);
                  var cell = row.insertCell(-1);
                  cell.innerHTML = src.cells[0].innerHTML.replace(/Url1/g,"Url"+num).replace(/(.*)(addImg)(.*)(\[)(\+)/i, "$1removeImg$3$4-");

           }
           // 删除图片上传
           function removeImg(obj)
           {
                  var row = rowindex(obj.parentNode.parentNode);
                  var tbl = document.getElementById('gallery-table');
                  tbl.deleteRow(row);
           }
		   var Browser = new Object();
            Browser.isMozilla = (typeof document.implementation != 'undefined') && (typeof document.implementation.createDocument != 'undefined') && (typeof HTMLDocument != 'undefined');
            Browser.isIE = window.ActiveXObject ? true : false;
            Browser.isFirefox = (navigator.userAgent.toLowerCase().indexOf("firefox") != - 1);
            Browser.isSafari = (navigator.userAgent.toLowerCase().indexOf("safari") != - 1);
            Browser.isOpera = (navigator.userAgent.toLowerCase().indexOf("opera") != - 1);
            function rowindex(tr)
            {
              if (Browser.isIE)
              {
                return tr.rowIndex;
              }
              else
              {
                table = tr.parentNode.parentNode;
                for (i = 0; i < table.rows.length; i ++ )
                {
                  if (table.rows[i] == tr)
                  {
                    return i;
                  }
                }
              }
            } 
		 </script>
<div class="attention">	     
		 <strong>说明：</strong>在您的空间模板里可以定义一些图片可以由用户自行上传更换，在希望用户可以更换的地方放上<span style="color:blue">{$ShowPicture1},{$ShowPicture2},{$ShowPicture3}...</span>进行调用即可，以下定义用户未上传自己图片时，显示的默认图片,其中的链接地址可选，如果不输入将不加链接，备注可以是广告文字，鼠标经过图片时显示等,如果勾选该图片为背景图，则该图片可以做为背景，不需要添加链接,在模板里只会输出图片的地址,非背景图请输入图片的宽和高。
	     <table width="100%" id="gallery-table"  align="center">
		 				<% 
				dim rss,pnum
				if KS.g("action")="modifytext" then
				  set rss=server.createobject("adodb.recordset")
				  rss.open "select * from KS_BlogSkin Where TemplateID=" & KS.chkclng(KS.G("id")) & " and isdefault=1 order by orderid,id",conn,1,1
				  pnum=0
				   do while not rss.eof
				     pnum=pnum+1
					%>
					<tr>
					 <td>
					 <%if pnum=1 then%>
					<a href="javascript:;" onclick="addImg(this)">[+]</a>
					 <%else%>
					<a href="javascript:;" onclick="removeImg(this)">[-]</a>
					 <%end if%>
				  图片名称 <input type="text" name="photonameUrl<%=pnum%>" value="<%=rss("photoname")%>" size="20"/> <span style='color:red'>*必填，如顶部banner,左边广告图等</span> <Br/>
				  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;图片地址 <input type="text" value="<%=rss("photourl")%>" name="imgUrl<%=pnum%>" id="Url<%=pnum%>" size="20" /> <button class='button' type='button' name='Submit' onClick="OpenThenSetValue('Include/SelectPic.asp?Currpath=<%=CurrPath%>',550,290,window,$('#Url<%=pnum%>')[0]);">选择</button>
				  <label><input type="checkbox" onclick="if (this.checked){$('#LUrl<%=pnum%>').hide()}else{$('#LUrl<%=pnum%>').show()}" name="TagUrl<%=pnum%>" value="1"<%if rss("isbg")=1 then response.write " checked"%>>背景图</label>
				  <span id="LUrl<%=pnum%>"<%if rss("isbg")=1 then response.write " style='display:none'"%>>
				  链接Url <input type="text" name="LinkUrl<%=pnum%>" value="<%=rss("linkurl")%>" size="20" /> <label><input type="checkbox" name="ModifyLinkUrl<%=pnum%>" value="1"<%if rss("modifylink")="1" then response.write " checked"%>>允许修链接</label></span>
				  宽<input style="text-align:center" type="text" name="widthUrl<%=pnum%>" value="<%=rss("width")%>" size="5"> 高<input type="text" name="heightUrl<%=pnum%>" style="text-align:center" value="<%=rss("height")%>" size="5">
				 
				  
				  <br/>
				  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;图片备注 <input type="text" name="NoteUrl<%=pnum%>" value="<%=rss("descript")%>" size="20" /><br/>
				   </td>
				  </tr>
					<%
				     rss.movenext
				   loop
				  %>
				   <script>
				  num=<%=pnum%>;
				  </script>

				  <%
				 end if
				%>
				<%if pnum=0 or KS.g("action")<>"modifytext" Then%>
			<tr>
				<td>

				  <a href="javascript:;" onclick="addImg(this)">[+]</a>
				  图片名称 <input type="text" name="photonameUrl1" size="20"/> <span style='color:red'>*必填，如顶部banner,左边广告图等</span> <Br/>
				  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;图片地址 <input type="text" name="imgUrl1" id="Url1" size="20" /> <input class='button' type='button' name='Submit' value='选择' onClick="OpenThenSetValue('Include/SelectPic.asp?Currpath=<%=CurrPath%>',550,290,window,$('#Url1')[0]);">
				  <label><input type="checkbox" onclick="if (this.checked){$('#LUrl1').hide()}else{$('#LUrl1').show()}" name="TagUrl1" value="1">背景图</label>
				  <span id="LUrl1">
				  链接Url <input type="text" name="LinkUrl1"  size="20" /> <label><input type="checkbox" name="ModifyLinkUrl1" value="1">允许修链接</label></span> 宽<input name="widthUrl1" style="text-align:center" type="text" size="5"> 高<input type="text" style="text-align:center" name="heightUrl1" size="5">
				  <br/>
				  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;图片备注 <input type="text" name="NoteUrl1" size="20" />
				</td>
			  </tr>
				<%end if %>  
		   </table>
	</div>	   
		   
		   
      </td>
    </tr>
</form>
  </table>
  
  <iframe src="../editor/ksplus/spacelabel.asp?flag=<%=flag%>" width="100%" height="300"></iframe>

<%
end sub


sub savePhoto(templateid,byref photots)
  dim i,orderid,num
  photots="":orderid=1:num=0
   conn.execute("delete from ks_blogskin where isdefault=1 and templateid=" & templateid)
  for i=1 to 100
   if request("photonameUrl" & i)<>"" Then
     num=num+1
     conn.execute("insert into KS_BlogSkin([templateid],[isdefault],[username],[photoname],[photourl],[linkurl],[Descript],[isbg],[orderid],[width],[height],[ModifyLink]) "&_
	 "values(" & templateid & ",1,'" & KS.C("AdminName") & "','" & replace(request("photonameUrl" & i),"'","''")&"','" & Request("imgUrl"& I) & "','" & request("linkurl"&i)&"','" & request("noteurl" & i) & "'," & KS.ChkClng(Request("TagUrl" & i)) & "," & orderid & "," & KS.ChkClng(request("widthurl"&i)) &"," & KS.ChkClng(request("heighturl"&i)) & "," & KS.ChkClng(Request("ModifyLinkUrl"&i))&")")
	 photots=photots &" " & Request("imgUrl"& I)
	 orderid=orderid+1
   end if
  next
  '删除超出原来可定义的无用的记录
  Conn.Execute("delete from ks_blogskin where templateid=" & templateid &" and orderid>" & num)
end sub

sub savetext()
	dim rs,sql,flag,photos
	set rs=server.CreateObject("adodb.recordset")
	sql="select * From KS_BlogTemplate where id=" & KS.chkclng(KS.g("id"))
	rs.open sql,conn,1,3
	rs("TemplateName")=trim(request("TemplateName"))
	rs("TemplateAuthor")=trim(request("TemplateAuthor"))
	rs("TemplateMain")=request("TemplateMain")
	rs("TemplatePic")=request("TemplatePic")
	rs("templatesub")=request("TemplateSub")
	rs("GroupID")=replace(request("GroupID")," ","")
	rs.update
	flag=rs("flag")
    call savePhoto(KS.chkclng(KS.g("id")),photos)
	rs.close:set rs=nothing
	'不是用户定义的图片就全部删除掉
	Conn.Execute("Delete From KS_UploadFiles Where ChannelID=1013 and InfoID=" & KS.chkclng(KS.g("id")) & " and filename not in(select photourl from ks_blogskin where templateid=" & KS.chkclng(KS.g("id")) & " and isdefault=0)")
	Call KS.FileAssociation(1013,KS.chkclng(KS.g("id")),request("TemplatePic")&" " & photos,0)
	response.Write  "<script>alert('模板修改成功!');location.href='KS.SpaceSkin.asp?flag=" & flag & "';</script>"
end sub
sub saveaddtext()
	dim rs,sql,photos
	set rs=server.CreateObject("adodb.recordset")
	sql="select top 1 * From KS_BlogTemplate"
	rs.open sql,conn,1,3
	rs.addnew
	rs("TemplateName")=trim(request("TemplateName"))
	rs("TemplateAuthor")=trim(request("TemplateAuthor"))
	rs("TemplatePic")=request("TemplatePic")
	rs("TemplateMain")=request("TemplateMain")
	rs("templatesub")=request("TemplateSub")
	rs("flag")=KS.chkclng(KS.g("flag"))
	rs("GroupID")=replace(request("GroupID")," ","")
	rs.update
	rs.movelast
	dim id:id=rs("id")
	rs.close:set rs=nothing
	call savePhoto(id,photos)
	Call KS.FileAssociation(1013,id,request("TemplatePic") & photos,0)
	response.Write  "<script>if (confirm('模板添加成功,继续添加吗？')==true){location.href='KS.SpaceSkin.asp?flag=" & ks.g("flag") & "&action=newtemplate';}else{location.href='KS.SpaceSkin.asp?flag=" & KS.g("flag") & "';}</script>"
end sub
sub savedefault()
	dim rs,isdefaultID
	isdefaultID=KS.ChkCLng(trim(request("radiobutton")))
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select id,isdefault From KS_BlogTemplate where flag=" & KS.chkclng(KS.g("flag")),conn,1,3
	while not rs.eof
		if isdefaultID=rs("id") then
			rs("isdefault")="true"
		else
			rs("isdefault")="false"
		end if
		rs.update
		rs.movenext
	wend
	rs.close
	set rs=nothing
	Response.Write"<script language=JavaScript>"
	Response.Write"alert(""修改成功！"");"
	Response.Write"window.history.go(-1);"
	Response.Write"</script>"
end sub


sub delconfig()
    conn.execute("delete from ks_UploadFiles where channelid=1013 and infoid=" & KS.ChkClng(KS.G("ID")))
	conn.execute("delete From KS_BlogTemplate where id="&KS.ChkCLng(KS.G("id")))
	conn.execute("delete From KS_BlogSkin where templateid="&KS.ChkCLng(KS.G("id")))
	response.Redirect "KS.SpaceSkin.asp?action=showconfig&flag=" &KS.g("flag")	
end sub


End Class
%>