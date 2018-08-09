<!--#include file="../../conn.asp"-->
<%
Dim Action:Action=Request("Action")
If Action="" Then Response.End()
Select Case Lcase(Action)
  case "ad" ad
  case "artphoto" ArtPhoto
  case "downphoto" DownPhoto
  case "flash" Flash
  case "flashplayer" flashPlayer
  case "getmusiclist" GetMusicList
  case "getspeciallist" GetSpecialList
  case "logo" Logo
  case "tags" tags
  case "moviedown" MovieDown
  case "moviepage" MoviePage
  case "moviephoto" MoviePhoto
  case "movieplay" MoviePlay
  case "productgroupphoto" ProductGroupPhoto
  case "productphoto" ProductPhoto
  case "status1" Status1
  case "status2" Status2
  case "status3" Status3
  case "supplyphoto" SupplyPhoto
  case "topuser" TopUser
  case "userdynamic" UserDynamic
End Select
%>
<%Sub Ad()%>
<html>
<head>
<title>插入对联广告参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
	if (document.myform.leftsrc.value=='')
	  {
	   alert('请输入左联地址!')
	   document.myform.leftsrc.focus();
	   return false;
	  }
	  if (document.myform.rightsrc.value=='')
	  {
	   alert('请输入右联地址!')
	   document.myform.rightsrc.focus();
	   return false;
	  }
	 if (document.myform.closesrc.value=='')
	  {
	   alert('请输入关闭图标地址!')
	   document.myform.closesrc.focus();
	   return false;
	  }
    Val = '{=JS_Ad("'+document.myform.leftsrc.value+'","'+document.myform.rightsrc.value+'","'+document.myform.closesrc.value+'",'+document.myform.speed.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
 
<link href="editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>对联广告设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td align="right"><div align="center">左联Flash地址</div></td>
    <td ><input name="leftsrc" type="text" id="leftsrc" size="60">
      100*250</td>
  </tr>
  <tr >
    <td align="right"><div align="center">右联Flash地址</div></td>
    <td ><input name="rightsrc" type="text" id="rightsrc" size="60">
      100*250</td>
  </tr>
  <tr >
    <td align="right"><div align="center">左右联底部关闭小图标</div></td>
    <td ><input name="closesrc" type="text" id="closesrc" value="/images/close.gif" size="40">
      输入&quot;0&quot;不显示</td>
  </tr>
  <tr >
    <td width="24%" align="right"><div align="center">滑动速度</div></td>
    <td width="76%" ><input name="speed" type="text" id="speed" value="0.8" size="8" onBlur="CheckNumber(this,'滚动速度');">
    范围 0.1~1.0 值越大,速度越快</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
 
<%End Sub

Sub ArtPhoto()
%>
<html>
<head>
<title>插入图片参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetPhoto('+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>显示图片设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">图片宽度：</div></td>
    <td width="60%" ><input name="PhotoWidth" type="text" onBlur="CheckNumber(this,'图片宽度');" size="6" value="130">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">图片高度：</div></td>
    <td ><input name="PhotoHeight" type="text" size="6" onBlur="CheckNumber(this,'图片高度');" value="90">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
 <%
End Sub
Sub DownPhoto()
%>
<html>
<head>
<title>插入下载缩略图参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetDownPhoto('+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>显示下载缩略图设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">下载缩略图宽度：</div></td>
    <td width="60%" ><input name="PhotoWidth" type="text" onBlur="CheckNumber(this,'下载缩略图宽度');" size="6" value="130">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">下载缩略图高度：</div></td>
    <td ><input name="PhotoHeight" type="text" size="6" onBlur="CheckNumber(this,'下载缩略图高度');" value="90">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub Flash()
%>
<html>
<head>
<title>插入Flash参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetFlash('+document.myform.FlashWidth.value+','+document.myform.FlashHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>显示Flash设置</LEGEND>
<table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">Flash宽度：</div></td>
    <td width="60%" ><input name="FlashWidth" type="text" onBlur="CheckNumber(this,'Flash播放器宽度');" size="6" value="550">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">Flash高度：</div></td>
    <td ><input name="FlashHeight" type="text" size="6" onBlur="CheckNumber(this,'Flash播放器高度');" value="380">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%End Sub
Sub FlashPlayer()
%>
<html>
<head>
<title>Flash播放器参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetFlashByPlayer('+document.myform.FlashWidth.value+','+document.myform.FlashHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>显示Flash播放器设置</LEGEND>
<table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">Flash播放器宽度：</div></td>
    <td width="60%" ><input name="FlashWidth" type="text" onBlur="CheckNumber(this,'Flash播放器宽度');" size="6" value="550">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">Flash播放器高度：</div></td>
    <td ><input name="FlashHeight" type="text" size="6" onBlur="CheckNumber(this,'Flash播放器高度');" value="380">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub GetMusicList()
%>
<html>
<head>
<title>歌曲播放列表参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var TypeID,Val,ShowSelect,type,ShowMouseTX,ShowDetailTF;
	
	for (var i=0;i<document.myform.ShowSelect.length;i++){
	 var KM = document.myform.ShowSelect[i];
	if (KM.checked==true)	   
		ShowSelect = KM.value
	}
	for (var i=0;i<document.myform.type.length;i++){
	 var KM = document.myform.type[i];
	if (KM.checked==true)	   
		type = KM.value
	}
	for (var i=0;i<document.myform.ShowMouseTX.length;i++){
	 var KM = document.myform.ShowMouseTX[i];
	if (KM.checked==true)	   
		ShowMouseTX = KM.value
	}
	for (var i=0;i<document.myform.ShowDetailTF.length;i++){
	 var KM = document.myform.ShowDetailTF[i];
	if (KM.checked==true)	   
		ShowDetailTF = KM.value
	}

    Val = '{=GetMusicList('+document.myform.TypeID.value+','+ShowSelect+','+type+','+document.myform.Num.value+','+document.myform.RowHeight.value+','+ShowMouseTX+','+ShowDetailTF+','+document.myform.Row.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>

<link href="Editor.css" rel="stylesheet">
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>歌曲播放列表参数设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td align="right"><div align="center">选择类别</div></td>
    <td >
	<select name="TypeID">
	 <option value='0'>-不指定任何类别-</option>
	 <option value='-1' style="color:red">-当前类别通用-</option>
	 <%
	  dim rs
	  set rs=server.createobject("adodb.recordset")
	  rs.open "select SclassID,Sclass from KS_MSSClass",conn,1,1
	  do while not rs.eof
	    response.write "<option value=""" & rs(0) & """>" & rs(1) & "</option>"
		rs.movenext
	  loop
	  rs.close
	  set rs=nothing
	  conn.close
	  set conn=nothing
	 %>
	</select>
	</td>
  </tr>
  <tr >
    <td align="right"><div align="center">显示选择框</div></td>
    <td ><input name="ShowSelect" type="radio" value="true" checked>
      是
        <input type="radio" name="ShowSelect" value="false">
        否</td>
  </tr>
  <tr >
    <td align="right"><div align="center">列表属性</div></td>
    <td ><input name="type" type="radio" value="0" checked>
      最新歌曲
        <input type="radio" name="type" value="1">
        推荐歌曲
        <input type="radio" name="type" value="2">
        热点歌曲</td>
  </tr>
  <tr >
    <td width="24%" align="right"><div align="center">列出多少首歌曲</div></td>
    <td width="76%" ><input name="Num" type="text" id="Num" value="10" size="8" onBlur="CheckNumber(this,'歌曲首数');">
      首 每行显示: 
        <input name="Row" type="text" id="Row" value="2" size="6" onBlur="CheckNumber(this,'歌曲首数');">
        首</td>
  </tr>
  <tr >
    <td align="right"><div align="center">歌曲之间的行距</div></td>
    <td ><input name="RowHeight" type="text" id="RowHeight" value="25" size="8" onBlur="CheckNumber(this,'歌曲首数');">
      px</td>
  </tr>
  <tr >
    <td align="right"><div align="center">鼠标经过是否特效</div></td>
    <td ><input name="ShowMouseTX" type="radio" value="true" checked>
是
  <input type="radio" name="ShowMouseTX" value="false">
否</td>
  </tr>
  <tr >
    <td align="right"><div align="center">列出是否显示详细</div></td>
    <td ><input name="ShowDetailTF" type="radio" value="true" checked>
是
  <input type="radio" name="ShowDetailTF" value="false">
否 　显示歌曲的详细，如下载，收藏等</td>
  </tr>
</table>
</FIELDSET></td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
<tr>
  <td height="30"><div align="center"><span class="STYLE1">备注：此标签音乐频道通用</span></div></td>
</tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub GetSpecialList()
%>
<html>
<head>
<title>专辑列表参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val,type,ShowMouseTX,ShowDetailTF;
	
	for (var i=0;i<document.myform.type.length;i++){
	 var KM = document.myform.type[i];
	if (KM.checked==true)	   
		type = KM.value
	}

	for (var i=0;i<document.myform.ShowDetailTF.length;i++){
	 var KM = document.myform.ShowDetailTF[i];
	if (KM.checked==true)	   
		ShowDetailTF = KM.value
	}

    Val = '{=GetMusicSpecialList('+type+','+document.myform.Num.value+','+document.myform.ColNum.value+','+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+','+document.myform.SpecialNameLen.value+','+ShowDetailTF+')}';  
    window.returnValue = Val;
    window.close();
}
</script>

<link href="Editor.css" rel="stylesheet">
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>专辑列表参数设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td align="right"><div align="center">列表属性</div></td>
    <td ><input name="type" type="radio" value="0" checked>
      最新专辑
        <input type="radio" name="type" value="1">
        推荐专辑
        <input type="radio" name="type" value="2">
        热点专辑</td>
  </tr>
  <tr >
    <td width="24%" align="right"><div align="center">列出多少张专辑</div></td>
    <td width="76%" ><input name="Num" type="text" id="Num" value="10" size="8" onBlur="CheckNumber(this,'列出多少张专辑');">
      张</td>
  </tr>
  <tr >
    <td align="right"><div align="center">专辑排列列数</div></td>
    <td ><input name="ColNum" type="text" id="ColNum" value="1" size="8" onBlur="CheckNumber(this,'专辑排列列数');">
      px</td>
  </tr>
  <tr >
    <td align="right"><div align="center">专辑图片的宽度</div></td>
    <td ><input name="PhotoWidth" type="text" id="PhotoWidth" value="90" size="8" onBlur="CheckNumber(this,'专辑图片的宽度');">
px</td>
  </tr>
  <tr >
    <td align="right"><div align="center">专辑图片的高度</div></td>
    <td ><input name="PhotoHeight" type="text" id="PhotoHeight" value="80" size="8" onBlur="CheckNumber(this,'专辑图片的高度');">
px</td>
  </tr>
  <tr >
    <td align="right"><div align="center">取专辑名称字数</div></td>
    <td ><input name="SpecialNameLen" type="text" id="SpecialNameLen" value="8" size="8" onBlur="CheckNumber(this,'取专辑名称字数');"> 
      字 一个汉字=两个英文字符 </td>
  </tr>
  <tr >
    <td align="right"><div align="center">是否显示发行公司及发行日期</div></td>
    <td ><input name="ShowDetailTF" type="radio" value="true" checked>
是
  <input type="radio" name="ShowDetailTF" value="false">
否 　显示歌曲的详细，如下载，收藏等</td>
  </tr>
</table>
</FIELDSET></td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
<tr>
  <td height="30"><div align="center"><span class="STYLE1">备注：些标签音乐频道通用</span></div></td>
</tr>
</table>
</form>
</body>
</html>
 <%End Sub
Sub Logo
%>
<html>
<head>
<title>插入网站Logo参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetLogo('+document.myform.FlashWidth.value+','+document.myform.FlashHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>显示Logo设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">Logo宽度：</div></td>
    <td width="60%" ><input name="FlashWidth" type="text" onBlur="CheckNumber(this,'Flash播放器宽度');" size="6" value="130">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">Logo高度：</div></td>
    <td ><input name="FlashHeight" type="text" size="6" onBlur="CheckNumber(this,'Flash播放器高度');" value="90">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub Tags()
%>
<html>
<head>
<title>插入Tags参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetTags('+document.myform.sorts.value+','+document.myform.num.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>显示Tags设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">显示Tags数：</div></td>
    <td width="60%" ><input name="num" type="text" id="num" value="50" size="6">
    个</td>
  </tr>
  <tr>
    <td align="right"><div align="center">Tags排序方式：</div></td>
    <td ><select name="sorts">
      <option value="1">点击数降序(热门Tags)</option>
      <option value="2">最后访问时间降序</option>
      <option value="3">添加时间</option>
    </select>
    </td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%End Sub
Sub MovieDown()
%>
<html>
<head>
<title>插入影片下载列表设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetMovieDownList('+document.myform.Num.value+',"'+document.myform.Navi.value+'")}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>显示影片下载列表设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">每行显示集数：</div></td>
    <td width="60%" ><input name="Num" type="text" onBlur="CheckNumber(this,'每行显示集数');" size="15" value="5">
     </td>
  </tr>
  <tr>
    <td align="right"><div align="center">导航图标：</div></td>
    <td ><input name="Navi" type="text" size="15" onBlur="CheckNumber(this,'导航图标');" value="/images/movienavi.gif">
    </td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub MoviePage()
%>
<html>
<head>
<title>插入内容页flv播放器设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetMoviePagePlay('+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>显示内容页flv播放器设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">影片宽度：</div></td>
    <td width="60%" ><input name="PhotoWidth" type="text" onBlur="CheckNumber(this,'影片宽度');" size="6" value="450">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">影片高度：</div></td>
    <td ><input name="PhotoHeight" type="text" size="6" onBlur="CheckNumber(this,'影片高度');" value="450">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub MoviePhoto()
%>
<html>
<head>
<title>插入影片图片参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetMoviePhoto('+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>显示影片图片设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">影片图片宽度：</div></td>
    <td width="60%" ><input name="PhotoWidth" type="text" onBlur="CheckNumber(this,'影片图片宽度');" size="6" value="250">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">影片图片高度：</div></td>
    <td ><input name="PhotoHeight" type="text" size="6" onBlur="CheckNumber(this,'影片图片高度');" value="250">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub MoviePlay()
%>
<html>
<head>
<title>插入影片播放列表设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetMoviePlayList('+document.myform.Num.value+',"'+document.myform.Navi.value+'")}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>显示影片播放列表设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">每行显示集数：</div></td>
    <td width="60%" ><input name="Num" type="text" onBlur="CheckNumber(this,'每行显示集数');" size="15" value="5">
     </td>
  </tr>
  <tr>
    <td align="right"><div align="center">导航图标：</div></td>
    <td ><input name="Navi" type="text" size="15" onBlur="CheckNumber(this,'导航图标');" value="/images/movienavi.gif">
    </td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub ProductGroupPhoto()
%>
<html>
<head>
<title>插入商品图片组参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetGroupPhoto('+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>显示商品图片组设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">预览图宽度：</div></td>
    <td width="60%" ><input name="PhotoWidth" type="text" onBlur="CheckNumber(this,'商品预览图宽度');" size="6" value="200">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">商品预览图高度：</div></td>
    <td ><input name="PhotoHeight" type="text" size="6" onBlur="CheckNumber(this,'商品预览图高度');" value="200">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub ProductPhoto()
%>
<html>
<head>
<title>插入商品缩略图参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetProductPhoto('+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>显示商品缩略图设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">商品缩略图宽度：</div></td>
    <td width="60%" ><input name="PhotoWidth" type="text" onBlur="CheckNumber(this,'商品缩略图宽度');" size="6" value="130">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">商品缩略图高度：</div></td>
    <td ><input name="PhotoHeight" type="text" size="6" onBlur="CheckNumber(this,'商品缩略图高度');" value="90">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub Status1()
%>
<html>
<head>
<title>插入状态栏打字效果参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
	if (document.myform.text.value=='')
	  {
	   alert('请输入文字!')
	   document.myform.text.focus();
	   return false;
	  }
    Val = '{=JS_Status1("'+document.myform.text.value+'",'+document.myform.speed.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<script language="JavaScript">
var msg = "欢迎您使用科汛网站管理系统! " ;
var interval = 120
var spacelen = 120;
var space10=" ";
var seq=0;
function KS_Status1() {
len = msg.length;
window.status = msg.substring(0, seq+1);
seq++;
if ( seq >= len ) {
seq = 0;
window.status = '';
window.setTimeout("KS_Status1();", interval );
}
else
window.setTimeout("KS_Status1();", interval );
}
KS_Status1();
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>状态栏打字效果设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td align="right"><div align="center">待显示的文字</div></td>
    <td ><input name="text" type="text" id="text" size="60"></td>
  </tr>
  <tr >
    <td align="right">&nbsp;</td>
    <td >如: 欢迎光临本站!!!</td>
  </tr>
  <tr >
    <td width="24%" align="right"><div align="center">滚动速度</div></td>
    <td width="76%" ><input name="speed" type="text" id="speed" value="120" size="8" onBlur="CheckNumber(this,'滚动速度');">值越大,速度越慢</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub Status2()
%>
<html>
<head>
<title>插入状态栏文字在状态栏上从右往左循环显示参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
	if (document.myform.text.value=='')
	  {
	   alert('请输入文字!')
	   document.myform.text.focus();
	   return false;
	  }
    Val = '{=JS_Status2("'+document.myform.text.value+'",'+document.myform.speed.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>

<script>
<!--
function KS_Status2(seed)
{ var m1 = "欢迎您使用科汛网站管理系统!" ;
var m2 = "" ;
var msg=m1+m2;
var out = " ";
var c = 1;
var speed = 120;
if (seed > 100)
{ seed-=2;
var cmd="KS_Status2(" + seed + ")";
timerTwo=window.setTimeout(cmd,speed);}
else if (seed <= 100 && seed > 0)
{ for (c=0 ; c < seed ; c++)
{ out+=" ";}
out+=msg; seed-=2;
var cmd="KS_Status2(" + seed + ")";
window.status=out;
timerTwo=window.setTimeout(cmd,speed); }
else if (seed <= 0)
{ if (-seed < msg.length)
{
out+=msg.substring(-seed,msg.length);
seed-=2;
var cmd="KS_Status2(" + seed + ")";
window.status=out;
timerTwo=window.setTimeout(cmd,speed);}
else { window.status=" ";
timerTwo=window.setTimeout("KS_Status2(100)",speed);
}
}
}
KS_Status2(100);
-->
</script>
      
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>文字在状态栏上从右往左循环显示效果设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td align="right"><div align="center">待显示的文字</div></td>
    <td ><input name="text" type="text" id="text" size="60"></td>
  </tr>
  <tr >
    <td align="right">&nbsp;</td>
    <td >如: 欢迎光临本站!!!</td>
  </tr>
  <tr >
    <td width="24%" align="right"><div align="center">滚动速度</div></td>
    <td width="76%" ><input name="speed" type="text" id="speed" value="120" size="8" onBlur="CheckNumber(this,'滚动速度');">
    值越大,速度越慢</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub Status3()
%>
<html>
<head>
<title>插入状态栏文字在状态栏上从右往左循环显示参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
	if (document.myform.text.value=='')
	  {
	   alert('请输入文字!')
	   document.myform.text.focus();
	   return false;
	  }
    Val = '{=JS_Status3("'+document.myform.text.value+'",'+document.myform.speed.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>

    <SCRIPT LANGUAGE="JavaScript">
<!--
var Message="欢迎您使用科汛网站管理系统! ";
var place=1;
function scrollIn() {
window.status=Message.substring(0, place);
if (place >= Message.length) {
place=1;
window.setTimeout("KS_Status3()",300);
} else {
place++;
window.setTimeout("scrollIn()",200);
}
}
function KS_Status3() {
window.status=Message.substring(place, Message.length);
if (place >= Message.length) {
place=1;
window.setTimeout("scrollIn()", 100);
} else {
place++;
window.setTimeout("KS_Status3()", 200);
}
}
KS_Status3();
-->
</SCRIPT>  
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>文字在状态栏上从右往左循环显示效果设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td align="right"><div align="center">待显示的文字</div></td>
    <td ><input name="text" type="text" id="text" size="60"></td>
  </tr>
  <tr >
    <td align="right">&nbsp;</td>
    <td >如: 欢迎光临本站!!!</td>
  </tr>
  <tr >
    <td width="24%" align="right"><div align="center">滚动速度</div></td>
    <td width="76%" ><input name="speed" type="text" id="speed" value="150" size="8" onBlur="CheckNumber(this,'滚动速度');">
    值越大,速度越慢</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%End Sub

Sub SupplyPhoto()
%>
<html>
<head>
<title>插入供求信息缩略图参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetSupplyPhoto('+document.myform.PhotoWidth.value+','+document.myform.PhotoHeight.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>显示供求信息缩略图设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">供求信息缩略图宽度：</div></td>
    <td width="60%" ><input name="PhotoWidth" type="text" onBlur="CheckNumber(this,'供求信息缩略图宽度');" size="6" value="130">
      px</td>
  </tr>
  <tr>
    <td align="right"><div align="center">供求信息缩略图高度：</div></td>
    <td ><input name="PhotoHeight" type="text" size="6" onBlur="CheckNumber(this,'供求信息缩略图高度');" value="90">
    px</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub TopUser()
%>
<html>
<head>
<title>插入用户登录排行参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetTopUser('+document.myform.num.value+','+document.myform.more.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>显示用户登录排行设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">显示用户数：</div></td>
    <td width="60%" ><input name="num" type="text" id="num" value="5" size="6">
      位</td>
  </tr>
  <tr>
    <td align="right"><div align="center">更多链接：</div></td>
    <td ><input name="more" type="text" id="more" value="more..." size="20"> 留空不输出</td>
  </tr>
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
Sub UserDynamic()
%>
<html>
<head>
<title>插入用户动态参数设置</title>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var Val;
    Val = '{=GetUserDynamic('+document.myform.num.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>显示用户动态标签设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td width="40%" align="right"><div align="center">显示最新：</div></td>
    <td width="60%" ><input name="num" type="text" id="num" value="10" size="6">
      条</td>
  </tr>
  
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
End Sub
%>