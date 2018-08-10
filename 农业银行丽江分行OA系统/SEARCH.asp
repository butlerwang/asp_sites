<%
key=trim(request("key"))
sm=trim(request("sm"))
%>
<html>
<head>
<title>全国邮政编码、电话区号查询</title>
<style>
a:link{font-size:9pt;color:#004080;text-decoration:none}
a:visited{text-decoration:none}
a:hover{color:red;TEXT-DECORATION:underline}
td {font-size:9pt}
BODY{font-size:9pt;scrollbar-face-color:#3DB5E2;scrollbar-3dlight-color:#C0C0C0;scrollbar-darkshadow-color:#C0C0C0;scrollbar-track-color:#C0C0C0;scrollbar-arrow-color:#C0C0C0;scrollbar-shadow-color:#C0C0C0;}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css"><!--
.white { font-size: 9pt; font-weight: bold; color: #FFFFFF}
A {font-size: 9pt; COLOR: #000000; TEXT-DECORATION: none}
A:visited {	font-size: 12px; COLOR: #333333}
A:active {	font-size: 12px; COLOR: #ff0000}
A:hover {	font-size: 12px; COLOR: #0000CC; text-decoration: underline}
.words { font-size: 12px; color: #ffffff}
.form1 {border: 2px #FFFFFF solid; border-color: #ffffff #ffffff #ffffff; background-color: #999999; font-family: "宋体"; font-size: 9pt; color: #FFFFFF}
.L15 {LINE-HEIGHT: 150%} .L20 {LINE-HEIGHT: 200%} .16f {FONT-SIZE: 16px; COLOR: #000066} .rf {COLOR: #FF6600; font-size: 70%} .bf {font-weight: bold; color: #FFFFFF; font-size: 12px}
a:link { font-size: 12px; color: #000000}
td { font-size: 12px}
--></style>

</head>

<body bgcolor="#C0C0C0">
<div align="center"> <div align="center"> </div><table width="83%" border="1" cellspacing="0" cellpadding="0" bordercolorlight="#666666" bordercolordark="#FFFFFF" bgcolor="#CCFFCC"> 
<tr> <td> <div align="center">省洲名称</div></td><td> <div align="center">地区名称</div></td><td> 
<div align="center">邮政编码</div></td><td> <div align="center">电话区号</div></td></tr> 
<%
if key<>"" then
connstr="DBQ="+server.mappath("ybqh.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
conn.open connstr
set rst=server.createobject("adodb.recordset")
rst.open "select * from ybqh",conn,1,1
do while (not rst.eof)
tmpstr=rst("sm")&getpy(rst("sm"))&rst("dq")&getpy(rst("dq"))&rst("yb")&" "&rst("qh")
if instr(1,tmpstr,key,1)>0 then
 %> <tr> <td><%=rst("sm")%></td><td><%=rst("dq")%></td><td><%
tmp=rst("yb")
if tmp="" then tmp="★★★"
response.write tmp
%></td><td><%=rst("qh")
  %></td></tr> <%
end if
rst.movenext
loop
%> </table><%
rst.close
conn.close
end if
%> <table width="610" border="1" cellspacing="0" cellpadding="0" bordercolorlight="#666666" bordercolordark="#FFFFFF"> 
<form name="form1" action="search.asp" method="post" > <tr bgcolor="#CCCCFF"> 
<td width="610" colspan="5"> 关键字： <input style="BORDER-BOTTOM: #000000 1px dotted; BORDER-LEFT: #000000 1px dotted; BORDER-RIGHT: #000000 1px dotted; BORDER-TOP: #000000 1px dotted; FONT-SIZE: 10pt" type="text" name="key" value="<%=key%>" size="20"> 
<input style="BORDER-BOTTOM: #000000 1px dotted; BORDER-LEFT: #000000 1px dotted; BORDER-RIGHT: #000000 1px dotted; BORDER-TOP: #000000 1px dotted; FONT-SIZE: 9pt" type="submit" name="submit" value="搜索"> 
<font color="#FF0000">输入一个省名、地区名、邮编、区号或其<font color="#66FF66"><b><font color="#000000">拼音首字母</font></b></font>或其中一部分 
</font> </td></tr> </form><tr bgcolor="#CCCCFF"> <td width="610" height="426" bgcolor="#C0C0C0" colspan="5"><h2><img src="glasses.gif" webstripperlinkwas="glasses.gif" alt="WB00730_.gif (767 字节)" width="46" height="38"><a name="按地图位置查询">按地图位置查询</a></h2><p align="center"><map name="FPMap0"> 
<area href="index.asp?sm=%CC%A8%CD%E5" webstripperlinkwas="taiwan.htm" shape="rect" coords="388, 282, 420, 313"> 
<area href="index.asp?sm=%CE%F7%B2%D8" webstripperlinkwas="xizang.htm" shape="polygon" coords="31, 186, 113, 183, 120, 214, 183, 230, 178, 263, 137, 268, 26, 211"> 
<area href="index.asp?sm=%B9%E3%CE%F7" webstripperlinkwas="guangxi.htm" shape="rect" coords="266, 301, 298, 317"> 
<area href="index.asp?sm=%C7%E0%BA%A3" webstripperlinkwas="qinhai.htm" shape="rect" coords="137, 159, 209, 209"> 
<area href="index.asp?sm=%BA%D3%C4%CF" webstripperlinkwas="henan.htm" shape="rect" coords="300, 191, 329, 214"> 
<area href="index.asp?sm=%C9%BD%CE%F7" webstripperlinkwas="shanxi.htm" shape="rect" coords="292, 155, 310, 187"> 
<area href="index.asp?sm=%B0%B2%BB%D5" webstripperlinkwas="anhui.htm" shape="polygon" coords="335, 213, 349, 203, 367, 233, 351, 239"> 
<area href="index.asp?sm=%C9%C2%CE%F7" webstripperlinkwas="shanxi2.htm" shape="rect" coords="268, 180, 288, 224"> 
<area href="index.asp?sm=%BC%AA%C1%D6" webstripperlinkwas="jilin.htm" shape="polygon" coords="364, 79, 378, 92, 394, 103, 406, 114, 422, 102, 412, 85, 393, 72, 366, 74"> 
<area href="index.asp?sm=%D4%C6%C4%CF" webstripperlinkwas="yunnan.htm" shape="polygon" coords="191, 262, 204, 262, 213, 278, 224, 281, 233, 291, 241, 305, 247, 314, 234, 323, 214, 322, 212, 327, 205, 329, 188, 317, 188, 309, 177, 293, 195, 275, 192, 266"> 
<area href="index.asp?sm=%D0%C2%BD%AE" webstripperlinkwas="xinjiang.htm" shape="polygon" coords="17, 157, 34, 157, 56, 162, 65, 165, 85, 164, 100, 169, 110, 172, 116, 171, 118, 159, 134, 155, 152, 138, 160, 111, 159, 104, 147, 98, 149, 83, 137, 61, 132, 55, 115, 65, 109, 65, 99, 63, 90, 77, 79, 75, 73, 75, 69, 91, 59, 96, 63, 120, 31, 117, 17, 130, 11, 130, 7, 132, 7, 138"> 
<area href="index.asp?sm=%C4%DA%C3%C9%B9%C5" webstripperlinkwas="neimenggu.htm" shape="polygon" coords="183, 120, 193, 130, 209, 146, 225, 158, 241, 162, 267, 158, 297, 145, 317, 136, 328, 132, 323, 106, 301, 111, 291, 114, 283, 122, 278, 126, 253, 134, 222, 131, 202, 120, 187, 120"> 
<area href="index.asp?sm=%BD%AD%CE%F7" webstripperlinkwas="jiangxi.htm" shape="polygon" coords="331, 249, 348, 249, 351, 249, 350, 271, 336, 284, 332, 279, 330, 267, 331, 253"> 
<area href="index.asp?sm=%BA%D3%B1%B1" webstripperlinkwas="hebei.htm" shape="polygon" coords="330, 130, 317, 135, 313, 153, 317, 167, 331, 167, 330, 145"> 
<area href="index.asp?sm=%B8%A3%BD%A8" webstripperlinkwas="fujian.htm" shape="polygon" coords="386, 266, 363, 260, 347, 273, 353, 282, 362, 290, 386, 270"> 
<area href="index.asp?sm=%C9%BD%B6%AB" webstripperlinkwas="shandong.htm" shape="polygon" coords="347, 162, 333, 170, 337, 186, 363, 182, 377, 166, 363, 164, 352, 164"> 
<area href="index.asp?sm=%C1%C9%C4%FE" webstripperlinkwas="liaoning.htm" shape="polygon" coords="361, 107, 377, 101, 386, 104, 399, 117, 384, 131, 371, 135, 362, 121"> 
<area href="index.asp?sm=%BA%A3%C4%CF" webstripperlinkwas="hainan.htm" shape="rect" coords="280, 346, 334, 361"> 
<area href="index.asp?sm=%D5%E3%BD%AD" webstripperlinkwas="zhejiang.htm" shape="rect" coords="365, 235, 399, 249"> 
<area href="index.asp?sm=%BA%DA%C1%FA%BD%AD" webstripperlinkwas="heilongjiang.htm" shape="rect" coords="377, 33, 437, 71"> 
<area href="index.asp?sm=%D6%D8%C7%EC" webstripperlinkwas="chongqing.htm" shape="rect" coords="214, 243, 247, 262"> 
<area href="index.asp?sm=%BD%AD%CB%D5" webstripperlinkwas="jiangsu.htm" shape="polygon" coords="351, 200, 368, 190, 387, 216, 372, 228"> 
<area href="index.asp?sm=%B8%CA%CB%E0" webstripperlinkwas="gansu.htm" shape="polygon" coords="216, 170, 237, 161, 247, 192, 254, 212, 236, 218"> 
<area href="index.asp?sm=%B9%E3%B6%AB" webstripperlinkwas="guangdong.htm" shape="rect" coords="311, 296, 344, 315"> 
<area href="index.asp?sm=%CB%C4%B4%A8" webstripperlinkwas="sichuan.htm" shape="rect" coords="200, 227, 245, 246"> 
<area href="index.asp?sm=%B9%F3%D6%DD" webstripperlinkwas="guizhou.htm" shape="rect" coords="250, 271, 277, 286"> 
<area href="index.asp?sm=%BA%FE%B1%B1" webstripperlinkwas="hubei.htm" shape="rect" coords="293, 225, 324, 241"> 
<area href="index.asp?sm=%C4%FE%CF%C4" webstripperlinkwas="ningxia.htm" shape="rect" coords="244, 156, 262, 191"> 
<area href="index.asp?sm=%C9%CF%BA%A3" webstripperlinkwas="shanghai.htm" shape="rect" coords="386, 212, 412, 228"> 
<area href="index.asp?sm=%BA%FE%C4%CF" webstripperlinkwas="hunan.htm" shape="rect" coords="297, 250, 317, 282"> 
<area href="index.asp?sm=%CC%EC%BD%F2" webstripperlinkwas="tianjin.htm" shape="rect" coords="333, 143, 360, 156"> 
<area href="index.asp?sm=%B1%B1%BE%A9" webstripperlinkwas="beijing.htm" shape="rect" coords="334, 127, 362, 142"></map><img rectangle="(334,127) (362, 142)  beijing.htm" rectangle="(300,191) (329, 214)  henan.htm" rectangle="(292,155) (310, 187)  shanxi.htm" rectangle="(268,180) (288, 224)  shanxi2.htm" polygon="(216,170) (237,161) (247,192) (254,212) (236,218) gansu.htm" polygon="(191,262) (204,262) (213,278) (224,281) (233,291) (241,305) (247,314) (234,323) (214,322) (212,327) (205,329) (188,317) (188,309) (177,293) (195,275) (192,266) yunnan.htm" polygon="(17,157) (34,157) (56,162) (65,165) (85,164) (100,169) (110,172) (116,171) (118,159) (134,155) (152,138) (160,111) (159,104) (147,98) (149,83) (137,61) (132,55) (115,65) (109,65) (99,63) (90,77) (79,75) (73,75) (69,91) (59,96) (63,120) (31,117) (17,130) (11,130) (7,132) (7,138) xinjiang.htm" polygon="(183,120) (193,130) (209,146) (225,158) (241,162) (267,158) (297,145) (317,136) (328,132) (323,106) (301,111) (291,114) (283,122) (278,126) (253,134) (222,131) (202,120) (187,120) neimenggu.htm" polygon="(331,249) (348,249) (351,249) (350,271) (336,284) (332,279) (330,267) (331,253) jiangxi.htm" polygon="(330,130) (317,135) (313,153) (317,167) (331,167) (330,145) hebei.htm" polygon="(386,266) (363,260) (347,273) (353,282) (362,290) (386,270) fujian.htm" polygon="(347,162) (333,170) (337,186) (363,182) (377,166) (363,164) (352,164) shandong.htm" polygon="(361,107) (377,101) (386,104) (399,117) (384,131) (371,135) (362,121) liaoning.htm" rectangle="(280,346) (334, 361)  hainan.htm" rectangle="(365,235) (399, 249)  zhejiang.htm" rectangle="(377,33) (437, 71)  heilongjiang.htm" rectangle="(214,243) (247, 262)  chongqing.htm" polygon="(351,200) (368,190) (387,216) (372,228) jiangsu.htm" polygon="(216,170) (237,161) (247,192) (254,212) (236,218) gansu.htm" rectangle="(311,296) (344, 315)  guangdong.htm" rectangle="(200,227) (245, 246)  sichuan.htm" rectangle="(250,271) (277, 286)  guizhou.htm" rectangle="(293,225) (324, 241)  hubei.htm" rectangle="(244,156) (262, 191)  ningxia.htm" rectangle="(386,212) (412, 228)  shanghai.htm" rectangle="(297,250) (317, 282)  hunan.htm" rectangle="(333,143) (360, 156)  tianjin.htm" rectangle="(334,127) (362, 142)  beijing.htm" src="chinamap.gif" webstripperlinkwas="images/chinamap.gif" border="0" usemap="#FPMap0"></p></td></tr> 
<tr> <td width="150" height="12" bgcolor="#C0C0C0"><div align="center"><a href="index.asp?sm=%D1%C7%D6%DE">亚洲</a></div></td><td width="150" height="12" bgcolor="#C0C0C0"><div align="center"><a href="index.asp?sm=%C5%B7%D6%DE">欧洲</a></div></td><td width="150" height="12" bgcolor="#C0C0C0"><div align="center"><a href="index.asp?sm=%C3%C0%D6%DE">美洲</a></div></td><td width="151" height="12" bgcolor="#C0C0C0"><div align="center"><a href="index.asp?sm=%B4%F3%D1%F3%D6%DE">大洋洲</a></div></td><td width="151" height="12" bgcolor="#C0C0C0"><div align="center"><a href="index.asp?sm=%B7%C7%D6%DE">非洲</a></div></td></tr> 
</table><!--#include file="py.asp" --></div>
</body>
</html>
