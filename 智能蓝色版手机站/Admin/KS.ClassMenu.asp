<%@language=vbscript CODEPAGE="65001" %>
<%
Option Explicit
Response.buffer = True
Server.ScriptTimeout=9999999
%>
<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="KS.ClassMenuParam.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KS:Set KS=New PublicCls
Dim strInstallDir,ComeUrl
If Not KS.ReturnPowerResult(0, "KMSL10008") Then          
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
Response.Write "<div class='topdashed'>"
Response.Write "<b>管理导航:</b>&nbsp;&nbsp;<a href='KS.ClassMenu.asp?Action=ShowConfig&ChannelID=" & ChannelID & "'>参数设置</a>"
Response.Write " | <a href='KS.ClassMenu.asp?Action=ShowCreate&ChannelID=" & ChannelID & "'>菜单生成</a>"
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;注：参数设置▲代表鼠标悬停时效果，▼代表鼠标移出时效果。</div>"


Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>"

Response.Write "  <tr>"
Response.Write "    <td width='70' height='30' style=""border-bottom:1px dashed #a7a7a7""><strong>菜单演示：</strong>"
Response.Write "    <td height='30' style=""border-bottom:1px dashed #a7a7a7"">"
Call ShowDemoMenu
Response.Write "    </td>"

Response.Write "  </tr></table>" & vbCrLf
Dim Action:Action=KS.G("Action")
If Action = "ShowConfig" Then
    Call ShowConfig
ElseIf Action = "SaveConfig" Then
    Call SaveConfig
ElseIf Action = "ShowCreate" Then
    Call ShowCreate_RootClass_Menu
ElseIf Action = "Create" Then
    Call Create_RootClass_Menu
Else
    Call ShowConfig
End If
Response.Write "</body></html>" & vbCrLf

Sub ShowConfig()
    Response.Write "<form method='POST' action='KS.ClassMenu.asp' id='myform' name='myform'>"
    Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"
	Response.Write "  <tr  class='title'>"
    Response.Write "    <td height='22' colspan='6'><strong>顶部栏目全局参数设置</strong> （注：部分特效只对特定的浏览器才有效）</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>选择频道：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write ReturnAllChannel()
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>每行显示数：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='MaxPerLine' type='text' id='MaxPerLine' value='" & MaxPerLine & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>生成文件名：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='JsFileName' type='text' id='JsFileName' value='" & JsFileName & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "  </tr>"
	
    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='6'><strong>顶部栏目菜单参数设置</strong> （注：部分特效只对特定的浏览器才有效）</td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>弹出方式：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_1' id='RCM_Menu_1'>"
    Response.Write "        <option value='1' "
    If RCM_Menu_1 = "1" Then Response.Write " selected"
    Response.Write "        >向左</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_1 = "2" Then Response.Write " selected"
    Response.Write "        >向右</option>"
    Response.Write "        <option value='3' "
    If RCM_Menu_1 = "3" Then Response.Write " selected"
    Response.Write "        >向上</option>"
    Response.Write "        <option value='4' "
    If RCM_Menu_1 = "4" Then Response.Write " selected"
    Response.Write "        >向下</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>横向偏移量：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_2' type='text' id='RCM_Menu_2' value='" & RCM_Menu_2 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>纵向偏移量：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_3' type='text' id='RCM_Menu_3' value='" & RCM_Menu_3 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项边距：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_4' type='text' id='RCM_Menu_4' value='" & RCM_Menu_4 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项间距：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_5' type='text' id='RCM_Menu_5' value='" & RCM_Menu_5 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项左边距：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_6' type='text' id='RCM_Menu_6' value='" & RCM_Menu_6 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项右边距：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_7' type='text' id='RCM_Menu_7' value='" & RCM_Menu_7 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单透明度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_8' type='text' id='RCM_Menu_8' value='" & RCM_Menu_8 & "' size='10'  title='0-100 完全透明-完全不透明'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单其它特效：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_9' type='text' id='RCM_Menu_9' value='" & RCM_Menu_9 & "' size='10' maxlength='200'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单弹出效果▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_10' id='RCM_Menu_10'>"
    Response.Write "        <option value='-1' "
    If RCM_Menu_10 = "-1" Then Response.Write " selected"
    Response.Write "        >无特效</option>"
    Response.Write "        <option value='0' "
    If RCM_Menu_10 = "0" Then Response.Write " selected"
    Response.Write "        >方形收缩</option>"
    Response.Write "        <option value='1' "
    If RCM_Menu_10 = "1" Then Response.Write " selected"
    Response.Write "        >方形扩散</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_10 = "2" Then Response.Write " selected"
    Response.Write "        >圆形收缩</option>"
    Response.Write "        <option value='3' "
    If RCM_Menu_10 = "3" Then Response.Write " selected"
    Response.Write "        >圆形扩散</option>"
    Response.Write "        <option value='4' "
    If RCM_Menu_10 = "4" Then Response.Write " selected"
    Response.Write "        >上拉效果</option>"
    Response.Write "        <option value='5' "
    If RCM_Menu_10 = "5" Then Response.Write " selected"
    Response.Write "        >下拉效果</option>"
    Response.Write "        <option value='6' "
    If RCM_Menu_10 = "6" Then Response.Write " selected"
    Response.Write "        >从左向右</option>"
    Response.Write "        <option value='7' "
    If RCM_Menu_10 = "7" Then Response.Write " selected"
    Response.Write "        >从右向左</option>"
    Response.Write "        <option value='8' "
    If RCM_Menu_10 = "8" Then Response.Write " selected"
    Response.Write "        >左右百叶</option>"
    Response.Write "        <option value='9' "
    If RCM_Menu_10 = "9" Then Response.Write " selected"
    Response.Write "        >上下百叶</option>"
    Response.Write "        <option value='10' "
    If RCM_Menu_10 = "10" Then Response.Write " selected"
    Response.Write "        >左右网格</option>"
    Response.Write "        <option value='11' "
    If RCM_Menu_10 = "11" Then Response.Write " selected"
    Response.Write "        >左右网格</option>"
    Response.Write "        <option value='12' "
    If RCM_Menu_10 = "12" Then Response.Write " selected"
    Response.Write "        >模糊效果</option>"
    Response.Write "        <option value='13' "
    If RCM_Menu_10 = "13" Then Response.Write " selected"
    Response.Write "        >左右关门</option>"
    Response.Write "        <option value='14' "
    If RCM_Menu_10 = "14" Then Response.Write " selected"
    Response.Write "        >左右开门</option>"
    Response.Write "        <option value='15' "
    If RCM_Menu_10 = "15" Then Response.Write " selected"
    Response.Write "        >上下关门</option>"
    Response.Write "        <option value='16' "
    If RCM_Menu_10 = "16" Then Response.Write " selected"
    Response.Write "        >上下开门</option>"
    Response.Write "        <option value='17' "
    If RCM_Menu_10 = "17" Then Response.Write " selected"
    Response.Write "        >左下拉开</option>"
    Response.Write "        <option value='18' "
    If RCM_Menu_10 = "18" Then Response.Write " selected"
    Response.Write "        >左上拉开</option>"
    Response.Write "        <option value='19' "
    If RCM_Menu_10 = "19" Then Response.Write " selected"
    Response.Write "        >右下拉开</option>"
    Response.Write "        <option value='20' "
    If RCM_Menu_10 = "20" Then Response.Write " selected"
    Response.Write "        >右上拉开</option>"
    Response.Write "        <option value='21' "
    If RCM_Menu_10 = "21" Then Response.Write " selected"
    Response.Write "        >上下条纹</option>"
    Response.Write "        <option value='22' "
    If RCM_Menu_10 = "22" Then Response.Write " selected"
    Response.Write "        >左右条纹</option>"
    Response.Write "        <option value='23' "
    If RCM_Menu_10 = "23" Then Response.Write " selected"
    Response.Write "        >随机特效</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单弹出效果▼：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_12' id='RCM_Menu_12'>"
    Response.Write "        <option value='-1' "
    If RCM_Menu_12 = "-1" Then Response.Write " selected"
    Response.Write "        >无特效</option>"
    Response.Write "        <option value='0' "
    If RCM_Menu_12 = "0" Then Response.Write " selected"
    Response.Write "        >方形收缩</option>"
    Response.Write "        <option value='1' "
    If RCM_Menu_12 = "1" Then Response.Write " selected"
    Response.Write "        >方形扩散</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_12 = "2" Then Response.Write " selected"
    Response.Write "        >圆形收缩</option>"
    Response.Write "        <option value='3' "
    If RCM_Menu_12 = "3" Then Response.Write " selected"
    Response.Write "        >圆形扩散</option>"
    Response.Write "        <option value='4' "
    If RCM_Menu_12 = "4" Then Response.Write " selected"
    Response.Write "        >上拉效果</option>"
    Response.Write "        <option value='5' "
    If RCM_Menu_12 = "5" Then Response.Write " selected"
    Response.Write "        >下拉效果</option>"
    Response.Write "        <option value='6' "
    If RCM_Menu_12 = "6" Then Response.Write " selected"
    Response.Write "        >从左向右</option>"
    Response.Write "        <option value='7' "
    If RCM_Menu_12 = "7" Then Response.Write " selected"
    Response.Write "        >从右向左</option>"
    Response.Write "        <option value='8' "
    If RCM_Menu_12 = "8" Then Response.Write " selected"
    Response.Write "        >左右百叶</option>"
    Response.Write "        <option value='9' "
    If RCM_Menu_12 = "9" Then Response.Write " selected"
    Response.Write "        >上下百叶</option>"
    Response.Write "        <option value='10' "
    If RCM_Menu_12 = "10" Then Response.Write " selected"
    Response.Write "        >左右网格</option>"
    Response.Write "        <option value='11' "
    If RCM_Menu_12 = "11" Then Response.Write " selected"
    Response.Write "        >左右网格</option>"
    Response.Write "        <option value='12' "
    If RCM_Menu_12 = "12" Then Response.Write " selected"
    Response.Write "        >模糊效果</option>"
    Response.Write "        <option value='13' "
    If RCM_Menu_12 = "13" Then Response.Write " selected"
    Response.Write "        >左右关门</option>"
    Response.Write "        <option value='14' "
    If RCM_Menu_12 = "14" Then Response.Write " selected"
    Response.Write "        >左右开门</option>"
    Response.Write "        <option value='15' "
    If RCM_Menu_12 = "15" Then Response.Write " selected"
    Response.Write "        >上下关门</option>"
    Response.Write "        <option value='16' "
    If RCM_Menu_12 = "16" Then Response.Write " selected"
    Response.Write "        >上下开门</option>"
    Response.Write "        <option value='17' "
    If RCM_Menu_12 = "17" Then Response.Write " selected"
    Response.Write "        >左下拉开</option>"
    Response.Write "        <option value='18' "
    If RCM_Menu_12 = "18" Then Response.Write " selected"
    Response.Write "        >左上拉开</option>"
    Response.Write "        <option value='19' "
    If RCM_Menu_12 = "19" Then Response.Write " selected"
    Response.Write "        >右下拉开</option>"
    Response.Write "        <option value='20' "
    If RCM_Menu_12 = "20" Then Response.Write " selected"
    Response.Write "        >右上拉开</option>"
    Response.Write "        <option value='21' "
    If RCM_Menu_12 = "21" Then Response.Write " selected"
    Response.Write "        >上下条纹</option>"
    Response.Write "        <option value='22' "
    If RCM_Menu_12 = "22" Then Response.Write " selected"
    Response.Write "        >左右条纹</option>"
    Response.Write "        <option value='23' "
    If RCM_Menu_12 = "23" Then Response.Write " selected"
    Response.Write "        >随机特效</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单弹出效果速度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_13' type='text' id='RCM_Menu_13' value='" & RCM_Menu_13 & "' size='10'  title='速度值：10-100'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单阴影效果：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_14' id='RCM_Menu_14'>"
    Response.Write "        <option value='0' "
    If RCM_Menu_14 = "0" Then Response.Write " selected"
    Response.Write "        >无阴影</option>"
    Response.Write "        <option value='1' "
    If RCM_Menu_14 = "1" Then Response.Write " selected"
    Response.Write "        >简单阴影</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_14 = "2" Then Response.Write " selected"
    Response.Write "        >复杂阴影</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单阴影深度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_15' type='text' id='RCM_Menu_15' value='" & RCM_Menu_15 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单阴影颜色：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_16' type='text' id='RCM_Menu_16' value='" & RCM_Menu_16 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单背景颜色：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_17' type='text' id='RCM_Menu_17' value='" & RCM_Menu_17 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单背景图片：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_18' type='text' id='RCM_Menu_18' value='" & RCM_Menu_18 & "' size='10' maxlength='200' title='只有当菜单项背景颜色设为透明色：transparent 时才有效'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>背景图片平铺模式：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_19' id='RCM_Menu_19'>"
    Response.Write "        <option value='0' "
    If RCM_Menu_19 = "0" Then Response.Write " selected"
    Response.Write "        >不平铺</option>"
    Response.Write "        <option value='1' "
    If RCM_Menu_19 = "1" Then Response.Write " selected"
    Response.Write "        >横向平铺</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_19 = "2" Then Response.Write " selected"
    Response.Write "        >纵向平铺</option>"
    Response.Write "        <option value='3' "
    If RCM_Menu_19 = "3" Then Response.Write " selected"
    Response.Write "        >完全平铺</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单边框类型：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Menu_20' id='RCM_Menu_20'>"
    Response.Write "        <option value='0' "
    If RCM_Menu_20 = "0" Then Response.Write " selected"
    Response.Write "        >无边框</option>"
    Response.Write "        <option value='1' "
    If RCM_Menu_20 = "1" Then Response.Write " selected"
    Response.Write "        >单实线</option>"
    Response.Write "        <option value='2' "
    If RCM_Menu_20 = "2" Then Response.Write " selected"
    Response.Write "        >双实线</option>"
    Response.Write "        <option value='5' "
    If RCM_Menu_20 = "5" Then Response.Write " selected"
    Response.Write "        >凹陷</option>"
    Response.Write "        <option value='6' "
    If RCM_Menu_20 = "6" Then Response.Write " selected"
    Response.Write "        >凸起</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单边框宽度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_21' type='text' id='RCM_Menu_21' value='" & RCM_Menu_21 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单边框颜色：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Menu_22' type='text' id='RCM_Menu_22' value='" & RCM_Menu_22 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "  </tr>"

    Response.Write "  <tr class='title'>"
    Response.Write "    <td height='22' colspan='6'><strong>菜单项参数设置</strong></td>"
    Response.Write "  </tr>"
'    response.write "  <tr class='tdbg'> "
'    response.write "    <td width='130' height='25'><strong>菜单项类型：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <select name='RCM_Item_1' id='RCM_Item_1'>"
'    response.write "        <option value='0' "
'   if RCM_Menu_1="0" then response.write " selected"
'    response.write "        >文本</option>"
'    response.write "        <option value='1' "
'   if RCM_Menu_1="1" then response.write " selected"
'    response.write "        >HTML</option>"
'    response.write "        <option value='2' "
'   if RCM_Menu_1="2" then response.write " selected"
'    response.write "        >图片</option>"
'    response.write "      </select>"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>菜单项名称：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_2' type='text' id='RCM_Item_2' value='" & RCM_Item_2 & "' size='10' >"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>图片文件：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_3' type='text' id='RCM_Item_3' value='" & RCM_Item_3 & "' size='10' >"
'    response.write "    </td>"
'    response.write "  </tr>"
'    response.write "  <tr class='tdbg'> "
'    response.write "    <td width='130' height='25'><strong>鼠标指在菜单项时，图片文件：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_4' type='text' id='RCM_Item_4' value='" & RCM_Item_4 & "' size='10' >"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>图片宽度：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_5' type='text' id='RCM_Item_5' value='" & RCM_Item_5 & "' size='10' >"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>图片高度：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_6' type='text' id='RCM_Item_6' value='" & RCM_Item_6 & "' size='10' >"
'    response.write "    </td>"
'    response.write "  </tr>"
'    response.write "  <tr class='tdbg'> "
'    response.write "    <td width='130' height='25'><strong>图片边框：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_7' type='text' id='RCM_Item_7' value='" & RCM_Item_7 & "' size='10' >"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>链接地址：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_8' type='text' id='RCM_Item_8' value='" & RCM_Item_8 & "' size='10' >"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>链接目标：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_9' type='text' id='RCM_Item_9' value='" & RCM_Item_9 & "' size='10' >"
'    response.write "    </td>"
'    response.write "  </tr>"
'    response.write "  <tr class='tdbg'> "
'    response.write "    <td width='130' height='25'><strong>链接状态栏显示：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_10' type='text' id='RCM_Item_10' value='" & RCM_Item_10 & "' size='10' >"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong>链接地址提示信息：</strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      <input name='RCM_Item_11' type='text' id='RCM_Item_11' value='" & RCM_Item_11 & "' size='10' >"
'    response.write "    </td>"
'    response.write "    <td width='130' height='25'><strong></strong></td>"
'    response.write "    <td width='120'>"
'    response.write "      "
'    response.write "    </td>"
'    response.write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项左图片：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_12' type='text' id='RCM_Item_12' value='" & RCM_Item_12 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项左图片▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_13' type='text' id='RCM_Item_13' value='" & RCM_Item_13 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>左图片宽度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_14' type='text' id='RCM_Item_14' value='" & RCM_Item_14 & "' size='10'  title='0为图像原始宽度'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>左图片高度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_15' type='text' id='RCM_Item_15' value='" & RCM_Item_15 & "' size='10'  title='0为图像原始高度'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>左图片边框大小：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_16' type='text' id='RCM_Item_16' value='" & RCM_Item_16 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项右图片：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_17' type='text' id='RCM_Item_17' value='" & RCM_Item_17 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项右图片▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_18' type='text' id='RCM_Item_18' value='" & RCM_Item_18 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>右图片宽度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_19' type='text' id='RCM_Item_19' value='" & RCM_Item_19 & "' size='10'  title='0为图像原始宽度'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>右图片高度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_20' type='text' id='RCM_Item_20' value='" & RCM_Item_20 & "' size='10'  title='0为图像原始高度'>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>右图片边框大小：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_21' type='text' id='RCM_Item_21' value='" & RCM_Item_21 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>文字水平对齐方式：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_22' id='RCM_Item_22'>"
    Response.Write "        <option value='0' "
    If RCM_Item_22 = "0" Then Response.Write " selected"
    Response.Write "        >左对齐</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_22 = "1" Then Response.Write " selected"
    Response.Write "        >居中</option>"
    Response.Write "        <option value='2' "
    If RCM_Item_22 = "2" Then Response.Write " selected"
    Response.Write "        >右对齐</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>文字垂直对齐方式：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_23' id='RCM_Item_23'>"
    Response.Write "        <option value='0' "
    If RCM_Item_23 = "0" Then Response.Write " selected"
    Response.Write "        >顶部</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_23 = "1" Then Response.Write " selected"
    Response.Write "        >居中</option>"
    Response.Write "        <option value='2' "
    If RCM_Item_23 = "2" Then Response.Write " selected"
    Response.Write "        >底部</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项背景颜色：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_24' type='text' id='RCM_Item_24' value='" & RCM_Item_24 & "' size='10'  title='透明色：transparent'>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>背景颜色是否显示：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_25' id='RCM_Item_25'>"
    Response.Write "        <option value='0' "
    If RCM_Item_25 = "0" Then Response.Write " selected"
    Response.Write "        >显示</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_25 = "1" Then Response.Write " selected"
    Response.Write "        >不显示</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项背景颜色▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_26' type='text' id='RCM_Item_26' value='" & RCM_Item_26 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>背景颜色是否显示▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_27' id='RCM_Item_27'>"
    Response.Write "        <option value='0' "
    If RCM_Item_27 = "0" Then Response.Write " selected"
    Response.Write "        >显示</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_27 = "1" Then Response.Write " selected"
    Response.Write "        >不显示</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项背景图片：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_28' type='text' id='RCM_Item_28' value='" & RCM_Item_28 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项背景图片▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_29' type='text' id='RCM_Item_29' value='" & RCM_Item_29 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>背景图片平铺模式：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_30' id='RCM_Item_30'>"
    Response.Write "        <option value='0' "
    If RCM_Item_30 = "0" Then Response.Write " selected"
    Response.Write "        >不平铺</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_30 = "1" Then Response.Write " selected"
    Response.Write "        >横向平铺</option>"
    Response.Write "        <option value='2' "
    If RCM_Item_30 = "2" Then Response.Write " selected"
    Response.Write "        >纵向平铺</option>"
    Response.Write "        <option value='3' "
    If RCM_Item_30 = "3" Then Response.Write " selected"
    Response.Write "        >完全平铺</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项边框类型：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='RCM_Item_32' id='RCM_Item_32'>"
    Response.Write "        <option value='0' "
    If RCM_Item_32 = "0" Then Response.Write " selected"
    Response.Write "        >无边框</option>"
    Response.Write "        <option value='1' "
    If RCM_Item_32 = "1" Then Response.Write " selected"
    Response.Write "        >单实线</option>"
    Response.Write "        <option value='2' "
    If RCM_Item_32 = "2" Then Response.Write " selected"
    Response.Write "        >双实线</option>"
    Response.Write "        <option value='5' "
    If RCM_Item_32 = "5" Then Response.Write " selected"
    Response.Write "        >凹陷</option>"
    Response.Write "        <option value='6' "
    If RCM_Item_32 = "6" Then Response.Write " selected"
    Response.Write "        >凸起</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项边框宽度：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_33' type='text' id='RCM_Item_33' value='" & RCM_Item_33 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项边框颜色：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_34' type='text' id='RCM_Item_34' value='" & RCM_Item_34 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项边框颜色▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_35' type='text' id='RCM_Item_35' value='" & RCM_Item_35 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项文字颜色：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_36' type='text' id='RCM_Item_36' value='" & RCM_Item_36 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'> "
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项文字颜色▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <input name='RCM_Item_37' type='text' id='RCM_Item_37' value='" & RCM_Item_37 & "' size='10' >"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项文字字体：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='FontName_RCM_Item_38' id='FontName_RCM_Item_38'>"
    Response.Write "        <option value='宋体' "
    If FontName_RCM_Item_38 = "宋体" Then Response.Write " selected"
    Response.Write "        >宋体</option>"
    Response.Write "        <option value=""黑体"" "
    If FontName_RCM_Item_38 = "黑体" Then Response.Write " selected"
    Response.Write "        >黑体</option>"
    Response.Write "        <option value=""楷体"" "
    If FontName_RCM_Item_38 = "楷体" Then Response.Write " selected"
    Response.Write "        >楷体</option>"
    Response.Write "        <option value=""仿宋"" "
    If FontName_RCM_Item_38 = "仿宋" Then Response.Write " selected"
    Response.Write "        >仿宋</option>"
    Response.Write "        <option value=""隶书"" "
    If FontName_RCM_Item_38 = "隶书" Then Response.Write " selected"
    Response.Write "        >隶书</option>"
    Response.Write "        <option value=""幼圆"" "
    If FontName_RCM_Item_38 = "幼圆" Then Response.Write " selected"
    Response.Write "        >幼圆</option>"
    Response.Write "        <option value=""Arial"" "
    If FontName_RCM_Item_38 = "Arial" Then Response.Write " selected"
    Response.Write "        >Arial</option>"
    Response.Write "        <option value=""Arial Black"" "
    If FontName_RCM_Item_38 = "Arial Black" Then Response.Write " selected"
    Response.Write "        >Arial Black</option>"
    Response.Write "        <option value=""Arial Narrow"" "
    If FontName_RCM_Item_38 = "Arial Narrow" Then Response.Write " selected"
    Response.Write "        >Arial Narrow</option>"
    Response.Write "        <option value=""Brush ScriptMT"" "
    If FontName_RCM_Item_38 = "Brush ScriptMT" Then Response.Write " selected"
    Response.Write "        >Brush Script MT</option>"
    Response.Write "        <option value=""Century Gothic"" "
    If FontName_RCM_Item_38 = "Century Gothic" Then Response.Write " selected"
    Response.Write "        >Century Gothic</option>"
    Response.Write "        <option value=""Comic Sans MS"" "
    If FontName_RCM_Item_38 = "Comic Sans MS" Then Response.Write " selected"
    Response.Write "        >Comic Sans MS</option>"
    Response.Write "        <option value=""Courier"" "
    If FontName_RCM_Item_38 = "Courier" Then Response.Write " selected"
    Response.Write "        >Courier</option>"
    Response.Write "        <option value=""Courier New"" "
    If FontName_RCM_Item_38 = "Courier New" Then Response.Write " selected"
    Response.Write "        >Courier New</option>"
    Response.Write "        <option value=""MS Sans Serif"" "
    If FontName_RCM_Item_38 = "MS Sans Serif" Then Response.Write " selected"
    Response.Write "        >MS Sans Serif</option>"
    Response.Write "        <option value=""Script"" "
    If FontName_RCM_Item_38 = "Script" Then Response.Write " selected"
    Response.Write "        >Script</option>"
    Response.Write "        <option value=""System"" "
    If FontName_RCM_Item_38 = "System" Then Response.Write " selected"
    Response.Write "        >System</option>"
    Response.Write "        <option value=""Times New Roman"" "
    If FontName_RCM_Item_38 = "Times New Roman" Then Response.Write " selected"
    Response.Write "        >Times New Roman</option>"
    Response.Write "        <option value=""Verdana"" "
    If FontName_RCM_Item_38 = "Verdana" Then Response.Write " selected"
    Response.Write "        >Verdana</option>"
    Response.Write "        <option value=""WideLatin"" "
    If FontName_RCM_Item_38 = "WideLatin" Then Response.Write " selected"
    Response.Write "        >Wide Latin</option>"
    Response.Write "        <option value=""Wingdings"" "
    If FontName_RCM_Item_38 = "Wingdings" Then Response.Write " selected"
    Response.Write "        >Wingdings</option>"
    Response.Write "      </select>"
    Response.Write "      <select name = 'FontSize_RCM_Item_38' id='FontSize_RCM_Item_38'>"
    Response.Write "        <option value=""9pt"" "
    If FontSize_RCM_Item_38 = "9pt" Then Response.Write " selected"
    Response.Write "        >9pt</option>"
    Response.Write "        <option value=""10pt"" "
    If FontSize_RCM_Item_38 = "10pt" Then Response.Write " selected"
    Response.Write "        >10pt</option>"
    Response.Write "        <option value=""12pt"" "
    If FontSize_RCM_Item_38 = "12pt" Then Response.Write " selected"
    Response.Write "        >12pt</option>"
    Response.Write "        <option value=""14pt"" "
    If FontSize_RCM_Item_38 = "14pt" Then Response.Write " selected"
    Response.Write "        >14pt</option>"
    Response.Write "        <option value=""16pt"" "
    If FontSize_RCM_Item_38 = "16pt" Then Response.Write " selected"
    Response.Write "        >16pt</option>"
    Response.Write "        <option value=""18pt"" "
    If FontSize_RCM_Item_38 = "18pt" Then Response.Write " selected"
    Response.Write "        >18pt</option>"
    Response.Write "        <option value=""24pt"" "
    If FontSize_RCM_Item_38 = "24pt" Then Response.Write " selected"
    Response.Write "        >24pt</option>"
    Response.Write "        <option value=""36pt"" "
    If FontSize_RCM_Item_38 = "36pt" Then Response.Write " selected"
    Response.Write "        >36pt</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "    <td width='130' height='25' class='clefttitle'><strong>菜单项文字字体▲：</strong></td>"
    Response.Write "    <td width='120'>"
    Response.Write "      <select name='FontName_RCM_Item_39' id='FontName_RCM_Item_39'>"
    Response.Write "        <option value='宋体' "
    If FontName_RCM_Item_39 = "宋体" Then Response.Write " selected"
    Response.Write "        >宋体</option>"
    Response.Write "        <option value=""黑体"" "
    If FontName_RCM_Item_39 = "黑体" Then Response.Write " selected"
    Response.Write "        >黑体</option>"
    Response.Write "        <option value=""楷体"" "
    If FontName_RCM_Item_39 = "楷体" Then Response.Write " selected"
    Response.Write "        >楷体</option>"
    Response.Write "        <option value=""仿宋"" "
    If FontName_RCM_Item_39 = "仿宋" Then Response.Write " selected"
    Response.Write "        >仿宋</option>"
    Response.Write "        <option value=""隶书"" "
    If FontName_RCM_Item_39 = "隶书" Then Response.Write " selected"
    Response.Write "        >隶书</option>"
    Response.Write "        <option value=""幼圆"" "
    If FontName_RCM_Item_39 = "幼圆" Then Response.Write " selected"
    Response.Write "        >幼圆</option>"
    Response.Write "        <option value=""Arial"" "
    If FontName_RCM_Item_39 = "Arial" Then Response.Write " selected"
    Response.Write "        >Arial</option>"
    Response.Write "        <option value=""Arial Black"" "
    If FontName_RCM_Item_39 = "Arial Black" Then Response.Write " selected"
    Response.Write "        >Arial Black</option>"
    Response.Write "        <option value=""Arial Narrow"" "
    If FontName_RCM_Item_39 = "Arial Narrow" Then Response.Write " selected"
    Response.Write "        >Arial Narrow</option>"
    Response.Write "        <option value=""Brush ScriptMT"" "
    If FontName_RCM_Item_39 = "Brush ScriptMT" Then Response.Write " selected"
    Response.Write "        >Brush Script MT</option>"
    Response.Write "        <option value=""Century Gothic"" "
    If FontName_RCM_Item_39 = "Century Gothic" Then Response.Write " selected"
    Response.Write "        >Century Gothic</option>"
    Response.Write "        <option value=""Comic Sans MS"" "
    If FontName_RCM_Item_39 = "Comic Sans MS" Then Response.Write " selected"
    Response.Write "        >Comic Sans MS</option>"
    Response.Write "        <option value=""Courier"" "
    If FontName_RCM_Item_39 = "Courier" Then Response.Write " selected"
    Response.Write "        >Courier</option>"
    Response.Write "        <option value=""Courier New"" "
    If FontName_RCM_Item_39 = "Courier New" Then Response.Write " selected"
    Response.Write "        >Courier New</option>"
    Response.Write "        <option value=""MS Sans Serif"" "
    If FontName_RCM_Item_39 = "MS Sans Serif" Then Response.Write " selected"
    Response.Write "        >MS Sans Serif</option>"
    Response.Write "        <option value=""Script"" "
    If FontName_RCM_Item_39 = "Script" Then Response.Write " selected"
    Response.Write "        >Script</option>"
    Response.Write "        <option value=""System"" "
    If FontName_RCM_Item_39 = "System" Then Response.Write " selected"
    Response.Write "        >System</option>"
    Response.Write "        <option value=""Times New Roman"" "
    If FontName_RCM_Item_39 = "Times New Roman" Then Response.Write " selected"
    Response.Write "        >Times New Roman</option>"
    Response.Write "        <option value=""Verdana"" "
    If FontName_RCM_Item_39 = "Verdana" Then Response.Write " selected"
    Response.Write "        >Verdana</option>"
    Response.Write "        <option value=""WideLatin"" "
    If FontName_RCM_Item_39 = "WideLatin" Then Response.Write " selected"
    Response.Write "        >Wide Latin</option>"
    Response.Write "        <option value=""Wingdings"" "
    If FontName_RCM_Item_39 = "Wingdings" Then Response.Write " selected"
    Response.Write "        >Wingdings</option>"
    Response.Write "      </select>"
    Response.Write "      <select name = 'FontSize_RCM_Item_39' id='FontSize_RCM_Item_39'>"
    Response.Write "        <option value=""9pt"" "
    If FontSize_RCM_Item_39 = "9pt" Then Response.Write " selected"
    Response.Write "        >9pt</option>"
    Response.Write "        <option value=""10pt"" "
    If FontSize_RCM_Item_39 = "10pt" Then Response.Write " selected"
    Response.Write "        >10pt</option>"
    Response.Write "        <option value=""12pt"" "
    If FontSize_RCM_Item_39 = "12pt" Then Response.Write " selected"
    Response.Write "        >12pt</option>"
    Response.Write "        <option value=""14pt"" "
    If FontSize_RCM_Item_39 = "14pt" Then Response.Write " selected"
    Response.Write "        >14pt</option>"
    Response.Write "        <option value=""16pt"" "
    If FontSize_RCM_Item_39 = "16pt" Then Response.Write " selected"
    Response.Write "        >16pt</option>"
    Response.Write "        <option value=""18pt"" "
    If FontSize_RCM_Item_39 = "18pt" Then Response.Write " selected"
    Response.Write "        >18pt</option>"
    Response.Write "        <option value=""24pt"" "
    If FontSize_RCM_Item_39 = "24pt" Then Response.Write " selected"
    Response.Write "        >24pt</option>"
    Response.Write "        <option value=""36pt"" "
    If FontSize_RCM_Item_39 = "36pt" Then Response.Write " selected"
    Response.Write "        >36pt</option>"
    Response.Write "      </select>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='40' colspan='6' align='center'>"
    Response.Write "      <input name='Action' type='hidden' id='Action' value='SaveConfig'>"
    Response.Write "      <input name='cmdSave' type='submit' id='cmdSave' value=' 保存设置 '  class='button'"
    If ObjInstalled = False Then Response.Write " disabled"
    Response.Write "      >"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
    Response.Write "</form>"
End Sub

Sub SaveConfig()
    If ObjInstalled = False Then
        Response.Write "<script>alert('你的服务器不支持 FSO(Scripting.FileSystemObject)!');</script>"
        Exit Sub
    End If

	Dim Param
    Param= "<" & "%" & vbCrLf
    Param=Param &  "'全局参数设置" & vbCrLf
    Param=Param &  "Const ChannelID=" & Chr(34) & Trim(request("ChannelID")) & Chr(34) & "      '模块ID" & vbCrLf
    Param=Param &  "Const MaxPerLine=" & Chr(34) & KS.ChkClng(Trim(request("MaxPerLine"))) & Chr(34) & "     '每行显示数量" & vbCrLf
    Param=Param &  "Const JsFileName=" & Chr(34) & FilterString(Trim(request("JsFileName"))) & Chr(34) & "      '生成的JS文件名" & vbCrLf
	Param=Param &  "" & vbCrLf
    Param=Param &  "'菜单显示参数设置" & vbCrLf
    Param=Param &  "Const RCM_Menu_1=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_1"))) & Chr(34) & "      '菜单弹出方式 1：左  2：右  3：上  4：下" & vbCrLf
    Param=Param &  "Const RCM_Menu_2=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_2"))) & Chr(34) & "      '菜单弹出横向偏移量" & vbCrLf
    Param=Param &  "Const RCM_Menu_3=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_3"))) & Chr(34) & "      '菜单弹出纵向偏移量" & vbCrLf
    Param=Param &  "Const RCM_Menu_4=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_4"))) & Chr(34) & "      '菜单项边距" & vbCrLf
    Param=Param &  "Const RCM_Menu_5=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_5"))) & Chr(34) & "      '菜单项间距" & vbCrLf
    Param=Param &  "Const RCM_Menu_6=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_6"))) & Chr(34) & "      '菜单项左边距" & vbCrLf
    Param=Param &  "Const RCM_Menu_7=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_7"))) & Chr(34) & "      '菜单项右边距" & vbCrLf
    Param=Param &  "Const RCM_Menu_8=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_8"))) & Chr(34) & "      '菜单透明度         0-100 完全透明-完全不透明" & vbCrLf
    Param=Param &  "Const RCM_Menu_9=" & Chr(34) & FilterString(Trim(request("RCM_Menu_9"))) & Chr(34) & "      '其它特效" & vbCrLf
    Param=Param &  "Const RCM_Menu_10=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_10"))) & Chr(34) & "        '鼠标指在菜单项时，菜单弹出效果" & vbCrLf
    Param=Param &  "Const RCM_Menu_11=" & Chr(34) & FilterString(Trim(request("RCM_Menu_11"))) & Chr(34) & "        '其它特效" & vbCrLf
    Param=Param &  "Const RCM_Menu_12=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_12"))) & Chr(34) & "        '鼠标移出菜单项时，菜单弹出效果" & vbCrLf
    Param=Param &  "Const RCM_Menu_13=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_13"))) & Chr(34) & "        '菜单弹出效果速度  10-100" & vbCrLf
    Param=Param &  "Const RCM_Menu_14=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_14"))) & Chr(34) & "        '弹出菜单阴影效果 0：none  1：simple  2：complex" & vbCrLf
    Param=Param &  "Const RCM_Menu_15=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_15"))) & Chr(34) & "        '弹出菜单阴影深度" & vbCrLf
    Param=Param &  "Const RCM_Menu_16=" & Chr(34) & FilterString(Trim(request("RCM_Menu_16"))) & Chr(34) & "        '弹出菜单阴影颜色" & vbCrLf
    Param=Param &  "Const RCM_Menu_17=" & Chr(34) & FilterString(Trim(request("RCM_Menu_17"))) & Chr(34) & "        '弹出菜单背景颜色" & vbCrLf
    Param=Param &  "Const RCM_Menu_18=" & Chr(34) & FilterString(Trim(request("RCM_Menu_18"))) & Chr(34) & "        '弹出菜单背景图片，只有当菜单项背景颜色设为透明色：transparent 时才有效" & vbCrLf
    Param=Param &  "Const RCM_Menu_19=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_19"))) & Chr(34) & "        '弹出菜单背景图片平铺模式。 0：不平铺  1：横向平铺  2：纵向平铺  3：完全平铺" & vbCrLf
    Param=Param &  "Const RCM_Menu_20=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_20"))) & Chr(34) & "        '弹出菜单边框类型 0：无边框  1：单实线  2：双实线  5：凹陷  6：凸起" & vbCrLf
    Param=Param &  "Const RCM_Menu_21=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Menu_21"))) & Chr(34) & "        '弹出菜单边框宽度" & vbCrLf
    Param=Param &  "Const RCM_Menu_22=" & Chr(34) & FilterString(Trim(request("RCM_Menu_22"))) & Chr(34) & "        '弹出菜单边框颜色" & vbCrLf
    Param=Param &  "Const RCM_Menu_23=" & Chr(34) & "#ffffff" & Chr(34) & "" & vbCrLf

    Param=Param &  "" & vbCrLf
    Param=Param &  "'菜单项参数设置" & vbCrLf
    Param=Param &  "Const RCM_Item_1=" & Chr(34) & "0" & Chr(34) & "      '菜单项类型  0--Txt  1--Html  2--Image" & vbCrLf
    Param=Param &  "Const RCM_Item_2=" & Chr(34) & "" & Chr(34) & "       '菜单项名称" & vbCrLf
    Param=Param &  "Const RCM_Item_3=" & Chr(34) & "" & Chr(34) & "       '菜单项为Image，图片文件" & vbCrLf
    Param=Param &  "Const RCM_Item_4=" & Chr(34) & "" & Chr(34) & "       '菜单项为Image，鼠标指在菜单项时，图片文件。" & vbCrLf
    Param=Param &  "Const RCM_Item_5=" & Chr(34) & "-1" & Chr(34) & "     '菜单项为Image，图片宽度" & vbCrLf
    Param=Param &  "Const RCM_Item_6=" & Chr(34) & "-1" & Chr(34) & "     '菜单项为Image，图片高度" & vbCrLf
    Param=Param &  "Const RCM_Item_7=" & Chr(34) & "0" & Chr(34) & "      '菜单项为Image，图片边框" & vbCrLf
    Param=Param &  "Const RCM_Item_8=" & Chr(34) & "" & Chr(34) & "       '菜单项链接地址" & vbCrLf
    Param=Param &  "Const RCM_Item_9=" & Chr(34) & "" & Chr(34) & "       '菜单项链接目标 如：_self  _blank" & vbCrLf
    Param=Param &  "Const RCM_Item_10=" & Chr(34) & "" & Chr(34) & "      '菜单项链接状态栏显示" & vbCrLf
    Param=Param &  "Const RCM_Item_11=" & Chr(34) & "" & Chr(34) & "      '菜单项链接地址提示信息" & vbCrLf
    Param=Param &  "Const RCM_Item_12=" & Chr(34) & FilterString(Trim(request("RCM_Item_12"))) & Chr(34) & "        '菜单项左图片" & vbCrLf
    Param=Param &  "Const RCM_Item_13=" & Chr(34) & FilterString(Trim(request("RCM_Item_13"))) & Chr(34) & "        '鼠标指在菜单项时，菜单项左图片" & vbCrLf
    Param=Param &  "Const RCM_Item_14=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Item_14"))) & Chr(34) & "        '菜单项左图片宽度，0为图像文件原始值" & vbCrLf
    Param=Param &  "Const RCM_Item_15=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Item_15"))) & Chr(34) & "        '菜单项左图片高度，0为图像文件原始值" & vbCrLf
    Param=Param &  "Const RCM_Item_16=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Item_16"))) & Chr(34) & "        '菜单项左图片边框大小" & vbCrLf
    Param=Param &  "Const RCM_Item_17=" & Chr(34) & FilterString(Trim(request("RCM_Item_17"))) & Chr(34) & "        '菜单项右图片。如：arrow_r.gif" & vbCrLf
    Param=Param &  "Const RCM_Item_18=" & Chr(34) & FilterString(Trim(request("RCM_Item_18"))) & Chr(34) & "        '鼠标指在菜单项时，菜单项右图片。如：arrow_w.gif" & vbCrLf
    Param=Param &  "Const RCM_Item_19=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Item_19"))) & Chr(34) & "        '菜单项右图片宽度，0为图像文件原始值" & vbCrLf
    Param=Param &  "Const RCM_Item_20=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Item_20"))) & Chr(34) & "        '菜单项右图片高度，0为图像文件原始值" & vbCrLf
    Param=Param &  "Const RCM_Item_21=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Item_21"))) & Chr(34) & "        '菜单项右图片边框大小" & vbCrLf
    Param=Param &  "Const RCM_Item_22=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Item_22"))) & Chr(34) & "        '菜单项文字水平对齐方式  0：左对齐  1：居中  2：右对齐" & vbCrLf
    Param=Param &  "Const RCM_Item_23=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Item_23"))) & Chr(34) & "        '菜单项文字垂直对齐方式  0：顶部  1：居中  2：底部" & vbCrLf
    Param=Param &  "Const RCM_Item_24=" & Chr(34) & FilterString(Trim(request("RCM_Item_24"))) & Chr(34) & "        '菜单项背景颜色  透明色：'transparent'" & vbCrLf
    Param=Param &  "Const RCM_Item_25=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Item_25"))) & Chr(34) & "        '菜单项背景颜色是否显示  0：显示  其它：不显示" & vbCrLf
    Param=Param &  "Const RCM_Item_26=" & Chr(34) & FilterString(Trim(request("RCM_Item_26"))) & Chr(34) & "        '鼠标指在菜单项时，菜单项背景颜色" & vbCrLf
    Param=Param &  "Const RCM_Item_27=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Item_27"))) & Chr(34) & "        '鼠标指在菜单项时，菜单项背景颜色是否显示。  0：显示  其它：不显示" & vbCrLf
    Param=Param &  "Const RCM_Item_28=" & Chr(34) & FilterString(Trim(request("RCM_Item_28"))) & Chr(34) & "        '菜单项背景图片" & vbCrLf
    Param=Param &  "Const RCM_Item_29=" & Chr(34) & FilterString(Trim(request("RCM_Item_29"))) & Chr(34) & "        '鼠标指在菜单项时，菜单项背景图片" & vbCrLf
    Param=Param &  "Const RCM_Item_30=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Item_30"))) & Chr(34) & "        '菜单项背景图片平铺模式。 0：不平铺  1：横向平铺  2：纵向平铺  3：完全平铺" & vbCrLf
    Param=Param &  "Const RCM_Item_31=" & Chr(34) & "3" & Chr(34) & "     '鼠标指在菜单项时，菜单项背景图片平铺模式。0-3" & vbCrLf
    Param=Param &  "Const RCM_Item_32=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Item_32"))) & Chr(34) & "        '菜单项边框类型 0：无边框  1：单实线  2：双实线  5：凹陷  6：凸起" & vbCrLf
    Param=Param &  "Const RCM_Item_33=" & Chr(34) & KS.ChkClng(Trim(request("RCM_Item_33"))) & Chr(34) & "        '菜单项边框宽度" & vbCrLf
    Param=Param &  "Const RCM_Item_34=" & Chr(34) & FilterString(Trim(request("RCM_Item_34"))) & Chr(34) & "        '菜单项边框颜色" & vbCrLf
    Param=Param &  "Const RCM_Item_35=" & Chr(34) & FilterString(Trim(request("RCM_Item_35"))) & Chr(34) & "        '鼠标指在菜单项时，菜单项边框颜色" & vbCrLf
    Param=Param &  "Const RCM_Item_36=" & Chr(34) & FilterString(Trim(request("RCM_Item_36"))) & Chr(34) & "        '菜单项文字颜色" & vbCrLf
    Param=Param &  "Const RCM_Item_37=" & Chr(34) & FilterString(Trim(request("RCM_Item_37"))) & Chr(34) & "        '鼠标指在菜单项时，菜单项文字颜色" & vbCrLf
    Param=Param &  "Const FontSize_RCM_Item_38=" & Chr(34) & FilterString(Trim(request("FontSize_RCM_Item_38"))) & Chr(34) & "        '菜单项文字大小" & vbCrLf
    Param=Param &  "Const FontName_RCM_Item_38=" & Chr(34) & FilterString(Trim(request("FontName_RCM_Item_38"))) & Chr(34) & "        '菜单项文字字体" & vbCrLf
    Param=Param &  "Const FontSize_RCM_Item_39=" & Chr(34) & FilterString(Trim(request("FontSize_RCM_Item_39"))) & Chr(34) & "        '鼠标指在菜单项时,菜单项文字大小" & vbCrLf
    Param=Param &  "Const FontName_RCM_Item_39=" & Chr(34) & FilterString(Trim(request("FontName_RCM_Item_39"))) & Chr(34) & "        '鼠标指在菜单项时,菜单项文字字体" & vbCrLf
    Param=Param &  "%" & ">"
   
    Call KS.WriteTOFile(strInstallDir & KS.Setting(89) & "KS.ClassMenuParam.asp", Param)
    Response.Write "<script>alert('顶部栏目菜单参数设置成功！');location.href='" & ComeUrl & "';</script>"
End Sub

Sub ShowCreate_RootClass_Menu()
    Response.Write "<br><table width='100%' border='0' cellspacing='1' cellpadding='2' class='ctable'>"
    Response.Write "  <tr class='sort'>"
    Response.Write "    <td height='22' align='center'><strong> 生 成 顶 部 栏 目 菜 单 </strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='150'>"
    Response.Write "<form name='myform' method='post' action='KS.ClassMenu.asp'>"
    Response.Write "<p align='center'>此操作将根据顶部栏目菜单参数设置中设置的参数生成自定义的菜单。</p>"
    Response.Write "<p align='center'><input name='Action' type='hidden' id='Action' value='Create'>"
    Response.Write "<input name='ChannelID' type='hidden' id='ChannelID' value='" & ChannelID & "'>"
    Response.Write "<input type='submit' name='Submit' value=' 生成顶部栏目菜单 ' class='button'></p>"
    Response.Write "</form>"
    Response.Write "    </td>"
    Response.Write "  </tr>"
    Response.Write "</table>"
End Sub

Sub Create_RootClass_Menu()
    strTopMenu = GetRootClass_Menu()
	If KS.Setting(97)="0" Then strTopMenu=Replace(strTopMenu,KS.GetDomain,KS.Setting(3))
	
	Call KS.WriteTOFile(KS.Setting(3) & KS.Setting(93) & JsFileName, strTopMenu)
	Response.Write "<br><table width='100%' border='0' cellspacing='1' cellpadding='2' class='ctable'>"
    Response.Write "  <tr class='sort'>"
    Response.Write "    <td height='22' align='center'><strong> 生 成 顶 部 栏 目 菜 单 </strong></td>"
    Response.Write "  </tr>"
    Response.Write "  <tr class='tdbg'>"
    Response.Write "    <td height='150'>"
	Response.Write "<br><p align='center'><font color=red><b>恭喜您！顶部菜单成功生成,请按以下提示完成最好操作。</b></font></p>"
    Response.Write "<p><b>第一步：将以下代码复制到您要调用模板的&lt;head&gt;&lt;/head&gt;之间。</b></p>"
	Response.Write "<input name='s1' value='&lt;script language=&quot;javascript&quot; type=&quot;text/javascript&quot; src=&quot;" & strInstallDir & "ks_inc/stm31.js&quot;&gt;&lt;/script&gt;' size='80'>&nbsp;<input class=""button"" onClick=""jm_cc('s1')"" type=""button"" value=""复制到剪贴板"" name=""button"">"
    Response.Write "<p><b>第二步：将以下代码复制到在模板里要显示的地方。</b></p>"
	Response.Write "<input name='s2' value='&lt;script language=&quot;javascript&quot; type=&quot;text/javascript&quot; src=&quot;" & KS.Setting(3) & KS.Setting(93) & JsFileName & "&quot;&gt;&lt;/script&gt;' size='80'>&nbsp;<input class=""button"" onClick=""jm_cc('s2')"" type=""button"" value=""复制到剪贴板"" name=""button1"">"
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

'=================================================
'函数名：GetRootClass_Menu
'作  用：得到栏目无级下拉菜单效果的HTML代码
'参  数：无
'返回值：栏目无级下拉菜单效果的HTML代码
'=================================================
Function GetRootClass_Menu()
    Dim Class_MenuTitle, strJS
    pNum = 1
    pNum2 = 0
    strJS = stm_bm() & vbCrLf
    strJS = strJS & stm_bp_h() & vbCrLf
    strJS = strJS & stm_ai() & vbCrLf
    
    strJS = strJS & stm_aix("p0i1", "p0i0", "网站首页", strInstallDir & "Index.asp", "_self", "", False) & vbCrLf
    strJS = strJS & stm_aix("p0i2", "p0i0", "|", "", "_self", "", False) & vbCrLf

    Dim sqlRoot, rsRoot, j,Param
	If Len(Channelid)>4 Then
	     Param=" and a.tn='" & ChannelID & "'"
	Else
	 if ChannelID<>0 Then 
	  Param=" and TN='0' And A.ChannelID=" & KS.ChkClng(ChannelID)
	 else
	  Param=" and TN='0'"
	 end if
	End If
	
	sqlRoot = "Select ID,FolderName,TN,FolderOrder,ClassType From KS_Class a inner join KS_Channel b on a.channelid=b.channelid Where  B.ChannelStatus=1 AND TopFlag=1" & Param & " And DelTF=0 Order By root,folderorder"
    Set rsRoot = KS.InitialObject("ADODB.Recordset")
    rsRoot.open sqlRoot, Conn, 1, 1
    If Not (rsRoot.bof And rsRoot.EOF) Then
        j = 3
        Do While Not rsRoot.EOF
		     If rsRoot("ClassType")="2" Then
             OpenTyKS_Class = "_blank"
			 Else
             OpenTyKS_Class = "_self"
			 End If
             Class_MenuTitle = ""
			  if not isnumeric(mid(rsRoot(0),3,3)) then
                strJS = strJS & stm_aix("p0i" & j & "", "p0i0", rsRoot(1),rsRoot(2), OpenTyKS_Class, Class_MenuTitle, False) & vbCrLf
			  Else
                strJS = strJS & stm_aix("p0i" & j & "", "p0i0", rsRoot(1), KS.GetFolderPath(rsRoot(0)), OpenTyKS_Class, Class_MenuTitle, False) & vbCrLf
                If Not Conn.Execute("Select ID From KS_Class Where TN='" & rsRoot(0) & "'").Eof Then
                    strJS = strJS & GetClassMenu(rsRoot(0), 0)
                End If
			  End If

            strJS = strJS & stm_aix("p0i2", "p0i0", "|", "", "_self", "", False) & vbCrLf
            j = j + 1
            rsRoot.movenext
            If (j - 2) Mod MaxPerLine = 0 And Not rsRoot.EOF Then
                strJS = strJS & "stm_em();" & vbCrLf
                strJS = strJS & stm_bm() & vbCrLf
                strJS = strJS & stm_bp_h() & vbCrLf
                strJS = strJS & stm_ai() & vbCrLf
            End If
        Loop
    End If
    rsRoot.Close
    Set rsRoot = Nothing
    strJS = strJS & "stm_em();" & vbCrLf

    GetRootClass_Menu = strJS
End Function

Function GetClassMenu(ID, ShowType)
    Dim sqlClass, rsClass, Sub_MenuTitle, k, strJS
    strJS = ""
    If pNum = 1 Then
        strJS = strJS & stm_bp_v("p" & pNum & "") & vbCrLf
    Else
        strJS = strJS & stm_bpx("p" & pNum & "", "p" & pNum2 & "", ShowType) & vbCrLf
    End If
    
    k = 0
    sqlClass = "select * from KS_Class where TN='" & ID & "' and topflag=1 order by root,folderorder"
    Set rsClass = KS.InitialObject("ADODB.Recordset")
    rsClass.open sqlClass, Conn, 1, 1
    Do While Not rsClass.EOF
		     If rsClass("ClassType")="2" Then
             OpenTyKS_Class = "_blank"
			 Else
             OpenTyKS_Class = "_self"
			 End If
            Sub_MenuTitle = ""
            If Not Conn.Execute("Select ID From KS_Class Where TN='" & rsClass("ID") & "'").Eof Then
                strJS = strJS & stm_aix("p" & pNum & "i" & k & "", "p" & pNum2 & "i0", rsClass("FolderName"), KS.GetFolderPath(rsClass("ID")), OpenTyKS_Class, Sub_MenuTitle, True) & vbCrLf
                pNum = pNum + 1
                pNum2 = pNum2 + 1
                strJS = strJS & GetClassMenu(rsClass("ID"), 1)
            Else
                
                strJS = strJS & stm_aix("p" & pNum & "i" & k & "", "p" & pNum2 & "i0", rsClass("FolderName"), KS.GetFolderPath(rsClass("ID")), OpenTyKS_Class, Sub_MenuTitle, False) & vbCrLf
            End If
        k = k + 1
        rsClass.movenext
    Loop
    rsClass.Close
    Set rsClass = Nothing
    strJS = strJS & "stm_ep();" & vbCrLf

    GetClassMenu = strJS
End Function

Function stm_bm()
    stm_bm = "stm_bm(['uueoehr',400,'','" & strInstallDir & "images/default/blank.gif',0,'','',0,0,0,0,0,1,0,0]);"
End Function

Function stm_bp_h()
    stm_bp_h = "stm_bp('p0',[0,4,0,0,2,2,0,0," & RCM_Menu_8 & ",'" & RCM_Menu_9 & "'," & RCM_Menu_10 & ",'" & RCM_Menu_11 & "'," & RCM_Menu_12 & "," & RCM_Menu_13 & ",0,0,'#000000','transparent','',3,0,0,'#000000']);"
End Function

Function stm_bp_v(bpID)
    stm_bp_v = "stm_bp('" & bpID & "',[1," & RCM_Menu_1 & "," & RCM_Menu_2 & "," & RCM_Menu_3 & "," & RCM_Menu_4 & "," & RCM_Menu_5 & "," & RCM_Menu_6 & "," & RCM_Menu_7 & "," & RCM_Menu_8 & ",'" & RCM_Menu_9 & "'," & RCM_Menu_10 & ",'" & RCM_Menu_11 & "'," & RCM_Menu_12 & "," & RCM_Menu_13 & "," & RCM_Menu_14 & "," & RCM_Menu_15 & ",'" & RCM_Menu_16 & "','" & RCM_Menu_17 & "','" & RCM_Menu_18 & "'," & RCM_Menu_19 & "," & RCM_Menu_20 & "," & RCM_Menu_21 & ",'" & RCM_Menu_22 & "']);"
End Function

Function stm_bpx(bpOID, bpTID, bpType)
    If bpType = 0 Then
        stm_bpx = "stm_bpx('" & bpOID & "','" & bpTID & "',[1," & RCM_Menu_1 & "," & RCM_Menu_2 & "," & RCM_Menu_3 & "," & RCM_Menu_4 & "," & RCM_Menu_5 & "," & RCM_Menu_6 & "," & RCM_Menu_7 & "," & RCM_Menu_8 & ",'" & RCM_Menu_9 & "'," & RCM_Menu_10 & ",'" & RCM_Menu_11 & "'," & RCM_Menu_12 & "," & RCM_Menu_13 & "," & RCM_Menu_14 & "," & RCM_Menu_15 & ",'" & RCM_Menu_16 & "','" & RCM_Menu_17 & "','" & RCM_Menu_18 & "'," & RCM_Menu_19 & "," & RCM_Menu_20 & "," & RCM_Menu_21 & ",'" & RCM_Menu_22 & "']);"
    Else
        stm_bpx = "stm_bpx('" & bpOID & "','" & bpTID & "',[1,2,-2,-3," & RCM_Menu_4 & "," & RCM_Menu_5 & ",0," & RCM_Menu_7 & "," & RCM_Menu_8 & ",'" & RCM_Menu_9 & "'," & RCM_Menu_10 & ",'" & RCM_Menu_11 & "'," & RCM_Menu_12 & "," & RCM_Menu_13 & "," & RCM_Menu_14 & "," & RCM_Menu_15 & ",'" & RCM_Menu_16 & "','" & RCM_Menu_17 & "','" & RCM_Menu_18 & "'," & RCM_Menu_19 & "," & RCM_Menu_20 & "," & RCM_Menu_21 & ",'" & RCM_Menu_22 & "']);"
    End If
End Function

Function stm_ai()
    stm_ai = "stm_ai('p0i0',[0,'|','','',-1,-1,0,'','_self','','','','',0,0,0,'','',0,0,0," & RCM_Item_22 & "," & RCM_Item_23 & ",'" & RCM_Item_24 & "'," & RCM_Item_25 & ",'" & RCM_Item_26 & "'," & RCM_Item_27 & ",'" & RCM_Item_28 & "','" & RCM_Item_29 & "'," & RCM_Item_30 & "," & RCM_Item_31 & "," & RCM_Item_32 & "," & RCM_Item_33 & ",'" & RCM_Item_34 & "','" & RCM_Item_35 & "','" & RCM_Item_36 & "','" & RCM_Item_37 & "','" & FontSize_RCM_Item_38 & " " & FontName_RCM_Item_38 & "','" & FontSize_RCM_Item_39 & " " & FontName_RCM_Item_39 & "','" & FontSize_RCM_Item_38 & " " & FontName_RCM_Item_38 & "','" & FontSize_RCM_Item_39 & " " & FontName_RCM_Item_39 & "']);"
End Function

Function stm_aix(mOID, mTID, mClassName, mClassFile, mOpenType, mMenuTitle, mSubClass)
    If mSubClass = False Then
        stm_aix = "stm_aix('" & mOID & "','" & mTID & "',[0,'" & mClassName & "','','',-1,-1,0,'" & mClassFile & "','" & mOpenType & "','" & mClassFile & "','" & EncodeJS(mMenuTitle) & "','','',0,0,0,'','',0,0,0," & RCM_Item_22 & "," & RCM_Item_23 & ",'" & RCM_Item_24 & "'," & RCM_Item_25 & ",'" & RCM_Item_26 & "'," & RCM_Item_27 & ",'" & RCM_Item_28 & "','" & RCM_Item_29 & "'," & RCM_Item_30 & "," & RCM_Item_31 & "," & RCM_Item_32 & "," & RCM_Item_33 & ",'" & RCM_Item_34 & "','" & RCM_Item_35 & "','" & RCM_Item_36 & "','" & RCM_Item_37 & "','" & FontSize_RCM_Item_38 & " " & FontName_RCM_Item_38 & "','" & FontSize_RCM_Item_39 & " " & FontName_RCM_Item_39 & "']);"
    ElseIf mSubClass = True Then
        stm_aix = "stm_aix('" & mOID & "','" & mTID & "',[0,'" & mClassName & "','','',-1,-1,0,'" & mClassFile & "','" & mOpenType & "','" & mClassFile & "','" & EncodeJS(mMenuTitle) & "','','',6,0,0,'" & strInstallDir & "images/default/arrow_r.gif','" & strInstallDir & "images/default/arrow_w.gif',7,7,0," & RCM_Item_22 & "," & RCM_Item_23 & ",'" & RCM_Item_24 & "'," & RCM_Item_25 & ",'" & RCM_Item_26 & "'," & RCM_Item_27 & ",'" & RCM_Item_28 & "','" & RCM_Item_29 & "'," & RCM_Item_30 & "," & RCM_Item_31 & "," & RCM_Item_32 & "," & RCM_Item_33 & ",'" & RCM_Item_34 & "','" & RCM_Item_35 & "','" & RCM_Item_36 & "','" & RCM_Item_37 & "','" & FontSize_RCM_Item_38 & " " & FontName_RCM_Item_38 & "','" & FontSize_RCM_Item_39 & " " & FontName_RCM_Item_39 & "']);"
    End If
End Function
    
Function EncodeJS(str)
    EncodeJS = Replace(Replace(Replace(Replace(Replace(str, Chr(10), ""), "\", "\\"), "'", "\'"), vbCrLf, "\n"), Chr(13), "")
End Function

Sub ShowDemoMenu()
    Response.Write "<script type='text/javascript' language='JavaScript1.2' src='" & strInstallDir & "KS_Inc/stm31.js'></script>"
    Response.Write "<script language='JavaScript' src='" & KS.Setting(3) & KS.Setting(93) & "/Menu.js'></script>"
End Sub

Function FilterString(strChar)
    If strChar = "" Or IsNull(strChar) Then
        FilterString = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "',%,<,>," & Chr(34) & ""
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    FilterString = tempChar
End Function

'取得网站的所有频道及其子栏目
Function ReturnAllChannel()
     Dim RS:Set RS=KS.InitialObject("ADODB.Recordset")
	  Dim SQL,K,ChannelStr:ChannelStr = ""
	   ChannelStr = "<select class='textbox' name=""ChannelID"" style=""width:200;border-style: solid; border-width: 1"">"
	   ChannelStr = ChannelStr & "<option value=""0"">-不限制-</option>"
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

%>
