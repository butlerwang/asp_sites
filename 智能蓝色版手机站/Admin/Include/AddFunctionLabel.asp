<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
Dim KS:Set KS=New PublicCls
Dim FolderID:FolderID=Request.QueryString("FolderID")
Dim SQL,I,RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
 RSC.Open "Select ChannelID,ChannelName,ChannelTable,ItemName,BasicType From KS_Channel Where ChannelStatus=1 Order By ChannelID",Conn,1,1
 If Not RSC.Eof Then
	SQL=RSC.GetRows(-1)
 End If
RSC.Close:Set RSC=Nothing

%>
<html>
<head>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="JavaScript" src="../../KS_Inc/jQuery.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>新建函数标签</title></head>
<link href="admin_style.css" rel="stylesheet">
<script language="javascript">
function CheckForm(){
 frames["LabelShow"].CheckForm();
}
</script>
<style>
li{margin:0px;padding:0px;list-style-type:none}
.list{border-bottom:1px #83B5CD solid;background:url(../images/titlebg.png); height:36px; font-size:13px; color:#555;padding-left:10px;}
.list li{display:block;float:left;border:1px solid #DEEFFA;background-color:#F7FBFE;height:20px;line-height:20px;margin:1px;padding:2px}

.submenu {z-index:999;position:absolute;white-space : nowrap; margin:0 ; border:1px solid #DEEFFA;display:none;background:url(../images/portalbox_bg.gif) no-repeat left top;left:-10px;background-color:#ffffff;top:22px}
.submunu_popup {line-height:18px;text-align:left;padding:8px;}
.submunu_popup a{line-height:18px;}
.rl{position:relative;}


</style>
<body topmargin="0" leftmargin="0" scroll=no>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="25">
	  <div class="list">
	   <li><a href="Label/GetGenericList.asp?FolderID=<%=FolderID%>" Target="LabelShow">万能列表</a></li>
	   <li><a href="Label/GetSlide.asp?FolderID=<%=FolderID%>" Target="LabelShow">通用幻灯</a></li>
	   <li><a href="Label/GetRolls.asp?FolderID=<%=FolderID%>" Target="LabelShow">滚动图片</a></li>
	   <li><a href="Label/GetMarquee.asp?FolderID=<%=FolderID%>" Target="LabelShow">滚动文字</a></li>
	   <li><a href="Label/GetNotRuleList.asp?FolderID=<%=FolderID%>" Target="LabelShow">不规则</a></li>
	   <li><a href="Label/GetCirClassList.asp?FolderID=<%=FolderID%>" Target="LabelShow">循环列表</a></li>
	   <li><a href="Label/GetRelativeList.asp?FolderID=<%=FolderID%>" Target="LabelShow">相关链接</a></li>
	   <li><a href="Label/GetPageList.asp?FolderID=<%=FolderID%>" Target="LabelShow">终级分页</a></li>
	  
	       
	      <span onMouseOut="$('#Menu_special').hide();">
	   	   <li class="rl" style="height:25px" onMouseOver="$('#Menu_special').show();"><a href="#">专 题</a> <img src="../images/d.gif" align="absmiddle" />
		   
		    <div class="submenu submunu_popup" id="Menu_special">
			<a href="Label/GetSpecialList.asp?FolderID=<%=FolderID%>" title="专题列表标签" Target="LabelShow">专题列表标签</a><br />
			<a href="Label/GetCirSpecialList.asp?FolderID=<%=FolderID%>" title="循环显示分类专题标签" Target="LabelShow">循环显示分类专题标签</a><br />
			<a href="Label/GetLastSpecialList.asp?FolderID=<%=FolderID%>" title="循环显示分类专题标签" Target="LabelShow">分页显示分类下的所有专题标签</a><br />
			</div></li>
		 </span>
		 
	      <span onMouseOut="$('#Menu_space').hide();">
	   	   <li class="rl" style="height:25px" onMouseOver="$('#Menu_space').show();"><a href="#">空 间</a> <img src="../images/d.gif" align="absmiddle" />
		   
		    <div class="submenu submunu_popup" id="Menu_space">
			<a href="Label/GetSpaceList.asp?FolderID=<%=FolderID%>" title="空间门户列表标签" Target="LabelShow">空间门户列表标签</a><br />
			<a href="Label/GetBlogInfoList.asp?FolderID=<%=FolderID%>" title="最新日志列表标签" Target="LabelShow">最新日志列表标签</a><br />
			<a href="Label/Getxclist.asp?FolderID=<%=FolderID%>" title="最新相册列表标签" Target="LabelShow">最新相册列表标签</a><br />
			<a href="Label/GetGrouplist.asp?FolderID=<%=FolderID%>" title="最新圈子列表标签" Target="LabelShow">最新圈子列表标签</a><br />
			<a href="Label/GetEnterPriseNewslist.asp?FolderID=<%=FolderID%>" title="企业新闻列表标签" Target="LabelShow">企业新闻列表标签</a><br />

			</div></li>
		 </span>
		 
	      <span  onMouseOut="$('#Menu_other').hide();">
	   	   <li class="rl"  style="height:25px" onMouseOver="$('#Menu_other').show();"><a href="#">其 它</a> <img src="../images/d.gif" align="absmiddle" />
		   
		    <div class="submenu submunu_popup" id="Menu_other">
			<a href="Label/GetLocation.asp?FolderID=<%=FolderID%>" title="网站位置导航标签" Target="LabelShow">网站位置导航标签</a><br />
			<a href="Label/GetAnnounceList.asp?FolderID=<%=FolderID%>" title="网站公告列表标签" Target="LabelShow">网站公告列表标签</a><br />
			<a href="Label/GetNavigation.asp?FolderID=<%=FolderID%>" title="栏目(频道)总导航标签" Target="LabelShow">栏目(频道)总导航标签</a><br />
			<a href="Label/GetLinkList.asp?FolderID=<%=FolderID%>" title="友情链接列表标签" Target="LabelShow">友情链接列表标签</a><br />
			 <%If KS.C_S(9,21)="1" Then%>
			<a href="Label/GetSjList.asp?FolderID=<%=FolderID%>" title="试卷列表调用标签" Target="LabelShow">试卷列表调用标签</a><br />
			 <%End If%>
			 <%If KS.C_S(5,21)="1" Then%>
			<a href="Label/GetGroupBuyList.asp?FolderID=<%=FolderID%>" title="团购列表调用标签" Target="LabelShow">团购列表调用标签</a><br />
			 <%End If%>
			<a href="Label/GetClubList.asp?FolderID=<%=FolderID%>" title="论坛帖子调用标签" Target="LabelShow">论坛帖子调用标签</a><br />
			<a href="Label/GetSlide.asp?from=club&FolderID=<%=FolderID%>" Target="LabelShow">论坛幻灯调用</a>
			</div></li> 
		 </span>
		 
		 <span  onMouseOut="$('#Menu_ask').hide();">
	   	   <li class="rl" style="height:25px" onMouseOver="$('#Menu_ask').show();"><a href="#">问 吧</a> <img src="../images/d.gif" align="absmiddle" />
		   
		    <div class="submenu submunu_popup" id="Menu_ask">
			<a href="Label/GetAQList.asp?FolderID=<%=FolderID%>" title="最新提问标签" Target="LabelShow">最新提问标签</a><br />
			<a href="Label/GetAAList.asp?FolderID=<%=FolderID%>" title="最新回答标签" Target="LabelShow">最新回答标签</a><br />
			<a href="Label/GetAskZJList.asp?FolderID=<%=FolderID%>" title="专家排行列表标签" Target="LabelShow">专家排行列表标签</a><br />
			</div></li>
		 </span>
		 
		 <%If KS.C_S(10,21)="1" Then%>
	      <span  onMouseOut="$('#Menu_job').hide();">
	   	   <li class="rl" style="height:25px" onMouseOver="$('#Menu_job').show();"><a href="#">招 聘</a> <img src="../images/d.gif" align="absmiddle" />
		   
		    <div class="submenu submunu_popup" id="Menu_job">
			<a href="Label/GetJobList.asp?FolderID=<%=FolderID%>" title="招聘职位列表标签" Target="LabelShow">招聘职位列表标签</a><br />
			<a href="Label/GetJobZWList.asp?FolderID=<%=FolderID%>" title="纯职位列表标签" Target="LabelShow">纯职位列表标签</a><br />
			<a href="Label/GetJobResumeList.asp?FolderID=<%=FolderID%>" title="简历列表标签" Target="LabelShow">简历列表标签</a><br />
			</div></li>
		 </span>
		 <%End If%>
		 
		 
	  </div>
	  <div style="clear:both"></div>
</td>
  </tr>
  <tr>
    <td valign="top">
	 <iframe name="LabelShow" ID="LabelShow" src="Label/GetGenericList.asp?PageTitle=新建通用列表标签&FolderID=<%=FolderID%>" style="width:100%;height:100%" frameborder="0" scrolling="auto"></iframe>	</td>
  </tr>
</table>
</body>
</html> 
