<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Session.asp"-->
<%
Dim SChannelID:SchannelID=request("schannelid")   'SchannelID=9999代表从自由标签/JS调用
Dim TemplateType:TemplateType=request("templateType")
Dim KS,KSCls,SQL,K,i,DIYFieldArr,ChannelID,FieldXML,FieldNode,FNode
On Error Resume Next
Set KS=New PublicCls
Set KSCls=New ManageCls
Dim DomainStr:DomainStr=KS.GetDomain
Dim RS:Set RS=Conn.Execute("Select ChannelID,BasicType,ChannelName,ItemName,ItemUnit,FieldBit,ModelEname From KS_Channel Where ChannelStatus=1 and channelid<>6  And ChannelID<>9 And ChannelID<>10 Order By ChannelID")
SQL=RS.GetRows(-1)
RS.Close:Set RS=Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script language="JavaScript" src="../../ks_inc/Common.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>标签插入</title>
<style type="text/css">
a{text-decoration: none;} /* 链接无下划线,有为underline */ 
a:link {color: #000000;} /* 未访问的链接 */
a:visited {color: #000000;} /* 已访问的链接 */
a:hover{color: #FF0000;text-decoration: underline;} /* 鼠标在链接上 */ 
a:active {color: #FF0000;} /* 点击激活链接 */
td	{font-family:  "Verdana, Arial, Helvetica, sans-serif"; font-size: 11.5px; color: #000000; text-decoration:none ; text-decoration:none ; }
BODY {
font-family:  "Verdana, Arial, Helvetica, sans-serif"; font-size: 11.5px;
FONT-SIZE: 9pt;
color: #000000;
text-decoration: none;
}
li{
list-style:none;
list-style-image:url(../Images/label0.gif);
margin-left:20px;
margin-bottom:2px;
}
</style>
</head>
<body topmargin="0" leftmargin="0">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="25"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="22" align="center" bgcolor="#0000FF"><strong><font color="#FFFFFF">网站系统---标签列表</font></strong></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td valign="top"> 
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr ParentID=""> 
          <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr  onmouseout="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
                <td><img src="../Images/home.gif" width="18" height="18"></td>
                <td height="20">标签导航</td>
              </tr>
              <tr onClick="ShowLabelTree('TY')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                <td width="672"><a href="#">网站通用标签</a></td>
              </tr>
              <tr> 
                <td colspan="2"> 
				   <div id="TY" style="display:none">
                    <li><a href="#" onClick="InsertLabel('{$GetSiteTitle}');">显示网站标题</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetSiteName}');">显示网站名称</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetSiteLogo}');">显示网站Logo(不带参数)</a></li>
					<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=Logo',250,130);">显示网站Logo(带参数)</a></li>
					<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=Tags',250,130);">显示热门Tags/最新Tags</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetSiteCountAll}');">显示网站信息统计</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetSiteOnline}');">显示在线人数(总在线：1人 用户：1人 游客：0人)</a></li>
					<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=TopUser',250,130);">显示活跃排行</a></li>
					<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=UserDynamic',250,130);">显示用户动态</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetSpecial}');">显示专题入口</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetFriendLink}');">显示友情链接入口</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetSiteUrl}');">显示网站URL</a></li>
					<li><a href="#" onClick="InsertLabel('{#GetFullDomain}');">显示网站完整URL(不管有没有启用相对路径,始终返回完整域名)</a></li>
					
					<li><a href="#" onClick="InsertLabel('{$GetInstallDir}');">显示网站安装路径</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetManageLogin}');">显示管理入口</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetCopyRight}');">显示版权信息</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetMetaKeyWord}');">显示针对搜索引擎的关键字</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetMetaDescript}');">显示针对搜索引擎的描述</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetWebmaster}');">显示站长</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetWebmasterEmail}');">显示站长EMail</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetClubInstallDir}');">论坛安装目录</a></li>
					<li><a href="#" onClick="InsertLabel('{$TodayGroupbuyLink}');">获得今日团购URL</a></li>
					<li><a href="#" onClick="InsertLabel('{$HistoryGroupbuyLink}');">获得往期团购URL</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetTemplateDir}');">模板路径</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetCssDir}');">CSS路径</a></li>
				 </div>
				 </td>
              </tr>
            </table></td>
        </tr>
      </table>
	  
	   <div onClick="ShowLabelTree('CommonJSLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
               <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">常用脚本特效标签</a></div>
              
				 <div id="CommonJSLabel" style="display:none">
				     <li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>editor/ksplus/labelparam.asp?action=ad',550,180);" class="LabelItem">对联广告</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_Time1}');" class="LabelItem">时间特效(样式:2006年4月8日)</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_Time2}');" class="LabelItem">时间特效(样式:2006年4月8日 星期六)</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_Time3}');" class="LabelItem">时间特效(样式:2007年6月1日 星期五【农历 4月...)</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_Time4}');" class="LabelItem">时间特效(样式:2006年4月8日 11:50:46 星期六)</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_Language}');" class="LabelItem">简繁转换</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_HomePage}');" class="LabelItem">设为首页</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_Collection}');" class="LabelItem">加入收藏</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_ContactWebMaster}');" class="LabelItem">联系站长</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_GoBack}');" class="LabelItem">返回上一页</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_WindowClose}');" class="LabelItem">关闭窗口</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_NoSave}');" class="LabelItem">页面不被别人"另存为"</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_NoIframe}');" class="LabelItem">页面不被别人放在框架中</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_NoCopy}');" class="LabelItem">防止网页信息被复制</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_DCRoll}');" class="LabelItem">双击滚屏特效</a></li>
					 <li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=Status1',550,150);" class="LabelItem">状态栏打字效果</a></li>
					 <li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=Status2',550,150);" class="LabelItem">文字在状态栏上从右往左循环显示</a></li>
					 <li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=Status3',550,150);" class="LabelItem">文字在状态栏上打字之后移动消失</a></li>
					</div>
               
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr ParentID=""> 
          <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr onClick="ShowLabelTree('SysFLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                <td width="672"><a href="#"><font color="blue">系统函数标签(KesionCMS入门标签)</font></a></td>
              </tr>
              <tr> 
                <td colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0" id="SysFLabel" style="display:none">
                    <tr> 
                      <td width="8%" align="right">&nbsp;</td>
                      <td height="20"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <%dim FolderRS,SqlStr
                          SqlStr = "Select * From KS_LabelFolder where FolderType=0 And ParentID='0'"
                         Set FolderRS = Conn.Execute(SqlStr)
                           if Not FolderRS.Eof then
	                    do while Not FolderRS.Eof
                           %>
                          <tr ParentID="<% = FolderRS("ParentID") %>" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                            <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td width="3%"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                                  <td width="97%"><span ShowFlag="False" TypeID="<% = FolderRS("ID") %>" onClick="SelectFolder(this)"><A href="#">
                                    <% = FolderRS("FolderName") %>
                                    </A></span></td>
                                </tr>
                              </table></td>
                          </tr>
                          <%
	 		        Response.Write(GetChildFolderList(0,0,FolderRS("ID"),""," style=""display:none;"" "))
                    Response.Write(GetLabelList(0,trim(FolderRS("ID")),"&nbsp;&nbsp;&nbsp;&nbsp;"," style=""display:none;"" "))
		        FolderRS.MoveNext
	            loop
              end if
               Response.Write(GetLabelList(0,"0","",""))
              %>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table>
	  
               
         <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr ParentID=""> 
              <td> 
			   <table width="100%" border="0" cellpadding="0" cellspacing="0">
               <tr onClick="ShowLabelTree('ContentLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                <td width="672"><a href="#"><font color="red">内容页标签</font></a></td>
               </tr>
               <tr> 
                <td colspan="2">
				    <table width="85%" align='center' border="0" cellspacing="0" cellpadding="0" id="ContentLabel" style="display:none">
                    <tr> 
					 <td>
					 <%
					  For K=0 To Ubound(SQL,2)
					   Call KSCls.LoadModelField(SQL(0,k),FieldXML,FieldNode)
					   Set Fnode=FieldXML.DocumentElement
					  %>
					 	<div onClick="ShowLabelTree('Content<%=SQL(6,K)%>')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
					<img src="../Images/Folder/folderclosed.gif" align="absmiddle" width="24" height="22"><a href="#"><%=SQL(3,K)%>内容页标签(<%=sql(2,k)%>)</a>
				     </div>	
					 <%Select Case SQL(1,K)%>
					  <%case 1%>
					  <div  id="Content<%=SQL(6,K)%>" style="display:none">
						<li><a href="#" onClick="InsertLabel('{$GetArticleUrl}');" class="LabelItem">当前<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">当前模型ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">当前<%=SQL(3,K)%>小ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">当前项目名称</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">当前项目单位</a></li>
						
						<li><a href="#" onClick="InsertLabel('{$GetArticleShortTitle}');" class="LabelItem"><%=SQL(3,K)%>简短标题</a></li>
						<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>editor/ksplus/labelparam.asp?action=ArtPhoto',250,130);" class="LabelItem">内容页图片</a></li>
						<%If Fnode.selectsinglenode("fielditem[@fieldname='title']/showonform").text="1" Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleTitle}');" class="LabelItem">完整标题</a></li>
						<%end if%>
						<%If Fnode.selectsinglenode("fielditem[@fieldname='keywords']/showonform").text="1" Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleKeyWord}');" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='keywords']/title").text%></a></li>
						<%end if%>
						<li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">取得<%=SQL(3,K)%>Tags</a></li>
						<%If Fnode.selectsinglenode("fielditem[@fieldname='intro']/showonform").text="1" Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleIntro}');" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='intro']/title").text%></a></li>
						<%end if%>
						<%If Fnode.selectsinglenode("fielditem[@fieldname='content']/showonform").text="1" Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleContent}');" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='content']/title").text%></a></li>
						<%end if%>
						<%If Fnode.selectsinglenode("fielditem[@fieldname='author']/showonform").text="1" Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleAuthor}');"><%=Fnode.selectsinglenode("fielditem[@fieldname='author']/title").text%></a></li>
						<%end if%>
						<%If Fnode.selectsinglenode("fielditem[@fieldname='origin']/showonform").text="1" Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleOrigin}');"><%=Fnode.selectsinglenode("fielditem[@fieldname='origin']/title").text%></a></li>
						<%end if%>
						<%If Fnode.selectsinglenode("fielditem[@fieldname='adddate']/showonform").text="1" Then%>
						<li><a href="#" onClick="InsertLabel('{$GetAddDate}');"><%=Fnode.selectsinglenode("fielditem[@fieldname='adddate']/title").text%>(格式:2012年10月1日)</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetDate}');"><%=Fnode.selectsinglenode("fielditem[@fieldname='adddate']/title").text%>(直接输出)</a></li>
						<%end if%>
						<li><a href="#" onClick="InsertLabel('{$GetModifyDate}');">修改时间(格式:2012年10月1日)</a></li>
						<%If Fnode.selectsinglenode("fielditem[@fieldname='hits']/showonform").text="1" Then%>
						 <li><a href="#" onClick="InsertLabel('{$GetHits}');" class="LabelItem"><%=SQL(3,K)%>人气(总浏览数)</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetHitsByDay}');" class="LabelItem"><%=SQL(3,K)%>本日浏览数</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetHitsByWeek}');" class="LabelItem"><%=SQL(3,K)%>本周浏览数</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetHitsByMonth}');" class="LabelItem"><%=SQL(3,K)%>本月浏览数</a></li>
						<%end if%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleInput}');"><%=SQL(3,K)%>录入(带链接)</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetUserName}');"><%=SQL(3,K)%>录入(不带链接)</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetRank}');"><%=SQL(3,K)%>推荐等级</a></li>
						<%If Fnode.selectsinglenode("fielditem[@fieldname='attribute']/showonform").text="1" Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleProperty}');">显示<%=SQL(3,K)%>的属性(热门、推荐、滚动、...)</a></li>
						<%end if%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleSize}');">显示<%=SQL(3,K)%>【字体:大 中 小】</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetArticleAction}');">显示【发表评论】【告诉好友】【打印...</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevArticle}');">显示上一<%=SQL(4,K)%><%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextArticle}');">显示下一<%=SQL(4,K)%><%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevUrl}');">显示上一<%=SQL(4,K)%><%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextUrl}');">显示下一<%=SQL(4,K)%><%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetShowComment}');">显示评论</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">发表评论</a></li>
						
                     <%Case 2%>					  
					  <div id="Content<%=SQL(6,K)%>" style="display:none">
						 <li><a href="#" onClick="InsertLabel('{$GetPictureUrl}');" class="LabelItem">当前<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">当前模型ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">当前<%=SQL(3,K)%>小ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">当前项目名称</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">当前项目单位</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureName}');" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='title']/title").text%></a></li>
						 <%If Fnode.selectsinglenode("fielditem[@fieldname='keywords']/showonform").text="1" Then%>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureKeyWord}');" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='keywords']/title").text%></a></li>
						 <%end if%>
						<li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">取得<%=SQL(3,K)%>Tags</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPictureSrc}');" class="LabelItem">取得<%=SQL(3,K)%>缩略图Src</a></li>
						
						 <li><a href="#" onClick="InsertLabel('{$ShowPictures}');"  style="color:red" class="LabelItem"><%=SQL(3,K)%>展示(根据添加图片时的展示方式自动显示)</a></li>
						  <%If Fnode.selectsinglenode("fielditem[@fieldname='picturecontent']/showonform").text="1" Then%>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureIntro}');" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='picturecontent']/title").text%></a></li>
						 <%end if%>
						 <%If Fnode.selectsinglenode("fielditem[@fieldname='author']/showonform").text="1" Then%>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureAuthor}');" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='author']/title").text%></a></li>
						 <%end if%>
						 <%If Fnode.selectsinglenode("fielditem[@fieldname='origin']/showonform").text="1" Then%>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureOrigin}');" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='origin']/title").text%></a></li>
						 <%end if%>
						 <%If Fnode.selectsinglenode("fielditem[@fieldname='addddate']/showonform").text="1" Then%>
						 <li><a href="#" onClick="InsertLabel('{$GetAddDate}');"><%=Fnode.selectsinglenode("fielditem[@fieldname='adddate']/title").text%>(格式:2012年10月1日)</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetDate}');"><%=Fnode.selectsinglenode("fielditem[@fieldname='adddate']/title").text%>(直接输出)</a></li>
						 <%end if%>
						 <li><a href="#" onClick="InsertLabel('{$GetModifyDate}');">修改时间(格式:2012年10月1日)</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetHits}');" class="LabelItem"><%=SQL(3,K)%>人气(总浏览数)</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetHitsByDay}');" class="LabelItem"><%=SQL(3,K)%>本日浏览数</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetHitsByWeek}');" class="LabelItem"><%=SQL(3,K)%>本周浏览数</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetHitsByMonth}');" class="LabelItem"><%=SQL(3,K)%>本月浏览数</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPictureInput}');"><%=SQL(3,K)%>录入(带链接)</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetUserName}');"><%=SQL(3,K)%>录入(不带链接)</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetRank}');"><%=SQL(3,K)%>推荐等级</a></li>
						 <%If Fnode.selectsinglenode("fielditem[@fieldname='attribute']/showonform").text="1" Then%>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureProperty}');" class="LabelItem">显示<%=SQL(3,K)%>属性(热门、滚动、推荐...</a></li>
						 <%end if%>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureAction}');">&nbsp;显示【我来评论】【我要...】</a><</li>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureVote}');" class="LabelItem">显示投它一票</a> </li>
						 <li><a href="#" onClick="InsertLabel('{$GetPicNums}');" class="LabelItem">显示总张数</a> </li>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureVoteScore}');" class="LabelItem">显示<%=SQL(3,K)%>得票数</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetPrevPicture}');" class="LabelItem">显示上一组<%=SQL(3,K)%></a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetNextPicture}');" class="LabelItem">显示下一组<%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevUrl}');">显示上一组<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextUrl}');">显示下一组<%=SQL(3,K)%>URL</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetShowComment}');">显示评论</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">发表评论</a></li>
				<%Case 3%>
				 <div id="Content<%=SQL(6,K)%>" style="display:none">
					 <li><a href="#" onClick="InsertLabel('{$GetDownUrl}');" class="LabelItem">当前<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">当前模型ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">当前<%=SQL(3,K)%>小ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">当前项目名称</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">当前项目单位</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownTitle}');" class="LabelItem"><%=SQL(3,K)%>名称+版本号</a></li>
					 <%If Fnode.selectsinglenode("fielditem[@fieldname='keywords']/showonform").text="1" Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownKeyWord}');" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='keywords']/title").text%></a></li>
					 <%end if%>
						<li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">取得<%=SQL(3,K)%>Tags</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownAddress}');" class="LabelItem">下载地址</a></li>
					 <%If Fnode.selectsinglenode("fielditem[@fieldname='photourl']/showonform").text="1" Then%>
					 <li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=DownPhoto',250,130);" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='photourl']/title").text%></a></li>
					 <%end if%>
					 
					 <li><a href="#" onClick="InsertLabel('{$GetDownSize}');" class="LabelItem">文件大小+MB(KB)</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownLanguage}');" class="LabelItem"><%=SQL(3,K)%>语言</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownType}');" class="LabelItem"><%=SQL(3,K)%>类别</a></li>
					  <%If Fnode.selectsinglenode("fielditem[@fieldname='platform']/showonform").text="1" Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownSystem}');" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='platform']/title").text%></a></li>
					 <%end if%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownPower}');" class="LabelItem">授权方式</a></li>
					 <%If Fnode.selectsinglenode("fielditem[@fieldname='content']/showonform").text="1" Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownIntro}');" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='content']/title").text%></a></li>
					 <%end if%>
					 <%If Fnode.selectsinglenode("fielditem[@fieldname='author']/showonform").text="1" Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownAuthor}');" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='author']/title").text%></a></li>
					 <%end if%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownInput}');" class="LabelItem"><%=SQL(3,K)%>录入(带链接)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetUserName}');" class="LabelItem"><%=SQL(3,K)%>录入(不带链接)</a></li>
					 <%If Fnode.selectsinglenode("fielditem[@fieldname='origin']/showonform").text="1" Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownOrigin}');" class="LabelItem">来 源</a></li>
					 <%end if%>
					 <li><a href="#" onClick="InsertLabel('{$GetAddDate}');" class="LabelItem">添加日期(格式:2012年10月1日)</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetModifyDate}');">修改时间(格式:2012年10月1日)</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetHits}');" class="LabelItem">总下载点击数</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetHitsByDay}');" class="LabelItem">本日点击数</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetHitsByWeek}');" class="LabelItem">本周点击数</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetHitsByMonth}');" class="LabelItem">本月点击数</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownLink}');" class="LabelItem">相关链接（演示地址+注册地址）</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownPoint}');" class="LabelItem">下载所需点卷</a></li>
					 <%If Fnode.selectsinglenode("fielditem[@fieldname='ysdz']/showonform").text="1" Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownYSDZ}');" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='ysdz']/title").text%></a></li>
					 <%end if%>
					 <%If Fnode.selectsinglenode("fielditem[@fieldname='zcdz']/showonform").text="1" Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownZCDZ}');" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='zcdz']/title").text%></a></li>
					 <%end if%>
					 <%If Fnode.selectsinglenode("fielditem[@fieldname='jymm']/showonform").text="1" Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownDecPass}');" class="LabelItem"><%=Fnode.selectsinglenode("fielditem[@fieldname='jymm']/title").text%></a></li>
					 <%end if%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownProperty}');" class="LabelItem">显示<%=SQL(3,K)%>属性(热门、推荐等）</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownAction}');">显示【我来评论】【我要...</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetRank}');" class="LabelItem">显示推荐星级</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetPrevDown}');" class="LabelItem">显示上一个<%=SQL(3,K)%></a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetNextDown}');" class="LabelItem">显示下一个<%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevUrl}');">显示上一个<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextUrl}');">显示下一个<%=SQL(3,K)%>URL</a></li>
					 <Li><a href="#" onClick="InsertLabel('{$GetShowComment}');">显示评论</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">发表评论</a></li>
				<%Case 4%>
				 <div id="Content<%=SQL(6,K)%>" style="display:none">
						<li><a href="#" onClick="InsertLabel('{$GetFlashUrl}');" class="LabelItem">当前<%=SQL(3,K)%> URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">当前模型ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">当前<%=SQL(3,K)%>小ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">当前项目名称</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">当前项目单位</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashName}');" class="LabelItem"><%=SQL(3,K)%>名称</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashKeyWord}');" class="LabelItem">当前<%=SQL(3,K)%>关键词</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">取得<%=SQL(3,K)%>Tags</a></li>
						<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=FlashPlayer',250,130);" class="LabelItem">查看<%=SQL(3,K)%>内容(播放器方式播放)</a></li>
						<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=Flash',250,130);" class="LabelItem">查看<%=SQL(3,K)%>内容(普通方式播放)</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashIntro}');" class="LabelItem"><%=SQL(3,K)%>简介</a> </li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashAuthor}');" class="LabelItem"><%=SQL(3,K)%>作者</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashOrigin}');" class="LabelItem">来 源</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetAddDate}');" class="LabelItem">添加日期(格式:2012年10月1日)</a></li>
					    <li><a href="#" onClick="InsertLabel('{$GetModifyDate}');">修改时间(格式:2012年10月1日)</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashSrc}');" class="LabelItem"><%=SQL(3,K)%>地址</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashFullScreen}');" class="LabelItem">全屏观看</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetHits}');" class="LabelItem"><%=SQL(3,K)%>人气(总浏览数)</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetHitsByDay}');" class="LabelItem"><%=SQL(3,K)%>本日浏览数</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetHitsByWeek}');" class="LabelItem"><%=SQL(3,K)%>本周浏览数</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetHitsByMonth}');" class="LabelItem"><%=SQL(3,K)%>本月浏览数</a></li>
		
						<li><a href="#" onClick="InsertLabel('{$GetFlashInput}');" class="LabelItem">动漫录入</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashProperty}');" class="LabelItem">显示<%=SQL(3,K)%>属性(热门、滚动、推荐、等级等）</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashAction}');">&nbsp;显示【我来评论】【我要收藏】【关闭窗口】</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashVote}');" class="LabelItem">显示投它一票</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashVoteScore}');" class="LabelItem">显示<%=SQL(3,K)%>得票数</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevFlash}');" class="LabelItem">显示上一个<%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextFlash}');" class="LabelItem">显示下一个<%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevUrl}');">显示上一个<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextUrl}');">显示下一个<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetShowComment}');">&nbsp;显示评论</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">&nbsp;发表评论</a></li>
			       <%Case 5%>
				<div id="Content<%=SQL(6,K)%>" style="display:none">
					<li><a href="#" onClick="InsertLabel('{$GetProductUrl}');" class="LabelItem">当前<%=SQL(3,K)%> URL</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductID}');" class="LabelItem">当前<%=SQL(3,K)%>编号(ID)</a></li>
						<li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">当前模型ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">当前<%=SQL(3,K)%>小ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">当前项目名称</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">当前项目单位</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductName}');" class="LabelItem"><%=SQL(3,K)%>名称</a></li>
					<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=ProductPhoto',250,130);" class="LabelItem">商品图片</a> </li>
					<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=ProductGroupPhoto',250,130);" class="LabelItem">显示商品图片组 <font color=red>new</font></a> </li>
					
					<li><a href="#" onClick="InsertLabel('{$GetProductKeyWord}');" class="LabelItem">当前<%=SQL(3,K)%>关键词</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">取得<%=SQL(3,K)%>Tags</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductIntro}');" class="LabelItem"><%=SQL(3,K)%>简介</a> </li>
					<li><a href="#" onClick="InsertLabel('{$GetProducerName}');" class="LabelItem">生 产 商</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetTrademarkName}');" class="LabelItem">取得商标</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetBrandName}');" class="LabelItem">取得品牌名称</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductModel}');" class="LabelItem"><%=SQL(3,K)%>型号</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductSpecificat}');" class="LabelItem"><%=SQL(3,K)%>规格</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetAddDate}');" class="LabelItem">添加时间(格式:2012年10月1日)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetModifyDate}');">修改时间(格式:2012年10月1日)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetServiceTerm}');" class="LabelItem">服务期限</a></li>
					<li><a href="#" onClick="InsertLabel('{$FL_Weight}');" class="LabelItem">单件重量</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetTotalNum}');" class="LabelItem">库存数量</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHasSold}');" class="LabelItem">显示已销售件数</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductUnit}');" class="LabelItem"><%=SQL(3,K)%>单位</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHits}');" class="LabelItem"><%=SQL(3,K)%>人气(总浏览数)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHitsByDay}');" class="LabelItem"><%=SQL(3,K)%>本日浏览数</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHitsByWeek}');" class="LabelItem"><%=SQL(3,K)%>本周浏览数</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHitsByMonth}');" class="LabelItem"><%=SQL(3,K)%>本月浏览数</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductType}');" class="LabelItem">销售类型</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetRank}');" class="LabelItem">推荐等级</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductProperty}');" class="LabelItem">显示<%=SQL(3,K)%>属性(热卖、特价、推荐等）</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductInput}');">&nbsp;显示商品录入</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetPrice}');" class="LabelItem">显示当前零售价</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetPrice_Member}');" class="LabelItem">显示商城价</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetPrice_Market}');" class="LabelItem">显示市场价</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetGroupPrice}');" class="LabelItem">自动取用户组价格 <font color=red>new</font></a></li>
					<li><a href="#" onClick="InsertLabel('{$GetScore}');" class="LabelItem">显示购物积分</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetAddCar}');" class="LabelItem">加入购物车</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetAddFav}');" class="LabelItem">加入收藏夹</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetPrevProduct}');" class="LabelItem">显示上一个<%=SQL(3,K)%></a></li>
					<li><a href="#" onClick="InsertLabel('{$GetNextProduct}');" class="LabelItem">显示下一个<%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevUrl}');">显示上一个<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextUrl}');">显示下一个<%=SQL(3,K)%>URL</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetShowComment}');">显示评论</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">发表评论</a></li>
			<%Case 7%>
			<div id="Content<%=SQL(6,K)%>" style="display:none">
				<li><a href="#" onClick="InsertLabel('{$GetMovieUrl}');" class="LabelItem">当前<%=SQL(3,K)%> URL</a></li>
				<li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">当前模型ID</a></li>
				<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">当前<%=SQL(3,K)%>ID</a></li>
				<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">当前项目名称</a></li>
				<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">当前项目单位</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieName}');" class="LabelItem"><%=SQL(3,K)%>名称</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieActor}');" class="LabelItem">主要演员</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieDirector}');" class="LabelItem"><%=SQL(3,K)%>导演</a> </li>
				<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=MoviePhoto',250,130);" class="LabelItem"><%=SQL(3,K)%>图片</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieKeyWord}');" class="LabelItem">当前<%=SQL(3,K)%>关键词</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">取得<%=SQL(3,K)%>Tags</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieLanguage}');" class="LabelItem"><%=SQL(3,K)%>语言</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieArea}');" class="LabelItem">出产地区</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieIntro}');" class="LabelItem">查看影片介绍</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieTime}');" class="LabelItem"><%=SQL(3,K)%>长度（播放时间）</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetScreenTime}');" class="LabelItem">上映时间</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetAddDate}');" class="LabelItem">添加日期(格式:2012年10月1日)</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetModifyDate}');">修改时间(格式:2012年10月1日)</a></li>
				<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=MoviePlay',250,130);" class="LabelItem">播放列表</a></li>
				<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=MoviePage',250,130);" class="LabelItem">内容页播放器(适合做flv,mtv视频交流类站点)</a></li>
				<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=MovieDown',250,130);" class="LabelItem">下载列表</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHits}');" class="LabelItem"><%=SQL(3,K)%>人气(总浏览数)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHitsByDay}');" class="LabelItem"><%=SQL(3,K)%>本日浏览数</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHitsByWeek}');" class="LabelItem"><%=SQL(3,K)%>本周浏览数</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHitsByMonth}');" class="LabelItem"><%=SQL(3,K)%>本月浏览数</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieNum}');" class="LabelItem">显示总集数</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetRank}');" class="LabelItem">显示推荐星级</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieInput}');" class="LabelItem"><%=SQL(3,K)%>录入</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetPoint}');" class="LabelItem">取得观看/下载点卷</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieProperty}');" class="LabelItem">显示影视属性(热门、滚动、推荐等）</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieVote}');" class="LabelItem">显示投它一票</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieVoteScore}');" class="LabelItem">显示<%=SQL(3,K)%>得票数</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetPrevMovie}');" class="LabelItem">显示上一部<%=SQL(3,K)%></a></li>
				<li><a href="#" onClick="InsertLabel('{$GetNextMovie}');" class="LabelItem">显示下一部<%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevUrl}');">显示上一部<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextUrl}');">显示下一部<%=SQL(3,K)%>URL</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetShowComment}');">显示评论</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">发表评论</a></li>
         <%Case 8%>
				<div id="Content<%=SQL(6,K)%>" style="display:none">
				  <li><a href="#" onClick="InsertLabel('{$GetGQInfoUrl}');" class="LabelItem">当前<%=SQL(3,K)%> URL</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetGQInfoID}');" class="LabelItem">当前<%=SQL(3,K)%>ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">当前模型ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">当前<%=SQL(3,K)%>小ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">当前项目名称</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">当前项目单位</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetGQTitle}');" class="LabelItem"><%=SQL(3,K)%>主题</a></li>
				  <li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>Editor/KSPlus/LabelParam.asp?action=SupplyPhoto',250,130);" class="LabelItem"><%=SQL(3,K)%>缩略图(带参数)</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetGQKeyWord}');" class="LabelItem">取得关键字</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">取得<%=SQL(3,K)%>Tags</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetPrice}');" class="LabelItem">价格说明</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetInfoType}');" class="LabelItem">信息类别</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetTransType}');" class="LabelItem">交易类别</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetValidTime}');" class="LabelItem">有 效 期</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetGQContent}');" class="LabelItem"><%=SQL(3,K)%>内容介绍</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetHits}');" class="LabelItem"><%=SQL(3,K)%>人气(总浏览数)</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetHitsByDay}');" class="LabelItem"><%=SQL(3,K)%>本日浏览数</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetHitsByWeek}');" class="LabelItem"><%=SQL(3,K)%>本周浏览数</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetHitsByMonth}');" class="LabelItem"><%=SQL(3,K)%>本月浏览数</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetAddDate}');" class="LabelItem">发布时间(格式:2012年10月1日)</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetModifyDate}');">修改时间(格式:2012年10月1日)</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetInput}');" class="LabelItem"><%=SQL(3,K)%>发布者(会员名称)</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetCompanyName}');" class="LabelItem">公司名称</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetContactMan}');" class="LabelItem">联系人</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetContactTel}');" class="LabelItem">联系电话</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetMobile}');" class="LabelItem">移动电话</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetFax}');" class="LabelItem">传真号码</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetAddress}');" class="LabelItem">详细地址</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetEmail}');" class="LabelItem">电子邮箱</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetPostCode}');" class="LabelItem">邮政编码</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetProvince}');" class="LabelItem">交易所在省份</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetCity}');" class="LabelItem">交易所在城市</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetHomePage}');" class="LabelItem">公司网址</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetPrevInfo}');" class="LabelItem">显示上一条<%=SQL(3,K)%></a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetNextInfo}');" class="LabelItem">显示下一条<%=SQL(3,K)%></a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetPrevUrl}');">显示上一条<%=SQL(3,K)%>URL</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetNextUrl}');">显示下一条<%=SQL(3,K)%>URL</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetShowComment}');">显示留言信息</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">发布留言</a></li>
				<%End Select%>
				<%If Fnode.selectsinglenode("fielditem[@fieldname='seooption']/showonform").text="1" Then%>
				   <li><a href="#" onClick="InsertLabel('{$GetSEOTitle}');">显示SEO页面标题</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetSEOKeyWords}');">显示SEO页面关键字</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetSEODescription}');">显示SEO页面描述</a></li>
				<%end if%>
					<div>============================</div>
					<div align='center'>自定义字段标签</div>
					<div>============================</div>
                          <%
						  if FieldXML.readystate=4 and FieldXML.parseError.errorCode=0 Then 
							  Dim DiyNode:Set DiyNode=FieldXML.DocumentElement.selectnodes("fielditem[fieldtype!=0]")
							  If DiyNode.Length>0 Then
							    For each Fnode In DiyNode
								  response.write " <li><a href=""#"" onClick=""InsertLabel('{$" & Fnode.selectsinglenode("@fieldname").text & "}');"">" & Fnode.selectsinglenode("title").text & "-{$" & Fnode.selectsinglenode("title").text & "}</a></li>"
								Next
							  End If
						  End If
				        %>

              </div>		
			   <%Next%>				  
					 </td>
					 </tr>
					 </table>
				</td>
			  </tr>
			  </table>
			  </td>
			  </tr>
			  </table>
			  
	  <div onClick="ShowLabelTree('ChannelClassLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
                <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">频道（栏目）专用标签</a>
	  </div>
				  <div id="ChannelClassLabel" style="display:none">  
				    <li><a href="#" onClick="InsertLabel('{$GetChannelID}');">显示当前模型ID</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetChannelName}');">显示当前模型名称</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetItemName}');" class="LabelItem">显示当前模型的项目名称</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetItemUnit}');" class="LabelItem">显示当前模型的项目单位</a></li>
				    =======================<br>    
				    <li><a href="#" onClick="InsertLabel('{$GetClassID}');">显示当前栏目ID</a></li>
				    <li><a href="#" onClick="InsertLabel('{$GetSmallClassID}');">显示当前栏目ClassID</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetClassName}');">显示当前栏目名称</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetClassUrl}');" class="LabelItem">显示当前栏目链接地址</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetClassPic}');" class="LabelItem">显示当前栏目图片</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetClassIntro}');" class="LabelItem">显示当前栏目介绍</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetClass_Meta_KeyWord}');" class="LabelItem">针对搜索引擎的关键字</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetClass_Meta_Description}');" class="LabelItem">针对搜索引擎的描述</a></li>
				    <li><a href="#" onClick="InsertLabel('{$GetParentID}');">显示父栏目ID</a></li>
				    <li><a href="#" onClick="InsertLabel('{$GetParentClassID}');">显示父栏目ClassID</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetParentUrl}');">显示父栏目链接地址</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetParentClassName}');">显示父栏目名称</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetTopClassName}');">显示一级栏目名称</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetTopClassUrl}');">显示一级栏目URL</a></li>
				 </div>
               
	  
       <div onClick="ShowLabelTree('SearchLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
                <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">搜索专用标签</a>
		</div>
				  <div id="SearchLabel" style="display:none">
				   <li><a href="#" onClick="InsertLabel('{$GetSearchByDate}');" class="LabelItem">高级日历搜索(小插件)</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetSearch}');" class="LabelItem">总站搜索</a></li>
				   <%
				   For K=0 To Ubound(SQL,2)
				    response.write "<li><a href=""#"" onClick=""InsertLabel('{$Get"  & SQL(6,K) & "Search}');"" class=""LabelItem"">" & SQL(2,K) & "搜索</a></li>"
				   Next
				   %>
				  </div>
			  

		
	    <div onClick="ShowLabelTree('AnnounceContent')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
               <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">公告内容页标签</a>
		</div>
		
		 <div id="AnnounceContent" style="display:none">
               <li><a href="#" onClick="InsertLabel('{$GetAnnounceTitle}');" class="LabelItem">公告标题</a></li>
			   <li><a href="#" onClick="InsertLabel('{$GetAnnounceAuthor}');" class="LabelItem">公告作者</a></li>
			   <li><a href="#" onClick="InsertLabel('{$GetAnnounceDate}');" class="LabelItem">公告发布(更新)时间</a></li>
			   <li><a href="#" onClick="InsertLabel('{$GetAnnounceContent}');" class="LabelItem">公告的具体内容</a></li>
		 </div>
			   
	  	  <div onClick="ShowLabelTree('LinkContent')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
                    <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">友情链接页标签</a>
		 </div>
             <div id="LinkContent" style="display:none">
                   <li><a href="#" onClick="InsertLabel('{$GetLinkCommonInfo}');" class="LabelItem">显示查看方式及申请友情链接等</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetClassLink}');" class="LabelItem">显示分类及友情链接站点搜索</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetLinkDetail}');" class="LabelItem">分页显示友情链接详细列表</a></li>
			</div>

	  	  <div onClick="ShowLabelTree('Special')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
                    <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">专题页标签</a>
		 </div>
             <div id="Special" style="display:none">
                   <li><a href="#" onClick="InsertLabel('{$GetSpecialID}');" class="LabelItem">当前专题ID</a></li>
                   <li><a href="#" onClick="InsertLabel('{$GetSpecialName}');" class="LabelItem">当前专题名称</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetSpecialPic}');" class="LabelItem">当前专题图片</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetSpecialNote}');" class="LabelItem">当前专题介绍</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetSpecialDate}');" class="LabelItem">当前专题添加时间</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetSpecialClassName}');" class="LabelItem">当前专题分类名称</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetSpecialClassURL}');" class="LabelItem">当前专题分类URL</a></li>
			</div>

		
	  	  <div onClick="ShowLabelTree('UserSystem');" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
                    <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">会员系统专用标签</a>
		  </div>
		  
          <div id="UserSystem" style="display:none">
					<li><a href="#" onClick="InsertLabel('{$GetUserLoginByScript}');" class="LabelItem">显示会员登录入口(Script调用)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetPopLogin}');" class="LabelItem">显示会员登录入口(跳窗)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetTopUserLogin}');" class="LabelItem">显示会员登录入口(横排)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetUserLogin}');" class="LabelItem">显示会员登录入口(竖排)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetAllUserList}');" class="LabelItem">显示所有注册会员列表(此标签仅限使用于会员列表页模板)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetUserRegLicense}');" class="LabelItem">显示新会员注册服务条款和声明</a></li>
					<li><a href="#" onClick="InsertLabel('{$Show_UserNameLimitChar}');" class="LabelItem">显示新会员注册时用户名最少字符数</a></li>
					<li><a href="#" onClick="InsertLabel('{$Show_UserNameMaxChar}');" class="LabelItem">显示新会员注册时用户名最多字符数</a></li>
					<li><a href="#" onClick="InsertLabel('{$Show_VerifyCode}');" class="LabelItem">显示新会员注册时验证码</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetUserRegResult}');" class="LabelItem">新会员注册成功信息</a>
               </div>
		  
		 <div onClick="ShowLabelTree('AdwLabel');" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">广告位通用标签</a></div>
             <div id="AdwLabel" style="display:none">
				<%  Dim RSObj:Set RSObj=server.createobject("adodb.recordset")
					SqlStr="select Place,PlaceName From KS_ADPlace"
					RSObj.open SqlStr,Conn,1,1
					do while not RSObj.eof 
                %>
                    <li><a href="#" onClick="InsertLabel('{=GetAdvertise(<%=RSObj(0)%>)}');" class="LabelItem"> <%=RSObj(1)%></a></li>
				<%RSOBj.MoveNext
				 Loop
				 RSObj.Close:SET RSObj=Nothing
				 %>
                 
		     </div>
   
	    <div onClick="ShowLabelTree('VoteLabel');" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">网站调查通用标签</a></div>
             <div id="VoteLabel" style="display:none">
				<%  Set RSObj=server.createobject("adodb.recordset")
					SqlStr="select ID,Title From KS_Vote"
					RSObj.open SqlStr,Conn,1,1
					do while not RSObj.eof 
                %>
                    <li><a href="#" onClick="InsertLabel('{=GetVote(<%=RSObj(0)%>)}');" class="LabelItem"><%=RSObj(1)%></a></li>
				<%RSOBj.MoveNext
				 Loop
				 RSObj.Close:SET RSObj=Nothing
				 %>
             </div>
			 
			 
		 <div onClick="ShowLabelTree('RssLabel');" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">RSS标签</a>
		 </div>
            <div id="RssLabel" style="display:none">
				<li><a href="#" onClick="InsertLabel('{$Rss}');" class="LabelItem">Rss标签显示</a></li>
				<li><a href="#" onClick="InsertLabel('{$RssElite}');" class="LabelItem">Rss推荐标签显示</a></li>
				<li><a href="#" onClick="InsertLabel('{$RssHot}');" class="LabelItem">Rss热门标签显示</a></li>
			</div>

	  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr ParentID=""> 
          <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr onClick="ShowLabelTree('DIYFunctionLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                <td width="672"><a href="#">用户自定义函数标签</a></td>
              </tr>
              <tr> 
                <td colspan="2">
				 <table width="100%" border="0" cellspacing="0" cellpadding="0" id="DIYFunctionLabel" style="display:none">
                    <tr> 
                      <td width="8%" align="right">&nbsp;</td>
                      <td height="20"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <%
                          SqlStr = "Select * From KS_LabelFolder where FolderType=5 And ParentID='0'"
                         Set FolderRS = Conn.Execute(SqlStr)
                           if Not FolderRS.Eof then
	                    do while Not FolderRS.Eof
                           %>
                          <tr ParentID="<% = FolderRS("ParentID") %>" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                            <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td width="3%"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                                  <td width="97%"><span ShowFlag="False" TypeID="<% = FolderRS("ID") %>" onClick="SelectFolder(this)"><A href="#">
                                    <% = FolderRS("FolderName") %>
                                    </A></span></td>
                                </tr>
                              </table></td>
                          </tr>
                          <%
	 		        Response.Write(GetChildFolderList(0,5,FolderRS("ID"),""," style=""display:none;"" "))
                    Response.Write(GetLabelList(5,trim(FolderRS("ID")),"&nbsp;&nbsp;&nbsp;&nbsp;"," style=""display:none;"" "))
		        FolderRS.MoveNext
	            loop
              end if
               Response.Write(GetLabelList(5,"0","",""))
              %>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table> 
          </td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr ParentID=""> 
          <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr onClick="ShowLabelTree('FreeLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                <td width="672"><a href="#">用户自定义静态标签</a></td>
              </tr>
              <tr> 
                <td colspan="2">
				 <table width="100%" border="0" cellspacing="0" cellpadding="0" id="FreeLabel" style="display:none">
                    <tr> 
                      <td width="8%" align="right">&nbsp;</td>
                      <td height="20"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <%
                          SqlStr = "Select * From KS_LabelFolder where FolderType=1 And ParentID='0'"
                         Set FolderRS = Conn.Execute(SqlStr)
                           if Not FolderRS.Eof then
	                    do while Not FolderRS.Eof
                           %>
                          <tr ParentID="<% = FolderRS("ParentID") %>" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                            <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td width="3%"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                                  <td width="97%"><span ShowFlag="False" TypeID="<% = FolderRS("ID") %>" onClick="SelectFolder(this)"><A href="#">
                                    <% = FolderRS("FolderName") %>
                                    </A></span></td>
                                </tr>
                              </table></td>
                          </tr>
                          <%
	 		        Response.Write(GetChildFolderList(0,1,FolderRS("ID"),""," style=""display:none;"" "))
                    Response.Write(GetLabelList(1,trim(FolderRS("ID")),"&nbsp;&nbsp;&nbsp;&nbsp;"," style=""display:none;"" "))
		        FolderRS.MoveNext
	            loop
              end if
               Response.Write(GetLabelList(1,"0","",""))
              %>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table> 
          </td>
        </tr>
      </table>

	   <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr ParentID=""> 
          <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr onClick="ShowLabelTree('SysJS')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                <td width="672"><a href="#">系统JS标签</a></td>
              </tr>
              <tr> 
                <td colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0" id="SysJS" style="display:none">
                    <tr> 
                      <td width="8%" align="right">&nbsp;</td>
                      <td height="20"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <%
                          SqlStr = "Select * From KS_LabelFolder where FolderType=2 And ParentID='0'"
                         Set FolderRS = Conn.Execute(SqlStr)
                           if Not FolderRS.Eof then
	                    do while Not FolderRS.Eof
                           %>
                          <tr ParentID="<% = FolderRS("ParentID") %>" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                            <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td width="3%"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                                  <td width="97%"><span ShowFlag="False" TypeID="<% = FolderRS("ID") %>" onClick="SelectFolder(this)"><A href="#"> 
                                    <% = FolderRS("FolderName") %>
                                    </A></span></td>
                                </tr>
                              </table></td>
                          </tr>
                          <%
	 		        Response.Write(GetChildFolderList(1,0,FolderRS("ID"),""," style=""display:none;"" "))
                    Response.Write(GetJSList(0,trim(FolderRS("ID")),"&nbsp;&nbsp;&nbsp;&nbsp;"," style=""display:none;"" "))
		        FolderRS.MoveNext
	            loop
              end if
               Response.Write(GetJSList(0,"0","",""))
              %>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr ParentID=""> 
          <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr onClick="ShowLabelTree('JSLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                <td width="672"><a href="#">自由JS标签</a></td>
              </tr>
              <tr> 
                <td colspan="2">
				  <table width="100%" border="0" cellspacing="0" cellpadding="0" id="JSLabel" style="display:none">
                    <tr> 
                      <td width="8%" align="right">&nbsp;</td>
                      <td height="20"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <%
                          SqlStr = "Select * From KS_LabelFolder where FolderType=3 And ParentID='0'"
                         Set FolderRS = Conn.Execute(SqlStr)
                           if Not FolderRS.Eof then
	                    do while Not FolderRS.Eof
                           %>
                          <tr ParentID="<% = FolderRS("ParentID") %>" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                            <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td width="3%"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                                  <td width="97%"><span ShowFlag="False" TypeID="<% = FolderRS("ID") %>" onClick="SelectFolder(this)"><A href="#"> 
                                    <% = FolderRS("FolderName") %>
                                    </A></span></td>
                                </tr>
                              </table></td>
                          </tr>
                          <%
	 		        Response.Write(GetChildFolderList(1,1,FolderRS("ID"),""," style=""display:none;"" "))
                    Response.Write(GetJSList(1,trim(FolderRS("ID")),"&nbsp;&nbsp;&nbsp;&nbsp;"," style=""display:none;"" "))
		        FolderRS.MoveNext
	            loop
              end if
               Response.Write(GetJSList(1,"0","",""))
              %>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table>
	 
</td>
  </tr>
  <tr>
    <td height="90" valign="top">
	<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="22" bgcolor="#0000FF"> 
            <div align="center"><font color="#FFFFFF"><strong>标签说明</strong></font></div></td>
        </tr>
        <tr> 
          <td valign="top" bgcolor="#efefef"> 
            <table width="272" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="143" height="25"> <img src="../Images/label2.gif" width="17" height="15"> 
                  网站通用标签</td>
                <td width="129" height="25"><img src="../Images/label1.gif" width="17" height="15"> 
                  频道内通用标签</td>
              </tr>
              <tr> 
                <td><img src="../Images/label0.gif"> 普通（函数）标签</td>
                <td><img src="../Images/label3.gif"> 自定义静态标签 </td>
              </tr>
              <tr> 
                <td><img src="../Images/JS0.gif" align="absmiddle"> 系统JS标签</td>
                <td><img src="../Images/JS1.gif" align="absmiddle"> 自由JS标签</td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<%
Set Conn = Nothing
Set KS=Nothing
Set KSCls=Nothing
Function GetLabelList(LabelType,TypeID,CompatStr,ShowStr)
	Dim ListSql,LabelRS
	ListSql = "Select * from KS_Label where LabelType=" & LabelType &" And FolderID='" & Trim(TypeID) & "' ORDER BY LabelFlag Desc"
	Set LabelRS = Conn.Execute(ListSql)
	IF LabelRS.EOF AND LabelRS.BOF THEN
       GetLabelList=""	 
	   LabelRS.close:Set LabelRS=Nothing
	  EXIT Function
	END IF
	do while Not LabelRS.Eof
	  	GetLabelList = GetLabelList & "<tr onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#EEFFFF'"" ParentID=""" & LabelRS("FolderID") & """ " & ShowStr & ">" & vbcrlf
		GetLabelList = GetLabelList & "<td height=22>" & vbcrlf
		GetLabelList = GetLabelList & "<table border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td>" & CompatStr &  "<img src=""../Images/Label" & trim(LabelRS("LabelFlag")) & ".gif""></td>"
		If LabelType=5 Then
		 GetLabelList = GetLabelList & "<td><A href=""#"" onclick=""InsertFunctionLabel('"&DomainStr&"Editor/KSPlus/InsertFunctionfield.asp?ID=" & Trim(LabelRS("ID")) & "',300,350)"">" & LabelRS("LabelName") & "</A></td></tr></table>"
		Else
		GetLabelList = GetLabelList & "<td><A href=""#"" onclick=""InsertLabel('" & Trim(LabelRS("LabelName")) & "')"">" & LabelRS("LabelName") & "</A></td></tr></table>"
		End If
		GetLabelList = GetLabelList & "</td>" & vbcrlf
		GetLabelList = GetLabelList & "</tr>" & vbcrlf
		LabelRS.MoveNext
	Loop
	Set LabelRS = Nothing
End Function
Function GetJSList(JSType,TypeID,CompatStr,ShowStr)
	Dim ListSql,JSRS
	ListSql = "Select * from KS_JSFile where JSType=" & JSType &" And FolderID='" & Trim(TypeID) & "' ORDER BY AddDate Desc"
	Set JSRS = Conn.Execute(ListSql)
	IF JSRS.EOF AND JSRS.BOF THEN
       GetJSList=""	 
	   JSRS.close
	   Set JSRS=Nothing
	  EXIT Function
	END IF
	do while Not JSRS.Eof
	  	GetJSList = GetJSList & "<tr onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#EEFFFF'"" ParentID=""" & JSRS("FolderID") & """ " & ShowStr & ">" & vbcrlf
		GetJSList = GetJSList & "<td height=22>" & vbcrlf
		GetJSList = GetJSList & "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf & "<tr>"  & vbcrlf & "<td>" & CompatStr &  "<img src=""../Images/JS" & trim(JSType) & ".gif""></td>"
		GetJSList = GetJSList & "<td><A href=""#"" onclick=""InsertLabel('" & Trim(JSRS("JSName")) & "')"">" & JSRS("JSName") & "</A></td>" & vbcrlf & "</tr>" & vbcrlf & "</table>"
		GetJSList = GetJSList & "</td>" & vbcrlf
		GetJSList = GetJSList & "</tr>" & vbcrlf
		JSRS.MoveNext
	Loop
	Set JSRS = Nothing
End Function
Function GetChildFolderList(GetType,LabelType,TypeID,CompatStr,ShowStr)
	Dim ChildFolderRS,ChildTypeListStr,TempStr
	Set ChildFolderRS = Conn.Execute("Select * FROM KS_LabelFolder where ParentID='" & TypeID & "'")
	TempStr = CompatStr & "&nbsp;&nbsp;&nbsp;&nbsp;"
	do while Not ChildFolderRS.Eof
	  	GetChildFolderList = GetChildFolderList & "<tr onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#EEEEEE'"" TypeFlag=""Class"" ParentID=""" & ChildFolderRS("ParentID") & """ " & ShowStr & ">" & vbcrlf
		GetChildFolderList = GetChildFolderList & "<td>" & vbcrlf
		GetChildFolderList = GetChildFolderList & "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf & "<tr>"  & vbcrlf & "<td>" & TempStr & "<img src=""../Images/Folder/folderclosed.gif""></td>"
		GetChildFolderList = GetChildFolderList & "<td><span TypeID=""" & ChildFolderRS("ID") & """ ShowFlag=""False"" onClick=""SelectFolder(this)""><a href=""#"">" & ChildFolderRS("FolderName") & "</a></span></td>" & vbcrlf & "</tr>" & vbcrlf & "</table>"
		GetChildFolderList = GetChildFolderList & "</td>" & vbcrlf
		GetChildFolderList = GetChildFolderList & "</tr>" & vbcrlf
		IF GetType=0 Then
		  GetChildFolderList = GetChildFolderList & vbcrlf & GetLabelList(LabelType,trim(ChildFolderRS("ID")),"&nbsp;&nbsp;&nbsp;&nbsp;" & TempStr,ShowStr) 
		Else
		  GetChildFolderList = GetChildFolderList & vbcrlf & GetJSList(LabelType,trim(ChildFolderRS("ID")),"&nbsp;&nbsp;&nbsp;&nbsp;" & TempStr,ShowStr) 
		End IF
		GetChildFolderList = GetChildFolderList & GetChildFolderList(GetType,LabelType,ChildFolderRS("ID"),TempStr,ShowStr)
		ChildFolderRS.MoveNext
	loop
	ChildFolderRS.Close
	Set ChildFolderRS = Nothing
End Function
%>
<script language="JavaScript">
function ShowLabelTree(Obj)
{
 switch (Obj)
  {case 'TY':
     if (document.all.TY.style.display!='')
       {document.all.TY.style.display='';}
     else
      {document.all.TY.style.display='none';} 
	  break;
	case 'CommonJSLabel':
     if (document.all.CommonJSLabel.style.display!='')
       {document.all.CommonJSLabel.style.display='';}
     else
      {document.all.CommonJSLabel.style.display='none';} 
	  break;
    case 'ChannelClassLabel':
     if (document.all.ChannelClassLabel.style.display!='')
       {document.all.ChannelClassLabel.style.display='';}
     else
      {document.all.ChannelClassLabel.style.display='none';} 
	  break;
   case 'SearchLabel':
        if (document.all.SearchLabel.style.display!='')
       {document.all.SearchLabel.style.display='';}
     else
      {document.all.SearchLabel.style.display='none';} 
	  break;
  <%For K=0 To Ubound(SQL,2)%>
   case 'Content<%=SQL(6,K)%>':
     if (document.all.Content<%=SQL(6,K)%>.style.display!='')
       {document.all.Content<%=SQL(6,K)%>.style.display='';}
     else
      {document.all.Content<%=SQL(6,K)%>.style.display='none';} 
	  break;
   <%Next%>
  case 'MusicLabel':
     if (document.all.MusicLabel.style.display!='')
       {document.all.MusicLabel.style.display='';}
     else
      {document.all.MusicLabel.style.display='none';} 
	  break;
   case 'AnnounceContent':
     if (document.all.AnnounceContent.style.display!='')
       {document.all.AnnounceContent.style.display='';}
     else
      {document.all.AnnounceContent.style.display='none';} 
	  break;
   case 'SysFLabel' :
   	  if (document.all.SysFLabel.style.display!='')
       {document.all.SysFLabel.style.display='';}
     else
      {document.all.SysFLabel.style.display='none';} 
	  break;
  case 'FreeLabel' :
      if (document.all.FreeLabel.style.display!='')
      {document.all.FreeLabel.style.display='';}
     else
      {document.all.FreeLabel.style.display='none';} 
	  break;
  case 'DIYFunctionLabel' :
      if (document.all.DIYFunctionLabel.style.display!='')
      {document.all.DIYFunctionLabel.style.display='';}
     else
      {document.all.DIYFunctionLabel.style.display='none';} 
	  break;  
  case 'DIYFieldLabel' :
      if (document.all.DIYFieldLabel.style.display!='')
      {document.all.DIYFieldLabel.style.display='';}
     else
      {document.all.DIYFieldLabel.style.display='none';} 
	  break; 
  case 'JSLabel' :
       if (document.all.JSLabel.style.display!='')
      {document.all.JSLabel.style.display='';}
     else
      {document.all.JSLabel.style.display='none';} 
	  break; 	   	  
  case 'SysJS' :
        if (document.all.SysJS.style.display!='')
      {document.all.SysJS.style.display='';}
     else
      {document.all.SysJS.style.display='none';} 
	  break; 
  case 'LinkContent':	   
        if (document.all.LinkContent.style.display!='')
      {document.all.LinkContent.style.display='';}
     else
      {document.all.LinkContent.style.display='none';} 
	  break; 
 case 'UserSystem':
      if (document.all.UserSystem.style.display!='')
      {document.all.UserSystem.style.display='';}
     else
      {document.all.UserSystem.style.display='none';} 
	  break; 
 case 'AdwLabel':
      if (document.all.AdwLabel.style.display!='')
      {document.all.AdwLabel.style.display='';}
     else
      {document.all.AdwLabel.style.display='none';} 
	  break; 
 case 'VoteLabel':
      if (document.all.VoteLabel.style.display!='')
      {document.all.VoteLabel.style.display='';}
     else
      {document.all.VoteLabel.style.display='none';} 
	  break; 
 case 'RssLabel':
      if (document.all.RssLabel.style.display!='')
      {document.all.RssLabel.style.display='';}
     else
      {document.all.RssLabel.style.display='none';} 
	  break; 
 case 'ContentLabel':
      if (document.all.ContentLabel.style.display!='')
      {document.all.ContentLabel.style.display='';}
     else
      {document.all.ContentLabel.style.display='none';} 
	  break; 
 case 'Special':
      if (document.all.Special.style.display!='')
      {document.all.Special.style.display='';}
     else
      {document.all.Special.style.display='none';} 
	  break; 
 }
}
function InsertLabel(LabelContent)
{
	window.returnValue=LabelContent;
	window.close();
}
function InsertFunctionLabel(Url,Width,Height)
{
window.returnValue = OpenWindow(Url,Width,Height,window);
window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
function SelectFolder(Obj)
{
	var CurrObj;
	if (Obj.ShowFlag=='True')
	{
		ShowOrDisplay(Obj,'none',true);
		Obj.ShowFlag='False';
	}
	else
	{
		ShowOrDisplay(Obj,'',false);
		Obj.ShowFlag='True';
	}
}
function ShowOrDisplay(Obj,Flag,Tag)
{
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if (CurrObj.ParentID==Obj.TypeID)
		{
			CurrObj.style.display=Flag;
			if (Tag) 
			if (CurrObj.TypeFlag=='Class') ShowOrDisplay(CurrObj.children(0).children(0).children(0).children(0).children(1).children(0),Flag,Tag);
		}
	}
}
</script> 
