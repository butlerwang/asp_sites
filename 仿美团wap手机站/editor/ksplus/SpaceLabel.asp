<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>空间标签</TITLE>
<META content="text/html; charset=utf-8" http-equiv=Content-Type>
<link href="editor.css" rel="stylesheet">
<style>
td{font-size:12px;}
body{background:#FFFFFF}
a{text-decoration:none;font-size:12px;color:#000000}
li{list-style-type:circle}
</style>
</HEAD>
<body>
<br>
<table align="center" width="95%" border="0" cellspacing="0" cellpadding="0">
		 <tr>
		  <td width="150" colspan=2><font color=red>可用标签说明</font></td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$BlogMain}');"><strong>{$BlogMain}</strong></a></td><td colspan=3>---显示日志主体部分(必须放在其它框架页模板)。</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowUserInfo}');">{$ShowUserInfo}</a></td><td>---显示用户信息。</td>
		     <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowComment}');">{$ShowComment}</a></td><td>---显示最新评论列表。</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowUserClass}');">{$ShowUserClass}</a></td><td>---显示专栏分类列表。</td>
		  <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowMessage}');">{$ShowMessage}</a></td><td>---显示最新留言列表。</td>
		  </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowBlogInfo}');">{$ShowBlogInfo}</a></td><td>---显示最新日志列表。</td>
		   <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowAnnounce}');">{$ShowAnnounce}</a></td><td>---显示最新公告。</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowBlogName}');">{$ShowBlogName}</a></td><td>---显示博客站点名称。</td>
		   <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowCalendar}');">{$ShowCalendar}</a></td><td>---显示日历搜索。</td>
		 </tr>
		
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowNavigation}');">{$ShowNavigation}</a></td><td>---显示首页导航等。</td>
		   <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowBlogTotal}');">{$ShowBlogTotal}</a></td><td>---显示统计信息等。</td>
		 </tr>
		 
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowMusicBox}');">{$ShowMusicBox}</a></td><td>---显示音乐播放器。</td>
		    <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowUserLogin}');">{$ShowUserLogin}</a></td><td>---显示会员登录框。</td>
		 </tr>
		
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowSearch}');">{$ShowSearch}</a></td><td>---显示搜索日志。</td>
		   <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowXML}');">{$ShowXML}</a></td><td>---显示RSS订阅。</td>
		 </tr>
		
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowUserName}');">{$ShowUserName}</a></td><td>---显示用户名。</td>
		   <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowUserID}');">{$ShowUserID}</a></td><td>---显示用户ID。</td>
		 </tr>
		 <tr>
		   <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowLogo}');">{$ShowLogo}</a></td><td>---显示Logo。</td>
		   <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowNewFresh}');">{$ShowNewFresh}</a></td><td>---显示1条新鲜事。</td>
		 </tr>
		
		 <tr>
	       <td colspan=4><li>{$ShowBannerSrc1},{$ShowBannerSrc2},{$ShowBannerSrc3}    ---显示Banner图片地址。</td>
		 </tr>
		
		 <tr>
		  <td width="150"><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowNewLog}');">{$ShowNewLog}</a></td><td>---最新4篇日志</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowNewAlbum}');">{$ShowNewAlbum}</a></td><td>---最新4张照片</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowNewInfo}');">{$ShowNewInfo}</a></td><td>---10条信息集</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowVisitor}');">{$ShowVisitor}</a></td><td>---最新访客</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowBlogDescript}');">{$ShowBlogDescript}</a></td><td>---博客描述</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowClubTopic}');">{$ShowClubTopic}</a></td><td>---列出10条论坛发表的最新话题</td>
		 </tr>

		<%if request("flag")="4" then%> 
		 
		 <tr>
	       <td colspan=4><br>================企业空间专用=======================</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowShortIntro}');" title="显示580个字的企业介绍">{$ShowShortIntro}</a></td><td>---企业简介(短)</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowIntro}');">{$ShowIntro}</a></td><td>---企业简介</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowNews}');">{$ShowNews}</a></td><td>---企业动态</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowSupply}');">{$ShowSupply}</a></td><td>---供应信息</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowProduct}');" title='一行显示4个，分两行显示最新产品'>{$ShowProduct}</a></td><td>---最新产品</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowProductList}');" title='纯文字显示最新产品'>{$ShowProductList}</a></td><td>---文本方式显示最新产品</td>
		 </tr>
		 <%end if%>
</table>

        <%response.end%>
		 <tr>
		  <td colspan=2><font color=red>副模板(日志)可用标签说明</font></td>
		 </tr>
		 <tr>
		  <td width="150"><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowLogTopic}');">{$ShowLogTopic}</a></td><td>---显示表情及日志标题</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowLogInfo}');">{$ShowLogInfo}</a></td><td>---显示发表时间、作者等信息</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowLogText}');">{$ShowLogText}</a></td><td>---显示日志正文</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowLogMore}');">{$ShowLogMore}</a></td><td>---显示阅读全文(次数)，回复(次数)，引用链接</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowTopic}');">{$ShowTopic}</a></td><td>---仅显示日志标题</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowAuthor}');">{$ShowAuthor}</a></td><td>---仅显示日志作者</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowAddDate}');">{$ShowAddDate}</a></td><td>---仅显示日志发布时间</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowEmot}');">{$ShowEmot}</a></td><td>---仅显示日志心情</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowWeather}');">{$ShowWeather}</a></td><td>---仅显示日志天气</td>
		 </tr>
		  <tr>
		  <td colspan=2><font color=red>副模板(个人档案)可用标签说明</font></td>
		 </tr>
		 <tr>
		  <td width="150"><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetUserName}');">{$GetUserName}</a></td><td>--用户名（昵称)</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetRealName}');">{$GetRealName}</a></td><td>---真实姓名</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetSex}');">{$GetSex}</a></td><td>---性别</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetBirthday}');">{$GetBirthday}</a></td><td>---出生日期</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetIDCard}');">{$GetIDCard}</a></td><td>---身份证号</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetOfficeTel}');">{$GetOfficeTel}</a></td><td>---办公电话</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetHomeTel}');">{$GetHomeTel}</a></td><td>---家庭电话</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetMobile}');">{$GetMobile}</a></td><td>---手机号码</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetFax}');">{$GetFax}</a></td><td>---传真号码</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetUserArea}');">{$GetUserArea}</a></td><td>---所在地区</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetAddress}');">{$GetAddress}</a></td><td>---联系地址</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetZip}');">{$GetZip}</a></td><td>---邮政编码</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetHomePage}');">{$GetHomePage}</a></td><td>---个人主页</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetUserFace}');">{$GetUserFace}</a></td><td>---用户头像</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetEmail}');">{$GetEmail}</a></td><td>---电子信箱</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetQQ}');">{$GetQQ}</a></td><td>---QQ 号码</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetICQ}');">{$GetICQ}</a></td><td>---ICQ号码</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetMSN}');">{$GetMSN}</a></td><td>---MSN账号</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetUC}');">{$GetUC}</a></td><td>---UC号码</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetSign}');">{$GetSign}</a></td><td>---个人签名</td>
		 </tr>
		 <tr>
		  <td colspan=2><font color=red>副模板(联系我们)可用标签说明</font></td>
		 </tr>
		 <tr>
		  <td width="150"><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetCompanyName}');">{$GetCompanyName}</a></td><td>--公司名称</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetBusinessLicense}');">{$GetBusinessLicense}</a></td><td>---营业执照</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetProfession}');">{$GetProfession}</a></td><td>---公司行业</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetOfficeTel}');">{$GetLegalPeople}</a></td><td>---企业法人</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetCompanyScale}');">{$GetCompanyScale}</a></td><td>---公司规模</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetRegisteredCapital}');">{$GetRegisteredCapital}</a></td><td>---注册资金</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetProvince}');">{$GetProvince}</a></td><td>---所在省份</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetCity}');">{$GetCity}</a></td><td>---所在城市</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetContactMan}');">{$GetContactMan}</a></td><td>---联 系 人</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetAddress}');">{$GetAddress}</a></td><td>---公司地址</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetZipCode}');">{$GetZipCode}</a></td><td>---邮政编码</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetTelphone}');">{$GetTelphone}</a></td><td>---联系电话</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetFax}');">{$GetFax}</a></td><td>---传真号码</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetWebUrl}');">{$GetWebUrl}</a></td><td>---公司网址</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetBankAccount}');">{$GetBankAccount}</a></td><td>---开户银行</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetAccountNumber}');">{$GetAccountNumber}</a></td><td>---银行账号</td>
		 </tr>

		 </table>

</BODY>
</HTML>
 
