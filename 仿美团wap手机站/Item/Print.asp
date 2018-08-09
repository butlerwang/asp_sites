<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%

Dim KS:Set KS=New PublicCls
Dim Channelid,ID,RS,ArticleContent,PayTF

ChannelID=KS.ChkClng(KS.S("M"))
ID=KS.ChkClng(KS.S("ID"))
if ID=0 Or ChannelID=0 then
	Response.Write"<script>alert(""错误的参数！"");location.href=""javascript:history.back()"";</script>"
    Response.End
end if
Set RS=Server.CreateObject("ADODB.RECORDSET")
RS.Open "Select a.*,ClassPurview From "& KS.C_S(ChannelID,2) & " a inner join ks_class b on a.tid=b.id Where a.ID=" & ID,Conn,1,1
IF RS.EOF AND RS.BOF THEN
  RS.CLOSE:SET RS=NOthing
  Call CloseConn()
  Set KS=Nothing
 	Response.Write"<script>alert(""错误的参数！"");location.href=""javascript:history.back()"";</script>"
    Response.End
END IF
	Dim InfoPurview:InfoPurview=Cint(RS("InfoPurview"))
	Dim ReadPoint:ReadPoint=Cint(RS("ReadPoint"))
	Dim ChargeType:ChargeType=Cint(RS("ChargeType"))
	Dim PitchTime:PitchTime=Cint(RS("PitchTime"))
	Dim ReadTimes:ReadTimes=Cint(RS("ReadTimes"))
	Dim ClassID:ClassID=RS("Tid")
	Dim KSUser:Set KSUser=New UserCls
	Dim UserLoginTF:UserLoginTF=Cbool(KSUser.UserLoginChecked)
	    
		 If ReadPoint>0 Then
			   IF UserLoginTF=false Then
				 Call GetNoLoginInfo
			   Else
				 Call PayPointProcess()
			   End If
		 ElseIf InfoPurview=2  Then 
			   IF UserLoginTF=false Then
				 Call GetNoLoginInfo
			   Else
					 IF InStr(RS("ArrGroupID"),KSUser.GroupID)=0 Then
					   ArticleContent="<div align=center>对不起，你没有查看本文的权限!</div>"
					 Else
						  Call PayPointProcess()
					 End If
			   End If
		 ElseIF InfoPurview=0 And (RS("ClassPurview")=1 Or RS("ClassPurview")=2) Then 
			  If UserLoginTF=false Then
			    Call GetNoLoginInfo
			  Else        
				Call PayPointProcess()
			  End If
		 Else
		   Call PayPointProcess()
		 End If   

	   '收费扣点处理过程
	   Sub PayPointProcess()
	     Dim UserChargeType:UserChargeType=KSUser.ChargeType
					   If Cint(ReadPoint)>0 Then
					     If UserChargeType=1 Then
							 Select Case ChargeType
							  Case 0:Call CheckPayTF("1=1")
							  Case 1:Call CheckPayTF("datediff('h',AddDate," & SqlNowString & ")<" & PitchTime)
							  Case 2:Call CheckPayTF("Times<" & ReadTimes)
							  Case 3:Call CheckPayTF("datediff('h',AddDate," & SqlNowString & ")<" & PitchTime & " or Times<" & ReadTimes)
							  Case 4:Call CheckPayTF("datediff('h',AddDate," & SqlNowString & ")<" & PitchTime & " and Times<" & ReadTimes)
							  Case 5:Call PayConfirm()
							  End Select
						Elseif UserChargeType=2 Then
				          If KSUser.GetEdays <=0 Then
						     ArticleContent="<div align=center>对不起，你的账户已过期 <font color=red>" & Edays & "</font> 天,此文需要在有效期内才可以查看，请及时与我们联系！</div>"
						  Else
						   Call GetContent()
						  End If
						Else
						 Call GetContent()
						end if
					   Else
						  Call GetContent()
					   End IF
	   End Sub
	   '检查是否过期，如果过期要重复扣点券
	   '返回值 过期返回 true,未过期返回false
	   Sub CheckPayTF(Param)
	    Dim SqlStr:SqlStr="Select top 1 Times From KS_LogPoint Where ChannelID=" & ChannelID & " And InfoID=" & ID & " And InOrOutFlag=2 and UserName='" & KSUser.UserName & "' And (" & Param & ") Order By ID"
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open SqlStr,conn,1,3
		IF RS.Eof And RS.Bof Then
			Call PayConfirm()	
		Else
		       RS.Movelast
			   RS(0)=RS(0)+1
			   RS.Update
			   Call GetContent()
		End IF
		 RS.Close:Set RS=nothing
	   End Sub
	   
	   Sub PayConfirm()
	     If UserLoginTF=false Then Call GetNoLoginInfo():Exit Sub
			 If Cint(KSUser.GetUserInfo("Point"))<ReadPoint Then
					 ArticleContent="<div align=center>对不起，你的可用" & KS.Setting(45) & "不足!阅读本文需要 <font color=red>" & ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) &",你还有 <font color=green>" & KSUser.GetUserInfo("Point") & "</font> " & KS.Setting(46) & KS.Setting(45) & "</div>,请及时与我们联系！" 
			 Else
					If PayTF="yes" Then
						IF Cbool(KS.PointInOrOut(ChannelID,RS("ID"),KSUser.UserName,2,ReadPoint,"系统","阅读收费" & KS.C_S(ChannelID,3) & "：<br>" & RS("Title"),0))=True Then Call GetContent()
					Else
						ArticleContent="<div align=center>阅读本文需要消耗 <font color=red>" & ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) &",你目前尚有 <font color=green>" & KSUser.GetUserInfo("Point") & "</font> " & KS.Setting(46) & KS.Setting(45) &"可用,阅读本文后，您将剩下 <font color=blue>" & KSUser.GetUserInfo("Point")-ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) &"</div><div align=center>你确实愿意花 <font color=red>" & ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) & "来阅读此文吗?</div><div>&nbsp;</div><div align=center><a href=""?ID=" & ID & "&PayTF=yes&Page=" & CurrPage &""">我愿意</a>    <a href=""" &DomainStr & """>我不愿意</a></div>"
					End If
			 End If
	   End Sub
	   Sub GetNoLoginInfo()
		   ArticleContent="<div align=center>对不起，你还没有登录，本文至少要求本站的注册会员才可查看!</div><div align=center>如果你还没有注册，请<a href=""../user/reg/""><font color=red>点此注册</font></a>吧!</div><div align=center>如果您已是本站注册会员，赶紧<a href=""../user/login/""><font color=red>点此登录</font></a>吧！</div>"
	   End Sub
	   Sub GetContent()
	     ArticleContent=Replace(RS("ArticleContent"),"[NextPage]","")
	   End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD><TITLE><%=rs("title")%>-打印文章</TITLE>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<SCRIPT language=JavaScript type=text/JavaScript>
function resizepic(thispic)
{
if(thispic.width>700) thispic.width=700;
}
//双击鼠标滚动屏幕的代码
var currentpos,timer;
function initialize()
{
timer=setInterval ("scrollwindow ()",30);
}
function sc()
{
clearInterval(timer);
}
function scrollwindow()
{
currentpos=document.body.scrollTop;
window.scroll(0,++currentpos);
if (currentpos !=document.body.scrollTop)
sc();
}
document.onmousedown=sc
document.ondblclick=initialize
</SCRIPT>
<META content="MSHTML 6.00.3790.2577" name=GENERATOR></HEAD>
<BODY onmouseup=document.selection.empty() oncontextmenu="return false" onselectstart="return false" ondragstart="return false"onbeforecopy="return false" oncopy=document.selection.empty() leftMargin=0 topMargin=0 onselect=document.selection.empty() marginheight="0" marginwidth="0">
<TABLE width=760 height="100%" border=0 align=center cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF" class=center_tdbgall style="WORD-BREAK: break-all">
  <TBODY>
  <TR>
    <TD class=main_title_760 align=right height=20><A class=class 
      href="javascript:window.print()"><IMG src="../Images/Default/printpage.gif" alt=打印本文 border=0 align=absMiddle>&nbsp;打印本文</A>&nbsp;&nbsp;<IMG alt=关闭窗口 src="../Images/Default/pageclose.gif" align=absMiddle border=0>&nbsp;<A 
      class=class href="javascript:window.close()">关闭窗口</A>&nbsp;&nbsp; </TD>
  </TR>
  <TR>
    <TD height="25" align=middle class=main_ArticleTitle><B><%=RS("Title")%></B></TD>
  </TR>
  <TR>
    <TD height="25" 
      align=middle class=Article_tdbgall>作者：<%=RS("Author")%>&nbsp;&nbsp;文章来源：<%=RS("Origin")%>&nbsp;&nbsp;点击数
      <%=RS("Hits")%>&nbsp;&nbsp;更新时间：<%=RS("AddDate")%>&nbsp;&nbsp;文章录入：<%=RS("Inputer")%>
  </TD></TR>
  <TR>
    <TD height="25">
      <HR align=center width="100%" color=#8ea7cd noShade SIZE=1>
    </TD>
  </TR>
  <TR>
    <TD valign="top"><%=KS.HtmlCode(ArticleContent)%></TD>
  </TR> 
</TBODY>
</TABLE>
</BODY>
</HTML> 
<%set ks=nothing
conn.close
set conn=nothing
%>