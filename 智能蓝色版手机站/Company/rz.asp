<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%

Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SiteIndex
        Private KS, KSR,Template,UserID,RS,RSC,Isqy
		Private TotalPut,MaxPerPage,CurrentPage
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		  MaxPerPage=10
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		%>
       <!--#include file="../KS_Cls/Kesion.IfCls.asp"-->
		<%
		Public Sub Kesion()
			If KS.S("page") <> "" Then
			  CurrentPage = CInt(Request("page"))
			Else
			  CurrentPage = 1
			End If
			UserID=KS.ChkClng(KS.S("UserID"))
			If UserID=0 Then KS.Die "error!"
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "select top 1 * From KS_User Where UserID=" & UserID,conn,1,1
			If RS.Eof And RS.Bof Then
			  RS.CLose:Set RS=Nothing
			  KS.Die "用户不存在！"
			End If

			Template = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & "企业空间/rz.html")
			FCls.RefreshType = "rzlist" '设置刷新类型，以便取得当前位置导航等
			FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
			
			isqy=false
			Set RSC=Server.CreateObject("ADODB.RECORDSET")
			RSC.Open "select top 1 * From KS_Enterprise Where UserName='" & RS("UserName") &"'",conn,1,1
			If Not RSC.Eof Then
			  isQy=true
			  Template=Replace(Template,"{$GetCompanyName}",RSC("CompanyName"))
			  Template=Replace(Template,"{$GetRegisteredCapital}",RSC("RegisteredCapital"))
			  Template=Replace(Template,"{$GetBusinessLicense}",RSC("BusinessLicense"))
			  Template=Replace(Template,"{$GetProvince}",RSC("Province"))
			  Template=Replace(Template,"{$GetCity}",RSC("City"))
			  Template=Replace(Template,"{$GetTelPhone}",RSC("Telphone"))
			  Template=Replace(Template,"{$GetLegalPeople}",RSC("LegalPeople"))
			  Template=Replace(Template,"{$GetFoundation}",RSC("Foundation"))
			  Template=Replace(Template,"{$GetBusiness}",RSC("Business"))
			  If IsDate(RSC("RZSJ")) Then
			  Template=Replace(Template,"{$GetRZSJ}",formatdatetime(RSC("RZSJ"),2))
			  End If


			  Template=Replace(Template,"{$GetContactMan}",LFCls.ReplaceDBNull(RSC("contactman"),"---"))
			  Template=Replace(Template,"{$GetMobile}",LFCls.ReplaceDBNull(RSC("mobile"),"---"))
			  Template=Replace(Template,"{$GetEmail}",LFCls.ReplaceDBNull(RSC("email"),"---"))
			  Template=Replace(Template,"{$GetAddress}",LFCls.ReplaceDBNull(RSC("address"),"---"))
			  Template=Replace(Template,"{$GetZipCode}",LFCls.ReplaceDBNull(RSC("zipcode"),"---"))
			  Template=Replace(Template,"{$GetQQ}",LFCls.ReplaceDBNull(RSC("qq"),"---"))
			  Template=Replace(Template,"{$GetCompanyScale}",LFCls.ReplaceDBNull(RSC("CompanyScale"),"---"))
			  Template=Replace(Template,"{$GetCompanyIntro}",LFCls.ReplaceDBNull(KS.HtmlCode(RSC("intro")),"---"))
			  Template=Replace(Template,"{$GetWebSite}",LFCls.ReplaceDBNull(RSC("weburl"),"http://"))
			  Template=Replace(Template,"{$GetUserID}",RS("userid"))
			Else
			  If KS.IsNul(RS("RealName")) Then
			  Template=Replace(Template,"{$GetCompanyName}",RS("UserName"))
			  Template=Replace(Template,"{$GetRealName}",RS("UserName"))
			  Else
			  Template=Replace(Template,"{$GetCompanyName}",RS("RealName"))
			  Template=Replace(Template,"{$GetRealName}",RS("RealName"))
			  End If
			  Template=Replace(Template,"{$GetTelphone}",LFCls.ReplaceDBNull(RS("officetel"),"---"))
			  Template=Replace(Template,"{$GetAddress}",LFCls.ReplaceDBNull(RS("address"),"---"))
			  Template=Replace(Template,"{$GetEmail}",LFCls.ReplaceDBNull(RS("email"),"---"))
			  Template=Replace(Template,"{$GetUserID}",LFCls.ReplaceDBNull(RS("userid"),"---"))
			End If
			RSC.Close
			
			Template=RexHtml_IF(Template)
			
			Dim RZInfo
			If isqy Then
			  If RS("IsRz")=1 Then 
			   RZInfo="<span class=""zzyz"" title=""此用户已经通过营业执照验证""><font>营业执照已验证</font></span>"
			  Else
			   RZInfo="<span class=""zzyzw"" title=""此用户营业执照未认证""><font style='color:red'>营业执照未认证</font></span>"
			  End If
			End If
			
			If RS("IsSfzRz")=1 Then 
			 RZInfo=RZInfo & "<span class=""nameyz"" title=""此用户已经通过身份证实名验证""><font>身份证已验证</font></span>"
			Else
			 RZInfo=RZInfo & "<span class=""nameyzw"" title=""此用户未通过身份证实名验证""><font style='color:red'>身份证未认证</font></span>"
			End If
			
			If RS("IsMobileRz")=1 Then 
			  RZInfo=RZInfo & "<span class=""sjyz"" title=""此用户已通过手机验证""><font>手机已验证</font></span>"
			Else
			  RZInfo=RZInfo & "<span class=""sjyzw"" title=""此用户未通过手机验证""><font style='color:red'>手机未认证</font></span>"
			End If
			
			If RS("IsEmailRz")=1 Then 
			 RZInfo=RZInfo & "<span class=""maileyz"" title=""此用户已通过邮箱验证""><font>邮箱已验证</font></span>"
			Else
			 RZInfo=RZInfo & "<span class=""maileyzw"" title=""此用户未通过邮箱验证""><font style='color:red'>邮箱未认证</font></span>"
			End If
			
			Template=Replace(Template,"{$GetRZInfo}",RZInfo)
			Template=Replace(Template,"{$GetZSNum}",Conn.Execute("Select count(1) From KS_EnterPriseZS Where UserName='" & RS("UserName") & "'")(0))
			
			Dim Str
			RSC.Open "Select * From KS_EnterpriseZS Where UserName='" & RS("UserName") & "' and status=1 order by id",conn,1,1
			Do While Not RSC.Eof
			  str=str & "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" &vbcrlf
			  str=str & "<tr><th rowspan=""6"" align=""center"" valign=""middle""><div class=""yzimg""><a href='" & rsc("photourl") & "' target='_blank'><img src=""" & rsc("photourl") & """ width=""254px"" height=""184px"" /></a></div></th><td>" & vbcrlf
			  str=str & "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""da_info"">" & vbcrlf
			  str=str & "<tr><td>证书名称： &nbsp;&nbsp;" & rsc("title") & "</td></tr>" &vbcrlf
			  str=str & "<tr><td>发证机关： &nbsp;&nbsp;" & rsc("fzjg") & "</td></tr>" &vbcrlf
			  str=str & "<tr><td>生效日期： &nbsp;&nbsp;" & rsc("sxrq") & "</td></tr>" &vbcrlf
			  str=str & "<tr><td>截止日期： &nbsp;&nbsp;" & rsc("jzrq") & "</td></tr>" &vbcrlf
			  str=str &"	</table></td></tr></table><br/>"
			RSC.MoveNext
			Loop
			RSC.Close
			Set RSC=Nothing
			
			Template=Replace(Template,"{$GetRYZSInfo}",str)
			
			RS.Close : Set RS=Nothing
			
			Template=KSR.KSLabelReplaceAll(Template)
		    Response.Write Template  
		End Sub
		
		
		
End Class
%>
