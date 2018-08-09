<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="config.asp"-->
<%

Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SiteIndex
        Private KS, KSR,str,c_str,curr_tips,pid,ads_str,s_str,astr,classname
		Private TotalPut,MaxPerPage,CurrentPage,Template,Province,City,Key
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
		Public Sub Kesion()
					If KS.S("page") <> "" Then
					  CurrentPage = CInt(Request("page"))
					Else
					  CurrentPage = 1
					End If
					Pid=KS.ChkClng(KS.S("Pid"))

				   Template = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & "企业空间/company_list.html")
				   FCls.RefreshType = "enterpriselist" '设置刷新类型，以便取得当前位置导航等
				   FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
				   call getclasslist()
				   call getcompanylist()
				   call getadslist()
				   call getsearchlist()
				   Template=Replace(Template,"{$ShowSmallClass}",str)
				   Template=Replace(Template,"{$ShowCurrTips}",curr_tips)
				   Template=Replace(Template,"{$ShowAds}",ads_str)
				   Template=Replace(Template,"{$ShowCompanyList}",c_str)
				   Template=Replace(Template,"{$ShowSearch}",s_str)
				   call getarealist()
		           Template=Replace(Template,"{$ShowAreaList}",astr)
				   Template=Replace(Template,"{$ShowClassName}",classname)
                   Template=Replace(Template,"{$ShowPage}",KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,false))
				   Template=KSR.KSLabelReplaceAll(Template)
		           Response.Write Template  
		End Sub
		
		Sub getadslist()
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 Dim Param:Param="Where adwz='1' and status=1"
		 If pid<>0 Then
		   Param=Param & " and Adtype=1 and bigclassid=" & pid
		 Else
		   Param=Param & " and Adtype=2 and smallclassid=" & ks.chkclng(ks.s("id"))
		 End If
		  Param=Param & " and datediff(" & DataPart_D & ",BeginDate," &SqlNowString & ")<datatimed"
		 rs.cursorlocation = 3
		 RS.Open "select top 1 * from KS_EnterPriseAD " & Param & " order by showtime",conn,1,3
         If RS.Eof Then
		  ads_str="<a href='" & default_ad_linkurl & "' target='_blank'><img border='0' src='" & default_ad_imgurl & "' width='" & default_ad_width & "' height='" & default_ad_height & "'></a>"
		 Else
		  ads_str="<a href='" & rs("url") & "' target='_blank'><img src='" & RS("PhotoUrl") & "' width='" & default_ad_width & "' height='" & default_ad_height & "'></a>"
		 End If
		 If Not RS.Eof Then
		 RS("ShowTime")=Now
		 RS.Update
		 End If
		 RS.Close:Set RS=Nothing
		End Sub
		
		
		Sub GetSearchList()
		  s_str="<form action='?' name='psform' method='get'>"
		  s_str=s_str & "企业搜索：<input type='text' name='key' size='30'>"
		  s_str=s_str & "&nbsp;<select name='pid'>"
		  dim rs:set rs=conn.execute("select id,classname from ks_enterpriseclass where parentid=0 order by orderid")
		  do while not rs.eof
		   if pid=rs(0) then
		   s_str=s_str & "<option value='" & rs(0) & "' selected>" & rs(1) & "</option>"
		   else
		   s_str=s_str & "<option value='" & rs(0) & "'>" & rs(1) & "</option>"
		   end if
		  rs.movenext
		  loop
		  s_str=s_str & "</select>&nbsp;<input onclick=""if(document.psform.key.value==''){alert('请输入关键字!');document.psform.key.focus();return false;}"" type='image' src='../images/btn2.gif' align='absmiddle'>"
		  rs.close:set rs=nothing
		  s_str=s_str & "</form>"
		End Sub
		
		Sub GetClassList()
		 Dim RS,I
		 Province=KS.CheckXSS(KS.S("Province"))
		 City=KS.CheckXSS(KS.S("City"))
		 Key=KS.CheckXSS(KS.S("Key"))
		 if Pid<>0 then
		     classname=LFCls.GetSingleFieldValue("Select classname from ks_enterpriseclass where id=" & Pid)
		     FCls.LocationStr="<a href='?pid=" & pid &"'>" & classname &"</a>"
		     curr_tips="<b>&nbsp;&nbsp;""<font color=#ff6600>" & classname & "</font>""&nbsp;&nbsp;的企业共 " & conn.execute("select count(id) from ks_enterprise where classid=" & pid)(0) &"  条</b>"
			 Set RS=Conn.Execute("select id,classname from ks_enterpriseclass where parentid=" & Pid & " order by orderid")
			 str="<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
			 Do While Not RS.Eof
			 str=str & "<tr>" & vbcrlf
			 for i=1 to 5
			   str=str & "<td width=""20%"" style=""padding:0px"">" & vbcrlf
			   str=str & "<div style=""height:26px;""><a href=""list.asp?id=" & rs(0) & """>" & rs(1) &"</a>(" & conn.execute("select count(id) from ks_enterprise where status=1 and smallclassid=" & rs(0))(0) &") </div>" & vbcrlf
	
			   str=str & "</td>" & vbcrlf
			   rs.movenext
			   if rs.eof then exit for
			 next
			 str=str & "</tr>"
			 Loop
			 str=str & "</table>" & vbcrlf
			 Template=replace(Template,"{$ShowTitleTips}",classname)
		 elseif ks.s("id")<>"" then
		   classname=LFCls.GetSingleFieldValue("Select classname from ks_enterpriseclass where id=" &ks.chkclng(ks.s("id")))
		   dim parentid:parentid=conn.execute("select parentid from ks_enterpriseclass where id=" &ks.chkclng(ks.s("id")))(0)
		   FCls.LocationStr="<a href='?pid=" & parentid & "'>" & LFCls.GetSingleFieldValue("Select classname from ks_enterpriseclass where id=" & parentid) &"</a> >> " & classname &"</a>"
		   curr_tips="<b>&nbsp;&nbsp;""<font color=#ff6600>" & classname & "</font>""&nbsp;&nbsp;的企业共 " & conn.execute("select count(id) from ks_enterprise where status=1 and smallclassid=" & ks.chkclng(ks.s("id")))(0) &"  条</b>"
		   Template=replace(Template,"{$ShowTitleTips}",classname)
		 elseif province<>"" then
		    classname=province& City
		    Dim total
			if KS.IsNul(City) then
			total=conn.execute("select count(id) from ks_enterprise where status=1 and province='" & province & "'")(0)
			else
			total=conn.execute("select count(id) from ks_enterprise where status=1 and city='" & city & "' and province='" & province & "'")(0)
			end if
		    FCls.LocationStr="地区<font color=red>""" & Province& City & """</font>"
			curr_tips="<b>&nbsp;&nbsp;所在地""<font color=#ff6600>" & Province & City & "</font>""&nbsp;&nbsp;的企业共 " &total &"  条</b>"
            Template=replace(Template,"{$ShowTitleTips}",province & city)
		 elseif Not KS.IsNul(key) Then
		    classname=key
			total=conn.execute("select count(id) from ks_enterprise where companyname like '%" &key & "%'")(0)
		    FCls.LocationStr="关键字<font color=red>""" & key & """</font>"
			curr_tips="<b>&nbsp;&nbsp;关键字""<font color=#ff6600>" & key & "</font>""&nbsp;&nbsp;的企业共 " &total &"  条</b>"
            Template=replace(Template,"{$ShowTitleTips}",key)
		 end if
		End Sub
		
		Sub GetCompanyList()
		 Dim SortStr,Param:Param=" where a.status=1"
		 if Not KS.IsNul(Province) Then
		  Param=Param & " and a.province='" & Province& "'"
		 ElseIf Not KS.IsNul(Key) Then
		   Param=Param & " and companyname like '%" & key & "%'"
		   if pid<>0 then
		  Param=Param & " and a.classid=" & Pid & " "
		  elseif KS.ChkClng(KS.S("ID"))<>0 Then
		  Param=Param & " and a.smallclassid=" & KS.ChkClng(KS.S("ID")) & " "
		  end if
		 Else
		  if pid<>0 then
		  Param=Param & " and a.classid=" & Pid & " "
		  else
		  Param=Param & " and a.smallclassid=" & KS.ChkClng(KS.S("ID")) & " "
		  end if
		 End If
		 If Not KS.IsNul(City) Then
		  Param=Param & " and a.city='" & City & "'"
		 End If
		 If KS.S("Recommend")="1" Then Param =Param & " and a.recommend=1"
		 
		 if ks.g("t")="1" then
		 SortStr = " order by a.adddate"
		 else
		 SortStr = " order by a.viptf desc,a.placevalue desc,a.id desc"
		 end if
		 Dim SqlStr:SqlStr="select b.logo,[Domain],a.username,a.companyname,a.province,a.city,a.recommend,a.intro,a.RegisteredCapital,a.address,a.telphone,a.viptf,a.placevalue,b.userid,a.isrz from ks_enterprise a inner join ks_blog b on a.username=b.username" & Param & SortStr
		 
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 rs.open SqlStr,conn,1,1
		 IF RS.Eof And RS.Bof Then
			  totalput=0
			  exit sub
		  Else
							TotalPut= Conn.Execute("Select count(*) from KS_Enterprise a " & Param)(0)
							If CurrentPage < 1 Then CurrentPage = 1
		
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
							Call ShowContent(RS)
			End IF
			
			c_str =c_str & "<div style='text-align:center'>" &  KS.ShowPagePara(totalPut, MaxPerPage, "", true, "家", CurrentPage, KS.QueryParam("page")) & "</div>"
			
			RS.Close
			Set RS=Nothing
		End Sub
		
		Sub ShowContent(RS)
		
		 Dim I,logo,n,url,msgUrl
 		 c_str="<div class=""productorder""><a href='?pid=" & Pid & "&id="&ks.g("id") & "&province=" & server.URLEncode(ks.s("province")) & "&provinceid=" &ks.s("provinceid") & "'>默认排序</a> <a href='?recommend=1&province=" & server.URLEncode(ks.s("province")) & "&provinceid=" &ks.s("provinceid") & "&pid=" & Pid & "&id="&ks.g("id") & "'>推荐企业</a> <a href='?t=1&pid=" & Pid & "&id="&ks.g("id") & "&province=" & server.URLEncode(ks.s("province")) & "&provinceid=" &ks.s("provinceid") & "'>加盟时间</a></div>"

		 c_str=c_str & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"">" & vbcrlf
         c_str=c_str & "<tr bgcolor=""#E7E7E7"">"
         c_str=c_str & "<td width=""111"" height=""26"" align=""center"">企业标志</td>"
         c_str=c_str & "<td width=""300"" align=""center"">公司/简要介绍</td>"
         c_str=c_str & "<td width=""100"" align=""center"">所在地区</div></td>"
         c_str=c_str & "<td width=""90"" align=""center"">注册资本(元)</td>"
        ' c_str=c_str & "<td width=""70"" align=""center"">诚信企业</td>"
        c_str=c_str & "<td width=""85"" align=""center"">积分排名</td>"
         c_str=c_str & "</tr>"
		 Do While Not RS.Eof
		 logo=trim(rs("logo"))
		 if logo="" or isnull(rs("logo")) then logo="/images/logo.jpg"
		 dim groupid:groupid=LFCls.GetSingleFieldValue("select groupid from ks_user where username='" & rs(2) & "'")
		 If KS.FoundInArr(KS.U_G(groupid,"powerlist"),"s01",",")=false Then
		     url="show.asp?username=" & rs(2)
			 msgUrl="#"
		 else
			 If KS.SSetting(14)<>"0" and rs(1)<>"" then 
			  if instr(rs(1),".")<>0 then
			  url="http://" & rs(1)
			  Else
			  url="http://" & rs(1) &"."& KS.SSetting(16)
			  End If
			  msgUrl="../space/?"&rs("userid") & "/message"
			 else
			  if KS.SSetting(21)="1" Then
			  url="../space/" & rs("userid") 
			  msgUrl="../space/"&rs("userid") & "/message"
			  Else
			  url="../space/?" & rs("userid")
			  msgUrl="../space/?"&rs("userid") & "/message"
			  end if
			 end if
		 end if
		 dim str
		 if rs("recommend")="1" then  str="<font color=green>荐</font>"	 else  str=""
		 if rs("isrz")="1" and IsBusiness then str=str & " <span style='font-size:12px;font-weight:normal;color:#999999'>[已实名认证]</span>"
		 

         n=n+1
		 if n mod 2=0 then
		 c_str=c_str & "<tr bgcolor=""#f6f6f6"">"
		 else
         c_str=c_str & "<tr>"
		 end if
         c_str=c_str & "<td width=""130"" height=""80"" align=""center""><a href='" & url & "' target='_blank'><img src=""" & logo & """ width=116 height=40 border='0'></a></td>"
         c_str=c_str & "<td width=""300"" style=""WORD-BREAK: break-all""><a href=""" & url & """ target=""_blank""><div style='font-weight:bold;font-size:14px;text-decoration:underline;margin:2px;'>" & RS("CompanyName") & " " & str &"</div></a>" & KS.Gottopic(KS.LoseHtml(KS.HtmlCode(RS("Intro"))),120) &"...<br/>公司地址：" & rs("address") & "<br/>联系电话："& rs("telphone") & "</td>"
         c_str=c_str & "<td width=""100"" align=""center"">" & RS("Province") & RS("City") & "</td>"
         c_str=c_str & "<td width=""90"" align=""center"">" & RS("RegisteredCapital") & "</td>"

		 c_str=c_str & "<td align='center'>" & rs("placevalue") & "</td>"
         c_str=c_str & "</tr>"
		 I=I+1
		If I >= MaxPerPage Then Exit Do
		 RS.MoveNext
		 Loop
         c_str=c_str & "</table>"
		End Sub
		Sub getarealist()
		  Dim RS,I,SQL,K,N
		  If Not KS.IsNul(Province) and KS.S("provinceid")<>"" then
		  Set RS=Conn.Execute("Select city from KS_Province where ParentID=" & KS.ChkClng(KS.S("Provinceid")) & " order by orderid")
		  Else
		  Set RS=Conn.Execute("Select id,city from KS_Province where parentid=0 order by orderid")
		  End If
		  
		  IF Not RS.Eof Then SQL=RS.GetRows(-1):RS.Close:Set RS=Nothing
		  If IsArray(SQL) Then
			  astr="<table border='0' width='100%'>" &vbcrlf
			  N=0
			  For i=0 To Ubound(SQL,2)
				astr=astr & "<tr>" &vbcrlf
				For K=1 To 3
				 If Province<>"" and KS.S("provinceid")<>"" then
				astr=astr & "<td><img src='../images/default/arrow_r.gif'> <a href=""list.asp?province=" & server.URLEncode(province) & "&provinceid=" & ks.s("provinceid") & "&city=" & server.URLEncode(sql(0,n)) & """>" & sql(0,n) & "</a></td>"
				 Else
				astr=astr & "<td><img src='../images/default/arrow_r.gif'> <a href=""list.asp?province=" & server.URLEncode(sql(1,n)) & "&provinceid=" & SQL(0,n) & """>" & sql(1,n) & "</a></td>"
				 End If
				n=n+1
				if n>Ubound(SQL,2) then Exit For
				Next
				astr=astr & "</tr>" &vbcrlf
				if n>Ubound(SQL,2) then Exit For
			 Next
			 astr=astr & "<tr><td colspan=3 align='center'><a href='index.asp'>返上一级导航</a></td></tr>"
			 astr=astr & "</table>" & vbcrlf
		 End If
		End Sub
End Class
%>
