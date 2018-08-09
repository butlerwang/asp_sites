<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.EscapeCls.asp"-->
<%


response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"
Dim KS:Set KS=New PublicCls
Dim ID:ID=KS.ChkClng(Request("id"))
If ID=0 Then KS.Die "error!"
Select Case Request("action")
 case "checklogin" checklogin
 case "savepost" savepost 
 case "getvotes" getvotes
 case "getgdlist" getgdlist
End Select

sub getgdlist()
  Dim role:role=ks.chkclng(ks.g("role"))
  Dim RS:Set RS=Server.CreateObject("adodb.recordset")
  dim xml,node,str
  rs.open "select top 30 username,userip,adddate,content,status from KS_PKGD where pkid=" & id & " and role=" & role & " order by id desc",conn,1,1
  if not rs.eof then
     set xml=ks.rstoxml(rs,"row","")
  end if
  rs.close
  set rs=nothing
  if isobject(xml) then
     dim n,content,UserIP,i,IpStr
	 n=0
     for each node in xml.documentelement.selectnodes("row")
	    UserIP=split(node.selectsinglenode("@userip").text,".")
		IpStr=""
		for i=0 to ubound(UserIP)
		   if i=3 then
		    ipstr=ipstr &"*"
		   else
		    ipstr=ipstr &UserIP(i)&"."
		   end if
		next
	   if node.selectsinglenode("@status").text="0" then
	    content="此观点未通过审核!"
	   else
	    content=node.selectsinglenode("@content").text
	   end if
	   content=replace(content,vbcrlf,"<br/>")
	   content=replace(content,Chr(13)&Chr(10),"<br/>")
	   content=replace(content,Chr(10),"<br/>")
	   str=str & "{""uname"":"""& node.selectsinglenode("@username").text &""",""comment_date"":""" & year(node.selectsinglenode("@adddate").text) & "-" & month(node.selectsinglenode("@adddate").text) & "-" & day(node.selectsinglenode("@adddate").text) &" " & hour(node.selectsinglenode("@adddate").text) & ":" & minute(node.selectsinglenode("@adddate").text) & """,""client_ip"":""" & ipstr & """,""comment_contents"":""" & content & """}"
	   n=n+1
	   if n<>xml.documentelement.selectnodes("row").length then str=str & ","
	 next
  end if
%>
var commentJsonVarStr___={"count":"3","comments":[
<%=str%>]};
 
<%
 if role=1 then
   ks.echo "showagree(commentJsonVarStr___);"
 elseif role=2 then
   ks.echo "showargue(commentJsonVarStr___);"
 else
   ks.echo "showother(commentJsonVarStr___);"
 end if
end sub

Sub checklogin()
   Dim Rs:Set Rs=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "Select top 1 * From KS_PKZT Where ID=" & ID,conn,1,1
   If RS.Eof And RS.Bof Then
      KS.Echo Escape("1|找不到PK主题!|null")
   Else
	   If RS("Status")="0" Then
		  KS.Echo Escape("1|该PK已锁定!|null")
	   ElseIf RS("TimeLimit")=1 And now>rs("enddate") Then
		  KS.Echo Escape("1|该PK已过期了!|null")
	   ElseIf KS.C("UserName")="" And RS("LoginTF")="1" Then
		  KS.Echo Escape("login||")
	   Else
	      KS.Echo "success||"
	   End If
  End If 
  RS.Close
  Set RS=Nothing
End Sub

Sub savepost()
   Dim Rs:Set Rs=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "Select top 1 * From KS_PKZT Where ID=" & ID,conn,1,1
   If RS.Eof And RS.Bof Then
      RS.Close: Set RS=Nothing
      KS.Echo Escape("1|找不到PK主题!|null")
	  Exit Sub
   Else
	   If RS("Status")="0" Then
		  KS.Echo Escape("1|该PK已锁定!|null")
	   ElseIf RS("TimeLimit")=1 And now>rs("enddate") Then
		  KS.Echo Escape("1|该PK已过期了!|null")
	   ElseIf KS.C("UserName")="" And RS("LoginTF")="1" Then
		  KS.Echo Escape("login||")
	   ElseIf RS("OnceTF")="1" and not conn.execute("select top 1 id from KS_PKGD where pkid=" & id & " and userip='" & KS.GetIP & "'").eof Then
		  KS.Echo Escape("1|您已PK过了,请不要重复PK!|null")
	   Else
	       Dim verify:verify=KS.ChkClng(rs("verifytf"))
	       RS.Close
		   Dim Content:Content=KS.DelSQL(UnEscape(request("Content")))
		   Dim Role:Role=KS.ChkClng(Request("role"))
		   If Content="" Then
		    KS.Echo Escape("1|请输入内容!|null")
		   Else
				  Dim UserName:UserName=KS.C("UserName")
				  If UserName="" Then UserName="网友"
				  RS.Open "select top 1 * from KS_PKGD",conn,1,3
				  RS.AddNew
					RS("PKID")=id
					RS("UserName")=UserName
					RS("UserIP")=KS.GetIP
					RS("Content")=content
					if verify=1 then
					RS("Status")=0
					else
					RS("Status")=1
					end if
					RS("AddDate")=Now
					RS("Role")=Role
				 RS.Update
				 If Role=1 Then
				  Conn.Execute("Update KS_PKZT Set ZFVotes=ZFVotes+1 where id=" & id)
				 ElseIf Role=2 Then
				  Conn.Execute("Update KS_PKZT Set FFVotes=FFVotes+1 where id=" & id)
				 Else
				  Conn.Execute("Update KS_PKZT Set SFVotes=SFVotes+1 where id=" & id)
				 End If
				 KS.Echo Escape("success||")
			    RS.Close
		   End If
	   End If
	   Set RS=Nothing
  End If 
End Sub

Sub getvotes()
     Dim Rs:Set Rs=Server.CreateObject("ADODB.RECORDSET")
	 RS.Open "Select top 1 * From KS_PKZT Where ID=" & ID,conn,1,1
	 If RS.Eof And RS.Bof Then
	    KS.Echo "0|0|0"
	 Else
	    KS.Echo rs("zfvotes") & "|" & rs("ffvotes") & "|" & rs("sfvotes")
	 End If
     RS.Close
	 Set RS=Nothing
End Sub
%>
