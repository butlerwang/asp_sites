<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%

Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SiteIndex
        Private KS, KSR,str,c_str,ClassID,Template,categoryname
		Private TotalPut,CurrentPage,MaxPerPage
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  MaxPerPage=10
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		   Dim I
		   ClassID=KS.ChkClng(KS.S("ClassID"))
	
		   Template = KSR.LoadTemplate(KS.Setting(102))
		   FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
		   Call GetPKList()
		   
		   Template=KSR.KSLabelReplaceAll(Template)
		   Response.Write Template  
		End Sub
		
		Sub GetPKList()
		   CurrentPage=KS.ChkClng(request("page"))
		   If CurrentPage=0 Then CurrentPage=1
		  dim rs,UserIP,ipstr,i,content,FaceStr,param
		  if ClassID<>0 then
		    param=" inner join ks_class b on a.classid=b.id where b.ClassID=" & classid
		  end if
		  set rs=server.createobject("adodb.recordset")
		  rs.open "select a.* from KS_PKZT a " & param & " order by a.id desc",conn,1,1
		   if rs.eof then
			 c_str=c_str & "没有PK主题！"
		   else
		 		    TotalPut= rs.recordcount
					If CurrentPage < 1 Then CurrentPage = 1
		
							If CurrentPage>1 and  (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
					 dim n:n=0
					 dim str,url,agreeNum ,argueNum,Total,zf,ff,m
					 m=(currentpage-1)*maxperpage
					Do While Not RS.Eof
					 n=n+1
					 m=M+1
					 agreeNum =rs("zfvotes")
					 argueNum = rs("ffvotes")
					 Total=agreeNum + argueNum+0.002
					 zf=formatpercent((agreeNum+0.001)/Total,2)
					 ff=formatpercent((argueNum+0.001)/Total,2)
					 
					 url="pk.asp?id=" & rs("id")
					 str=str &"<div class='listPk'>" & vbcrlf
					 str=str &"<table width='100%' border='0' cellspacing='0' cellpadding='0'>" & vbcrlf
					 str=str &"	  <tr valign='top'>" & vbcrlf
					 str=str & "	<td><h1><span class='PKdate'>" & rs("adddate") & "</span>  <a href='" & url & "' target='_blank'>" & rs("title") & "</a></h1>" &vbcrlf
					 
					 str=str & "	 <table border='0' cellspacing='0' cellpadding='0' height='19' class='number'>" &vbcrlf
					 str=str & "   <tr>" &vbcrlf
					 str=str & "	  <td width='35' align='center' valign='bottom'><a href='" & url & "' target='_blank'>观点A</a></td>" &vbcrlf
					 str=str &"		  <td width='60' align='center'><span class='red'>" & agreeNum & "</span></td>"
					 str=str & "	  <td width='83'><div class='exponentBj'><table width='70' border='0' align='center' cellpadding='0' cellspacing='0'>" &vbcrlf
					 str=str & "   <tr>" &vbcrlf
					 str=str & "   <td width='" & zf & "' style='border-right:1px #CC0000 solid;'><div class='zhengfang' style='width:100%;'></div></td>" &vbcrlf
					 str=str & "   <td width='" & ff & "'><div class='fanfang' style='width:100%;'></div></td>"&vbcrlf
					 
					 str=str & "  </tr>" &vbcrlf
					 str=str & "</table>" &vbcrlf
					 str=str & "</div></td>"
					 str=str & " <td width='60' align='center'><span class='LightGrey01'>" & argueNum &"</span></td>"&vbcrlf
					 str=str & " <td width='35' align='center' valign='bottom'><a href='" & url & "' target='_blank'>观点B</a></td>" &vbcrlf
					 str=str & "</tr>"
					 str=str & "</table>" &vbcrlf
					 str=str & "</td></tr>"
					 str=str & "</table>"
					 str=str & "</div>"
					 if n>=maxperpage or rs.eof then exit do
					  RS.MoveNext
					Loop
					RS.Close
					Set RS=Nothing
		  end if
		   Template=Replace(Template,"{$GetPKList}",str)
		   Template=Replace(Template,"{$ShowPage}","<div style='text-align:right'>" &  KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,false) & "</div>")
		   	
		End Sub
		
		
		
		
End Class
%>
