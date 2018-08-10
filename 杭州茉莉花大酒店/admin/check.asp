<%
'If Session("name") = "" then
'response.Write "<script LANGUAGE=javascript>alert('网络超时，或者你还没有登陆! ');window.location.href='login.asp';</'script>"
'response.End
'end if
 '检测网站功能模块
 Private Function webConfig()
   Set oRs_web=Server.CreateObject("adodb.recordset")
   sql="select Template from Sbe_WebConfig"
   oRs_web.open sql,conn,1,1
   IF oRs_web.eof and oRs_web.bof Then oRs_web.Close:set oRs_web=Nothing:Exit Function
   webConfig=oRs_web("Template")
   oRs_web.Close:set oRs_web=Nothing
 End Function
 Private Sub check_name(intID)   
 intID=intID
  select Case intID 
   Case 0
   Case 1
     str="<input name=""checkbox"" type=""checkbox"" value=""1"" checked>企业信息"
   Case 2
     str="<input name=""checkbox"" type=""checkbox"" value=""2"" checked>客房中心"
   Case 3
     str="<input name=""checkbox"" type=""checkbox"" value=""3"" checked>资讯中心"
   Case 4
     str="<input name=""checkbox"" type=""checkbox"" value=""4"" checked>店铺形象"
   Case 5
     str="<input name=""checkbox"" type=""checkbox"" value=""5"" checked>权限管理"
  Case 6
     str="<input name=""checkbox"" type=""checkbox"" value=""6"" checked>人事招聘"
  Case 7
     str="<input name=""checkbox"" type=""checkbox"" value=""7"" checked>在线留言"
  Case 8
     str="<input name=""checkbox"" type=""checkbox"" value=""8"" checked>在线预定"
  Case 9
     str="<input name=""checkbox"" type=""checkbox"" value=""9"" checked>楼盘标志"
  end select
   Response.Write str
End Sub
%>