<%
'If Session("name") = "" then
'response.Write "<script LANGUAGE=javascript>alert('���糬ʱ�������㻹û�е�½! ');window.location.href='login.asp';</'script>"
'response.End
'end if
 '�����վ����ģ��
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
     str="<input name=""checkbox"" type=""checkbox"" value=""1"" checked>��ҵ��Ϣ"
   Case 2
     str="<input name=""checkbox"" type=""checkbox"" value=""2"" checked>�ͷ�����"
   Case 3
     str="<input name=""checkbox"" type=""checkbox"" value=""3"" checked>��Ѷ����"
   Case 4
     str="<input name=""checkbox"" type=""checkbox"" value=""4"" checked>��������"
   Case 5
     str="<input name=""checkbox"" type=""checkbox"" value=""5"" checked>Ȩ�޹���"
  Case 6
     str="<input name=""checkbox"" type=""checkbox"" value=""6"" checked>������Ƹ"
  Case 7
     str="<input name=""checkbox"" type=""checkbox"" value=""7"" checked>��������"
  Case 8
     str="<input name=""checkbox"" type=""checkbox"" value=""8"" checked>����Ԥ��"
  Case 9
     str="<input name=""checkbox"" type=""checkbox"" value=""9"" checked>¥�̱�־"
  end select
   Response.Write str
End Sub
%>