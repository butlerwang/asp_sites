<!--#include file="check.asp"-->
<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨����ϵͳ</title>
<link href="include/style.css" rel="stylesheet" type="text/css">
</head>
<base target="main">
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
Call OpenData()
%>
<table width="100%" height="93" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="100%" height="93" background="images/top.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="34%" height="41">&nbsp;&nbsp;<img src="images/logo.gif" width="180" height="49"></td>
          <td width="40%">&nbsp;</td>
          <td width="26%">&nbsp;</td>
        </tr>
        <tr valign="middle"> 
          <td height="37" colspan="3"><table border="1" cellpadding="0" id="button_img">
            <tr>
			    <td><a href="index.asp" target="_top">��ҳ</a></td>
			  <%
				IF instr(webConfig," 5")>=1 Then'��վ��������
				  IF instr(session("manconfig")," 5")>=1 Then'��վ����Ȩ������	
				%>
                <td><a href="member/">Ȩ�޹���</a></td>
				<%
				  End IF
				End IF
				%>
				<%
				IF instr(webConfig," 1")>=1 Then'��վ��������
				  IF instr(session("manconfig")," 1")>=1 Then'��վ����Ȩ������	
				%>
                <td><a href="Company/">��ҵ��Ϣ</a></td>
				<%
				  End If
				End IF
				%>
				<%
				IF instr(webConfig," 2")>=1 Then'��վ��������
				  IF instr(session("manconfig")," 2")>=1 Then'��վ����Ȩ������	
				%>
                <td><a href="product/">�ͷ�����</a></td>
				<%
				  End IF
				End IF
				%>
				<%
				IF instr(webConfig," 3")>=1 Then'��վ��������
				  IF instr(session("manconfig")," 3")>=1 Then'��վ����Ȩ������	
				%>
                <td><a href="news/">��Ѷ����</a></td>
				<%
				  End If
				End IF
				%>
				<%
				IF instr(webConfig," 6")>=1 Then'��վ��������
				  IF instr(session("manconfig")," 6")>=1 Then'��վ����Ȩ������	
				%>
                <td><a href="job/">������Ƹ</a></td>
				<%
				  End If
				End IF
				%>
				<%
				IF instr(webConfig," 7")>=1 Then'��վ��������
				  IF instr(session("manconfig")," 7")>=1 Then'��վ����Ȩ������	
				%>
                <td><a href="ServeCenter/">��������</a></td>
				<%
				  End If
				End IF
				%>
				<%
				IF instr(webConfig," 4")>=1 Then'��վ��������
				  IF instr(session("manconfig")," 4")>=1 Then'��վ����Ȩ������	
				%>
                <td><a href="Down/">��������</a></td>
				<%
				  End If
				End IF
				%>
                <%
				IF instr(webConfig," 8")>=1 Then'��վ��������
				  IF instr(session("manconfig")," 8")>=1 Then'��վ����Ȩ������	
				%>
                <td><a href="orders/">����Ԥ��</a></td>             
				<%
				  End If
				End IF
				%>
                <%
				IF instr(webConfig," 9")>=1 Then'��վ��������
				  IF instr(session("manconfig")," 9")>=1 Then'��վ����Ȩ������	
				%>
                <td><a href="weblink/">¥�̱�־</a></td>             
				<%
				  End If
				End IF
				%>
				<%if Session("flag")=99 then%>
				<td><a href="manage/">��վ����</a></td>
				<%end if%>
			<td><%if session("name") ="" then%><a href="quit.asp"><span style="font-weight:bold;color:#F00; ">��¼</span></a><%else%><a href="quit.asp"><span style="font-weight:bold;color:#F00; ">�˳�</span></a><%end if%></td>
			</tr>
          </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<%Call CloseDataBase()%>
</body>
</html>