
<%
Setmailer=Server.CreateObject("ASPMAIL.ASPMailCtrl.1")
name="ffff"
email="lj_lw@ynmail.com"
subject="��ӭ���´�����"
memo="��ӭ���´��������ԣ�"
mailserver="smtp.163.net"
result=mailer.SendMail(mailserver,name,email,subject,memo)
%>

