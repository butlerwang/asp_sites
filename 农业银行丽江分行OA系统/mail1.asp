
<%
Setmailer=Server.CreateObject("ASPMAIL.ASPMailCtrl.1")
name="ffff"
email="lj_lw@ynmail.com"
subject="欢迎您下次再来"
memo="欢迎您下次再来留言！"
mailserver="smtp.163.net"
result=mailer.SendMail(mailserver,name,email,subject,memo)
%>

