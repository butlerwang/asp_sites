
<%@ Language=VBScript %>
<%
dim msvr,muid,mpwd,mdb,mto
msvr=Request.Form("txtsvr")
muid=Request.Form("txtuid")
mpwd=Request.Form("txtpwd")
mdb=Request.Form("txtdb")
mto=Request.Form("txtto")
if mpwd="" then mpwd="''"

on error resume next
set dmosvr=server.CreateObject("SQLDMO.SQLServer")
dmosvr.connect msvr,muid,mpwd

if err.number>0 then Response.Redirect("http:backuperr.asp?err="&err.number)

mdevname="Backup_"&muid&"_"&mdb
set dmodev=server.CreateObject("SQLDMO.BackupDevice")
dmodev.name=mdevname
dmodev.type=2
dmodev.PhysicalLocation=mto
dmosvr.BackupDevices.Add dmodev

if err.number>0 then Response.Redirect("http:backuperr.asp?err="&err.number)

set dmobak=server.CreateObject("SQLDMO.Backup")
dmobak.database=mdb
dmobak.devices=mdevname
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<p><strong>���ڱ�������, ���Ե�...</strong></p>
<%
dmobak.sqlbackup dmosvr
if err.number>0 then Response.Redirect("http:backuperr.asp?err="&err.number)

dmosvr.backupdevices(mdevname).remove
set dmobak=nothing
set dmodev=nothing
dmosvr.disconnect
set dmosvr=nothing
%>
<p><strong>���� '<%=mdb%>' ���ݱ��ݳɹ�!</strong></p>
</BODY>
</HTML>