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

mdevname="Restore_"&muid&"_"&mdb
dmosvr.backupdevices(mdevname).remove
err.clear

set dmodev=server.CreateObject("SQLDMO.BackupDevice")
dmodev.name=mdevname
dmodev.type=2
dmodev.PhysicalLocation=mto
dmosvr.BackupDevices.Add dmodev

if err.number>0 then Response.Redirect("http:backuperr.asp?err="&err.number)

set dmores=server.CreateObject("SQLDMO.Restore")
dmores.database=mdb
dmores.devices=mdevname
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<body class="bg_frame_up">

<p><strong>数据恢复中, 请稍等...</strong></p>
<%
dmores.sqlrestore dmosvr
if err.number>0 then Response.Redirect("http:backuperr.asp?err="&err.number)

set dmores=nothing
set dmodev=nothing
dmosvr.disconnect
set dmosvr=nothing
%>
<p><strong>数据 '<%=mdb%>' 数据恢复成功!</strong></p>
</BODY>
</HTML> 

