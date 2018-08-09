<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%

Dim KS
Set KS=New PublicCls
Dim ID
ID = KS.ChkClng(KS.S("ID"))
ChannelID=KS.ChkClng(KS.S("m"))
If ChannelID=0 Then Response.End()
Response.Write "document.write('" & Conn.Execute("Select Score From " & KS.C_S(ChannelID,2) &" Where ID=" & ID)(0) & "');"
Call CloseConn()
Set KS=Nothing
%> 
