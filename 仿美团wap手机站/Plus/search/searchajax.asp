<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.EscapeCls.asp"-->
<%


response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="utf-8"

dim ks:set ks=new publiccls
dim orderstr,Param,ModelTable
dim ChannelID:ChannelID=KS.ChkClng(request("m"))
If ChannelID=0 Then
	ModelTable="KS_ItemInfo"
ElseIf ChannelID=102 Then
	ModelTable="KS_AskTopic"
Else
    ModelTable=KS.C_S(ChannelID,2)
End If

If ChannelID=102 Then
	Param=" Where LockTopic=0"
	orderstr=" order by topicid desc"
Else
	Param=" Where Verific=1 and deltf=0"
	orderstr=" order by id desc"
End If

		   
dim searchText:searchText=KS.DelSQL(Unescape(Request("searchText")))
if ks.isnul(searchText) then ks.die ""
dim rs:set rs=server.CreateObject("adodb.recordset")
rs.open "select top 100 title from " & ModelTable &" " & Param & " and title like '%" & searchText & "%'" & orderstr,conn,1,1
do while not rs.eof
 response.write rs(0) & "@@"
rs.movenext
loop
rs.close
set rs=nothing
closeconn
set ks=nothing
%>