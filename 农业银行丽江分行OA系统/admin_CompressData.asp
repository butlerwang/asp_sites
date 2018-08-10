
<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="data.asp"-->
<!-- #include file="check.asp" -->
<%
if Session("Urule")="c" then
response.redirect "error.asp?id=admin"
end if
%>
<title>数据压缩</title>
<link rel="stylesheet" type="text/css" href="forum.css">

<BODY bgcolor="#ffffff" alink="#333333" vlink="#333333" link="#333333" topmargin="20">
<form action="Admin_CompressData.asp">
<tr align=center>
<td><br>输入数据库所在相对路径,并且输入数据库名称 </td>
</tr>
<tr align=center>
<td><input type="text" name="dbpath" value=db\system1.asa></td>
</tr>
<tr align=center><br>
<td><input type="checkbox" name="boolIs97" value="True">如果您使用 Access 97 数据库请选择<br>
(系统默认为 Access 2000 数据库)<br>
<input type="submit" value="开始压缩"><br></td>
</tr>
</table>
<%
Dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")

If dbpath <> "" Then
dbpath = server.mappath(dbpath)
	response.write(CompactDB(dbpath,boolIs97))
End If
%>
</p></td>
            </tr>
        </table>
        </td>
    </tr>
</table>
<%
	end sub
%>
<%
Const JET_3X = 4

Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath
strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(dbPath) Then
Set Engine = CreateObject("JRO.JetEngine")

	If boolIs97 = "True" Then
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb;" _
		& "Jet OLEDB:Engine Type=" & JET_3X
	Else
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
	End If

fso.CopyFile strDBPath & "temp.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
Set fso = nothing
Set Engine = nothing

	CompactDB = "你的数据库, " & dbpath & ", 已经压缩成功!" & vbCrLf

Else
	CompactDB = "数据库名称或路径不正确. 请重试!" & vbCrLf
End If

End Function
%>
</form>