﻿<%@ LANGUAGE=VBScript CodePage=65001%>
<% response.charset="utf-8" %>
<% session.codepage=65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> <strong><br>
      </strong> <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td class="back_southidc" height="25"> <div align="center"><strong>备份数据库</strong></div>
            <%
if request("action")="Backup" then
call backupdata()
else
%></td>
        </tr>
        <tr class="tr_southidc"> 
          <form method="post" action="Manage_backup.asp?action=Backup">
            <td><br> 
              <table width="91%" border="0" align="center" cellpadding="0" cellspacing="2">
                <tr> 
                  <td height="25"><strong>备份数据库文件</strong>[需要FSO权限]</td>
                </tr>
                <tr> 
                  <td height="22"> 当前数据库路径</td>
                </tr>
                <tr> 
                  <td height="22"><input type=text size=50 name=DBpath value="<%=db%>"></td>
                </tr>
                <tr> 
                  <td height="22"><input type="hidden" size=50 name=bkfolder value=Databackup ></td>
                </tr>
                <tr> 
                  <td height="22">备份数据库名称[如备份目录有该文件，将覆盖，如没有，将自动创建]</td>
                </tr>
                <tr> 
                  <td height="22"><input type=text size=30 name=bkDBname value="<%=date()%>"></td>
                </tr>
                <tr> 
                  <td height="22"><div align="center"> 
                      <input type=submit value="确定">
                    </div></td>
                </tr>
                <tr> 
                  <td height="22"><br> <br>
                    本程序的默认数据库文件为<%=db%><br>
                    您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
                    注意：所有路径都是相对与程序空间根目录的相对路径</td>
                </tr>
                <tr> 
                  <td height="22">&nbsp;</td>
                </tr>
              </table></td>
          </form>
        </tr>
      </table>
      <%end if%>
      <% 
sub backupdata() 
Dbpath=request.form("Dbpath") 
Dbpath=server.mappath(Dbpath) 
bkfolder=request.form("bkfolder") 
bkdbname=request.form("bkdbname") 
Set Fso=server.createobject("scripting.filesystemobject") 
if fso.fileexists(dbpath) then 
If CheckDir(bkfolder) = True Then 
fso.copyfile dbpath,bkfolder& "\"& bkdbname 
else 
MakeNewsDir bkfolder 
fso.copyfile dbpath,bkfolder& "\"& bkdbname & ".asa"
end if 
response.write "<center>备份数据库成功，备份的数据库为 " & bkfolder & "\" & bkdbname & ".asa</center>" 
Else 
response.write "找不到您所需要备份的文件。" 
End if 
end sub 
'------------------检查某一目录是否存在------------------- 
Function CheckDir(FolderPath) 
folderpath=Server.MapPath(".")&"\"&folderpath 
Set fso1 = CreateObject("Scripting.FileSystemObject") 
If fso1.FolderExists(FolderPath) then 
'存在 
CheckDir = True 
Else 
'不存在 
CheckDir = False 
End if 
Set fso1 = nothing 
End Function 
'-------------根据指定名称生成目录--------- 
Function MakeNewsDir(foldername) 
Set fso1 = CreateObject("Scripting.FileSystemObject") 
Set f = fso1.CreateFolder(foldername) 
MakeNewsDir = True 
Set fso1 = nothing 
End Function 
%></td>
  </tr>
</table>
