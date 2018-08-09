<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Server.ScriptTimeout=9999999

Dim KSCls
Set KSCls = New DB_BackUp
KSCls.Kesion()
Set KSCls = Nothing

Class DB_BackUp
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Sub Kesion()
		
		 With KS
		If Trim(Request.ServerVariables("HTTP_REFERER"))="" Then
			  .echo "<script>alert('请不要，非法提交！');history.back();</script>"
			Response.end
		 End If
		 
		 
		  .echo "<html>"
		  .echo "<head>"
		  .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		  .echo "<title>备份数据库</title>"
		  .echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		  .echo "<script src=""../ks_inc/jquery.js""></script>"
		if KS.G("Action")<>"ExecSql" then
		  .echo ("<body oncontextmenu=""return false;"">")
		  .echo "<ul id='menu_top'>"
		  .echo "<li class='parent' onclick=""location.href='?Action=BackUp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>备份数据库</span></li>"
		  .echo "<li class='parent' onclick=""location.href='?Action=Restore';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/s.gif' border='0' align='absmiddle'>恢复数据库</span></li>"
		  .echo "<li class='parent' onclick=""location.href='?Action=Compact';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/verify.gif' border='0' align='absmiddle'>"
		If DataBaseType=1 Then
		   .echo "MSSQL数据库日志清理"
		Else
		   .echo "压缩修复数据库"
		End If
		  .echo "</span></li>"
		  .echo "</ul>"
	    elseif ks.g("flag")<>"Result" then
		  .echo ("<body oncontextmenu=""return false;"">")
		  .echo "<ul id='menu_top'>"
		  .echo "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		  .echo "  <tr>"
		  .echo "    <td height=""23"" align=""left"" valign=""top"">"
		  .echo "	<td align=""center""><strong>在线执行SQL语句</strong></td>"
		  .echo "    </td>"
		  .echo "  </tr>"
		  .echo "</table>"
		  .echo "</ul>"
		end if
		   Select Case KS.G("Action") 
		    Case "BackUp"
			   If Not KS.ReturnPowerResult(0, "KMST10007") Then                '检查管理员组操作(增和改)的权限检查
		          Call KS.ReturnErr(1, "")
				  Response.End
			   Else
			     Call Db_BackUp()
			  End If
		   Case "Restore"
			   If Not KS.ReturnPowerResult(0, "KMST10007") Then                '检查恢复数据库的权限
				  Call KS.ReturnErr(1, "")
				  Response.End
			   Else
			     Call Db_Restore()
			  End If
		   Case "Compact"
		     	If Not KS.ReturnPowerResult(0, "KMST10007") Then                '检查压缩数据库的权限
				  Call KS.ReturnErr(1, "")
				  Response.End
				Else
				  Call Db_Compact()
			  End If
		   Case "ExecSql"
		     If Not KS.ReturnPowerResult(0, "KMST10009") Then                '检查在线执行SQL语句
				  Call KS.ReturnErr(1, "")
			  Response.End
			  Else
			    Call Db_ExecSQL()
		      End If
		  End Select
		  End With
		End Sub

		
		'备份
	  Sub Db_BackUp()
		With KS
		  .echo "<table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		  .echo "  <tr> "
		  .echo "    <td align=""center"" valign=""top"">"
		  .echo "       <table width=""560"" border=""0"" cellpadding=""2"" cellspacing=""1"">"
		  .echo "        <tr> "
		  .echo "          <td height=""22""> "
		
		 Dim CurrPath,BackPath,tempArr,bkdbname    
		if request("Flag")="Backup" then
		   bkdbname=request.form("bkdbname")
		  If InStr(lcase(bkdbname),".asp")>0 or InStr(lcase(bkdbname),".asa")>0 or InStr(lcase(bkdbname),".php")>0 or InStr(lcase(bkdbname),".cer")>0 or InStr(lcase(bkdbname),".cdx")>0 or right("00000"&lcase(bkdbname),4)<>".bak" Then
			KS.echo "<script>alert('备份文件不正确，扩展名只能是.bak');$(parent.document).find('#ajaxmsg').toggle(false);history.back();</script>"
			Set KS = Nothing:Response.End
		   End If

           If DataBaseType=0 Then
			  CurrPath=request.form("Dbpath")
			  TempArr=replace(CurrPath,"/","\")
			  TempArr=split(TempArr,"\")
			  BackPath=Replace(CurrPath,TempArr(Ubound(TempArr)),"")
			  if KS.backupdata(CurrPath,BackPath & bkdbname)=true then
			     .echo "<div align=center><font color=green>系统主数据库备份成功!</font></div><div align=center>备份的主数据库为:" & backpath & Bkdbname & "</div>"
			  Else
			    .echo ("<font color=red>操作失败!</font>")
			  End IF
			  .echo "<script>$(parent.document).find('#ajaxmsg').toggle(false);</script>"
		  Else
			  If Left(bkdbname,1)<>"/" and Left(bkdbname,1)<>"\" Then bkdbname="/" & bkdbname
			  CurrPath=bkdbname
			  TempArr=replace(CurrPath,"/","\")
			  TempArr=split(TempArr,"\")
			  BackPath=Replace(CurrPath,TempArr(Ubound(TempArr)),"")
			  KS.CreateListFolder BackPath
			  conn.execute   "backup database  [" & DataBaseName &"]  to  disk='"& Server.MapPath(bkdbname) &"'" 
			  on   error   resume   next   
			  If   err   Then   
				   .echo ("<font color=red>操作失败!原因：" &err.description & "</font>")
			  Else   
				   .echo "<div align=center><font color=green>系统主数据库备份成功!</font></div><div align=center>备份的主数据库为:" & Bkdbname & "</div>"
			  End   If   
                   .echo "<script>$(parent.document).find('#ajaxmsg').toggle(false);</script>"
		  End If
		elseif request("Flag")="Backup1" then
		  CurrPath=request.form("Dbpath")
		  TempArr=replace(CurrPath,"/","\")
		  TempArr=split(TempArr,"\")
		  BackPath=Replace(CurrPath,TempArr(Ubound(TempArr)),"")
		  bkdbname=request.form("bkdbname")
		  If InStr(lcase(bkdbname),".asp")>0 or InStr(lcase(bkdbname),".asa")>0 or InStr(lcase(bkdbname),".php")>0 or InStr(lcase(bkdbname),".cer")>0 or InStr(lcase(bkdbname),".cdx")>0 or right("00000"&lcase(bkdbname),4)<>".bak" Then
			KS.echo "<script>alert('备份文件不正确，扩展名只能是.bak');history.back();</script>"
			Set KS = Nothing:Response.End
		   End If

		  if KS.backupdata(CurrPath,BackPath & bkdbname)=true then
		     .echo "<div align=center><font color=green>系统采集数据库备份成功!</font></div><div align=center>备份的采集数据库为:" & backpath & Bkdbname & "</div>"
		  Else
		     .echo ("<font color=red>操作失败!</font>")
		 End IF
		end if
		
		  .echo "</td>"
		  .echo "        </tr>"
		  .echo "        <tr> "
		  .echo "		     <td> "
		
		if DataBaseType=0 then
				  .echo "              <fieldset>"
				  .echo "          <form method=""post"" action=""?Action=BackUp&Flag=Backup"">"
				  .echo "	<legend>系统主数据库</legend>"
				  .echo "	<table width=""91%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center""> 当前数据库路径</td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center""><input type=text size=50 name=DBpath value=""" &server.mappath(DBPath) & """ readonly></td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22""></td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center"">备份数据库名称[如果备份目录存在该文件将覆盖，否则将自动创建]</td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center""><input type=text size=50 name=bkDBname value=""Data(" & year(now)&month(now)&day(now) & ").bak""></td>"
				  .echo "                </tr>"
				  .echo "              </table>"
				  .echo "			  </fieldset>"
				  .echo "			  <table width=""100%"" border=""0"">"
				  .echo "			   <tr>"
				  .echo "			   <td height=""50"" align=center>"
				  .echo "			     <input type=submit onclick=""$(parent.document).find('#ajaxmsg').toggle(true);"" value=""确定备份"" class=""button"">"
				  .echo "			   </td>"
				  .echo "			   </tr>"
				  .echo "			   </form></table>"
		Else
				  .echo "              <fieldset>"
				  .echo "          <form method=""post"" action=""?Action=BackUp&Flag=Backup"">"
				  .echo "	<legend>系统主数据库(MSSQL)</legend>"
				  .echo "	<table width=""91%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center""> 当前数据库</td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center""><input type=text size=50 name=DBpath value=""" & DataBaseName & """ readonly></td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22""></td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center"">备份数据库名称[如果备份目录存在该文件将覆盖，否则将自动创建]</td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center""><input type=text size=50 name=bkDBname value=""/KS_Data/SQL(" & year(now)&month(now)&day(now) & ").bak""></td>"
				  .echo "                </tr>"
				  .echo "              </table>"
				  .echo "			  </fieldset>"
				  .echo "			  <table width=""100%"" border=""0"">"
				  .echo "			   <tr>"
				  .echo "			   <td height=""50"" align=center style='color:green'>"
				  .echo "			     <br/>tips:一般的虚拟主机是不允许mssql数据库使用该项功能备份的，如备份不成功，请联系主机商备份！<br/><br/> <input type=submit onclick=""$(parent.document).find('#ajaxmsg').toggle(true);"" value=""确定备份"" class=""button"">"
				  .echo "			  </td>"
				  .echo "			   </tr>"
				  .echo "			   </form></table>"
		end if
		
		  .echo "              <fieldset>"
		
		  .echo "          <form method=""post"" action=""?Action=BackUp&Flag=Backup1"">"
		  .echo "	<legend>系统采集数据库</legend>"
		  .echo "	<table width=""91%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""center""> 当前数据库路径</td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""center""><input type=text size=50 name=DBpath value=""" &server.mappath(CollectDBPath) & """ readonly></td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22""></td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""center"">备份数据库名称[如果备份目录存在该文件将覆盖，否则将自动创建]</td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""center""><input type=text size=50 name=bkDBname value=""Collect(" & year(now)&month(now)&day(now) & ").bak""></td>"
		  .echo "                </tr>"
		  .echo "              </table>"
		  .echo "			  </fieldset>"
		  .echo "			  <table width=""100%"" border=""0"">"
		  .echo "			   <tr>"
		  .echo "			   <td height=""50"" align=center>"
		  .echo "			     <input type=submit value=""确定备份"" class=""button"">"
		  .echo "			   </td>"
		  .echo "			   </tr>"
		  .echo "          </form>"
		  .echo "			   </table>"
		
		if DataBaseType=0 then			  
		  .echo "			  主数据库完整路径为：&nbsp;&nbsp;&nbsp;<font color=red>" & server.mappath(dbpath) & "</font><br>"
		end if
		  .echo "              采集数据库完整路径为：<font color=red>" & server.mappath(CollectDBPath) & "</font><br></td>"
		  .echo "        </tr>"
		  .echo "      </table>"
		  .echo "     </td>"
		  .echo "  </tr>"
		  .echo "</table>"
		  .echo "</body>"
		  .echo "</html>"
		 End With
		End Sub
		
		'恢复
		Sub Db_Restore()
		  With KS
			  .echo "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
			  .echo "  <tr> "
			  .echo "    <td align=""center"" valign=""top""> <br> <strong><br>"
			  .echo "      </strong> <table width=""560"" border=""0"" cellpadding=""2"" cellspacing=""1"">"
			  .echo "        <tr> "
			  .echo "          <td height=""25"" align=""center""> "
					
			if request("submit1")="恢复选中的备份文件" then
				if Request.Form("backname")="0" then
				    .echo ("<script>alert('没有备份文件!');history.back();</script>")
				  Response.End
				end if
			   if  RestoreDatabase(Request.Form("backname"),Request("Flag"))=true then
				  if request("Flag")="main" then
				   .echo "<div align=center><font color=green>操作成功！</font></div><div align=center>主数据库已从<font color=red>" & Request.Form("backname") & "</font>备份中恢复!</div>"
				  else
				   .echo "<div align=center><font color=green>操作成功！</font></div><div align=center>采集数据库已从<font color=red>" & Request.Form("backname") & "</font>备份中恢复!</div>"
				  end if
			   else
				   .echo "<font color=red>操作失败!</font>"
			   end if
			elseif request("submit1")="删除选中的备份文件" then
			   if Request.Form("backname")="0" then
				    .echo ("<script>alert('没有备份文件!');history.back();</script>")
				  Response.End
				end if
			  if  DeleteFile(Request.Form("backname"))=true then
				   .echo "<div align=center><font color=green>操作成功！</font></div><div align=center>备份文件<font color=red>" & Request.Form("backname") & "</font>已删除!</div>"
			   else
				   .echo "<font color=red>操作失败!</font>"
			   end if
			end if  
			
			
			  .echo "</td>"
			  .echo "        </tr>"
			  .echo "        <tr> "
					 
			  .echo "           <td>"
			
			if DataBaseType=0 then
			  .echo "          <form method=""post"" name=""restoreform"" action=""KS.Database.asp?Action=Restore&Flag=main"">"
			  .echo "<br> <fieldset>"
			  .echo "	<legend>主数据库已备份的文件</legend>"
						
						
			  .echo "              <table width=""96%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
			  .echo "                <tr> "
			  .echo "                  <td height=""22"" align=""center""><strong>选择备份文件：</strong>"
			
								dim  tempStr,strCurDir,CurrDataBase,CurrLdb,Fso,Dir,s
								dim havebackfile:havebackfile=false
								tempStr=replace(dbpath,"/","\")
								tempStr=split(tempStr,"\")
								strCurDir=replace(dbpath,tempStr(ubound(tempStr)),"")
								strCurDir=server.mappath(strCurDir)
								
								 CurrDataBase=tempStr(ubound(tempStr))
								 CurrLdb=left(CurrDataBase,len(CurrDataBase)-4) & ".ldb"
								
			  .echo "				    <select name=""backname"">"
								 
							  set fso = KS.InitialObject(KS.Setting(99))
							  set dir = fso.GetFolder(strCurDir)
							  for each s in dir.Files
								 if s.name<>CurrDataBase and s.name<>Currldb and lcase(right(s.name,4))<>".dat" then
								  havebackfile=true
								  .echo "<option value=""" & strCurDir &"\" & s.name & """>" & s.name & "</option>"
								end if
							  next
							  if havebackfile=false then
							     .echo "<option value=""0"">---还没有备份的主数据库文件---</option>"
							   end if
			  .echo "		            </select></td>"
			  .echo "               </tr>"
			  .echo "              </table>"
	
			  .echo "			  </fieldset>"
			  .echo "			  <table width=""100%"" border=""0"">"
			  .echo "			   <tr><td height=""50"" align=center>"
			  .echo "			     <input type=""submit"" name=""submit1"" "
			if havebackfile=false then   .echo "disabled"
			  .echo " value=""恢复选中的备份文件"" class=""button"" onclick=""return(confirm('确定恢复数据库吗？此操作不可逆'))"">"
			  .echo "			     <input name=""submit1"" type=""submit"" "
			if havebackfile=false then   .echo " disabled" 
			  .echo " value=""删除选中的备份文件"" class=""button"" onclick=""return(confirm('确定删除选中的备份文件吗？此操作不可逆'))""/>"
			  .echo "			   </td>"
			  .echo "			   </tr>"
			  .echo "		      </table>"
			  .echo "          </form>"
			  .echo "		    </td>"
			  .echo "        </tr>"
			  .echo "      </table>"
			  .echo "     </td>"
			  .echo "  </tr>"
			  .echo "</table>"
			
			end if
			  .echo "             <br><table width=""560"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
			  .echo "                <tr> "
			
			  .echo "          <form method=""post"" name=""restoreform"" action=""KS.Database.asp?Action=Restore&Flag=collect"">"
					 
			  .echo "          <td><fieldset>"
			  .echo "	<legend>采集数据库已备份的文件</legend>"
						
						
			  .echo "              <table width=""96%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
			  .echo "                <tr> "
			  .echo "                  <td height=""22"" align=""center""><strong>选择备份文件：</strong>"
			
								
								havebackfile=false
								tempStr=replace(CollectDBPath,"/","\")
								tempStr=split(tempStr,"\")
								strCurDir=replace(CollectDBPath,tempStr(ubound(tempStr)),"")
								strCurDir=server.mappath(strCurDir)
								
								 CurrDataBase=tempStr(ubound(tempStr))
								 CurrLdb=left(CurrDataBase,len(CurrDataBase)-4) & ".ldb"
								
			  .echo "				    <select name=""backname"">"
								 
							  set fso = KS.InitialObject(KS.Setting(99))
							  set dir = fso.GetFolder(strCurDir)
							  for each s in dir.Files
								 if s.name<>CurrDataBase and s.name<>Currldb then
								  havebackfile=true
								  .echo "<option value=""" & strCurDir &"\" & s.name & """>" & s.name & "</option>"
								end if
							  next
							  if havebackfile=false then
							     .echo "<option value=""0"">---还没有备份的采集数据库文件---</option>"
							   end if
			  .echo "		            </select></td>"
			  .echo "               </tr>"
			  .echo "              </table>"
			  .echo "			  </fieldset>"
			
			  .echo "			  <table width=""100%"" border=""0"">"
			  .echo "			   <tr><td height=""50"" align=center>"
			  .echo "			     <input type=""submit"" name=""submit1"" "
			if havebackfile=false then   .echo "disabled"
			  .echo " value=""恢复选中的备份文件"" class=""button"" onclick=""return(confirm('确定恢复数据库吗？此操作不可逆'))"">"
			  .echo "			     <input name=""submit1"" type=""submit"" "
			if havebackfile=false then   .echo " disabled" 
			  .echo " value=""删除选中的备份文件"" class=""button"" onclick=""return(confirm('确定删除选中的备份文件吗？此操作不可逆'))""/>"
			  .echo "			   </td>"
			  .echo "			   </tr>"
			  .echo "		      </table>"
			
			  .echo "          </form>"
			  .echo "			   </td>"
			  .echo "			   </tr>"
			  .echo "		      </table>"
			  .echo "</body>"
			  .echo "</html>"
			End With
		End Sub
		' 恢复数据库
		Public Function RestoreDatabase(BackName,Flag)
				dim fso,sFileName
				RestoreDatabase=false
				on error resume next
				set fso = KS.InitialObject(KS.Setting(99))
				IF Flag="main" Then  '主数据库
				  sFileName = DbPath
				  Conn.Close
				  fso.CopyFile BackName, server.mappath(DbPath), True
				  if err then
					RestoreDatabase=false
				  else
					RestoreDatabase=true
				  end if
				  conn.Open ConnStr
				Elseif Flag="collect" Then  '采集数据库
				 
				 sFileName = CollectDBPath
				  fso.CopyFile BackName, server.mappath(CollectDBPath), True
				  if err then
					RestoreDatabase=false
				  else
					RestoreDatabase=true
				  end if
				Else 
				  RestoreDatabase=false
				  Exit Function
				End IF
			   IF err Then
				RestoreDatabase=false
			   End IF
			End Function
		
		'**************************************************
		'函数名：DeleteFile
		'作  用：删除指定文件
		'参  数：FileStr要删除的文件
		'返回值：成功返回true 否则返回Flase
		'**************************************************
		Function DeleteFile(FileStr)
		   Dim fso
		   On Error Resume Next
		   Set fso = KS.InitialObject(KS.Setting(99))
			If fso.FileExists(FileStr) Then
				fso.DeleteFile FileStr, True
			Else
			DeleteFile = True
			End If
		   Set fso = Nothing
		   If Err.Number <> 0 Then
		   Err.Clear
		   DeleteFile = False
		   Else
		   DeleteFile = True
		   End If
		End Function


       '恢复
	   Sub Db_Compact()
	    With KS
		  .echo "<table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		  .echo "  <tr> "
		  .echo "    <td align=""center"" valign=""top""> <br> <strong><br>"
		  .echo "      </strong> <table width=""560"" border=""0"" cellpadding=""2"" cellspacing=""1"">"
		  .echo "        <tr> "
		  .echo "          <td height=""25"" align=""center""> "
				  
		if request("Flag")="Backup" then
		  If DataBaseType=0 Then
		   if  CompactDatabase(DBPath,ConnStr)=true then
			   .echo "<font color=green>主数据库压缩和修复成功!</font>"
		   else
			   .echo "<font color=red>操作失败!</font>"
		   end if
		  Else
			 conn.execute("DUMP TRANSACTION [" & DataBaseName & "] WITH  NO_LOG")
			 conn.execute("DBCC SHRINKDATABASE([" & DataBaseName & "])")
			  on   error   resume   next   
			  If   err   Then   
				   .echo ("<font color=red>操作失败!</font>")
			  Else   
				   .echo "<div align=center><font color=green>您的mssql数据库日志已清空!</font></div>"
			  End   If   
                 .echo "<script>$(parent.document).find('#ajaxmsg').toggle(false);</script>"
		  End If
		elseif Request("Flag")="Backup1" then
		   if  CompactCollectDatabase(CollectDBPath,CollcetConnStr)=true then
			   .echo "<font color=green>采集数据库压缩和修复成功!</font>"
		   else
			   .echo "<font color=red>操作失败!</font>"
		   end if
		end if
		
		  .echo "</td>"
		  .echo "        </tr>"
		
		if DataBaseType=0 then
		  .echo "        <tr> "
		  .echo "          <form method=""post"" action=""?Action=Compact&Flag=Backup"">"
				 
		  .echo "            <td> <fieldset>"
		  .echo "	<legend>主数据库信息</legend>"
					
					dim filesize:filesize=KS.GetFieSize(server.mappath(DBPath))
					dim ReclaimedSpace:ReclaimedSpace=CLng(conn.Properties("Jet OLEDB:Compact Reclaimed Space Amount").Value)
					dim LocaleIdentifier:LocaleIdentifier=Conn.Properties("Locale Identifier").Value
		  .echo "              <table width=""96%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
		  .echo "                <tr> "
		  .echo "                  <td width=""23%"" height=""22"" align=""right""><strong>数据库路径：</strong></td>"
		  .echo "                  <td width=""77%""><font color=#ff6600>" & server.mappath(DBPath) & "</font></td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""right""><strong>压缩前大小：</strong></td>"
		  .echo "                  <td height=""22"">" & FormatNumber(filesize, 0, False, False, True) & " 字节</td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""right""><strong>压缩后大小：</strong></td>"
		  .echo "                  <td height=""22"">" & FormatNumber(filesize - ReclaimedSpace, 0, False, False, True) & " 字节 (总计可以减少" & FormatNumber(ReclaimedSpace, 0, True, False, True)& " 字节)</td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""right""><strong>地区标识符：</strong></td>"
		  .echo "                  <td height=""22"">" & GetLocaleName(LocaleIdentifier) & "</td>"
		  .echo "                </tr>"
					   
		  .echo "              </table>"
		  .echo "			  </fieldset>"
		  .echo "			  <table width=""100%"" border=""0"">"
		  .echo "			   <tr><td height=""50"" align=center>"
		  .echo "			     <input type=submit value=""开始压缩"" class=""button"">"
		  .echo "			   </td>"
		  .echo "			   </tr>"
		  .echo "			   </table>"
		  .echo "			  </td>"
		  .echo "          </form>"
		  .echo "        </tr>"
		Else
		  .echo "        <tr> "
		  .echo "          <form method=""post"" action=""?Action=Compact&Flag=Backup"">"
				 
		  .echo "            <td> <fieldset>"
		  .echo "	<legend>主数据库信息(MSSQL)</legend>"
					
		
		  .echo "              <table width=""96%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
		on error resume next
		Dim RS,I
		Set RS=Conn.Execute("select name, convert(float,size) * (8192.0/1024.0)/1024. from dbo.sysfiles")			
		For I=0 To 1
			  .echo "                <tr> "
			  .echo "                  <td height=""22""><strong>文件" & RS(0) & "大小：</strong>" & rs(1) & " MB</td>"
			  .echo "                </tr>"
		 RS.MoveNext
	    Next
	   RS.Close:Set RS=Nothing
		if err then KS.AlertHintScript "对不起,您的服务器不支持此操作!"
					   
		  .echo "              </table>"
		  .echo "			  </fieldset>"
		  .echo "			  <table width=""100%"" border=""0"">"
		  .echo "			   <tr><td height=""50"" align=center>"
		  .echo "			     <input type=submit value=""开始清理日志"" onclick=""$(parent.document).find('#ajaxmsg').toggle(true);"" class=""button"">"
		  .echo "			   </td>"
		  .echo "			   </tr>"
		  .echo "			   </table>"
		  .echo "			  </td>"
		  .echo "          </form>"
		  .echo "        </tr>"		
		end if
		  .echo "        <tr> "
		  .echo "          <form method=""post"" action=""?Action=Compact&Flag=Backup1"">"
				 
		  .echo "            <td><br> <fieldset>"
		  .echo "	<legend>采集数据库信息</legend>"
		
						 conn.close
						Set conn = KS.InitialObject("ADODB.Connection")
						conn.open CollcetConnStr
		
					filesize=KS.GetFieSize(server.mappath(CollectDBPath))
					ReclaimedSpace=CLng(conn.Properties("Jet OLEDB:Compact Reclaimed Space Amount").Value)
					LocaleIdentifier=Conn.Properties("Locale Identifier").Value
		  .echo "              <table width=""96%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
		  .echo "                <tr> "
		  .echo "                  <td width=""23%"" height=""22"" align=""right""><strong>数据库路径：</strong></td>"
		  .echo "                  <td width=""77%""><font color=#ff6600>" & server.mappath(CollectDBPath) & "</font></td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""right""><strong>压缩前大小：</strong></td>"
		  .echo "                  <td height=""22"">" & FormatNumber(filesize, 0, False, False, True) & " 字节</td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""right""><strong>压缩后大小：</strong></td>"
		  .echo "                  <td height=""22"">" & FormatNumber(filesize - ReclaimedSpace, 0, False, False, True) & " 字节 (总计可以减少" & FormatNumber(ReclaimedSpace, 0, True, False, True)& " 字节)</td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""right""><strong>地区标识符：</strong></td>"
		  .echo "                  <td height=""22"">" & GetLocaleName(LocaleIdentifier) & "</td>"
		  .echo "                </tr>"
					   
		  .echo "              </table>"
		  .echo "			  </fieldset>"
		  .echo "			  <table width=""100%"" border=""0"">"
		  .echo "			   <tr><td height=""50"" align=center>"
		  .echo "			     <input type=submit value=""开始压缩"" class=""button"">"
		  .echo "			   </td>"
		  .echo "			   </tr>"
		  .echo "			   </table>"
		  .echo "			  </td>"
		  .echo "          </form>"
		  .echo "        </tr>"
		
		
		  .echo "      </table>"
		  .echo "	  说明：避免不可预测的错误发生，请在压缩之前备份原始数据库！"
		  .echo "     </td>"
		  .echo "  </tr>"
		  .echo "</table>"
		  .echo "</body>"
		  .echo "</html>"
		End With
		End Sub
		
		'**********************************************************************
		'函数名：CompactDatabase
		'作用：压缩主数据库
		'参数：DBPath--数据库位置,ConnStr---数据库连接字符串
		'**********************************************************************   
		 Public Function CompactDatabase(DBPath, ConnStr)
				On Error Resume Next
				Dim strTempFile, fso, jro, ver, strCon, strTo, LCID
				Set fso = KS.InitialObject(KS.Setting(99))
				strTempFile = DBPath
				strTempFile = Left(strTempFile, InStrRev(strTempFile, "\")) & fso.GetTempName
				Set jro = KS.InitialObject("JRO.JetEngine")
				LCID = Conn.Properties("Locale Identifier").Value
				'关闭数据库
				Conn.Close
				strTo = "Provider=Microsoft.Jet.OLEDB.4.0; Locale Identifier=" & LCID & "; Data Source=" & Server.MapPath(strTempFile) & "; Jet OLEDB:Engine Type=" & ver
				
				jro.CompactDatabase ConnStr, strTo
				CompactDatabase = False
				If Err Then
					fso.DeleteFile Server.MapPath(strTempFile)
				Else
					fso.DeleteFile Server.MapPath(DBPath)
					fso.MoveFile Server.MapPath(strTempFile), Server.MapPath(DBPath)
					If Err Then
						fso.DeleteFile Server.MapPath(strTempFile)
					Else
						CompactDatabase = True
					End If
				End If
				Set jro = Nothing
				Set fso = Nothing
				'重新打开数据库
				Conn.Open ConnStr
		End Function
		'**********************************************************************
		'函数名：CompactDatabase
		'作用：压缩采集数据库
		'参数：DBPath--数据库位置,ConnStr---数据库连接字符串
		'**********************************************************************   
		 Public Function CompactCollectDatabase(DBPath, ConnStr)
				On Error Resume Next
				Dim strTempFile, fso, jro, ver, strCon, strTo, LCID
				
				Set conn = KS.InitialObject("ADODB.Connection")
				conn.open CollcetConnStr
				
				Set fso = KS.InitialObject(KS.Setting(99))
				strTempFile = DBPath
				strTempFile = Left(strTempFile, InStrRev(strTempFile, "\")) & fso.GetTempName
				Set jro = KS.InitialObject("JRO.JetEngine")
				LCID = Conn.Properties("Locale Identifier").Value
				'关闭数据库
				Conn.Close
				strTo = "Provider=Microsoft.Jet.OLEDB.4.0; Locale Identifier=" & LCID & "; Data Source=" & Server.MapPath(strTempFile) & "; Jet OLEDB:Engine Type=" & ver
				
				jro.CompactDatabase ConnStr, strTo
				CompactCollectDatabase = False
				If Err Then
					fso.DeleteFile Server.MapPath(strTempFile)
				Else
					fso.DeleteFile Server.MapPath(DBPath)
					fso.MoveFile Server.MapPath(strTempFile), Server.MapPath(DBPath)
					If Err Then
						fso.DeleteFile Server.MapPath(strTempFile)
					Else
						CompactCollectDatabase = True
					End If
				End If
				Set jro = Nothing
				Set fso = Nothing
				'重新打开数据库
				Conn.Open ConnStr
		End Function
		
		'得到数据库的地区标识符	
		Function GetLocaleName(lcid)
				Select Case lcid
					Case 1033	GetLocaleName = "常规"
					Case 2052	GetLocaleName = "中文标点"
					Case 133124	GetLocaleName = "中文笔画"
					Case 1028	GetLocaleName = "中文笔画(台湾)"
					Case 197636	GetLocaleName = "中文拼音(台湾)"
					Case 1050	GetLocaleName = "克罗地亚语"
					Case 1029	GetLocaleName = "捷克语"
					Case 1061	GetLocaleName = "爱沙尼亚语"
					Case 1036	GetLocaleName = "法语"
					Case 66615	GetLocaleName = "格鲁吉亚语(现代)"
					Case 66567	GetLocaleName = "德语(电话簿)"
					Case 1038	GetLocaleName = "匈牙利语"
					Case 66574	GetLocaleName = "匈牙利语(技术术语)"
					Case 1039	GetLocaleName = "冰岛语"
					Case 1041	GetLocaleName = "日语"
					Case 66577	GetLocaleName = "日语(Unicode)"
					Case 1042	GetLocaleName = "韩语"
					Case 66578	GetLocaleName = "韩语(Unicode)"
					Case 1062	GetLocaleName = "拉脱维亚语"
					Case 1036	GetLocaleName = "立陶宛语"
					Case 1071	GetLocaleName = "FYRO 马其顿语"
					Case 1044	GetLocaleName = "挪威语/丹麦语"
					Case 1045	GetLocaleName = "波兰语"
					Case 1048	GetLocaleName = "罗马尼亚语"
					Case 1051	GetLocaleName = "斯洛伐克语"
					Case 1060	GetLocaleName = "斯洛文尼亚语"
					Case 1034	GetLocaleName = "西班牙语(传统)"
					Case 3082	GetLocaleName = "西班牙语(西班牙)"
					Case 1053	GetLocaleName = "瑞典语/芬兰语"
					Case 1054	GetLocaleName = "泰国语"
					Case 1055	GetLocaleName = "土耳其语"
					Case 1058	GetLocaleName = "乌克兰语"
					Case 1066	GetLocaleName = "越南语"
					Case Else	GetLocaleName = "未知"
				End Select
			End Function

       '在线执行SQL
	   Sub Db_ExecSQL()
	   With KS
		  .echo "<html>"
		  .echo "<head>"
		  .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		  .echo "<title>在线执行SQL语句</title>"
		  .echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		  .echo "<script src=""../ks_inc/jquery.js""></script>"
		Dim Flag:Flag=KS.G("Flag")
		IF Flag="Result" Then 
		  .echo ("<body style=""margin:1;"">")
		 Call ExeSQL
		Else
		  .echo ("<body scroll=no>")
    %>
		<script language="javascript">
	<!--
	 function CheckForm()
	 {
	 if ($('textarea[name=Sql]').val()=='')
	  {
	  alert('请输入SQL查询语句！');
	  $('textarea[name=Sql]').focus();
	  return false;
	  }
	  ExeSQLFrame.location.href="KS.Database.asp?Action=ExecSql&Flag=Result&SQL="+escape($('textarea[name=Sql]').val().replace('+','ksaddks'));
	  return false;
	  }
	-->
	</script>
	<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	<form name="SqlForm" method="post" Action="?Action=ExecSql" onsubmit="return CheckForm()">
	<tr height="50">
	  <td>
	  <textarea name="Sql" rows="5" wrap="OFF" style="width:100%;"></textarea>
	  <input type="hidden" name="Flag" value="Exec">
	  </td>
	</tr>
	<tr height="25">
	 <td align="center">
	  <input type="submit" name="submit1" class="button" value="立即执行"><span style="color:red">一次可以执行多条SQL语句，多条语句请用回车换行隔开，如果您没有一定的SQL基础，建议不要使用！</span>
	  </td>
	</tr>
	</form>
	  <tr> 
		<td valign="_top"><iframe id="ExeSQLFrame" scrolling="auto" src="KS.Database.asp?Action=ExecSql&Flag=Result" style="width:100%;height:93%" frameborder=1></iframe></td>
	  </tr>
	</table>
	<% End iF%>
	</BODY>
	</HTML>
<% End With
  End Sub
  Sub ExeSQL()
        Dim SelectSQLTF,ExecSQLErrorTF,ExeResultNum,ExeResult,FiledObj,i
		Dim Sql:Sql =replace(request.querystring("Sql"),"ksaddks","+")
	    if SQL="" Then Exit Sub
		sql=split(sql,vbcrlf)
		For I=0 To Ubound(sql)
		  if (Sql(i)<>"") Then
				If Instr(1,lcase(Sql(i)),"delete from ks_log")<>0 then
					Call KS.AlertHistory("对不起，不能删除日志表数据！",-1)
						Exit Sub
				End If
				SelectSQLTF = (LCase(Left(Trim(Sql(i)),6)) = "select")
				Conn.Errors.Clear
				On Error Resume Next
				if SelectSQLTF = True then
					  Set ExeResult = Conn.Execute(Sql(i),ExeResultNum)
				else
					  Conn.Execute Sql(i),ExeResultNum
				end if
				 
				If Conn.Errors.Count<>0 Then
					  ExecSQLErrorTF = True
					  Set ExeResult = Conn.Errors
				Else
					  ExecSQLErrorTF = False
				End If
				if ExecSQLErrorTF = True then
				%>
				<table width="100%" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
				  <tr bgcolor="F4F4EA"> 
					<td height="20" nowrap> 
					  <div align="center">错误号</div></td>
					<td height="20" nowrap> 
					  <div align="center">来源</div></td>
					<td height="20" nowrap> 
					  <div align="center">描述</div></td>
					<td height="20" nowrap> 
					  <div align="center">帮助</div></td>
					<td height="20" nowrap> 
					  <div align="center">帮助文档</div></td>
				  </tr>
				  <tr height="20" bgcolor="#FFFFFF"> 
					<td nowrap> 
					  <% = Err.Number %> </td>
					<td nowrap> 
					  <% = Err.Description %> </td>
					<td nowrap> 
					  <% = Err.Source %> </td>
					<td nowrap> 
					  <% = Err.Helpcontext %> </td>
					<td nowrap> 
					  <% = Err.HelpFile %> </td>
				  </tr>
				</table>
				<%
				else
				%>
				<table border="0" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
				  <%
					if SelectSQLTF = True then
				%>
				  <tr>
				<%
						For Each FiledObj In ExeResult.Fields
				%>
					<td nowrap bgcolor="F4F4EA" height="26"><div align="center">
						<% = FiledObj.name %>
					  </div></td>
				<%
						next
				%>
				  </tr>
				<%
						do while Not ExeResult.Eof
				%>
				  <tr height="20" nowrap bgcolor="#ffffff" onMouseOver="this.style.background='#F5f5f5'" onMouseOut="this.style.background='#FFFFFF'">
				<%
							For Each FiledObj In ExeResult.Fields
				%>
					<td> 
					  <div align="center">
						<%
						 if IsNull(FiledObj.value) then
							KS.echo("&nbsp;")
						 else
							KS.Echo (FiledObj.value)
						 end if
						 %>
					  </div></td>
				<%
							next
				%>
				  </tr>
				<%
							ExeResult.MoveNext
						loop
					else
				%>
				  <tr>
					<td bgcolor="F4F4EA" height="26">
				<div align="center">执行结果</div></td>
				  </tr>
				  <tr>
					<td height="20" bgcolor="#FFFFFF">
				<div align="center">
						<% = ExeResultNum & "条纪录被影响"%>
					  </div></td>
				  </tr>
				<%
					end if
				%>
				</table>
				<%
				  end if
			 end if
		  Next
		 End Sub
End Class
%> 
