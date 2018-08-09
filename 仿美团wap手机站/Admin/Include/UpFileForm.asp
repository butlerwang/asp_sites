<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%

Dim KSCls
Set KSCls = New UpFileForm
KSCls.Kesion()
Set KSCls = Nothing

Class UpFileForm
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim ChannelID, UpLoadFrom,UploadDir
		ChannelID = KS.G("ChannelID")
		If ChannelID = "" Then ChannelID = 0
		UpLoadFrom = ChannelID
		UploadDir=Request("Path")
		If UploadDir="" Then UploadDir=KS.GetUpFilesDir
		If Right(UploadDir,1)<>"/" Then UploadDir=UploadDir & "/"
		
		Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<title>上传文件</title>"
		Response.Write "<link rel=""stylesheet"" href=""" & KS.GetDomain & "Editor/ksplus/Editor.css"">"
		Response.Write "<link href=""admin_style.css"" rel=""stylesheet"" type=""text/css"">"
		Response.Write "</head>"
		Response.Write "<body onselectstart=""return false;"" topmargin=""0"" leftmargin=""0"">"
		Response.Write "<div align=""center"">"
		Response.Write "  <table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		Response.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""swfupload.asp?from=Common"">"
		Response.Write "      <tr>"
		Response.Write "        <td>"
		Response.Write "          <div align=""center"">"
		Response.Write "            <table width=""90%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		Response.Write "              <tr>"
		Response.Write "                <td height=""30""> &nbsp;&nbsp;上传文件个数"
		Response.Write "                  <input name=""UpFileNum"" type=""text"" value=""5"" size=""6"">"
		Response.Write "                  <input type=""button"" class=""button"" name=""Submit42"" value=""确定设定"" onClick=""AddUpFile();"">"
		Response.Write "                  <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
		Response.Write "                  添加水印</td>"
		Response.Write "              </tr>"
		Response.Write "              <tr>"
		Response.Write "                <td height=""30"" id=""FilesList""> </td>"
		Response.Write "              </tr>"
		Response.Write "            </table>"
		Response.Write "            </div>"
		Response.Write "        </td>"
		Response.Write "        <td width=""30%"" valign=""top""><br><br> <fieldset style=""width:100%;"">"
		Response.Write "          <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		Response.Write "            <tr>"
		Response.Write "              <td height=""20"">"
		Response.Write "                <div align=""center"">命名规则</div></td>"
		Response.Write "            </tr>"
		Response.Write "            <tr>"
		Response.Write "              <td height=""20"">"
		Response.Write "                <div align=""left"">"
		Response.Write "                  <input type=""radio"" name=""AutoReName"" value=""0"">"
		Response.Write "                  原名称不变</div></td>"
		Response.Write "            </tr>"
		Response.Write "            <tr>"
		Response.Write "              <td height=""20"">"
		Response.Write "                <div align=""left"">"
		Response.Write "                  <input type=""radio"" name=""AutoReName"" value=""1"">"
		Response.Write "                  &quot; 副件&quot;+文件名</div></td>"
		Response.Write "            </tr>"
		Response.Write "            <tr>"
		Response.Write "              <td height=""20"">"
		Response.Write "                <div align=""left"">"
		Response.Write "                  <input type=""radio"" name=""AutoReName"" value=""2"">"
		Response.Write "                  随机数+扩展名</div></td>"
		 Response.Write "           </tr>"
		Response.Write "            <tr>"
		Response.Write "              <td height=""20""><input type=""radio"" name=""AutoReName"" value=""3"">"
		Response.Write "              随机数+文件名</td>"
		Response.Write "            </tr>"
		Response.Write "            <tr>"
		Response.Write "              <td height=""20"">"
		Response.Write "                <div align=""left"">"
		Response.Write "                  <input name=""AutoReName"" type=""radio"" value=""4"" checked>"
		Response.Write "                  20060101121022</div></td>"
		Response.Write "            </tr>"
		Response.Write "          </table>"
		Response.Write "        </fieldset></td>"
		Response.Write "      </tr>"
		Response.Write "      <tr>"
		Response.Write "        <td height=""40"" colspan=""2""> <table align=""center"" width=""60%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		Response.Write "            <tr>"
		Response.Write "              <td> <div align=""center"">"
		Response.Write "                  <input class=""button"" type=""submit"" id=""BtnSubmit"" name=""Submit"" onClick=""PromptInfo();"" value="" 确 定 "">"
		Response.Write "                  <input name=""Path"" value=""" & UploadDir & """ type=""hidden"" id=""Path"">"
		Response.Write "                  <input name=""UpLoadFrom"" value=""" & UpLoadFrom & """ type=""hidden"" id=""UpLoadFrom"">"
		Response.Write "                </div></td>"
		Response.Write "              <td><div align=""center"">"
		Response.Write "                  <input class=""button"" type=""reset"" id=""ResetForm"" name=""Submit3"" value="" 重 填 "">"
		Response.Write "                </div></td>"
		Response.Write "              <td><div align=""center"">"
		Response.Write "                  <input class=""button"" onClick=""dialogArguments.location.reload();window.close();"" type=""button"" name=""Submit2"" value="" 关 闭 "">"
		Response.Write "                </div></td>"
		Response.Write "            </tr>"
		Response.Write "          </table></td>"
		 Response.Write "     </tr>"
		Response.Write "    </form>"
		Response.Write "  </table>"
		Response.Write "</div>"
		Response.Write "<div id=""LayerPrompt"" style=""position:absolute; z-index:1; left: 112px; top: 28px; background-color: #ffffee; layer-background-color: #00CCFF; border: 1px solid #f9c943; width: 300px; height: 63px; visibility: hidden;"">"
		Response.Write "  <table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "    <tr>"
		Response.Write "      <td><div>&nbsp;请稍等，正在上传文件<img src='../../images/default/wait.gif' align='absmiddle'></div></td>"
		Response.Write "      <td style='display:none' width=""35%""><div align=""left""><font id=""ShowInfoArea"" size=""+1""></font></div></td>"
		Response.Write "    </tr>"
		Response.Write "  </table>"
		Response.Write "</div>"
		Response.Write "</body>"
		Response.Write "</html>"
		Response.Write "<script language=""JavaScript"">" & vbCrLf
		Response.Write "var ForwardShow=true;" & vbCrLf
		Response.Write "function AddUpFile()" & vbCrLf
		Response.Write " {" & vbCrLf
		Response.Write "  var UpFileNum = document.all.UpFileNum.value;" & vbCrLf
		Response.Write "  if (UpFileNum=='')" & vbCrLf
		Response.Write "    UpFileNum=5;" & vbCrLf
		Response.Write "  var i,Optionstr;" & vbCrLf
		Response.Write "      Optionstr = '<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""0"">';" & vbCrLf
		Response.Write "  for (i=1;i<=UpFileNum;i++)" & vbCrLf
		Response.Write "      {" & vbCrLf
		Response.Write "       Optionstr = Optionstr+'<tr><td>&nbsp;文&nbsp;件&nbsp;'+i+'</td><td>&nbsp;<input type=""file"" accept=""html"" size=""20"" class=""textbox"" name=""File'+i+'"">&nbsp;</td></tr>';" & vbCrLf
		Response.Write "       }" & vbCrLf
		Response.Write "    Optionstr = Optionstr+'</table>';" & vbCrLf
		Response.Write "    document.all.FilesList.innerHTML = Optionstr;" & vbCrLf
		Response.Write "  }" & vbCrLf
		Response.Write "function ShowPromptMessage()" & vbCrLf
		Response.Write "{ " & vbCrLf
		Response.Write "    var TempStr=ShowInfoArea.innerText;" & vbCrLf
		Response.Write "    if (ForwardShow==true)" & vbCrLf
		Response.Write "    {" & vbCrLf
		Response.Write "        if (TempStr.length>4) ForwardShow=false;" & vbCrLf
		Response.Write "        ShowInfoArea.innerText=TempStr+'.';" & vbCrLf
		Response.Write "    } " & vbCrLf
		Response.Write "    else" & vbCrLf
		Response.Write "    {" & vbCrLf
		Response.Write "        if (TempStr.length==1) ForwardShow=true;" & vbCrLf
		Response.Write "        ShowInfoArea.innerText=TempStr.substr(0,TempStr.length-1);" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "}" & vbCrLf
		Response.Write "function PromptInfo()" & vbCrLf
		Response.Write "{" & vbCrLf
		Response.Write "    document.all.ResetForm.disabled=true;" & vbCrLf
		Response.Write "    LayerPrompt.style.visibility='visible';" & vbCrLf
		Response.Write "    return true;" & vbCrLf
		Response.Write "}" & vbCrLf
		Response.Write "window.setInterval('ShowPromptMessage()',150)" & vbCrLf
		Response.Write "AddUpFile();" & vbCrLf
		Response.Write "</script>" & vbCrLf
		End Sub
End Class
%> 
