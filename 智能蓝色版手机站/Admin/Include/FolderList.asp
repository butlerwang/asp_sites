<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%

Dim KSCls
Set KSCls = New FolderList
KSCls.Kesion()
Set KSCls = Nothing

Class FolderList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Function Kesion()
		Dim CurrPath, FsoObj, FolderObj, SubFolderObj, FileObj, I, FsoItem
		Dim ParentPath, FileExtName, AllowShowExtNameStr
		AllowShowExtNameStr = "htm,html,shtml"
		CurrPath = Request("CurrPath")
		If CurrPath = "" Then CurrPath = "/"
		Set FsoObj = KS.InitialObject(KS.Setting(99))
		Set FolderObj = FsoObj.GetFolder(Server.MapPath(CurrPath))
		Set SubFolderObj = FolderObj.SubFolders
		Set FileObj = FolderObj.Files
		
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>"
		Response.Write "<link href='Admin_Style.CSS' rel='stylesheet'>"
		Response.Write "<body topmargin='0' leftmargin='0' scroll=yes>"
		Response.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		Response.Write "  <tr>"
		Response.Write "    <td height='20' class='sort'> <div align='center'><font color='#000000'>名称</font></div></td>"
		Response.Write "    <td height='20' class='sort'> <div align='center'><font color='#000000'>类型</font></div></td>"
		Response.Write "    <td height='20' class='sort'> <div align='center'><font color='#000000'>修改日期</font></div></td>"
		Response.Write "  </tr>"
		
		 For Each FsoItem In SubFolderObj
		 
		Response.Write "  <tr>"
		Response.Write "    <td height='20'>"
		Response.Write "        <table border='0' cellspacing='0' cellpadding='0'>"
		Response.Write "          <tr title='双击鼠标进入此目录'>"
		Response.Write "          <td><img src='../Images/Folder/folderclosed.gif' width='24' height='22'></td>"
		Response.Write "            <td> <span class='FolderItem' Path='" & FsoItem.name & "' onDblClick=""OpenFolder('" & FsoItem.name & "');"" onClick='SelectFolder(this);'>"
		Response.Write FsoItem.name
		Response.Write "              </span> </td>"
		Response.Write "          </tr>"
		Response.Write "        </table>"
		Response.Write "      </div></td>"
		Response.Write "    <td height='20'>"
		Response.Write "      <div align='center'>目录</div></td>"
		Response.Write "    <td height='20'>"
		Response.Write "      <div align='center'>" & FsoItem.size & "</div></td>"
		Response.Write "  </tr>"
		  Next
		For Each FsoItem In FileObj
			FileExtName = LCase(Mid(FsoItem.name, InStrRev(FsoItem.name, ".") + 1))
			If KS.CheckFileShowOrNot(AllowShowExtNameStr, FileExtName) = True Then
		
		Response.Write "<tr title='单击选择文件'>"
		Response.Write "    <td height='20'>"
		Response.Write "      <table width='100%' border='0' cellspacing='0' cellpadding='0'>"
		Response.Write "        <tr>"
		Response.Write "          <td width='3%'>&nbsp;</td>"
		Response.Write "          <td width='97%'><span class='FolderItem' File='" & FsoItem.name & "' onDblClick=""parent.SelectFile('" & replace(FsoItem.name,"'","\'") & "');"" onClick=""SelectFile(this,'" & replace(FsoItem.name,"'","\'") & "');"">"
		Response.Write FsoItem.name
		Response.Write "            </span></td>"
		Response.Write "        </tr>"
		Response.Write "      </table>"
		Response.Write "    </td>"
		Response.Write "    <td height='20'> <div align='center'>"
		Response.Write FsoItem.Type
		Response.Write "      </div></td>"
		Response.Write "    <td height='20'> <div align='center'>"
		Response.Write FsoItem.DateLastModified
		Response.Write "      </div></td>"
		Response.Write "  </tr>"
		
			End If
		Next
		
		Response.Write "</table></body></html>"
		
		Set FsoObj = Nothing
		Set FolderObj = Nothing
		Set FileObj = Nothing
		
		Response.Write "<script language='JavaScript'>"
		Response.Write "var CurrPath='" & CurrPath & "';"
		Response.Write "var FileName='';"
		Response.Write "function SelectFile(Obj,file)"
		Response.Write "{"
		Response.Write "    for (var i=0;i<document.all.length;i++)"
		Response.Write "    {"
		Response.Write "        if (document.all(i).className=='FolderSelectItem') document.all(i).className='FolderItem';"
		Response.Write "    }"
		Response.Write "    Obj.className='FolderSelectItem';"
		Response.Write "    FileName=file;"
		Response.Write "}"
		Response.Write "function SelectFolder(Obj)"
		Response.Write "{   FileName='';"
		Response.Write "    for (var i=0;i<document.all.length;i++)"
		Response.Write "    {"
		Response.Write "        if (document.all(i).className=='FolderSelectItem') document.all(i).className='FolderItem';"
		Response.Write "    }"
		Response.Write "    Obj.className='FolderSelectItem';"
		Response.Write "}"
		Response.Write "function OpenFolder(Obj)"
		Response.Write "{ "
		Response.Write "    var SubmitPath='';"
		Response.Write "    if (CurrPath=='/') SubmitPath=CurrPath+Obj;"
		Response.Write "    else SubmitPath=CurrPath+'/'+Obj;"
		Response.Write "    location.href='FolderList.asp?CurrPath='+SubmitPath;"
		Response.Write "    AddFolderList(parent.document.getElementById('FolderSelectList'),SubmitPath,SubmitPath);"
		Response.Write "}"
		Response.Write "function AddFolderList(SelectObj, Label, LabelContent)"
		Response.Write "{"
		Response.Write "    var i=0,AddOption;"
		Response.Write "    if (!SearchOptionExists(SelectObj,Label))"
		Response.Write "    {"
		Response.Write "        AddOption = document.createElement('OPTION');"
		Response.Write "        AddOption.text=Label;"
		Response.Write "        AddOption.value=LabelContent;"
		Response.Write "        SelectObj.options.add(AddOption);"
		Response.Write "        SelectObj.options[SelectObj.length-1].selected=true;"
		Response.Write "    }"
		Response.Write "}"
		Response.Write "function SearchOptionExists(Obj, SearchText)"
		Response.Write "{"
		Response.Write "    var i;"
		Response.Write "    for(i=0;i<Obj.length;i++)"
		Response.Write "    {"
		Response.Write "        if (Obj.options[i].text==SearchText)"
		Response.Write "        {"
		Response.Write "            Obj.options[i].selected=true;"
		Response.Write "            return true;"
		Response.Write "        }"
		Response.Write "    }"
		Response.Write "    return false;"
		Response.Write "}"
		Response.Write "</script>"
		End Function
End Class
%> 
