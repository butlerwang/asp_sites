<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<%
Dim TempClassList,InstallDir,CurrPath,JSConfig,KS,KSRObj,FolderID,TempSymbol
Dim JSID,JSRS,SQLStr,JSName,JSFunctionFlag,Descript,Action,RSCheck,FileUrl,JSType,JSFileName
Dim KeyWord,SearchType,StartDate,EndDate
  
'收集搜索参数
KeyWord=Request("KeyWord")
SearchType=Request("SearchType")
StartDate = Request("StartDate")
EndDate = Request("EndDate")

FileUrl=Request("FileUrl") '便于添加完毕后返回
Set KS=New PublicCls
Set KSRObj=New Refresh
	JSFileName=Replace(Replace(Trim(Request.Form("JSFileName")),"""",""),"'","")
	if instr(JSFileName,";")<>0 or instr(lcase(JSFileName),".asp")<>0 or instr(lcase(JSFileName),".php")<>0 or instr(lcase(JSFileName),".cer")<>0 or instr(lcase(JSFileName),".asa")<>0 then
       Call KS.AlertHistory("JS名称格式不合法!",-1)
	   Set KS=Nothing
	   Response.End
	end if

Set JSRS=Server.CreateObject("Adodb.RecordSet")
Select Case Request.Form("Action")
 Case "Add" 
    JSName= Replace(Replace(Trim(Request.Form("JSName")),"""",""),"'","")
    Descript=Replace(Trim(Request.Form("Descript")),"'","")
    JSConfig=Trim(Request.Form("JSConfig"))
	JSType=Request.Form("JSType")
	FolderID=Request.Form("ParentID")
    IF FolderID="" Then FolderID="0"
	IF JSType="" Then JSType=0
    IF JSName="" THEN
       Call KS.AlertHistory("JS名称不能为空!",-1)
	   Set KS=Nothing
	   Response.End
    END IF
	IF UCASE(Right(JSFileName,3))<>".JS" THEN
	  Call KS.AlertHistory("JS文件名的扩展名必须是.js",-1)
	  Set KS=Nothing
	  Response.End
	END IF
    IF JSConfig="" THEN
      Call KS.AlertHistory("JS内容不能为空!",-1)
	  Set KS=Nothing
	  Response.End
    END IF
	JSName="{JS_" & JSName & "}"
	JSRS.Open "Select JSName From [KS_JSFile] Where JSName='" & JSName & "' Or JSFileName='" & JSFileName &"'",Conn,1,1
	IF Not JSRS.EOF Then
	  if Trim(JSRS("JSName"))=JSName Then
	   Response.Write("<script>alert('JS名称已经存在!');location.href='" & FileUrl & "?Action=Add&FolderID=" & FolderID &"';</script>")
	  else
	   Response.Write("<script>alert('JS文件名已经存在!');location.href='" & FileUrl & "?Action=Add&FolderID=" & FolderID &"';</script>")
	  end if
	  JSRS.Close
	  Conn.Close
	  Set JSRS=Nothing
	  Set Conn= Nothing
	  Set KS=Nothing
	  Response.End
	ELSE
	    JSRS.Close
		JSRS.Open "Select * From [KS_JSFile] Where (JSID is Null)",Conn,1,3
		JSRS.AddNew
		  Do While True
		    '生成ID  年+6位随机
            JSID = Year(Now()) & KS.MakeRandom(6)
            Set RSCheck = conn.execute("Select JSID from [KS_JSFile] Where JSID='" & JSID & "'")
             If RSCheck.EOF And RSCheck.BOF Then
              RSCheck.Close
			  Set RSCheck=Nothing
              Exit Do
             End If
          Loop
		 JSRS("JSID")=JSID
		 JSRS("JSName")=JSName
		 JSRS("JSFileName")=JSFileName
		 JSRS("Description")=Descript
		 JSRS("JSConfig")=JSConfig
		 JSRS("JSType")=JSType
		 JSRS("AddDate")=now
		 JSRS("OrderID")=1
		 JSRS("FolderID")=FolderID
		 JSRS.Update
		IF JSType=0 Then
		    TempSymbol="&OpStr=JS管理  >> 系统JS &ButtonSymbol=SysJSList"
		 ELSE
		    TempSymbol="&OpStr=JS管理  >> 自由JS &ButtonSymbol=FreeJSList"
		 END IF
		 KSRObj.RefreshJS(JSName)
		 JSRS.Close
		 Set JSRS=Nothing
		 Set KSRObj=Nothing
     	Response.Write("<script>if (confirm('成功提示:\n\n添加JS成功,继续添加JS吗?')){location.href='" & FileUrl & "?Action=Add&FolderID=" & FolderID & "';}else{top.frames['BottomFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & FolderID &TempSymbol &"';top.frames['MainFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "include/JS_Main.asp?FolderID=" & FolderID &"&JSType=" & JSType & "';}</script>") 
	END IF
Case "Edit"
    Dim Page
	Page=Request.Form("Page")
    JSID=Trim(Request.Form("JSID"))
    JSName= Replace(Replace(Trim(Request.Form("JSName")),"""",""),"'","")
    Descript=Replace(Trim(Request.Form("Descript")),"'","")
    JSConfig=Trim(Request.Form("JSConfig"))
	JSType=Request.Form("JSType")
	FolderID=Request.Form("ParentID")
	IF FolderID="" Then FolderID="0"
	IF JSType="" Then JSType=0
    IF JSName="" THEN
       Call KS.AlertHistory("JS名称不能为空!",-1)
	   Set KS=Nothing
	   Response.End
    END IF
    IF JSConfig="" THEN
      Call KS.AlertHistory("JS内容不能为空!",-1)
	  Set KS=Nothing
	  Response.End
    END IF
	JSName="{JS_" & JSName & "}"
	JSRS.Open "Select JSName From [KS_JSFile] Where JSID <>'" & JSID &"' AND JSName='" & JSName & "'",Conn,1,1
	IF Not JSRS.EOF Then
	  Response.Write("<script>alert('JS名称已经存在!');location.href='" & FileUrl & "?Page=" & Page & "&JSID=" & JSID & "';</script>")
	  JSRS.Close
	  Conn.Close
	  Set JSRS=Nothing
	  Set Conn= Nothing
	  Set KS=Nothing
	  Response.End
	ELSE
	    JSRS.Close
		JSRS.Open "Select * From [KS_JSFile] Where JSID='" & JSID &"'",Conn,1,3
		 JSRS("JSName")=JSName
		 JSRS("Description")=Descript
		 JSRS("JSConfig")=JSConfig
		 JSRS("JSType")=JSType
		 JSRS("FolderID")=FolderID
		 JSRS.Update
		 KSRObj.RefreshJS(JSName)

		 IF KeyWord="" Then
		    IF JSType=0 Then
		        TempSymbol="&OpStr=JS管理  >> 系统JS &ButtonSymbol=SysJSList"
		    ELSE
		        TempSymbol="&OpStr=JS管理  >> 自由JS &ButtonSymbol=FreeJSList"
		    END IF
     	   Response.Write("<script>alert('成功提示:\n\nJS修改成功!');top.frames['BottomFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & FolderID &TempSymbol &"';top.frames['MainFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "include/JS_Main.asp?FolderID="& FolderID & "&Page=" & Page & "&JSType=" & JSType & "';</script>") 
		 ELSE
		    IF JSType=0 Then
		        TempSymbol="OpStr=JS管理  >> <font color=red>搜索系统JS结果</font>&ButtonSymbol=SysJSSearch"
		    ELSE
		        TempSymbol="OpStr=JS管理  >> <font color=red>搜索自由JS结果</font>&ButtonSymbol=FreeJSSearch"
		    END IF
     	   Response.Write("<script>alert('成功提示:\n\nJS修改成功!');top.frames['BottomFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?" &TempSymbol &"';top.frames['MainFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "include/JS_Main.asp?KeyWord="& KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate &"&EndDate=" & EndDate &"&Page=" & Page & "&JSType=" & JSType & "';</script>") 
		 END IF
	END IF
		 JSRS.Close
		 Set JSRS=Nothing
		 Set KSRObj=Nothing
End Select
%>
 
