<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%


Dim KSCls
Set KSCls = New SpecialIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SpecialIndex
        Private KS, KSRFObj
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRFObj = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		  Dim FileContent,FsoIndex:FsoIndex=KS.Setting(5)
				   FileContent = KSRFObj.LoadTemplate(KS.Setting(111))
				   FCls.RefreshType = "SpecialIndex"  '设置刷新类型，以便取得当前位置导航等
				   FCls.RefreshFolderID = "0"         '设置当前刷新目录ID 为"0" 以取得通用标签
				   FCls.CurrSpecialID="" '清除当前专题ID
				   If Trim(FileContent) = "" Then FileContent = "首页模板不存在!"
				    FileContent=KSRFObj.KSLabelReplaceAll(FileContent)
		   Response.Write FileContent  
		End Sub
End Class
%>
