<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%

Dim KSCls
Set KSCls = New SiteMaps
KSCls.Kesion()
Set KSCls = Nothing

Class SiteMaps
        Private KS, KSR,Maps,ClassXml,TreeStr
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		           Dim FileContent
		           Dim MapTemplatePath:MapTemplatePath=KS.Setting(3) & KS.Setting(90) & "common/map.html"  '模板地址
				   FileContent = KSR.LoadTemplate(MapTemplatePath)    
				   FCls.RefreshType = "map" '设置刷新类型，以便取得当前位置导航等
				   FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
				   Call MapList()
				   FileContent=Replace(FileContent,"{$ShowMap}",Maps)
				   FileContent=KSR.KSLabelReplaceAll(FileContent)
				   response.write FileContent
		End Sub
		
		Sub MapList()
		    Call KS.LoadClassConfig()
			Dim Node,TJ,SpaceStr,k
			Set ClassXML=Application(KS.SiteSN&"_class")
				If IsOBject(ClassXml) Then
				  For Each Node In ClassXML.DocumentElement.SelectNodes("class[@ks26=1][@ks10=1]")
				    if ks.chkclng(ks.c_s(Node.SelectSingleNode("@ks12").text,21))=1 then
				      TJ=Node.SelectSingleNode("@ks10").text
				      TreeStr = TreeStr  & "<div class='maplist'>"&vbcrlf & "   <div class=""classname"">" & KS.GetClassNP(Node.SelectSingleNode("@ks0").text) &"</div>" &vbcrlf                   
					  Call SubMapList(Node.SelectSingleNode("@ks0").text)
					  TreeStr=TreeStr & "</div>"
					end if
				  Next
				End If
			 Maps=TreeStr
	  End Sub
	 
	 Sub SubMapList(TN)
	   Dim Node,K,TJ,SpaceStr
	   If IsOBject(ClassXml) Then
				  For Each Node In ClassXML.DocumentElement.SelectNodes("class[@ks26=1][@ks13='" & TN &"']")
				   if ks.chkclng(ks.c_s(Node.SelectSingleNode("@ks12").text,21))=1 then
						If Node.SelectSingleNode("@ks19").text>0 Then
	                     TreeStr = TreeStr & "<div class=""maplist2"">" & vbcrlf & "   <span class=""classname2"">" & KS.GetClassNP(Node.SelectSingleNode("@ks0").text) & "：</span>"
						 Call SubMapList2(Node.SelectSingleNode("@ks0").text)  
						 TreeStr = TreeStr & "</div>"&vbcrlf   
						Else
	                     TreeStr = TreeStr & "<span>" & KS.GetClassNP(Node.SelectSingleNode("@ks0").text) & "</span>"
						End If
				   end if
				  Next
		End If
	 End Sub
	 Sub SubMapList2(TN)
	   Dim Node,K,TJ,SpaceStr
	   If IsOBject(ClassXml) Then
				  For Each Node In ClassXML.DocumentElement.SelectNodes("class[@ks26=1][@ks13='" & TN &"']")
				     if ks.chkclng(ks.c_s(Node.SelectSingleNode("@ks12").text,22))=1 then
	                     TreeStr = TreeStr & "<span>" & KS.GetClassNP(Node.SelectSingleNode("@ks0").text) & "</span>"
					 end if
				  Next
		End If
	 End Sub
End Class
%>
