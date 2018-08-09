<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%


Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SiteIndex
        Private KS, KSRFObj,KMRFObj,FileContent
		Private CurrPage,RSObj,PerPageNumber,PageStyle
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRFObj = New Refresh
		  Set KMRFObj= New RefreshFunction
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KMRFObj=Nothing
		End Sub
		Public Sub Kesion()
		 Dim ClassID:ClassID=KS.ChkClng(KS.S("ClassID"))
		 CurrPage=KS.ChkClng(KS.S("Page"))
		 If CurrPage<=0 Then CurrPage=CurrPage+1
		 IF ClassID=0 Then Exit Sub
		 Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		  RSObj.Open "Select * From KS_SpecialClass Where ClassID=" & ClassID,Conn,1,1
		  If RSObj.Eof And RSObj.Bof Then
		  Call KS.Alert("您要查看的栏目已删除。或是您非法传递注入参数!",""):Exit Sub
		  End If
		  FCls.FromAspPage=True
		  FCls.RefreshType = "ChannelSpecial"
		  FCls.RefreshFolderID = ClassID
		  FileContent = KSRFObj.LoadTemplate(RSObj("TemplateID"))
		  FileContent = KSRFObj.ReplaceSpecialClass(FileContent)
		  FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
		   		  
		  
		   Dim LabelParamStr:LabelParamStr=Application("PageParam")
		   If LabelParamStr<>"" And Not IsNull(LabelParamStr) Then
				 Dim XMLDoc,XMLSql,LabelStyle
				 Dim ParamNode,IncludeSubClass,ModelID,OrderStr,PrintType,PageStyle,PicStyle,FieldStr,Param
				 Dim PerPageNumber,TotalPut,PageNum,TempStr,SQLStr,RS
				 Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				 If XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
					 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
					 PrintType       = ParamNode.getAttribute("printtype") : If Not IsNumeric(PrintType) Then PrintType=1
					 PageStyle       = ParamNode.getAttribute("pagestyle") : If PageStyle="" Or IsNull(PageStyle) Then PageStyle=1
					 PerPageNumber   = ParamNode.getAttribute("num") : If Not IsNumeric(PerPageNumber) Then PerPageNumber=10
					 SqlStr="Select S.specialid,S.classid,S.SpecialName,S.SpecialEname,S.FsoSpecialIndex,S.SpecialAddDate as AddDate,S.PhotoUrl,S.SpecialNote As Intro,S.creater,C.ClassName as SpecialClassName From KS_Special S Inner Join KS_SpecialClass C On S.ClassID=C.ClassID Where S.ClassID=" & ClassID & " Order By S.SpecialID Desc"
					 
						Set RS=Server.CreateObject("ADODB.RECORDSET")
						RS.Open SqlStr, Conn, 1, 1
						If RS.EOF And RS.BOF Then
							TempStr = "<p>此分类下没有专题</p>"
						Else
							PerPageNumber=cint(PerPageNumber)
							TotalPut = Conn.Execute("Select count(S.specialid) From KS_Special S Inner Join KS_SpecialClass C On S.ClassID=C.ClassID Where S.ClassID=" & ClassID)(0)
							if (TotalPut mod PerPageNumber)=0 then
									PageNum = TotalPut \ PerPageNumber
							else
									PageNum = TotalPut \ PerPageNumber + 1
							end if
							If CurrPage >1 and (CurrPage - 1) * PerPageNumber < totalPut Then
								RS.Move (CurrPage - 1) * PerPageNumber
							Else
								CurrPage = 1
							End If
							Set XMLSQL=KS.ArrayToXml(RS.GetRows(PerPageNumber),RS,"row","root")
							Call KMRFObj.LoadPageParam(XMLSQL,ParamNode,0)
							LabelStyle=Application("LabelStyle")
							TempStr = KMRFObj.ExplainSpecialListLabelBody(LabelStyle)
							XMLSql=Empty
							
							FCls.PageStyle=PageStyle       '分页样式
							FCls.TotalPage=PageNum         '总页数
							TempStr = TempStr & KS.GetPrePageList(FCls.PageStyle,"个",FCls.TotalPage,CurrPage,TotalPut,PerPageNumber) & KS.GetPageList("?ClassID=" & ClassID,FCls.PageStyle,CurrPage,FCls.TotalPage, True)
							
						End If
					
						RS.Close:Set RS=Nothing					
						XMLDoc= Empty : Set ParamNode=Nothing
				 End If
		   End If

		  
		 
			 FileContent=Replace(FileContent,"{Tag:Page}",TempStr)

		  Response.Write FileContent 
		  RSObj.Close:Set RSObj=Nothing
		End Sub
		
End Class
%>
