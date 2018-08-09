<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%


Dim KSCls
Set KSCls = New Special
KSCls.Kesion()
Set KSCls = Nothing

Class Special
        Private KS, KSRFObj,KMRFObj
		Private CurrPage,PageStyle,PerPageNumber,ID
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
		  
		  CurrPage=KS.ChkClng(KS.S("Page"))
		  If CurrPage<=0 Then CurrPage=CurrPage+1

		  Dim FileContent
		  ID=KS.ChkClng(KS.S("ID"))
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select * From KS_Special Where SpecialID=" & ID,Conn,1,1
		  If RS.Eof And RS.Bof Then
		  Call KS.Alert("您要查看的专题已删除。或是您非法传递注入参数!",""):Exit Sub
		  End If
		  If KS.Setting(78)="1" Then response.Redirect(KS.GetSpecialPath(ID,RS("SpecialEname"),RS("FsoSpecialIndex"))):rs.close:set rs=nothing:conn.close:set conn=nothing
		  
		   FCls.RefreshType = "Special"
		   FCls.RefreshFolderID = RS("ClassID")
		   FCls.CurrSpecialID = ID
		   FileContent = KSRFObj.LoadTemplate(RS("TemplateID"))
		   FileContent = KSRFObj.ReplaceSpecialContent(FileContent,RS)
		                           
		   If Trim(FileContent) = "" Then FileContent = "专题页模板不存在!"
		     FileContent=KSRFObj.KSLabelReplaceAll(FileContent)
			 FileContent=Replace(FileContent,"{KS:PageList}",KS.GetPageList("?ID=" & ID,FCls.PageStyle,CurrPage,FCls.TotalPage,true))
		   RS.Close
			
			
			Dim LabelParamStr:LabelParamStr=Application("PageParam")
			If LabelParamStr<>"" And Not IsNull(LabelParamStr) Then
				 Dim XMLDoc,XMLSql,LabelStyle,KMRFOBJ,SQLStr,TableName
				 Dim ParamNode,IncludeSubClass,ModelID,OrderStr,PrintType,PageStyle,PicStyle,ShowPicFlag,FieldStr,Param
				 Dim PerPageNumber,TotalPut,PageNum,TempStr
				 Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				 If XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
					 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
					 ModelID         = ParamNode.getAttribute("modelid") : If Not IsNumeric(ModelID) Then ModelID=1
					 PrintType       = ParamNode.getAttribute("printtype") : If Not IsNumeric(PrintType) Then PrintType=1
					 PageStyle       = ParamNode.getAttribute("pagestyle") : If PageStyle="" Or IsNull(PageStyle) Then PageStyle=1
					 PicStyle        = ParamNode.getAttribute("picstyle")
					 OrderStr        = ParamNode.getAttribute("orderstr") : If OrderStr="" Or IsNull(OrderStr) Then OrderStr="ID Desc"
					 ShowPicFlag     = ParamNode.getAttribute("showpicflag") : If ShowPicFlag="" Or IsNull(ShowPicFlag) Then ShowPicFlag=false
					 PerPageNumber   = ParamNode.getAttribute("num") : If Not IsNumeric(PerPageNumber) Then PerPageNumber=10
					 
					 Param = " Where I.Verific=1 And I.DelTF=0"
					 Param= Param & KS.GetSpecialPara(ModelID,ID)
					 
					 Set KMRFObj= New RefreshFunction
					 Set KMRFObj.ParamNode=ParamNode
				     Call KMRFObj.LoadField(ModelID,PrintType,PicStyle,ShowPicFlag,FieldStr,TableName,Param)
				
					If Lcase(Left(Trim(OrderStr),2))<>"id" Then  OrderStr=OrderStr & ",I.ID Desc"
					If ModelID=0 Then			
					 TableName="[KS_ItemInfo]"
					Else
					 TableName=KS.C_S(ModelID,2)
					End If
					SqlStr = "SELECT " & FieldStr & " FROM " & TableName & " I " & Param & " ORDER BY I.IsTop Desc," & OrderStr
					'response.write sqlstr
					Set RS=Server.CreateObject("ADODB.RECORDSET")
					RS.Open SqlStr, Conn, 1, 1
					If RS.EOF And RS.BOF Then
						TempStr = "<p>此专题下没有信息</p>"
					Else
						PerPageNumber=cint(PerPageNumber)
						TotalPut = Conn.Execute("select Count(id) from " & TableName & " I " & Param)(0)
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
						Call KMRFObj.LoadPageParam(XMLSQL,ParamNode,ModelID)
						LabelStyle=Application("LabelStyle")
						TempStr = KMRFObj.ExplainGerericListLabelBody(LabelStyle)
						XMLSql=Empty
						
						FCls.PageStyle=PageStyle       '分页样式
						FCls.TotalPage=PageNum         '总页数
						TempStr = TempStr & KS.GetPrePageList(FCls.PageStyle,"条记录",FCls.TotalPage,CurrPage,TotalPut,PerPageNumber) & KS.GetPageList("?ID=" & ID,FCls.PageStyle,CurrPage,FCls.TotalPage, True)
						
					End If
				
					RS.Close:Set RS=Nothing					
					XMLDoc= Empty : Set ParamNode=Nothing
				End If	
				
			End If
			
			FileContent=Replace(FileContent,"{Tag:Page}",TempStr)
			
		    KS.Echo FileContent
	   End Sub
	   
End Class
%>
