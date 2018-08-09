<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%> 
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.CollectCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
Response.Buffer = True
Server.ScriptTimeout = 999
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim KSCls
Set KSCls = New Collect_ItemCollection
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_ItemCollection
        Private KS
		Private KMCObj
		Private ConnItem
		Private Action, ItemID, CollecType
		Private FoundErr, ErrMsg
		Private SqlItem, RsItem
		Private Arr_Item, Arr_Filters, Arr_Historys, myCache, CollecTest, Content_View,Arr_Field
		Private CacheTemp
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KMCObj=New CollectPublicCls
		  Set ConnItem = KS.ConnItem()
		End Sub
        Private Sub Class_Terminate()
		 Call KS.CloseConnItem()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KMCObj=Nothing
		End Sub
		Sub Kesion()
		
		FoundErr = False
		CacheTemp = KS.SiteSN
		
		'检察表单
		Call DelNews
		Call CheckForm
		If FoundErr <> True Then
		   Call SetCache
		   If FoundErr <> True Then
				 ErrMsg = "<meta http-equiv=""refresh"" content=""3;url=Collect_ItemCollecFast.asp?ItemNum=1&ListNum=1&NewsSuccesNum=0&NewsFalseNum=0&ImagesNumAll=0"">"
		   End If
		End If
		If FoundErr = True Then
		   Call KS.AlertHistory(ErrMsg,-1)
		Else
		   Call Main
		End If
		End Sub
		Sub Main()
		
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<title>采集系统</title>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Admin_Style.css"">"
		Response.Write "<style type=""text/css"">"
		Response.Write "<!--" & vbCrLf
		Response.Write ".STYLE1 {" & vbCrLf
		Response.Write "    color: #FF0000;" & vbCrLf
		Response.Write "    font-weight: bold;" & vbCrLf
		Response.Write "}" & vbCrLf
		Response.Write "-->" & vbCrLf
		Response.Write "</style>"
		Response.Write "</head>"
		'Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"" oncontextmenu=""return false"">"
		Response.Write "<div class=""topdashed sort"">采集系统采集管理</div>"
		
		Response.Write "<br>"
		Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
		Response.Write "    <tr>"
		Response.Write "      <td height=""100"" colspan=""2"" align=center>"
		Response.Write "        <p><br>"
		Response.Write "          <br>"
		Response.Write "          <br>"
		Response.Write "      欢迎使用KesionCMS自带采集系统，正在初始化数据，请稍后...      </p>"
		Response.Write "        <p><span class=""STYLE1"">使用声明: 采集信息如果涉及到版权问题与科兴信息技术有限公司无关!</span><br>"
		 Response.Write "         <br>"
		Response.Write ErrMsg & "         </p></td>"
		Response.Write "    </tr>"
		Response.Write "</table>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
		
		Sub CheckForm()
		
		   '提取表单
		   Action = Trim(Request("Action"))
		   ItemID = Trim(Request("ItemID"))
		   CollecType = Trim(Request("CollecType"))
		   CollecTest = Trim(Request("CollecTest"))
		   Content_View = Trim(Request("Content_View"))
		   Session("taskf")=Request("f")
		   '检察表单
		   If Action <> "Start" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "●参数不足!\n"
		   End If
		   If ItemID = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "●请您选择项目!\n"
		   Else
			  If InStr(ItemID, ",") > 0 Then
				 ItemID = Replace(ItemID, " ", "")
			  End If
		   End If
	
		   If CollecTest = "yes" Then
			  CollecTest = True
		   Else
			  CollecTest = False
		   End If
		   If Content_View = "yes" Then
			  Content_View = True
		   Else
			  Content_View = False
		   End If
		End Sub
		Sub SetCache()
		   '项目信息
		   SqlItem = "select * From KS_CollectItem where ItemID in(" & ItemID & ")"
		   Set RsItem = Server.CreateObject("adodb.recordset")
		   RsItem.Open SqlItem, ConnItem, 1, 1
		   If Not RsItem.EOF Then
			  Arr_Item = RsItem.GetRows()
		   End If
		   RsItem.Close:Set RsItem = Nothing
		
		   Set myCache = New ClsCache
		   myCache.name = CacheTemp & "items"
		   Call myCache.clean
		   If IsArray(Arr_Item) = True Then
			  myCache.add Arr_Item, DateAdd("n", 1000, Now)
		   Else
			  FoundErr = True
			  ErrMsg = ErrMsg & "发生意外错误！"
		   End If
		
		   '自定义字段
		   SqlItem="Select FieldID,FieldName,BeginStr,EndStr,ShowType,ItemID,ChannelID From KS_FieldRules Where ItemID in(" & ItemID & ") Order by OrderID"
		   Set RsItem = Server.CreateObject("adodb.recordset")
		   RsItem.Open SqlItem, ConnItem, 1, 1
		   If Not RsItem.EOF Then
			  Arr_Field = RsItem.GetRows()
			 Application.Lock()
		     Set Application("CollectFieldRules")=KS.ArrayToXml(Arr_Field,RsItem,"row","")
		     Application.UnLock()

		   End If
		   RsItem.Close:Set RsItem = Nothing
		
		   Set myCache = New ClsCache
		   myCache.name = CacheTemp & "Field"
		   Call myCache.clean
		   If IsArray(Arr_Field) = True Then

			  myCache.add Arr_Field, DateAdd("n", 1000, Now)
		   End If
		   
		
		   '过滤信息
		   SqlItem = "select * From KS_Filters where Flag=True"
		   Set RsItem = Server.CreateObject("adodb.recordset")
		   RsItem.Open SqlItem, ConnItem, 1, 1
		   If Not RsItem.EOF Then
			  Arr_Filters = RsItem.GetRows()
		   End If
		   RsItem.Close:Set RsItem = Nothing
		
		   myCache.name = CacheTemp & "filters"
		   Call myCache.clean
		   If IsArray(Arr_Filters) = True Then
			  myCache.add Arr_Filters, DateAdd("n", 1000, Now)
		   End If
		
		   '历史记录
		   SqlItem = "select NewsUrl,Title,CollecDate,Result From KS_History"
		   Set RsItem = Server.CreateObject("adodb.recordset")
		   RsItem.Open SqlItem, ConnItem, 1, 1
		   If Not RsItem.EOF Then
			  Arr_Historys = RsItem.GetRows()
		   End If
		   RsItem.Close
		   Set RsItem = Nothing
		
		   myCache.name = CacheTemp & "Historys"
		   Call myCache.clean
		   If IsArray(Arr_Historys) = True Then
			  myCache.add Arr_Historys, DateAdd("n", 1000, Now)
		   End If
		
		   '其它信息
		   myCache.name = CacheTemp & "collectest"
		   Call myCache.clean
		   myCache.add CollecTest, DateAdd("n", 1000, Now)
		
		   myCache.name = CacheTemp & "contentview"
		   Call myCache.clean
		   myCache.add Content_View, DateAdd("n", 1000, Now)
		
		   Set myCache = Nothing
		End Sub
		Sub DelNews()
		   ConnItem.Execute ("Delete From KS_NewsList")
		End Sub
End Class
%>
 
