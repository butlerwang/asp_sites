<%
Response.CodePage=65001
Response.Charset="utf-8"
Dim SqlNowString,DataPart_D,DataPart_Y,DataPart_H,DataPart_S,DataPart_W,DataPart_M,CurrentPage
Dim Conn,DBPath,CollectDBPath,DataServer,DataUser,DataBaseName,DataBasePsw,ConnStr,CollcetConnStr
Const DataBaseType=0                 '系统数据库类型，"1"为MS SQL2000数据库，"0"为MS ACCESS 2000数据库
Const MsxmlVersion=".3.0"                '系统采用XML版本设置 

Const EnableSiteManageCode = True        '是否启用后台管理认证密码 是： True  否： False 
Const SiteManageCode = "8888"      '后台管理认证密码，请修改，这样即使有人知道了您的后台用户名和密码也不能登录后台
Const IsBusiness=False              

 
If DataBaseType=0 then
	'如果是ACCESS数据库，请认真修改好下面的数据库的文件名
	DBPath       = "/KS_Data/XunWang.mdb"      'ACCESS数据库的文件名，请使用相对于网站根目录的的绝对路径
Else
	 '如果是SQL数据库，请认真修改好以下数据库选项
	 DataServer   = "(local)"                                  '数据库服务器IP
	 DataUser     = "sa"                                       '访问数据库用户名
	 DataBaseName = "XunWang"                                '数据库名称
	 DataBasePsw  = "989066"                                   '访问数据库密码 
End if

'采集数据库路径
CollectDBPath="\KS_Data\Collect\KS_Collect.Mdb"

'=============================================================== 以下代码请不要自行修改========================================
CurrentPage=Request("Page")
If Not IsNumeric(CurrentPage) Then CurrentPage=1
If CurrentPage<1 Then CurrentPage=1
Call OpenConn
Sub OpenConn()
    On Error Resume Next
    If DataBaseType = 1 Then
       ConnStr="Provider = Sqloledb; User ID = " & datauser & "; Password = " & databasepsw & "; Initial Catalog = " & databasename & "; Data Source = " & dataserver & ";"
	   SqlNowString = "getdate()"
	   DataPart_D   = "d"
	   DataPart_Y   = "year"
	   DataPart_H   = "hour"
	   DataPart_S   = "s"
	   DataPart_W   = "week"
       DataPart_M   = "month"
    Else
       ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DBPath)
	   SqlNowString = "Now()"
	   DataPart_D   = "'d'"
	   DataPart_Y   = "'yyyy'"
	   DataPart_H   = "'h'"
	   DataPart_S   = "'s'"
	   DataPart_W   = "'w'"
       DataPart_M   = "'m'"
    End If
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.open ConnStr
    If Err Then Err.Clear:Set conn = Nothing:Response.Write "数据库连接出错，请检查Conn.asp文件中的数据库参数设置。出错原因:<br/>" & Err.Description:Response.End
	CollcetConnStr ="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(CollectDBPath)
End Sub
Sub CloseConn()
    On Error Resume Next
	Conn.close:Set Conn=nothing
End sub

'====================================如果频道启用二级域名,请正确配置以下参数,否则可能导致会员不能登录==========================
Const EnabledSubDomain =false       rem 网站频道是否启用二级域名 true表示启用 false表示没有启用
Const RootDomain = ""       rem 网站主域名根,如果有多个子域名,必须设置
'=============================================二级域名配置结束========================================================


'==============================================全局变量类开始==============================
Dim GCls:Set GCls=New GlobalVarCls
Class GlobalVarCls
    Public Sql_Use
    Public StaticPreList,StaticPreContent,StaticExtension,ClubPreContent,ClubPreList
	Private Sub Class_Initialize()
	   StaticPreList    = "list"                 rem 内容模型伪静态列表前缀 不能包含"?"及"-"
	   staticPreContent = "thread"               rem 内容模型伪静态内容前缀 
	   StaticExtension  = ".html"                rem 内容模型伪静态扩展名
	   ClubPreContent   = "forumthread"          rem 伪静态小论坛帖子前缀地址 
	   ClubPreList      = "forum"                rem 伪静态小论坛版面列表前缀地址
	End Sub
    Private Sub Class_Terminate()
		 Set GCls=Nothing
	End Sub
	
	Public Function Execute(Command)
		If Not IsObject(Conn) Then OpenConn()
		On Error Resume Next
		Set Execute = Conn.Execute(Command)
		If Err Then
				Response.Write("查询语句为：" & Command & "<br>")
				Response.Write("错误信息为：" & Err.Description & "<br>")
			Err.Clear
			Set Execute = Nothing
			Response.End()
		End If
		Sql_Use = Sql_Use + 1
	End Function
	
	Function GetUrl() 
		On Error Resume Next 
		Dim strTemp 
		If LCase(Request.ServerVariables("HTTPS")) = "off" Then 
		 strTemp = "http://"
		Else 
		 strTemp = "https://"
		End If 
		strTemp = strTemp & Request.ServerVariables("SERVER_NAME") 
		If Request.ServerVariables("SERVER_PORT") <> 80 Then 
		 strTemp = strTemp & ":" & Request.ServerVariables("SERVER_PORT") 
		end if
		strTemp = strTemp & Request.ServerVariables("URL") 
		If Trim(Request.QueryString) <> "" Then 
		 strTemp = strTemp & "?" & Trim(Request.QueryString) 
		end if
		GetUrl = strTemp 
	End Function

	'====================标志来访地址================
	Public Property Let ComeUrl(ByVal strVar) 
			Session("M_ComeUrl") = strVar 
	End Property 
			
	Public Property Get ComeUrl
			ComeUrl= Session("M_ComeUrl")
	End Property 
	'================================================
End Class
'==============================================全局临时变量类结束==============================
%>
