<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%

Dim KS:Set KS=New PublicCls

	 '==========================推广积分===============================================
		  If KS.Setting(140)="1" Then
		   	Dim ComeUrl:ComeUrl=Request.ServerVariables("HTTP_REFERER")
			Dim QParam:QParam=Split(Lcase(ComeUrl),"uid=")
			If Ubound(QParam)>=1 Then
		    Dim UserName:UserName=Split(QParam(1),"&")(0)
			End If
			If UserName<>"" Then
			  If Not Conn.Execute("Select Top 1 UserName From KS_User Where UserName='" & UserName & "'").Eof Then
			    Dim UserIP:UserIP=KS.GetIP()
				If ComeUrl="" Then ComeUrl="★直接输入或书签导入★"
			    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
				Dim SQL:SQL="Select top 1 * From KS_PromotedPlan Where UserName='" & UserName & "' And UserIP='" & UserIP & "' And DateDiff(" & DataPart_D & ",AddDate," & SqlNowString & ")<1"
				RS.Open SQL ,conn,1,3
				If RS.Eof And RS.Bof Then
				  RS.AddNew
				  RS("UserName") = UserName
				  RS("UserIP")   = UserIP
				  RS("AddDate")  = Now
				  RS("ComeUrl")  = KS.URLDecode(ComeUrl)
				  RS("Score")    = KS.Setting(141)
				  RS("AllianceUser")="-"
				  RS.Update
				  RS.Close

				  Call KS.ScoreInOrOut(UserName,1,KS.Setting(141),"系统","成功推荐一个IP:" & UserIP & "访问!",0,0)
				Else 
				  RS.Close
				End IF
				Set RS=Nothing
			  End If
			End If
		  End If
 '==========================推广积分结束========================================
Set KS=Nothing
Call CloseConn()
%>
