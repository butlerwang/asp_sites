<%

Class LoginCheckCls1
		Private ComeUrl
		Private TrueSiteUrl
		Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		Sub Run()
		ComeUrl = Trim(Request.ServerVariables("HTTP_REFERER"))
		TrueSiteUrl = Trim(Request.ServerVariables("HTTP_HOST"))
		  If KS.IsNul(KS.C("AdminName")) Or KS.IsNul(KS.C("AdminPass")) Or KS.IsNul(KS.C("PowerList"))="" Or KS.IsNUL(KS.C("UserName")) Then
			Response.Write ("<script>top.location.href='/';</script>")
            Response.End()
		  Else
			 If Check=false Then
				 Response.Write ("<script>top.location.href='/';</script>")
				 Response.End()
			 End If
		 End If
		End Sub
		Function Check()
		   if KS.IsNul(KS.C("UserName")) Or  KS.IsNul(KS.C("PassWord")) Or KS.IsNul(KS.C("AdminName")) Or KS.IsNul(KS.C("AdminPass")) Then Check=false :Exit Function
		  Dim ChkRS:Set ChkRS = Server.CreateObject("ADODB.RecordSet")
			 ChkRS.Open "Select top 1 * From KS_Admin Where UserName='" & KS.R(KS.C("AdminName")) & "'",Conn, 1, 1
			 If ChkRS.EOF And ChkRS.BOF Then
			   Check=false
			 Else
			   If ChkRS("PassWord")=KS.C("AdminPass") Then
			   Check=true
			   Else
			    Check=false
			   End If
			 End If
		     ChkRS.Close:Set ChkRS = Nothing
		End Function
End Class
%> 
