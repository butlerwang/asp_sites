<%
	 Public Function RexHtml_IF(ByVal str)
        Dim Reg
        Set Reg = New RegExp
        Reg.IgnoreCase = True
        Reg.Global = True
       ' On Error Resume Next
        Dim Matches, Match
        Dim TestIF,KS
		Set KS=New PublicCls
        
        Reg.Pattern = "{ElseIf([\s\S]*?):(.+?)}([\s\S]*?){Else\1}([\s\S]*?){/ElseIf\1}"
        Set Matches = Reg.Execute(str)
        For Each Match In Matches
            Execute ("If " & Match.SubMatches(1) & " Then TestIf = True Else TestIf = False")
            If TestIF Then str = Replace(str, Match.Value, Match.SubMatches(2)) Else str = Replace(str, Match.Value, Match.SubMatches(3)) ' 替换
            If Err Then Response.Write "<font color=red>执行Else" & Match.SubMatches(0) & "标签失败 [" & Match.SubMatches(1) & "]" & Err.Description & "</font>": Err.Clear: Response.End
        Next
        
        Reg.Pattern = "{If([\s\S]*?):(.+?)}([\s\S]*?){/If\1}"
        Set Matches = Reg.Execute(str)
        For Each Match In Matches
            Execute ("If " & Match.SubMatches(1) & " Then TestIf = True Else TestIf = False")
            If TestIF Then str = Replace(str, Match.Value, Match.SubMatches(2)) Else str = Replace(str, Match.Value, "") ' 替换
            If Err Then Response.Write "<font color=red>执行IF" & Match.SubMatches(0) & "标签失败 [" & Match.SubMatches(1) & "]" & Err.Description & "</font>": Err.Clear: Response.End
        Next
        
     Set Matches = Nothing
     Set Reg = Nothing
    
	 If RegExists("{ElseIf([\s\S]*?):(.+?)}([\s\S]*?){Else\1}([\s\S]*?){/ElseIf\1}", str) Or RegExists("{If([\s\S]*?):(.+?)}([\s\S]*?){/If\1}", str) Then str = RexHtml_IF(str) '再次替换
    
     RexHtml_IF = str
    End Function
	
	Public Function RegExists(ByVal Pattern, ByVal TestContent)
		Dim Reg
		Set Reg = New RegExp
        Reg.IgnoreCase = True
        Reg.Global = True
		Reg.Pattern = Pattern
		RegExists = Reg.Test(TestContent)
	End Function
	

%>