<%
Dim Immediate,Templates
Immediate = true
Sub Echo(sStr)
    If Immediate Then
        Response.Write    sStr
    Else
        Templates    = Templates&sStr 
    End If 
End Sub 

Sub Scan(sTemplate)
    Dim iPosLast, iPosCur
    iPosLast    = 1
    While True 
        iPosCur    = InStr(iPosLast, sTemplate, "{@")
        If iPosCur>0 Then
            Echo    Mid(sTemplate, iPosLast, iPosCur-iPosLast)
            iPosLast    = Parse(sTemplate, iPosCur+2)
        Else 
            Echo    Mid(sTemplate, iPosLast)
            Exit Sub  
        End If 
   Wend 
End Sub 

Function Parse(sTemplate, iPosBegin)
    Dim iPosCur, sToken, sValue, sTemp
    iPosCur        = InStr(iPosBegin, sTemplate, "}")
    sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
    iPosBegin    = iPosCur+1
    iPosCur       = InStr(sTemp, ".")
	if iPosCur>1 Then
    sToken        = Left(sTemp, iPosCur-1)
	End If
    sValue        = Mid(sTemp, iPosCur+1) 

    Select Case sValue
        Case "begin"
            sTemp            = "{@" & ( sToken & ".end}" )
            iPosCur            = InStr(iPosBegin, sTemplate, sTemp)
            KSCls.ParseArea      sToken, Mid(sTemplate, iPosBegin, iPosCur-iPosBegin)
            iPosBegin        = iPosCur+Len(sTemp)
        Case Else
            KSCls.ParseNode sToken, sValue 
    End Select 
    Parse    = iPosBegin
End Function 
%>
