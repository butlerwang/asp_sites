<%

Class CtoECls
     Private Sub Class_Initialize()
	 End Sub
	 Private Sub Class_Terminate()
	 End Sub
	 function CTOE(str) 
	    dim codestr:codestr=Join(UTF8ToANSIArray(str), ",") 
		dim n,tt
		tt=split(codestr,",")
		for n=0 to ubound(tt)
		 CTOE=CTOE&getpychar(tt(n))
		next
     end function 
	  
		Public Function LShift(ByVal lValue, ByVal iBit) 
			LShift = lValue * (2 ^ iBit) 
		End Function 
		
		'整合高低位字节为整数 
		Public Function MAKEWORD(ByVal iHigh, ByVal iLow) 
			MAKEWORD = (iHigh And &HFF) Or LShift((iLow And &HFF), 8) 
		End Function 
		
		'将Unicode字符串转换成GBK编码数组 
		'该函数用于CodePage=65001环境 
		Public Function UTF8ToANSIArray(ByVal strData) 
			Dim objStream 
			Dim ret, i, k 
			Set objStream = Server.CreateObject("ADODB.Stream") 
			objStream.Type = 2 
			objStream.Mode = 3 
			objStream.Charset = "gbk" 
			objStream.Open 
			objStream.WriteText strData 
			objStream.Position = 0 
			objStream.Type = 1 
			ReDim ret(objStream.Size - 1) 
			For i = 0 To UBound(ret) 
				ret(i) = AscB(objStream.Read(1)) 
			Next 
			k = 0 
			For i = 0 To UBound(ret) 
				If ret(i) < 128 Then 
					ret(k) = ret(i) 
				Else 
					ret(k) = MAKEWORD(ret(i + 1), ret(i)) 
					i = i + 1 
				End If 
				k = k + 1 
			Next 
			ReDim Preserve ret(k - 1) 
			objStream.Close 
			Set objStream = Nothing 
			UTF8ToANSIArray = ret 
		End Function 
		
		function getpychar(tmp)
		'tmp=65536+asc(char)
		if(tmp>=45217 and tmp<=45252) then 
		getpychar= "a"
		elseif(tmp>=45253 and tmp<=45760) then
		getpychar= "b"
		elseif(tmp>=45761 and tmp<=46317) then
		getpychar= "c"
		elseif(tmp>=46318 and tmp<=46825) then
		getpychar= "d"
		elseif(tmp>=46826 and tmp<=47009) then 
		getpychar= "e"
		elseif(tmp>=47010 and tmp<=47296) then 
		getpychar= "f"
		elseif(tmp>=47297 and tmp<=47613) then 
		getpychar= "g"
		elseif(tmp>=47614 and tmp<=48118) then
		getpychar= "h"
		elseif(tmp>=48119 and tmp<=49061) then
		getpychar= "j"
		elseif(tmp>=49062 and tmp<=49323) then 
		getpychar= "k"
		elseif(tmp>=49324 and tmp<=49895) then 
		getpychar= "l"
		elseif(tmp>=49896 and tmp<=50370) then 
		getpychar= "m"
		elseif(tmp>=50371 and tmp<=50613) then 
		getpychar= "n"
		elseif(tmp>=50614 and tmp<=50621) then 
		getpychar= "o"
		elseif(tmp>=50622 and tmp<=50905) then
		getpychar= "p"
		elseif(tmp>=50906 and tmp<=51386) then 
		getpychar= "q"
		elseif(tmp>=51387 and tmp<=51445) then 
		getpychar= "r"
		elseif(tmp>=51446 and tmp<=52217) then 
		getpychar= "s"
		elseif(tmp>=52218 and tmp<=52697) then 
		getpychar= "t"
		elseif(tmp>=52698 and tmp<=52979) then 
		getpychar= "w"
		elseif(tmp>=52980 and tmp<=53688) then 
		getpychar= "x"
		elseif(tmp>=53689 and tmp<=54480) then 
		getpychar= "y"
		elseif(tmp>=54481 and tmp<=62289) then
		getpychar= "z"
		else
		getpychar=chr(tmp)
		end if
		end function
 
End Class
%>