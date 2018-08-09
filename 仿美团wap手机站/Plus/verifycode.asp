<%
' 仿discuz论坛验证码
' version 1.0
' http://www.mysuc.com/
' Copyright (C) 2009 by hayden
Option Explicit
Response.buffer=true
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Addheader   "cache-control","no-cache"   
Response.AddHeader   "Pragma","no-cache"
Response.ContentType = "Image/BMP"
'If Trim(Request.ServerVariables("HTTP_REFERER"))="" Then response.End()

	' 明确变量申明
	' 此处设置置会话超时为3分钟
	Session.Timeout = 3

	' 随机运行，以确保一个随机号码检索
	Randomize	
	
	' 声明变量
	Dim jpeg
	Dim pixelsAcross,textColour,CodeTotal
	Dim sessionnaem 
	Dim randomNumber : randomNumber = Int(Rnd * 7)+1
	
	' 验证码全局变量名称设定
	sessionnaem = "Verifycode"
	' 验证码个数
	CodeTotal = 4
	' 文字颜色
	textColour =  randomFomtcolor(randomNumber)
	
	' 验证码缩进距离
	pixelsAcross = Int(Rnd * 20)+3
	
	' 创建一个jpeg对象
	on error resume next
	Set jpeg = Server.CreateObject("Persits.jpeg")
	if err then  '不支持aspjpeg则用原来的
	  err.clear
	  Com_CreatValidCode
	  response.end
	end if
	
	' 打开随机背景图片
	drawBackGroud randomNumber,170,60
	' 绘制字符
	doString
	' 随机线
	drawLines
	' 随机圆
	drawCircle
	' 随机矩形
	'drawBar

	' 返回的二进制，申明本页为一个JPEG图片类型
	jpeg.SendBinary
	Set jpeg = Nothing 

	' 函数（drawBackGroud）：打开背景图片
	Function drawBackGroud(srandom,swidth,sheight)
		Jpeg.Open Server.MapPath("background/background"&srandom&".jpg")
		Jpeg.Width = swidth
		Jpeg.Height = sheight	
	End Function
	
	' 函数（drawLines）：绘制随机线
	Sub drawLines
		jpeg.Canvas.Pen.Color = &HADCD3C
		jpeg.Canvas.DrawLine 0, Int(Rnd * jpeg.Height), jpeg.Width, Int(Rnd * jpeg.Height)
	End Sub
	
	' 函数（drawBar）：绘制随机矩形框
	Sub drawBar
		jpeg.Canvas.Brush.Solid = False '填充
		'矩形边框颜色
		jpeg.Canvas.Pen.Color = &H9CCF00
		'绘制矩形框
		jpeg.Canvas.Bar Int(Rnd * jpeg.Width), Int(Rnd * jpeg.Height), Int(Rnd * 50)+20,Int(Rnd * 50)+20
	End Sub

	' 函数（drawCircle）：绘制随机圆
	Sub drawCircle
		jpeg.Canvas.Brush.Solid = False '填充
		jpeg.Canvas.Pen.Color = &H8080FF
		jpeg.Canvas.Circle Int(Rnd * jpeg.Width), Int(Rnd * jpeg.Height), Int(Rnd * 10)+5
		jpeg.Canvas.Pen.Color = &HEEEEEE
		jpeg.Canvas.Circle Int(Rnd * jpeg.Width), Int(Rnd * jpeg.Height), Int(Rnd * 10)+10
	End Sub

	' 函数（doString）：绘制验证码字符
	Sub doString
		Dim theString
		Dim x
	
		' 获取坠机字符串
		theString = createRandomString()
		
		' 循环通过字符串的每个字符
		For x = 1 to len(theString)

			' 在验证码图片当前位置打印字符
			addLetter Mid(theString, x, 1)
			
		Next

	End Sub

	' 函数（addLetter）在验证码图片当前位置打印字符
	Sub addLetter(theLetter)	
		' 字体的颜色
		jpeg.Canvas.Font.Color = textColour
		' 字体阴影
		jpeg.Canvas.Font.ShadowColor = &HFFFFFF
		' 是否为粗体　故不做随机判断，而是直接设定加粗
		'if doTextStyle then
			jpeg.Canvas.Font.Bold = True
		'End If
		if doTextStyle then  '下划线
			'jpeg.Canvas.Font.Underlined  = True
		End If	
		' 是否为斜体
		if doTextStyle then
			jpeg.Canvas.Font.Italic   = True
		End If		
		' 字体
		jpeg.Canvas.Font.Family = "Arial Black"'randomFont()		
		' 字体大小
		jpeg.Canvas.Font.Size = randomFontSize()
		
		' 文字清晰度
		jpeg.Canvas.Font.Quality = 8
		
		' 背景色　当前使用了背景图，故此处注释掉
		'jpeg.Canvas.Font.BkColor = backColour
		
		' 字体背景模式(处理平滑)
		jpeg.Canvas.Font.BkMode = "transparent"
		' 绘制字符
		jpeg.canvas.print pixelsAcross, Int(Rnd * 5), theLetter
		' 字符宽度
		pixelsAcross = pixelsAcross + Int(Rnd * 10)+30
	End Sub
	
	' 返回随机真假值各机率为50%
	Function doTextStyle()
		if Rnd() > 0.5 then
			doTextStyle = true
		else
			doTextStyle = false
		end if
	End Function

	' 返回验证码中各字符的随机大小
	Function randomFontSize()
		Dim theNumber
		' 获取一个随机大小，范围(40-60)
		theNumber = Int(Rnd * 20) + 40
		randomFontSize = theNumber
		
	End Function

	' 返回随机验证码文字颜色
	Function randomFomtcolor(srandomm)
		Dim arrFomtcolor(8)
		arrFomtcolor(1) = &HBDE3FF
		arrFomtcolor(2) = &HD68618
		arrFomtcolor(3) = &H086529
		arrFomtcolor(4) = &H637594
		arrFomtcolor(5) = &Hffffff
		arrFomtcolor(6) = &HBDDBF7
		arrFomtcolor(7) = &H08756B
		arrFomtcolor(8) = &H295131
		randomFomtcolor = arrFomtcolor(srandomm)
	End Function 
	
	' 返回随机字体
	Function randomFont()
		Dim theNumber	
		Dim font	
		' 取得1-6区间内一随机字符
		theNumber = Int(Rnd * 5) + 1
		' 随机字体
		if theNumber =1 then
			font = "Arial Black"
		elseif theNumber =2 then
			font = "Courier New"
		elseif theNumber =3 then
			font = "Helvetica"
		elseif theNumber =4 then
			font = "Times New Roman"
		elseif theNumber =5 then
			font = "Verdana"
		else
			font = "Geneva"
		end If
		randomFont = font
	
	End Function
	
	' 返回随机验证证字符串
	Function createRandomString
		Dim outputString
		Dim x
        For x = 0 To CodeTotal-1
			' 英文字符出现机率60%, 数字出现机率40%
			if rnd() < 0.6 then
				' 返回一个随机英文字符
            	outputString = outputString & Chr(Int((26 * rnd()) + 65))
			else
				' 返回一个随机数字
				outputString = outputString & Chr(Int((10 * rnd()) + 48))
			end if
        Next
		Session(sessionnaem) = outputString
        createRandomString = outputString	
	End Function
	
'自带验证码
Sub Com_CreatValidCode()
        Randomize
        Dim i, ii, iii
        Const cOdds = 5 ' 杂点出现的机率
        Const cAmount = 10 ' 文字数量
        Const cCode = "0123456789"
        
        ' 颜色的数据(字符，背景)
        Dim vColorData(1)
        vColorData(0) = ChrB(0) & ChrB(0) & ChrB(2)  ' 蓝0，绿0，红0（黑色）
        vColorData(1) = ChrB(255) & ChrB(255) & ChrB(255) ' 蓝250，绿236，红211（浅蓝色）
        
        ' 随机产生字符
        Dim vCode(4), vCodes
        For i = 0 To 3
          vCode(i) = Int(Rnd * cAmount)
          vCodes = vCodes & Mid(cCode, vCode(i) + 1, 1)
        Next
        Session("Verifycode") = vCodes  '记录入Session
        ' 字符的数据
        Dim vNumberData(9)
        vNumberData(0) = "1110000111110111101111011110111101001011110100101111010010111101001011110111101111011110111110000111"
        vNumberData(1) = "1111011111110001111111110111111111011111111101111111110111111111011111111101111111110111111100000111"
        vNumberData(2) = "1110000111110111101111011110111111111011111111011111111011111111011111111011111111011110111100000011"
        vNumberData(3) = "1110000111110111101111011110111111110111111100111111111101111111111011110111101111011110111110000111"
        vNumberData(4) = "1111101111111110111111110011111110101111110110111111011011111100000011111110111111111011111111000011"
        vNumberData(5) = "1100000011110111111111011111111101000111110011101111111110111111111011110111101111011110111110000111"
        vNumberData(6) = "1111000111111011101111011111111101111111110100011111001110111101111011110111101111011110111110000111"
        vNumberData(7) = "1100000011110111011111011101111111101111111110111111110111111111011111111101111111110111111111011111"
        vNumberData(8) = "1110000111110111101111011110111101111011111000011111101101111101111011110111101111011110111110000111"
        vNumberData(9) = "1110001111110111011111011110111101111011110111001111100010111111111011111111101111011101111110001111"
        ' 输出图像文件头
        Response.BinaryWrite ChrB(66) & ChrB(77) & ChrB(230) & ChrB(4) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) &_
          ChrB(0) & ChrB(0) & ChrB(54) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(40) & ChrB(0) &_
          ChrB(0) & ChrB(0) & ChrB(40) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(10) & ChrB(0) &_
          ChrB(0) & ChrB(0) & ChrB(1) & ChrB(0)
        
        ' 输出图像信息头
        Response.BinaryWrite ChrB(24) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(176) & ChrB(4) &_
          ChrB(0) & ChrB(0) & ChrB(18) & ChrB(11) & ChrB(0) & ChrB(0) & ChrB(18) & ChrB(11) &_
          ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) &_
          ChrB(0) & ChrB(0)
        
        For i = 9 To 0 Step -1  ' 历经所有行
                For ii = 0 To 3  ' 历经所有字
                        For iii = 1 To 10 ' 历经所有像素
                                ' 逐行、逐字、逐像素地输出图像数据
                                If Rnd * 99 + 1 < cOdds Then ' 随机生成杂点
                                        If Mid(vNumberData(vCode(ii)), i * 10 + iii, 1) Then
                                                Response.BinaryWrite vColorData(0)
                                        Else
                                                Response.BinaryWrite vColorData(1)
                                        End If
                                Else
                                        Response.BinaryWrite vColorData(Mid(vNumberData(vCode(ii)), i * 10 + iii, 1))
                                End If
                        Next
                Next
        Next
End Sub

%>
