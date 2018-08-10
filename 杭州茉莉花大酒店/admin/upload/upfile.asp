<!--#include file="upfile_class.asp"-->
<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="lib.asp"-->
<%call OpenData()
  Set Rs=Server.CreateObject("adodb.recordset")
  Sql="Select UpfileType,UpfileSize,PicAuto,PicAutoType,PicPercent,PicHeight,PicWidth,Watermark,WatermarkSize,WatermarkWord From Sbe_WebConfig"
  Rs.Open Sql,Conn,1,1%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>杭州古秀服饰有限公司</title>
<STYLE TYPE="text/css">
<!--
td {font-size: 9pt}
.bt {border-left:1px solid #C0C0C0; border-top:1px solid #C0C0C0; font-size: 9pt; border-right-width: 1; border-bottom-width: 1; height: 20px; width: 80px; background-color: #EEEEEE; cursor: hand; border-right-style:solid; border-bottom-style:solid}
.tx1 { width: 200 ;height: 20px; font-size: 9pt; border: 1px solid; border-color: black black #000000; color: #0000FF}
-->
</STYLE>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#F5F8FA">
<table width="353" height="13" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="353" height="20">&nbsp; 
      <%
dim pass
pass=false
set upfile=new upfile_class ''建立上传对象
upfile.AllowExt=rs(0)   '设置上传文件类型名单
'upfile.NoAllowExt="asp;exe;htm;html;aspx;cs;vb;js;css;"	'设置上传类型的黑名单
upfile.GetData (rs(1)*1024)   '取得上传数据,限制最大上传10M
if upfile.isErr then  '如果出错
    select case upfile.isErr
	case 1
	Response.Write "你没有上传数据"
	case 2
	Response.Write "你上传的文件超出限制,最大"&rs(1)&"K"
	end select
else
    FPath=Server.mappath("../../uploadfile")
  	for each formName in upfile.file '列出所有上传了的文件
	   set oFile=upfile.file(formname)
	   FileName=oFile.filename
	  ' upfile.SaveToFile formname,FPath b&"\"&FileName   ''保存文件 也可以使用AutoSave来保存,参数一样,但是会自动建立新的文件名
       FileName=upfile.AutoSave(formname,FPath&"\"&FileName)
	   if upfile.iserr then 
		Response.Write upfile.errmessage
		else
		pass=true '上传成功
		Form_Name=upfile.form("Form_Name")
		UploadFile=upfile.form("UploadFile")
		end if
	 set oFile=nothing
	next
end if
set upfile=nothing  '删除此对象



if pass then

'================ 设置 ===========

if UploadFile ="Bpic" Then

   FileExt=LCase(Mid(FileName,InStrRev(FileName, ".")+1))
   If inStr("gif|jpg|jpeg|bmp",FileExt)>0 Then
      '生成略缩图
      IF rs(2) Then
	      Set Jpeg = Server.CreateObject("Persits.Jpeg") 
          Jpeg.Open Server.MapPath("../../uploadfile/"&FileName)
		  if rs(3)=1 Then
		     Width=Jpeg.OriginalWidth * rs(4)/100
			 Height=Jpeg.OriginalHeight * rs(4)/100
		  else
		     if rs(5)=0 and rs(6)=0 then
			   Width=Jpeg.OriginalWidth
			   Height=Jpeg.OriginalHeight
			 elseif rs(5)=0 then
			    Width=Jpeg.OriginalWidth * rs(6) / Jpeg.OriginalHeight
				Height=rs(6)
			 elseif rs(6)=0 then
			    Width=rs(5)
				Height=rs(5) * Jpeg.OriginalHeight / Jpeg.OriginalWidth
			 else
			    Width=rs(5)
				Height=rs(6)
			 end if		  
		  end if
		  Jpeg.Width=Width
          Jpeg.Height=Height
		  Spic=GetNewFileName()&"."&FileExt
		  Jpeg.Save Server.MapPath("../../uploadfile/"&Spic) '保存 
          response.Write("<b><font color=#009900>√</font></b> 略缩图生成成功！<script language=""JavaScript"">parent."&Form_Name&".Spic.value="""&Spic&""";</script>")
		  		  
	  End If
   
      '加水印,以下以前有
	 'If Rs(7) Then
	      'Set Jpeg = Server.CreateObject("Persits.Jpeg") 
          'Jpeg.Open Server.MapPath("../../uploadfile/"&FileName)
          'Jpeg.Canvas.Font.Color = &H999999'  颜色 
          'Jpeg.Canvas.Font.Family = "黑体" '字体 
          'Jpeg.Canvas.Font.Bold = false  '是否加粗
         ' Jpeg.Canvas.Font.Size = rs(8)
		 '以上以前有
          'Jpeg.Canvas.Font.BkMode="Opaque" '白色背景
		  
         ' Jpeg.Canvas.Pen.Color = &H000000' black 颜色 
         ' Jpeg.Canvas.Pen.Width = 1 '画笔宽度
          'Jpeg.Canvas.Brush.Solid = False '是否加粗处理 
         ' Jpeg.Canvas.Bar 1, 1, Jpeg.Width, Jpeg.Height 
		           
		  'Jpeg.Width=Jpeg.OriginalWidth
         ' Jpeg.Height=Jpeg.OriginalHeight
		   '以下以前有
		  'Jpeg.Canvas.Print 2, 2, "  "&rs(9)&"  "
          'Jpeg.Save Server.MapPath("../../uploadfile/"&FileName) '保存 		  
	 ' End If  
	  '以上以前有
   End If
End If
'=================================

  response.Write("<b><font color=#009900>√</font></b> 文件上传成功！<script language=""JavaScript"">parent."&Form_Name&"."&UploadFile&".value="""&FileName&""";</script>")
else
  response.Write("<b>[<a href='javascript:window.history.back(-1)'>返回</a>]</b>")
end if


Public Function GetNewFileName()
    Randomize
	dim ranNum
	dim dtNow
	dtNow=Now()
	ranNum=int(90000*rnd)+10000
	'以下这段由webboy提供
	GetNewFileName=year(dtNow) & right("0" & month(dtNow),2) & right("0" & day(dtNow),2) & right("0" & hour(dtNow),2) & right("0" & minute(dtNow),2) & right("0" & second(dtNow),2) & ranNum
End Function


%>
	</td>
  </tr>
</table>
</body>
</html>

