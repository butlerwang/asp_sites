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
<title>���ݹ���������޹�˾</title>
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
set upfile=new upfile_class ''�����ϴ�����
upfile.AllowExt=rs(0)   '�����ϴ��ļ���������
'upfile.NoAllowExt="asp;exe;htm;html;aspx;cs;vb;js;css;"	'�����ϴ����͵ĺ�����
upfile.GetData (rs(1)*1024)   'ȡ���ϴ�����,��������ϴ�10M
if upfile.isErr then  '�������
    select case upfile.isErr
	case 1
	Response.Write "��û���ϴ�����"
	case 2
	Response.Write "���ϴ����ļ���������,���"&rs(1)&"K"
	end select
else
    FPath=Server.mappath("../../uploadfile")
  	for each formName in upfile.file '�г������ϴ��˵��ļ�
	   set oFile=upfile.file(formname)
	   FileName=oFile.filename
	  ' upfile.SaveToFile formname,FPath b&"\"&FileName   ''�����ļ� Ҳ����ʹ��AutoSave������,����һ��,���ǻ��Զ������µ��ļ���
       FileName=upfile.AutoSave(formname,FPath&"\"&FileName)
	   if upfile.iserr then 
		Response.Write upfile.errmessage
		else
		pass=true '�ϴ��ɹ�
		Form_Name=upfile.form("Form_Name")
		UploadFile=upfile.form("UploadFile")
		end if
	 set oFile=nothing
	next
end if
set upfile=nothing  'ɾ���˶���



if pass then

'================ ���� ===========

if UploadFile ="Bpic" Then

   FileExt=LCase(Mid(FileName,InStrRev(FileName, ".")+1))
   If inStr("gif|jpg|jpeg|bmp",FileExt)>0 Then
      '��������ͼ
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
		  Jpeg.Save Server.MapPath("../../uploadfile/"&Spic) '���� 
          response.Write("<b><font color=#009900>��</font></b> ����ͼ���ɳɹ���<script language=""JavaScript"">parent."&Form_Name&".Spic.value="""&Spic&""";</script>")
		  		  
	  End If
   
      '��ˮӡ,������ǰ��
	 'If Rs(7) Then
	      'Set Jpeg = Server.CreateObject("Persits.Jpeg") 
          'Jpeg.Open Server.MapPath("../../uploadfile/"&FileName)
          'Jpeg.Canvas.Font.Color = &H999999'  ��ɫ 
          'Jpeg.Canvas.Font.Family = "����" '���� 
          'Jpeg.Canvas.Font.Bold = false  '�Ƿ�Ӵ�
         ' Jpeg.Canvas.Font.Size = rs(8)
		 '������ǰ��
          'Jpeg.Canvas.Font.BkMode="Opaque" '��ɫ����
		  
         ' Jpeg.Canvas.Pen.Color = &H000000' black ��ɫ 
         ' Jpeg.Canvas.Pen.Width = 1 '���ʿ��
          'Jpeg.Canvas.Brush.Solid = False '�Ƿ�Ӵִ��� 
         ' Jpeg.Canvas.Bar 1, 1, Jpeg.Width, Jpeg.Height 
		           
		  'Jpeg.Width=Jpeg.OriginalWidth
         ' Jpeg.Height=Jpeg.OriginalHeight
		   '������ǰ��
		  'Jpeg.Canvas.Print 2, 2, "  "&rs(9)&"  "
          'Jpeg.Save Server.MapPath("../../uploadfile/"&FileName) '���� 		  
	 ' End If  
	  '������ǰ��
   End If
End If
'=================================

  response.Write("<b><font color=#009900>��</font></b> �ļ��ϴ��ɹ���<script language=""JavaScript"">parent."&Form_Name&"."&UploadFile&".value="""&FileName&""";</script>")
else
  response.Write("<b>[<a href='javascript:window.history.back(-1)'>����</a>]</b>")
end if


Public Function GetNewFileName()
    Randomize
	dim ranNum
	dim dtNow
	dtNow=Now()
	ranNum=int(90000*rnd)+10000
	'���������webboy�ṩ
	GetNewFileName=year(dtNow) & right("0" & month(dtNow),2) & right("0" & day(dtNow),2) & right("0" & hour(dtNow),2) & right("0" & minute(dtNow),2) & right("0" & second(dtNow),2) & ranNum
End Function


%>
	</td>
  </tr>
</table>
</body>
</html>

