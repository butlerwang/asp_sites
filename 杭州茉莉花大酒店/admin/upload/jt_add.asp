<!--#include file="upload_wj.inc"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<%
set upload=new lx_add
if upload.form("act")="uploadfile" then
	filepath=trim(upload.form("filepath"))
	filelx=trim(upload.form("filelx"))
	
	i=0
	for each formName in upload.File
		set file=upload.File(formName)
 
 fileExt=lcase(file.FileExt)	'�õ����ļ���չ��������.
 if file.filesize<100 then
	response.write "<script language=javascript>alert('����ѡ����Ҫ�ϴ����ļ���');history.go(-1);</script>"
	response.end
 end if
 if (fileExt="swf") then
 checkup=1
 else
 if (fileExt="jpg") then
 checkup=1
 else
 if (fileExt="gif") then
 checkup=1
 else
	response.write "<script language=javascript>alert('���ļ����Ͳ����ϴ���');history.go(-1);</script>"
	response.end
 end if
 end if
 end if
 
 if checkup<>1 then
 	response.write "<script language=javascript>alert('���ļ����Ͳ����ϴ���');history.go(-1);</script>"
	response.end
 end if
 
 if file.filesize>(500*1024) then
		response.write "<script language=javascript>alert('ͼƬ�ļ���С���ܳ���500K��');history.go(-1);</script>"
		response.end
 end if

 randomize
 ranNum=int(90000*rnd)+10000
 filename=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
%>
<%
 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
  'file.SaveAs Server.mappath(filename)   ''�����ļ�
  file.SaveToFile Server.mappath(FileName)
  'response.write file.FileName&"�����ϴ��ɹ�!����<br>"
  'response.write "���ļ�����"&FileName&"<br>"
  'response.write "���ļ����Ѹ��Ƶ������λ�ã��ɹرմ��ڣ�"
  if filelx="swf" then
  response.write "<script>window.opener.document."&upload.form("FormName")&".size.value='"&int(file.FileSize/1024)&" K'</script>"
  end if
  response.write "<script>window.opener.document."&upload.form("FormName")&"."&upload.form("EditName")&".value='"&FileName&"'</script>"
  response.Write "<script>window.opener.dochangepic();</script>"
  'response.Write "<script>window.opener.mmm1.s</script>"
  'response.write "<script>window.opener.document."&upload.form("FormName")&"."&upload.form("image")&".src='"&FileName&"'</script>"
%>
<%
end if
set file=nothing
next
set upload=nothing
end if

%>
<script language="javascript">
window.alert("�ļ��ϴ��ɹ�!�벻Ҫ�޸����ɵ����ӵ�ַ��");
window.close();
</script>
