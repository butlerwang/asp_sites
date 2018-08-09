<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->

<% 
Response.Expires = 0
dim KS:Set KS=New PublicCls
dim del:del=KS.ChkClng(Request("del"))
Dim KSUser:Set KSUser=New UserCls
if KSUser.UserLoginChecked=false then
 set ks=nothing : set ksuser=nothing
 ks.die "error login"
end if


Dim Pic:Pic = Request("p")
If KS.IsNul(Pic) Then
 KS.Die "<script>alert('您没有上传图片');window.close();</script>"
ElseIf instr(lcase(pic),".gif")=0 and instr(lcase(pic),".jpg")=0 and instr(lcase(pic),".png")=0 and instr(lcase(pic),".jpeg")=0 Then
 KS.Die "<script>alert('非图片文件!');window.close();</script>"
ElseIf left(lcase(pic),4)="http" and instr(lcase(pic),lcase(ks.getdomain))=0 Then
 KS.Die "<script>alert('非本站图片不能处理!');window.close();</script>"
End If
Dim PointX:PointX = KS.ChkClng(KS.S("x"))
Dim PointY:PointY = KS.ChkClng(KS.S("y"))
Dim CutWidth:CutWidth = KS.ChkClng(KS.S("w"))
Dim CutHeight:CutHeight = KS.ChkClng(KS.S("h"))
Dim PicWidth:PicWidth = KS.ChkClng(KS.S("pw"))
Dim PicHeight:PicHeight = KS.ChkClng(KS.S("ph"))

on error resume next
Set Jpeg = Server.CreateObject("Persits.Jpeg")
if err then 
 err.clear
 KS.Die "<script>alert('服务器不支持aspJpeg组件!');</script>"
end if
Jpeg.Open Server.MapPath(Pic)

'缩放切割图片
Jpeg.Width = PicWidth
Jpeg.Height = PicHeight
Jpeg.Crop PointX, PointY, CutWidth + PointX, CutHeight + PointY

Dim filename:filename=split(pic,"/")(ubound(split(pic,"/")))
filename=split(filename,".")(0)&"_S."&split(filename,".")(1)


Dim SaveName

If KS.IsNul(KS.C("AdminName")) Then
SaveName=KS.ReturnChannelUserUpFilesDir(0,KSUser.GetUserInfo("UserID")) & filename
Else
SaveName=KS.GetUpFilesDir() & "/" &  filename
End If

Jpeg.Quality=80
Jpeg.Save Server.MapPath(SaveName)        '保存图片到磁盘

if del="1" and KS.C("AdminName")<>"" then
  call KS.DeleteFile(Pic)
end if
 
'输出图片
'Response.ContentType = "image/jpeg"
'Jpeg.SendBinary

If KS.Setting(97)="1" Then
  If Left(SaveName,1)="/" Then SaveName=Right(SaveName,Len(SaveName)-1) 
  If left(lcase(SaveName),4)<>"http" Then
   SaveName=KS.GetDomain & SaveName
  End If
End If


Set KS=Nothing
CloseConn
%>
<script type="text/javascript">
if (document.all){
window.returnValue='<%=SaveName%>';
}else{
top.window.opener.setVal('<%=SaveName%>');
}
top.close();
window.onunload=CheckReturnValue;
function CheckReturnValue()
{
    if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>