<%@language=vbscript codepage="65001" %>
<%
Option Explicit
Response.Buffer = True
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!-- #include File="../../KS_Cls/Kesion.CommonCls.asp" -->
<!-- #include File="../../conn.asp" -->
<!-- #include File="../../Plus/Session.asp" -->
<%
Dim Login:Set Login=New LoginCheckCls1
Call Login.Run()
Set Login=Nothing
Dim KS:Set KS=New PublicCls
Dim ID, sql, rs

ID = KS.R(KS.S("id"))
Call Main

Call CloseConn
Set KS=Nothing
Sub Main()
    Set rs=Conn.Execute("select * from KS_Label where ID='" & ID & "'") 
    If rs.bof and rs.EOF Then
        response.write "标签不存在"
    Else
%>
<html>
<head>
<title>自定义函数标签参数输入框</title>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<script src="../../ks_Inc/common.js" language="javascript"></script>
<script src="../../ks_Inc/jQuery.js" language="javascript"></script>
<script language="javascript">
function objectTag(itotal) {
        var TempStr="";
        for(i=0;i<itotal;i++){
		    if ($('#Field'+ i).val()=='')
			 {
			 alert('请输入'+$('#Param'+i).html());
			 $('#Field'+i).focus();
			 return false;
			 }
            if(i<itotal-1){
                TempStr =TempStr + $('#Field'+ i).val() + ","; 
            }else{
                TempStr=TempStr + $('#Field'+ i).val(); 
            }
        }
	    var reval = '<%=Replace(rs("LabelName"),"}","") %>('+TempStr+')}';  
	    window.returnValue = reval;
	    window.close();
}
</script>
<link href='Editor.css' rel='stylesheet' type='text/css'>
</head>
<body>
<form name="myform">
<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>
  <tr>
    <td colspan="2" align="center"><strong>请输入动态函数标签参数</strong><hr></td>
  </tr>
<%
   Dim arrFieldParam,FieldParams,FieldParam, i
   FieldParams=Split(rs("Description"),"@@@")
   If Ubound(FieldParams)>0 and FieldParams(1)<>"" Then
       FieldParam=FieldParams(1)
       ArrFieldParam=Split(FieldParam,vbcrlf)
       For i = 0 To UBound(arrFieldParam)
          response.write "<tr><td align='right'><span id='Param" & I &"'>" & arrFieldParam(i) & "</span>：</td><td><input type=""text"" id='Field" & i & "' name='Field" & i & "'></td></tr>"
       Next
    response.write "<tr><td colspan=2 align='center'><input TYPE='button' value=' 确 定 ' onCLICK='objectTag(" & UBound(arrFieldParam)+1 & ")'></td>" 
 Else
 %>
 <script>window.returnValue='<%=Replace(rs("LabelName"),"}","") %>()}';window.close();</script>"
 <%
 response.end
 End If  
%>
  </tr>
</table>
<br>
<hr>
<font color=red>说明：自定义函数标签的调用格式如下：<br><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{SQL_标签名称(<font color=blue>参数1,参数2...</font>)}</font>

</form>
</body>
</html>
<%

    End If
    Set rs = Nothing
End Sub
%>
 
