<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%
Call OpenData()
tid=request("tid")
CompanyID=request("id")
If IsSubmit then  
 Set rs=server.createobject("adodb.recordset")
  If CompanyID = "" Then
	msg = "��Ϣ��ӳɹ�!"
	Rs.open "Select * from Sbe_Company where id Is null",conn,1,3	
	Rs.addnew
   Else
	msg = "��Ϣ�޸ĳɹ���"
	Rs.open "Select * from Sbe_Company where ID=" & clng(CompanyID) ,conn,1,3	
  End if  
    Rs("Tid")=request("select")
	Rs("Content")=request("content")
	Rs("Photo")=request("companyphoto")
	Rs("jianjie")=request("jianjie")
	Rs("Uploadfile")=request("Uploadfile")
	rs.update
  rs.close  
  Set rs=nothing	
	Call MessageBoxOK(msg) '�����ʾ
ElseIF Len(tid)>0 Then	
	Dim StrSQL
	Dim objRec
	StrSQL = "Select * from Sbe_Company Where tID=" & tid
	Set objRec = Conn.Execute(StrSQL)
	With ObjRec
		If .Eof And .Bof Then
			content=""
			companyphoto=""
			'id=""
			'tid=""
		Else	
		    id=objRec("id")    
			content = objRec("content")
			companyphoto = objRec("Photo")
			tid=objRec("tid")
			jianjie=objRec("jianjie")
			Uploadfile=objRec("Uploadfile")
		End If
	End With
	objRec.Close:set objRec=Nothing
End if
Private Sub MessageBoxOK(strValue)

	With Response
		.Write "<script>" & vbcrlf
		.Write "alert('"+strValue+"');" & vbcrlf
		.Write "this.location.href='"& Request.ServerVariables("HTTP_REFERER") &"'" & vbcrlf
		.Write "</script>" & vbcrlf
	End With
End Sub
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�����Ѷ</title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<script language="javascript">
function CheckForm()
{
  if(document.add.select.value == ""){
   alert("��Ϣ�����Ϊ��!");
   document.add.select.focus();
   return false;
  }
if (eWebEditor1.getHTML()==""){    
      alert("ϵͳ��ʾ\n���ݲ���Ϊ��");
     return false;
    }
}	
</script>
</head>

<body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" >
  <tr> 
    <td height="25"><font color="#6A859D">��ҵ��Ϣ&gt;&gt; ������Ϣ</font></td>
  </tr>
  <tr>
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>


<br>
<form name="add" method="post" OnSubmit="return CheckForm();">
<table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0" id="sbe_table">
  <tr align="center">
    <td colspan="3" class="sbe_table_title">��ҵ��Ϣ</td>
  </tr>
  <tr>
    <td width="16%" align="right">��Ϣ���:</td>
    <td colspan="2">
	     <select name="select" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location='?tid='+jmpURL;} else {this.selectedIndex=0 ;}">
          <option>��ѡ��...</option>
          <%
		    Call ShowClass("Sbe_Company",tid)%>
        </select> </td>
  </tr>
<tr class="display">
    <td align="right">���:</td> 
    <td colspan="2"><textarea name="jianjie" id="textarea" style="display:none"><%=jianjie%></textarea><iframe ID="eWebEditor1" src="../eWebEditor/ewebeditor.asp?id=jianjie&style=s_mini1" frameborder="0" scrolling="no" width="100%" HEIGHT="200"></iframe></td>
  </tr>
  <tr>
    <td align="right">��Ϣ����:</td>
    <td colspan="2"><textarea name="content" id="textarea" style="display:none"><%=content%></textarea><input name="Uploadfile" type="hidden" value="<%=Uploadfile%>"><iframe ID="eWebEditor1" src="../eWebEditor/ewebeditor.asp?id=content&style=sbe&savefilename=uploadfile" frameborder="0" scrolling="no" width="100%" HEIGHT="350"></iframe></td>
  </tr>
  <tr class="display">
    <td align="right">�ϴ���ҳͼƬ:</td> 
    <td width="23%"><input name="companyphoto" type="text" value="<%=companyphoto%>" size="25"></td>
    <td width="61%"><iframe src="../upload/upload.asp?Form_Name=add&UploadFile=companyphoto" width="100%" height="25" frameborder="0" scrolling="no"></iframe></td>
  </tr>
  <tr align="center">
    <td colspan="3"><input type="hidden" name="id" value="<%=id%>"><input name="Submit" type="submit" class="sbe_button" value="�ύ">
    <input name="Submit2" type="reset" class="sbe_button" value="����"></td>
  </tr>
</table>
</form>
<%Call CloseDataBase()%>
</body>
</html>
