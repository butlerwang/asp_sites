<%
response.Expires = -1
response.ExpiresAbsolute = Now() - 1
response.Expires = 0
response.CacheControl = "no-cache"
Response.CodePage=65001
Response.CharSet="utf-8"
Dim fieldname, num, dbname, dbtype, isknow,isidarr,isid,datasourcetype

fieldname = Trim(Request("fieldname"))
dbname = Trim(Request("dbname"))
datasourcetype=request("datasourcetype")
isidarr=split(fieldname,".")
isid=false
if ubound(isidarr)=1 then
  if lcase(isidarr(1))="id" and datasourcetype="0" then
    isid=true
  end if
end if

If dbname = "" Then dbname = 0
dbtype = Trim(Request("dbtype"))
If dbtype = "" Then dbtype = 0
isknow = False
%>
<html>
<head>
<title>�ֶ���������</title>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>
<link href='Admin_Style.css' rel='stylesheet' type='text/css'>
<script language = 'JavaScript'>
function changemode(){
    var dbname=document.myform.ftype.value;
    if(dbname=='Text'){
    input1.style.display='';
    }else{
    input1.style.display='none';
    }
    if(dbname=='Num'){
    input2.style.display='';
    }else{
    input2.style.display='none';
    }
    if(dbname=='Date'){
    input3.style.display='';
    }else{
    input3.style.display='none';
    }
}
function changeDate(){
    var dbname=document.myform.Datetype.value;
    if(dbname=='3'){
    document.myform.Datemb.value="2";
    }else{
        document.myform.Datemb.value="YY-MM-DD hh:mm:ss";
    }
}
function submitdate(){
    var dbname=document.myform.ftype.value;
    if(dbname=='Text'){
        dbname="{@" + document.myform.FieldName.value + "(" + dbname + "," + document.myform.CatNum.value + "," + document.myform.CatType.value + "," + document.myform.OutSplit.value + ","+document.myform.NullChar.value+")}";
    }
    if(dbname=='Num'){
	    for (var i=0;i<document.myform.OutType.length;i++){
            if (document.myform.OutType[i].checked){
                var OutType=document.myform.OutType[i].value;
        }
        }
       dbname="{@" + document.myform.FieldName.value + "(" + dbname + "," + OutType + "," + document.myform.XiaoShu.value + ")}";
    }
    if(dbname=='Date'){
    dbname="{@" + document.myform.FieldName.value + "(" + dbname + "," + document.myform.Datemb.value + ")}";
    }
   
    parent.InsertValue(dbname);
	parent.closeWindow();
}
</script>
</head>
<body>
<table width="100%">
<form method='post' action='' name='myform'>
    <tr class="tdbg"><td><strong>�ֶ����ƣ�</strong><input name='FieldName' type='text' id='FieldName' size='20' value="<% =fieldname %>" readonly></td></tr>
    <tr class="tdbg"><td><strong>������ͣ�</strong><select name="ftype" style="width:150" onChange="changemode()">
	<option value='Text'>�ı���</option>
<%
If dbtype = 4 Or dbtype=12 Then
    response.write "<option value='Num' selected>������</option>"
    isknow = True
Else
    response.write "<option value='Num'>������</option>"
End If
If dbtype = 5 Then
    response.write "<option value='Date' selected>ʱ����</option>"
    isknow = True
Else
    response.write "<option value='Date'>ʱ����</option>"
End If

%>
</select></td></tr>
<%
If isknow = False Then
    response.write "<tbody id='input1' style='display:'>"
Else
    response.write "<tbody id='input1' style='display:none'>"
End If
%>
    <tr class="tdbg"><td><strong>������ȣ�</strong><input name='CatNum' type='text' id='gotopic' size='6' value=0>
    &nbsp;&nbsp;&nbsp;<font color='#FF0000'>Ϊ0�򲻽ض�</font></td></tr>
	<tr class="tdbg"><td><strong>�ض���ʾ��</strong><Input name='CatType' type='text' value='...' size="6">
	  &nbsp;&nbsp;&nbsp;<font color='#FF0000'>Ϊ0����ʾ�κα�־</font></td>
	</tr>
    <tr class="tdbg"><td><strong>���˴���</strong><select name='OutSplit'><option value='0' selected>����HTML���</option><option value='1'>������HTML���</option><option value='2'>����HTML���</option></select></td></tr>
	    <tr class="tdbg"><td><strong>�ֶ�ֵΪ��ʱ�����</strong><input title='(��ͼƬֵΪ�գ������һ��Ĭ�ϵ�ͼƬ "/Images/defaule.gif")' name='NullChar' type='text' id='NullChar' size='20' value=""></td></tr>

</tbody>

<%
If (dbtype = 4 Or dbtype=12) Then
    response.write "<tbody id='input2' style='display:'>"
Else
    response.write "<tbody id='input2' style='display:none'>"
End If
%>
    <tr class="tdbg"><td><strong>�����ʽ��</strong><Input type='radio' name='OutType' value='0' checked onClick="input21.style.display='';input22.style.display='none'">
    ԭ�� 
        <Input type='radio' name='OutType' value='1' onClick="input21.style.display='none';input22.style.display=''">С�� <Input type='radio' name='OutType' value='2' onClick="input21.style.display='none';input22.style.display='none'">�ٷ���</td></tr>
<%
        If dbtype = 12 Then
        response.write "<tbody id='input21' style='display:'>"
        Else
        response.write "<tbody id='input21' style='display:none'>"
        End If
%>
</tbody>
    <tbody id='input22' style='display:none'><tr class="tdbg"><td><strong>С��λ����</strong><input name='XiaoShu' type='text' id='XiaoShu' size='4' value=2></td></tr></tbody>
</tbody>


<%
If dbtype = 5 Then
    response.write "<tbody id='input3' style='display:'>"
Else
    response.write "<tbody id='input3' style='display:none'>"
End If
%>
    
    <tr class="tdbg">
      <td><strong>�����ʽ��</strong>
        <input name='Datemb' type='text' id='Datemb' size='28' value="YYYY-MM-DD">
		<br>
		<font color=red>YYYY:��(4λ) YY:��(2λ) ��MM:�� ��DD:��<br>
		hh:ʱ�� mm:�֡� ss:��</font></td>
    </tr>
</tbody>


<%
If dbtype = 11 Then
    response.write "<tbody id='input4' style='display:'>"
Else
    response.write "<tbody id='input4' style='display:none'>"
End If
%>
    
</tbody>


<%
If (LCase(fieldname) = "tid" or LCase(fieldname) = "ks_class.id" or LCase(fieldname) = "id"  or isid or LCase(fieldname) = "newsid" Or LCase(fieldname) = "picid" Or LCase(fieldname) = "downid" Or LCase(fieldname) = "flashid" Or LCase(fieldname) = "proid" or LCase(fieldname) = "movieid" or LCase(fieldname) = "gqid" or LCase(fieldname) = "classid") and datasourcetype="0" Then
    response.write "<tbody id='input5' style='display:'>"
Else
    response.write "<tbody id='input5' style='display:none'>"
End If

%>
<tr class="tdbg"><td><strong>�����ʽ��</strong>
<Input type='radio' name='outype' value=0>
��� 
<Input type='radio' name='outype' value='1' checked>
����Url 
<Input type='radio' name='outype' value='2'> 
�ֶ�ֵ </td>
</tr>
</tbody>

<tr class="tdbg"><td  style="text-align:center"><input class="button" type='button' value="����" onClick="submitdate();">&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' class="button" value="ȡ��" onClick="parent.closeWindow();"></td></tr>
<tr class="tdbg" height="100%"><td>&nbsp;<input name='Fieldnum' id='Fieldnum' value="<% =num %>" type='hidden'><br>&nbsp;<br>&nbsp;</td></tr>
</form>
</table>
</body>
</html>
 
