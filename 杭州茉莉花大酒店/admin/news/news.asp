<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
 <%
If Session("name") = "" then
response.Write "<script LANGUAGE=javascript>alert('���糬ʱ�������㻹û�е�½! ');this.location.href='../login.asp';</script>"
end if
 	temp_check_rights = Split(session("manconfig"),",")
	for temp_rights_count=0 to ubound(temp_check_rights)
	    if trim(temp_check_rights(temp_rights_count)) = "3" then
			rights_check_passkey = trim(temp_check_rights(temp_rights_count))
		end if
	next
	if rights_check_passkey <> "3" then
Session.Abandon()
	Response.Write "<Script Language=JavaScript>alert('����Ȩ�޲��㣬�벻Ҫ�Ƿ�������������ģ�飬���������˺Ž���ϵͳ�Զ�ɾ��!');this.location.href='../login.asp';</Script>"
	Response.end
	end if%>
<%
Call OpenData()
 CompanyID = Trim(Request("ID"))
 tid=trim(Request("tid"))
 
If IsSubmit then 
  sqlsize ="select * from Sbe_News_Class where ID ="&tid
  set rssize=conn.execute(sqlsize)
  if not (rssize.eof and rssize.bof) then
    if rssize("ChildNum") >0  then
	   Response.Write "<Script Language=JavaScript>alert('����Ѷ�����»����ӷ���,����ȷѡ�����!');history.go(-1);</Script>"
	   Response.end
	end if
  end if 
  rssize.close
 set rssize=nothing 
  Uploadfile=request.Form("Uploadfile")
  Dim msg  
  Set rs=server.createobject("adodb.recordset")
  If len(CompanyID)<=0 Then
   set rs_max=server.CreateObject("adodb.recordset")
     sql="select max(sequence) as maxid from Sbe_News"
     rs_max.open sql,conn,1,1
     if isnull(rs_max("maxid")) then
        sequence=1
     else
        sequence=rs_max("maxid")+1
     end if
     rs_max.close
     set rs_max=nothing	 
	'msg = "��Ѷ��ӳɹ�!"
	Rs.open "Select * from Sbe_News where id Is null",conn,1,3	
	Rs.addnew
	Rs("Sequence")= sequence
  Else
	'msg = "��Ѷ�޸ĳɹ���"
	Rs.open "Select * from Sbe_News where ID=" & clng(CompanyID) ,conn,1,3	
  End if
  Rs("tid")= tid
  Rs("title")=Request.Form("title")
  Rs("writer")= Request.Form("writer")
  Rs("keyword")= Request.Form("keyword")
  if  Request.Form("newsdate")<>"" then
  Rs("newsdate")= Request.Form("newsdate")
  end if
  Rs("content")= Request.Form("content")
  Rs("pic")= Request.Form("pic")
 ' Rs("tuijian")= Request.Form("tuijian")
 ' Rs("PhotoNew")= Request.Form("PhotoNew")
  Rs("detail")=Request.Form("detail")
  Rs("Bpic")=Request.Form("Bpic")
  Rs("spic")=Request.Form("spic")
  Rs("Show")=request("Show")
  Rs("leibie")=request("leibie")
  Rs("Uploadfile")=Uploadfile
  rs.update
  rs.close
  Set rs=nothing	
   If len(CompanyID)<=0 Then
	Response.Write"<script>alert('��Ѷ���ӳɹ�');this.location.href='news.asp?tid="&tid&"';</script>"
   Else
    Response.Write"<script>alert('��Ѷ�޸ĳɹ�');this.location.href='main.asp';</script>"
   End IF

ElseIF Len(CompanyID)>0 Then	
	Dim StrSQL
	Dim objRec
	StrSQL = "Select * from Sbe_News Where ID=" & CompanyID	
	Set objRec=server.createobject("adodb.recordset")
	 objRec.open StrSQL,conn,1,1
	With ObjRec
		If .Eof And .Bof Then
			Response.Write "<Script>alert('����ʧ��');history.back();</script>" 
			Response.End
		Else
			title = objRec("title")
			tid= objRec("tid")  
            writer= objRec("writer")
            newsdate= objRec("newsdate")
            content= objRec("content")
			tuijian= objRec("tuijian")          
            PhotoNew= objRec("PhotoNew")
			keyword=objRec("keyword")
			pic=objRec("pic")
			detail=objRec("detail")
			Show=objRec("Show")
			Spic=objRec("Spic")
			Bpic=objRec("Bpic")
			leibie=objRec("leibie")
			Uploadfile=objRec("Uploadfile")
		End If
	End With
	objRec.Close:set objRec=Nothing
	   webname="��Ѷ�޸�"
elseif Len(CompanyID)=0 Then 
	   leibie=1
	   webname="��Ѷ���"
End if
'Private Sub MessageBoxOK(strValue,tid)
	'With Response
		'.Write "<script>" & vbcrlf
		'.Write "alert('"+strValue+"');" & vbcrlf
		'.Write "this.location.href='"& request.Cookies("refer_page") &"?tid="& tid &"';" & vbcrlf
		'.Write "</'script>" & vbcrlf
	'End With
'End Sub
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=webname%></title>
<link href="../include/style.css" rel="stylesheet" type="text/css">
<script language="javascript">
  function foreColor()
   {
    var arr = showModalDialog("../eWebEditor/Dialog/selcolor.htm", "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0");
    if (arr != null) document.add.title.value='<font color='+arr+'>'+document.add.title.value+'</font>'
    else document.add.title.focus();
}

function clk(value){
 add.writer.value=value;
}
</script>
<script language="JavaScript">
function check(){
  if(add.tid.value==""){
     alert("��ѡ����࣡");
	 add.tid.focus();
	 return false;
  }
  if(add.title.value==""){
     alert("����д��Ѷ���ƣ�");
	 add.title.focus();
	 return false;
  }
 document.add.Submit.disabled=true;
 document.add.value="���Ժ�..."
  return true;
}
</script>
<SCRIPT language=javascript>
function show_user_rights_menu(menu_id)
{
if (menu_id==0)
{
eval("show_user_rights.style.display=\"none\";");
}
else
{
eval("show_user_rights.style.display=\"\";");
}
}
</SCRIPT>
<script language="JavaScript" src="../include/meizzDate.js"></script>
</head>

<body>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" >
  <tr> 
    <td height="25"><font color="#6A859D">��Ѷ���� &gt;&gt;<%=webname%> </font></td>
  </tr>
  <tr>
    <td height="1" background="../images/dot.gif"></td>
  </tr>
</table>


<br>
<form name="add" method="post" action="" onSubmit="return check()">
<table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0" id="sbe_table">
  <tr align="center">
    <td colspan="3" class="sbe_table_title"><%=webname%></td>
  </tr>
  <tr>
    <td width="16%" align="right">��Ѷ���:</td>
    <td colspan="2">
	     <select name="tid" class="sbe_button" style="width:150px;">
        <option value="">��Ѷ���</option>
          <%
		    Call ShowClass("Sbe_News",tid)%>
        </select> </td>
  </tr>
  <tr>
    <td align="right">��Ѷ����:</td>
    <td colspan="2"><input name="title" type="text" class="input" id="title" style="width:300px;" value="<%=title%>" maxlength="200">
    <!--<img src="../images/forecolor.gif" width="20" height="20" class="Ico" onClick="foreColor();">--></td>
  </tr>
  <tr>
    <td align="right">��Ѷ��Դ:</td>
    <td colspan="2"><input name="writer" type="text" class="input" id="writer" value="<%=writer%>" maxlength="50">
      ѡ��:<%Call news_come_Class()%> ---<a href="news_come_class.asp" onClick="window.open(this.href,'', 'height=350,width=400,toolbar=no,location=no,status=no,menubar=no');return false">��Ѷ��Դ����</a></td>
  </tr>
  <tr class="display">
    <td align="right">�����Ѷ�ؼ���:</td>
    <td colspan="2"><input name="keyword" type="text" class="input" id="keyword" value="<%=keyword%>"></td>
  </tr>
<!--<tr >
    <td align="right">��Ѷ�Ƽ�:</td>
    <td colspan="2"> <input type="radio" name="Tuijian" value="1" <%'Call ReturnSel(tuijian,true,2)%>  >
        �� &nbsp;&nbsp; <input name="Tuijian" type="radio" value="0"  <%'Call ReturnSel(tuijian,false,2)%> >
        ��</td>
  </tr>-->
  <tr class="display">
    <td align="right">ͼƬ��Ѷ:</td>
    <td colspan="2"><input type="radio" name="PhotoNew" value="1"  <%Call ReturnSel(PhotoNew,true,2)%> onclick=show_user_rights_menu(1)>
        �� &nbsp;&nbsp; <input name="PhotoNew" type="radio" value="0"  <%Call ReturnSel(PhotoNew,false,2)%> onclick=show_user_rights_menu(0)>��</td>
  </tr>
  <tr id="show_user_rights" <%if PhotoNew=false then response.write("style='display:none;'") end if%>>
    <td align="right">�ϴ�ͼƬ:</td> 
    <td width="23%"><input name="pic" type="text" class="input" value="<%=pic%>" size="25"></td>
    <td width="61%"><iframe src="../upload/upload.asp?Form_Name=add&UploadFile=pic" width="70%" height="25" frameborder="0" scrolling="no"></iframe> ͼƬ�ߴ磺94*68</td>
  </tr>
  <tr <%=banben_display%>>
    <td align="right">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��:</td>
    <td colspan="2"> <input type="radio" name="leibie" value="1" <%Call ReturnSel(leibie,1,2)%>>
        �� &nbsp;&nbsp; <input name="leibie" type="radio" value="2" <%Call ReturnSel(leibie,2,2)%>>
        Ӣ</td>
  </tr>
  <tr class="display">
    <td align="right">��Ҫ˵��:</td>
    <td colspan="2"><textarea name="detail" cols="50" rows="3" class="input" id="detail"><%=detail%></textarea></td>
  </tr>
  
  <tr>
    <td align="right">��ϸ����:</td>
    <td colspan="2"><textarea name="content" id="textarea" class="display"><%=content%></textarea><input name="Uploadfile" type="hidden" class="input" id="Uploadfile" value="<%=Uploadfile%>"><iframe ID="eWebEditor1" src="../eWebEditor/ewebeditor.asp?id=content&style=sbe&savefilename=uploadfile" frameborder="0" scrolling="no" width="100%" HEIGHT="350"></iframe></td>
  </tr>
  <tr>
    <td align="right">���ʱ��:</td>
    <td colspan="2"><input name="newsdate" type="text" class="input" id="newsdate" onFocus="setday(this)" <%if newsdate="" then response.Write ("value='"&date()&"'") else response.Write ("value='"&newsdate&"'") end if%>></td>
  </tr>
  <tr>
    <td align="right">�Ƿ���ʾ:</td>
    <td colspan="2"> <input type="radio" name="Show" value="1" <%if CompanyID="" then%>checked <%else%> <%Call ReturnSel(Show,true,2)%> <%end if%>>
        �� &nbsp;&nbsp; <input name="Show" type="radio" value="0" <%if CompanyID<>"" then%> <%Call ReturnSel(Show,false,2)%> <%end if%>>
        ��</td>
  </tr>
  <tr align="center">
    <td colspan="3"><input type="hidden" name="ID" value="<%=CompanyID%>"><input name="Submit" type="submit" class="sbe_button" value="�ύ">
    <input name="Submit2" type="reset" class="sbe_button" value="����"></td>
  </tr>
</table>
</form>
<%Call CloseDataBase()%>
</body>
</html>
<%
Private Sub news_come_Class()
'��ȡ��Ѷ��Դ
 Set oRs=Conn.Execute("select * from news_come_class order by id asc")
 IF oRs.Eof and oRs.bof Then Exit Sub
 Do While not oRs.eof 
  response.write "<a href=""javascript:clk('"& oRs("title") &"');"" >"& oRs("title") &"</a>/"& vbCrLf
 oRs.Movenext
 Loop
End Sub
%>