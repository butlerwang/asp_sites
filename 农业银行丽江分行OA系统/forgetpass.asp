<!--#INCLUDE FILE="data.asp" -->
<html><head><title>��������칫ϵͳ----��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<Script Language="javaScript">
    function  validate1()
    {
       
        if  (document.myform.Userid.value=="")
        {
            alert("��¼�ʺŲ���Ϊ��");
            document.myform.Userid.focus();
            return false ;
        }
        }
function  validate2()
    {
        if  (document.myform.answer.value=="")
        {
            alert("����𰸲���Ϊ��");
            document.myform.answer.focus();
            return false ;
        }
        
}
</Script>

<link rel="stylesheet" href="oa.css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form method="post" action="forgetpass.asp" name="myform">
  <table width="100%" border="0" cellspacing="1" cellpadding="2">
    <tr > 
      <td class="heading"> 
        <div align="center">
          <center> 
        <table width="81%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000" bordercolorlight="#000000">
          <tr> 
            <td width="2%" align="right"></td>
            <td align="left" height="25">
              <p align="center"><font color="#FFFFFF"><b>�� �� �� ��</b></font></p>
            </td>
            <td width="3%"></td>
          </tr>
        </table>
          </center>
        </div>
      </td>
    </tr>
  </table>
  <div align="center">
  <table width="80%" border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF" bordercolor="#FFFFFF" bordercolorlight="#000000">
 <%if request("one")="" then%>
    <tr> 
      <td width="17%" valign="top">
        <p align="right">��¼�ʺ�:</font></p>
      </td>
    <center>
      <td width="83%"> 
        <input type="text" name="Userid" class="form" size="24"><INPUT TYPE="hidden" name="one" value="one">
      </td>
    </tr>
    </center>
  </table>
  </div><BR>
  <div align="center"><input type=image  src="images/next.gif"  onclick="return  validate1()">&nbsp;&nbsp;      
  <a  href="javaScript:window.close()"><img   border="0" src="images/close_1.gif"></a></div>    
  <%
   elseif request("Userid")<>"" then
    Set rs= Server.CreateObject("ADODB.Recordset") 
    strSql="select * from user where �û���='"&request("Userid")&"'"
    rs.open strSql,Conn,1,1 
      if rs.eof then
  %>
  <tr> 
    <center>
      <td colspan=2 align=center> 
         <FONT COLOR="red"><B>���ʺŲ����ڣ�</B></FONT>
      </td>
    </tr>
    </center>
  </table>
  </div><BR><div align="center">&nbsp;&nbsp;      
  <a  href="javascript:history.back(1)"><img border="0" src="images/previous.gif"></a>&nbsp;&nbsp;      
  <a  href="javaScript:window.close()"><img   border="0" src="images/close_1.gif"></a></div> 
  <%else%>
  <tr> 
      <td width="17%" valign="top">
        <p align="right">�� �� ��:</font></p>
      </td>
    <center>
      <td width="83%"> 
      <%=rs("�û���")%>
      </td>
    </tr>
    </center>
  <tr> 
      <td width="17%" valign="top">
        <p align="right">��������:</font></p>
      </td>
    <center>
      <td width="83%"> 
      <%=rs("����")%>
      </td>
    </tr>
    </center>
  <tr> 
      <td width="17%" valign="top">
        <p align="right">�����:</font></p>
      </td>
    <center>
      <td width="83%"> 
        <input type="text" name="answer" class="form" size="24"><INPUT TYPE="hidden" name="one" value="one"><INPUT TYPE="hidden" name=user value="<%=rs("id")%>">
      </td>
    </tr>
    </center>
  </table>
  </div><BR><div align="center">&nbsp;&nbsp;      
  <a  href="javascript:history.back(1)"><img border="0" src="images/previous.gif"></a>&nbsp;&nbsp;<input type=image  src="images/next.gif" onclick="return  validate2()">&nbsp;&nbsp;       
  <a  href="javaScript:window.close()"><img   border="0" src="images/close_1.gif"></a></div> 
  <%end if%>
  <%
    elseif request("answer")<>"" then 
    Set rs= Server.CreateObject("ADODB.Recordset") 
    strSql="select * from user where id="&request("user")
    rs.open strSql,Conn,1,1 
       if rs("��")<>request("answer") then 
 %>
  <tr> 
    <center>
      <td colspan=2 align=center> 
         <FONT COLOR="red"><B>�𰸴���</B></FONT>
      </td>
    </tr>
    </center>
  </table>
  </div><BR><div align="center">&nbsp;&nbsp;      
  <a  href="javascript:history.back(1)"><img border="0" src="images/previous.gif"></a>&nbsp;&nbsp;      
  <a  href="javaScript:window.close()"><img   border="0" src="images/close_1.gif"></a></div> 
 <%else%>  
    <tr> 
      <td width="17%" valign="top">
        <p align="right">�� �� ��:</font></p>
      </td>
    <center>
      <td width="83%"> 
      <%=rs("�û���")%>
      </td>
    </tr>
    </center>
  <tr> 
      <td width="17%" valign="top">
        <p align="right"> ��  �� :</font></p>
      </td>
    <center>
      <td width="83%"> 
      <%=rs("����")%>
      </td>
    </tr>
    </center>
  
  </table>
  </div><BR><div align="center">&nbsp;&nbsp;      
  <a  href="javascript:history.back(1)"><img border="0" src="images/previous.gif"></a>&nbsp;&nbsp;       
  <a  href="javaScript:window.close()"><img   border="0" src="images/close_1.gif"></a></div> 
  <%end if%>
  <%end if%> 
      
     
     
</form>     
     
<div align="center">     
  <center>     
  <table border="1" width="80%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF" bordercolor="#C0C0C0">     
    <tr>     
      <td width="100%">     
      ϵͳ������:</font>    
      <ul>    
        <li>������������ע���ʺ�</li>   
        <li>��ϵͳ��֤�ʺŴ��ں���ʾ������ʾ����</li>   
        <li>���������������</li>  
        <li>��ȫ��ȷ����ʾ��ע��ʱ����</li> 
        <li>�����κ�������͹���Ա��ϵ</li> 
      </ul>       
      </td>       
    </tr>       
  </table>       
  </center>       
</div>       
       
</body>       
</html>

