<!--#include file="data.asp"-->
<!--#include file="html.asp"-->
<%if Session("Urule")<>"a" then
	Response.write "��û���㹻Ȩ��"
	response.end
end if%>

<%
response.buffer=false
dim sql
dim rs
dim id


 name=htmlencode2(request("name"))
 password=request("password")
 userid=htmlencode2(request("userid"))
 question=htmlencode2(request("question"))
 answer=htmlencode2(request("answer"))
 email=request("email")
 tel=request("tel")
 department=htmlencode2(request("company"))
 ip= Request.ServerVariables("REMOTE_ADDR")
 nowtime=now()
sj=cstr(year(nowtime))+"-"+cstr(month(nowtime))+"-"+cstr(day(nowtime))+" "+cstr(hour(nowtime))+":"+right("0"+cstr(minute(nowtime)),2)+":"+right("0"+cstr(second(nowtime)),2)

set rs=server.createobject("ADODB.recordset") 
rs.Open "SELECT * FROM user where id="&request("id"),conn,3,3 
rs("�û���")=userid
rs("����")=password
rs("����")=email
rs("����")=department
rs("����")=question
rs("��")=answer
rs("�绰")=tel
rs("����")=name
rs("Ȩ��")=request("admin")
rs("mobile")=request("mobile")
rs("ilevel")=request("ilevel")
rs.update 
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<table width="100%" border="0" cellspacing="1" cellpadding="2">
    <tr > 
      <td class="heading"> 
        <div align="center">
          <center> 
        <table width="81%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000" bordercolorlight="#000000" style="font-size:9pt">
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
  <table width="80%" border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF" bordercolor="#FFFFFF" bordercolorlight="#000000" style="font-size:9pt">
    <tr> 
      <td valign="top">
        <p align="center">���������Ѿ��ɹ��޸�</p>
        </font></td>
   </tr>
    
  </table>
  </div>
  <div align="center"><a  href="javascript:history.back(1)"><img border="0" src="images/previous.gif"></a>&nbsp;&nbsp;      
  <a  href="userchk.asp" target=main><img border="0" src="images/close_1.gif" onclick="window.close();"></a>     
       
     
  </div> 