<!--#include file="data.asp"-->
<!--#include file="html.asp"-->

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
set rs=server.createobject("ADODB.recordset") 
rs.Open "SELECT * FROM user where id="&Session("Uid"),conn,1,3 
rs("用户名")=userid
rs("密码")=password
rs("信箱")=email
rs("部门")=department
rs("问题")=question
rs("答案")=answer
rs("电话")=tel
rs("姓名")=name
rs("mobile")=request("mobile")
rs("iPageSize")=request("iPageSize")
rs("iAdd")=request("iAdd")
Session("iPageSize")=request("iPageSize")
Session("iAdd")=request("iAdd")
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
              <p align="center"><font color="#FFFFFF"><b>修 改 资 料</b></font></p>
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
        <p align="center">您的资料已经成功修改</p>
        </font></td>
   </tr>
    
  </table>
  </div>
  <div align="center"><a  href="passwd.asp"><img border="0" src="images/previous.gif"></a>   
       
     
  </div> 