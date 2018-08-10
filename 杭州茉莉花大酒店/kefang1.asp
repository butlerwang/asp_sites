<!-- #include file="inc/conn.asp"-->
<!-- #include file="Check_Sql.asp"-->
<!-- #include file="inc/lib.asp"-->
<%OpenData()%>
<%set rs=server.CreateObject("adodb.recordset")%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>杭州茉莉花大酒店</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<link href="css/css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="811" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="1%"><img id="kefang1_03" src="images/kefang1_03.jpg" width="7" height="25" alt="" /></td>
                    <td width="96%" background="images/kefang1_04.jpg"></td>
                    <td width="3%"><img id="kefang1_05" src="images/kefang1_05.jpg" width="28" height="25" alt="" /></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="1%"><img id="kefang1_08" src="images/kefang1_08.jpg" width="7" height="351" alt="" /></td>
                    <td width="99%" background="images/kefang1_09.jpg">
					<%
		            page=Cint(request("page"))
		            activepage=request.QueryString("activepage")
					sql="select num,content,price,id from sbe_product where tid="&request("id")&""
					rs.open sql,conn,1,1
			if not rs.eof then
						rs.pagesize=1
                      iCount=rs.RecordCount '记录总数
                      iPageSize=rs.PageSize
                      maxpage=rs.PageCount 
						if activepage = "next" then
							page = page + 1
							else if activepage = "up" then
								page = page - 1
									else if activepage = "first" then
										page = 1
											else if activepage = "last" then
												page = rs.pagecount
												end if
										end if
								end if
						end if
					

						if page=0 then
							page=1
						end if
						
						if page > rs.pagecount then
							page = rs.pagecount
						end if
	
						rs.absolutepage = CInt(page)
						rowcount = 0
                   
					 %>
					   <%do while ( not rs.eof and rowcount < rs.pagesize )%>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="55%"><div align="center">
                          <table width="51%" border="1" cellpadding="0" cellspacing="1" bordercolor="#40260F">
                              <tr>
                                <td bordercolor="#40260F"><img id="kefang1_16" src="images/kefang1_16.jpg" width="393" height="278" alt="" /></td>
                              </tr>
                                      </table>
                        </div></td>
                        <td width="45%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="23%" class="notice1"><div align="right">数　　量：</div></td>
                            <td width="77%" class="notice1"><%=rs(0)%></td>
                          </tr>
                          <tr>
                            <td valign="top" class="notice1"><div align="right">客房描述：</div></td>
                            <td class="notice2"> <%=HTMLcode(rs(1))%></td>
                          </tr>
                          <tr valign="top">
                            <td class="notice1"><div align="right">价&nbsp; &nbsp; 格：</div></td>
                            <td class="notice2"><%=HTMLcode(rs(2))%></td>
                          </tr>
                          <tr>
                            <td height="40"  colspan="2" valign="top" ><table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td width="7%">&nbsp;</td>
                                <td width="93%"><a href="online.asp?id=<%=rs(3)%>&idd=<%=request("id")%>"><img id="kefang1_31" src="images/kefang1_31.jpg" width="109" height="22" alt=""  border="0"/></a></td>
                              </tr>
                            </table></td>
                          </tr>
                          <tr>
                            <td height="40" colspan="2" align="right" valign="bottom">
							<%if request("id")=9 then%>
							<img id="kefang1_41" src="images/kefang1_41.jpg" width="108" height="27" alt="" />
							<%elseif request("id")=18 then%>
							<img id="kefang1_41" src="images/kefang1_422.jpg" alt="" />
							<%elseif request("id")=33 then%>
							<img id="kefang1_41" src="images/kefang1_4222.jpg" alt="" />
							<%elseif request("id")=34 then%>
							<img id="kefang1_41" src="images/kefang1_42222.jpg" alt="" />
							<%elseif request("id")=35 then%>
							<img id="kefang1_41" src="images/kefang1_422.jpg" alt="" />
							<%elseif request("id")=36 then%>
							<img id="kefang1_41" src="images/kefang1_4222222.jpg" alt="" />
							<%elseif request("id")=37 then%>
							<img id="kefang1_41" src="images/kefang1_422222.jpg" alt="" />
							<%end if%>
							
							
							</td>
                          </tr>
                          <tr>
                            <td colspan="2"><img id="kefang1_44" src="images/kefang1_44.jpg" width="386" height="11" alt="" /></td>
                          </tr>
                          <tr>
                            <td colspan="2" align="right" valign="top" class="ziti4"><%call PageControl(iCount,maxpage,page)%></td>
                          </tr>
                        </table></td>
                      </tr>
                    </table>
						<%rs.movenext
						rowcount = rowcount + 1
						loop
						end if
						rs.close%>
					</td>
                    <td width="0%"><img src="images/11_12.jpg" width="3" height="351" /></td>
                  </tr>
                </table></td>
              </tr>
            </table>
</body>
</html>
