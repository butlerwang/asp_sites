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
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img id="yule_01" src="images/yule_01.jpg" width="1003" height="6" alt="" /></td>
  </tr>
  <tr>
    <td><table width="1003" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="89%">&nbsp;</td>
            <td width="11%"><img id="resources_03" src="images/resources_03.jpg" width="114" height="118" alt="" /></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="72" height="107" background="images/resources_05.jpg">&nbsp;</td>
            <td width="782"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td background="images/resources_07.jpg"><script type="text/javascript" src="swfobject.js"></script>
        
<div id="10" style="width: 260px; height: 50px"></>
  This text is replaced by the Flash movie.</div>

<script type="text/javascript">
   var so = new SWFObject("resources.swf", "mymovie", "260", "50",  "#000000");
           so.addParam("quality", "best");
           so.addParam("wmode", "transparent");
           so.addParam("menu", "false");
           so.addParam("scale", "noscale");
           so.addParam("flashVars", document.location.search.substr(1));
   so.write("10");
     </script></td>
              </tr>
              <tr>
                <td background="images/resources_11.jpg"><img id="resources_10" src="images/resources_10.jpg" width="399" height="35" alt="" /></td>
              </tr>
              <tr>
                <td background="images/resources_14.jpg"><img id="resources_13" src="images/resources_13.jpg" width="577" height="22" alt="" /></td>
              </tr>
            </table></td>
            <td width="149" align="right"><img id="resources_09" src="images/resources_09.jpg" width="149" height="107" alt="" /></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="8%"><img id="resources_16" src="images/resources_16.jpg" width="79" height="165" alt="" /></td>
            <td colspan="2"><table width="73%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="141" valign="top" background="images/resources_17.jpg" ><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td valign="top" class="ziti1">
                            <%
		            page=Cint(request("page"))
		            activepage=request.QueryString("activepage")
						sql="select Department,AddDate,Content,id from sbe_job where show=-1 order by id desc"
							rs.open sql,conn,1,1
			if not rs.eof then
						rs.pagesize=4
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
					◆ <a href="#" onClick="window.open ('resources1.asp?id=<%=rs(3)%>', 'newwindow', 'height=513, width=807, top=0, left=0, toolbar=no, menubar=no, scrollbars=no, resizable=no,location=no, status=no') 
"><%=rs(0)%>：<%=rs(2)%></a> [<%=rs(1)%>] <br />
						<%rs.movenext
						rowcount = rowcount + 1
						loop
						end if
						rs.close%>
					</td>
                  </tr>
                  <tr>
                    <td align="left" class="ziti2"><%call PageControl(iCount,maxpage,page)%></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td><img id="resources_22" src="images/1111111_09.jpg" width="456" height="24" alt="" /></td>
                    </tr>
                </table></td>
              </tr>
            </table></td>
            <td width="37%"><table width="103%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="104" background="images/resources_18.jpg">&nbsp;</td>
              </tr>
              <tr>
                <td><img id="resources_21" src="images/resources_21.jpg" width="319" height="61" alt="" /></td>
              </tr>
            </table></td>
            <td width="149" align="right" valign="top"><img src="images/resources_20.jpg" alt="" name="resources_20" width="149" height="165" id="resources_20" /></td>
          </tr>
          
        </table></td>
      </tr>
      <tr>
        <td><img id="resources_26" src="images/resources_26.jpg" width="1003" height="125" alt="" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
     <td><!--#include file="down.asp"--></td>
  </tr>
</table>
</body>
</html>
