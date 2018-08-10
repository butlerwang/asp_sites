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
<table width="1003" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img  src="images/yule_01.jpg" width="1003" height="6" ></td>
  </tr>
  <tr>
    <td><table width="1003" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="1003" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="787" background="images/2222_03.jpg">&nbsp;</td>
            <td><img id="news_03" src="images/news_03.jpg" width="216" height="135" alt="" /></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="1003" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="54" background="images/news_05.jpg">&nbsp;</td>
            <td width="578"><table width="97%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td background="images/news_07.jpg"><script type="text/javascript" src="swfobject.js"></script>
        
<div id="7" style="width: 255px; height: 51px"></>
  This text is replaced by the Flash movie.</div>

<script type="text/javascript">
   var so = new SWFObject("news1.swf", "mymovie", "255", "51",  "#000000");
           so.addParam("quality", "best");
           so.addParam("wmode", "transparent");
           so.addParam("menu", "false");
           so.addParam("scale", "noscale");
           so.addParam("flashVars", document.location.search.substr(1));
   so.write("7");
     </script></td>
              </tr>
              <tr>
                <td height="33" background="images/news_11.jpg"><img id="news_10" src="images/news_10.jpg" width="418" height="33" alt="" /></td>
              </tr>
              <tr>
                <td><img id="news_13" src="images/news_13.jpg" width="578" height="21" alt="" /></td>
              </tr>
            </table></td>
            <td width="371"><img  src="images/news_09.jpg" width="370" height="105" ></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="44%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="55" height="144" background="images/news_15.jpg">&nbsp;</td>
                    <td valign="top" background="images/news_16.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td height="110" valign="top" class="ziti4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <%
		            page=Cint(request("page"))
		            activepage=request.QueryString("activepage")
							if request("type")<>"" then
							sql="select id,title,newsdate,sequence from sbe_news where tid="&request("type")&" and show=-1 order by sequence desc"
							else
							sql="select id,title,newsdate,sequence from sbe_news where tid=1 and show=-1 order by sequence desc"
							end if
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
							<tr>
                              <td width="16" height="25"><a href="news1.asp">◆</a></td>
                              <td width="280"><a href="#" onClick="window.open ('news1.asp?sequence=<%=rs(3)%>&type=<%=request("type")%>&id=<%=rs(0)%>', 'newwindow', 'height=635, width=807, top=0, left=0, toolbar=no, menubar=no, scrollbars=no, resizable=no,location=no, status=no') 
"><%=gotTopic(rs(1),40)%></a></td>
                              <td width="75">[<%=rs(2)%>]</td>
                            </tr>
						<%rs.movenext
						rowcount = rowcount + 1
						loop
						end if
						rs.close%>
                          </table>
                         </td>
                      </tr>
                      <tr>
                        <td><%call PageControl(iCount,maxpage,page)%> </td>
                      </tr>
                    </table></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="14%" background="images/news_19.jpg">&nbsp;</td>
                    <td width="51%" background="images/news_21.jpg"><img src="images/news_20.jpg" alt="" name="news_20" width="227" height="24" border="0" usemap="#news_20Map" id="news_20" /></td>
                    <td width="35%" align="right" valign="top" background="images/news_21.jpg"><img src="images/news_21.jpg" alt="" width="163" height="24" border="0" usemap="#news_21Map" id="news_21" /></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td><img  src="images/news_22.jpg" width="445" height="107"></td>
              </tr>
            </table></td>
            <td width="56%" background="images/1111111_17.jpg"><script type="text/javascript" src="swfobject.js"></script>
        
<div id="6" style="width: 558px; height: 275px"></>
  This text is replaced by the Flash movie.</div>

<script type="text/javascript">
   var so = new SWFObject("news.swf", "mymovie", "558", "275",  "#000000");
           so.addParam("quality", "best");
           so.addParam("wmode", "transparent");
           so.addParam("menu", "false");
           so.addParam("scale", "noscale");
           so.addParam("flashVars", document.location.search.substr(1));
   so.write("6");
     </script></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
     <td><!--#include file="down.asp"--></td>
  </tr>
</table>

<map name="news_20Map">
<area shape="rect" coords="4,2,114,22" href="news.asp?type=1">
<area shape="rect" coords="120,2,224,20" href="news.asp?type=4">
</map>
<map name="news_21Map"><area shape="rect" coords="98,13,99,14" href="#"></map></body>
</html>
