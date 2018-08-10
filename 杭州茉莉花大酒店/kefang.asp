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
<style type="text/css">
<!--
#znss {
	position:absolute;
	width:200px;
	height:115px;
	z-index:1;
}
-->
</style>
<script language="javascript">
                function correctPNG()
                {
                for(var i=0; i<document.images.length; i++)
                {
                 var img = document.images[i]
                 var imgName = img.src.toUpperCase()
                 if (imgName.substring(imgName.length-3, imgName.length) == "PNG")
                 {
                 var imgID = (img.id) ? "id='" + img.id + "' " : ""
                 var imgClass = (img.className) ? "class='" + img.className + "' " : ""
                 var imgTitle = (img.title) ? "title='" + img.title + "' " : "title='" + img.alt + "' "
                 var imgStyle = "display:inline-block;" + img.style.cssText
                 if (img.align == "left") imgStyle = "float:left;" + imgStyle
                 if (img.align == "right") imgStyle = "float:right;" + imgStyle
                 if (img.parentElement.href) imgStyle = "cursor:hand;" + imgStyle 
                 var strNewHTML = "<span " + imgID + imgClass + imgTitle
                 + " style=\"" + "width:" + img.width + "px; height:" + img.height + "px;" + imgStyle + ";"
                 + "filter:progid:DXImageTransform.Microsoft.AlphaImageLoader"
                 + "(src=\'" + img.src + "\', sizingMethod='scale');\"></span>"
                 img.outerHTML = strNewHTML
                 i = i-1
                 }
                }
                }
                window.attachEvent("onload", correctPNG);
        </script>
<link href="css/css.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
#Layer1 {
	position:absolute;
	width:200px;
	height:115px;
	z-index:1;
}
-->
</style>
<STYLE type=text/css> 
body { 
SCROLLBAR-FACE-COLOR: #DBC188 ; 
SCROLLBAR-HIGHLIGHT-COLOR: #FFFFFF; 
SCROLLBAR-SHADOW-COLOR: #743106; 
SCROLLBAR-3DLIGHT-COLOR: #743106; 
SCROLLBAR-ARROW-COLOR: #743106; 
SCROLLBAR-TRACK-COLOR: #FAF1CB; 
SCROLLBAR-DARKSHADOW-COLOR: #ffffff 
} 
</STYLE> 
</head>

<body>
<table width="1003" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img id="yule_01" src="images/yule_01.jpg" width="1003" height="6" alt="" /></td>
  </tr>
  <tr>
    <td height="131">&nbsp;</td>
  </tr>
  <tr>
    <td height="66" valign="top"><table width="1003" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="1003" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="56" height="111" background="images/kefang_04.jpg">&nbsp;</td>
            <td width="571" valign="top"><table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="331" background="images/kefang_06.jpg"><script type="text/javascript" src="swfobject.js"></script>
        
<div id="9" style="width: 258px; height: 56px"></>
  This text is replaced by the Flash movie.</div>

<script type="text/javascript">
   var so = new SWFObject("kefang1.swf", "mymovie", "258", "56",  "#000000");
           so.addParam("quality", "best");
           so.addParam("wmode", "transparent");
           so.addParam("menu", "false");
           so.addParam("scale", "noscale");
           so.addParam("flashVars", document.location.search.substr(1));
   so.write("9");
     </script></td>
                <td width="239" background="images/kefang_06.jpg">&nbsp;</td>
              </tr>
              <tr>
                <td background="images/kefang_13.jpg"><img id="kefang_11" src="images/kefang_11.jpg" width="331" height="33" alt="" /></td>
                <td background="images/kefang_13.jpg">&nbsp;</td>
              </tr>
              <tr>
                <td colspan="2"><img id="kefang_14" src="images/kefang_14.jpg" width="571" height="22" alt="" /></td>
              </tr>
            </table></td>
            <td width="287" background="images/kefang_08.jpg">&nbsp;</td>
            <td width="89"><img id="kefang_10" src="images/kefang_10.jpg" width="89" height="111" alt="" /></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="57" height="65" background="images/kefang_15.jpg">&nbsp;</td>
                    <td width="475" valign="top" background="images/kefang_17.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td height="40" valign="top">&nbsp;<span class="ziti">&nbsp;&nbsp;拥有准三星级标准装修的温馨、舒适的经济型客房119间，同时您还可以享受本酒店为您提供的免费市话、宽带上网、24小时免费停车等特色服务。</span></td>
                      </tr>
                      <tr valign="top">
                        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="18%"><a href="?id=9" onClick="document.all.Layer1.style.display='block';"><img src="images/kefang_19.jpg" alt="" width="95" height="25" border="0" id="kefang_19" /></a></td>
                            <td width="17%"><a href="?id=18" onClick="document.all.Layer1.style.display='block';"><img src="images/kefang_20.jpg" alt="" width="89" height="25" border="0" id="kefang_20" /></a></td>
                            <td width="17%"><a href="?id=33" onClick="document.all.Layer1.style.display='block';"><img src="images/kefang_21.jpg" alt="" width="90" height="25" border="0" id="kefang_21" /></a></td>
                            <td width="19%"><a href="?id=34" onClick="document.all.Layer1.style.display='block';"><img src="images/kefang_22.jpg" alt="" width="99" height="25" border="0" id="kefang_22" /></a></td>
                            <td width="29%"><a href="?id=35" onClick="document.all.Layer1.style.display='block';"><img src="images/kefang_23.jpg" alt="" width="102" height="25" border="0" id="kefang_23" /></a></td>
                          </tr>
                        </table></td>
                      </tr>
                    </table></td>
                  </tr>
                </table></td>
              </tr>
			  

              <tr>
                <td>
				
				<div id="Layer1" style="position:absolute;width:806px;border:0px ;height:372px;top: 52px; margin-left:100px" <%if request("id")="" then%> style="display:none"<%end if%>>
<table width="817" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="98%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="97%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="1%"><img id="kefang1_03" src="images/kefang1_03.jpg" width="7" height="25" alt="" /></td>
                  <td width="96%" background="images/kefang1_04.jpg"></td>
                  <td width="3%"><img id="kefang1_05" src="images/kefang1_05.jpg" width="28" height="25" alt="" onClick="document.all.Layer1.style.display='none';return false;" /></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td><table width="817" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="1%"><img id="kefang1_08" src="images/kefang1_08.jpg" width="7" height="351" alt="" /></td>
                  <td width="99%" background="images/kefang1_09.jpg">					
				  <%if request("id")>0 then
		            page=Cint(request("page"))
		            activepage=request.QueryString("activepage")
					sql="select num,content,price,id,spic from sbe_product where tid="&request("id")&" and show=-1"
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
                                <td bordercolor="#40260F">
								<%if rs(4)<>"" then%>
								<img id="kefang1_16" src="uploadfile/<%=rs(4)%>" width="393" height="278" alt="" />
								<%end if%>
								</td>
                              </tr>
                                      </table>
                        </div></td>
                        <td width="45%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td colspan="2" class="notice1">
							<div style=" width:100%;overflow:auto;height:166;">
							<table width="100%" border="0" cellspacing="0" cellpadding="0">
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
                            </table>
							</div>
							</td>
                            </tr>
                          
                          <tr>
                            <td height="40"  colspan="2" valign="top" ><table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td width="7%">&nbsp;</td>
                                <td width="93%"><a href="online.asp?id=<%=rs(3)%>&idd=<%=request("id")%>" target="_blank"><img id="kefang1_31" src="images/kefang1_31.jpg" width="109" height="22" alt=""  border="0"/></a></td>
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
							<img id="kefang1_41" src="images/19.jpg" alt="" />
							<%elseif request("id")=36 then%>
							<img id="kefang1_41" src="images/kefang1_4222222.jpg" alt="" />
							<%elseif request("id")=37 then%>
							<img id="kefang1_41" src="images/kefang1_422222.jpg" alt="" />
							<%end if%>							</td>
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
						rs.close
						end if%>
</td>
                  <td width="0%"><img src="images/11_12.jpg" width="3" height="351" /></td>
                </tr>
            </table></td>
          </tr>
        </table></td>
        <td  width="3%"><img src="images/1111_07.png" width="7" height="376" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><img src="images/22_25.png" width="825" height="11" /></td>
  </tr>
</table>				</div>
                  <img id="kefang_24" src="images/kefang_24.jpg" width="532" height="208" alt="" /></td>
              </tr>
            </table></td>
            <td valign="top"><img id="kefang_18" src="images/kefang_18.jpg" width="471" height="273" alt="" /></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
     <td><!--#include file="down.asp"--></td>
  </tr>
</table>
</body>
</html>
