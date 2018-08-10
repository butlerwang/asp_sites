<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title><script language="javascript">
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
</head>

<body>
<table width="538" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="538"><table border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="7"><img id="kefang1_03" src="images/kefang1_03.jpg" width="7" height="25" alt="" /></td>
                  <td width="510" background="images/kefang1_04.jpg"></td>
                  <td width="28"><img src="images/kefang1_05.jpg" alt="" name="kefang1_05" width="28" height="25" id="kefang1_05" /></td>
                </tr>
            </table></td>
            <td width="17" rowspan="2"><img src="images/9_07.png" width="16" height="376" /></td>
          </tr>
          <tr>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr valign="middle">
                  <td width="7"><img id="kefang1_08" src="images/kefang1_08.jpg" width="7" height="351" alt="" /></td>
                  <td width="535" background="images/kefang1_09.jpg"><table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="55%" valign="middle"><div align="center">
                            <table width="51%" border="1" cellpadding="0" cellspacing="1" bordercolor="#40260F">
                              <tr>
                                <td bordercolor="#40260F"><img src="images/44444.jpg" alt="" name="kefang1_16" width="394" height="279" id="kefang1_16" /></td>
                              </tr>
                            </table>
                        </div></td>
                      </tr>
                  </table></td>
                  <td width="3"><img src="images/11_12.jpg" width="3" height="351" /></td>
                </tr>
    </table></td>
          </tr>
  <tr>
    <td colspan="2"><img src="images/91_27.png" width="562" height="13" /></td>
    </tr>
</table>
</td></tr></table>
</body>
</html>
