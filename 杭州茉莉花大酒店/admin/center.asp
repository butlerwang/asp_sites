<html>
<title>关闭左边菜单</title>
<script language="javascript">
<!--
var displayBar=true;
function switchBar()
{
	if (displayBar)
	{
		parent.frame.cols="0,10,*";
		displayBar=false;
		//obj.src="images/admin_show_left.gif";
		//obj.title="打开左边管理菜单";
	}
	else{
		parent.frame.cols="170,10,*";
		displayBar=true;
		//obj.src="images/admin_hide_left.gif";
		//obj.title="关闭左边管理菜单";
	}
}
-->
</script>

<link href="include/style.css" rel="stylesheet" type="text/css">
<body>
<table height="100%" border=0 cellpadding=0 cellspacing=0>
  <tr> 
   
    <td width="10" valign="middle" style="border-left:1px solid #ccc;border-right:1px solid #ccc; CURSOR: hand" bgcolor="#eeeeee" onClick="switchBar()">&nbsp; </td>
   
  </tr>
</table>
</body>
</html>