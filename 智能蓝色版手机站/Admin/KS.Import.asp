<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%


Dim KSCls
Set KSCls = New Import
KSCls.Kesion()
Set KSCls = Nothing

Class Import
        Private KS,KSCls,ChannelID,IConnStr,Iconn,tempField
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		 Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		Sub Kesion()
		 If KS.S("Action")="testsource" Then
		   Call testsource()
		   Exit Sub
		 End If
		 With KS
			.echo "<html>"
			.echo "<title>基本参数设置</title>"
			.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
			.echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			.echo "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
           %>
		    <script type="text/javascript">
			function datachanage(){
				  switch (parseInt($('#datasourcetype').val()))
					{
					 case 1:
					  $('#datasourcestr').val('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=/数据库.mdb');
					  $('#tablename').val('Table1');
					  break;
					 case 3:
					  $('#datasourcestr').val('Provider=Sqloledb; User ID=用户名; Password=密码; Initial Catalog=数据库名称; Data Source =(local);');
					  $('#tablename').val('Table1');
					  break;
					 case 2:
					  $('#datasourcestr').val('driver={microsoft excel driver (*.xls)};dbq=/数据库.xls');
					  $('#tablename').val('Sheet1$');
					  break;
					}
			}
			function testsource()
		    {
			  var str = $('#datasourcestr').val();
			  var datatype=$('#datasourcetype').val();
			  if (str=='')
			  {
				alert('请输入连接字符串!');
				$('#datasourcestr').focus();
				return false;
			  }
			  var url = 'KS.Import.asp';
			  $.get(url,{action:"testsource",datatype:datatype,str:escape(str)},function(d){
				if (d=='true')
				 alert('恭喜，测试通过!')
				else
				 alert('对不起，字符串连接有误!');
			  });
		    } 
			function checkNext()
			{
			  if ($("#channelid>option:selected").val()==0){
			     alert('请选择要导入的模型!');
				 return false;
			  }
			  if ($("#datasourcestr").val()=='')
			  {
			    alert('请输入数据源连接字串!');
				$("#datasourcestr").focus();
				return false;
			  }
			  if ($("#tablename").val()=='')
			  {
			    alert('请输入数据表名!');
				$("#tablename").focus();
				return false;
			  }
			   return true;
			}
		   function getClass(v){
		      if (v==1){
			   $("#stid1").show();
			   $("#stid2").hide();
			  }else{
			   $("#stid1").hide();
			   $("#stid2").show();
			  }
		   }
		   function getTemplate(v){
		      if (v==1){
			   $("#stemplate1").show();
			   $("#stemplate2").hide();
			  }else{
			   $("#stemplate1").hide();
			   $("#stemplate2").show();
			  }
		   }
		   function getFname(v){
		      if (v==1){
			   $("#sfname").hide();
			  }else{
			   $("#sfname").show();
			  }
		   }
		   
		   function getUnit(v){
		      if (v==1){
			   $("#sunit").hide();
			  }else{
			   $("#sunit").show();
			  }
		   }
		   function getPoint(v){
		      if (v==1){
			   $("#spoint").hide();
			  }else{
			   $("#spoint").show();
			  }
		   }
		   
			</script>
		   <%
			.echo "</head>"
			.echo "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
		  
		  select case request("action")
		   case "Step2" Step2
		   case "Step3" Step3
		   case else
		     call step1
		  end select
		  	.echo "</body>"
			.echo "</html>"
        End With
	   End Sub
	   
	   
	   '
	   Sub Step1()
		 With KS
			.echo "      <div class='topdashed sort'>"
			.echo "      第一步 数据批量导入主数据设置"
			.echo "      </div>"
			.echo "<form action=""?Action=Step2"" method=""post"" name=""DownParamForm"">"
			.echo "  <table width=""100%"" border=""0"" align=""center"" cellspacing=""1"" class=""ctable"">"
			.echo "    <tr class='tdbg'>"
			.echo "      <td width=""150"" height=""30"" class='clefttitle' align='right'><strong>要导入的模型</strong></td>"
			.echo "      <td><select id='channelid' name='channelid'>"
			.echo " <option value='0'>---请选择目标模型---</option>"
	
			If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,Node
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1 and (@ks6=3 or @ks6=1 or @ks6=5)]")
			   .echo "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "(" & Node.SelectSingleNode("@ks2").text & ")</option>"
			next
			.echo "</select>"			
			.echo "     </td>"
			.echo "    </tr>"
			.echo "    <tr class='tdbg'>"
			.echo "      <td width=""150"" height=""30"" class='clefttitle' align='right'><strong>数据源类型</strong></td>"
            .echo "      <td><select name=""datasourcetype"" id=""datasourcetype"" onchange=""datachanage()""><option value='1'>access</option><option value='2'>Excel</option><option value='3'>MS SQL</option></select></td>"
			.echo "    </tr>"
			.echo "    <tr class='tdbg'>"
			.echo "      <td width=""150"" height=""30"" class='clefttitle' align='right'><strong>连接字符串</strong></td>"
            .echo "      <td><textarea name='datasourcestr' id='datasourcestr' cols='70' rows='3'>Provider=Microsoft.Jet.OLEDB.4.0;Data Source=/数据库.mdb</textarea>"
			.echo "     &nbsp;<input class='button' id='testbutton' name='testbutton' type='button' value='测试' onclick='testsource();'><br><font color=green>说明:Access/Excel数据源支持相对路径,如Provider=Microsoft.Jet.OLEDB.4.0;Data Source=/1.mdb,表示连接根目录下的1.mdb数据库</font></td>"
			.echo "    </tr>"
			.echo "    <tr class='tdbg'>"
			.echo "      <td width=""150"" height=""30"" class='clefttitle' align='right'><strong>数据表名称</strong></td>"
            .echo "      <td><input type='text' name='tablename' id='tablename' value='Table1' /></td>"
			.echo "    </tr>"
			.echo "  </table>"
			.echo " <div style='text-align:center;padding:20px'><input type='submit' onclick='return(checkNext())' value=' 下一步 ' class='button' name='button1'></div>"
			.echo "</form>"
			End With
		End Sub
		
		Sub testsource()
			response.cachecontrol="no-cache"
			response.addHeader "pragma","no-cache"
			response.expires=-1
			response.expiresAbsolute=now-1
			Response.CharSet="utf-8"
			on error resume next
		   dim str:str=unescape(request("str"))
		   If KS.G("DataType")="1" or KS.G("DataType")="2" Then str=LFCls.GetAbsolutePath(str)
		   dim tconn:Set tconn = Server.CreateObject("ADODB.Connection")
			tconn.open str
			If Err Then 
			  Err.Clear
			  Set tconn = Nothing
			  KS.Echo "false"
			else
			  KS.Echo "true"
			end if
		End Sub
		
		Sub OpenImporIConn()
				   if not isobject(IConn) then
					on error resume next
					Set IConn = Server.CreateObject("ADODB.Connection")
					IConn.open IConnStr
					If Err Then 
					  Err.Clear
					  Set IConn = Nothing
					  Response.Write "<script>alert('数据源连接失败,请检查数据库连接!');history.back();</script>"
					  response.end
					end if
				   end if		
		End Sub
       '**************************************************
		'过程名：ShowChird
		'作  用：显示指定数据表的字段列表
		'参  数：无
		'**************************************************
		Function ShowField(fieldname)
				if request("tablename")="" then
				 response.write "<script>alert('表名称必须输入！');history.back();</script>"
				 response.end
				end if
				dim dbname:dbname=request("tablename")
				if tempField="" Then
					dim rs:Set rs=Iconn.OpenSchema(4)
					if request("datasourcetype")<>"2" then
					'Do Until rs.EOF or rs("Table_name") = trim(dbname)
					'	rs.MoveNext
					'Loop
					end if
					'Do Until rs.EOF or rs("Table_name") <> trim(dbname)
					Do Until rs.EOF
					  tempField=tempField & "<option value='"&lcase(rs("column_Name"))&"'>·"&rs("column_Name")&"</option>"
					  rs.MoveNext
					loop
				    rs.close:set rs=nothing
			   End If
			   ShowField=replace(tempField,"value='" & lcase(fieldname) & "'","value='" & lcase(fieldname) & "' selected")
		End Function	
		
		
		Sub Step2()
		   ChannelID=KS.ChkClng(Request("ChannelID"))
		   If ChannelID=0 Then 
		     KS.AlertHintScript "请选择要导入的模型!"
		   End If
		   With KS
			.echo "      <div class='topdashed sort'>"
			.echo "      第二步 数据批量导入字段设置"
			.echo "      </div>"
			IConnStr=Request("datasourcestr")
			If KS.G("datasourcetype")="1" or KS.G("datasourcetype")="2" Then IConnStr=LFCls.GetAbsolutePath(IConnStr)
			if IConnStr="" Then
			  KS.AlertHintScript "请输入连接字符串!"
			End If
			
			
			OpenImporIConn()
			.echo "<table width='100%' style='margin-top:10px' border='0' align='center'  cellspacing='1' class='ctable'>"
			.echo "<form name='myform' id='myform' action='KS.Import.asp?action=Step3' method='post'>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='Title'><option value='0'>-此项必选-</option>"
			.echo ShowField("title")
			.echo "	</select> =>	</td>"
			.echo "	<td>" & KS.C_S(ChannelID,3) 
			If KS.C_S(ChanneliD,6)=5 Then .echo "名称" Else .echo "标题"
			.echo "(Title)* <font color=#999999>此字段如果值为空将直接跳过不导入</font>"
			.echo "<br/><strong>重复处理:</strong><label><input type='checkbox' name='titlerepet' value='1' onclick=""if (this.checked){alert('您选中了标题重复跳过不导入!');}"">遇该字段值有重复时不导入</label>"
			.echo "</td></tr>"
			
			If KS.C_S(ChanneliD,6)=5 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='ProID'><option value='0'>-自动生成商品ID-</option>"
			.echo ShowField("ProID")
			.echo "	</select> =>	</td>"
			.echo "	<td>商品ID(ProID)* <font color=#999999>保证源数据是不重复，否则请选择自动生成</font></td></tr>"
			End If
			
			If KS.C_S(ChanneliD,6)=1 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='FullTitle'><option value='0'>-此项不导入-</option>"
			.echo ShowField("fulltitle")
			.echo "	</select> =>	</td>"
			.echo "	<td>完整标题(FullTitle)*</td></tr>"
			End If
			
			'===================================栏目ID=====================================
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>所属栏目:</td><td><label><input type='radio' value='1' name='tidtype' onclick=""getClass(1)"" checked/>直接导入指定的栏目</label> <br/><label><input type='radio' onclick=""getClass(2)"" name='tidtype' value='2'>读取数据源的栏目ID</label>"
			.echo "	</td></tr>"
			
			.echo "<tr class='tdbg' id='stid1'><td height='25' align='right' class='clefttitle'></td><td><select size='1' name='tid1' id='tid1' style='width:160px'>"
			.echo " <option value='0'>--请选择栏目--</option>"
			.echo KS.LoadClassOption(ChannelID,false)& " </select> =>栏目ID(Tid)*</td></tr>"
			
			.echo "<tr class='tdbg' id='stid2' style='display:none'><td height='25' align='right' class='clefttitle'></td><td><select name='tid2'><option value='0'>-此项不导入-</option>"
			.echo ShowField("tid")
			.echo "	</select> =>栏目ID(Tid)*</td></tr>"
			'=================================================================================
			
			'==================================模板ID=======================================================
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>绑定模板:</td><td><label><input type='radio' value='1' name='templatetype' onclick=""getTemplate(1)"" checked/>选择模板并绑定</label> <br/><label><input type='radio' onclick=""getTemplate(2)"" name='templatetype' value='2'>读取数据源的模板</label>"
			.echo "	</td></tr>"
			.echo "<tr class='tdbg' id='stemplate1'><td height='25' align='right' class='clefttitle'></td><td><input id='TemplateID' name='TemplateID' readonly size=20 class='textbox' value='{@TemplateDir}/"& KS.C_S(ChannelID,1) & "/内容页.html'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]") &"  =>模板(TemplateID)*</td></tr>"
			
			.echo "<tr class='tdbg' id='stemplate2' style='display:none'><td height='25' align='right' class='clefttitle'></td><td><select name='templateid2'><option value='0'>-此项不导入-</option>"
			.echo ShowField("templateid")
			.echo "	</select> =>模板(TemplateID)*</td></tr>"
			'=================================================================================================
			
			'==================================文件名=======================================================
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>文件名:</td><td><label><input type='radio' value='1' name='Fnametype' onclick=""getFname(1)"" checked/>自动生成</label> <br/><label><input type='radio' onclick=""getFname(2)"" name='Fnametype' value='2'>读取数据源的文件名</label>"
			.echo "	</td></tr>"
			.echo "<tr class='tdbg' id='sfname' style='display:none'><td height='25' align='right' class='clefttitle'></td><td><select name='Fname'><option value='0'>-此项不导入-</option>"
			.echo ShowField("Fname")
			.echo "	</select> =>文件名(Fname)*</td></tr>"
			'=================================================================================================
			
			
			If KS.C_S(ChanneliD,6)=5 Then
			
			'==================================商品单位=======================================================
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>商品单位:</td><td><label><input type='radio' value='1' name='Unittype' onclick=""getUnit(1)"" checked/>手工指定单位</label> &nbsp;&nbsp;&nbsp;<span id='sunit1'><input type='text' name='myunit' value='件' size='5'> 如：件，套，个等</span> <br/><label><input type='radio' onclick=""getUnit(2)"" name='Unittype' value='2'>读取数据源的单位</label>"
			.echo "	</td></tr>"
			.echo "<tr class='tdbg' id='sunit' style='display:none'><td height='25' align='right' class='clefttitle'></td><td><select name='Unit'><option value='0'>-此项不导入-</option>"
			.echo ShowField("Unit")
			.echo "	</select> =>商品单位(Unit)</td></tr>"
			'=================================================================================================
			'==================================积分=======================================================
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>获赠积分:</td><td><label><input type='radio' value='1' name='Pointtype' onclick=""getPoint(1)"" checked/>手工指定获赠积分</label> &nbsp;&nbsp;&nbsp;<span id='sunit1'><input type='text' name='myPoint' value='-1' size='5'> 输入-1，将自动设置会员价为获赠积分数</span> <br/><label><input type='radio' onclick=""getPoint(2)"" name='Pointtype' value='2'>读取数据源的单位</label>"
			.echo "	</td></tr>"
			.echo "<tr class='tdbg' id='spoint' style='display:none'><td height='25' align='right' class='clefttitle'></td><td><select name='Point'><option value='0'>-此项不导入-</option>"
			.echo ShowField("Point")
			.echo "	</select> =>获赠积分(Point)</td></tr>"
			'=================================================================================================
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='ProIntro'><option value='0'>-此项不导入-</option>"
			.echo ShowField("ProIntro")
			.echo "	</select> =>	</td>"
			.echo "	<td>商品介绍(ProIntro)*</td></tr>"
			

			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='Price'><option value='0'>-此项不导入-</option>"
			.echo ShowField("Price")
			.echo "	</select> =>	</td>"
			.echo "	<td>参考价(Price)*</td></tr>"

			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='Price_Member'><option value='0'>-此项不导入-</option>"
			.echo ShowField("Price_Member")
			.echo "	</select> =>	</td>"
			.echo "	<td>商城价(Price_Member)*</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='Weight'><option value='0'>-此项不导入-</option>"
			.echo ShowField("Weight")
			.echo "	</select> =>	</td>"
			.echo "	<td>商品重量(Weight)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='TotalNum'><option value='0'>-此项不导入-</option>"
			.echo ShowField("TotalNum")
			.echo "	</select> =>	</td>"
			.echo "	<td>库存数量(TotalNum)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='AlarmNum'><option value='0'>-此项不导入-</option>"
			.echo ShowField("AlarmNum")
			.echo "	</select> =>	</td>"
			.echo "	<td>库存报警下限数(AlarmNum)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='ProModel'><option value='0'>-此项不导入-</option>"
			.echo ShowField("ProModel")
			.echo "	</select> =>	</td>"
			.echo "	<td>商品型号(ProModel)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='ProSpecificat'><option value='0'>-此项不导入-</option>"
			.echo ShowField("ProSpecificat")
			.echo "	</select> =>	</td>"
			.echo "	<td>商品规格(ProSpecificat)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='ProducerName'><option value='0'>-此项不导入-</option>"
			.echo ShowField("ProducerName")
			.echo "	</select> =>	</td>"
			.echo "	<td>生 产 商(ProducerName)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='TrademarkName'><option value='0'>-此项不导入-</option>"
			.echo ShowField("TrademarkName")
			.echo "	</select> =>	</td>"
			.echo "	<td>商品商标(TrademarkName)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='ServiceTerm'><option value='0'>-此项不导入-</option>"
			.echo ShowField("ServiceTerm")
			.echo "	</select> =>	</td>"
			.echo "	<td>服务期限(ServiceTerm)</td></tr>"
			End If
			
			
			
			If KS.C_S(ChanneliD,6)=3 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downlb'><option value='0'>-此项不导入-</option>"
			.echo ShowField("downlb")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "类别(DownLB)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downyy'><option value='0'>-此项不导入-</option>"
			.echo ShowField("downyy")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "语言(DownYY)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downsq'><option value='0'>-此项不导入-</option>"
			.echo ShowField("downsq")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "授权(DownSQ)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downsize'><option value='0'>-此项不导入-</option>"
			.echo ShowField("downsize")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "大小(DownSize)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downpt'><option value='0'>-此项不导入-</option>"
			.echo ShowField("downpt")
			.echo "	</select> =>	</td>"
			.echo "	<td>系统平台(DownPT)</td></tr>"
			
			End If
			
			If KS.C_S(ChanneliD,6)<>5 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='photourl'><option value='0'>-此项不导入-</option>"
			.echo ShowField("photourl")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "图片(PhotoUrl)</td></tr>"
		    Else
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='photourl'><option value='0'>-此项不导入-</option>"
			.echo ShowField("photourl")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "小图(PhotoUrl)"
			
			.Echo "<br/><strong>是否淘宝图片：</strong><label><input onclick=""$('#picpath1').show()"" checked type='radio' name='pictype1' value='1'>是</label> <label><input onclick=""$('#picpath1').hide()"" type='radio' name='pictype1' value='0'>否</label>"
			.Echo "<div id='picpath1'>图片路径<input type='text' name='mypicpath1'  value='" & KS.Setting(3) & "images/" & Year(Now) & "-" & Month(Now) & "/'> <br/><font color=#999999>请先将原tbi图片复制到以上文件夹下,执行导入的同时系统会自动重命名为jpg格式</font></div>"

			.echo "</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='BigPhoto'><option value='0'>-此项不导入-</option>"
			.echo ShowField("BigPhoto")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "大图(BigPhoto)"
			.Echo "<br/><strong>是否淘宝图片：</strong><label><input onclick=""$('#picpath').show()"" checked type='radio' name='pictype' value='1'>是</label> <label><input onclick=""$('#picpath').hide()"" type='radio' name='pictype' value='0'>否</label>"
			.Echo "<div id='picpath'>图片路径<input type='text' name='mypicpath'  value='" & KS.Setting(3) & "images/" & Year(Now) & "-" & Month(Now) & "/'> <br/><font color=#999999>请先将原tbi图片复制到以上文件夹下,执行导入的同时系统会自动重命名为jpg格式</font></div>"
			.Echo "</td></tr>"
			End If
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='keywords'><option value='0'>-此项不导入-</option>"
			.echo ShowField("keywords")
			.echo "	</select> =>	</td>"
			.echo "	<td>关键字(KeyWords)</td></tr>"
			
		If KS.C_S(ChanneliD,6)<>5 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='author'><option value='0'>-此项不导入-</option>"
			.echo ShowField("author")
			.echo "	</select> =>	</td>"
			If KS.C_S(ChannelID,6)=3 Then
			.echo "	<td>作者开发商(Author)</td></tr>"
			Else
			.echo " <td>" &KS.C_S(ChannelID,3) & "作者(Author)</td></tr>"
			End If
		End If
		
		If KS.C_S(ChanneliD,6)<>5 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='Origin'><option value='0'>-此项不导入-</option>"
			.echo ShowField("Origin")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "来源(Origin)</td></tr>"
	    End If
		If KS.C_S(ChannelID,6)=3 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downurls'><option value='0'>-此项不导入-</option>"
			.echo ShowField("downurls")
			.echo "	</select> =>	</td>"
			.echo "	<td>下载地址(DownUrls)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downcontent'><option value='0'>-此项不导入-</option>"
			.echo ShowField("downcontent")
			.echo "	</select> =>	</td>"
			.echo "	<td>软件介绍(DownContent)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='ysdz'><option value='0'>-此项不导入-</option>"
			.echo ShowField("ysdz")
			.echo "	</select> =>	</td>"
			.echo "	<td>演示地址(YSDZ)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='zcdz'><option value='0'>-此项不导入-</option>"
			.echo ShowField("zcdz")
			.echo "	</select> =>	</td>"
			.echo "	<td>注册地址(ZCDZ)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='jymm'><option value='0'>-此项不导入-</option>"
			.echo ShowField("JYMM")
			.echo "	</select> =>	</td>"
			.echo "	<td>解压密码(JYMM)</td></tr>"
		  ElseIf KS.C_S(ChannelID,6)=1 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='intro'><option value='0'>-此项不导入-</option>"
			.echo ShowField("intro")
			.echo "	</select> =>	</td>"
			.echo "	<td>" & KS.C_S(ChannelID,3) & "简介(Intro)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='articlecontent'><option value='0'>-此项不导入-</option>"
			.echo ShowField("articlecontent")
			.echo "	</select> =>	</td>"
			.echo "	<td>" & KS.C_S(ChannelID,3) & "内容(ArticleContent)</td></tr>"
		  End If			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='adddate'><option value='0'>-此项不导入-</option>"
			.echo ShowField("adddate")
			.echo "	</select> =>	</td>"
			.echo "	<td>添加日期(AddDate)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='inputer'><option value='0'>-此项不导入-</option>"
			.echo ShowField("inputer")
			.echo "	</select> =>	</td>"
			.echo "	<td>" & KS.C_S(ChannelID,3) & "录入(Inputer)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='rank'><option value='0'>-此项不导入-</option>"
			.echo ShowField("rank")
			.echo "	</select> =>	</td>"
			.echo "	<td>" & KS.C_S(ChannelID,3) & "等级(Rank)</td></tr>"
			
			Dim FieldXML,FieldNode,FNode
			Call KSCls.LoadModelField(ChannelID,FieldXML,FieldNode)
			For Each FNode In FieldNode
	         If KS.ChkClng(FNode.SelectSingleNode("fieldtype").text)<>0 Then
				.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
				.echo "		<select name='" & FNode.SelectSingleNode("@fieldname").text & "'><option value='0'>-此项不导入-</option>"
				.echo ShowField(FNode.SelectSingleNode("@fieldname").text)
				.echo "	</select> =>	</td>"
				.echo "	<td>" & FNode.SelectSingleNode("title").text & "(" & FNode.SelectSingleNode("@fieldname").text & ")</td></tr>"
		     End If
			Next
			
			.echo "<tr class='tdbg'><td height='35' colspan='2' class='clefttitle' style=""text-align:left""><strong>说明:</strong><br/>1.建议按各个模型的主数据表结构制作数据源数据库文件，如文章模型对应的数据表结构请参考KS_Article;<br/>2.所属栏目如果是选择“读取数据源的栏目ID”，那么请确保您整理的栏目ID是已存在的栏目（即在后台的栏目管理可以看到的栏目）； </td></tr>"


			.echo "</table>"
			.echo "<input type='hidden' name='channelid' value='" & channelid & "'/>"
			.echo "<input type='hidden' name='datasourcetype' value='" & request("datasourcetype") & "'/>"
			.echo "<input type='hidden' name='datasourcestr' value='" & request("datasourcestr") & "'/>"
			.echo "<input type='hidden' name='tablename' value='" & request("tablename") & "'/>"
			
			.echo "<div style='padding:10px;text-align:center'><input type='submit' onclick=""return(confirm('请认真检查各导入项，确定无误后再点击确认！'))"" value=' 下 一 步 ' class='button'</div>"
			.echo "</form>"
           End With
		End Sub
		
		'步骤三
		Sub Step3()
		  %>
		  <div class='topdashed sort'>第三步 数据批量导入执行页面</div>
		
		<div style="text-align:center">			 
			 <div style="margin-top:50px;border:1px dashed #cccccc;width:500px;height:80px">
			 <br>
			<div id="message">
			  <br>操作提示栏！
			</div>
			</div>
	    </div>
		<br/><br/><br/>
		  <%
		   ChannelID=KS.ChkClng(Request("ChannelID"))
		   If ChannelID=0 Then 
		     KS.AlertHintScript "请选择要导入的模型!"
		   End If

		  IF KS.G("Title")="0" Then
		   KS.Echo "<script>alert('" & KS.C_S(ChannelID,3) & "名称选项必须选择!');history.back();</script>"
		   response.end
		  End If
		  
		  If KS.G("tidtype")="2" And KS.G("Tid2")="0" Then
		   Call KS.AlertHistory("栏目选项必须选择",-1)
		   response.end
		  ElseIf KS.G("Tidtype")="1" and KS.G("Tid1")="0" Then
		   Call KS.AlertHistory("所属栏目必须选择",-1)
		   response.end
		  End If
		  
		  If KS.G("templatetype")="2" And KS.G("templateid2")="0" Then
		   Call KS.AlertHistory("模板选项必须选择",-1)
		   response.end
		  End If
		  
		  Server.ScriptTimeOut=999999
			IConnStr=Request("datasourcestr")
			If KS.G("datasourcetype")="1" or KS.G("datasourcetype")="2" Then IConnStr=LFCls.GetAbsolutePath(IConnStr)
			if IConnStr="" Then
			  KS.AlertHintScript "请输入连接字符串!"
			End If
			OpenImporIConn()
		 Dim TableName:TableName=Request("TableName")
		 Dim Total,n,i,msg,errnum,t,Intro,succnum
		 Dim IRS:Set IRS=Server.CreateOBject("ADODB.RECORDSET")
    	 Dim RS:Set RS=Server.CreateObject("ADODB.RecordSet")
		 IRS.Open "Select * From [" & TableName & "]",iConn,1,1
		 'Dim IRS:Set IRS=iConn.Execute("Select * From [" & TableName & "]")
		 Total=IRS.RecordCount
		 If Total<=0 Then Total=IConn.Execute("Select count(*) From [" & TableName & "]")(0)
		 n=0:t=0:errnum=0:succnum=0
		 For I=0 To Total
			 n=n+1
           if KS.IsNul(IRS(KS.G("Title"))) Then
		    ErrNum=ErrNum+1
			Msg=msg & "源数据库记录没有标题，所以跳过!<br/>"
		   ElseIf Instr(IRS(KS.G("Title")),"<")<>0 or Instr(IRS(KS.G("Title")),">")<>0 or Instr(IRS(KS.G("Title")),"'")<>0 Then
		    ErrNum=ErrNum+1
			Msg=msg & "标题有特殊字符跳过!<br/>"
		   Else
				 Dim Tid,TemplateID,ProID
				 If KS.G("tidtype")="1" Then 
				   Tid=KS.G("Tid1")
				 Else
				   Tid=IRS(KS.G("Tid2"))
				 End If
				 If KS.G("templatetype")="1" Then
				    TemplateID=KS.G("TemplateID")
				 Else
				    TemplateID=IRS(KS.G("templateid2"))
				 End If
				 If KS.C_S(ChanneliD,6)=5 Then
				   ProID=KS.GetInfoID(ChannelID)
				 End If
				 Dim Param
				 IF KS.G("titlerepet")="1" Then
				   Param="Where [Title]='" &KS.LoseHtml(IRS(KS.G("Title"))) & "' and tid='" & tid & "'"
				 Else
				   Param="Where 1=0"
				 End If
				 RS.Open "Select top 1 * From [" & KS.C_S(ChannelID,2) & "] " &Param,conn,1,3
				 If RS.Eof and RS.Bof Then
				   RS.AddNew
				   RS("Title")=Left(KS.LoseHtml(IRS(KS.G("Title"))),200)
				   RS("Tid")=Tid
				   RS("TemplateID")=TemplateID
				   If KS.G("Fnametype")="2" Then
				    RS("Fname")=IRS(KS.G("Fname"))
				   End If
				   
				   '商城系统
				   If KS.C_S(ChanneliD,6)=5 Then
				        RS("BrandID")=0
				     If KS.G("ProID")<>"0" Then
						RS("ProID")=IRS(KS.G("ProID"))
					 Else
					    RS("ProID")=ProID
					 End If
					 If KS.G("ProIntro")<>"0" Then
					    RS("ProIntro")=IRS(KS.G("ProIntro")) & " "
					 Else
					    RS("ProIntro")=" "
					 End If
					 Intro=KS.LoseHtml(RS("ProIntro"))
					 
					
					 If KS.G("Price")<>"0" Then
					   If IsNumeric(IRS(KS.G("Price"))) Then
					    RS("Price")=IRS(KS.G("Price"))
					   Else
					    RS("Price")=0
					   End If
					 Else
					   RS("Price")=0
					 End If
					
					 If KS.G("Price_Member")<>"0" Then
					   If IsNumeric(IRS(KS.G("Price_Member"))) Then
					    RS("Price_Member")=IRS(KS.G("Price_Member"))
					   Else
					    RS("Price_Member")=0
					   End If
					 Else
					   RS("Price_Member")=0
					 End If
					 
					 If KS.G("Pointtype")="1" Then
					   If KS.G("MyPoint")<>"-1" Then
					     RS("Point")=KS.ChkClng(KS.G("MyPoint"))
					   Else
					     RS("Point")=KS.ChkClng(RS("Price_Member"))
					   End If
					 Else
					     RS("Point")=KS.ChkClng(IRS(KS.G("Point")))
					 End If
					 
					 If KS.G("Unittype")="1" Then
					     RS("Unit")=KS.G("MyUnit")
					 Else
					     RS("Unit")=IRS(trim(KS.G("Unit")))
					 End If
					 
					 
					 If KS.G("Weight")<>"0" Then
					   If IsNumeric(IRS(KS.G("Weight"))) Then
					    RS("Weight")=IRS(KS.G("Weight"))
					   Else
					    RS("Weight")=0
					   End If
					 Else
					   RS("Weight")=0
					 End If
					 If KS.G("TotalNum")<>"0" Then
					   If IsNumeric(IRS(KS.G("TotalNum"))) Then
					    RS("TotalNum")=IRS(KS.G("TotalNum"))
					   Else
					    RS("TotalNum")=1000
					   End If
					 Else
					   RS("TotalNum")=1000
					 End If
					 If KS.G("AlarmNum")<>"0" Then
					   If IsNumeric(IRS(KS.G("AlarmNum"))) Then
					    RS("AlarmNum")=IRS(KS.G("AlarmNum"))
					   Else
					    RS("AlarmNum")=10
					   End If
					 Else
					   RS("AlarmNum")=10
					 End If
					 If KS.G("ProModel")<>"0" Then
					    RS("ProModel")=Left(IRS(KS.G("ProModel")) & " ",200)
					 Else
					    RS("ProModel")=""
					 End If
					 If KS.G("ProSpecificat")<>"0" Then
					    RS("ProSpecificat")=Left(IRS(KS.G("ProSpecificat")) & " ",50)
					 Else
					    RS("ProSpecificat")=""
					 End If
					 If KS.G("ProducerName")<>"0" Then
					    RS("ProducerName")=Left(IRS(KS.G("ProducerName")) & " ",50)
					 Else
					    RS("ProducerName")=""
					 End If
					 If KS.G("TrademarkName")<>"0" Then
					    RS("TrademarkName")=Left(IRS(KS.G("TrademarkName")) & " ",50)
					 Else
					    RS("TrademarkName")=""
					 End If
					 If KS.G("ServiceTerm")<>"0" Then
					   If IsNumeric(IRS(KS.G("ServiceTerm"))) Then
					    RS("ServiceTerm")=IRS(KS.G("ServiceTerm"))
					   Else
					    RS("ServiceTerm")=0
					   End If
					 Else
					   RS("ServiceTerm")=0
					 End If
					 
					   If KS.G("BigPhoto")<>"0" Then
						   RS("BigPhoto")=Left(IRS(KS.G("BigPhoto")),255)
						   If KS.G("pictype")="1"  And Not KS.IsNUL(RS("BigPhoto")) Then
						     Dim PicPath,PhysicalPath,FsoObj,FileObj
							 PicPath=KS.G("mypicpath")
							 
							 Set FsoObj = KS.InitialObject(KS.Setting(99))
							 If (RS("BigPhoto")<> "") Then
								PhysicalPath = Server.MapPath(PicPath) & "\" & Split(RS("BigPhoto"),":")(0) & ".tbi"
								If FsoObj.FileExists(PhysicalPath) = True Then
									PhysicalPath = Server.MapPath(PicPath) & "\" & Split(RS("BigPhoto"),":")(0)&".jpg"
									If FsoObj.FileExists(PhysicalPath) = False Then
										Set FileObj = FsoObj.GetFile(Server.MapPath(PicPath) & "\" & Split(RS("BigPhoto"),":")(0) & ".tbi")
										FileObj.name = Split(RS("BigPhoto"),":")(0)&".jpg"
										Set FileObj = Nothing
									End If
								End If
							End If
							  RS("BigPhoto")=PicPath & Split(RS("BigPhoto"),":")(0)&".jpg"
						   End If
					   Else
						RS("BigPhoto")=""
					   End If

				   End If
				   
				If KS.C_S(Channelid,6)=1 Then
					   If KS.G("Intro")<>"0" Then
						RS("Intro")=IRS(KS.G("Intro"))
					   End If
					   If KS.G("ArticleContent")<>"0" Then
					    If KS.IsNUL(IRS(KS.G("ArticleContent"))) Then
						 RS("ArticleContent")=" "
						Else
					     RS("ArticleContent")=IRS(KS.G("ArticleContent"))
						End If
					   Else
					     RS("ArticleContent")=" "
					   End If
					   
					   If KS.G("FullTitle")<>"0" Then
					    RS("FullTitle")=IRS(KS.G("FullTitle"))
					   End If
					   Intro=RS("Intro")
				ElseIf KS.C_S(Channelid,6)=3 Then   '下载
				   If KS.G("DownPT")<>"0" Then
				    RS("DownPT")=IRS(KS.G("DownPT"))
				   End If
				   If KS.G("DownUrls")<>"0" Then
				    RS("DownUrls")=IRS(KS.G("DownUrls"))
				   End If
				   If KS.G("DownContent")<>"0" Then
				    RS("DownContent")=IRS(KS.G("DownContent"))
				   Else
				    RS("DownContent")=" "
				   End If
				   If KS.G("YSDZ")<>"0" Then
				    RS("YSDZ")=IRS(KS.G("YSDZ"))
				   End If
				   If KS.G("DownLB")<>"0" Then
				    RS("DownLB")=IRS(KS.G("DownLB"))
				   End If
				   If KS.G("DownYY")<>"0" Then
				    RS("DownYY")=IRS(KS.G("DownYY"))
				   End If
				   If KS.G("DownSQ")<>"0" Then
				    RS("DownSQ")=IRS(KS.G("DownSQ"))
				   End If
				   If KS.G("DownSize")<>"0" Then
				    RS("DownSize")=IRS(KS.G("DownSize"))
				   End If
				   If KS.G("ZCDZ")<>"0" Then
				    RS("ZCDZ")=IRS(KS.G("ZCDZ"))
				   End If
				   If KS.G("JYMM")<>"0" Then
				    RS("JYMM")=IRS(KS.G("JYMM"))
				   End If
				    Intro=RS("DownContent")
				End If

               If KS.C_S(ChanneliD,6)<>5 Then  
				   If KS.G("PhotoUrl")<>"0" Then
				    RS("PhotoUrl")=IRS(KS.G("PhotoUrl"))
				   End If
			   Else
				   If KS.G("PhotoUrl")<>"0" Then
						   RS("PhotoUrl")=Left(IRS(KS.G("PhotoUrl")),255)
						   If KS.G("pictype")="1" And Not KS.IsNUL(RS("PhotoUrl")) Then
							 PicPath=KS.G("mypicpath1")
							 
							 Set FsoObj = KS.InitialObject(KS.Setting(99))
							 If (RS("PhotoUrl")<> "") Then
								PhysicalPath = Server.MapPath(PicPath) & "\" & Split(RS("PhotoUrl"),":")(0) & ".tbi"
								If FsoObj.FileExists(PhysicalPath) = True Then
									PhysicalPath = Server.MapPath(PicPath) & "\" & Split(RS("PhotoUrl"),":")(0)&".jpg"
									If FsoObj.FileExists(PhysicalPath) = False Then
										Set FileObj = FsoObj.GetFile(Server.MapPath(PicPath) & "\" & Split(RS("PhotoUrl"),":")(0) & ".tbi")
										FileObj.name = Split(RS("PhotoUrl"),":")(0)&".jpg"
										Set FileObj = Nothing
									End If
								End If
							End If
							  RS("PhotoUrl")=PicPath & Split(RS("PhotoUrl"),":")(0)&".jpg"
						   End If
					   Else
						RS("BigPhoto")=""
					   End If
				End If   
				   
				   If KS.G("KeyWords")<>"0" Then
				    RS("KeyWords")=IRS(KS.G("KeyWords"))
				   End If
				   If KS.IsNUL(RS("KeyWords")) Then
				    RS("KeyWords")=""
				   End If
				   
				 If KS.C_S(ChanneliD,6)<>5 Then  
				   If KS.G("Author")<>"0" Then
				    RS("Author")=IRS(KS.G("Author"))
				   End If
				   If KS.G("Origin")<>"0" Then
				    RS("Origin")=IRS(KS.G("Origin"))
				   End If
				 End If

				   If KS.G("Inputer")<>"0" Then
				    RS("inputer")=IRS(KS.G("Inputer"))
				   Else
				    RS("Inputer")=KS.C("AdminName")
				   End If
				   If KS.G("AddDate")<>"0" Then
				    If IsDate(IRS(KS.G("AddDate"))) Then
				     RS("AddDate")=IRS(KS.G("AddDate"))
					Else
					 RS("AddDate")=Now
					End If
				   Else
				    RS("AddDate")=Now
				   End If
				   If KS.G("Rank")<>"0" Then
				    RS("Rank")=IRS(KS.G("Rank"))
				   End If
				   
				   
				   Dim FieldXML,FieldNode,FNode
					Call KSCls.LoadModelField(ChannelID,FieldXML,FieldNode)
					For Each FNode In FieldNode
					 If KS.ChkClng(FNode.SelectSingleNode("fieldtype").text)<>0 Then
					   IF KS.G(FNode.SelectSingleNode("@fieldname").text)<>"0" Then
					   RS(FNode.SelectSingleNode("@fieldname").text)=IRS(KS.G(FNode.SelectSingleNode("@fieldname").text))
					   End If
					 End If
					Next
				   
				   RS("recommend")=0
				   RS("popular")=0
				   RS("IsTop")=0
				   RS("Rolls")=0
				   RS("slide")=0
				   RS("Strip")=0
				   RS("RefreshTF")=0
				   RS("verific")=1
				   RS("Hits")=0
				   RS("HitsByDay")=0
				   RS("HitsByWeek")=0
				   RS("HitsByMonth")=0
				   RS("LastHitsTime")=now
				   RS.Update
				   RS.MoveLast
				   Dim InfoID:InfoID=RS("ID")
				   If KS.G("Fnametype")="1" Then
				     RS("Fname")=RS("ID") & ".html"
					 RS.Update
				   End If
				   
				   Call LFCls.InserItemInfo(ChannelID,InfoID,RS("Title"),RS("Tid"),Intro,RS("KeyWords"),RS("PhotoUrl"),RS("Inputer"),RS("Verific"),RS("Fname"))
                   succnum=succnum+1
				Else
				 msg=msg & "名称:" & KS.LoseHtml(IRS(KS.G("Title"))) & "已存在，所以跳过<br/>"
				 ErrNum=ErrNum+1
				End If
				RS.Close
		 End If
		  	Response.Write "<script>document.all.message.innerHTML='<br>共<font color=red>" & Total & "</font> 条数据，正在导入第<font color=red>" & n & "</font>条！出错跳过<font color=blue>" & ErrNum & "</font>条!';</script>" &vbcrlf
			Response.Flush
		  IRS.MoveNext
		  If cint(n)>=cint(Total) Then Exit For
		Next
		 IRS.Close:Set IRS=Nothing:Set RS=Nothing
		 Response.Write "<script>document.all.message.innerHTML='<br>恭喜！总<font color=red>" & Total & "</font>条，成功导入 <font color=red>" & succnum & "</font> 条数据！出错 <font color=blue>" & errnum &"</font> 条';</script>"&vbcrlf
		 
		 if msg<>"" then
		   response.write "<strong>以下记录重复没有再导入:</strong><br/><font color=red>" & msg & "</font>"&vbcrlf
		 end if
		End Sub


End Class
%> 
