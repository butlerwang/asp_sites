<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<!--#include file="Label/LabelFunction.asp"-->
<%

Dim KSCls
Set KSCls = New LabelSQLCls
KSCls.Kesion()
Set KSCls = Nothing

Class LabelSQLCls
        Private KS
		Private ActionStr,LabelID, LabelRS, SQLStr, LabelName, Descript, LabelContent, LabelFlag, ParentID,Action, Page, RSCheck, FolderID,FieldParam,SQLType,ItemName,pagenum,dbname1,LabelIntro,PageStyle,note,tconn
		Private datasourcetype,datasourcestr,ajax
		Private KeyWord, SearchType, StartDate, EndDate
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Call KS.DelCahe(KS.SiteSn & "_sqllabellist")
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		KeyWord = Request("KeyWord")
		SearchType = Request("SearchType")
		StartDate = Request("StartDate")
		EndDate = Request("EndDate")
		Action = Request.QueryString("Action")
		Page = Request("Page")
		Dbname1=KS.G("dbname1")
		FolderID = Request.QueryString("FolderID")
		LabelName=Request("LabelName")
		ItemName=Request("ItemName")
		SQLType=Request("SQLType")
		LabelID = Request("LabelId")
		PageStyle=Request("PageStyle")
		Note=Request("Note")
		
		datasourcetype=KS.ChkClng(Request("datasourcetype"))
		datasourcestr=Request("datasourcestr")
		
		IF KS.G("action")="testsource" Then
		  call testsource():exit sub
		ElseIf KS.G("action")="testlabelname" then
		  call testlabelname():exit sub
		end if
		
        Call OpenExtConn()
		With KS
		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		.echo "<title>新建标签</title>"
		.echo "</head>"
		.echo "<link href=""Admin_Style.CSS"" rel=""stylesheet"">"
		.echo "<script language=""JavaScript"" src=""../../ks_inc/Common.js""></script>"
		.echo "<script language=""JavaScript"" src=""../../ks_inc/jquery.js""></script>"
		.echo "<script language=""JavaScript"" src='../../ks_inc/lhgdialog.js'></script>"
		%>
		<script type="text/javascript">
		  function ChangeSqlType(num){
		   if (num==1){$("#pagearea").show()}else{$("#pagearea").hide() }
		  }
	  </script>
		<%
		.echo "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Select Case KS.G("Action")
			 Case "ShowClassID" ShowClassID
			 Case "AddNewSubmit" Call AddLabelSave()
			 Case "EditSubmit" Call EditLabelSave()
			 Case "Step2"  Call Step2()
			 Case "Step1" Call Step1()
			 Case "Edit" Call Step0
			 Case Else Call Step0()
			End Select
		.echo "</body>"
		.echo "</html>"
		End With
	  End Sub
	  
	  sub testlabelname()
	        Dim LabelID:LabelID=request.QueryString("labelid")
			Dim RS:Set RS = Server.CreateObject("Adodb.RecordSet")
			if labelid<>"" then 
			 RS.Open "Select LabelName From [KS_Label] Where id<>'" & labelid & "' and LabelName='" & "{SQL_" & LabelName & "}" & "'", Conn, 1, 1
			else
			 RS.Open "Select LabelName From [KS_Label] Where LabelName='" & "{SQL_" & LabelName & "}" & "'", Conn, 1, 1
			end if
			If Not RS.EOF Then
			 KS.Echo "false"
			Else
			 KS.Echo "true"
			end if
			rs.close:set rs=nothing
	  end sub
	  
	  Sub testsource()
	  on error resume next
	   dim str:str=request("str")
	   If KS.G("DataType")="1" or KS.G("DataType")="5" or KS.G("DataType")="6"  Then str=LFCls.GetAbsolutePath(str)
	   dim tconn:Set tconn = Server.CreateObject("ADODB.Connection")
		tconn.open str
		If Err Then 
		  Err.Clear
		  Set tconn = Nothing
		  KS.Echo "false"
		else
		  KS.Echo "true"
		end if
	  end sub
	  Sub Step0()
	    With KS
		 .echo "<body>"
	 	 .echo " <table width='100%' height='25' border='0' cellpadding='0' cellspacing='1' bgcolor='#efefef' class='sort'>"
		 .echo "       <tr><td><div align='center'><font color='#990000'>"
		 .echo "第一步:为SQL标签建立数据源"
		 .echo "    </font></div></td></tr>"
		 .echo "    </table>"
		 
		If LabelID <> "" Then
		    Dim FieldParamArr
		    ActionStr="Step2"
			Set LabelRS = Server.CreateObject("Adodb.Recordset")
			SQLStr = "SELECT top 1 * FROM [KS_Label] Where ID='" & LabelID & "'"
			LabelRS.Open SQLStr, Conn, 1, 1
			If Not LabelRS.Eof Then
				LabelName = Replace(Replace(LabelRS("LabelName"), "{SQL_", ""), "}", "")
				FolderID=LabelRS("FolderID")
				LabelContent = Server.HTMLEncode(LabelRS("LabelContent"))
				FieldParamArr= Split(LabelRS("Description"),"@@@")
			End IF
			LabelIntro =FieldParamArr(0)
			If Ubound(FieldParamArr)>=1 Then
			FieldParam =FieldParamArr(1)
			SQLType= FieldParamArr(2)
			ItemName=FieldParamArr(3)
			PageStyle=FieldParamArr(4)
			Ajax=FieldParamArr(5)
			datasourcetype=FieldParamArr(6)
			datasourcestr=FieldParamArr(7)
			Note=FieldParamArr(8)
			if datasourcetype<>0 then Call OpenExtConn()

			End If
			LabelRS.Close
		Else
		  ItemName="篇"
		  PageStyle=1
		  ActionStr="Step1"
		  Ajax=0
		  datasourcestr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=数据库.mdb"
		End If
			  %>
	    <script type="text/javascript">
		  function CheckForm()
		  {
		   if ($('#LabelName').val()=='')
		   {
		      alert('请输入标签名称!');
			  $('#LabelName').focus();
			  return;
		   }
		   if ($('#lbtf').val()=='false')
		   {
		      alert('标签名称不可用，请重输!');
			  $('#LabelName').focus();
			  return;
		   }
		   $('#myform').submit();
		  }
		  function changeconnstr()
		  {
		    if ($('#datasourcetype').val()==0)
			{
			  $('#datasourcestr').attr("disabled",true);
			  $('#testbutton').attr("disabled",true);
			  $('#lbt').show();
			 }
			else
			{
			  $('#testbutton').attr("disabled",false);
			  $('#datasourcestr').attr("disabled",false);
			//  $('lbt').style.display='none';
			}
		    switch (parseInt($('#datasourcetype').val()))
		    {
			 case 1:
			  $('#datasourcestr').val('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=数据库.mdb');
			  break;
			 case 2:
			  $('#datasourcestr').val('Provider=Sqloledb; User ID=用户名; Password=密码; Initial Catalog=数据库名称; Data Source =(local);');
			  break;
			 case 3:
		      $('#datasourcestr').val('DSN=数据源名;UID=用户名;PWD=密码');
			  break;
			 case 4:
		      $('#datasourcestr').val('driver={microsoft odbc for oracle};uid=用户名;pwd=密码;server=服务器');
			  break;
			 case 5:
		      $('#datasourcestr').val('driver={microsoft excel driver (*.xls)};dbq=数据库名称');
			  break;
			 case 6:
		      $('#datasourcestr').val('driver={microsoft dbase driver (*.dbf)};dbq=数据库名称');
			  break;
			 case 7:
			  alert('连接mysql数据源,需要服务器支持mysql odbc 3.51 driver数据源');
		      $('#datasourcestr').val('driver={mysql odbc 3.51 driver};server=服务器名称;database=数据库名称;user name=用户名;password=密码;');
			  break;
			}
		
		  }
		  function testlabelname()
		  {
		  var LabelName = $('#LabelName').val();
		  var url = 'LabelSQL.asp';
  		  $.get(url,{action:"testlabelname",labelid:"<%=labelid%>",labelname:LabelName},function(d){
		    if (d=='true')
			  $('#labelmessage').html('<font color=blue>恭喜，可以使用该名称!</font>');
			else
			  $('#labelmessage').html('<font color=red>对不起，该名称不可用，已存在!</font>');
			  $('#lbtf').val(d);
		  });
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
			  var url = 'LabelSQL.asp';
			  $.get(url,{action:"testsource",datatype:datatype,str:str},function(d){
				if (d=='true')
				 alert('恭喜，测试通过!')
				else
				 alert('对不起，字符串连接有误!');
			  });
		 } 
           
		</script>
		<br>
	     <table border='0' cellspacing='1' cellpadding='1' width='95%' align='center' class='ctable'>
		  <form action="?action=<%=ActionStr%>" method="post" id="myform" name="myform">
		   <input name='lbtf' id='lbtf' type='hidden'>
		   <input type='hidden' name='labelid' value='<%=labelid%>'>
		  <tr class='tdbg'>
		    <td class='clefttitle' align='right'><strong>标签名称:</strong></td>
		    <td><input name="LabelName" id="LabelName" value="<%=LabelName%>" onblur='testlabelname()' style="width:200;"> <font color=red>*</font><span id='labelmessage'></span><br>例如标签名称：&quot;推荐文章列表&quot;，则在模板中调用：<font color="#FF0000">&quot;{SQL_推荐文章列表(参数1,参数2...)}&quot;</font>。</td>
		  </tr>
		  <tr class='tdbg'>
		   <td width="80" height="30" class='clefttitle' align='right'><strong>标签目录:</strong></td>
		   <td><%=ReturnLabelFolderTree(FolderID, 5)%><font color=""#FF0000"">请选择标签归属目录，以便日后管理标签</font></td>
		  </tr>
		  <tr class='tdbg'>
		   <td width="80" height="30" class='clefttitle' align='right'><strong>数 据 源:</strong></td>
		   <td>
		     <select name="datasourcetype" id="datasourcetype" style="width:290px" onChange="changeconnstr()">
			   <option value="0"<%if datasourcetype=0 then .echo " selected"%>>KesionCMS主数据库</option>
			   <option value="1"<%if datasourcetype=1 then .echo  " selected"%>>Access数据源</option>
			   <option value="2"<%if datasourcetype=2 then .echo  " selected"%>>MS SQL数据源</option>
			   <option value="3"<%if datasourcetype=3 then .echo  " selected"%>>ODBC数据源</option>
			   <option value="4"<%if datasourcetype=4 then .echo  " selected"%>>Oracle数据源</option>
			   <option value="5"<%if datasourcetype=5 then .echo  " selected"%>>Excel数据源</option>
			   <option value="6"<%if datasourcetype=6 then .echo  " selected"%>>Dbase数据源</option>
			   <option value="7"<%if datasourcetype=7 then .echo  " selected"%>>MYSQL数据源(需支持mysql odbc 3.51 driver)</option>
			 </select>
		   </td>
		  </tr>
		  <tr class='tdbg'>
		   <td width="80" height="25" class='clefttitle' align='right'><strong>连接字符串:</strong></td>
		   <td><textarea <%if datasourcetype=0 then .echo  " disabled"%> name="datasourcestr" id="datasourcestr" cols="70" rows="3"><%=Datasourcestr%></textarea>
		     &nbsp;<input class='button' id="testbutton" name="testbutton" <%if datasourcetype=0 then .echo " disabled"%> type='button' value='测试' onclick='testsource();'>
			 <br><font color=green>说明:外部Access数据源支持相对路径,如Provider=Microsoft.Jet.OLEDB.4.0;Data Source=/数据库.mdb</font>
		   </td>
		  </tr>
		  
		  <%If LabelID <> "" Then%>
		  <tr class='tdbg'>
		   <td width="80" height="25" class='clefttitle' align='right'><strong>查询语句:</strong></td>
		   <td><textarea name="LabelIntro" cols="80" style="width:98%" rows="4"><%=LabelIntro%></textarea>
		   </td>
		  </tr>
		  <%End If%>
		  
		  <tr class='tdbg'>
		   <td width="80" height="30" class='clefttitle' align='right'><strong>Ajax调用:</strong></td>
		   <td><input type='radio' value='1' name='ajax'<%if ajax=1 then .echo  " checked"%>>是&nbsp;<input type='radio' value='0' name='ajax'<%if ajax=0 then .echo  " checked"%>>否
		   </td>
		  </tr>
		  <tr class='tdbg' id='lbt'>
		   <td width="80" height="45" class='clefttitle' align='right'><strong>标签类型:</strong></td>
		   <td>

		    <input type="radio" name="SQLType" value="0" <%if sqltype=0 then .echo  " checked"%> onclick='ChangeSqlType(this.value);'>普通标签  
			<input type="radio" name="SQLType" value="1"<%if sqltype=1 then .echo  " checked"%> onclick='ChangeSqlType(this.value);'>终级分页标签<font color=red>(内外部数据库均适用，一个页面只能放一个分页标签)</font>
			
			<table border='0' id='pagearea' <%if sqltype=0 then .echo  " style=display:none"%>>
			 <tr><td>分页项目单位：<input type="text" value="<%=itemname%>" class="textbox" name="ItemName" size="6"> 如：篇、组、个、部等</td><td width='250'>&nbsp;&nbsp;&nbsp;<%=ReturnPageStyle(PageStyle)%></td>
			 </tr>
			 </table>
			</td>

		  </tr>
		  
		  <tr class='tdbg'>
		   <td width="80" height="25" class='clefttitle' align='right'><strong>简要说明:</strong></td>
		   <td><textarea name="note" cols="80" style="width:98%" rows="9"><%=note%></textarea>
		   </td>
		  </tr>
		  
		  </form>
		 </table>
	  <%
	   End With
	  End Sub
	  
	  '第二步
	  Sub Step1()
	  %>
	  <script language="javascript">
	  function checkfield(){
		var strtmpp='' ;
		strtmpp= "<table border='1' cellpadding='2' cellspacing='1'  width='98%' class='border'><tr align='center'>";
		<%if datasourcetype=0 then%>
		strtmpp = strtmpp + "<td title='通用标签'><font color=red>通用标签=></font></td>";
		strtmpp = strtmpp +" <td title='当前模型ID' style='cursor:pointer;' onclick=AddParamToSql2('{$CurrChannelID}')>{$CurrChannelID}</td>";
		strtmpp = strtmpp + "<td style='cursor:pointer;'  onclick=AddParamToSql2('{$CurrClassID}') title='当前文章、图片、下载、动漫等的通用栏目ID，利用它可以构造出通用的自定义函数标签.如 Select id,title From KS_Article Where Tid=‘{$CurrClassID}’'>{$CurrClassID}</td>";
		strtmpp = strtmpp + "<td style='cursor:pointer;'  onclick=AddParamToSql2('{$CurrClassChildID}') title=\"包含子栏目的通用栏目ID,以“，”号隔开,如 Select ID,Title,AddDate From KS_Article Where Tid in({$CurrClassChildID})\">{$CurrClassChildID}</td>";
		strtmpp = strtmpp + "<td style='cursor:pointer' onclick=AddParamToSql2('{$CurrInfoID}') title='当前信息（文章，图片，下载等）的ID,如Select ID,Intro From KS_Article Where ID={$CurrInfoID}'>{$CurrInfoID}</td>";
		strtmpp = strtmpp + "<td style='cursor:pointer' onclick='AddParamToSql2('{$CurrSpecialID}')' title=\"当前专题ID（只限在专题页使用）,如Select ID,Intro From KS_Article Where specialid like '%{$CurrInfoID}%'\">{$CurrSpecialID}</td>";
		strtmpp = strtmpp + "</tr>";
		<%End if%>
		
		strtmpp = strtmpp + "<tr align='center'>";
		var fieldtemp = document.myform.FieldParam.value.split("\n");
			for(i=0;i<fieldtemp.length;i++){
				strtmpp = strtmpp + "<td style='cursor:pointer;' onclick='AddParamToSql(" + i + ")'>" + fieldtemp[i] + "</td>";
				if(((i+1)%5) == 0){
					strtmpp = strtmpp + "</tr><tr align='center'>";
				}
			}
			strtmpp = strtmpp + "</table>";
			document.getElementById ("ParamList").innerHTML=strtmpp;
     }
	 var pos=null;
	 function setPos()
	 { if (document.all){
			document.myform.LabelIntro.focus();
		    pos = document.selection.createRange();
		  }else{
		    pos = document.getElementById("LabelIntro").selectionStart;
		  }
	 }
	 //插入
	function InsertValue(Val)
	{  if (pos==null) {alert('请先定位要插入的位置!');return false;}
		if (document.all){
			  pos.text=Val;
		}else{
			   var obj=$("#LabelIntro");
			   var lstr=obj.val().substring(0,pos);
			   var rstr=obj.val().substring(pos);
			   obj.val(lstr+Val+rstr);
		}
	 }
	 function AddParamToSql(input){
		if (input != null){
			InsertValue("{$Param(" + input + ")}");
		}
	}
	function AddParamToSql2(input){
		if (input != null){
		   if (document.all)
		   {
		      myform.LabelIntro.focus();
		      var str = document.selection.createRange();
              str.text = input;
		   }else{
			InsertValue(input);
		   }
		}
	}
	
	  function CheckForm()
		{ var form=document.myform; 
		  if (form.LabelName.value=='')   
		  { alert('请输入标签名称!');
			  form.LabelName.focus();
			  return false; 
			} 
			  form.submit(); 
			  $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=标签管理 >> <font color=red>自定义SQL标签</font>&ButtonSymbol=LabelAdd';
			  return true;
		}
		 function changedb()
		 {
		  var dbname1=$('#dbname1').val();
		  var dbname2=$('#dbname2').val();
		  var LabelName=$('#LabelName').val();
		  var ParentID=$('#ParentID').val();
		  var Page=$('#Page').val();
		  var LabelID=$('#LabelID').val();
		  var PageStyle=$('#PageStyle').val();
		  var Ajax=$('#Ajax').val();
		  var SQLType=$('#SQLType').val();
		  var ItemName=$('#ItemName').val();
		  var datasourcetype=$('#datasourcetype').val();
		  var datasourcestr=$('#datasourcestr').val();
		  var Note=$('#Note').val();
		  location.href='LabelSQL.asp?action=Step1&Flag=addfield&Ajax='+Ajax+'&Note='+Note+'&datasourcetype='+datasourcetype+'&datasourcestr='+datasourcestr+'&dbname1='+dbname1+'&dbname2='+dbname2+'&LabelName='+LabelName+'&ParentID='+ParentID+'&Page='+Page+'&LabelID='+LabelID+'&SQLType='+SQLType+'&ItemName='+ItemName+'&PageStyle='+PageStyle;
		 }
		function addfield(){
			document.myform.LabelIntro.value='';
			var select=document.myform.field;
			var select2=document.myform.field2;
			for(i=0;i<select.length;i++){
				if(document.myform.field[i].selected==true){
					if(document.myform.dbname2.value==''){
						if (document.myform.LabelIntro.value==''){
							document.myform.LabelIntro.value=document.myform.field[i].value;
						}else{
							document.myform.LabelIntro.value=document.myform.LabelIntro.value+","+document.myform.field[i].value;
						}
					}else{
						if (document.myform.LabelIntro.value==''){
							document.myform.LabelIntro.value=document.myform.dbname1.value + "." + document.myform.field[i].value;
						}else{
							document.myform.LabelIntro.value=document.myform.LabelIntro.value + "," + document.myform.dbname1.value + "." + document.myform.field[i].value;
						}
					}
				}
			}
			if(document.myform.dbname2.value==''){
				if(document.myform.pagenum.value>0){
				<% if datasourcetype=5 then%>
					document.myform.LabelIntro.value="select " + document.myform.LabelIntro.value + " from [<%=dbname1%>]";
				}else{
					document.myform.LabelIntro.value="select top 10 " + document.myform.LabelIntro.value + " from [<%=dbname1%>]";
				}
				<%else%>
					document.myform.LabelIntro.value="select " + document.myform.LabelIntro.value + " from <%=dbname1%>";
				}else{
					document.myform.LabelIntro.value="select top 10 " + document.myform.LabelIntro.value + " from <%=dbname1%>";
				}
				<%end if%>
			}else{
				for(i=0;i<select2.length;i++){
					if(document.myform.field2[i].selected==true){
						if (document.myform.LabelIntro.value==''){
							document.myform.LabelIntro.value=document.myform.dbname2.value + "." + document.myform.field2[i].value;
						}else{
							document.myform.LabelIntro.value=document.myform.LabelIntro.value + "," + document.myform.dbname2.value + "." + document.myform.field2[i].value;
						}
					}
				}
				if(document.myform.dbname1.value==''){
					if(document.myform.pagenum.value>0){
						document.myform.LabelIntro.value="select " + document.myform.LabelIntro.value + " from <%=KS.G("dbname2")%>";
					}else{
						document.myform.LabelIntro.value="select top 10 " + document.myform.LabelIntro.value + " from <%=KS.G("dbname2")%>";
					}
				}else{
					if(document.myform.bg1.value==''){
						if(document.myform.pagenum.value>0){
							document.myform.LabelIntro.value="select " + document.myform.LabelIntro.value + " from <%=dbname1%>,<%=KS.G("dbname2")%>";
						}else{
							document.myform.LabelIntro.value="select top 10 " + document.myform.LabelIntro.value + " from <%=dbname1%>,<%=KS.G("dbname2")%>";
						}
					}else{
						if(document.myform.pagenum.value>0){
							document.myform.LabelIntro.value="select " + document.myform.LabelIntro.value + " from <%=dbname1%>,<%=KS.G("dbname2")%> where ";
						}else{
							document.myform.LabelIntro.value="select top 10 " + document.myform.LabelIntro.value + " from <%=dbname1%>,<%=KS.G("dbname2")%> where ";
						}
						document.myform.LabelIntro.value=document.myform.LabelIntro.value + "<%=dbname1%>." + document.myform.bg1.value + " = " + "<%=KS.G("dbname2")%>." + document.myform.bg2.value;
					}
				}
			}
		}
		</script>
		<script type="text/javascript">
		function ShowIframe()
		{  var p=new parent.KesionPopup()
			p.PopupCenterIframe("查看栏目<=>ID对照表","include/LabelXML.asp?action=ShowClassID",600,350,"auto")
		}
		</script>
	  <%
	  FolderID=KS.G("ParentID")
	  If SQLType="" Then SQLType=0:pagenum=0 else pagenum=1
	  IF ItemName="" Then ItemName="篇"
	   With KS
	    .echo "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
		.echo "  <form name=""myform"" id=""myform"" method=post action=""?action=Step2"" onSubmit=""return(CheckForm())"">"
		.echo "    <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""5"">"
		.echo "    <input type=""hidden"" name=""LabelID"" id=""LabelID"" value=""" & LabelID & """>"
		.echo "    <input type=""hidden"" name=""Page"" id=""Page"" value=""" & Page & """>"
		.echo "    <input type='hidden' name='pagenum' id=""pagenum"" value='" & pagenum &"' id='pagenum'>"	
		
		.echo " <input type='hidden' value=""" & LabelName & """ id=""LabelName"" name=""LabelName"" style=""width:200;"">	"
		.echo " <input type=""hidden"" name=""ParentID"" id=""ParentID"" value=""" & FolderID & """>"
		.echo " <input type=""hidden"" name=""SQLType"" id=""SQLType"" value=""" & SQLType & """>"
		.echo " <input type=""hidden"" name=""ItemName"" id=""ItemName"" value=""" & ItemName & """ size=""6""> "
		.echo " <input type=""hidden"" name=""PageStyle"" id=""PageStyle"" value=""" & PageStyle & """ size=""6""> "
		.echo " <input type=""hidden"" name=""Note"" id=""Note"" value=""" & note & """ size=""6""> "
		
		.echo " <input type=""hidden"" name=""Ajax"" id=""Ajax"" value=""" & KS.G("Ajax") & """>"
		.echo " <input type=""hidden"" name=""datasourcetype"" id=""datasourcetype"" value=""" & datasourcetype & """ size=""6""> "
		.echo " <input type=""hidden"" name=""datasourcestr"" id=""datasourcestr"" value=""" & datasourcestr & """> "
		
		
		.echo " <tr>"
		.echo "   <td height=""25"" colspan=""2""> "
		 .echo "      <table width='100%' height='25' border='0' cellpadding='0' cellspacing='1' bgcolor='#efefef' class='sort'>"
		 .echo "       <tr><td><div align='center'><font color='#990000'>"
		  If Action = "EditLabel" Then
		   .echo "修改自定义函数标签"
		   Else
		   .echo "第二步:构造SQL查询语句"
		  End If
		.echo "    </font></div></td><td><a href='javascript:ShowIframe()'><u>查看栏目<=>ID对照表</u></a></td></tr>"
		.echo "    </table>"
		.echo " </td>"
		.echo "    </tr>"
		.echo "    <tr><td colspan=""2"" align=""center"" height=""25"" class=""title""><strong>构 造 SQL 查 询 语 句</strong></td></tr>"
		.echo "    <tr>"
		.echo "      <td height=""30"" colspan=2>"
		%>
		<table style="margin-top:5px" width="100%" border=0 cellpadding='2' cellspacing='1' class='border'>
			<tr class="tdbg">
			  <td width=100 height="28" style="text-align:center"><strong>主表：</strong></td>
			  <td>
			  <select name='dbname1' id='dbname1' onChange='changedb()' class="textbox" style="WIDTH: 250px;" >
			  <option value=''>请选择一个数据表</option>
			  <%showmain(1)%>
			  </select>     </td>
			  <td style="text-align:center" width=100><strong>从表：</strong></td>
			  <td>
			  <select name='dbname2' id='dbname2' class="textbox" onChange='changedb()' style="WIDTH: 250px;" >
			  <option value=''>请选择一个数据表</option>
			  <%showmain(2)%>
			  </select>     </td>
			</tr>
			<tr class="tdbg" <%If dbname1<>"" and KS.G("dbname2")<>"" then .echo "" else .echo " style='display:none'"%>>
			  <td height="28" style="text-align:center"><strong>约束字段：</strong></td>
			  <td><select name='bg1' class="textbox" style='width:250px;'>
			  <Option value=''>选择主表字段</Option>
               <%
				if KS.G("flag")="addfield" then
				ShowChird(dbname1)
				end if
			 %>			  </select>			  </td>
			  <td style="text-align:center"><strong>&lt;&lt; 等于 &gt;&gt;</strong></td>
			  <td><select class="textbox" name='bg2' style='width:250px;'><option value=''>选择从表字段</option>
		  <%
			if KS.G("flag")="addfield" then
			ShowChird(KS.G("dbname2"))
			end if
			%>
			  </select>			  </td>
		  </tr>
			<tr class="tdbg">
			  <td style="text-align:center" width=100><strong>选择字段：</strong><br><br><font color=#ff0000>请选择需要调用的字段名称,按Ctrl或Shift键多选</font></td>
			  <td width=100>
			<Select class="textbox" style="WIDTH: 250px; HEIGHT: 210px" onChange='addfield()' multiple size=1 name="field">
			<%
			if dbname1="" then .echo "<Option value=0>请先选择一个表</Option>"
			if KS.G("flag")="addfield" then
			ShowChird(dbname1)
			end if
			%>
			  </Select></td>
			  <td style="text-align:center"><strong>&gt;&gt;&gt;</strong></td>
			  <td>
		<Select class="textbox" style="WIDTH: 250px; HEIGHT: 210px" onChange='addfield()' multiple size=2 name="field2">
		  <%
		  if KS.G("dbname2")="" then .echo "<Option value=0>请先选择一个表</Option>"
			if KS.G("flag")="addfield" then
			ShowChird(KS.G("dbname2"))
			end if
			%>
			  </Select></td>
		</tr>
		    <tr class="tdbg">
		      <td style="text-align:center"><strong>参数说明：</strong></td>
		      <td colspan=2 valign="middle"><textarea class="textbox" name='FieldParam' cols='55' rows='3' id='FieldParam' onKeyUp="checkfield();" style="height:60px"></textarea></td>
	          <td valign="middle"><font color='#FF0000'>*(不可改) 输入函数列表参数,每行一个,不带参数请留空。</font></td>
	      </tr>
	      <tr class="tdbg">
            <td width='100' style="text-align:center"><strong>SQL查询语句：</strong></td>
            <td colspan=3><div id="ParamList">
			<%if datasourcetype=0 then%>
			<table border='1' cellpadding='2' cellspacing='1'  width='98%' class='border'>
			<tr align='center'><td title='通用标签'><font color=red>通用标签=></font></td><td style='cursor:pointer;' onClick="AddParamToSql2('{$CurrChannelID}')" title='当前模型ID'>{$CurrChannelID}</td><td style='cursor:pointer;' onClick="AddParamToSql2('{$CurrClassID}')" title='当前文章、图片、下载、动漫等的通用栏目ID，利用它可以构造出通用的自定义函数标签.如 Select id,title From KS_Article Where Tid=‘{$CurrClassID}’'>{$CurrClassID}</td><td style='cursor:pointer;'  onclick="AddParamToSql2('{$CurrClassChildID}')" title="包含子栏目的通用栏目ID,以“，”号隔开,如 Select ID,Title,AddDate From KS_Article Where Tid in({$CurrClassChildID})">{$CurrClassChildID}</td><td style="cursor:pointer" onClick="AddParamToSql2('{$CurrInfoID}')" title="当前信息（文章，图片，下载等）的ID,如Select ID,Intro From KS_Article Where ID={$CurrInfoID}">{$CurrInfoID}</td><td style="cursor:pointer" onClick="AddParamToSql2('{$CurrSpecialID}')" title="当前专题ID（只限在专题页使用）,如Select ID,Intro From KS_Article Where specialid like '%{$CurrInfoID}%'">{$CurrSpecialID}</td></tr>
			</table>
			<%end if%>
			</div><textarea name='LabelIntro' onClick="setPos()" class="textbox" cols='97' rows='5' style='width:98%;height:80px' id='LabelIntro'>select top 10 * from KS_Article</textarea>
			<br>
			<font color=red>特别提示：</font>
			</td>
	   </tr>
	   
	   </table>
		<%
		.echo "      </td></tr>"
		.echo "</form></table>"
	  End With
	 End Sub
	 '**************************************************
	'过程名：ShowMain
	'作  用：显示数据表列表
	'参  数：无
	'**************************************************
	Sub ShowMain(Num)
		dim rs,tablename,temptable,modeltablestr
		With KS
		if datasourcetype=0 then
			Dim rsc:set rsc=conn.execute("select itemname,channeltable,channelname from ks_channel where channelid<>6 And ChannelID<>9 and channelstatus=1 order by channelid")
			if not rsc.eof then
				 .echo "<optgroup  style=""color:blue;"" label=""=============模型数据表============="">"
				 do while not rsc.eof
				   modeltablestr=modeltablestr & rsc(1) & ","
				   if KS.G("dbname"&num)= rsc(1) then
					.echo "<option value='" & rsc(1) & "' selected>" & rsc(0) & "数据表(" & rsc(2) &"|" & rsc(1) & ")</option>"
				   else
					.echo "<option value='" & rsc(1) & "'>" & rsc(0) & "数据表(" & rsc(2) &"|" & rsc(1) & ")</option>"
				   end if
					rsc.movenext
				 loop
				   if KS.G("dbname"&num)= "KS_Class" then
					.echo "<option value='KS_Class' selected style=""color:red"">模型栏目表</option>"
				   else
					.echo "<option value='KS_Class' style=""color:red"">模型栏目表</option>"
				   end if
				 modeltablestr=modeltablestr &"ks_class,"
				 .echo "<optgroup  label=""=============其它表============="">"
			end if
			rsc.close:set rsc=nothing
		 end if
		 
		 if datasourcetype=0 then
		 Set rs = Conn.OpenSchema(4)
		 else
		 Set rs = tConn.OpenSchema(4)
		 end if
		tablename=""
		Do While Not rs.EOF
			'temptable=Lcase(rs("Table_name"))
			temptable=rs("Table_name")
			if temptable <> tablename and temptable <> "KS_Admin" and temptable <> "KS_NotDown" and lcase(left(temptable,4)) <> "msys" and lcase(left(temptable,3)) <> "sys" and KS.FoundInArr(modeltablestr, temptable, ",")=false then
			'if (temptable ="KS_Article" or temptable = "KS_Photo" or temptable = "KS_DownLoad" or temptable = "KS_Flash" or temptable = "KS_Movie" or temptable = "KS_GQ" or temptable = "KS_Product" or temptable = "KS_Class") and temptable <> tablename then
			    if KS.G("dbname"&num)= temptable then
				.echo "<option value='" & temptable & "' selected>" & temptable & "</option>"
				else
				.echo "<option value='" & temptable & "'>" &temptable & "</option>"
				end if
				Tablename = temptable
			end if
		rs.MoveNext
		Loop
		rs.close:set rs=nothing
	 End With
	End Sub
	 '**************************************************
	'过程名：ShowChird
	'作  用：显示指定数据表的字段列表
	'参  数：无
	'**************************************************
	Sub ShowChird(dbname)
		dim rs
		if dbname<>"" then	
		   if datasourcetype<>0 then
		    Set rs=Tconn.OpenSchema(4)
		   else
			Set rs = Conn.OpenSchema(4)	
		   end if
		   
			Do Until rs.EOF or rs("Table_name") = trim(dbname)
				rs.MoveNext
			Loop
			Dim UserFieldArr,CommonFieldArr,CommonField
			
			if datasourcetype=0 then
				Dim rsc:set rsc=server.createobject("adodb.recordset")
				rsc.open "select channelname,itemname,BasicType from ks_channel where channelid<>6 And ChannelID<>9 and channeltable='" & dbname & "'",conn,1,1
				if not rsc.eof then
					CommonField=GetCommonField(rsc(0),rsc(2),CommonFieldArr,rsc(1))
					KS.echo CommonField
					Do Until rs.EOF or rs("Table_name") <> trim(dbname)
					   if left(lcase(rs("column_Name")),3)="ks_" then
						 if UserFieldArr="" then
						 UserFieldArr=UserFieldArr & rs("column_Name")
						 else
						 UserFieldArr=UserFieldArr & "," & rs("column_Name")
						 end if
					   elseif KS.FoundInArr(CommonFieldArr, lcase(rs("column_Name")), ",")=false and lcase(rs("column_Name"))<>"orderid" Then
						KS.echo "<option value='"&rs("column_Name")&"'>·"&GetFieldName(rs("column_Name"),rsc(1))&"</option>"
					   end if
						rs.MoveNext
					loop
					KS.echo GetUserField(UserFieldArr)
				else
				   
					if lcase(dbname)="ks_class" then KS.echo  GetCommonField("栏目",0,CommonFieldArr,"栏目")
					Do Until rs.EOF or rs("Table_name") <> trim(dbname)
					  if lcase(dbname)<>"ks_class" then
					   KS.echo "<option value='"&rs("column_Name")&"'>·"&rs("column_Name")&"</option>"
					  else
						KS.echo "<option value='"&rs("column_Name")&"'>·"&GetFieldName(rs("column_Name"),"")&"</option>"
					  end if
					   rs.MoveNext
					loop
				end if
				rsc.close:set rsc=nothing
			else 
					Do Until rs.EOF or rs("Table_name") <> trim(dbname)
					   KS.echo "<option value='"&rs("column_Name")&"'>·"&rs("column_Name")&"</option>"
					   rs.MoveNext
					loop
			End If
			rs.close:set rs=nothing
		End if
	End Sub
	
	'自定义字段
	Function GetUserField(UserFieldArr) 
	  Dim i
	  GetUserField="<optgroup  style=""color:red"" label=""=====用户自定义字段====="">"
	  UserFieldArr=Split(UserFieldArr,",")
	  For I=0 TO Ubound(UserFieldArr)
	   GetUserField= GetUserField&"<option value=""" &UserFieldArr(i) &""">·" &UserFieldArr(i) &"</option>"
	  Next
	End Function
	
	'常用字段列表
	Function GetCommonField(ChannelName,BasicType,byref CommonFieldArr,itemname)
	  select case BasicType
	     Case 0
		  CommonFieldArr="classid,id,foldername,createdate,creater"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====栏目表的常用字段====="">"
		  GetCommonField=GetCommonField &"<option value=""ClassID"">·栏目自动编号ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""ID"">·栏目编号ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""FolderName"">·栏目名称</option>"
		  GetCommonField=GetCommonField &"<option value=""CreateDate"">·栏目创建时间</option>"
		  GetCommonField=GetCommonField &"<option value=""Creater"">·栏目创建者</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""=====以下字段不建议使用====="">"
	     Case 1
		  CommonFieldArr="id,tid,title,author,editor,origin,inputer,adddate,hits,articlecontent,photourl,rank,Intro"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====" & ChannelName &"的常用字段====="">"
		  GetCommonField=GetCommonField &"<option value=""ID"">·" & itemname & "自动编号ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""Tid"">·" & itemname & "栏目ID(Url|名称)</option>"
		  GetCommonField=GetCommonField &"<option value=""Title"">·" & itemname & "标题</option>"
		  GetCommonField=GetCommonField &"<option value=""Author"">·" & itemname & "作者</option>"
		  GetCommonField=GetCommonField &"<option value=""Inputer"">·" & itemname & "录入员</option>"
		  GetCommonField=GetCommonField &"<option value=""Origin"">·" & itemname & "来源</option>"
		  GetCommonField=GetCommonField &"<option value=""Adddate"">·" & itemname & "添加/更新时间</option>"
		  GetCommonField=GetCommonField &"<option value=""Hits"">·" & itemname & "浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""rank"">·阅读等级</option>"
		  GetCommonField=GetCommonField &"<option value=""Intro"">·" & itemname & "导读</option>"
		  GetCommonField=GetCommonField &"<option value=""Articlecontent"">·" & itemname & "详细内容</option>"
		  GetCommonField=GetCommonField &"<option value=""PhotoUrl"">·" & itemname & "图片地址</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""=====" & ChannelName & "的其它字段====="">"
		 Case 2
		  CommonFieldArr="id,tid,title,author,photourl,origin,inputer,adddate,hits,hitsbyday,hitsbyweek,hitsbymonth,picturecontent,score,rank"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====" & ChannelName & "的常用字段====="">"
		  GetCommonField=GetCommonField &"<option value=""ID"">·" & itemname & "自动编号ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""Tid"">·" & itemname & "栏目ID(Url|名称)</option>"
		  GetCommonField=GetCommonField &"<option value=""Title"">·" & itemname & "名称</option>"
		  GetCommonField=GetCommonField &"<option value=""PhotoUrl"">·" & itemname & "地址</option>"
		  GetCommonField=GetCommonField &"<option value=""Author"">·" & itemname & "作者</option>"
		  GetCommonField=GetCommonField &"<option value=""Inputer"">·" & itemname & "录入员</option>"
		  GetCommonField=GetCommonField &"<option value=""Origin"">·" & itemname & "来源</option>"
		  GetCommonField=GetCommonField &"<option value=""Adddate"">·" & itemname & "添加/更新时间</option>"
		  GetCommonField=GetCommonField &"<option value=""Hits"">·总浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByDay"">·日浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByWeek"">·周浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByMonth"">·日浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""Score"">·得票数</option>"
		  GetCommonField=GetCommonField &"<option value=""Rank"">·推荐等级</option>"
		  GetCommonField=GetCommonField &"<option value=""picturecontent"">·" & itemname & "介绍</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""====" & ChannelName & "的其它字段====="">"
		 Case 3
		  CommonFieldArr="id,tid,title,author,downlb,downyy,downsq,downsize,ysdz,photourl,origin,inputer,adddate,hits,hitsbyday,hitsbyweek,hitsbymonth,downcontent"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====" & ChannelName & "的常用字段====="">"
		  GetCommonField=GetCommonField &"<option value=""ID"">·" & itemname & "自动编号ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""Tid"">·" & itemname & "栏目ID(Url|名称)</option>"
		  GetCommonField=GetCommonField &"<option value=""Title"">·" & itemname & "名称</option>"
		  GetCommonField=GetCommonField &"<option value=""PhotoUrl"">·" & itemname & "缩略图</option>"
		  GetCommonField=GetCommonField &"<option value=""Author"">·" & itemname & "作者</option>"
		  GetCommonField=GetCommonField &"<option value=""DownLB"">·" & itemname & "类别</option>"
		  GetCommonField=GetCommonField &"<option value=""DownYY"">·" & itemname & "语言</option>"
		  GetCommonField=GetCommonField &"<option value=""DownSQ"">·" & itemname & "授权</option>"
		  GetCommonField=GetCommonField &"<option value=""DownSize"">·" & itemname & "大小</option>"
		  GetCommonField=GetCommonField &"<option value=""YSDZ"">·演示地址</option>"
		  GetCommonField=GetCommonField &"<option value=""Inputer"">·" & itemname & "录入员</option>"
		  GetCommonField=GetCommonField &"<option value=""Origin"">·" & itemname & "来源</option>"
		  GetCommonField=GetCommonField &"<option value=""Adddate"">·" & itemname & "添加/更新时间</option>"
		  GetCommonField=GetCommonField &"<option value=""Hits"">·总浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByDay"">·日浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByWeek"">·周浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByMonth"">·日浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""downcontent"">·" & itemname & "介绍</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""=====" & ChannelName & "的其它字段====="">"
		 Case 4
		  CommonFieldArr="id,tid,title,author,photourl,origin,inputer,adddate,hits,hitsbyday,hitsbyweek,hitsbymonth,flashcontent,score,rank"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====" & ChannelName & "的常用字段====="">"
		  GetCommonField=GetCommonField &"<option value=""ID"">·" & itemname & "自动编号ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""Tid"">·" & itemname & "栏目ID(Url|名称)</option>"
		  GetCommonField=GetCommonField &"<option value=""Title"">·" & itemname & "名称</option>"
		  GetCommonField=GetCommonField &"<option value=""PhotoUrl"">·图片地址</option>"
		  GetCommonField=GetCommonField &"<option value=""Author"">·" & itemname & "作者</option>"
		  GetCommonField=GetCommonField &"<option value=""Inputer"">·" & itemname & "录入员</option>"
		  GetCommonField=GetCommonField &"<option value=""Adddate"">·" & itemname & "添加/更新时间</option>"
		  GetCommonField=GetCommonField &"<option value=""Hits"">·总浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByDay"">·日浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByWeek"">·周浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByMonth"">·日浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""Score"">·得票数</option>"
		  GetCommonField=GetCommonField &"<option value=""Rank"">·推荐等级</option>"
		  GetCommonField=GetCommonField &"<option value=""flashcontent"">·" & itemname & "介绍</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""=====" & ChannelName & "的其它字段====="">"
		 Case 5
		  CommonFieldArr="id,tid,title,author,photourl,bigphoto,promodel,inputer,adddate,hits,prospecificat,producername,trademarkname,ProIntro,score,rank,price,price_member,price_market,price_original,serviceterm,totalnum,unit,discount"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====" & ChannelName & "的常用字段====="">"
		  GetCommonField=GetCommonField &"<option value=""ID"">·" & itemname & "自动编号ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""Tid"">·" & itemname & "栏目ID(Url|名称)</option>"
		  GetCommonField=GetCommonField &"<option value=""Title"">·" & itemname & "名称</option>"
		  GetCommonField=GetCommonField &"<option value=""PhotoUrl"">·" & itemname & "小图</option>"
		  GetCommonField=GetCommonField &"<option value=""BigPhoto"">·" & itemname & "大图</option>"
		  GetCommonField=GetCommonField &"<option value=""Price"">·市场价格</option>"
		  GetCommonField=GetCommonField &"<option value=""Price_Member"">·会员价格</option>"
		  GetCommonField=GetCommonField &"<option value=""ServiceTerm"">·服务年限</option>"
		  GetCommonField=GetCommonField &"<option value=""TotalNum"">·库存数量</option>"
		  GetCommonField=GetCommonField &"<option value=""ProModel"">·" & itemname & "型号</option>"
		  GetCommonField=GetCommonField &"<option value=""Unit"">·" & itemname & "单位</option>"
		  GetCommonField=GetCommonField &"<option value=""Inputer"">·" & itemname & "录入员</option>"
		  GetCommonField=GetCommonField &"<option value=""Adddate"">·上市时间</option>"
		  GetCommonField=GetCommonField &"<option value=""Hits"">·总浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""ProSpecificat"">·商品规格</option>"
		  GetCommonField=GetCommonField &"<option value=""ProducerName"">·生产商</option>"
		  GetCommonField=GetCommonField &"<option value=""TrademarkName"">·品牌/商标</option>"
		  GetCommonField=GetCommonField &"<option value=""Rank"">·推荐等级</option>"
		  GetCommonField=GetCommonField &"<option value=""ProIntro"">·" & itemname & "介绍</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""=====" & ChannelName & "的其它字段====="">"
		 Case 7
		  CommonFieldArr="id,tid,title,movieact,photourl,movietime,moviedy,adddate,hits,screentime,movieyy,moviedq,moviecontent,score,rank,inputer"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====" & ChannelName & "的常用字段====="">"
		  GetCommonField=GetCommonField &"<option value=""ID"">·" & itemname & "自动编号ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""Tid"">·" & itemname & "栏目ID(Url|名称)</option>"
		  GetCommonField=GetCommonField &"<option value=""Title"">·" & itemname & "名称</option>"
		  GetCommonField=GetCommonField &"<option value=""PhotoUrl"">·图片地址</option>"
		  GetCommonField=GetCommonField &"<option value=""MovieAct"">·主要演员</option>"
		  GetCommonField=GetCommonField &"<option value=""MovieDY"">·" & itemname & "导演</option>" 
		  GetCommonField=GetCommonField &"<option value=""MovieTime"">·播放长度</option>"
		  GetCommonField=GetCommonField &"<option value=""ScreenTime"">·上映时间</option>"
		  GetCommonField=GetCommonField &"<option value=""MovieYY"">·" & itemname & "语言</option>"
		  GetCommonField=GetCommonField &"<option value=""MovieDQ"">·出产地区</option>"		  
		  GetCommonField=GetCommonField &"<option value=""Adddate"">·" & itemname & "添加/更新时间</option>"
		  GetCommonField=GetCommonField &"<option value=""Hits"">·总浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""Score"">·得票数</option>"
		  GetCommonField=GetCommonField &"<option value=""Rank"">·推荐等级</option>"
		  GetCommonField=GetCommonField &"<option value=""MovieContent"">·" & itemname & "介绍</option>"
		  GetCommonField=GetCommonField &"<option value=""Inputer"">·" & itemname & "录入员</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""=====" & ChannelName & "的其它字段====="">"
		 Case 8
		  CommonFieldArr="id,tid,title,author,photourl,address,contactman,province,adddate,hits,city,companyname,tel,gqcontent,zip,username,fax,email,homepage,validdate,price"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====" & ChannelName & "的常用字段====="">"
		  GetCommonField=GetCommonField &"<option value=""ID"">·信息自动编号ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""Tid"">·" & itemname & "栏目ID(Url|名称)</option>"
		  GetCommonField=GetCommonField &"<option value=""Title"">·" & itemname & "主题名称</option>"
		  GetCommonField=GetCommonField &"<option value=""gqcontent"">·" & itemname & "的详细内容</option>"
		  GetCommonField=GetCommonField &"<option value=""PhotoUrl"">·" & itemname & "图片地址</option>"
		  GetCommonField=GetCommonField &"<option value=""Adddate"">·" & itemname & "添加/更新时间</option>"
		  GetCommonField=GetCommonField &"<option value=""ValidDate"">·有效天数</option>"
		  GetCommonField=GetCommonField &"<option value=""Price"">·价格</option>"
		  GetCommonField=GetCommonField &"<option value=""Hits"">·浏览次数</option>"
		  GetCommonField=GetCommonField &"<option value=""UserName"">·发布会员名</option>"
		  GetCommonField=GetCommonField &"<option value=""ContactMan"">·联系人</option>"
		  GetCommonField=GetCommonField &"<option value=""Address"">·联系地址</option>"
		  GetCommonField=GetCommonField &"<option value=""Tel"">·联系电话</option>"
		  GetCommonField=GetCommonField &"<option value=""Fax"">·传真号码</option>"
		  GetCommonField=GetCommonField &"<option value=""Email"">·电子邮箱</option>"
		  GetCommonField=GetCommonField &"<option value=""Zip"">·邮政编码</option>"
		  GetCommonField=GetCommonField &"<option value=""HomePage"">·主页地址</option>"
		  GetCommonField=GetCommonField &"<option value=""Province"">·所在省份</option>"
		  GetCommonField=GetCommonField &"<option value=""City"">·所在城市</option>"
		  GetCommonField=GetCommonField &"<option value=""CompanyName"">·公司名称</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""=====" & ChannelName & "的其它字段====="">"
		 case else
		  GetCommonField=""
	  end Select
	End Function

	Function GetFieldName(EField,itemname)
	   if datasourcetype<>0 then GetFieldName=Efield:exit function
	  Select Case Lcase(EField)
	  case "classid"
	     GetFieldName="栏目自动编号ClassID"
	   case "fnametype"
	     GetFieldName="生成的信息扩展名"
	   case "creater"
	     GetFieldName="栏目创建者"
	   case "createdate"
	     GetFieldName="栏目创建时间"
	   case "templateid"
	     GetFieldName="栏目下的信息模板ID"
	   case "channelid"
	     GetFieldName="模块ID"
	   case "cirspecialshowtf"
	     GetFieldName="循环栏目专题显示标志"
	   case "classbasicinfo"
	     GetFieldName="栏目信息配置集合"
	   case "classdefinecontent"
	     GetFieldName="栏目自设内容集合"
	   case "classpurview"
	     GetFieldName="栏目浏览权限ID"
	   case "commenttf"
	     GetFieldName="栏目下信息允行评论标志"
	   case "defaultarrgroupid"
	     GetFieldName="栏目下默认指定会员组的查看权限"
	   case "defaultchargetype"
	     GetFieldName="栏目下默认重复收费方式"
	   case "defaultdividepercent"
	     GetFieldName="栏目下与投稿者的默认分成比例"
	   case "defaultpitchtime"
	     GetFieldName="栏目下的默认重复收费查看次数"
	   case "defaultreadpoint"
	     GetFieldName="栏目下的默认收费点数"
	   case "defaultreadtimes"
	     GetFieldName="栏目下的默认重复收费查看次数"
	   case "folder"
	     GetFieldName="目录英文名称"
	   case "folderdomain"
	     GetFieldName="栏目绑定的域名"
	   case "folderfsoindex"
	     GetFieldName="栏目生成的首页名称"
	   case "folderorder"
	     GetFieldName="栏目排序序号"
	   case "foldertemplateid"
	     GetFieldName="栏目的模板ID"
	   case "specialtemplateid"
	     GetFieldName="频道专题列表页模板ID"
	   case "tn"
	     GetFieldName="父栏目ID"
	   case "tj"
	     GetFieldName="栏目深度"
	   case "ts"
	     GetFieldName="栏目ID集合列表"
	   case "topflag"
	     GetFieldName="顶部导航显示标志"
	   case "movieact"
	     GetFieldName= itemname & "演员"
	   case "moviedq"
	     GetFieldName=itemname & "地区"
	   case "moviedy"
	     GetFieldName=itemname & "导演"
	   case "movietime"
	     GetFieldName="播放长度"
	   case "screentime"
	     GetFieldName="上映时间"
	   case "movieyy"
	     GetFieldName=itemname & "语言"
	   case "moviecontent"
	     GetFieldName=itemname & "介绍"
	   case "movietype"
	     GetFieldName=itemname & "播放格式ID"
	   case "movieurls"
	     GetFieldName=itemname & "播放地址"
	   case "serverid"
	     GetFieldName="播放服务器ID"
	   case "alarmnum"
	     GetFieldName="下限报警数"
	   case "producttype"
	     GetFieldName="销售类型ID"
	   case "isspecial"
	     GetFieldName="特价标志"
	   case "price_member"
	     GetFieldName="会员价格"
	   case "price_market"
	     GetFieldName="市场价格"
	   case "price_original"
	     GetFieldName="原始零售价"
	   case "discount"
	     GetFieldName="折扣"
	  case "serviceterm"
	     GetFieldName="服务年限"
	   case "totalnum"
	     GetFieldName="库存量"
	   case "promodel"
	     GetFieldName=itemname & "型号"
	  case "unit"
	     GetFieldName=itemname & "单位" 
	  case "producername"
	     GetFieldName="生产商" 
	  case "prospecificat"
	     GetFieldName=itemname & "规格" 
	  case "trademarkname"
	     GetFieldName="品牌/商标" 
	  case "point"
	     GetFieldName="购物积分" 
	  case "prointro"
	     GetFieldName=itemname & "介绍" 
	   case "newsid","picid","downid","flashid","movieid","proid","gqid"
	     GetFieldName="系统生成的唯一ID(url)"
	   case "picurls"
	     GetFieldName="图片地址集合"
	   case "score"
	     GetFieldName="得票数"
	   Case "adddate"
	    GetFieldName="添加/更新时间"
	   Case "tid"
	    GetFieldName="栏目ID(Url|名称)"
	   case "arrgroupid"
	    GetFieldName="有权查看的会员组ID"
	   case "articlecontent"
	    GetFieldName=itemname & "详细内容"
	   case "inputer"
	    GetFieldName=itemname & "录入员"
	   case "photourl"
	    GetFieldName="图片地址"
	   case "bigphoto"
	    GetFieldName="大图片地址"
	   case "hitsbyday"
	    GetFieldName="日浏览次数"
	   case "hitsbyweek"
	    GetFieldName="周浏览次数"
	   case "hitsbymonth"
	    GetFieldName="月浏览次数"
	   case "picturecontent"
	    GetFieldName=itemname & "介绍"
	   case "author"
	    GetFieldName="作者"
	   case "origin"
	    GetFieldName="来源"
	   case "picurl"
	    GetFieldName="图片地址"
	   case "downlb"
	    GetFieldName="软件类别"
	   case "downpt"
	    GetFieldName="软件平台"
	   case "downsize"
	    GetFieldName="软件大小"
	   case "downsq"
	    GetFieldName="授权方式"
	   case "downyy"
	    GetFieldName=itemname & "语言"
	   case "ysdz"
	    GetFieldName=itemname & "演示地址"
	   case "zcdz"
	    GetFieldName=itemname & "注册地址"
	   case "downversion"
	    GetFieldName=itemname & "版本"
	   case "downurls"
	    GetFieldName="下载地址集合"
	   case "inputer"
	    GetFieldName=itemname & "录入员"
	   case "downcontent"
	    GetFieldName=itemname & "简介"
	   case "flashurl"
	    GetFieldName=itemname & "地址"
	   case "flashcontent"
	    GetFieldName=itemname & "介绍"		
	   case "typeid"
	    GetFieldName="交易类别ID"			
	   case "gqcontent"
	   	GetFieldName="供求详细内容"		
	   case "validdate"
	   	GetFieldName="有效天数"		
	   case "price"
	   	GetFieldName="价格"
	   case "username"
	    GetFieldName="用户名"
	   case "contactman"
	    GetFieldName="联系人"
	   case "address"
	    GetFieldName="联系地址"
	   case "tel"
	    GetFieldName="联系电话"
	   case "fax"
	    GetFieldName="传真号码"
	   case "email"
	    GetFieldName="电子邮箱"
	   case "zip"
	    GetFieldName="邮政编码"
	   case "homepage"
	    GetFieldName="公司主页"		
	   case "province"
	    GetFieldName="省份"		
	   case "city"
	    GetFieldName="城市"		
	   case "companyname"
	    GetFieldName="公司名称"		
	   case "beyondsavepic"
	    GetFieldName="远程存图标志"
	   case "changes"
	    GetFieldName="转向链接标志"
	   case "chargetype"
	    GetFieldName="重复收费方式"
	   case "comment"
	    GetFieldName="允许评论标志"
	   case "deltf"
	    GetFieldName="放入回收站标志"
	   case "dividepercent"
	     GetFieldName="投稿分成比例"
	   case "fname"
	     GetFieldName="生成的文件名"
	   case "fulltitle"
	     GetFieldName="文章完整标题"
	  case "infopurview"
	     GetFieldName="阅读权限方式"
	  case "istop"
	     GetFieldName="置顶标志"
	  case "jsid"
	     GetFieldName="加入的JSID列表"
	  case "keywords"
	     GetFieldName="关键字"
	  case "picnews"
	     GetFieldName="图片新闻标志"
	  case "pitchtime"
	     GetFieldName="重复收费小时数"
	  case "popular"
	     GetFieldName="热门标志"
	  case "rank"
	     GetFieldName="等级"
	  case "readpoint"
	     GetFieldName="需要的阅读点数"
	  case "readtimes"
	     GetFieldName="阅读指定次数重新收费"
	  case "recommend"
	     GetFieldName="推荐标志"
	  case "refreshtf"
	     GetFieldName="已生成标志"
	  case "rolls"
	     GetFieldName="滚动标志"
	  case "showcomment"
	     GetFieldName="标题旁显示评论标志"
	  case "slide"
	     GetFieldName="幻灯片标志"
	  case "specialid"
	     GetFieldName="专题ID"
	  case "strip"
	     GetFieldName="头条标志"
	  case "templateid"
	     GetFieldName="模板ID"
	  case "titlefontcolor"
	     GetFieldName="标题颜色"
	  case "titlefonttype"
	     GetFieldName="标题加粗+斜体标志"
	  case "titletype"
	     GetFieldName="图文标志"
	   case "hits"
	     GetFieldName="点击次数"
	   case "verific"
	     GetFieldName="审核标志"
	   case "id"
	     GetFieldName="自动编号ID(Url)"
	   case "foldername"
	     GetFieldName="栏目名称"
	   case "lasthitstime"
	     GetFieldName="最后被浏览的时间"
	   case "title"
	      If InStr(lcase(LabelIntro),"ks_article") then
	       GetFieldName=itemname & "标题"
		  else
		   GetFieldName=itemname & "名称"
		  end if
	 Case "Intro"
	     GetFieldName="导读内容"
	   Case else
	    GetFieldName=efield
	  End Select
	   'GetFieldName=Efield&"(" & GetFieldName&"）"
	  ' GetFieldName=GetFieldName
	End Function
	 
	  '第三步
	 Sub Step2()
	    Dim FieldParam,FieldParamArr,LoopTimes,NoRecord
		LabelName = Request.Form("LabelName")
		FolderID=KS.G("ParentID")
		Ajax=KS.G("Ajax")
	    SQLType=KS.G("SQLType")
		PageStyle=KS.G("PageStyle")
		ItemName=KS.G("ItemName")
        if datasourcetype<>0 then Call OpenExtConn()
		With KS
		Set LabelRS = Server.CreateObject("Adodb.RecordSet")
		If LabelID <> "" Then
		    ActionStr="EditSubmit"
			Set LabelRS = Server.CreateObject("Adodb.Recordset")
			SQLStr = "SELECT top 1 * FROM [KS_Label] Where ID='" & LabelID & "'"
			LabelRS.Open SQLStr, Conn, 1, 1
			If Not LabelRS.Eof Then
				LabelContent = Server.HTMLEncode(LabelRS("LabelContent"))
				FieldParamArr= Split(LabelRS("Description")&"@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@","@@@")
			End IF
			If Not KS.IsNul(Request("LabelIntro")) Then
			LabelIntro=Request("LabelIntro")
			Else
			LabelIntro =FieldParamArr(0)
			End If
			If Ubound(FieldParamArr)>=1 Then
				FieldParam = FieldParamArr(1)
			End If
			NoRecord = FieldParamArr(9)
			LabelRS.Close
		Else
		  LabelIntro=request("LabelIntro")
		  FieldParam=KS.G("FieldParam")
		  ActionStr="AddNewSubmit"
		  LoopTimes=GetLoopTimes(LabelIntro)
		  NoRecord ="没有记录!"
		  LabelContent="[loop=" & LoopTimes &"]请在此输入循环内容[/loop]"
		End If
		 Call SqlValid(LabelIntro)
		%>
		<script src="../../ks_inc/kesion.box.js"></script>
		<script language="javascript">
		var pos=null;
		function setPos()
		{ if (document.all){
			document.myform.LabelContent.focus();
		    pos = document.selection.createRange();
		  }else{
		    pos = document.getElementById("LabelContent").selectionStart;
		  }
		}
		function FieldInsertCode(fieldname,dbtype,dbname)
		{ 
		   if(pos==null) {alert('请先定位插入位置!');return false;}
		   var link="Admin_FieldParam.asp?fieldname=" + fieldname + "&dbtype="+ dbtype + "&dbname=" + dbname+"&datasourcetype=<%=datasourcetype%>";
		  var p=new KesionPopup()
		  p.PopupImgDir="../";
		  p.PopupCenterIframe('插入字段标签',link,350,230,'no');
		}
		
		//插入到循环体
		function InsertValue(Val)
		{
			 if (document.all){
			  pos.text=Val;
			 }else{
			   var obj=$("#LabelContent");
			   var lstr=obj.val().substring(0,pos);
			   var rstr=obj.val().substring(pos);
			   obj.val(lstr+Val+rstr);
			 }
		}
		function FieldInsertCode1(Val)
		{ 
		
		  if (Val!=''){
		   InsertValue(Val);
		   }
		}
		</script>
		<script language = 'JavaScript'>

		function show_ln(txt_ln,txt_main){
			var txt_ln  = document.getElementById(txt_ln);
			var txt_main  = document.getElementById(txt_main);
			txt_ln.scrollTop = txt_main.scrollTop;
			while(txt_ln.scrollTop != txt_main.scrollTop)
			{
				txt_ln.value += (i++) + '\n';
				txt_ln.scrollTop = txt_main.scrollTop;
			}
			return;
		}
		
		</script>
		<script type="text/javascript">
		function ShowIframe()
		{  var p=new parent.KesionPopup()
			p.PopupCenterIframe("查看栏目<=>ID对照表","include/LabelXML.asp?action=ShowClassID",600,350,"auto")
		}
		</script>
		<%
		.echo "<table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		.echo "  <form name=""myform"" id=""myform"" method=post action=""LabelSQL.asp"">"
		.echo "    <input type='hidden' name='keyword' id='keyword' value='" & KeyWord & "'>"
		.echo "    <input type='hidden' name='Searchtype' id='Searchtype' value='" & searchtype & "'>"
		.echo "    <input type=""hidden"" name=""LabelFlag"" id='LabelFlag' value=""3"">"
		.echo "    <input type=""hidden"" name=""LabelID"" id='LabelID' value=""" & LabelID & """>"
		.echo "    <input type=""hidden"" name=""FolderID"" id='FolderID' value=""" & FolderID & """>"
		.echo "    <input type=""hidden"" name=""Page"" id='Page' value=""" & Page & """>"
		.echo "    <input type=""hidden"" name=""FieldParam"" id='FieldParam' value=""" & FieldParam & """>"
		.echo "    <input type='hidden' name='Action' id='Action' value='" & ActionStr & "'>"
		
		.echo " <input type=""hidden"" name=""LabelName"" id=""LabelName"" value=""" &LabelName & """>"
		.echo " <input type=""hidden"" name=""SQLType"" id=""SQLType"" value=""" &SQLType & """>"
		.echo " <input type=""hidden"" name=""Ajax"" id=""Ajax"" value=""" & Ajax & """>"
		.echo " <input type=""hidden"" name=""ItemName"" id=""ItemName"" value=""" & ItemName & """>"
		.echo " <input type=""hidden"" name=""PageStyle"" id=""PageStyle"" value=""" & PageStyle & """>"
		.echo " <input type=""hidden"" name=""Note"" id=""Note"" value=""" & note & """ size=""6""> "
		
		.echo " <input type=""hidden"" name=""datasourcetype"" id=""datasourcetype"" value=""" & datasourcetype & """ size=""6""> "
		.echo " <input type=""hidden"" name=""datasourcestr"" id=""datasourcestr"" value=""" & datasourcestr & """> "

		.echo " <tr>"
		.echo "   <td height=""25"" colspan=""2"" bgcolor='#efefef' class='sort'> "
		.echo "    <div align='center'><font color='#990000'>"
		.echo "第三步：建立标签样式（循环内容）</font> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='javascript:ShowIframe()'><u>查看栏目<=>ID对照表</u></a>"
		.echo " </div></td>"
		.echo "    </tr>"
		.echo "    <tr style='display:none'>"
		.echo "      <td width=""60"" height=""19"">标签目录</td>"
		.echo "      <td>" & ReturnLabelFolderTree(FolderID, 1) & "</td>"
		.echo "    </tr>"
		.echo "    <tr class=""tableBorder1"">"
		.echo "      <td height=""16""><div align=""left"">查询语句</div></td>"
		.echo "      <td><textarea name=""LabelIntro"" rows=""4"" style=""width:98%;"">" & LabelIntro & "</textarea></td>"
		.echo "    </tr>"
		.echo "    <tr class=""tableBorder1"">"
		.echo "      <td height=""16""><div align=""left"">查询无记录时输出内容</div></td>"
		.echo "      <td class='tips'><input style='height:23px;' size='60' type=""text"" class=""textbox"" name=""NoRecord"" value=""" & NoRecord &"""/> 可留空,支持HTML语法。</td>"
		.echo "    </tr>"
        .echo "    <tr class=""tableBorder1"">"
		.echo "      <td width=""60"" height=""30"" nowrap align=center><strong>可用字段</strong></td>"
		.echo "      <td>"
		 Dim FieldName,dbtype,I,J,ClickStr,isidarr,isid
		 Dim RSField:Set RSField=Server.CreateObject("ADODB.RECORDSET")
		 Call OpenExtConn()
		 if datasourcetype<>0 then
		 RSField.Open ClearParam(LabelIntro),tConn,1,3
		 else
		 RSField.Open ClearParam(LabelIntro),Conn,1,3
		 end if
		  .echo "<table style=""table-layout:fixed"" border=1 bordercolordark=""#999999"" bordercolorlight=""#FFFFFF"" width=""710"" cellpadding='0' cellspacing='0'>"
		  .echo "<tr class='tdbg' height='20'>"
		  For I=0 To RSField.Fields.count-1
		     dbtype=RSField.Fields(i).type
			 FieldName=RSField.Fields(i).name
			 isidarr=split(FieldName,".")
				isid=false
				if ubound(isidarr)=1 then
				  if lcase(isidarr(1))="id" then
					isid=true
				  end if
				end if
			 If (Lcase(FieldName)="tid" or Lcase(FieldName)="id" or isid or Lcase(FieldName)="newsid" Or Lcase(FieldName)="picid" or Lcase(FieldName)="downid" or Lcase(FieldName)="flashid" or Lcase(FieldName)="proid" or Lcase(FieldName)="movieid" or Lcase(FieldName)="gqid" or Lcase(FieldName)="classid") and  datasourcetype=0  Then
			   
			    Dim sChannelID
			   	If DataBaseType=1 Then
				  dim rsc:set rsc=server.CreateObject("adodb.recordset")
				  rsc.open "Select ChannelID From KS_Channel Where charindex(channeltable,'" & ReplaceBC(LabelIntro) & "')>0",conn,1,1
				  if not rsc.eof then
				    sChannelID=rsc(0)
				  end if
				  rsc.close:set rsc=nothing
				Else
				  if not Conn.Execute("Select ChannelID From KS_Channel Where Instr('" & ReplaceBC(LabelIntro) & "',channeltable)>0").eof then
				   sChannelID=Conn.Execute("Select ChannelID From KS_Channel Where Instr('" & ReplaceBC(LabelIntro) & "',channeltable)>0")(0)
				  end if 
				End If
			   if sChannelID>=1 then
			     ClickStr="FieldInsertCode('" & FieldName & "',"&dbtype&"," & sChannelID & ")"
			   ElseIf Instr(Lcase(LabelIntro),"ks_class") then 
			     ClickStr="FieldInsertCode('" & FieldName & "',"&dbtype&",100)"
			   Else 
			     ClickStr="FieldInsertCode('" & FieldName & "',"&dbtype&",0)"
			   End If
			 Else
			   ClickStr="FieldInsertCode('" & FieldName & "',"&dbtype&",0)"
			 End IF
			 If j=5 Then j=0:.echo "</tr><tr class='tdbg' height='20'>"
			  J=J+1
			  if instr(FieldName,".")=0 then
			  .echo "<td  width=""20%"" align=""center"" onMouseOut=""this.className='tdbg'"" onMouseOver=""this.className='tdbgmouseover'"" style=""cursor:pointer;"" onClick=""" & ClickStr & """>" & GetFieldName(trim(FieldName),"") & "</td>"
			 else
			  .echo "<td  width=""20%"" align=""center"" onMouseOut=""this.className='tdbg'"" onMouseOver=""this.className='tdbgmouseover'"" style=""cursor:pointer;"" onClick=""" & ClickStr & """>" & split(FieldName,".")(0)&"." & GetFieldName(trim(split(FieldName,".")(1)),"") & "</td>"
			 end if
		 Next
		 For I=J+1 to 5
		  .echo "<td class='tdbg' height='25'>&nbsp;</td>"
		 Next
		 .echo "</tr>"
		 .echo "</table>"
		 .echo  "</td>"
		 .echo "  </tr>"
		 If FieldParam<>"" Then
		 .echo "<tr class=""tableBorder1""><td width=""60"" height=""30"" nowrap align=center><strong>函数参数</strong></td>"
		 .echo "<td>"
		 .echo "<table border=1 bordercolordark=""#999999"" bordercolorlight=""#FFFFFF"" width=""100%"" cellpadding='0' cellspacing='0'>"
		 .echo "<tr class='tdbg' height='20'>"
		 FieldParamArr=Split(FieldParam,vbcrlf)
		 J=0
		 For I=0 To Ubound(FieldParamArr)
		   If j=5 Then j=0:.echo "</tr><tr class='tdbg' height='20'>"
		   J=J+1
		 .echo "<td  width=""20%"" align=""center"" onMouseOut=""this.className='tdbg'"" onMouseOver=""this.className='tdbgmouseover'"" style=""cursor:pointer;"" onClick=""FieldInsertCode1('{$Param(" & I & ")}');"">" & FieldParamArr(I) &"</td>"
		 Next
		 For I=J+1 to 5
		  .echo "<td class='tdbg' height='25'>&nbsp;</td>"
		 Next
		 .echo "</tr>"
		.echo "</table>"
		.echo "</td></tr>"
		End If
		
		 .echo "   <tr class=""tableBorder1"">"
 
		 .echo "	<td align='center'><strong>循 环 体</strong>{$AutoID}</td><td height='230' valign=""top""><textarea id='txt_ln' name='rollContent' cols='6' style='width:35px;overflow:hidden;height:100%;background-color:highlight;border-right:0px;text-align:right;font-family: tahoma;font-size:12px;font-weight:bold;color:highlighttext;cursor:default;' readonly>"
		 Dim N
		 For N=1 To 3000
			.echo N & "&#13;&#10;"
		 Next
		 .echo"</textarea>"
		 .echo "<textarea name='LabelContent'  onclick='setPos()' onkeyup='setPos()' id='LabelContent' style='width:670px;height:100%' rows='15' id='txt_main' onscroll=""show_ln('txt_ln','LabelContent')"" wrap='on'>" & LabelContent & "</textarea>" & vbNewLine
		 .echo "	<script>for(var i=3000; i<=3000; i++) document.getElementById('txt_ln').value += i + '\n';</script>"
		 .echo "   </td></tr>"
		 
		 .echo "   <tr class=""tableBorder1"">"
 
		 .echo "	<td><strong>简要描述</strong></td><td><font color=red>1、SQL标签定义规则</font><br>循环体格式：[loop=n]循环标签的内容[/loop]<br>其中n表示循环次数，且n满足n>=0。loop为循环关键字，此循环体可以重复使用,但是不能嵌套。<font color=red><br>2、SQL标签字段规则</font><br>字段格式：{$Field(FieldName,OutType,Param,...)}<br>FieldName&nbsp;&nbsp;--数据库表的字段名称<br>OutType&nbsp;&nbsp;&nbsp;&nbsp;--输出类型 支持：文本(Text)、日期(Date)、数据(Num)、对象URL(GetInfoUrl)，栏目URL(GetClassUrl) 5种类型<br><font color=red>3、支持使用{ReqNum(字符串)}或{ReqStr(字符串)}来取得Url的参数值ex.asp?ClassID=100,那么{ReqNum(ClassID)} 将得到100<br/><font color=red>4.当个人/企业空间要使用sql标签时,可以用<font color=red>""{$GetUserName}""</font>取得当前空间的用户名 </font><br>如:select top 10 id,title from ks_article where inputer='{$GetUserName}' order by id desc<br/><font color=red>5.查询KS_ItemInfo表时,允许使用{$GetItemUrl}得到文档的URL链接</font><br/>如:当有查询到KS_ItemInfo表时,可以使用{$GetItemUrl}来得到文档的URL,但要在查询的SQL语句中包含ChannelID,InfoID,Tid,Fname四个字段。<br/>举例：select top 10 <span style='color:green'>i.channelid,i.infoid,i.tid,i.fname</span>,title,diggnum from KS_digglist d inner join ks_iteminfo i on d.infoid=i.infoid where d.channelid=i.channelid order by diggnum desc<br/>循环体:[loop=10]&lt;a href=""<span style='color:green'>{$GetItemUrl}</span>"">{$Field(title,Text,0,...,0,)}&lt;/a>[/loop]"

		 .echo "   </td></tr>"
		

		.echo "  </form>"
		.echo "</table>"
		.echo "<script language=""JavaScript"">" & vbCrLf
		.echo "<!--" & vbCrLf
		.echo "function CheckForm()" & vbCrLf
		.echo "{ var form=document.myform;"
		.echo "  if (form.LabelName.value=='')"
		.echo "   {"
		.echo "    alert('请输入标签名称!');"
		.echo "    form.LabelName.focus();"
		.echo "    return false;"
		.echo "   }"
		 .echo " if (form.LabelContent.value==''||form.LabelContent.value=='[loop="&LoopTimes&"]请在此输入循环内容[/loop]')"
		 .echo " {"
		 .echo "   alert('请输入标签循环内容!');"
		 .echo "   form.LabelContent.focus();"
		 .echo "   return false;"
		 .echo "  }"
		 .echo "  form.submit();"
		 .echo "  return true;"
		.echo "}" & vbCrLf
		.echo "//-->" & vbCrLf
		.echo "</script>"
		
		Set Conn = Nothing
		
		End With
End Sub

'保存
Sub AddLabelSave()
			LabelName = KS.G("LabelName")
			Descript = Request("LabelIntro")
			FieldParam = Request("FieldParam")
			LabelContent = Trim(Request.Form("LabelContent"))
			LabelFlag = KS.G("LabelFlag")
			FolderID = KS.G("FolderID")
			SQLType =KS.G("SQLType")
			ItemName=KS.G("ItemName")
			PageStyle=KS.G("PageStyle")
			Ajax=KS.G("Ajax")
			If LabelName = "" Then
			   Call KS.AlertHistory("标签名称不能为空!", -1)
			   Set KS = Nothing
			   Exit Sub
			End If
			If SQLType=1 And ItemName="" Then
			  Call KS.AlertHistory("分页项目不能为空!", -1)
			  Set KS = Nothing
			  Exit Sub
			End IF
			If LabelContent = "" Then
			  Call KS.AlertHistory("标签内容不能为空!", -1)
			  Set KS = Nothing
			  Exit Sub
			End If
			LabelName = "{SQL_" & LabelName & "}"
			Set LabelRS = Server.CreateObject("Adodb.RecordSet")
			LabelRS.Open "Select top 1 LabelName From [KS_Label] Where LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
			  Call KS.AlertHistory("标签名称已经存在!", -1)
			  LabelRS.Close
			  Conn.Close
			  Set LabelRS = Nothing
			  Set Conn = Nothing
			  Set KS = Nothing
			 Exit Sub
			Else
				LabelRS.Close
				LabelRS.Open "Select  top 1 * From [KS_Label] Where (ID is Null)", Conn, 1, 3
				LabelRS.AddNew
				  Do While True
					'生成ID  年+12位随机
					LabelID = Year(Now()) & KS.MakeRandom(10)
					Set RSCheck = Conn.Execute("Select ID from [KS_Label] Where ID='" & LabelID & "'")
					 If RSCheck.EOF And RSCheck.BOF Then
					  RSCheck.Close
					  Set RSCheck = Nothing
					  Exit Do
					 End If
				  Loop
				 LabelRS("ID") = LabelID
				 LabelRS("LabelName") = LabelName
				 LabelRS("LabelContent") = LabelContent
				 LabelRS("LabelFlag") = LabelFlag
				 LabelRS("Description") = Descript &"@@@"&FieldParam&"@@@"&SQLType&"@@@"&ItemName&"@@@"&PageStyle&"@@@"&Ajax&"@@@"& datasourcetype &"@@@" &datasourcestr & "@@@" & note & "@@@" & KS.G("NoRecord")
				 LabelRS("FolderID") = FolderID
				 LabelRS("AddDate") = Now
				 LabelRS("LabelType") = 5
				 LabelRS("OrderID") = 1
				 LabelRS.Update
				 Call KS.FileAssociation(1021,2,LabelContent,0)
				KS.echo "<script>$.dialog.confirm('恭喜，添加标签成功,继续添加标签吗?',function(){location.href='LabelSQL.asp?Action=AddNew&LabelType=5&FolderID=" & FolderID & "';},function(){$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & FolderID & "&OpStr=标签管理 >> 自义定函数标签&ButtonSymbol=DIYFunctionLabel';parent.frames['MainFrame'].location.href='Label_Main.asp?LabelType=5&FolderID=" & FolderID & "';});</script>"

			End If
	End Sub
	
	'保存修改
	Sub EditLabelSave()
			LabelID = Trim(Request.Form("LabelID"))
			FolderID = Request.Form("FolderID")
			LabelName = Replace(Replace(Trim(Request.Form("LabelName")), """", ""), "'", "")
			Descript = Request("LabelIntro")
			FieldParam = Request("FieldParam")
			SQLType =KS.G("SQLType")
			ItemName=KS.G("ItemName")
			PageStyle=KS.G("PageStyle")
			Ajax=KS.G("Ajax")
			Call SqlValid(Descript)
			LabelContent = Trim(Request.Form("LabelContent"))
			LabelFlag = Request.Form("LabelFlag")
			If LabelName = "" Then
			   Call KS.AlertHistory("标签名称不能为空!", -1)
			   Set KS = Nothing
			   Exit Sub
			End If
			If SQLType=1 And ItemName="" Then
			  Call KS.AlertHistory("分页项目不能为空!", -1)
			  Set KS = Nothing
			  Exit Sub
			End IF
			If LabelContent = "" Then
			  Call KS.AlertHistory("标签内容不能为空!", -1)
			  Set KS = Nothing
			  Exit Sub
			End If
			LabelName = "{SQL_" & LabelName & "}"
			Set LabelRS = Server.CreateObject("Adodb.RecordSet")
			LabelRS.Open "Select LabelName From [KS_Label] Where ID <>'" & LabelID & "' AND LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
			  Call KS.AlertHistory("标签名称已经存在!", -1)
			  LabelRS.Close:Conn.Close:Set LabelRS = Nothing:Set Conn = Nothing
			  Set KS = Nothing
			  Exit Sub
			Else
				LabelRS.Close
				LabelRS.Open "Select top 1 * From [KS_Label] Where ID='" & LabelID & "'", Conn, 1, 3
				 LabelRS("LabelName") = LabelName
				 LabelRS("LabelContent") = LabelContent
				 LabelRS("LabelFlag") = LabelFlag
				 LabelRS("FolderID") = FolderID
				 LabelRS("Description") = Descript & "@@@"&FieldParam&"@@@"&SQLType&"@@@"&ItemName&"@@@"&PageStyle&"@@@"&Ajax&"@@@"& datasourcetype &"@@@" &datasourcestr& "@@@" & note & "@@@" & KS.G("NoRecord")
				 LabelRS("AddDate") = Now
				 LabelRS.Update
				 '遍历所有标签内容，找出所有标签的图片
				 Dim Node,UpFiles,RCls
				 UpFiles=LabelContent
				 LabelRS.Close
				 LabelRS.Open "Select LabelContent From KS_Label Where LabelType=5",conn,1,1
                 Do While Not LabelRS.Eof
				     UpFiles=UpFiles & LabelRS(0)
				     LabelRS.MoveNext
				 Loop
				 LabelRS.Close
				 Set LabelRS=Nothing
				 Call KS.FileAssociation(1021,2,UpFiles,1)


				 '遍历及入库结束
				 
				 If KeyWord = "" Then
				   	KS.Echo "<script>$.dialog.tips('<br/>恭喜，标签修改成功!',1,'success.gif',function(){$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & FolderID & "&OpStr=标签管理  >> 自定义函数标签&ButtonSymbol=DIYFunctionLabel';location.href='Label_main.asp?Page=" & Page & "&LabelType=5&FolderID=" & FolderID & "';});</script>"

				 Else
				   	KS.Echo "<script>$.dialog.tips('恭喜，标签修改成功!',1,'success.gif',function(){$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?OpStr=标签管理 >> <font color=red>搜索自定义函数标签结果</font>&ButtonSymbol=DIYFunctionSearch';location.href='Label_main.asp?Page=" & Page & "&LabelType=5&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';});</script>"
				 End If
			End If
	End Sub
	
	Sub SqlValid(SqlStr)
	     On Error Resume Next
		 if datasourcetype<>0 then
		 tConn.Execute(ClearParam(SqlStr))
		 else
		 Conn.Execute(ClearParam(SqlStr))
		 end if
		 If Err Then 
		  KS.Echo "<script>alert('" & replace(err.description,"'","\'") & "');history.back();</script>"
		  response.end
		 End If
	End Sub
	function ClearParam(byval SqlStr)
	     sqlstr=Lcase(sqlstr)
	     if instr(SqlStr,"where")<>0 then sqlstr=split(SqlStr,"where")(0)
	     Dim I
		 For I=0 To 100
		  SqlStr=Replace(SqlStr,"{$param(" & I & ")}",1)
		 Next
		  SqlStr=Replace(SqlStr,"{$currclasschildid}","'1'")
		  SqlStr=Replace(SqlStr,"{$currchannelid}",1)
		  SqlStr=Replace(SqlStr,"{$currclassid}",1)
		  SqlStr=Replace(SqlStr,"{$currinfoid}",1)
		  SqlStr=Replace(SqlStr,"{$currspecialid}",1)
		  SqlStr=Replace(SqlStr,"{$getusername}",1)
		  ClearParam=ReplaceRequest(SqlStr)
		 exit function
     End function
	 
'替换request的值,支持ReqNum和ReqStr两个标签
		Function ReplaceRequest(Content)
		     Dim regEx, Matches, Match,TempStr,QStr,ReqType
			 Set regEx = New RegExp
			 regEx.Pattern= "{(ReqNum|ReqStr)[^{}]*}"
			 regEx.IgnoreCase = True
			 regEx.Global = True
			 Set Matches = regEx.Execute(Content)
			 For Each Match In Matches
				On Error Resume Next
				TempStr = Match.Value
				ReqType=Split(TempStr,"(")(0)
				QStr=Replace(Split(TempStr,"(")(1),")}","")
				If ReqType="{ReqNum" Then
				 Content=Replace(Content,TempStr,1)
				Else
				 Content=Replace(Content,TempStr,"1")
				End If
			Next
			ReplaceRequest=Content
		End Function

  
 Function GetLoopTimes(SqlStr)
		 Dim regEx, Matches, Match
		 Set regEx = New RegExp
		 regEx.Pattern = "top\s?[\d]*\d"
		 regEx.IgnoreCase = True
		 regEx.Global = True
		 Set Matches = regEx.Execute(SqlStr)
		 If Matches.count > 0 Then 
		  GetLoopTimes=Trim(Split(lcase(Matches.item(0)),"top")(1))
         End If
		 regEx.Pattern = "top\s?{\$Param\([^}]*}"
		 'regEx.Pattern = "top[^}]*}"
		 regEx.IgnoreCase = True
		 regEx.Global = True
		 Set Matches = regEx.Execute(SqlStr)
		 If Matches.count > 0 Then 
		  GetLoopTimes=Trim(Split(Matches.item(0),"top")(1))
         End If
		If GetLoopTimes="" Then GetLoopTimes=10
  End Function
  

	
	Sub OpenExtConn()
		if datasourcetype<>0 then 
		   '外部access自动转换相对路径为绝对路径
		   Dim connstr:connstr=datasourcestr
		   if datasourcetype=1 or datasourcetype=5 or datasourcetype=6 Then connstr=LFCls.GetAbsolutePath(connstr)  
		   if not isobject(tconn) then
			on error resume next
		    Set tconn = Server.CreateObject("ADODB.Connection")
			tconn.open connstr
			If Err Then 
			  Err.Clear
			  Set tconn = Nothing
			  KS.Echo "<script>alert('外部数据库连接失败!');history.back();</script>"
			  response.end
			end if
		   end if
		end if
	End Sub
	
	Function ReplaceBC(ByVal C)
	 C=Replace(C,"'","")
	 C=Replace(C,"(","")
	 C=Replace(C,")","")
	 ReplaceBC=C
	End Function
	
	Sub ShowClassID()
	%>
		 <script type="text/javascript">
						  function copyToClipboard(txt) {
							 if(window.clipboardData) {
									 window.clipboardData.clearData();
									 window.clipboardData.setData("Text", txt);
							 } else if(navigator.userAgent.indexOf("Opera") != -1) {
								  window.location = txt;
							 } else if (window.netscape) {
								  try {
									   netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");
								  } catch (e) {
									   alert("被浏览器拒绝！\n请在浏览器地址栏输入'about:config'并回车\n然后将'signed.applets.codebase_principal_support'设置为'true'");
								  }
								  var clip = Components.classes['@mozilla.org/widget/clipboard;1'].createInstance(Components.interfaces.nsIClipboard);
								  if (!clip)
									   return;
								  var trans = Components.classes['@mozilla.org/widget/transferable;1'].createInstance(Components.interfaces.nsITransferable);
								  if (!trans)
									   return;
								  trans.addDataFlavor('text/unicode');
								  var str = new Object();
								  var len = new Object();
								  var str = Components.classes["@mozilla.org/supports-string;1"].createInstance(Components.interfaces.nsISupportsString);
								  var copytext = txt;
								  str.data = copytext;
								  trans.setTransferData("text/unicode",str,copytext.length*2);
								  var clipid = Components.interfaces.nsIClipboard;
								  if (!clip)
									   return false;
								  clip.setData(trans,null,clipid.kGlobalClipboard);
							 }
								  alert("复制成功！")
						}
		 </script>
	 <body class="tdbg">
	 <table width="100%" cellpadding="0" cellspacing="0">
	   <tr><td colspan="4" align="center" height="25" class="title"><strong>(栏 目 <=> ID)对 照 表</strong></td></tr>
	   <tr class="tdbg">
		<td colspan=4>
		  <table border=0>
		   <tr>
		   <td width="30"></td>
		   <td><%
		   GetClassIDTable()
		   %></td>
		   </tr>
		   </table>
		</td>
	   </tr>
	 <table>
	 </body>
	<%
	End Sub
	
  Function GetClassIDTable()
  
		Dim Node,K,SQL,NodeText,Pstr,TJ,SpaceStr
		KS.LoadClassConfig()
		For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1]")
		      SpaceStr=""
			  TJ=Node.SelectSingleNode("@ks10").text
			  If TJ>1 Then
				 For k = 1 To TJ - 1
					SpaceStr = SpaceStr & "──"
				 Next
				KS.Echo "<li>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & SpaceStr & Node.SelectSingleNode("@ks1").text & "&nbsp;&nbsp;&nbsp;"  & Node.SelectSingleNode("@ks0").text & " <input type='button' value='复制' class='button' onclick=""copyToClipboard('"&Node.SelectSingleNode("@ks0").text&"')""></li>"
			  Else
				KS.Echo "<li><img src='../images/folder/domain.gif' align='absmiddle'>" & Node.SelectSingleNode("@ks1").text & "&nbsp;&nbsp;&nbsp;&nbsp;" & Node.SelectSingleNode("@ks0").text & " <input type='button' value='复制' class='button' onclick=""copyToClipboard('"&Node.SelectSingleNode("@ks0").text&"')""></li>"
			  End If
		Next
	End Function
End Class
%> 
