<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Article_list_to_html.asp" -->
<!-- #include file="../inc/x_to_html/post_index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Blank_Content_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Recruit_list_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Case_List_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Search_index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/SiteMap_index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/order_index_to_html.asp" -->

<link rel="stylesheet" type="text/css" href="/css/common.css"  />

<script language="javascript">

//ȫѡJS
function unselectall(){
if(document.form1.chkAll.checked){
document.form1.chkAll.checked = document.form1.chkAll.checked&0;
}
}
function CheckAll(form){
for (var i=0;i<form.elements.length;i++){
var e = form.elements[i];
if (e.Name != 'chkAll'&&e.disabled==false)
e.checked = form.chkAll.checked;
}
}
</script>



<!-- ��Ѷ�������������˵� ��ʼ -->
<script language="JavaScript">
<!--
<%
'�������ݱ��浽����
Dim count2,rsClass2,sqlClass2
set rsClass2=server.createobject("adodb.recordset")
sqlClass2="select id,pid,ppid,name from category where ppid=2   order by id " 
rsClass2.open sqlClass2,cn,1,1
%>
var subval2 = new Array();
//����ṹ��һ����ֵ,������ֵ,������ʾֵ
<%
count2 = 0
do while not rsClass2.eof
%>
subval2[<%=count2%>] = new Array('<%=rsClass2("pID")%>','<%=rsClass2("ID")%>','<%=rsClass2("Name")%>')
<%
count2 = count2 + 1
rsClass2.movenext
loop
rsClass2.close
%>

<%
'�������ݱ��浽����
Dim count3,rsClass3,sqlClass3
set rsClass3=server.createobject("adodb.recordset")
sqlClass3="select id,pid,ppid,name from category where ppid=3   order by id" 
rsClass3.open sqlClass3,cn,1,1
%>
var subval3 = new Array();
//����ṹ��������ֵ,������ֵ,������ʾֵ
<%
count3 = 0
do while not rsClass3.eof
%>
subval3[<%=count3%>] = new Array('<%=rsClass3("pID")%>','<%=rsClass3("ID")%>','<%=rsClass3("Name")%>')
<%
count3 = count3 + 1
rsClass3.movenext
loop
rsClass3.close
%>

function changeselect1(locationid)
{
    document.form1.pid.length = 0;
    document.form1.pid.options[0] = new Option('ѡ���������','');
    document.form1.ppid.length = 0;
    document.form1.ppid.options[0] = new Option('ѡ����������','');
    for (i=0; i<subval2.length; i++)
    {
        if (subval2[i][0] == locationid)
        {document.form1.pid.options[document.form1.pid.length] = new Option(subval2[i][2],subval2[i][1]);}
    }
}

function changeselect2(locationid)
{
    document.form1.ppid.length = 0;
    document.form1.ppid.options[0] = new Option('ѡ����������','');
    for (i=0; i<subval3.length; i++)
    {
        if (subval3[i][0] == locationid)
        {document.form1.ppid.options[document.form1.ppid.length] = new Option(subval3[i][2],subval3[i][1]);}
    }
}
//-->
</script><!-- ��Ѷ�������������˵� ���� -->

	<%
Call header()
%>

<%'����

if request.querystring("action")="create" then

'������ҳ
if request.form("index")=1 then
call index_to_html()
end if

'������Ŀ
if request.form("Class_List")=1 then

cid=request.form("cid")
pid=request.form("pid")
ppid=request.form("ppid")

if cid<>"" or pid<>"" or ppid<>""  then

if ppid<>"" then
n_sql=" id="&ppid&" "
end if

if pid<>"" and ppid="" then
n_sql=" id="&pid&" "
end if

if cid<>"" and pid="" and ppid="" then
n_sql=" id="&cid&" "
end if
sql="select [id],ppid,ClassType,Html_Yes,index_push from [category] where "&n_sql&"    order by [time] desc"
else
sql="select [id],ppid,ClassType,Html_Yes,index_push from [category]  order by [time] desc"
end if


set rs_create=server.createobject("adodb.recordset")
rs_create.open(sql),cn,1,1
if not rs_create.eof then
do while not rs_create.eof
ClassID=rs_create("id")

'����
if rs_create("ClassType")=1 then
call Article_list_to_html(ClassID)
end if

'��Ʒ
if rs_create("ClassType")=2 then
call Case_List_to_html(ClassID)
end if

'��Ƹ
if rs_create("ClassType")=4 then
call Recruit_list_to_html(ClassID)
end if

'��ҳ
if rs_create("ClassType")=5  then
call Blank_Content_to_html(ClassID)
end if

rs_create.movenext
loop
end if
rs_create.close
set rs_create=nothing
end if


'����������ҳ���б�
if request.form("post_index")=1 then
call post_index_to_html()
end if

'��������ҳ
if request.form("search_index")=1 then
call search_index_to_html()
end if

'������վ��ͼ
if request.form("SiteMap_index")=1 then
call SiteMap_to_html()
end if

'���ɶ���ҳ
if request.form("order_index")=1 then
call order_index_to_html()
end if


response.Write "<script language='javascript'>alert('���³ɹ���');history.go(-1);</script>"

end if
%>


	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>������Ŀ</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='TipTitle'>&nbsp;�� ������ʾ</td>
          </tr>
          <tr>
            <td height="30" valign="top" class="TipWords"><p>1����Ŀ������Ҫ����������ҳ������б�ҳ������ɡ�</p>
                <p>2��������ȫѡ�������ɣ���Ŀ���࣬����ʱ����Ի�������³�ʱ��</p></td>
          </tr>
          <tr>
            <td height="10" ></td>
          </tr>
        </table>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="100">
              <form name="form1" method="post" action="?action=create">
                <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
				      <tr > <td height="32" colspan="2"  class="TitleHighlight3"> &nbsp;
                      <label>
                      <input type="checkbox" name="index" value="1">
                      <span style="font-weight: bold">������ҳ</span></label></td>
                  </tr>
				      <tr>
                        <td height="10" colspan="2" class="forumRowt" style="line-height:200%;"></td>
				      </tr>	

				      <tr >
                    <td height="32" colspan="2"  class="TitleHighlight3"> &nbsp;
                      <label>
                      <input type="checkbox" name="Class_List" value="1">
                      <span style="font-weight: bold">������Ŀ</span></label></td>
                  </tr>
				      <tr>
                    <td height="35" colspan="2" class="forumRowt" style="line-height:200%;"><table width="100%" border="0" cellspacing="3" cellpadding="0">
                      <tr class="forumRowHighLight">
                        <td width="8%" height="30" class="forumRowHighLight">&nbsp;ѡ����Ŀ</td>
                        <td width="92%" class="forumRowHighLight">&nbsp;
                            <%
Dim count1,rsClass1,sqlClass1
set rsClass1=server.createobject("adodb.recordset")
sqlClass1="select id,pid,ppid,name from category where ppid=1  order by id" 
rsClass1.open sqlClass1,cn,1,1
%>
                            <select name="cid" id="cid" onChange="changeselect1(this.value)">
                              <option value="">ѡ��һ������</option>
                              <%
count1 = 0
do while not rsClass1.eof
response.write"<option value="&rsClass1("ID")&">"&rsClass1("Name")&"</option>"
count1 = count1 + 1
rsClass1.movenext
loop
rsClass1.close
%>
                            </select>
                          &nbsp;&nbsp;
                          <select name="pid" id="pid"  onchange="changeselect2(this.value)">
                            <option value="">ѡ���������</option>
                          </select>
                          &nbsp;&nbsp;
                          <select name="ppid" id="ppid">
                            <option value="">ѡ����������</option>
                        </select></td>
                      </tr>
                 
                    </table>                     </td>
                  </tr>
			      <tr > <td height="32" colspan="2"  class="TitleHighlight3"> &nbsp;
             <input type="checkbox" name="post_index" value="1">
                      <span style="font-weight: bold">�ÿ�����</span></td>
                  </tr>
				      <tr>
                        <td height="10" colspan="2" class="forumRowt" style="line-height:200%;"></td>
				      </tr>					  
				  			
			      <tr > <td height="32" colspan="2"  class="TitleHighlight3"> &nbsp;
             <input type="checkbox" name="search_index" value="1">
                      <span style="font-weight: bold">����ҳ��</span></td>
                  </tr>
				      <tr>
                        <td height="10" colspan="2" class="forumRowt" style="line-height:200%;"></td>
				      </tr>	
			      <tr > <td height="32" colspan="2"  class="TitleHighlight3"> &nbsp;
             <input type="checkbox" name="SiteMap_index" value="1">
                      <span style="font-weight: bold">��վ��ͼ</span></td>
                  </tr>
				      <tr>
                        <td height="10" colspan="2" class="forumRowt" style="line-height:200%;"></td>
				      </tr>	                      
 			      <tr > <td height="32" colspan="2"  class="TitleHighlight3"> &nbsp;
             <input type="checkbox" name="order_index" value="1">
                      <span style="font-weight: bold">����ҳ��</span></td>
                  </tr>
				      <tr>
                        <td height="10" colspan="2" class="forumRowt" style="line-height:200%;"></td>
				      </tr>	                       	  
                  <tr>
                    <td width="12%" height="50"><label>
                      &nbsp;
                      <input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>
                    ȫѡ/ȫ��ѡ</label></td>
                    <td width="88%"><div align="center">
                      <input type="submit" name="Submit" value="�ύ">
                    </div></td>
                  </tr>
                </table>
              </form>            </td>
          </tr>
      </table>
	    </td>
	</tr>
	</table>

<%
Call DbconnEnd()
 %>