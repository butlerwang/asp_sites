<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="page_next.asp" -->

<%
'�����ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FileName,FolderName from web_Models_type where [id]=5"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Model_FileName=rs_1("FileName")
if rs_1("FolderName")<>"" then
Model_FolderName="/"&rs_1("FolderName")
end if
end if
rs_1.close
set rs_1=nothing%>
<!-- ���������˵� ��ʼ -->
<script language="JavaScript">
<!--
<%
'�������ݱ��浽����
Dim count2,rsClass2,sqlClass2
set rsClass2=server.createobject("adodb.recordset")
sqlClass2="select id,pid,ppid,name from category where ppid=2 and ClassType=1 order by id " 
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
sqlClass3="select id,pid,ppid,name from category where ppid=3 and ClassType=1 order by id" 
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
</script><!-- ���������˵� ���� -->
<% '����ģ��
act=request.querystring("act")
juhaoyongSearchKeywordsBiaoji=request.form("juhaoyongSearchKeywordsBiaoji")
keywords=trim(request.form("keywords"))

if juhaoyongSearchKeywordsBiaoji="1" then
cid=request.form("cid")
pid=request.form("pid")
ppid=request.form("ppid")
else
cid=request.QueryString("cid")
pid=request.QueryString("pid")
ppid=request.QueryString("ppid")
end if

if cid="-1" then cid=""
if pid="-1" then pid=""
if ppid="-1" then ppid=""

if juhaoyongSearchKeywordsBiaoji="1" then

	if cid="" and pid="" and  ppid="" then
	s_sql="select id,title,cid,pid,ppid,file_path,image,index_push,slide_yes,special_yes,view_yes,headline,hit,ip,time from article where [title] like '%"&keywords&"%' and ArticleType=1 order by id desc"
	elseif pid="" and ppid="" then
	search_sql=" and cid='"&cid&"'"
	s_sql="select id,title,cid,pid,ppid,file_path,image,index_push,slide_yes,special_yes,view_yes,headline,hit,ip,time from article where [title] like '%"&keywords&"%'"&search_sql&" and ArticleType=1  order by id desc"
	elseif ppid="" then
	search_sql=" and pid='"&pid&"'"
	s_sql="select id,title,cid,pid,ppid,file_path,image,index_push,slide_yes,special_yes,view_yes,headline,hit,ip,time from article where [title] like '%"&keywords&"%'"&search_sql&" and ArticleType=1  order by id desc"
	else
	search_sql=" and ppid='"&ppid&"'"
	s_sql="select id,title,cid,pid,ppid,file_path,image,index_push,slide_yes,special_yes,view_yes,headline,hit,ip,time from article where [title] like '%"&keywords&"%'"&search_sql&" and ArticleType=1  order by id desc"
	end if

else
	s_sql="select id,title,cid,pid,ppid,file_path,image,index_push,slide_yes,special_yes,view_yes,headline,hit,ip,time from article where  cid='"&cid&"' and pid='"&pid&"' and ppid='"&ppid&"' and ArticleType=1  order by id desc"

end if

%>
<script language="JavaScript">
<!--
function ask(msg) {
	if( msg=='' ) {
		msg='���棺ɾ���󽫲��ɻָ���ȷ��Ҫɾ����';
	}
	if (confirm(msg)) {
		return true;
	} else {
		return false;
	}
}
//-->
</script>
<script language="javascript">

//ȫѡJS
function unselectall(){
if(document.form2.chkAll.checked){
document.form2.chkAll.checked = document.form2.chkAll.checked&0;
}
}
function CheckAll(form){
for (var i=0;i<form.elements.length;i++){
var e = form.elements[i];
if (e.Name != 'chkAll'&&e.disabled==false)
e.checked = form.chkAll.checked;
}
}
</script>	<%
Call header()
%>

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>�����б�</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='TipTitle'>&nbsp;�� ������ʾ</td>
          </tr>
          <tr>
            <td height="30" valign="top" class="TipWords"><p>1�������б���ʾ������ӵ��������£���ʾ��δ��ˡ������½���������վ����ʾ��</p>
                <p>2��ɾ�����½���ͬ��ɾ�����ݿ��еļ�¼�����µľ����ַ�����ء�</p>
            </td>
          </tr>
          <tr>
            <td height="10" ></td>
          </tr>
        </table>

	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| <a href="article_add.asp?juhaoyong_cid=<%=cid%>&juhaoyong_pid=<%=pid%>&juhaoyong_ppid=<%=ppid%>">����µ�����</a></td>
          </tr>
      </table>

<ul class=HoutaiArticleDaohang><li>
<%if cid<>"" then%>
&nbsp;<img src=images/juhaoyongdaohangjiantou.gif border="0" />��<a href=Article_list.asp?cid=<%=cid%>&act=search><%=juhaoyongGetLanmuName(cid)%></a>
<%end if%>

<%if pid<>"" then%>
>> <a href=Article_list.asp?cid=<%=cid%>&pid=<%=pid%>&act=search><%=juhaoyongGetLanmuName(pid)%></a>
<%end if%>

<%if ppid<>"" then%>
>> <a href=Article_list.asp?cid=<%=cid%>&pid=<%=pid%>&ppid=<%=ppid%>&act=search><%=juhaoyongGetLanmuName(ppid)%></a>
<%end if%>
</li></ul>

<%
if cid<>"" and (pid="" and ppid="") then
juhaoyongGetLowerFolderNameList=juhaoyongGetLowerFolderNameList(2,cid,cid)
elseif ppid="" then 
juhaoyongGetLowerFolderNameList=juhaoyongGetLowerFolderNameList(3,pid,cid)
end if
%>
 <form name="form2" method="post" action="Article_Del.asp?action=AllDel&juhaoyong_cid=<%=cid%>&juhaoyong_pid=<%=pid%>&juhaoyong_ppid=<%=ppid%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">
 	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="2">
          <tr>
            <%if juhaoyongSearchKeywordsBiaoji<>"1" then%><td width="2%" height="30" class="TitleHighlight">&nbsp;</td><%end if%>
            <td width="4%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">���</div></td>
            <td width="33%" height="30" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">���±���</div></td>
            <td width="24%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">���·���</div></td>
            <td width="6%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">ip/pv</div></td>
            <td width="6%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">���</div></td>
            <td width="17%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">���ʱ��</div></td>
            <td width="8%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">���²���</div></td>
          </tr>
<% '�����б�ģ��
strFileName="article_list.asp" 
pageno=25
set rs = server.CreateObject("adodb.recordset")
rs.Open (s_sql),cn,1,1
rscount=rs.recordcount
if not rs.eof and not rs.bof then
call showsql(pageno)
rs.move(rsno)
for p_i=1 to loopno
%>
<% if p_i mod 2 =0 then
class_style="forumRow"
else
class_style="forumRowHighLight"
end if%>
<%

%>
          <tr >
            <%if juhaoyongSearchKeywordsBiaoji<>"1" then%><td   height="30" class='<%=class_style%>'><div align="center"><input type="checkbox" name="Selectitem" id="Selectitem" value="<%=rs("id")%>"></div></td><%end if%>
            <td   height="30" class='<%=class_style%>'><div align="center"><%=rs("id")%></div></td>
            <td class='<%=class_style%>' >&nbsp;<a href="<%=Model_FolderName&"/"&rs("file_path")%>" target="_blank" title="<%=rs("title")%>"><%=left(rs("title"),26)%></a><%if rs("image")<>"" then%>&nbsp;[<span style="color: #FF0000">ͼ</span>]<%end if%><%if rs("index_push")=1 then%>&nbsp;[<span style="color: #FF0000">��</span>]<%end if%><%if rs("slide_yes")=1 then%>&nbsp;[<span style="color: #FF0000">�õ�</span>]<%end if%><%if rs("special_yes")=1 then%>&nbsp;[<span style="color: #FF0000">ר��</span>]<%end if%></td>
            <td class='<%=class_style%>' >&nbsp;
			<% '������ʾ
			cid=cint(rs("cid"))

			set rs1=server.createobject("adodb.recordset")
			sql="select name from category where id="&cid&""
			rs1.open(sql),cn,1,1
			if not rs1.eof and not rs1.bof then
			response.write rs1("name")
			response.write "&nbsp;>&nbsp;"
			end if
			rs1.close

			
			if rs("pid")<>"" then
            pid=cint(rs("pid"))
			sql="select name from category where id="&pid&""
			rs1.open(sql),cn,1,1
				if not rs1.eof and not rs1.bof then
				response.write rs1("name")
				response.write "&nbsp;>&nbsp;"
				end if
			rs1.close
			end if
						
			
			if rs("ppid")<>"" then
            ppid=cint(rs("ppid"))
			sql="select name from category where id="&ppid&""
			rs1.open(sql),cn,1,1
				if not rs1.eof and not rs1.bof then
				response.write rs1("name")
				end if
			rs1.close
			end if
			set rs1=nothing
			%>            </td>
            <td class='<%=class_style%>' ><div align="center"><%=rs("ip")%>/<%=rs("hit")%></div></td>
            <td class='<%=class_style%>' ><div align="center"><a href="article_view_yes.asp?id=<%=rs("id")%>&juhaoyong_cid=<%=rs("cid")%>&juhaoyong_pid=<%=rs("pid")%>&juhaoyong_ppid=<%=rs("ppid")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>"><%if rs("view_yes")=1 then%>�����<%else%><span style="color: #FF0000">δ���</span><% end if%></a></div></td>
            <td class='<%=class_style%>' ><div align="center"><%=rs("time")%></div></td>
            <td class='<%=class_style%>' >
            <div align="center"><a href="article_edit.asp?id=<%=rs("id")%>&juhaoyong_cid=<%=rs("cid")%>&juhaoyong_pid=<%=rs("pid")%>&juhaoyong_ppid=<%=rs("ppid")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>">�޸�</a> | <a href="javascript:if(ask('���棺ɾ���󽫲��ɻָ���ȷ��Ҫɾ����')) location.href='article_del.asp?id=<%=rs("id")%>&juhaoyong_cid=<%=rs("cid")%>&juhaoyong_pid=<%=rs("pid")%>&juhaoyong_ppid=<%=rs("ppid")%>&page=<%=page%>&act=<%=act%>&keywords=<%=keywords%>';">ɾ��</a>            </div></td>
          </tr>
		  		  <%
		  rs.movenext
		  next
		  else
response.write "<div align='center'><span style='color: #FF0000'>�������£�</span></div>"
		  end if 
		  rs.close
		  set rs=nothing
		  %>
		     <%if juhaoyongSearchKeywordsBiaoji<>"1" then%><tr><td height="35"  colspan="8" >&nbsp;<input name='chkAll' type='checkbox' id='chkAll' onclick='CheckAll(this.form)' value='checkbox'>
                    ȫѡ/ȫ��ѡ&nbsp;<input type="submit" name="Submit" value="ɾ��ѡ��"></td></tr>
			 <%end if%>
            <tr  >
              <td height="35"  colspan="8" ><div align="center">
                <%call showpage(strFileName,rscount,pageno,false,true,"")%>
           </div></td>
		    </tr>
      </table> </form>  
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="20" class='forumRow'>&nbsp;</td>
          </tr>
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| ��������</td>
          </tr>
          <tr>
            <td height="70"><form name="form1" method="post" action="?act=search">
              <div align="center"><%
Dim count1,rsClass1,sqlClass1
set rsClass1=server.createobject("adodb.recordset")
sqlClass1="select id,pid,ppid,name from category where ppid=1 and ClassType=1 order by id" 
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
            </select>&nbsp;
            <label>
<input name="keywords" type="text"  size="35" maxlength="40">
              </label>
                <label>
                       &nbsp;<input name="juhaoyongSearchKeywordsBiaoji" type="hidden" value="1">
                       <input type="submit" name="Submit" value="�� ��">
                </label>
              </div>
            </form>
            </td>
          </tr>
      </table>
	    <br></td>
	</tr>
	</table>

<%
Call DbconnEnd()
 %>