<!--#INCLUDE FILE="data.asp" -->
<%

	dim id
	id = request("id")

set rs=server.CreateObject("adodb.recordset")
sql="SELECT * FROM userinfo WHERE id =1"
rs.Open sql,conn,1,1
response.contenttype="x-mixed-replace"
Response.BinaryWrite rs("photo")
rs.Close


function ImageUp(formsize,formdata)          '��������Ĺ����ǽ�ȡ���е�ͼ�񲿷֡�
    bncrlf=chrb(13) & chrb(10)               '���ɺ������Ժ�������Լ�����ʹ���ˡ�
    divider=leftb(formdata,instrb(formdata,bncrlf)-1)
    datastart=instrb(formdata,bncrlf&bncrlf)+4
    dataend=instrb(datastart+1,formdata,divider)-datastart
    imageup=midb(formdata,datastart,dataend)
end function

'-------------------------
%>