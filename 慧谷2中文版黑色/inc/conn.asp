<%
dim cn,connstr
Set cn=Server.CreateObject("ADODB.Connection")
DataName="#NYIKUGY5434231#.ini" '���ݿ�����
DataFolder="/huigurdata" '���ݿ����ļ��У�һ���ڸ�Ŀ¼��
DataPath=DataFolder&"/"&DataName '���ݿ��·��
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DataPath)
cn.open connstr

%>