<%
dim cn,connstr
Set cn=Server.CreateObject("ADODB.Connection")
DataName="#fgxfnchbsfdgdfgfdg.mdb" '���ݿ�����
DataFolder="/D21293" '���ݿ����ļ��У�һ���ڸ�Ŀ¼��
DataPath=DataFolder&"/"&DataName '���ݿ��·��
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DataPath)
cn.open connstr

%>