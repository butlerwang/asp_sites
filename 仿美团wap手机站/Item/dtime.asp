<!--#include file="../Conn.asp"-->
<%

dim dtime,id

id=request("id")

if not isnumeric(id) then response.end

dtime=conn.execute("select adddate from ks_article where id=" & id)(0)

response.write "document.writeln('" & GetTimeFormat(dtime) & "');"

Function GetTimeFormat(DateTime)
        if DateDiff("n",DateTime,now)<5 then
      GetTimeFormat="刚刚"
     elseif DateDiff("n",DateTime,now)<60 then
      GetTimeFormat=DateDiff("n",DateTime,now) & " 分钟前"
     elseif DateDiff("h",DateTime,now)<5 Then
      GetTimeFormat=DateDiff("h",DateTime,now) & " 小时前"
     else
      GetTimeFormat=formatdatetime(DateTime,2)
     end if
 End Function
%>
