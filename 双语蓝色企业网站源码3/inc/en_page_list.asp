<div class="t_page">
<%'===========��ʾ��ҳ�Ĺ��̵��ã�Ҫ�������ݿ�򿪺��ͷ���Դǰ
call listPages()  
rs.close  '�ͷ���Դ
set rs=nothing%> 

 <%'==========��ҳ���̿�ʼ��Ҳ�ɵ�������һ�ļ��ڱ��ļ�ǰ��������
 sub listPages() '������̿�ʼ%>
<% 
	   			 n=cint(request.querystring("page"))
				  if n=0 then 
 n=1
 end if
response.write "Current Page��<span class='FontRed'>"&n&"</span>/"&maxpagecount&"&nbsp;<a href=?page=1&q="&keywords_all&" >Home</a> " 

if n>=2 then
response.write"<a href=?page="&n-1&"&q="&keywords_all&" title='to page"&n-1&"'>Pre</a> "
end if

for i=pagestart to pageend
            if i=0 then 
            i=1
            end if
if n=i then 
classi="class='black_link'" 
else
classi=""
end if
            strurl="<span "&classi&"><a href=?page="&i&"&q="&keywords_all&" title='to page"&i&"' >"&i&"</a></span>"
response.write strurl
response.write " "

 next

 if n<>pageend then
 n=n+1
 end if
  response.write"<a href=?page="&n&"&q="&keywords_all&" title='to page"&n&"'>next</a>"

            %> 

<%end sub '������̽���
'==========��ҳ���̽���%>	
</div>