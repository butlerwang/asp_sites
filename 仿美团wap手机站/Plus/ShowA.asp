<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%

Dim KSCls
Set KSCls = New Advertise
KSCls.Kesion()
Set KSCls = Nothing

Class Advertise
        Private KS
		Private getplace,getshow,adsrs,adssql,adsrsp,adssqlp,adsrss,adssqls,getip,getggwlxsz,getggwhei,getggwwid
        Private ttarg,DomainStr,GaoAndKuan,advertvirtualvalue
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Sub Kesion()
		  Select Case KS.S("Action")
		   Case "Daima"
		     Call AdvertiseDaima()
		   Case "AdOpen"
		     Call AdvertiseAdOpen()
		  End Select
		End Sub
		
 '代码
  Sub AdvertiseDaima()
         response.write "<body>"
  	    if KS.S("id")<>"" and isnumeric(KS.S("ID")) then
			dim adssql
			dim adsrs:set adsrs=server.createobject("adodb.recordset")
			adssql="Select top 1 intro from KS_Advertise where id="&KS.ChkClng(KS.S("id"))&" order by time"
			adsrs.open adssql,conn,1,1       
			if not adsrs.eof then
			response.write adsrs(0)
			end if
			adsrs.close:set adsrs=nothing
			conn.close:set conn=nothing
		else
			response.write "<center><br><br>无效广告。</center>"
		end if
		response.write "</body>"
  End Sub

 Sub AdvertiseAdOpen()
 %>
     <html>
	 <head>
	 <script type="text/javascript" src="../ks_inc/jquery.js"></script>
	 <script type="text/javascript">
	 function addHits(c,id){if(c==1){try{jQuery.getScript('ajaxs.asp?action=HitsGuangGao&id='+id,function(){});}catch(e){}}}
	 </script>
	 <style type="text/css">
	 body{font-size:12px}
	 
	 </style>
	 </head>
	 <body topmargin="0" leftmargin="0">
	<%
	Dim DomainStr:DomainStr=KS.GetDomain
	Dim ttarg:ttarg="_top"
	Dim GaoAndKuan:GaoAndKuan=""
	Dim Adsrs:Set adsrs=server.createobject("adodb.recordset")
	Dim adssql:adssql="Select top 1 id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei,clicks,url from KS_Advertise where id="&KS.Chkclng(KS.S("i"))
	adsrs.open adssql,Conn,3,3
	adsrs("show")=adsrs("show")+1
	adsrs("time")=now()
	adsrs.Update
	if adsrs("window")=0 then
	ttarg = "_blank"
	end if
	
	if isnumeric(adsrs("hei")) then
	GaoAndKuan=" height="&adsrs("hei")&" "
	else
	
	if right(adsrs("hei"),1)="%" then
	if isnumeric(Left(len(adsrs("hei"))-1))=true then
	 GaoAndKuan=" height="&adsrs("hei")&" "
	end if
	end if
	
	end if
	
	if isnumeric(adsrs("wid")) then
	GaoAndKuan=GaoAndKuan&" width="&adsrs("wid")&" "
	else
	if right(adsrs("wid"),1)="%" then
	if isnumeric(Left(len(adsrs("wid"))-1))=true then 
	GaoAndKuan=GaoAndKuan&" width="&adsrs("wid")&" "
	end if
	end if
	end if
	
	 Select Case adsrs("xslei")
				Case "txt"%>
				<span onClick="addHits(<%=adsrs("clicks")%>,<%=adsrs("id")%>)"><a title="<%=adsrs("sitename")%>" href="<%=adsrs("url")%>" target="<%=ttarg%>"><%=adsrs("sitename")%></a></span>
	<%          Case "gif"%>
	                <span onClick="addHits(<%=adsrs("clicks")%>,<%=adsrs("id")%>)">
	                <a title="<%=adsrs("intro")%>" href="<%=adsrs("url")%>" target="<%=ttarg%>"><img border=0  <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>" /></a> 
				    </span>
	<%          Case "swf"%><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http:/download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"; <%=GaoAndKuan%>><param name=movie value="<%=adsrs("gif_url")%>"><param name=quality value=high>
	  <%          Case "dai"%><%=adsrs("intro")%>
	  <embed src="<%=adsrs("gif_url")%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"></embed></object>
	<%          Case else%><a title="<%=adsrs("intro")%>" href="<%=adsrs("url")%>" target="<%=ttarg%>"><img border=0  <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>" /></a>
	<%
	 End Select%>
	 <%
	adsrs.close
	set adsrs=nothing
	Conn.close
	set Conn=nothing 
	%>
	 </body>
	</html>
<%
 End Sub
End Class
 %>  
