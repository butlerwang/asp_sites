<!--#include file="../check.asp"-->
<!--#include file="../include/conn.asp"-->
<!--#include file="../include/lib.asp"-->
<%
Table_name=request.querystring("Table_name")
ItemID=request.querystring("ItemID")
intID=request.querystring("id")
Call OpenData()
Call Del(Table_name,ItemID,intID)
Call CloseDataBase()
Call page_back("É¾³ý³É¹¦")
%>