<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceApp.asp"-->
<%

Dim KS,KSBCls,KSR
Set KS=New PublicCls
Set KSBCls=New BlogCls
dim TemplateID,Tp
TemplateID=KS.ChkClng(KS.S("TemplateID"))
Tp=KSBCls.GetTemplatePath(TemplateID,"TemplateMain")
Tp=Replace(tp,"{$GetInstallDir}",KS.Setting(3))
Tp=Replace(tp,"{$GetSiteUrl}",KS.Setting(2))
KS.Echo Tp
Set KS=Nothing
Set KSBCls=Nothing
call closeconn()
%>