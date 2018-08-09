<%

'出现频率较高的标签直接在这里替换掉,提高性能
If Instr(F_C,"{$GetInstallDir}")<>0 Then F_C=Replace(F_C,"{$GetInstallDir}",KS.Setting(3))
If Instr(F_C,"{$GetSiteUrl}")<>0 Then F_C=Replace(F_C,"{$GetSiteUrl}",Domainstr)

%>