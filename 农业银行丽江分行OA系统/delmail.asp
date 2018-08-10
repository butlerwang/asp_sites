<%
'=====删除文件函数======================================================================================	
function FileDel(FileName)

		Dim fso, f2
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		If fso.FileExists(Server.Mappath("file/"+FileName)) Then
			Set f2 = fso.GetFile(Server.Mappath("file/"+FileName))
			f2.Delete
			FileDel=1
		else
			FileDel=2
		end if
		set f2=nothing
		set fso=nothing

end function
'=====删除文件函数结束===========================================================================================

'======清空信箱附件=================================================================================================================
function DelAll(box)
			
			Record2.open("select iaddfile from "+box+rs("用户名")+" ") 'where iaddfile<>''

			while not Record2.eof
				FileDel Record2("iaddfile")
				Record2.movenext
			wend
			Record2.close
			con2.Execute("DROP TABLE "+box+rs("用户名"))
end function		
			
'======清空当前信箱结束=================================================================================================================================================================
%>
