'this script attempts to ping a list of pc and see whether KBxxxxx exists
'Created by Michael Goulart - GOS-AD
'ExxonMobil 2009 
'Ver 1.0

on error resume next
dim strcomputer


x_online=0
y_offline=0
z_notinstalled=0


If Wscript.Arguments.Count <> 2 Then
  Wscript.Echo "Syntax Error. Correct syntax is:"
  Wscript.Echo
  Wscript.Echo "cscript kb.vbs <drive:filename.ext> <kbxxxxx>"
  Wscript.Echo "Example : cscript kb.vbs h:\Serverlist.txt kb835732"
  Wscript.Echo "where <filename.ext> should contains list of machines names and KBxxxxx you want to search."
  Wscript.Quit
End If

filename1 = Wscript.Arguments(0)
filename2 = Wscript.Arguments(1) & "_status.txt"
KBx = Wscript.Arguments(1) 


set objdict = createobject("scripting.dictionary")
set objfso = createobject("scripting.filesystemobject")
set objshell = wscript.createobject("wscript.shell")


'check existence of filename
if objfso.fileExists(filename1) = false then
	wscript.echo "Filename : " & filename1 & " does NOT exist."
	wscript.quit
end if


wscript.echo "This might take several minutes if you have many machines.Please be Patient =)"

Set objfile = objfso.CreateTextFile(filename2, True)
wscript.echo filename2 & " created in the directory where you run the script."

set objtextfile = objfso.opentextfile(filename1,1)


'read the workstation list file and stored into dictionary
i=0
do until objtextfile.atendofstream
	strnextline = objtextfile.readline
	objdict.add i, strnextline
	i = i+1
loop


for each objitem in objdict
	wscript.stdout.write(".")
	strcomputer = objdict.item(objitem)
	If isOnline(strComputer) Then
		
		if isinstalled(strcomputer, kbx) then

			objfile.writeline strcomputer & ", online, " & kbx & " installed"

		else
			objfile.writeline strcomputer & ", online, " & kbx & " NOT found. Please verify the machine."
			z_notinstalled = z_notinstalled + 1

		end if
		x_online = x_online + 1
	Else
		objfile.writeline strcomputer & ", offline."
		y_offline = y_offline + 1
	End If
next
	
summary


'==================================
'procedures and functions below
'==================================


sub summary
	objfile.writeline
	objfile.writeline "Total Online : " & x_online
	objfile.writeline "Total Offline : " & y_offline
	objfile.writeline "Machines not installed with patch : " & z_notinstalled
	objfile.writeline now
	wscript.echo vbcrlf & "Completed. Please see " & filename2 & " for information."
end sub


Function isOnline(strComp)
	isOnline = False
	Set objExec = objShell.Exec("cmd /C ping -n 1 " & strComp)
	Set objRegExp = New RegExp
	objRegExp.IgnoreCase = True
	objRegExp.Pattern = "reply"

	If objRegExp.Test(objExec.StdOut.ReadAll) Then isOnline = True

	Set objExec = Nothing
	Set objRegExp = Nothing
End Function 


Function isInstalled(strcomp,kbx)

	isinstalled = false

	Set objWMI = GetObject("WinMGMTS://" & strcomp & "/Root/CIMv2")
	strWQL = "SELECT * FROM Win32_QuickFixEngineering WHERE HotFixID = '" & kbx & "'"
	Set colResults = objWMI.ExecQuery(strWQL)


	For Each objItem In colResults
		isinstalled = true
	Next

	If isinstalled = true then isinstalled = true

	If IsObject(objItem) Then Set objItem = Nothing
	If IsObject(objWMI) Then Set objWMI = Nothing 

End Function
