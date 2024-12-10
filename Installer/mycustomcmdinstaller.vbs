'Continues on error
on error resume next

'Sets the shell
set oShell = Wscript.CreateObject("Wscript.Shell")

'Sets the filesystem
set filesys = CreateObject("Scripting.FileSystemObject")

'Gets this script name and path
strFile = Wscript.ScriptFullName

'Gets only the file name
filename=filesys.GetFileName(strFile)

'Gets the path where this script is located
pathFile = filesys.GetParentFolderName(strFile)

'Sets the path location for the ccadons folder
cadons = pathFile & "\ccaddons"

'Gets the current username
uname= CreateObject("Wscript.Network").Username

'Sets the install folder to the users appdata
installfolder="C:\Users\" & uname & "\AppData\Roaming\mycustomcmd"

'Sets the path for the customaddons folder
customfolder = installfolder & "\customcmdaddons"

'Sets the path for the customcmd batch file
batinstalllocation = installfolder & "\customcmd.bat"

'Sets the path for the fullscreen script
vbsinstalllocation = customfolder & "\dosfullscreen.vbs"

'Sets the backup install location
mydocsbackup=oShell.SpecialFolders("MyDocuments") & "\mycustomcmd-installer"

'Sets the backup install location for the customaddons
mydocsbackupspecial=mydocsbackup & "\customcmdaddons"

'Sets the backup install location for this script
mybackfileinstaller=mydocsbackup & "\" & filename

'Trys to creates all the required folders
sub mkdirs()
	on error resume next
	filesys.createfolder(installfolder)
	filesys.createfolder(customfolder)
	filesys.createfolder(mydocsbackup)
	filesys.createfolder(mydocsbackupspecial)
end sub

'Creates the customcmd batch file
sub makecustomcmd()
	'Continues on error
	on error resume next
	
	'Creates the arrays
	set customarr = createobject("system.collections.arraylist")
	set mcustomarr = createobject("system.collections.arraylist")
	set webarr = createobject("system.collections.arraylist")
	set mwebarr = createobject("system.collections.arraylist")
	set progarr = createobject("system.collections.arraylist")
	set mprogarr = createobject("system.collections.arraylist")

	'Creates 
	lines=""
	
	'Gets the folder ccadons
	set gfolder = filesys.GetFolder(cadons)
	
	'Gets the files in the ccadons folder
	set colfiles = gfolder.Files
	
	'Loops through all the files in the ccadons folder and adds it to the lines variable
	for each ofile in colfiles
		set rfile = filesys.OpenTextFile(ofile)					
		lines = lines + vbNewLine + rfile.readall()					
		rfile.close					
	next

	'Creates an array out of the lines variable
	insplit=split(lines,vbNewLine)

	'Loops through the insplit array
	for each ln in insplit
		'Gets the count of the string based on how many | chars
		count = Len(ln) - Len(Replace(ln,"|",""))

		'Splits the line using the | char
		lsplit=split(ln, "|")

		'This section below looks for 2 | and adds the variable to the proper sections.
		'Example: Command|Menu Description|Program
		if count = 2 then

			'Sets skip count
			skip=0

			
			'###Websites Section###
			
			'Looks for http in the first 4 characters in the third spot and checks to see if the skip variable is 0
			if left(lsplit(2), 4) = "http" and skip = 0 then

				'Creates the if statement in the batch file
				cline = "	if " & """" & "%incomm%" & """" & "==" & """" & lsplit(0) & """" & " set runthis=web " & """" & lsplit(2) & """"		
				
				'Checks the array if the string exists and if doesn't then adds it
				if webarr.contains(cline) = false then
					webarr.add cline		
				end if
				
				'Creates the menu item for the print out
				rline="	echo 	" & lsplit(1)
				
				'Checks if the line exists in the array and if doesn't it adds it to the array
				if mwebarr.contains(rline) = false then
					mwebarr.add rline
				end if
				
				'Sets skip variable to one
				skip=1
			end if
		
			'###Programs###
			
			'Looks for the extentions exe, vbs, bat and checks to see if the skip variable is 0
			if right(lsplit(2), 3) = "exe" or right(lsplit(2), 3) = "vbs" or right(lsplit(2), 3) = "cmd" or right(lsplit(2), 3) = "bat" and skip = 0 then
				
				'Creates the if statement in the batch file
				cline = "	if " & """" & "%incomm%" & """" & "==" & """" & lsplit(0) & """" & " set runthis=" & """" & lsplit(2) & """"		
				
				'Checks the array if the string exists and if doesn't then adds it
				if progarr.contains(cline) = false then
					progarr.add cline		
				end if
				
				'Creates the menu item for the print out
				rline="	echo 	" & lsplit(1)
				
				'Checks if the line exists in the array and if doesn't it adds it to the array
				if mprogarr.contains(rline) = false then
					mprogarr.add rline
				end if
				
				'Sets skip variable to one
				skip = 1
			end if
		
			'###Custom Scripts###
			
			'Checks to see if skip is 0 for the rest that may not be captured by the previous if statements
			if skip = 0 then
				
				'Creates the if statement in the batch file
				cline = "	if " & """" & "%incomm%" & """" & "==" & """" & lsplit(0) & """" & " set runthis=" & """" & lsplit(2) & """"
				
				'Checks the array if the string exists and if doesn't then adds it
				if customarr.contains(cline) = false then			
					customarr.add cline		
				end if
				
				'Creates the menu item for the print out
				rline="	echo 	" & lsplit(1)
				
				'Checks if the line exists in the array and if doesn't it adds it to the array
				if mcustomarr.contains(rline) = false then
					mcustomarr.add rline
				end if
			end if
		end if
	
	
		'This section below works exactly the same as the previous section except it looks for 3 | in relation to program varables.
		'Example: Command|Menu Description|Program|Program Variable
		if count = 3 then
			skip=0
			'Websites
			if left(lsplit(2), 4) = "http" and skip = 0 then	
				cline = "	if " & """" & "%incomm%" & """" & "==" & """" & lsplit(0) & """" & " set runthis=web " & """" & lsplit(2) & """" & " " & lsplit(3)		
				if webarr.contains(cline) = false then
					webarr.add cline		
				end if			
				rline="	echo 	" & lsplit(1)
				if mwebarr.contains(rline) = false then
					mwebarr.add rline
				end if
				skip=1
			end if
		
			'Programs
			if right(lsplit(2), 3) = "exe" or right(lsplit(2), 3) = "vbs" or right(lsplit(2), 3) = "cmd" or right(lsplit(2), 3) = "bat" and skip = 0 then	
				cline = "	if " & """" & "%incomm%" & """" & "==" & """" & lsplit(0) & """" & " set runthis=" & """" & lsplit(2) & """" & " " & lsplit(3)		
				if progarr.contains(cline) = false then			
					progarr.add cline		
				end if			
				rline="	echo 	" & lsplit(1)
				if mprogarr.contains(rline) = false then
					mprogarr.add rline
				end if
				skip = 1
			end if
		
			'Custom Scripts
			if skip = 0 then
				cline = "	if " & """" & "%incomm%" & """" & "==" & """" & lsplit(0) & """" & " set runthis=" & """" & lsplit(2) & """" & " " & lsplit(3)
				if customarr.contains(cline) = false then			
					customarr.add cline		
				end if			
				rline="	echo 	" & lsplit(1)
				if mcustomarr.contains(rline) = false then
					mcustomarr.add rline
				end if
			end if
		end if
	next
	
	'Sorts the arrays alphabetically
	customarr.sort()
	mcustomarr.sort()
	webarr.sort()
	mwebarr.sort()
	progarr.sort()
	mprogarr.sort()

	'The different sections of the batch script
	sect1="@echo off & SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION#n#rem Declares echo output and expansions#n##n#rem sets the title of the batch#n#TITLE Custom CMD#n##n#rem Sets the color scheme of the terminal#n#color 71#n##n#rem The start function is used when the script loads#n#:start#n#	rem Gets the default web browser from the registry#n#	for /f ^tokens=2*^ %%a in ($reg QUERY HKEY_CLASSES_ROOT\http\shell\open\command /ve$) do ( set ibrow=%%b )#n#	#n#	rem Sets the path for full screen script from the customcmdaddons folder#n#	set fullscreen=%~d0customcmdaddons\dosfullscreen.vbs#n#	#n#	rem Checks to see if the full screen script exists and if it does then it run it#n#	if EXIST ^%fullscreen%^ (#n#		start ^^ ^%fullscreen%^#n#	)#n#	#n#	rem Changes to the users directory#n#	cd C:\Users\%USERNAME%\#n#	#n#	rem Goes to the listcommands to print out the menu of commands#n#	goto :listincomms#n#	#n#rem The beginning of the script for user input#n#:begin#n#	#n#	rem Prints a blank line#n#	echo.#n#	#n#	rem Sets the variables to appropriate pre values#n#	set incomm=#n#	set runthis=#n#	set skip=#n#	set count=0#n#	#n#	rem Asks for user input#n#	set /P incomm=Enter incomm:#n#	#n#	rem Gets the first character from the input#n#	set qchar=%incomm:~0,1%#n#	#n#	rem Checks to see if the qchar variable is a quote and is so it goes to rcomm function to run command#n#	if ^%qchar%%qchar%^ == ^^^^ (#n#		set runthis=%incomm%#n#		goto :rcomm#n#	)#n#	#n#	rem Looks for quotes in string and if they exist it goes to hasarguments function#n#	echo %incomm% | findstr /C:^^^^ 1>nul#n#	if errorlevel 1 (#n#		echo.#n#	) else (#n#		set runthis=%incomm%#n#		goto :hasarguments#n#	)#n#	#n#	rem If input is empty it loops back to the beginning#n#	if ^%incomm%^==^^ goto :begin#n#	#n#	rem Calls the lower case function to set the input to lowercase#n#	call :tolower incomm#n#	#n#	rem general commands built into the script#n#	if ^%incomm%^==^reload^ set runthis=exit & start ^^ ^%~f0^#n#	if ^%incomm%^==^home^ set runthis=cd C:\Users\%USERNAME%\#n#	if ^%incomm%^==^gdate^ set runthis=echo %date% %time%#n#	if ^%incomm%^==^listcommands^ goto :listincomms#n#	if ^%incomm%^==^web^ set runthis=%ibrow%#n#	if ^%incomm%^==^exit^ exit#n#	#n#	rem custom commands#n#	:: VBS Break#n#"
	sect2="#n#	rem programs#n#	if ^%incomm%^==^notes^ set runthis=^C:\Program Files (x86)\Microsoft Office\Office16\ONENOTE.EXE^#n#	if ^%incomm%^==^snippet^ set runthis=^C:\WINDOWS\system32\SnippingTool.exe^#n#	if ^%incomm%^==^word^ set runthis=^C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE^#n#	if ^%incomm%^==^excel^ set runthis=^C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE^#n#	:: VBS Break#n#"
	sect3="#n#	rem websites#n#	:: VBS Break#n#"
	sect4="	#n#	rem Gets the first character from the runthis variable#n#	set qchar=%runthis:~0,1%#n#	#n#	rem Checks to see if the qchar variable is a quote and is so it goes to rcomm function to run command#n#	if ^%qchar%%qchar%^ == ^^^^ (#n#		set runthis=%runthis%#n#		goto :rcomm#n#	)#n#	#n#	rem Looks for quotes in string and if they exist it goes to hasarguments function#n#	echo %runthis% | findstr /C:^^^^ 1>nul#n#	if errorlevel 1 (#n#		echo.#n#	) else (#n#		set runthis=%runthis%#n#		goto :hasarguments#n#	)#n#	#n#	rem If runthis value is empty it just gets the command from the user input#n#	if ^%runthis%^ == ^^ set runthis=%incomm%#n#	goto :hasarguments#n#	#n#	rem If input is empty it loops back to the beginning#n#	if ^%incomm%^==^^ goto :begin#n#	#n#	rem Goes to hasarguments function#n#	goto :hasarguments#n#	#n#rem Arguments functions#n#:hasarguments#n#	#n#	rem Gets the first part of the input#n#	for /F ^tokens=1^ %%a in (^%runthis%^) do set b=%%a#n#	#n#	rem Gets the remainder part of the input#n#	call set args=%%runthis:%b%=%%#n#	#n#	rem Checks the first part of the input and calls the proper function#n#	if ^%b%^ == ^word^ (#n#		set runthis=^C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE^ %args%#n#	)#n#	if ^%b%^ == ^excel^ (#n#		set runthis=^C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE^ %args%#n#	)#n#	if ^%b%^ == ^notes^ (#n#		set runthis=^C:\Program Files (x86)\Microsoft Office\Office16\ONENOTE.EXE^ %args%#n#	)#n#	if ^%b%^ == ^web^ (#n#		set runthis=%ibrow% %args%#n#	)#n#	#n#	rem Goes to run the command function#n#	goto :rcomm#n#	#n#rem Run commands function#n#:rcomm#n#	#n#	rem Sets the counter to negative 1#n#	set count=-1#n#	#n#	rem Counts the number of periods#n#	for %%a in (%runthis:.= %) do set /a count+=1#n#	#n#	rem Looks for a period in string#n#	echo %runthis% | findstr /C:^.^ 1>nul#n#	#n#	rem If error level is 1 then it runs the command internally#n#	rem If error level is not 1 then it looks to see if count is 4 then it runs the command internally#n#	rem If count is not 4 then it runs the command externally#n#	if errorlevel 1 (#n#	 	%runthis%#n#	) else (#n#		if %count% == 4 (#n#			%runthis%#n#		) else (#n#			start ^^ %runthis%#n#		)#n#	)#n#	#n#	rem Loops back to the beginning#n#	goto :begin#n##n#rem Function to convert input to lowercase	#n#:tolower#n#for %%L in (a b c d e f g h i j k l m n o p q r s t u v w x y z) DO SET %1=!%1:%%L=%%L!#n#goto :EOF#n##n#rem The menu function#n#:listincomms#n#	echo ---------------#n#	echo General Commands#n#	echo ----------------#n#	echo 	exit - Exits this script#n#	echo 	listincomms - Shows this list of Commands#n#	echo 	reload - Reloads the script#n#	echo 	web - Loads a web page#n#	echo ----------------#n#	echo Custom Commands#n#	echo ---------------#n#:: VBS Break#n#"
	sect5="	echo ---------------#n#	echo Websites#n#	echo ---------------#n#:: VBS Break#n#"
	sect6="	echo ---------------#n#	echo Programs#n#	echo ---------------#n#	echo 	excel - Opens Microsoft Excel#n#	echo 	notes - Opens Microsoft Access#n#	echo 	snippet - Opens the Snipping Tool#n#	echo 	word - Opens Microsoft Word#n#	:: VBS Break#n#"
	sect7="	echo ---------------#n#	goto :begin#n#	:: VBS Break#n#"	






	'Replaces all the key characters in each section with the proper characters.
	s1=replace(sect1, "#n#", vbNewLine)
	s1=replace(s1,"^","""")
	s1=replace(s1,"$","'")
	s2=replace(sect2, "#n#", vbNewLine)
	s2=replace(s2,"^","""")
	s2=replace(s2,"$","'")
	s3=replace(sect3, "#n#", vbNewLine)
	s3=replace(s3,"^","""")
	s3=replace(s3,"$","'")
	s4=replace(sect4, "#n#", vbNewLine)
	s4=replace(s4,"^","""")
	s4=replace(s4,"$","'")
	s5=replace(sect5, "#n#", vbNewLine)
	s5=replace(s5,"^","""")
	s5=replace(s5,"$","'")
	s6=replace(sect6, "#n#", vbNewLine)
	s6=replace(s6,"^","""")
	s6=replace(s6,"$","'")
	s7=replace(sect7, "#n#", vbNewLine)
	s7=replace(s7,"^","""")
	s7=replace(s7,"$","'")
	
	'Writes the batch script
	set ofile = filesys.CreateTextFile(batinstalllocation)
	
	'Writes the first section
	ofile.write s1 & vbNewLine
	
	'Writes the if custom array
	ofile.write Join(customarr.ToArray, vbNewLine)
	ofile.write vbNewLine
	
	'Writes the second section
	ofile.write s2 & vbNewLine
	
	'Writes the if program array
	ofile.write Join(progarr.ToArray, vbNewLine)
	ofile.write vbNewLine
	
	'Writes the third section
	ofile.write s3 & vbNewLine
	
	'Writes the if web array
	ofile.write Join(webarr.ToArray, vbNewLine)
	ofile.write vbNewLine
	
	'Writes the fourth section
	ofile.write s4 & vbNewLine
	
	'Writes the menu custom array
	ofile.write Join(mcustomarr.ToArray, vbNewLine)
	ofile.write vbNewLine
	
	'Writes the fifth section
	ofile.write s5 & vbNewLine
	
	'Writes the menu web array
	ofile.write Join(mwebarr.ToArray, vbNewLine)
	ofile.write vbNewLine
	
	'Writes the sixth section
	ofile.write s6 & vbNewLine
	
	'Writes the menu program array
	ofile.write Join(mprogarr.ToArray, vbNewLine)
	ofile.write vbNewLine
	
	'Writes the seventh section
	ofile.write s7 & vbNewLine
	
	'Closes the file
	ofile.close()
end sub

'Creates the full screen vb script
sub mkfullscreen()
	'Continues on error
	on error resume next
	
	'The code for the full screen vbs script
	ln="$Sets the shell#n#Set WshShell = WScript.CreateObject(^WScript.Shell^)#n#$Activates the F11 key#n#WshShell.SendKeys ^{F11}^#n#"
	
	'Replaces all the key characters in each section with the proper characters.
	ln=replace(ln, "#n#", vbNewLine)
	ln=replace(ln,"^","""")
	ln=replace(ln,"$","'")
	
	'Creates the full screen vbs script
	set ofile = filesys.CreateTextFile(vbsinstalllocation)
	
	'Writes the code for the full screen vbs script
	ofile.write ln
	
	'Closes the full screen vbs script
	ofile.close()
end sub

'Creates a backup installion folder and files in the User's My Documents
sub copytodocuments()
	'Continues on error
	on error resume next
	
	'Copies the cadons folder
	filesys.copyfolder cadons, mydocsbackupspecial
	
	'Copies this installation script
	filesys.copyfile strFile, mybackfileinstaller
end sub

sub mkshortcut()
	'Creates the custom cmd shortcut in the start menu
	set ofile = filesys.CreateTextFile("C:\Users\" + uname + "\AppData\Roaming\Microsoft\Windows\Start Menu\mycustomcmd\mycustomcmd.url")
	
	'Writes the code for the full screen vbs script
	ofile.writeline "[INTERNETShortcut]"
	ofile.writeline "URL=file:///C:\Users\" + uname + "\AppData\Roaming\mycustomcmd\customcmd.bat"
	ofile.writeline "IconFile=C:\Windows\System32\cmd.exe"
	ofile.writeline "IconIndex=0"
	ofile.writeline "HotKey=0"
	ofile.writeline "IDList="
	ofile.writeline "[{000214A0-0000-0000-C000-000000000046}]"
	ofile.writeline "Prop3=19,9"

	
	'Closes the full screen vbs script
	ofile.close()
	
	'Creates the custom cmd shortcut on the Desktop
	' set ofile = filesys.CreateTextFile("C:\Users\" + uname + "\Desktop\mycustomcmd.url")
	
	'Writes the code for the full screen vbs script
	ofile.writeline "[INTERNETShortcut]"
	ofile.writeline "URL=file:///C:\Users\" + uname + "\AppData\Roaming\mycustomcmd\customcmd.bat"
	ofile.writeline "IconFile=C:\Windows\System32\cmd.exe"
	ofile.writeline "IconIndex=0"
	ofile.writeline "HotKey=0"
	ofile.writeline "IDList="
	ofile.writeline "[{000214A0-0000-0000-C000-000000000046}]"
	ofile.writeline "Prop3=19,9"

	
	'Closes the full screen vbs script
	ofile.close()
end sub

'Calls all the functions in order
mkdirs()
makecustomcmd()
mkfullscreen()
copytodocuments()
mkshortcut()
msgbox("Installation Complete")

